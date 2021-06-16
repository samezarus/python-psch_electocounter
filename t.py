import serial
import libscrc
from datetime import date, datetime, timedelta
import openpyxl
import logging
import pymysql
import sys


# Конфигурация модуля логов
logger = logging.getLogger('psch2.py')
logger.setLevel(logging.INFO)
fh = logging.FileHandler('log.txt')
formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] [%(name)s] [%(message)s]')
fh.setFormatter(formatter)
logger.addHandler(fh)


def str_to_hex(s):
    """
    Преобразование строки в нстойщий hex
    string -> hex ('6800' - > '\x68\x00')
    """

    return bytes.fromhex(s)


def int_to_hex_str(i):
    """
    Преобразование инта в строку хекса

    i: 0->255

    1 dec -> 01 hex, 20 dec -> 14 hex
    """

    return f'{i:0>2X}'


def validate_strhex(s):
    """
    Функция дополняет "слово"(пара байт)
    ps: Стоит обратить внимание, что если три символа, то 0 ставится в начале
    """
    result = s

    if len(s) == 3:
        result = f'0{s}'

    if len(s) == 2:
        result = f'00{s}'

    if len(s) == 1:
        result = f'000{s[0]}'

    return result


def half_hour_time():
    """
    :return: Возвращает список получасовых промежутков (00:00 - 00:30, 00:30 - 01:00 ... 23:30 - 23:59)
    """

    def add_zero(s_):
        result_ = s_
        if len(result_) == 1:
            result_ = '0' + result_

        return result_

    result = list()

    items_list = [i for i in range(0, 48)]

    l1 = 0
    l_fl = True
    l2 = 0

    r1 = 0
    r_fl = True
    r2 = 30
    for i in range(len(items_list)):
        l1s = add_zero(str(l1))
        l2s = add_zero(str(l2))
        r1s = add_zero(str(r1))
        r2s = add_zero(str(r2))

        if i == 47:
            r1s = '23'
            r2s = '59'

        s = l1s + ':' + l2s + '-' + r1s + ':' + r2s
        result.append(s)

        if r_fl:
            r1 += 1
            r2 = 0
            r_fl = False
        else:
            r2 = 30
            r_fl = True

        if l_fl:
            l1 = r1 - 1
            l2 = 30
            l_fl = False
        else:
            l1 = r1
            l2 = 0
            l_fl = True

    return result


def get_crc(s):
    """
    CRC-16/MODBUS

    s (string) - вида '6800' ('\x68\x00')
    result (string)
    """
    result = ''

    tx = str_to_hex(s)
    crc16 = libscrc.modbus(tx)
    s_hex = str(hex(crc16))

    if len(s_hex) == 6:
        result = f'{s_hex[4]}{s_hex[5]}{s_hex[2]}{s_hex[3]}'

    if len(s_hex) == 5:
        result = f'{s_hex[3]}{s_hex[4]}0{s_hex[2]}'

    return result


def make_true_date(s_date):
    """
    Преобразоване даты вида ддммгг(180221) в 2021-02-18
    """

    try:
        result = datetime.strptime(s_date, '%d%m%y')
        #result = f'{str(result)[0:4]}.{str(result)[5:7]}.{str(result)[8:10]}'
        result = f'{str(result)[0:10]}'
    except:
        result = s_date

    return result


def make_true_date_time(s_date, s_time):
    """
    Для БД или графиков

    s_date: ддммгг (180221)
    s_time: чч:мм-чч:мм (01:00-01:30)

    return гггг-мм-дд чч:мм (2021-02-18 01:30:00) (берётся второе время получасовки)
    """

    dt = f'{s_date} {s_time[6:11]}'

    try:
        result = datetime.strptime(dt, '%d%m%y %H:%M')
    except:
        result = datetime.now()

    return result


def mysql_execute(db_connection, query, commit_flag, result_type):
    """
    Функция для выполнния любых типов запросов к MySQL
    :dbCursor:   Указатель на курсор БД
    :query:      Запрос к БД
    :commitFlag: Делать ли коммит (True - Делать)
    :resultType: Тип результата (one - первую строку результата, all - весь результат)
    """

    result = None

    error_flag = False

    if db_connection:
        dbCursor = db_connection.cursor()
        try:
            dbCursor.execute(query)
        except:
            error_flag = True
            logger.error(f'Ошибка при выполнении запроса: {query}')

        if not error_flag:
            if commit_flag == True:
                db_connection.commit()

            if result_type == 'one':
                result = dbCursor.fetchone()

            if result_type == 'all':
                result = dbCursor.fetchall()

    return result


class PowerProfileItem:
    """
    Элемент профиля мощности
    """
    def __init__(self):
        self.a_plus = 0.0  # A+ кВт
        self.a_minus = 0.0  # A- кВт
        self.r_plus = 0.0  # R+ квар
        self.r_minus = 0.0  # R- квар
        self.date_param = ''  # Дата снятия (ддммгг)
        self.time_param = ':-:'  # Временной промежуток снятия (21:00-21:30 ...)
        self.date_time = datetime.now()  # Дата время для построения графика, время будет из второй части получасовки


class PSCH:
    def __init__(self, params_):
        self.global_error = False  # Флаг глобальной ошибки, после которой невозможно работать метадам класса

        logger.info('Инициализация приложения')

        #
        self.port = None
        self.counter_factory_number = params_['counter_factory_number']  # Заводской номер электросчётчика
        self.counter_identifier = params_[
            'counter_identifier']  # Идентификатор электросчётчика (десятичное значение)
        self.counter_divide = params_[
            'counter_divide']  # Постоянная электросчётчика в зависимости от типа и варианта исполнения
        self.counter_transform = params_[
            'counter_transform']  # Коэффициент трансформации (Следует узнать у энергетика)
        self.counter_password = params_['counter_password']  # Пароль для доступа к электросчётчику
        self.xlsx_template = params_['xlsx_template']  # Шаблон для выгруки ексел
        self.xlsx_result = params_['xlsx_result']  # Результирующий файл ексель
        self.prevmonth = ''  # Параметр для результирующего фаля ексель. гггг_мм (2021_06)
        self.mysql_host = params_['mysql_host']  #
        self.mysql_db = params_['mysql_db']  #
        self.mysql_user = params_['mysql_user']  #
        self.mysql_password = params_['mysql_password']  #


        try:
            self.port = serial.Serial(
                port=params_['port_name'],  # Имя com-порта
                baudrate=params_['port_baudrate'],  # Скорость соединения
                parity=params_['port_parity'],  # Четность
                stopbits=params_['port_stopbits'],  # Стоповые биты
                bytesize=params_['port_bytesize'],  # Размер байт
                timeout=params_['port_timeout']  # Таймаут
            )

            logger.info('Инициализация com-порта прошла успешно')


        except:
            logger.error('Инициализация com-порта произошла с ошибкой')
            self.global_error = True

    def prepare_command(self, cmd):
        """
        Добавляет в конец команды crc

        cmd (string) - команда для отправки в com-порт
        result (string)
        """

        result = ''
        if not self.global_error:
            try:
                crc = get_crc(cmd)
                result = f'{cmd}{crc}'
            except:
                logger.error(f'Ошибка при добавлении CRC к команде: {cmd}')
                self.global_error = True

        return result

    def send_to_port(self, port_, counter_identifier_, cmd):
        """
        Посылает запрос в com-порт.

        cmd - должен быть без crc на конце. Функция сама его подставит
        """

        result = b''

        cmd_print = False  # Флаг печати ввода/вывода команд в консоль

        if not self.global_error:
            try:
                port_.flushInput()
                port_.flushOutput()

                cmd = self.prepare_command(f'{int_to_hex_str(counter_identifier_)}{cmd}')

                hex_cmd = str_to_hex(cmd)

                if cmd_print:
                    print(f'TX:    {hex_cmd.hex()}')

                port_.write(hex_cmd)

                end_time = datetime.now() + timedelta(seconds=port_.timeout)

                while True:
                    result += port_.read(port_.inWaiting())

                    if datetime.now() > end_time:
                        break

                if cmd_print:
                    print(f'RX:    {result.hex()}')
                    print('')
            except:
                logger.error(f'Ошибка при отправке команды электросчётчику №: {cmd}')
                self.global_error = True

        return result.hex()

    def test_counter(self, port_, counter_identifier_):
        """
        Проверка, доступен ли счётчик
        Проверка строиться на посылке счётчику короткой строки, если в ответ пришла посылаемая строка,
        то тест пройден
        """

        result = False
        cmd = f'00'

        if not self.global_error:
            try:
                r = self.send_to_port(port_, counter_identifier_, cmd)
                if len(r) > 0:
                    logger.info(f'Тест связи с электросчётчиком №: {counter_identifier_} пройден')
                    result = True
            except:
                logger.error(f'Тест связи с электросчётчиком №: {counter_identifier_} не пройден')
                self.global_error = True

        return result

    def open_channel(self, port_, counter_identifier_, counter_password_):
        """
        Открытие канала связи со счётчиком
        """

        result = False
        hex_password = ''

        if not self.global_error:
            try:
                hex_password = counter_password_.encode().hex()
            except:
                logger.error(f'Ошибка при конвертации пароля в HEX строку: {counter_password_}')
                self.global_error = True

        if not self.global_error:
            try:
                cmd = f'01{hex_password}'
                r = self.send_to_port(port_, counter_identifier_, cmd)

                etalon_ansver = str_to_hex(self.prepare_command(f'{int_to_hex_str(counter_identifier_)}00')).hex()

                if len(r) != 0:
                    if etalon_ansver == r:
                        result = True
            except:
                logger.error(f'Ошибка при открытии канала с электросчётчиком №: {counter_identifier_}')
                self.global_error = True

        return result

    def close_channel(self, port_, counter_identifier_):
        """
        Закрытие канала связи со счётчиком
        """

        result = False

        if not self.global_error:
            try:
                cmd = f'02'
                r = self.send_to_port(port_, counter_identifier_, cmd)

                etalon_ansver = str_to_hex(self.prepare_command(f'{int_to_hex_str(counter_identifier_)}00')).hex()

                if len(r) != 0:
                    if etalon_ansver == r:
                        result = True
            except:
                logger.error(f'Ошибка при закрытии канала с электросчётчиком №: {counter_identifier_}')
                self.global_error = True

        return result

    def read_power_profile_pointer_on_date(self, port_, counter_identifier_, date_):
        """
        Поиск указателя базового массива профиля мощности на заданную дату
        """

        r = ''

        if not self.global_error:
            # Прочитать версию ПО счетчика
            cmd = f'0803'
            self.send_to_port(port_, counter_identifier_, cmd)

            # Прочитать установленные программируемые флаги из счетчика
            cmd = f'0809'
            self.send_to_port(port_, counter_identifier_, cmd)

            # Прочитать время интегрирования мощности массива профиля счетчика
            cmd = f'0806'
            self.send_to_port(port_, counter_identifier_, cmd)

            # Прочитать текущий указатель первого (или единственного) базового массива профиля мощности счетчика
            cmd = f'0804'
            self.send_to_port(port_, counter_identifier_, cmd)

            #  Найти адрес заголовка на дату
            cmd = f'032800FFFFFF{date_}FF1E'
            self.send_to_port(port_, counter_identifier_, cmd)

            # Найти указатель базаового массива профиля мощности на начало искомой даты
            dt_end = datetime.now() + timedelta(seconds=10)  # Время не больше которого должен идти поиск
            p = '1'
            cmd = f'081800'
            while p != '0':
                r = self.send_to_port(port_, counter_identifier_, cmd)

                if len(r) == 16:
                    p = r[3]
                    r = f'{r[8]}{r[9]}{r[10]}{r[11]}'

                if dt_end < datetime.now():
                    logger.error(
                        f'Ошибка при попытке найти указатель электросчётка №: {counter_identifier_} на дату: {make_true_date(date_)}')
                    self.global_error = True
                    break

                """
                try:
                    p = r[3]
                    r = f'{r[8]}{r[9]}{r[10]}{r[11]}'
                except:
                    logger.error(f'Ошибка при попытке найти указатель электросчётка №: {counter_identifier_} на дату: {date_}')
                    self.global_error = True
                    break
                """

        return r

    def read_7bit_header(self, port_, counter_identifier_, date_, pointer_):
        """
        Прочитать 7 байт информации (заголовок профиля) из памяти № 03h c адреса "pointer"
        """

        result = False

        if not self.global_error:
            cmd = f'0603{pointer_}07'
            r = self.send_to_port(port_, counter_identifier_, cmd)

            part = f'00{date_}011E'

            if r.find(part) == -1:
                result = True

        return result

    def read_transformation_coefficient(self, port_, counter_identifier_):
        """
        Прочитать установленные коэффициенты трансформации счетчика
        """

        result = {
            'kn': 0,  # Кн
            'kt': 0,  # Кт
            'dimensionality': 0,  # Признак размерности кВт ч
            'whole_part': 0,  # Целая часть Кн*Кт/100000
            'fractional_part': 0  # Дробная часть Кн*Кт/100
        }

        cmd = f'0802'
        r = self.send_to_port(port_, counter_identifier_, cmd)

        if not self.global_error:
            if len(r) == 26:
                result['kn'] = int(f'{r[2]}{r[3]}{r[4]}{r[5]}', 16)
                result['kt'] = int(f'{r[6]}{r[7]}{r[8]}{r[9]}', 16)
                result['dimensionality'] = int(f'{r[10]}{r[11]}', 16)
                result['whole_part'] = int(f'{r[12]}{r[13]}', 16)
                result['fractional_part'] = int(f'{r[14]}{r[15]}{r[16]}{r[17]}{r[18]}{r[19]}{r[20]}{r[21]}', 16)

        return result

    def read_power_profile_line(self, port_, counter_identifier_, index_, pointer_):
        """
        Прочитать первую или очередную строку с данными профиля мощности

        index_ не должен быть равным 0 (проблемы CRC). Только 1 -> 255
        """

        result = ''

        ma = '03'  # № адреса памяти
        bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

        cmd = f'0C{index_}{ma}{pointer_}{bytes_count}'
        r = self.send_to_port(port_, counter_identifier_, cmd)

        # Отсекаем номер счётчика и индекс, отсекаем CRC
        if not self.global_error:
            try:
                result = r[4:-4]
            except:
                logger.error(f'Ошибка чтении строки с данными профиля мощности')
                self.global_error = True

        return result

    def prepare_power_profile_item(self, ppi1, ppi2, hhx, divide_, transform_):
        """
        Парсим данные
        ppi1: элеиент (PowerProfileItem()) перваой получасовки часа
        ppi2: элеиент (PowerProfileItem()) второй получасовки часа
        hhx:
        """

        if not self.global_error:
            #if len(hhx) == 32:
            try:
                ppi1.a_plus = round((int(hhx[0:4], 16) / divide_) * transform_, 2)
                ppi1.a_minus = round((int(hhx[4:8], 16) / divide_) * transform_, 2)
                ppi1.r_plus = round((int(hhx[8:12], 16) / divide_) * transform_, 2)
                ppi1.r_minus = round((int(hhx[12:16], 16) / divide_) * transform_, 2)

                ppi2.a_plus = round((int(hhx[16:20], 16) / divide_) * transform_, 2)
                ppi2.a_minus = round((int(hhx[20:24], 16) / divide_) * transform_, 2)
                ppi2.r_plus = round((int(hhx[24:28], 16) / divide_) * transform_, 2)
                ppi2.r_minus = round((int(hhx[28:32], 16) / divide_) * transform_, 2)
            except:
                logger.error(f'Ошибка при парсинге получасовок часа')
                self.global_error = True

    def read_power_profile(self, port_, counter_identifier_, pointer_, date_, divide_, transform_):
        """
        Прочитать все значения профиля мощности на дату date_
        pointer -  значение из функции read_power_profile_pointer_on_date()
        :date: ddmmyy
        """

        result = []

        pointer_ = validate_strhex(pointer_)

        fl = False  # Флаг удачного поиска 24-х пар получасовок

        bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

        data = ''

        if not self.global_error:
            try:
                #print(f'Чтение профиля мощности за {make_true_date(date_)}')
                logger.info(f'Чтение профиля мощности за {make_true_date(date_)}')

                # Циклично пытаемся найти пары получкасовок
                # Данных пар не обязательно должно быть 24 (по две на час)
                for i in range(1, 255):
                    index = int_to_hex_str(i)

                    data += self.read_power_profile_line(port_, counter_identifier_, index, pointer_)

                    dec_pointer = int(pointer_, 16) + int(bytes_count, 16)

                    if dec_pointer < 65535:  # Пока не вышли за пределы адреса FFFFh
                        pointer_ = int_to_hex_str(dec_pointer)
                        pointer_ = validate_strhex(pointer_)
                    else:  # Обнуляем указатель
                        pointer_ = '0000'

                    # Точно прерываем цикл, так уже на всякий случай получили лишнюю строку ответа
                    if fl:
                        break

                    # Если нашли идентификатор 24-й пары получасовок
                    # то не выходим, а даём получить ещё строку ответа во избежание
                    # ситуации, когда признак пары получасовок находится в конце строки текущего ответа
                    if data.find(f'23{date_}') != -1:
                        fl = True

                # Если успешно нашли 24-ю(последнюю) пару получасовок
                if fl:
                    hht = half_hour_time()

                    pos = 0  # Индекс данных в получасовках (для получения времени из hht)

                    """
                    Результат всегда будет содержать 48 получасовок за сутки
                    И не беда, что данные будут не на все получасовки (такое случается 
                    из-за постоянного перезатирания данных)
                    """
                    for i in range(0, 24):
                        true_date = make_true_date(date_)

                        item1 = PowerProfileItem()  # Первая получасовка часа
                        item1.date_param = true_date
                        item1.time_param = hht[pos]
                        item1.date_time = make_true_date_time(date_, hht[pos])
                        result.append(item1)

                        item2 = PowerProfileItem()  # Вторая получасовка часа
                        item2.date_param = true_date
                        item2.time_param = hht[pos + 1]
                        item2.date_time = make_true_date_time(date_, hht[pos + 1])
                        result.append(item2)

                        # Идентификатор пары получасовок
                        if i > 9:
                            h = f'{i}{date_}'
                        else:
                            h = f'0{i}{date_}'

                        l = data.split(h)

                        if len(l) > 1:
                            hhx = l[1]
                            if len(hhx) > 39:
                                hhx = hhx[8:40]
                                self.prepare_power_profile_item(item1, item2, hhx, divide_, transform_)

                        pos += 2
            except:
                logger.error(f'Ошибка при чтении профиля мощности за {make_true_date(date_)}')
                self.global_error = True

        return result

    def get_prevday_power_profile(self, port_, counter_identifier_, divide_, transform_):
        """
        Прочитать все значения профиля мощности на вчера
        """

        result = []

        dtn = datetime.now()  # Текущий тайм стемп
        ydtn = dtn - timedelta(days=1)  # Вчерашний таймстемп (нужна исключительно дата)

        date_param = f'{str(ydtn)[8:10]}{str(ydtn)[5:7]}{str(ydtn)[2:4]}'  # Формат даты для посылки в электросчётчик

        if not self.global_error:
            pointer_ = self.read_power_profile_pointer_on_date(port_, counter_identifier_, date_param)

            if self.read_7bit_header(port_, counter_identifier_, date_param, pointer_):
                self.read_transformation_coefficient(port_, counter_identifier_)

                result = self.read_power_profile(port_, counter_identifier_, pointer_, date_param, divide_, transform_)

        return result

    def get_prevmonth_power_profile(self, port_, counter_identifier_, divide_, transform_):
        """
        Прочитать все значения профиля мощности за прошлый месяц
        """

        result = []  # Массив суточных массивов профилей мощности

        # последний день предыдущего месяца
        last_day_prev_month = date.today().replace(day=1) - timedelta(days=1)

        # Первый день предыдущего месяца
        first_day_prev_month = date.today().replace(day=1) - timedelta(days=last_day_prev_month.day)

        d = first_day_prev_month

        self.prevmonth = f'{str(d)[0:4]}_{str(d)[5:7]}'

        if not self.global_error:
            while d != last_day_prev_month + timedelta(days=1):
                date_param = f'{str(d)[8:10]}{str(d)[5:7]}{str(d)[2:4]}'

                pointer_ = self.read_power_profile_pointer_on_date(port_, counter_identifier_, date_param)

                if self.read_7bit_header(port_, counter_identifier_, date_param, pointer_):
                    #self.read_transformation_coefficient(port_, counter_identifier_)

                    day_data = self.read_power_profile(port_,
                                                       counter_identifier_,
                                                       pointer_,
                                                       date_param,
                                                       divide_,
                                                       transform_)

                    result.extend(day_data)

                d += timedelta(days=1)

            self.close_channel(port_, counter_identifier_)

        return result

    def print_power_profile(self, power_profile_items_):
        """
        Вывод в консоль значений профиля мощности
        """

        for item in power_profile_items_:
            print(f'{item.date_time} |'
                  f'{item.date_param} '
                  f'{item.time_param} | '
                  f'{item.a_plus} | '
                  f'{item.a_minus} | '
                  f'{item.r_plus} | '
                  f'{item.r_minus} |')

    def power_profile_to_xlsx(self, power_profile_items_, template_xlsx_, result_xlsx_):
        """
        Сохранение профиля мощности в эсель файл result_xlsx_, по шаблону template_xlsx_
        """
        wb = None

        if not self.global_error:
            try:
                wb = openpyxl.load_workbook(filename=template_xlsx_)
            except:
                logger.error(f'Ошибка при попытке открыть шаблон {template_xlsx_}')
                self.global_error = True

        if not self.global_error:
            # Обрабатываем вкладку "Профили нагрузки"
            try:
                page2 = wb['Профили нагрузки']
                page2['O1'] = ''  # Название устройства
                page2['O2'] = ''  # Адрес объекта
                page2['O3'] = ''  # Адрес устройства
                page2['O4'] = ''  # Идентификатор устройства
                page2['O5'] = ''  # Заводской номер
                page2['AG1'] = datetime.now()

                for i, item in enumerate(power_profile_items_, start=0):
                    page2[f"""A{i + 13}"""] = f'{item.date_param}  {item.time_param}'
                    page2[f"""J{i + 13}"""] = f'{str(item.a_plus)}'
            except:
                logger.error(f'Ошибка при добавлении данных в шаблон {template_xlsx_}')
                self.global_error = True

        if not self.global_error:
            try:
                wb.save(filename=result_xlsx_)
                logger.info(f'Успешное сохранение профиля нагрузки в файл {result_xlsx_}')
            except:
                logger.error(f'Ошибка при сохранении файла {result_xlsx_}')
                self.global_error = True

    def power_profile_to_mysql(self, power_profile_items_):
        """

        """
        if not self.global_error:
            db = None

            try:
                db = pymysql.connect(
                    host=self.mysql_host,
                    db=self.mysql_db,
                    user=self.mysql_user,
                    password=self.mysql_password,
                    cursorclass=pymysql.cursors.DictCursor)

                logger.info(f'Успешное подключение к БД {self.mysql_host}.{self.mysql_db}')
            except:
                logger.error(f'Ошибка при подключении к БД {self.mysql_host}.{self.mysql_db}')

            if db != None:
                #
                query = f"select counterID from counters where serialNumber = '{self.counter_factory_number}'"
                mysql_result = mysql_execute(db, query, False, 'one')

                if mysql_result != None:
                    counter_id = mysql_result['counterID']

                    for item in power_profile_items_:
                        query = f"insert into loadprofiles (counterID, dt, activePowerConsumed, reactiveEnergyConsumed) " \
                            f" select '{counter_id}', " \
                            f" '{item.date_time}', " \
                            f" {item.a_plus}, " \
                            f" {item.r_plus} " \
                            f" FROM (SELECT 1) as dummytable " \
                            f" WHERE NOT EXISTS (SELECT 1 FROM loadprofiles WHERE " \
                            f" counterID='{counter_id}' and " \
                            f" dt='{item.date_time}' " \
                            f")"
                        mysql_execute(db, query, True, 'one')

                db.close()

    def power_profile_to_mysql_by_days(self, port_, counter_identifier_, divide_, transform_, days_count_):
        """
        Записывает в БД профиль мощности за указанное количество дней
        """

        dtn = datetime.now()  # Текущий тайм стемп

        i = 0

        while i != days_count_+1:
            i += 1
            ydtn = dtn - timedelta(days=i)
            date_param = f'{str(ydtn)[8:10]}{str(ydtn)[5:7]}{str(ydtn)[2:4]}'  # Формат даты для посылки в электросчётчик

            if not self.global_error:
                pointer_ = self.read_power_profile_pointer_on_date(port_, counter_identifier_, date_param)

                #if self.read_7bit_header(port_, counter_identifier_, date_param, pointer_):
                day_data = self.read_power_profile(port_,
                                                    counter_identifier_,
                                                    pointer_,
                                                    date_param,
                                                    divide_,
                                                    transform_)
                self.power_profile_to_mysql(day_data)



########################################################################################################################

ext_cmd = ''  # Параметр переданный скрипту

if len(sys.argv) > 1:
    ext_cmd = sys.argv[1]

print(ext_cmd)

params = {
    'port_name': 'COM3',  # Имя com-порта
    'port_baudrate': 9600,  # Скорость соединения
    'port_parity': 'N',  # Четность
    'port_stopbits': 1,  # Стоповые биты
    'port_bytesize': 8,  # Размер байт
    'port_timeout': 0.3,  # Таймаут
    'counter_factory_number': '1103181104',  # Заводской номер электросчётчика
    'counter_identifier': 104,  # Идентификатор электросчётчика (десятичное значение)
    'counter_divide': 1250,  # Постоянная счетчика в зависимости от типа и варианта исполнения (ПСЧ-4ТМ.05МК)
    'counter_transform': 400,  # Коэффициент трансформации (Следует узнать у энергетика)
    'counter_password': '000000',  # Пароль для доступа к электросчётчику
    'xlsx_template': 'template.xlsx',  # Шаблон для выгруки ексел
    'xlsx_result': 'result.xlsx',  # Результирующий файл ексель
    'mysql_host': 'localhost',  #
    'mysql_db': '',  #
    'mysql_user': '',  #
    'mysql_password': ''  #
}

psch = PSCH(params)

if psch.test_counter(psch.port, psch.counter_identifier):
    logging.info(f'Тест электросчётчика {psch.counter_identifier} пройден')

    if psch.open_channel(psch.port, psch.counter_identifier, psch.counter_password):
        # тесты
        #ext_cmd = '-test'
        if ext_cmd == '-test':
            date_param = '270521'
            pointer = psch.read_power_profile_pointer_on_date(psch.port,
                                                              psch.counter_identifier,
                                                              date_param)
            items = psch.read_power_profile(psch.port,
                                            psch.counter_identifier,
                                            pointer,
                                            date_param,
                                            psch.counter_divide,
                                            1)
            psch.print_power_profile(items)

        # Сохранение в ексель
        #ext_cmd = '-xlsx'
        if ext_cmd == '-xlsx':
            #items = psch.read_power_profile_pointer_on_date(psch.port, psch.counter_identifier, '270521')

            # Профиль мощности за вчера (для ускорения тестов)
            # items = psch.get_prevday_power_profile(psch.port, psch.counter_identifier, psch.counter_divide, 1)

            # Профиль мощности за прошлый месяц
            items = psch.get_prevmonth_power_profile(psch.port,
                                                     psch.counter_identifier,
                                                     psch.counter_divide,
                                                     1)

            fn = f'C:/temp/Приморский край, Владивосток, Народный проспект, 20/{psch.counter_identifier}_{psch.prevmonth}.xlsx'
            psch.power_profile_to_xlsx(items,
                                       psch.xlsx_template,
                                       fn)

        # Сохранение в БД
        #ext_cmd = '-mysql'
        if ext_cmd == '-mysql':
            # Профиль мощности за 90
            days_count = 90
            psch.power_profile_to_mysql_by_days(psch.port,
                                                psch.counter_identifier,
                                                psch.counter_divide,
                                                1,
                                                days_count)
    else:
        """
        Одна из причин - это неверный пароль 
        """
        logging.error(f'Неудалось открыт канал с электросчётком {psch.counter_identifier}')
else:
    logging.error(f'Тест электросчётчика {psch.counter_identifier} не пройден')
