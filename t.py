import serial
import libscrc
from datetime import date, datetime, timedelta
import csv


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
        self.date_time = datetime.now() # Дата время для построения графика, время будет из второй части получасовки


def str_to_hex(s):
    # string -> hex ('6800' - > '\x68\x00')

    return bytes.fromhex(s)


def int_to_hex_str(i):
    """
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


def prepare_command(cmd):
    """
    Добавляет в конец команды crc

    cmd (string) - команда для отправки в com-порт
    result (string)
    """

    crc = get_crc(cmd)
    result = cmd + crc

    return result


def send_to_port(port_, counter_identifier_, cmd, cmd_print):
    """
    Посылает запрос в com-порт.

    cmd - должен быть без crc на конце. Функция сама его подставит
    """

    cmd_print = False

    port_.flushInput()
    port_.flushOutput()

    cmd = prepare_command(f'{int_to_hex_str(counter_identifier_)}{cmd}')

    hex_cmd = str_to_hex(cmd)

    if cmd_print:
        print(f'TX    {hex_cmd.hex()}')

    port_.write(hex_cmd)

    result = b''

    end_time = datetime.now() + timedelta(seconds=port_.timeout)

    while True:
        result += port_.read(port.inWaiting())

        if datetime.now() > end_time:
            break
    if cmd_print:
        print(f'RX    {result.hex()}')
        print('')

    return result.hex()


def test_counter(port_, counter_identifier_):
    """
    Проверка, доступен ли счётчик
    """

    result = False
    cmd = f'00'

    try:
        r = send_to_port(port_, counter_identifier_, cmd, True)
        if len(r) > 0:
            result = True
    except:
        pass

    return result


def open_channel(port_, counter_identifier_, counter_password):
    """
    Открытие канала связи со счётчиком
    """

    result = False

    hex_password = counter_password.encode().hex()

    cmd = f'01{hex_password}'
    r = send_to_port(port_, counter_identifier_, cmd, True)

    etalon_ansver = str_to_hex(prepare_command(f'{int_to_hex_str(counter_identifier_)}00')).hex()

    if len(r) != 0:
        if etalon_ansver == r:
            result = True

    return result


def close_channel(port_, counter_identifier_):
    """
    Открытие канала связи со счётчиком
    """

    result = False

    cmd = f'02'
    r = send_to_port(port_, counter_identifier_, cmd, True)

    etalon_ansver = str_to_hex(prepare_command(f'{int_to_hex_str(counter_identifier_)}00')).hex()

    if len(r) != 0:
        if etalon_ansver == r:
            result = True

    return result


def read_power_profile_pointer_on_date(port_, counter_identifier_, date_):
    """
    Поиск первого (или единственного) указателя базового массива профиля мощности
    на заданную дату
    """

    cmd = f'0803'
    send_to_port(port_, counter_identifier_, cmd, True)

    cmd = f'0809'
    send_to_port(port_, counter_identifier_, cmd, True)

    cmd = f'0806'
    send_to_port(port_, counter_identifier_, cmd, True)

    cmd = f'0804'
    send_to_port(port_, counter_identifier_, cmd, True)

    cmd = f'032800FFFFFF{date_}FF1E'
    send_to_port(port_, counter_identifier_, cmd, True)

    # Найти указатель базаового массива профиля мощности на начало искомой даты
    r = ''
    p = '1'
    while p != '0':
        cmd = f'081800'
        r = send_to_port(port_, counter_identifier_, cmd, True)
        p = r[3]

    r = f'{r[8]}{r[9]}{r[10]}{r[11]}'

    return r


def read_7bit_header(port_, counter_identifier_, date_, pointer_):
    """
    Прочитать 7 байт информации (заголовок профиля) из памяти № 03h c адреса "pointer"
    """
    result = True

    cmd = f'0603{pointer_}07'
    r = send_to_port(port_, counter_identifier_, cmd, True)

    part = f'00{date_}011E'

    if r.find(part) != -1:
        result = False

    return result


def read_transformation_coefficient(port_, counter_identifier_):
    """
    Прочитать установленные коэффициенты трансформации счетчика
    """

    result = {
        'kn': 0,              # Кн
        'kt': 0,              # Кт
        'dimensionality': 0,  # Признак размерности кВт ч
        'whole_part': 0,      # Целая часть Кн*Кт/100000
        'fractional_part': 0  # Дробная часть Кн*Кт/100
    }

    cmd = f'0802'
    r = send_to_port(port_, counter_identifier_, cmd, True)

    if len(r) == 26:
        result['kn'] = int(f'{r[2]}{r[3]}{r[4]}{r[5]}', 16)
        result['kt'] = int(f'{r[6]}{r[7]}{r[8]}{r[9]}', 16)
        result['dimensionality'] = int(f'{r[10]}{r[11]}', 16)
        result['whole_part'] = int(f'{r[12]}{r[13]}', 16)
        result['fractional_part'] = int(f'{r[14]}{r[15]}{r[16]}{r[17]}{r[18]}{r[19]}{r[20]}{r[21]}', 16)

    return result


def read_power_profile_line(port_, counter_identifier_, index_, pointer_):
    """
    Прочитать первую или очередную строку с данными профиля мощности

    index_ не должен быть равным 0 (проблемы CRC). Только 1 -> 255
    """

    ma = '03'  # № адреса памяти
    bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

    cmd = f'0C{index_}{ma}{pointer_}{bytes_count}'
    r = send_to_port(port_, counter_identifier_, cmd, True)

    # Отсекаем номер счётчика и индекс, отсекаем CRC
    return r[4:-4]


def prepare_power_profile_item(ppi1, ppi2, hhx, divide_, transform_):
    """
    Парсим данные получасовки, и возвращаем её в виде класса PowerProfileItem()
    """

    if len(hhx) == 32:
        ppi1.a_plus = round((int(hhx[0:4], 16) / divide_) * transform_, 2)
        ppi1.a_minus = round((int(hhx[4:8], 16) / divide_) * transform_, 2)
        ppi1.r_plus = round((int(hhx[8:12], 16) / divide_) * transform_, 2)
        ppi1.r_minus = round((int(hhx[12:16], 16) / divide_) * transform_, 2)

        ppi2.a_plus = round((int(hhx[16:20], 16) / divide_) * transform_, 2)
        ppi2.a_minus = round((int(hhx[20:24], 16) / divide_) * transform_, 2)
        ppi2.r_plus = round((int(hhx[24:28], 16) / divide_) * transform_, 2)
        ppi2.r_minus = round((int(hhx[28:32], 16) / divide_) * transform_, 2)


def read_power_profile(port_, counter_identifier_, pointer_, date_, divide_, transform_):
    """
    Прочитать все строки профиля мощности на дату
    pointer -  значение из функции read_power_profile_pointer_on_date()
    """

    print(f'Чтение профиля мощности за {make_true_date(date_)}')

    pointer_ = validate_strhex(pointer_)

    result = []

    fl = False  # Флаг удачного поиска 24-х пар получасовок

    bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

    data = ''

    # Циклично пытаемся найти пары получкасовок
    # Данных пар не обязательно должно быть 24 (по две на час)
    for i in range(1, 255):
        index = int_to_hex_str(i)

        data += read_power_profile_line(port_, counter_identifier_, index, pointer_)

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
        И не беда, что данные будут не на все получасовки (такое случается из-за постоянного перезатирания данных)
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
                    prepare_power_profile_item(item1, item2, hhx, divide_, transform_)

            pos += 2

    return result


def get_prevmonth_power_profile(port_, counter_identifier_, divide_, transform_):
    """
    Получение профиля мощности за прошлый месяц от текущего
    """

    result = []  # Массив суточных массивов профилей мощности

    # последний день предыдущего месяца
    last_day_prev_month = date.today().replace(day=1) - timedelta(days=1)

    # Первый день предыдущего месяца
    first_day_prev_month = date.today().replace(day=1) - timedelta(days=last_day_prev_month.day)

    d = first_day_prev_month
    while d != last_day_prev_month + timedelta(days=1):
        date_param = f'{str(d)[8:10]}{str(d)[5:7]}{str(d)[2:4]}'

        pointer_ = read_power_profile_pointer_on_date(port_, counter_identifier_, date_param)

        if read_7bit_header(port_, counter_identifier_, date_param, pointer_):
            read_transformation_coefficient(port_, counter_identifier_)

            day_data = read_power_profile(port_, counter_identifier_, pointer_, date_param, divide_, transform_)

            result.append(day_data)

        d += timedelta(days=1)

    return result



########################################################################################################################


port = serial.Serial(
    port='COM3',
    baudrate=9600,
    parity='N',
    stopbits=1,
    bytesize=8,
    timeout=0.3
    )

counter_identifier = 104
divide = 1250  # Делитель в зависимости от модели счётчика (ПСЧ-4ТМ.05МК)
transform = 400  # Коэффициент трансформации (по хорошему его нужно прописывать в электросчётчике, а потом читать)

online = test_counter(port, counter_identifier)
if online:
    print(f'Счётчик {counter_identifier} доступен')

    channel = open_channel(port, counter_identifier, '000000')
    if channel:
        print(f'Счётчик {counter_identifier} канал открыт')

        """
        # Профиль на дату
        date_p = '160221'

        pointer = read_power_profile_pointer_on_date(port, counter_identifier, date_p)
        print(f'{pointer} -- указатель на дату {date_p}')

        if read_7bit_header(port, counter_identifier, date_p, pointer):
            tc = read_transformation_coefficient(port, counter_identifier)
            print(tc)
            
            ppl = read_power_profile(port, counter_identifier, pointer, date_p, divide, transform)
            for item in ppl:
                print(f'{item.date_time} |'
                      f'{item.date_param} '
                      f'{item.time_param} | '
                      f'{item.a_plus} | '
                      f'{item.a_minus} | '
                      f'{item.r_plus} | '
                      f'{item.r_minus} |')
        """

        # Профиль за предыдущий месяц
        mont_data = get_prevmonth_power_profile(port, counter_identifier, divide, transform)

        for day_item in mont_data:
            for hour_item in day_item:
                print(f'{hour_item.date_time} |'
                      f'{hour_item.date_param} '
                      f'{hour_item.time_param} | '
                      f'{hour_item.a_plus} | '
                      f'{hour_item.a_minus} | '
                      f'{hour_item.r_plus} | '
                      f'{hour_item.r_minus} |')

        channel = close_channel(port, counter_identifier)

        if channel:
            print(f'Счётчик {counter_identifier} канал закрыт')

port.close()
