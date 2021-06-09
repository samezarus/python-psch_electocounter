#import psch
import serial
import libscrc
from datetime import date, datetime, timedelta
from time import sleep


class PowerProfileItem:
    """
    Элемент профиля мощности
    """
    def __init__(self):
        self.a_plus = 0.0  # A+ кВт
        self.a_minus = 0.0  # A- кВт
        self.r_plus = 0.0  # R+ кВт
        self.r_minus = 0.0  # R- кВт
        self.date_param = ''  # Дата снятия (ддммгг)
        self.doubleTimeParam = ':-:'  # Временной промежуток снятия (21:00-21:30 ...)


def str_to_hex(s):
    # string -> hex ('6800' - > '\x68\x00')

    return bytes.fromhex(s)


def int_to_hex_str(i):
    """
    i: 0->255

    1 dec -> 01 hex, 20 dec -> 14 hex
    """

    #return '{:x}'.format(i)
    return (f'{i:0>2X}')


def validate_strhex(s):
    """
    Функция дополняет "слово"(пара байт)
    ps: Стоит обратить внимание, что если три символа, то 0 ставится в начале
    """
    result = s

    if len(s) == 3:
        result = f'0{s}'

    if len(s) == 2:
        #result = f'0{s[0]}0{s[1]}'
        result = f'00{s}'

    if len(s) == 1:
        result = f'000{s[0]}'

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
    sHex = str(hex(crc16))

    if len(sHex) == 6:
        result = sHex[4] + sHex[5] + sHex[2] + sHex[3]

    if len(sHex) == 5:
        result = f'{sHex[3]}{sHex[4]}0{sHex[2]}'

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


def send_to_port(port, counter_identifier, cmd, cmd_print):
    """
    Посылает запрос в com-порт.

    cmd - должен быть без crc на конце. Функция сама его подставит
    """

    cmd_print = False

    port.flushInput()
    port.flushOutput()

    cmd = prepare_command(f'{int_to_hex_str(counter_identifier)}{cmd}')

    hexCmd = str_to_hex(cmd)

    if cmd_print:
        print(f'TX    {hexCmd.hex()}')

    port.write(hexCmd)

    result = b''

    endTime = datetime.now() + timedelta(seconds=port.timeout)

    while True:
        result += port.read(port.inWaiting())

        if datetime.now() > endTime:
            break
    if cmd_print:
        print(f'RX    {result.hex()}')
        print('')

    return result.hex()


def test_counter(port, counter_identifier):
    """
    Проверка, доступен ли счётчик
    """

    result = False
    cmd = f'00'

    try:
        r = send_to_port(port, counter_identifier, cmd, True)
        if len(r) > 0:
            result = True
    except:
        result = None

    return result


def open_channel(port, counter_identifier, counter_password):
    """
    Открытие канала связи со счётчиком
    """

    result = False

    hexPassword = counter_password.encode().hex()

    cmd = f'01{hexPassword}'
    r = send_to_port(port, counter_identifier, cmd, True)

    etalon_ansver = str_to_hex(prepare_command(f'{int_to_hex_str(counter_identifier)}00')).hex()

    if len(r) != 0:
        if etalon_ansver == r:
            result = True

    return result


def close_channel(port, counter_identifier):
    """
    Открытие канала связи со счётчиком
    """

    result = False

    cmd = f'02'
    r = send_to_port(port, counter_identifier, cmd, True)

    etalon_ansver = str_to_hex(prepare_command(f'{int_to_hex_str(counter_identifier)}00')).hex()

    if len(r) != 0:
        if etalon_ansver == r:
            result = True

    return result


def read_power_profile_pointer_on_date(port, counter_identifier, date):
    """
    Поиск первого (или единственного) указателя базового массива профиля мощности
    на заданную дату
    """

    cmd = f'0803'
    r = send_to_port(port, counter_identifier, cmd, True)

    cmd = f'0809'
    r = send_to_port(port, counter_identifier, cmd, True)

    cmd = f'0806'
    r = send_to_port(port, counter_identifier, cmd, True)

    cmd = f'0804'
    r = send_to_port(port, counter_identifier, cmd, True)

    cmd = f'032800FFFFFF{date}FF1E'
    r = send_to_port(port, counter_identifier, cmd, True)

    # Найти указатель базаового массива профиля мощности на начало искомой даты
    #cmd = f'081800'
    #r = send_to_port(port, counter_identifier, cmd, True)

    p = '1'
    while p != '0':
        cmd = f'081800'
        r = send_to_port(port, counter_identifier, cmd, True)
        p = r[3]

    r = r[8] + r[9] + r[10] + r[11]

    return r


def read_7bit_header(port, counter_identifier, date, pointer):
    """
    Прочитать 7 байт информации (заголовок профиля) из памяти № 03h c адреса "pointer"
    """
    result = True

    cmd = f'0603{pointer}07'
    r = send_to_port(port, counter_identifier, cmd, True)

    part = f'00{date}011E'

    if r.find(part) != -1:
        result = False

    return result


def read_transformation_coefficient(port, counter_identifier):
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
    r = send_to_port(port, counter_identifier, cmd, True)

    if len(r) == 26:
        result['kn'] = int(f'{r[2]}{r[3]}{r[4]}{r[5]}', 16)
        result['kt'] = int(f'{r[6]}{r[7]}{r[8]}{r[9]}', 16)
        result['dimensionality'] = int(f'{r[10]}{r[11]}', 16)
        result['whole_part'] = int(f'{r[12]}{r[13]}', 16)
        result['fractional_part'] = int(f'{r[14]}{r[15]}{r[16]}{r[17]}{r[18]}{r[19]}{r[20]}{r[21]}', 16)

    return result


def read_power_profile_line(port, counter_identifier, index, pointer):
    """
    Прочитать первую или очередную строку с данными профиля мощности

    index не должен быть равным 0 (проблемы CRC). Только 1 -> 255
    """

    ma = '03'  # № адреса памяти
    bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

    cmd = f'0C{index}{ma}{pointer}{bytes_count}'
    r = send_to_port(port, counter_identifier, cmd, True)

    # Отсекаем номер счётчика и индекс, отсекаем CRC
    return r[4:-4]


def create_power_profile_item(hhx, date, divide):
    ppi = PowerProfileItem()

    if len(hhx) == 16:
        ppi.a_plus = int(hhx[0:4], 16) / divide
        ppi.a_minus = int(hhx[4:8], 16) / divide
        ppi.r_plus = int(hhx[8:12], 16) / divide
        ppi.r_minus = int(hhx[12:16], 16) / divide
        ppi.date_param = date
    else:
        print(f'{hhx} -- Не удалось получить данные')

    return ppi


def read_power_profile(port, counter_identifier, pointer, date, divide):
    """
    Прочитать все строки профиля мощности на дату
    pointer -  значение из функции read_power_profile_pointer_on_date()
    """

    pointer = validate_strhex(pointer)

    result = []

    fl = False  # Флаг удачного поиска 24-х пар получасовок

    bytes_count = '82'  # Количество байт для считывания (82 в проприетарной утилите)

    data = ''

    # Циклично пытаемся найти
    for i in range(1, 255):
        index = int_to_hex_str(i)

        data += read_power_profile_line(port, counter_identifier, index, pointer)

        dec_pointer = int(pointer, 16) + int(bytes_count, 16)

        if dec_pointer < 65535:
            pointer = int_to_hex_str(dec_pointer)
            pointer = validate_strhex(pointer)
        else:
            pointer = '0000'

        # Точно прерываем цикл, так уже на всякий случай получили лишнюю строку ответа
        if fl:
            break

        # Если нашли признак 24-й пары получасовок
        # то не выходим, а даём получить ещё строку ответа во избежание
        # ситуации, когда признак находится в конце строки текущего ответа
        if data.find(f'23{date}') != -1:
            fl = True

    #print(data)

    # Если успешно нашли 24-и пары получасовок
    if fl:
        for i in range(0, 24):
            if i > 9:
                h = f'{i}{date}'
            else:
                h = f'0{i}{date}'


            # Сложно рассказать зачем это )
            repeat_count = data.count(h)
            if data.count(h) > 1:
                s = ''
                for j in range(repeat_count):
                    s += h
                h = s

            l = data.split(h)

            # Строка очищенная от "мусора", в ней данные двух получасовок к примеру
            # 00:00->00:30 и 00:30->01:00
            hh = l[1][8:40]

            hh1 = hh[0:16]
            hh2 = hh[16:32]

            item1 = create_power_profile_item(hh1, date, divide)
            item2 = create_power_profile_item(hh2, date, divide)

            result.append(item1)
            result.append(item2)

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

online = test_counter(port, counter_identifier)
if online:
    print(f'Счётчик {counter_identifier} доступен')

    channel = open_channel(port, counter_identifier, '000000')
    if channel:
        print(f'Счётчик {counter_identifier} канал открыт')

        date = '190521'

        pointer = read_power_profile_pointer_on_date(port, counter_identifier, date)
        print(f'{pointer} -- указатель на дату {date}')

        if read_7bit_header(port, counter_identifier, date, pointer):
            tc = read_transformation_coefficient(port, counter_identifier)
            print(tc)

            divide = 1250 # Делитель в зависимости от модели счётчика
            ppl = read_power_profile(port, counter_identifier, pointer, date, divide)

            for item in ppl:
                print(f'{item.a_plus} - {item.a_minus} - {item.r_plus} - {item.r_minus}')


        channel = close_channel(port, counter_identifier)

        if channel:
            print(f'Счётчик {counter_identifier} канал закрыт')

port.close()
