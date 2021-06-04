#import psch
import serial
import libscrc
from datetime import date, datetime, timedelta

def str_to_hex(s):
    # string -> hex ('6800' - > '\x68\x00')
    return bytes.fromhex(s)


def decint_hex(i):
    """
    Преобразование десятичного (int) -? hex
    """
    return'{:x}'.format(i)


def validate_4strhex_char(s):
    """
    Функция дополняет "слово"(пара байт)
    ps: Стоит обратить внимание, что если три символа, то 0 ставится в начале
        хотя может быть ситуация, когда 0 надо ставить в s[2] !!!!!!!!!!!!!!!
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

    result = ''

    crc = get_crc(cmd)
    result = cmd + crc

    return result


def send_to_port(port, counter_identifier, cmd):
    """
    Посылает запрос в com-порт.

    cmd - должен быть без crc на конце. Функция сама его подставит
    """

    fl = False

    port.flushInput()
    port.flushOutput()

    cmd = prepare_command(f'{decint_hex(counter_identifier)}{cmd}')

    hexCmd = str_to_hex(cmd)

    port.write(hexCmd)

    result = b''

    endTime = datetime.now() + timedelta(seconds=port.timeout)

    while True:
        result += port.read(port.inWaiting())

        if datetime.now() > endTime:
            break

    return result.hex()


def test_counter(port, counter_identifier):
    """
    Проверка, доступен ли счётчик
    """

    result = False
    cmd = f'00'

    try:
        r = send_to_port(port, counter_identifier, cmd)
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
    r = send_to_port(port, counter_identifier, cmd)

    etalon_ansver = str_to_hex(prepare_command(f'{decint_hex(counter_identifier)}00')).hex()

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
    r = send_to_port(port, counter_identifier, cmd)

    etalon_ansver = str_to_hex(prepare_command(f'{decint_hex(counter_identifier)}00')).hex()

    if len(r) != 0:
        if etalon_ansver == r:
            result = True

    return result


def get_current_power_profile_params(port, counter_identifier):
    """
    Функцмя получения текущего(на данный момент времени) указателя(pointer) первого (или единственного) базового массива профиля мощности

    """

    # Команда: 08h 04h
    cmd = f'0804'

    result = {
        'pointer': '',
        'half_time_param': '',
        'hour_param': '',
        'date_param': ''
    }

    r = send_to_port(port, counter_identifier, cmd)

    if len(r) == 20:
        result['pointer'] = validate_4strhex_char(r[12] + r[13] + r[14] + r[15])

        result['half_time_param'] = r[2]

        result['hour_param'] = r[4] + r[5]

        result['date_param'] = r[6] + r[7] + r[8] + r[9] + r[10] + r[11]

    return result


def get_start_day_power_profile_pointer(current_power_profile_params):
    """
    Функция находит указатель на начало !!! ТЕКУЩИХ СУТОК !!! первого (или единственного) базового массива профиля мощности
    От него в последствии можно отталкиваться, получая целые сутки прибавляя или отнимая 576(int)
    """

    result = ''

    if current_power_profile_params['pointer'] != '' and \
        current_power_profile_params['half_time_param'] != '' and \
        current_power_profile_params['hour_param'] != '':

        dec_pointer = int(current_power_profile_params['pointer'], 16)
        hour_count = int(current_power_profile_params['hour_param'])

        dec_pointer = dec_pointer - hour_count*8 - hour_count*16

        if current_power_profile_params['half_time_param'] == '8':
            dec_pointer -= 8
        else:
            dec_pointer -= 16

        result = validate_4strhex_char(decint_hex(dec_pointer))

    return str(result)


def get_start_day_ago_power_profile_pointer(self, dayCount):
    """
    Функция находит указатель на начало дней/суток (dayCount) тому назад первого (или единственного) базового массива профиля мощности
    """

    result = ''

    dayStart = validate_4strhex_char(self.startDayPowerProfilePointer)

    data = int(dayStart, 16) - (576 * dayCount)
    self.startDayPowerProfilePointer = decint_hex(data)

    self.startDayPowerProfilePointer = validate_4strhex_char(self.startDayPowerProfilePointer)

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

        current_power_profile_params = get_current_power_profile_params(port, counter_identifier)
        print(current_power_profile_params)

        start_day_power_profile_pointer = get_start_day_power_profile_pointer(current_power_profile_params)
        print(f'Указатель на начало текущих суток: {start_day_power_profile_pointer}')

        channel = close_channel(port, counter_identifier)
        if channel:
            print(f'Счётчик {counter_identifier} канал закрыт')

port.close()