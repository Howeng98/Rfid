#!/usr/bin/python3
import os
import sys
import time
import json
import socket
import signal
import logging
import schedule
import datetime
import _thread as thread
from happig_database import Database

DATABASE = None
connection = 0
IP_setting = {}
MYSQL_setting = {}
POSTGRE_setting = {}
server = None

BUFFER_SIZE = 1024


def broadcast_send(server, msg):
    ''' Send via UDP '''
    server.sendto(msg, ('<broadcast>', 1501))


def get_setting_data():
    setting_parameter = []
    dir_path = os.path.dirname(os.path.realpath(__file__))
    with open(os.path.join(dir_path, 'setting.json')) as f:
        setting_parameter = json.load(f)
    global IP_setting, POSTGRE_setting, MYSQL_setting
    IP_setting = setting_parameter["IP"]
    MYSQL_setting = setting_parameter["MYSQL"]
    POSTGRE_setting = setting_parameter["POSTGRE"]


def receive(server, myaddr):
    ''' Return RFID devices '''
    devicelist = []
    while True:
        try:
            data, addr = server.recvfrom(1024)
        except socket.error as e:
            print(e)
            break
        if addr[0] != myaddr:  # Skip server
            devicelist.append((data[:6].hex().upper(), addr[0]))
    return devicelist


def searchdevice(server):
    ''' Broadcast via UDP, only RFID devices and server will response'''
    msg = b'Read Parameter-12345678901234567890123456789012.'
    broadcast_send(server, msg)


def select_device(server, mac):
    ''' Send the same command when double click the device in Wenshing Windows app '''
    macdash = mac2macdash(mac)
    cmd = bytearray('Read Http-12345678901234567890%s.' % (macdash), 'ascii')
    broadcast_send(server, cmd)
    cmd = bytearray('Read Dns Website1234567890123:%s.' % (macdash), 'ascii')
    broadcast_send(server, cmd)


def mac2macdash(mac):
    chunk_size = 2
    mac = [mac[i:i+chunk_size] for i in range(0, len(mac), chunk_size)]
    return '-'.join(mac)


def setup_via_net(server, mac, device_ip):
    ''' Send the same command when double click the device in Wenshing Windows application '''
    subnet_mask = [255, 255, 255, 0]
    device_ip = list(map(int, device_ip.split('.')))
    default_gateway = device_ip[:]
    default_gateway[3] = 1
    device_port = int(IP_setting["DevicePort"])
    destination_ip = list(map(int, get_ip_address().split('.')))
    destination_port = int(IP_setting["DestinationPort"])
    baud_rate = int(IP_setting["BaudRate"])
    delay_send = 50  # ms
    # TODO let device ip can be changed, and let defaultgate use its ip first 3 number
    hexs = mac + mac + intArray2hex(device_ip) + int2hex(device_port, 2) + intArray2hex(destination_ip) + int2hex(destination_port, 2) + intArray2hex(
        default_gateway) + intArray2hex(subnet_mask) + ' 01 ' + int2hex(baud_rate, 3) + ' 00 00 00 ' + int2hex(delay_send) + ' 00 01 00 be 0f 20 b0'
    bs = bytearray.fromhex(hexs)
    sum_ = 213
    for b in bs:
        sum_ += b
    bs = bs + bytes([sum_ % 256])
    broadcast_send(server, bs)


def int2hex(val, num_bytes=1):
    intArray = []
    for _ in range(num_bytes):
        intArray.append(val % 256)
        val = val // 256
    intArray = reversed(intArray)
    return intArray2hex(intArray)


def get_ip_address():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    return s.getsockname()[0]

def get_devices(server):
    ''' Return RFID devices '''
    myaddr = get_ip_address()
    searchdevice(server)
    devices = receive(server, myaddr)
    return devices

def set_and_get_devices(server):
    ''' Set all RFID devices to listen to server '''
    devices = get_devices(server)
    print('-' * 50)
    for mac, addr in devices:
        print("set device: mac = %s, address = %s" % (mac2macdash(mac), addr))
        select_device(server, mac)
        setup_via_net(server, mac, addr)
    # device_idx = int(input("Select device: "))
    # try:
    #   mac, addr = devices[device_idx]
    #   print('Device', mac, 'is selected')
    # except Exception as e:
    #   print(e)
    #   return
    # select_device(server, mac)
    # setup_via_net(server, mac, addr)
    return devices


def hex2int(h):
    return int(h, 16)


def intArray2hex(ints):
    return ' '.join([hex(i)[2:].zfill(2) for i in ints])


def read_code():
    ''' Keep scanning rfid tag, send ctrl-c to terminate '''
    keep_scanning = [True]

    def signal_handler(sig, frame):
        print('received ctrl-c')
        keep_scanning[0] = False

    sigint_handler_old = signal.signal(signal.SIGINT, signal_handler)
    while keep_scanning[0]:
        try:
            data = connection.recv(BUFFER_SIZE)
        except Exception as e:
            print(e)
            continue
        print('received "%s"' % data)
        scan_decode(data)
    signal.signal(signal.SIGINT, sigint_handler_old)


def scan_decode(val):
    epcs = []
    try:
        while len(val) > 0:
            if val[0] == hex2int('02') and val[1] == hex2int('53'):
                data_len = val[2] + 2
                tid = intArray2hex(val[3:data_len])
                print('Tag TID:', tid,'\n\n')
                # TODO
            elif val[0] == hex2int('02') and val[1] == hex2int('54'):
                data_len = val[2] + 2
                # print('RSSI value being received: ', val[3])
                # print(val[4])
                # print(val[5])
                # print(val[6])
                freq = val[4] + val[5] * 256
                print(freq)
                if val[6] >= 0x80:
                    freq += 0x0E * 256 * 256
                else:
                    freq += 0x0D * 256 * 256
                # print('scan frequency %d kHz' % (freq), end='\t\t')
                epc = intArray2hex(val[10:data_len])
                print('Tag EPC:', epc,'\n\n')
                epcs.append(epc)
                val = val[data_len+2:]
            else:
                print("Unexpected tag !!!!!!:")
                break
    except Exception as e:
        print(e)
    return epcs


def print_commands():
    print('###########################################')
    commands = ["Scan", "Change Power", "Change IP & Port", "AT Comment"]
    print("Select command")
    for idx, command in enumerate(commands):
        print('    %d %s' % (idx+1, command + ' command'))
    print('')


def read_selection():
    s = input('Select: ')
    return s


def change_power():
    send('AT+0001-SetPower:30dB\x0d\x0a')


def change_ip():
    pass


def send(send_string):
    print('Send: ' + send_string, '')
    try:
        connection.sendall(str.encode(send_string))
        data = connection.recv(1024)
        print('Recv: ', data)
    except Exception as e:
        print(e)


def stop_scan():
    send('AT+0001-Scan:0\x0d\x0a')


def start_scan():
    send('AT+0001-Scan:1\x0d\x0a')

connection.sendall
def At_comment():
    cm = input('AT Comment: ')
    cm += '\x0d\x0a'
    send(cm)


def thread_func(thread_idx, sock, mac_addr_pairs):
    ''' Open a thread to listen to RFID device and save all received data to database '''
    thread_name = 'Thread %d' % thread_idx
    print(thread_name, 'is listening')
    sock.listen(1)
    connection, cli_addr = sock.accept()
    thread_mac = 'NO'
    for mac, addr in mac_addr_pairs:
        if cli_addr[0] == addr:
            thread_mac = mac
    print(thread_name, 'connected to', cli_addr, ', mac =', thread_mac)    
    while True:
        tags = []
        try:
            data = connection.recv(BUFFER_SIZE)
            print(thread_name, 'recv:', str(data))
            tags = scan_decode(data)
        except ConnectionResetError as e:
            # Reconnect
            print(e)
            print('reconnecting...')
            reconnect = False
            while not reconnect:
                global server
                devices = get_devices(server)
                for device in devices:
                    mac, addr = device
                    if mac == thread_mac:
                        select_device(server, mac)
                        setup_via_net(server, mac, addr)
                        sock.listen(1)
                        try:
                            connection, cli_addr = sock.accept()
                            reconnect = True
                            break
                        except Exception as e:
                            print(e)
                time.sleep(1)
        except Exception as e:
            print(e)
        time_now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        (shed_id, sensor_id, tag_id, read_time)
        ##for tag in tags:
        ##    DATABASE.insert_tag(3, mac2macdash(thread_mac), tag, time_now)    

def iniatialize():
    get_setting_data()
    dir_path = os.path.dirname(os.path.realpath(__file__))
    save_path = os.path.join(dir_path, "DB_log")
    global DATABASE
    DATABASE = Database(save_path, MYSQL_setting, POSTGRE_setting)


def main():
    iniatialize()

    global server
    server = socket.socket(
        socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    server.settimeout(0.2)
    myaddr = get_ip_address()
    server.bind((myaddr, 1501))

    devices = set_and_get_devices(server)
    print(devices)
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    ip_address = get_ip_address()
    try:
        port = int(IP_setting['DestinationPort'])
    except Exception as e:
        port = 8080  # Default port
        print(e)
    server_address = (ip_address, port)
    print('starting up on %s port %s' % server_address)
    sock.bind(server_address)

    for idx in range(len(devices)):
        thread.start_new_thread(thread_func, (idx, sock, devices))

    while True:
        schedule.run_pending()  # TODO No need to use schedule: modify happig_log.py, create directory if it's not exist when saving log file
        time.sleep(3)

    ############### For debug only (check whether the code can work) ##################
    # Receive the data in small chunks and retransmit it
    while True:
        try:
            print_commands()
            s = read_selection()
            stop_scan()
            if s == '1':
                start_scan()
                read_code()
            elif s == '2':
                change_power()
            elif s == '3':
                change_ip()
            elif s == '4':
                At_comment()
            else:
                print(s)
                start_scan()
                break
            start_scan()
        except Exception as e:
            print(e)
    connection.close()


if __name__ == '__main__':
    main()
