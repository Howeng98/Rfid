import hid
import signal

def hex2int(h):
  return int(h, 16)

def intArray2hex(ints):
  return ' '.join([hex(i)[2:].zfill(2) for i in ints])

def intArray2int(ints):
  val = 0
  for pwr, i in enumerate(ints):
    val += i*(16**pwr)
  return val

def int2hex(val, num_bytes=1):
  intArray = []
  for _ in range(num_bytes):
    intArray.append(val % 256)
    val = val // 256
  return intArray2hex(intArray)

def test_command_sd():
  send = bytearray.fromhex('07 04 03 03 01 04')
  return send

def scan_command_sd():
  send = bytearray.fromhex('07 11 00 86 00 02 00 00 00 0D 8C 00 05 00 00 01 01 00 01 06')
  return send

def stop_command_sd():
  send = bytearray.fromhex('08 0A 00 8C 00 05 00 00 01 00 00 00 00')
  return send

def RBA_command_sd():

  send = bytearray.fromhex('81')
  tag_epc = input('EPC of TAG:')
  epc_len = len(tag_epc.replace(' ', ''))//2
  byte_after3 = 14 + epc_len
  send = send + bytearray.fromhex(intArray2hex([byte_after3]))
  send = send + bytearray.fromhex('00 06 00')
  byte_after8 = 9 + epc_len
  send = send + bytearray.fromhex(intArray2hex([byte_after8]))
  send = send + bytearray.fromhex('00 00 02 00 00 01 20 00 60 00')
  send = send + bytearray.fromhex(tag_epc)
  send = send + bytearray.fromhex('AA')

  print(send)
  Hid.write(send)
  ret_data = read_data()
  print(ret_data)

  send = bytearray.fromhex('82')
  pwd = input('Access Password of TAG:')
  pwd_len = len(pwd.replace(' ', ''))//2
  byte_after3 = 10 + pwd_len
  send = send + bytearray.fromhex(intArray2hex([byte_after3]))
  send = send + bytearray.fromhex('00 08 00')
  byte_after9 = 5 + pwd_len
  send = send + bytearray.fromhex(intArray2hex([byte_after9]))
  send = send + bytearray.fromhex('00 81')
  bank_area = input('Bank Area: 00 = Reserverd / 01 = EPC / 02 = TID / 03 = User :')
  send = send + bytearray.fromhex(bank_area)
  start_pos = input('Start position of Area:')
  send = send + bytearray.fromhex(int2hex(int(start_pos), 3))
  read_num = input('How many byte to read:')
  send = send + bytearray.fromhex(int2hex(int(read_num)))
  send = send + bytearray.fromhex(pwd)
  return send

def WBA_command_sd():

  send = bytearray.fromhex('81')
  tag_epc = input('EPC of TAG:')
  epc_len = len(tag_epc.replace(' ', ''))//2
  byte_after3 = 14 + epc_len
  send = send + bytearray.fromhex(intArray2hex([byte_after3]))
  send = send + bytearray.fromhex('00 06 00')
  byte_after8 = 9 + epc_len
  send = send + bytearray.fromhex(intArray2hex([byte_after8]))
  send = send + bytearray.fromhex('00 00 02 00 00 01 20 00 60 00')
  send = send + bytearray.fromhex(tag_epc)
  send = send + bytearray.fromhex('AA')

  print(send)
  Hid.write(send)
  ret_data = read_data()
  print(ret_data)

  send = bytearray.fromhex('82')
  wt_data = input('Write data to TAG (word):')
  wt_len = len(wt_data.replace(' ', ''))//2
  bank_area = input('Bank Area: 00 = Reserverd / 01 = EPC / 02 = TID / 03 = User :')
  if bank_area == '01':
    byte_after3 = 14 + wt_len + 2
    byte_after8 = 9 + wt_len + 2
  else:
    byte_after3 = 14 + wt_len
    byte_after8 = 9 + wt_len
  send = send + bytearray.fromhex(intArray2hex([byte_after3]))
  send = send + bytearray.fromhex('00 07 00')
  send = send + bytearray.fromhex(intArray2hex([byte_after8]))
  send = send + bytearray.fromhex('00 02')
  send = send + bytearray.fromhex(bank_area)

  if bank_area == '01':
    modify_num = 0x30 + (wt_len-12)*4
    send = send + bytearray.fromhex('01 00 00 00')
  else:
    start_pos = input('Start position of Area:')
    send = send + bytearray.fromhex(int2hex(int(start_pos), 4))
  pwd = input('Access Password of TAG:')
  send = send + bytearray.fromhex(pwd)
  send = send + bytearray.fromhex(int2hex(modify_num, 2))
  send = send + bytearray.fromhex(wt_data)
  print(send)
  return send

def modify_power():
  send = bytearray.fromhex('C0 0A 00 69 00 05 00 04 02 00')
  power = input('Power (11~1A) :')
  send = send + bytearray.fromhex(power)
  send = send + bytearray.fromhex('03 00')
  return send

def scan_command_send():
  cmd = bytearray.fromhex('07 11 00 86 00 02 00 00 00 0D 8C 00 05 00 00 01 01 00 01 06')
  Hid.write(cmd)

  keep_scanning = [True]
  def signal_handler(sig, frame):
    print(' received ctrl-c')
    keep_scanning[0] = False

  sigint_handler_old = signal.signal(signal.SIGINT, signal_handler)
  while keep_scanning[0]:
    scan_decode(read_data())

  signal.signal(signal.SIGINT, sigint_handler_old)
  stop_commamd_send()

def stop_commamd_send():
  cmd = bytearray.fromhex('08 0A 00 8C 00 05 00 00 01 00 00 00 00')
  Hid.write(cmd)
  s = read_data() # TODO check if valid
  print('Stop scanning')


def scan_decode(val):
  if val[11] == hex2int('AA'):
    print('old tag', end='\t\t')
  elif val[11] == hex2int('80'):
    print('tag >= 64 byte', end='\t\t')
  else:
    print('new tag', end='\t\t')

  try:
    print('-%d dBm' % ((16**2) - val[12]), end='\t\t')
    print('scan frequency %d kHz' % (intArray2int(val[13:16])), end='\t\t')
    print('TAG:', intArray2hex(val[17:19]), end='\t\t')
    print('EPC:', intArray2hex(val[19:-1]))
  except Exception as e:
    print(e)


def print_commands():
  commands = ["Test", "Scan", "Stop", "Read Bank Area", "Write Bank Area",
      "Unlock Lock & Prelock", "Kill Tag", "Set", "Change Power"]
  print("Select command")
  for idx, command in enumerate(commands):
    print('    %d %s' % (idx+1, command + ' command'))

def read_selection():
  s = input('Select: ')
  return s

Hid = None

def main():
  VENDOR_ID = 0x2202
  PRODUCT_ID = 0x1007
  data = b'\x07\x04\x03\x03\x01\x04'
  data = bytearray.fromhex('07 11 00 86 00 02 00 00 00 0D 8C 00 05 00 00 01 01 00 01 06')

  global Hid
  Hid = hid.device(VENDOR_ID, PRODUCT_ID)
  Hid.open(VENDOR_ID, PRODUCT_ID)
  print(Hid.get_manufacturer_string())
  print(Hid.get_product_string())
#######################
  while True:
    print_commands()
    s = read_selection()
    print(s)
    if s == '2':
      scan_command_send()
    elif s == '4':
      ret = RBA_command_sd()
      Hid.write(ret)
      print(read_data())
      ret = RBA_command_sd()
      Hid.write(ret)
      print(read_data())
    elif s == '6':
      lock_command_send()
    elif s == '9':
      ret = modify_power()
      Hid.write(ret)
      print(intArray2hex(read_data()))
    else:
      print(len(s))
      print(s)
      break


  print("Closing the device")
  Hid.close()

def read_data():
  data = Hid.read(64)
  data_num = data[1]
  return data[:data_num+2]

if __name__ == '__main__':
  main()
