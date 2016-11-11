import _winreg as winreg
import math
import os
import pyscreenshot as ImageGrab
import platform
from psutil import virtual_memory
import re
import socket
import subprocess
import collections
import sys
import time
import wmi
import ctypes
import sys

"""
    TO-DO list:
    - gather Network information:
        - Driver version, date and manufacturer
        - Hardware specs
    - list of existing users
    -
"""

_ntuple_diskusage = collections.namedtuple('usage', 'total used free')

def disk_usage(path):
    _, total, free = ctypes.c_ulonglong(), ctypes.c_ulonglong(), \
                     ctypes.c_ulonglong()
    if sys.version_info >= (3,) or isinstance(path, unicode):
        fun = ctypes.windll.kernel32.GetDiskFreeSpaceExW
    else:
        fun = ctypes.windll.kernel32.GetDiskFreeSpaceExA
    ret = fun(path, ctypes.byref(_), ctypes.byref(total), ctypes.byref(free))
    if ret == 0:
        raise ctypes.WinError()
    used = total.value - free.value
    return _ntuple_diskusage(total.value, used, free.value)


def convert_filesize(input, size=0):
    sizes = ['empty', 'byte', 'Kb', 'Mb', 'Gb', 'Tb', 'Pb', 'Eb', 'Zb']
    if size >= len(sizes):
        return 'unknown file size'
    if input < math.pow(1024, size):
        return '{}{}'.format(round(input / math.pow(1024, size-1), 2), sizes[size])
    else:
        return convert_filesize(input, size+1)

def main():

    products = ['HM700', 'HM800', 'RAID700', 'RAID800', 'UCP2000', 'UCP4000', 'Other']
    for idx, val in enumerate(products):
        print '{}) {}'.format(idx + 1, val)

    while 1:
        regex_1 = r'\b[1-7]\b'
        regex_2 = r'(exit)'
        product = raw_input('Which product? type (1-{}) or type "exit": '.format(len(products)))
        if re.match(regex_1, product):
            break
        if re.match(regex_2, product):
            exit()

    dcs = ['ADC', 'CMC', 'EDC', 'ISCC']
    for idx, val in enumerate(dcs):
        print '{}) {}'.format(idx+1, val)

    while 1:
        regex_1 = r'\b[1-{}]\b'.format(len(dcs))
        regex_2 = r'(exit)'
        dc = raw_input('Please enter DC number (1-{}) or type "exit": '.format(len(dcs)))
        if re.match(regex_1, dc):
            break
        if re.match(regex_2, dc):
            exit()

    os.system('cls')

    output = []
    output.append('\n=== Environment ===')
    output.append('Product | {}'.format(products[int(product) - 1]))
    output.append('DC | {}'.format(dcs[int(dc)-1]))

    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\CentralProcessor\0',
                          'ProcessorNameString', 'Processor brand'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\CentralProcessor\0',
                          '~MHz', 'Processor MHz'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS', 'SystemManufacturer',
                          'System manufacturer'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS',
                          'SystemVersion', 'System version'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS',
                          'SystemProductName', 'System product name'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS',
                          'BIOSVendor', 'BIOS vendor'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS',
                          'BIOSVersion', 'BIOS version'))
    output.append(get_windows_reg(winreg.HKEY_LOCAL_MACHINE, r'Hardware\Description\System\BIOS',
                          'BIOSReleaseDate', 'BIOS release'))

    path = 'C:\\'
    output.append('{} | {}'.format('Python version', str(sys.version)))
    output.append('{} | {}'.format('Platform', str(platform.platform())))
    output.append('{} | {}'.format('Node', str(platform.node())))
    output.append('{} | {}'.format('Host Name', str(socket.gethostname())))
    output.append('{} | {}'.format('RAM', convert_filesize(virtual_memory().total)))
    output.append('{} {} | {}'.format('DISK Total', path, convert_filesize(disk_usage(path)[0])))
    output.append('{} {} | {}'.format('DISK Used', path, convert_filesize(disk_usage(path)[1])))
    output.append('{} {} | {}'.format('DISK Free', path, convert_filesize(disk_usage(path)[2])))
    output.append('{} | {}'.format('CD-ROM Drive', get_subprocess_command(['wmic', 'cdrom', 'get', 'drive'])))
    output.append('{} | {}'.format('whoami', get_subprocess_command(['whoami'])))


    output.append('\n=== Halcyon ===')
    output.append('{} | {}'.format('Halcyon Path', get_subprocess_command(['echo', '%HALCYONPATH%'])))
    # os.environ # treat as dict, then the key is HALCYONPATH

    # output.append('\n=== Installed Programs ===')
    # w = wmi.WMI()
    # for p in w.Win32_Product():
    #     output.append(u''.join(p.Caption).encode('utf-8').strip())

    output.append('\n=== IP Config ===')
    output.append(get_subprocess_command(['ipconfig', '-all']))

    # output.append('\n=== Installed Drivers ===')
    # output.append(get_subprocess_command(['driverquery', '-FO', 'list', '-v']))

    img = ImageGrab.grab()
    saveas = os.path.join(r"C:\capture_logs", time.strftime('%Y_%m_%d_%H_%M_%S') + '_screenshot.jpg')
    img.save(saveas)

    path = os.path.join(r"C:\capture_logs", time.strftime('%Y_%m_%d_%H_%M_%S') + '_log.txt')
    with open(path, 'w') as f:
        f.write('\n'.join(output))

    print '\n'.join(output)


def get_windows_reg(directory, path, regkey, title):
    key = winreg.OpenKey(directory, path)
    data = str(winreg.QueryValueEx(key, regkey)[0])
    if regkey is 'path':
        paths = data.split(';')
        data = ''
        for p in paths:
            data = ';\n'.join([data, p])
    winreg.CloseKey(key)
    return "{} | {}".format(title, data)


def get_subprocess_command(arguments):
    proc = subprocess.check_output(arguments, shell=True, stderr=subprocess.STDOUT)
    return str(proc)


if __name__ == "__main__":
    main()
