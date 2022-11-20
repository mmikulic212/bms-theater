#!/usr/bin/env python

import sys
import winreg
import re

PROGRAM_VERSION = "1"
PROGRAM_DATE = "2022/11/20"

BMS_VER = "Falcon BMS 4.36"
BMS_REG_KEY = "SOFTWARE\\WOW6432Node\\Benchmark Sims"
BMS_THEATER_LIST_FILE = "\\Data\\Terrdata\\theaterdefinition\\theater.lst"

def run():
    print(f' Falcon BMS Theater Changer v{PROGRAM_VERSION} ({PROGRAM_DATE})')
    bms_key = open_reg_key(winreg.HKEY_LOCAL_MACHINE, BMS_REG_KEY )
    bms_versions = get_reg_subkeys(bms_key)
    bms_version = ''
    print('Found Falcon BMS version:')
    for i,vers in enumerate(bms_versions, start=1):
        print(f' {i}: {vers}')
    while True:
        try:
            choosed = int(input(f'Choose Falcon BMS version (1-{i}):'))
            bms_version = bms_versions[choosed-1]
            print(bms_version)
        except (TypeError, ValueError, IndexError):
            continue
        break
    bms_key.Close()
    
    bms_key = open_reg_key(winreg.HKEY_LOCAL_MACHINE, BMS_REG_KEY+'\\'+bms_version)
    bms_basedir = get_reg_value(bms_key,"baseDir")[0]
    print(bms_basedir)
    bms_curtheater = get_reg_value(bms_key,"curTheater")[0]
    bms_key.Close()
    print(bms_curtheater)
    theater_names = []
    with open(bms_basedir+BMS_THEATER_LIST_FILE) as theater_list_file:
        print(f'Found theaters for {bms_version}:')
        for theater_list_line in theater_list_file:
            try:
                with open(bms_basedir+"\\Data\\"+theater_list_line.rstrip()) as tdf_file:
                    for tdf_line in tdf_file.readlines():
                        m = re.search(r"name (.*)",tdf_line)
                        if m:
                            theater_names.append(m.group(1))
            except FileNotFoundError:
                continue

    for i,tea in enumerate(theater_names, start=1):
        if bms_curtheater in tea:
            print(f' {i}: {tea} <<-- current')
        else:
            print(f' {i}: {tea}')

    chosen_theater = bms_curtheater
    while True:
        try:
            choosed = int(input(f'Choose Theater for {bms_version} (1-{i}):'))
            chosen_theater = theater_names[choosed-1]
        except (TypeError, ValueError, IndexError):
            continue
        break
    
    bms_key = open_reg_key(winreg.HKEY_LOCAL_MACHINE, BMS_REG_KEY+'\\'+bms_version, winreg.KEY_ALL_ACCESS)
    set_reg_value(bms_key,"curTheater",chosen_theater)
    bms_key.Close()
    print(f"{chosen_theater} was set in registry for {bms_version} as current theater.")


def open_reg_key(root_reg:str, key:str, access:int = winreg.KEY_READ) -> winreg.HKEYType:
    root_key = winreg.ConnectRegistry(None, root_reg)
    open_key = winreg.OpenKey(root_key, f"{key}\\", 0, access)
    return open_key

def close_reg_key(open_key:winreg.HKEYType) -> None:
    open_key.Close()

def get_reg_subkeys(open_key:winreg.HKEYType) -> list:

    subkeys = []
    i = 0
    while True:
        try:
            subkeys.append(winreg.EnumKey(open_key,i))
            i+=1
        except OSError:
            break
    return subkeys

def get_reg_value(open_key:winreg.HKEYType, value_name:str) -> str:
    return winreg.QueryValueEx(open_key, value_name)

def set_reg_value(open_key:winreg.HKEYType, value_name:str, value:str) -> None:
    winreg.SetValueEx(open_key, value_name, 0, winreg.REG_SZ, value)
    return

if __name__ == '__main__':
    try:
        run()
    except KeyboardInterrupt:
        exit("Quitting...")
