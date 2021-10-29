#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   xlwings_install.py
@Author  :   Billy Zhou
@Time    :   2021/10/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[0]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import os
import xlwings as xw
from subprocess import check_output
from src.setup.conda_check import find_conda_path
from src.setup.conda_check import add_conda_path
from src.setup.conda_check import check_conda_bat
from src.setup.conda_check import create_conda_env
from src.setup.xlwings_check import install_xlwings_addin
from src.setup.xlwings_check import check_xlwings_config


if __name__ == '__main__':
    # init
    usr_folder = Path(r'C:\Users\{}'.format(os.getlogin()))
    xlwings_filename = 'xlwings.xlam'
    xlstart_path = usr_folder.joinpath(r'AppData\Roaming\Microsoft\Excel\XLSTART')
    addins_path = usr_folder.joinpath(r'AppData\Roaming\Microsoft\AddIns')
    xlstart_path.mkdir(parents=True, exist_ok=True)
    addins_path.mkdir(parents=True, exist_ok=True)

    # find the basepath of conda
    conda_basepath = find_conda_path(folder_keyword='conda')

    # check and add conda_basepath to sys path
    add_conda_path(conda_basepath)

    # check conda.bat and update conda environment variables
    check_conda_bat(conda_basepath, add_path=True)

    # check and create conda env named pyexcel
    create_conda_env('pyexcel', cwdPath.joinpath('requirements_conda_pyexcel_win.yaml'))

    # check and add pyexcel to sys path
    add_conda_path(conda_basepath.joinpath('envs\\pyexcel'))

    # check and install xlwings addins
    install_xlwings_addin(xlwings_filename, xlstart_path, addins_path, force_to_shutdown=False)

    # check and update xlwings configs
    check_xlwings_config(usr_folder.joinpath('.xlwings').joinpath('xlwings.conf'), conda_basepath, force_to_update=True)

    # check and create PERSONAL.xlsb
    if not xlstart_path.joinpath('PERSONAL.xlsb').exists():
        wb = xw.Book()
        wb.save(str(xlstart_path.joinpath('PERSONAL.xlsb')))
        wb.app.kill()
        print('PERSONAL.xlsb created')
    try:
        print('open the XLSTART path. Please import udf_scripts in PERSONAL.xlsb')
        check_output('start %windir%\\explorer.exe {}'.format(xlstart_path), shell=True)
    except Exception as e:
        log.error('An error occurred. {}'.format(e.args))
