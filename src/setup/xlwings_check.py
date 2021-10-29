#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   xlwings_check.py
@Author  :   Billy Zhou
@Time    :   2021/10/28
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import os
import shlex
import xlwings as xw
from subprocess import check_output
from src.setup.conda_check import find_conda_path


def install_xlwings_addin(xlwings_filename: str, xlstart_path: Path, addins_path: Path, force_to_shutdown: bool = False) -> Path:
    """install xlwings addins and move xlwings.xlam to addins folder"""
    # check_excel_status
    if xw.apps.count:
        if force_to_shutdown:
            for app in xw.apps:
                app.kill()
        else:
            raise PermissionError('Excel applications are running. Please shutdown excel and try again')

    # check xlwings.xlam file
    if addins_path.joinpath(xlwings_filename).exists():
        addins_path.joinpath(xlwings_filename).unlink()
    if xlstart_path.joinpath(xlwings_filename).exists():
        xlstart_path.joinpath(xlwings_filename).unlink()
    # install
    for line in check_output(shlex.split('conda.bat activate pyexcel && xlwings addin install')).decode('utf-8').split('\n'):
        if 'Successfully installed the xlwings add-in! Please restart Excel.' in line:
            print('xlwings addins is installed')
    xlstart_path.joinpath(xlwings_filename).rename(addins_path.joinpath(xlwings_filename))
    print(f'{xlwings_filename} is moved from {xlstart_path} to {addins_path}')
    return addins_path.joinpath(xlwings_filename)


def check_xlwings_config(
        config_filepath: Path, conda_basepath: Path,
        force_to_create: bool = False, force_to_update: bool = False):
    """check and update the variables in xlwings configs"""
    # init
    status = {
        '"INTERPRETER",': False,
        '"PYTHONPATH",': False,
        '"UDF MODULES",': False,
    }

    # create if config_file not exists
    if not Path(config_filepath).exists():
        check_output(shlex.split('conda.bat activate pyexcel && xlwings config create')).decode('utf-8').split('\n')
        print('xlwings config created')
    elif force_to_create:
        check_output(shlex.split('conda.bat activate pyexcel && xlwings config create --force')).decode('utf-8').split('\n')
        print('xlwings config created')

    # check and update the main variables status
    with open(Path(config_filepath), 'r+') as f:
        lines = f.readlines()
        pop_list = []
        for line in lines:
            for key in status.keys():
                if key in line:
                    if force_to_update:
                        pop_list.append(line)
                    else:
                        status[key] = True
        for i in pop_list:
            lines.pop(lines.index(i))
        # add if not exists
        for key in status.keys():
            if not status[key]:
                if key == '"INTERPRETER",':
                    lines.append('{}"{}"\n'.format(key, Path(
                        conda_basepath).joinpath(r'envs\pyexcel\pythonw.exe')))
                elif key == '"PYTHONPATH",':
                    lines.append('{}"{}"\n'.format(key, cwdPath.joinpath('bin')))
                elif key == '"UDF MODULES",':
                    lines.append('{}"{}"\n'.format(key, 'main'))
        f.seek(0)
        f.writelines(lines)

        print('xlwings config updated')
        lines_print = ''.join(lines)
        print(f'The config of xlwings now are: \n{lines_print}')


if __name__ == '__main__':
    # init
    usr_folder = Path(r'C:\Users\{}'.format(os.getlogin()))
    xlwings_filename = 'xlwings.xlam'
    xlstart_path = usr_folder.joinpath(r'AppData\Roaming\Microsoft\Excel\XLSTART')
    addins_path = usr_folder.joinpath(r'AppData\Roaming\Microsoft\AddIns')
    xlstart_path.mkdir(parents=True, exist_ok=True)
    addins_path.mkdir(parents=True, exist_ok=True)

    if False:
        install_xlwings_addin(xlwings_filename, xlstart_path, addins_path, force_to_shutdown=False)

    if False:
        conda_basepath = find_conda_path(folder_keyword='conda')
        check_xlwings_config(
            usr_folder.joinpath('.xlwings').joinpath('xlwings.conf'), conda_basepath,
            force_to_create=False, force_to_update=True)
