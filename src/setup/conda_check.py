#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   conda_check.py
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

import time
import yaml
import shlex
import shutil
from subprocess import check_output
from src.utils.setup import proc_run
from src.utils.input_check import input_checking_YN


def find_conda_path(folder_keyword: str = 'conda') -> Path:
    """find conda path in the path of current interpreter"""
    conda_basepath = ''
    interpreter_path = Path(sys.executable)
    for part in interpreter_path.parts:
        if folder_keyword in part:
            conda_basepath = ''.join(interpreter_path.parts[:interpreter_path.parts.index(part) + 1])
    if not conda_basepath:
        raise FileNotFoundError('Please install anaconda or miniconda at first and use them to run this script')
    else:
        return Path(conda_basepath)


def check_conda_bat(conda_basepath: Path, add_path: bool = True):
    """check whether conda.bat command runs, if not, try to add conda path to sys path"""
    try:
        check_output(shlex.split('conda.bat')).decode('utf-8').split('\n')
        print('conda.bat runs succeed')
    except Exception:
        print('Failed to run command "conda.bat". Try to add sys env variables of conda')
        if add_path:
            add_conda_path(conda_basepath)
        else:
            raise FileNotFoundError("conda.bat is not in sys path. Please add it manually")


def add_conda_path(conda_basepath: Path, path_overflow: bool = False, show_sys_path: bool = False):
    """add the paths under conda_basepath to sys_path"""
    # check conda path in sys_path or not
    path_to_add = ''
    sys_path = check_output(shlex.split('echo %path%'), shell=True).decode('utf-8').split(';')
    sys_path_lower = [p.rstrip().lower() for p in sys_path]
    if show_sys_path:
        print(f'win_path_lower: {sys_path_lower}')
    for subpath in ['', 'Scripts', r'Library\bin']:
        conda_path = Path(conda_basepath).joinpath(subpath)
        if not conda_path.exists():
            print('Path[{}] not exists'.format(conda_path))
        else:
            if str(conda_path).lower() not in sys_path_lower:
                path_to_add += str(conda_path) + ';'
                print('Path[{}] wait to add'.format(str(conda_path)))
            else:
                print('Path[{}] already in sys_path'.format(conda_path))

    if path_to_add:
        # get usr_path
        usr_path_to_backup = ''
        get_usr_path_cmd = 'reg query HKEY_CURRENT_USER\\Environment'
        # get_usr_path_cmd = 'powershell -NoProfile -Command "(Get-ItemProperty HKCU:\\Environment).PATH"'
        for line in check_output(get_usr_path_cmd, shell=True).decode('utf-8').strip().split('\n'):
            if 'Path    REG_SZ' in line.strip():
                usr_path_to_backup = line.strip()[len('Path    REG_SZ    '):]
            elif 'Path    REG_EXPAND_SZ' in line.strip():
                usr_path_to_backup = line.strip()[len('Path    REG_EXPAND_SZ    '):]

        if not usr_path_to_backup:
            print('Failed to get the usr path variable')
        else:
            print("Your usr path in usr variables: \n{}".format(usr_path_to_backup))
            # backup the usr path to file
            filename = cwdPath.joinpath('usrpath_backup_{}.txt'.format(time.strftime(
                '%Y%m%d%H%M%S', time.localtime())))
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(usr_path_to_backup)
                print(f'The usr path is saved to {filename}')

            # check the length of paths
            if len(usr_path_to_backup) + len(path_to_add) > 1024:
                print(f'The len of usr_path is {len(usr_path_to_backup)}.')
                print(f'The len of path_to_add is {len(path_to_add)}.')
                if not path_overflow:
                    raise OverflowError(
                        'The len of path to write is over 1024. Please add path[{}] to env manually'.format(
                            path_to_add[:-1]))
                else:
                    print('Froce to write sys_path in usr variables. The overflow part will be truncated.')
            if usr_path_to_backup[-1] == ';':
                usr_path_to_backup = usr_path_to_backup[:-1]
            add_result = check_output(f'setx path "{usr_path_to_backup};{path_to_add[:-1]}"', shell=True).decode('utf-8')
            print(add_result.strip())
            usr_path = check_output(get_usr_path_cmd, shell=True).decode('utf-8').strip()
            print("Path[{}] added to usr path".format(path_to_add[:-1]))
            print("Your usr variables now are: \n{}".format(usr_path.strip()))


def check_conda_env(env_name: str):
    """check the status of conda env named env_name"""
    conda_env_status = False
    for line in check_output(shlex.split('conda.bat info --envs')).decode('utf-8').split('\n'):
        if line[:len(env_name)] == env_name:
            conda_env_status = True
    return conda_env_status


def check_conda_settings(usr_folder: Path, force_to_cover: bool = True):
    """check and update settings in .condarc file"""
    def copy_default():
        src = cwdPath.joinpath(r'res\dev\.condarc')
        shutil.copy2(src, conda_settings)
        print(f'settings_file[{src}] copied to [{conda_settings}].')

    conda_settings = Path(usr_folder).joinpath('.condarc')
    if not conda_settings.exists():
        copy_default()
    else:
        # init
        def_channels = [
            'https://mirrors.bfsu.edu.cn/anaconda/pkgs/main',
            'https://mirrors.bfsu.edu.cn/anaconda/pkgs/r',
            'https://mirrors.bfsu.edu.cn/anaconda/pkgs/msys2'
        ]
        cus_channels = {
            'conda-forge': 'https://mirrors.bfsu.edu.cn/anaconda/cloud',
            'msys2': 'https://mirrors.bfsu.edu.cn/anaconda/cloud',
            'bioconda': 'https://mirrors.bfsu.edu.cn/anaconda/cloud',
            'menpo': 'https://mirrors.bfsu.edu.cn/anaconda/cloud',
            'pytorch': 'https://mirrors.bfsu.edu.cn/anaconda/cloud',
            'simpleitk': 'https://mirrors.bfsu.edu.cn/anaconda/cloud'
        }

        print(f'settings_file[{conda_settings}] already exists.')
        with open(conda_settings, 'r') as f:
            settings = yaml.load(f.read(), Loader=yaml.Loader)
        log.debug(f'settings before checked: {settings}')
        if settings is None:
            copy_default()
        elif not isinstance(settings, dict):
            raise TypeError(f'Error type[{type(settings)}] of settings file.')
        else:
            # check value of show_channel_urls
            if isinstance(settings.get('show_channel_urls', None), bool):
                if not settings['show_channel_urls']:
                    settings['show_channel_urls'] = True

            # check value of channels
            if isinstance(settings.get('channels', None), list):
                if 'defaults' not in settings['channels']:
                    settings['channels'].append('defaults')
            else:
                settings['channels'] = ['defaults']

            # check value of default_channels
            if isinstance(settings.get('default_channels', None), list):
                for ch in def_channels:
                    if ch not in settings['default_channels']:
                        settings['default_channels'].append(ch)
            else:
                settings['default_channels'] = def_channels

            # check value of custom_channels
            if isinstance(settings.get('custom_channels', None), dict):
                for alias in cus_channels.keys():
                    if alias not in settings['custom_channels'].keys():
                        settings['custom_channels'][alias] = cus_channels[alias]
                    else:
                        if force_to_cover:
                            settings['custom_channels'][alias] = cus_channels[alias]
                        else:
                            tips = "Change the custom channel[{}] from \n   value[{}] \nto value[{}]".format(
                                alias, settings['custom_channels'][alias], cus_channels[alias]
                            )
                            if input_checking_YN(tips, default_Y=False) == 'Y':
                                settings['custom_channels'][alias] = cus_channels[alias]
                                print('Changed.')
                            else:
                                print('Canceled.')
            else:
                settings['custom_channels'] = cus_channels

        log.debug(f'settings after checked: {settings}')
        # update conda settings file
        if settings is not None:
            with open(conda_settings, 'w') as f:
                yaml.dump(settings, f, Dumper=yaml.Dumper)


def create_conda_env(env_name: str, configfile: Path, force_to_install: bool = False):
    """create conda env named env_name from configfile"""
    def install(env_name, configfile):
        print(f'Installing the conda env[{env_name}]. It might takes a few minutes...')
        install_cmd = 'conda.bat env create -f "{}"'.format(str(configfile).replace('\\', '/'))
        proc_run(install_cmd)

    # clean cache to make sure to use the updated channels
    for line in check_output(shlex.split('conda.bat clean -i')).decode('utf-8').split('\n'):
        print(line)

    status = check_conda_env(env_name)
    if status:
        print(f'Conda environment[{env_name}] already exists')
    if not status:
        install(env_name, configfile)
    elif force_to_install:
        remove_conda_env(env_name, force_to_remove=True)
        install(env_name, configfile)


def remove_conda_env(env_name: str, force_to_remove: bool = False):
    """remove conda env named env_name"""

    if not check_conda_env(env_name):
        print(f'Conda environment[{env_name}] not exists')
    else:
        print(f'Try to remove the conda env[{env_name}]. It might takes a few minutes...')
        if force_to_remove:
            remove_cmd = 'conda.bat remove -y -n "{}" --all'.format(env_name)
        else:
            remove_cmd = 'conda.bat remove -n "{}" --all'.format(env_name)
        proc_run(remove_cmd, break_lines=['Proceed ([y]/n)?'])


if __name__ == '__main__':
    import os
    usr_folder = Path(os.path.expanduser('~'))

    if True:
        conda_basepath = find_conda_path(folder_keyword='conda')
        print(f"conda_basepath: {conda_basepath}")

    if False:
        check_conda_bat(conda_basepath, add_path=True)

    if False:
        print(check_conda_env('pyexcel'))
        print(check_conda_env('testconda'))

    if False:
        check_conda_settings(usr_folder, force_to_cover=True)

    if False:
        create_conda_env('testconda', cwdPath.joinpath('res/dev/testconda.yaml'), force_to_install=True)

    if False:
        add_conda_path(conda_basepath.joinpath('envs/testconda'), path_overflow=True)

    if False:
        remove_conda_env('testconda', force_to_remove=False)
