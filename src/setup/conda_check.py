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
import shlex
from subprocess import check_output, Popen, PIPE, STDOUT, TimeoutExpired


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


def create_conda_env(env_name: str, configfile: Path):
    if check_conda_env(env_name):
        print(f'Conda environment[{env_name}] already exists')
    else:
        print(f'Installing the conda env[{env_name}]. It might takes a few minutes...')
        install_cmd = 'conda.bat env create -f "{}"'.format(str(configfile).replace('\\', '/'))
        proc = Popen(shlex.split(install_cmd), stdin=PIPE, stdout=PIPE, stderr=STDOUT)
        while True:
            line = proc.stdout.readline()
            if not line:
                break
            print(line.decode('utf-8').rstrip())


def remove_conda_env(env_name: str, force_to_remove: bool = False):
    def terminated_read(stdout, terminators: str) -> str:
        buf = []
        while True:
            if stdout.readable():
                r = stdout.read(1).decode('utf-8')
                # print(r)
                buf.append(r)
                if r in terminators:
                    break
        return "".join(buf)

    if not check_conda_env(env_name):
        print(f'Conda environment[{env_name}] not exists')
    else:
        print(f'Try to remove the conda env[{env_name}]. It might takes a few minutes...')
        if force_to_remove:
            remove_cmd = 'conda.bat remove -y -n "{}" --all'.format(env_name)
        else:
            remove_cmd = 'conda.bat remove -n "{}" --all'.format(env_name)
        print(shlex.split(remove_cmd))
        proc = Popen(shlex.split(remove_cmd), stdin=PIPE, stdout=PIPE, stderr=STDOUT)
        try:
            while True:
                line = terminated_read(proc.stdout, "\n?")
                print(line.rstrip())
                if not line:
                    break
                elif line.rstrip() == 'Proceed ([y]/n)?':
                    proc.stdin.write(input().encode('utf-8'))
                    proc.stdin.close()
        except TimeoutExpired:
            print('Time out. Kill the process')
            proc.kill()


if __name__ == '__main__':
    if True:
        conda_basepath = find_conda_path(folder_keyword='conda')
        print(f"conda_basepath: {conda_basepath}")

    if False:
        check_conda_bat(conda_basepath, add_path=True)

    if False:
        print(check_conda_env('pyexcel'))

    if False:
        create_conda_env('testconda', cwdPath.joinpath('res/dev/testconda.yaml'))

    if False:
        add_conda_path(conda_basepath.joinpath('envs/testconda'), path_overflow=True)

    if False:
        remove_conda_env('testconda', force_to_remove=False)
