#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   util.py
@Author  :   Billy Zhou
@Time    :   2021/11/04
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import codecs
import shlex
from subprocess import Popen, PIPE, STDOUT, TimeoutExpired


def proc_run(cmd: str, break_lines: list = [], shell: bool = False, input_encode: str = 'utf-8', output_decode: str = 'utf-8'):
    def terminated_read(stdout: codecs.StreamReader, terminators: str) -> str:
        buf = []
        while stdout.readable():
            r = stdout.read(1)
            # print(r)
            buf.append(r)
            if r in terminators:
                break
        return "".join(buf)

    proc = Popen(shlex.split(cmd), shell=shell, stdin=PIPE, stdout=PIPE, stderr=STDOUT)
    try:
        try:
            proc.stdout = codecs.getreader(output_decode)(proc.stdout)
        except LookupError:
            raise LookupError

        while True:
            line = terminated_read(proc.stdout, "\n?")
            print(line.rstrip())
            if not line:
                break
            elif line.rstrip() in break_lines:
                proc.stdin.write(input().encode(input_encode))
                proc.stdin.close()
    except TimeoutExpired:
        print('Time out. Kill the process')
    finally:
        proc.kill()


if __name__ == '__main__':
    if False:
        proc_run('ping www.baidu.com', output_decode='gbk')

    if True:
        proc_run('ls', output_decode='gbk')
