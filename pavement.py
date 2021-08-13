import glob
import os
import sys
import subprocess
import locale

import pypandoc

from paver.easy import *

TEST_DIR = "tests"
FUNCTION_TEST = "test_function.py"

README_MD = "README.md"
README_RST = "README.rst"


@task
@consume_args
def test(args):
    verbose = False
    if "-v" in args:
        verbose = True
        args.remove("-v")
    if len(args) == 0:
        cmd_arg = " ".join(collect_unittests(TEST_DIR))
    elif args[0] == "all":
        cmd_arg = f"discover {TEST_DIR}"
    elif args[0] == "function":
        cmd_arg = f"{TEST_DIR}/test_function.py"
    cmd_arg += " -v" if verbose else ""
    do_test(cmd_arg)


@task
def rst():
    if not os.path.exists(README_RST):
        print("generate", README_RST)
        convert(README_MD, README_RST)
    elif mtime(README_RST) < mtime(README_MD):
        print("update", README_RST)
        convert(README_MD, README_RST)


def mtime(file):
    return os.stat(file).st_mtime


def convert(md_file, rst_file):
    with open(rst_file, "w") as f:
        f.write(pypandoc.convert(md_file, "rst"))


def collect_unittests(dir):
    files = glob.glob(f"{dir}/*.py")
    return [f for f in files if FUNCTION_TEST not in f]


def do_test(cmd_arg):
    cmd = "python -m unittest " + cmd_arg
    exec_cmd(cmd)


def exec_cmd(cmd):
    print("exec:", cmd)
    proc = subprocess.Popen(
        cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT
    )
    while True:
        line = proc.stdout.readline()
        if line:
            sys.stdout.write(line.decode(locale.getpreferredencoding()))
        if not line and proc.poll() is not None:
            return
