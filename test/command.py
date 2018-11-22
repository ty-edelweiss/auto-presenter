#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import subprocess


def main(arg):
    """TODO: Docstring for main.
    :arg: list
    :returns: void

    """
    print(f"Running command `{' '.join(arg)}`")
    status = subprocess.check_call(arg)
    print(f"Status code = {status}")


if __name__ == "__main__":
    cmd = sys.argv[1:] if len(sys.argv) > 1 else []
    main(cmd)
