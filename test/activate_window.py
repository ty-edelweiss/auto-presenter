#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from AppKit import (
    NSRunningApplication,
    NSApplicationActivateIgnoringOtherApps
)


def main(arg):
    """TODO: Docstring for main.
    :arg: int
    :returns: void

    """
    app = NSRunningApplication.runningApplicationWithProcessIdentifier_(arg)
    app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)


if __name__ == "__main__":
    pid = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    main(pid)
