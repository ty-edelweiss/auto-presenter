#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from Quartz import (
    CGWindowListCopyWindowInfo,
    kCGWindowListExcludeDesktopElements,
    kCGWindowListOptionOnScreenOnly,
    kCGNullWindowID
)
from Foundation import NSSet


def main(arg):
    """TODO: Docstring for main.
    :arg: string
    :returns: void

    """
    options = kCGWindowListExcludeDesktopElements & kCGWindowListOptionOnScreenOnly
    wl = CGWindowListCopyWindowInfo(options, kCGNullWindowID)
    w = NSSet.setWithArray_(wl)
    print("\nList of windows:")
    if arg:
        print([info for info in w if info["kCGWindowOwnerName"] == arg])
    else:
        print(w)
    print("\n")


if __name__ == "__main__":
    owner = sys.argv[1] if len(sys.argv) > 1 else ""
    main(owner)
