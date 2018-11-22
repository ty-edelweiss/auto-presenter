#!/usr/bin/env python
# -*- coding: utf-8 -*-

from AppKit import NSWorkspace


def main():
    """TODO: Docstring for main.
    :returns: void

    """
    app = NSWorkspace.sharedWorkspace().frontmostApplication()
    active_app_name = app.localizedName()
    print("\nAcativate window:")
    print(active_app_name)
    print("\n")


if __name__ == "__main__":
    main()
