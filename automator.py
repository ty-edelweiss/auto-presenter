#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import os
import sys
import time
import glob
import yaml
import signal
import pickle
import logging
import subprocess
from datetime import date
from optparse import OptionParser

import pyautogui
from Quartz import (
    CGWindowListCopyWindowInfo,
    kCGWindowListExcludeDesktopElements,
    kCGWindowListOptionOnScreenOnly,
    kCGNullWindowID
)
from AppKit import (
    NSRunningApplication,
    NSApplicationActivateIgnoringOtherApps
)
from Foundation import NSSet

logging.basicConfig(format="%(asctime)s [%(name)s] - %(levelname)s: %(message)s", level=logging.DEBUG)
LOGGER_ = logging.getLogger(__file__)

LOADING_ = [".    ", "..   ", "...  ", ".... ", "....."]

CONFIG_ = "config.yaml"
OWNER_KEY_ = "kCGWindowOwnerName"
NAME_KEY_ = "kCGWindowName"
BOUNDS_KEY_ = "kCGWindowBounds"
PID_KEY_ = "kCGWindowOwnerPID"

BASE_APP_ = "zoom.us"
BASE_WINDOWS_ = {
    "en": [
        "^Zoom - Free Account$",
        "^Zoom *(Participant ID: [0-9]+    Meeting ID: [0-9\-]+)*$",
        "^Select a window or an application that you want to share+$",
    ],
    "ja": [
        "^Zoom - 無料アカウント$",
        "^Zoom *(ミーティング ID: [0-9\-]+)*$",
        "^共有するウィンドウまたはアプリケーションの選択$",
    ],
}
INTERVAL_ = 1
RETRY_ = 2

WINDOWS_ = None
LOGGER_.setLevel(logging.INFO)


class Cache(object):

    _instance = None

    caching_ = True
    cache_file_ = None

    def __new__(cls):
        raise NotImplementedError('Cannot initialize via Constructor')

    @classmethod
    def __internal_new__(cls):
        return super().__new__(cls)

    @classmethod
    def get_instance(cls):
        if not cls._instance:
            cls._instance = cls.__internal_new__()
        return cls._instance

    @classmethod
    def set(cls, caching: bool, config: dict, path: str):
        cls.caching_ = caching
        cls.cache_file_ = os.path.join(path, config["env"]["cache"])

    @classmethod
    def load(cls):
        if not cls.caching_ or not os.path.isfile(cls.cache_file_):
            return {}
        with open(cls.cache_file_, "rb") as fp:
            cache = pickle.load(fp, encoding="bytes")
        return cache

    @classmethod
    def save(cls, cache: dict):
        if cls.caching_:
            with open(cls.cache_file_, "wb") as fp:
                pickle.dump(cache, fp)
        LOGGER_.debug("(Cache) -> %s", cache)


class Scheduler(object):

    def __init__(self, config: dict, path: str):
        self.stack_ = Cache.get_instance().load()
        self.path_ = path
        self.config_ = config

    def calc_priority(self, filepath):
        group = os.path.dirname(filepath).split("/")[-1]
        m = re.match(r"^([1-9]).+$", os.path.basename(filepath))
        speciality = 10 - int(m.groups()[0]) if m else 0
        return self.config_["data"]["groups"][group] + speciality

    def notify(self, filekey: str):
        self.stack_[filekey]["completed"] = True
        Cache.get_instance().save(self.stack_)
        return True

    def assign(self):
        corr_table = self.config_["env"]["extensions"]
        files = [f for ext in corr_table for f in glob.glob(os.path.join(self.path_, f"**/*{ext}"))]
        LOGGER_.debug(files)
        for filepath in files:
            filekey = "/".join(filepath.split("/")[-2:])
            if filekey not in self.stack_:
                root, ext = os.path.splitext(os.path.basename(filepath))
                self.stack_[filekey] = {
                    "target": root,
                    "app": corr_table[ext],
                    "path": filepath,
                    "priority": self.calc_priority(filepath),
                    "completed": False
                }
        if len(self.stack_) == 0:
            return None
        Cache.get_instance().save(self.stack_)
        tasks = sorted(self.stack_.values(), key=lambda t: t["priority"] * int(not t["completed"]))
        new_task = tasks.pop()
        LOGGER_.debug(new_task)
        return new_task if not new_task["completed"] else None


def window_info(owner: str, compress: bool = True):
    options = kCGWindowListExcludeDesktopElements & kCGWindowListOptionOnScreenOnly
    wl = CGWindowListCopyWindowInfo(options, kCGNullWindowID)
    w = NSSet.setWithArray_(wl)
    if not compress:
        return [wi for wi in w if wi[OWNER_KEY_] == owner]
    info = {wi[NAME_KEY_]: wi for wi in w if NAME_KEY_ in wi and wi[OWNER_KEY_] == owner}
    return info


def window_activate(pid: int):
    app = NSRunningApplication.runningApplicationWithProcessIdentifier_(pid)
    app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)


def active_application():
    script = "from AppKit import NSWorkspace;" \
    "print(NSWorkspace.sharedWorkspace().frontmostApplication().localizedName(), end='')"
    cmd = ["python",  "-c", script]
    output = subprocess.check_output(cmd)
    return output.decode(encoding="utf-8")


def window_startup(target: str, app: str, path: str):
    cmd = ["open",  path]
    status = subprocess.check_call(cmd)
    if status:
        print_e(f"ERROR: {app} cannot be opened")
        exit(1)
    prog = re.compile(f"^{target}.*$")
    damping = 4
    info = None
    while True:
        all_info = window_info(app)
        info = [v for k, v in all_info.items() if prog.match(k)]
        if len(info) > 0:
            time.sleep(INTERVAL_)
            break
        elif len(all_info):
            time.sleep(INTERVAL_)
        else:
            time.sleep(damping * INTERVAL_)
            damping = damping - 1 if damping > 1 else 1
        LOGGER_.debug(f"Starting up the window :: {app} -> {target}")
    return info[0]


def base_operation(zoom_col: int):
    counter = 0
    active_app = BASE_APP_
    while active_app == BASE_APP_:
        if counter > RETRY_:
            break

        all_info = window_info(BASE_APP_)

        prog = re.compile(WINDOWS_[0])
        win_info = [v for k, v in all_info.items() if prog.match(k)]
        if len(win_info) > 0:
            pid = win_info[0][PID_KEY_]
            window_activate(pid)
        else:
            print_e(f"ERROR: {BASE_APP_} is not ready to start a presentation.")
            exit(1)

        time.sleep(INTERVAL_)

        prog = re.compile(WINDOWS_[1])
        win_info = [v for k, v in all_info.items() if prog.match(k)]
        if len(win_info) > 0:
            pyautogui.press("esc")
            bounds = win_info[0][BOUNDS_KEY_]
            window_x, window_y = bounds["X"] + (0.5 * bounds["Width"]), bounds["Y"] + (0.55 * bounds["Height"])
            pyautogui.moveTo(window_x, window_y)
            pyautogui.click(window_x, window_y)
        else:
            print_e(f"ERROR: {BASE_APP_} is not ready to start a presentation.")
            exit(1)

        time.sleep(INTERVAL_)
        all_info = window_info(BASE_APP_)

        prog = re.compile(WINDOWS_[2])
        win_info = [v for k, v in all_info.items() if prog.match(k)]
        if len(win_info) > 0:
            bounds = win_info[0][BOUNDS_KEY_]
            window_x, window_y = bounds["X"] + (0.2 * zoom_col * bounds["Width"]), bounds["Y"] + (0.5 * bounds["Height"])
            pyautogui.moveTo(window_x, window_y)
            pyautogui.click(window_x, window_y)
            window_x, window_y = bounds["X"] + (bounds["Width"] - 50), bounds["Y"] + (bounds["Height"] - 20)
            pyautogui.moveTo(window_x, window_y)
            pyautogui.click(window_x, window_y)
        else:
            print_e(f"ERROR: {BASE_APP_} is not ready to start a presentation.")
            exit(1)

        time.sleep(INTERVAL_)
        active_app = active_application()

        counter = counter + 1

    return active_app


def operation(owner: str, info: dict, zoom_col: int = 0):
    if zoom_col == 0:
        active_app = active_application()
    else:
        active_app = base_operation(zoom_col)

        LOGGER_.debug(active_app)
        time.sleep(INTERVAL_)

    counter = 0
    new_win = 0
    win_num = len(window_info(active_app, compress=False))
    while not new_win:
        if counter > RETRY_:
            break

        if active_app == "Microsoft PowerPoint":
            pyautogui.hotkey("shift", "command", "enter")
        elif active_app == "Skim" or active_app == "Keynote":
            pyautogui.hotkey("option", "command", "p")
        elif active_app == "Preview" or active_app == "プレビュー":
            pyautogui.hotkey("shift", "command", "f")

        time.sleep(2 * INTERVAL_)
        new_win = len(window_info(active_app, compress=False)) - win_num

        counter = counter + 1


def automation(scheduler: object, **kwargs):
    task = scheduler.assign()
    if task is None:
        return task
    task_info = window_startup(task["target"], task["app"], task["path"])
    operation(task["app"], task_info, **kwargs)
    return task


def print_b(message: str, endl: str = "\n"):
    print(f"\033[1;39m{message}\033[0;39m", end=endl)


def print_e(message: str, endl: str = "\n"):
    print(f"\033[0;31m{message}\033[0;39m", end=endl)


def signal_handler(signal, frame):
    print_b("\rToday's meeting ended. Have a nice day!")
    sys.exit(0)


def adjust_cache(origin_cache):
    print("\033[1;31mBegin adjustments for automation before the meeting\033[0;39m")

    tmp = sorted([(k, v["priority"]) for k, v in origin_cache.items()], key=lambda x: x[1], reverse=True)
    key_map = {str(idx+1): key for idx, (key, _) in enumerate(tmp)}

    cache = origin_cache.copy()
    while True:
        cmd = input("\033[1;31m>> \033[0;39m")
        if cmd == "exit":
            break
        elif cmd == "delete":
            print("\033[1;39m[Succeeded]\033[0;39m Deleted all caches")
            cache = {}
            key_map = {}
        elif cmd == "list":
            print("\033[1;39m[Succeeded]\033[0;39m List all caches ---------->")
            for idx, key in key_map.items():
                print(f"\033[1;31m[{idx}]\033[0;39m {key}")
                lines = [f"{k}: {v}" for k, v in cache[key].items()]
                print("\n".join(lines))
            print("<---------------------------------- End")
        elif cmd in key_map:
            key = key_map[cmd]
            print("\033[1;39m[Succeeded]\033[0;39m Select the key from among (" + ", ".join(cache[key]) + ")")
            property_key = input(f"\033[1;31m{key} >> \033[0;39m")
            if property_key not in cache[key]:
                print("\033[1;39m[Failed]\033[0;39m Invalid key.")
            elif isinstance(cache[key][property_key], bool):
                print(f"\033[1;39m[Succeeded]\033[0;39m Input new bool value. (now: {cache[key][property_key]})")
                new_value = input(f"\033[1;31m[bool] {key}::{property_key} >> \033[0;39m")
                if new_value.lower() == "true":
                    cache[key][property_key] = True
                elif new_value.lower() == "false":
                    cache[key][property_key] = False
                else:
                    print("Invalid type.")
            elif isinstance(cache[key][property_key], int):
                print(f"\033[1;39m[Succeeded]\033[0;39m Input new int value. (now: {cache[key][property_key]})")
                try:
                    new_value = input(f"\033[1;31m[int] {key}::{property_key} >> \033[0;39m")
                    cache[key][property_key] = int(new_value)
                except ValueError:
                    print("\033[1;39m[Failed]\033[0;39m Invalid type.")
            else:
                print(f"\033[1;39m[Succeeded]\033[0;39m Input new string value. (now: {cache[key][property_key]})")
                input(f"\033[1;31m[str] {key}::{property_key} >> \033[0;39m")
                if new_value:
                    cache[key][property_key] = new_value
                else:
                    print("\033[1;39m[Failed]\033[0;39m Invalid type.")
        else:
            print("\033[1;39m[Failed]\033[0;39m Invalid command.")

    return cache


def command_line(usage):
    parser = OptionParser(usage=usage)
    parser.add_option("-v", "--verbose",
                      action="store_true", dest="verbose", default=False,
                      help="don't print debug messages to stdout")
    parser.add_option("-n", "--non-cache",
                      action="store_true", dest="non_cache", default=False,
                      help="disable functions for caching")
    parser.add_option("-l", "--lang",
                      type="str", dest="lang", default="en",
                      help=f"select language settings on an environment [en, ja]")
    parser.add_option("-m", "--manual",
                      action="store_true", dest="manual", default=False,
                      help="require user's input for next steps")
    parser.add_option("-a", "--adjust",
                      action="store_true", dest="adjust", default=False,
                      help="manually adjust cache for autmation.")
    parser.add_option("--zoom-col",
                      type="int", dest="zoom_column", default=0,
                      help=f"select a column for click on {BASE_APP_} [1,2,3,4]")

    options, _ = parser.parse_args()

    return options


def main():
    global WINDOWS_

    options = command_line("usage: %prog [options]")

    signal.signal(signal.SIGINT, signal_handler)

    if options.verbose:
        LOGGER_.setLevel(logging.DEBUG)
        LOGGER_.debug("Running on debug mode")

    WINDOWS_ = BASE_WINDOWS_[options.lang]

    with open(CONFIG_, "r") as fp:
        config = yaml.load(fp)
        LOGGER_.debug(config)

    today = date.today()
    print_b(f"Meeting :: {today}")
    path = os.path.join(
        os.path.expanduser(config["data"]["assets"]),
        today.strftime("%Y%m%d")
    )

    if not os.path.isdir(path):
        os.mkdir(path)
        os.chmod(path, 0o777)
    for group in config["data"]["groups"].keys():
        directory = os.path.join(path, group)
        if not os.path.isdir(directory):
            os.mkdir(directory)
            os.chmod(directory, 0o777)

    Cache.get_instance().set(not options.non_cache, config, path)
    if not options.non_cache and options.adjust:
        new_cache = adjust_cache(Cache.get_instance().load())
        Cache.get_instance().save(new_cache)

    scheduler = Scheduler(config, path)

    all_info = window_info(BASE_APP_)
    if WINDOWS_[0][1:-1] not in all_info:
        print_e(f"ERROR: {BASE_APP_} is not ready to start meeting.")
        exit(1)

    tick = 0

    print_b("Let's get started today's meeting!")
    wait = False
    while True:
        if not wait:
            presenter = automation(scheduler, zoom_col=options.zoom_column)
            if presenter is not None:
                print(f"\r              ", end="\r")
                print_b(f"Next speaker is the author of {presenter['target']}.")
                tick = 0
                wait = True
            elif not options.manual:
                print(f"\rWaiting for new documents {LOADING_[tick]} ", end="")
                time.sleep(config["env"]["sleep"])
                tick = tick + 1 if tick < len(LOADING_) - 1 else 0
            else:
                print_b("Next docuemnt doesn't exist.")
                if input("Did you finish today's meeting? (y/n) :: ") == "y":
                    break
        else:
            print(f"\rGive a presentation {LOADING_[tick]} ", end="")
            all_info = window_info(presenter["app"])
            prog = re.compile(f"^{presenter['target']}.*$")
            win_info = [v for k, v in all_info.items() if prog.match(k)]
            if not len(win_info):
                print(f"\rGive a presentation {LOADING_[len(LOADING_)-1]} [done]")
                task_key = "/".join(presenter["path"].split("/")[-2:])
                scheduler.notify(task_key)
                tick = 0
                wait = False
            else:
                time.sleep(config["env"]["sleep"])
                tick = tick + 1 if tick < len(LOADING_) - 1 else 0
    print_b("Today's meeting ended. Have a nice day!")


if __name__ == "__main__":
    main()
