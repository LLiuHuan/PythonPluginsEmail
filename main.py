# -*- coding: utf-8 -*-

import os
import sys
import importlib
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
from watchdog.events import FileSystemEventHandler

from PluginManager import PluginManager
from PluginManager import __ALLMODEL__

from apscheduler.schedulers.background import BackgroundScheduler
__SCHEDULER = BackgroundScheduler()
__SCHEDULER.start()

import signal

PLUGINS = {}

# 设置相应信号处理的handler
signal.signal(signal.SIGINT, exit)
signal.signal(signal.SIGTERM, exit)

isStop = False


def modified(fileName):
    try:
        item = PLUGINS.get(fileName, "")
        if item != "":
            __SCHEDULER.remove_job(fileName)

            __SCHEDULER.add_job(item.Start,
                                'cron',
                                month='*',
                                day='*',
                                hour='*',
                                minute=str(item.minute),
                                id=fileName)
    except Exception as e:
        pass


def deleted(fileName):
    try:
        __SCHEDULER.remove_job(fileName)
        PLUGINS.pop(fileName)
    except Exception as e:
        pass


def created():
    try:
        PLUGINS = {}
        __SCHEDULER.remove_all_jobs()
    except Exception as e:
        pass


# 监听配置变化
class FileMonitorHandler(FileSystemEventHandler):
    def __init__(self, **kwargs):
        super(FileMonitorHandler, self).__init__(**kwargs)
        # 监控目录 目录下面以device_id为目录存放各自的图片
        self._watch_path = "./plugins"

    # 重写文件改变函数，文件改变都会触发文件夹变化
    def on_modified(self, event):
        if not event.is_directory and event.src_path.count('\\') == 1:
            fileName = os.path.basename(event.src_path).split('.')[0]
            modified(fileName)

    # 创建文件
    def on_created(self, event):
        if not event.is_directory and event.src_path.count('\\') == 1:
            print('创建了文件夹', event.src_path)
            created()
            main()

    # 移动文件
    def on_moved(self, event):
        if not event.is_directory and event.src_path.count('\\') == 1:
            print("移动了文件", event.src_path)
            fileName = os.path.basename(event.src_path).split('.')[0]
            deleted(fileName)

    # 删除文件
    def on_deleted(self, event):
        print("删除了文件", event.src_path)
        if not event.is_directory and event.src_path.count('\\') == 1:
            fileName = os.path.basename(event.src_path).split('.')[0]
            deleted(fileName)

    # 都会触发
    # def on_any_event(self, event):
    # print("都会触发")
    # print()


def init(addonPath):
    if not os.path.exists(addonPath):
        os.mkdir(addonPath)


def main():
    global PLUGINS
    #加载所有插件
    PluginManager.LoadAllPlugin()
    #遍历所有接入点下的所有插件
    for SingleModel in __ALLMODEL__:
        plugins = SingleModel.GetPluginObject()
        for item in plugins:
            PLUGINS[item.filename] = item
            #调用接入点的公共接口
            __SCHEDULER.add_job(item.Start,
                                'cron',
                                **item.corn,
                                id=item.filename)


def exit(signum, frame):
    global isStop
    isStop = True
    print('You choose to stop me.')
    __SCHEDULER.shutdown()
    observer.stop()


if __name__ == "__main__":
    main()
    event_handler = FileMonitorHandler()
    observer = Observer()
    observer.schedule(event_handler, path="./plugins",
                      recursive=True)  # recursive递归的
    observer.start()

    import time
    while not isStop:
        time.sleep(1)
