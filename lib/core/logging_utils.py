# -*- coding: utf-8 -*-
import os
import datetime


def write_log(log_dir, log_name, message):
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    path = os.path.join(log_dir, log_name)
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(path, 'a') as f:
        f.write('[{}] {}\n'.format(timestamp, message))
    return path
