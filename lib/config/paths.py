# -*- coding: utf-8 -*-
import os


def get_extension_root():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(current_dir, '..', '..'))

BASE_DIR = get_extension_root()
DATA_DIR = os.path.join(BASE_DIR, 'data')
MASTER_DIR = os.path.join(DATA_DIR, 'master')
CACHE_DIR = os.path.join(DATA_DIR, 'cache')
TEMP_DIR = os.path.join(DATA_DIR, 'temp')
LOG_DIR = os.path.join(DATA_DIR, 'logs')
EXPORT_DIR = os.path.join(DATA_DIR, 'exports')


def ensure_runtime_dirs():
    for path in [DATA_DIR, MASTER_DIR, CACHE_DIR, TEMP_DIR, LOG_DIR, EXPORT_DIR]:
        if not os.path.exists(path):
            os.makedirs(path)
