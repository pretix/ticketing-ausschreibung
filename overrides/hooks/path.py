import os
import sys


def on_config(config, *args, **kwargs):
    sys.path.insert(0, os.getcwd())
    config["markdown_extensions"].append("yamlloader")