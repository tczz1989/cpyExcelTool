import configparser
import os
import glob


def get_input_config():
    cf = configparser.ConfigParser()
    cf.read("..\\config.ini")  # 读取配置文件，如果写文件的绝对路径，就可以不用os模块
    items = cf.items("input")  # 获取指定section 的option值
    return items


def get_output_config():
    cf = configparser.ConfigParser()
    cf.read("..\\config.ini")  # 读取配置文件，如果写文件的绝对路径，就可以不用os模块
    items = cf.items("output")  # 获取指定section 的option值
    return items


def get_input_filenames():
    files = glob.glob(os.path.join("../input", '*.xls*'))
    # for file in files:
    #     print(file)
    return files
