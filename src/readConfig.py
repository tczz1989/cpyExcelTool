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


def get_output_filenames():
    files = os.listdir("../output/")
    return files


def config_is_xlsx(config):
    file_format = os.path.splitext(config[0][1])[-1].lower()
    if file_format == '.xlsx':
        return True
    elif file_format == '':
        config[0] = ('filename', config[0][1] + '.xlsx')
        return True
    elif file_format == '.':
        config[0] = ('filename', config[0][1] + 'xlsx')
        return True
    else:
        return False


def get_config():
    input_config = get_input_config()
    output_config = get_output_config()
    # print(input_config, output_config)
    get_input_filenames()
    if config_is_xlsx(input_config) and config_is_xlsx(output_config):
        pass
    else:
        raise Exception("config input or output file name error!\n")
    return input_config, output_config
