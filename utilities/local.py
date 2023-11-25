import os
import time
import xml.etree.ElementTree as Et

CONFIG_PATH = r'../.resx/_config.xml'


# io

def get_config(config_name: str):
    # 解析XML文件
    pathname = __file__
    os.chdir(os.path.dirname(pathname))
    tree = Et.parse(CONFIG_PATH)
    # 获取根节点
    root = tree.getroot()
    # 读取配置信息
    return root.find(config_name).text


def config_is_exist() -> bool:
    pathname = __file__
    os.chdir(os.path.dirname(pathname))
    return os.path.exists(CONFIG_PATH)


def create_config_file():
    pathname = __file__
    os.chdir(os.path.dirname(pathname))
    # 创建根元素
    root = Et.Element('config')
    datasource = Et.SubElement(root, 'datasource')
    windowtitle = Et.SubElement(root, 'windowtitle')
    # defaults
    Et.SubElement(root, 'spaceoffset').text = "0"
    Et.SubElement(root, 'chscharcomp').text = "0"
    Et.SubElement(root, 'autoclose').text = "True"
    Et.SubElement(root, 'autoreset').text = "True"

    # 从控制台获取数据并设置为元素的文本内容
    datasource.text = input('请输入数据源路径：')
    windowtitle.text = input('请输入打开xlsx文件时窗体的后缀名：')

    # 创建ElementTree对象
    tree = Et.ElementTree(root)
    # 将ElementTree对象写入XML文件
    tree.write(CONFIG_PATH)
    time.sleep(1)


# getters

def get_directory() -> str:
    return get_config("datasource")


def get_window_title_suffix() -> str:
    return get_config("windowtitle")


def get_output_space_offset() -> int:
    return int(get_config("spaceoffset"))


def get_chinese_char_comp() -> float:
    return float(get_config("chscharcomp"))


def enable_auto_close() -> bool:
    return bool(get_config("autoclose"))


def enable_auto_restart() -> bool:
    return bool(get_config("autoreset"))


# debug

if __name__ == "__main__":
    print(get_directory())
    print(get_window_title_suffix())
    print(enable_auto_close())
    print(enable_auto_restart())
