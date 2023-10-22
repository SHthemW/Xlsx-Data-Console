import os
import xml.etree.ElementTree as Et


def get_config(config_name: str):
    # 解析XML文件
    pathname = __file__
    os.chdir(os.path.dirname(pathname))
    tree = Et.parse(r'../_config.xml')
    # 获取根节点
    root = tree.getroot()
    # 读取配置信息
    return root.find(config_name).text


def get_directory() -> str:
    return get_config("datasource")

def get_window_title_suffix() -> str:
    return get_config("windowtitle")


if __name__ == "__main__":
    print(get_directory())
    print(get_window_title_suffix())
