import os
import xml.etree.ElementTree as Et


# 定义要搜索的目录
def get_directory() -> str:
    pathname = __file__
    os.chdir(os.path.dirname(pathname))

    with open(r'../_config.txt', 'r') as file:
        data = file.read()
    # 使用ElementTree解析数据
    root = Et.fromstring(data)
    # 获取<datasource>标签中的文本，并赋值给directory
    return root.text.strip()


if __name__ == "__main__":
    path = r'/'
    # 获取当前目录下的所有文件
    files = [os.path.join(path, file) for file in os.listdir(path)]
    # 遍历文件列表，输出文件名
    for file in files:
        print(file)
