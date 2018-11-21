# -*-coding:utf-8
import binascii
import re
import sys


# 保存至文件
def savefile(savepath, content):
    with open(savepath, 'wb') as fp:
        fp.write(content.encode())  # str转换为二进制存储
        # fp.close()


def readfile(path):
    with open(path, 'rb') as fp:
        # content = binascii.b2a_hex(fp.read()).decode('utf-8')  # 二进制转换为str
        # fp.close()
        content = " ".join(re.findall(r'.{2}', binascii.b2a_hex(fp.read()).decode('utf-8')))
    return content


# content="I love machine learning!"
# savefile("d:\\mytest1.txt",content)


def main(filename):
    content1 = readfile(filename)
    # print(content1)
    pathname = "".join(re.findall(r'(.+?)\.', filename, flags=re.IGNORECASE))
    with open(pathname + ".txt", 'w', encoding='utf-8') as fp:
        fp.write(content1)


if __name__ == '__main__':
    sys.argv.append('')
    main(sys.argv[1])
