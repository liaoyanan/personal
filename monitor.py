from datetime import datetime
import time
import hashlib
import os


# 计算文件md5
def get_md5(file_path, Bytes=1024):
    md5 = hashlib.md5()
    with open(file_path, 'rb') as f:
        while 1:
            data = f.read(Bytes)
            if data:
                md5.update(data)
            else:
                break
    result = md5.hexdigest()
    return result


# 获取目录下所有文件
def get_file(path):
    file_path = []
    file = os.walk(path)
    for root, dirs, files in file:
        for file in files:
            file_path.append(root + '\\' + file)
    return file_path


def compare(file1,file2):
    st1 = os.stat(file1)
    st2 = os.stat(file2)
    #encoding_method = sys.getdefaultencoding()
    #windows上带encoding='gbk'，Linux去掉
    with open(file1, 'r', encoding='gbk') as fp1, open(file2, 'r', encoding="gbk") as fp2:
        while True:
            b1 = fp1.readline()  # 读取指定大小的数据进行比较
            b2 = fp2.readline()
            if b1 != b2:
                print("发现变更：\n"+str(b1) + str(b2)+"\n")
                return False
            if not b1:
                print("未发现文件变更\n")
                return True

def firstchk(mpath):
    #第一次计算所有文件的md5值
    with open('init.log', 'w+') as fp:
        for spath in mpath:
            file_path_list = get_file(spath)
            for file_path in file_path_list:
                #fp.write(file_path.replace(spath, '') + '|' + get_md5(file_path) + '\n')
                fp.write(file_path + '|' + get_md5(file_path) + '\n')


def nowcheck(mpath):
    #每次重新计算当前文件的md5值
    with open('new_md5.log', 'w+') as fp:
        for spath in mpath:
            file_path_list = get_file(spath)
            for file_path in file_path_list:
                #print(file_path)
                #fp.write(file_path.replace(spath, '') + '|' + get_md5(file_path) + '\n')
                fp.write(file_path + '|' + get_md5(file_path) + '\n')


if __name__ == '__main__':
    #web服务路径
    dirs=(r'D:\安全服务', r'C:\Tcl', r'D:\方案文档', r'C:\AppData', r'D:\PyCharm Community Edition 2019.3')
    firstchk(dirs)
    interval = 600
    print("成功获取到初始md5值，开始监控，每" + str(interval) + "秒进行一次扫描：")
    for i in range(999):
        time.sleep(interval)
        print(str(datetime.now()) + "第" + str(i) + "次对比中......", end="")
        nowcheck(dirs)
        compare("init.log", "new_md5.log")
        os.remove("new_md5.log")
