import xlwings as xw
import configparser
import os, shutil
import csv
import tkinter
from tkinter import filedialog
import pandas as pd
import time
now = time.localtime()
date = time.strftime("%m.%d",now)
cf = configparser.ConfigParser()

def mycopyfile(srcfile, dstfile):
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
    else:
        fpath, fname = os.path.split(dstfile)  # 分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)  # 创建路径
        shutil.copyfile(srcfile, dstfile)  # 复制文件
        print("copy %s -> %s" % (srcfile, dstfile))


# 全局窗口
wd = tkinter.Tk()
ct = tkinter.Label(wd, text='')
ct.pack()


def box():
    cf.read('ex.ini')
    count = cf.get('dir', 'count')
    wd.geometry('300x200')
    rp = tkinter.Button(wd, text='重置人工写入次数 ', command=rpc)
    choose = tkinter.Button(wd, text='选择文件', command=choosefile)
    label = tkinter.Button(wd, text = '制作标签表',command = icon)
    handle = tkinter.Button(wd, text = '定位支撑',command = handles)
    refresh = tkinter.Button(wd, text='重新生成导入及人工表格', command=newtwo)
    ct['text'] = count
    rp.pack()
    choose.pack()
    label.pack()
    handle.pack()
    refresh.pack()
    wd.mainloop()

def newtwo():
    mycopyfile(oneloc+'\\文档\\导入.csv', oneloc+ '\\导入.csv')

def rpc():
    os.chdir('E:\Pyxel')
    cf.set('dir', 'count', '3')
    with open('ex.ini', 'w+') as f:
        cf.write(f)
    ct['text'] = 3

def handles(): #定位支撑表
    os.chdir('E:\Pyxel')
    root = tkinter.Tk()
    filez = filedialog.askopenfilenames(parent=root, title='Choose a file',
                                        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    final = pd.DataFrame()
    for file in root.tk.splitlist(filez):
        print('打开' + file)
        op = dozc(file)
        final = final.append(op)

    final.to_excel(oneloc + '\\定位支撑' + '.xlsx',index=None, header = ['区间号','锚段号','编号','切割长度'])
    print('ALL OVER')



def icon():
    os.chdir('E:\Pyxel')
    root = tkinter.Tk()
    filez = filedialog.askopenfilenames(parent=root, title='Choose a file',
                                        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    orcsv = pd.DataFrame() #空dataframe
    ctcsv = pd.DataFrame() #统计吊弦数据
    for file in root.tk.splitlist(filez): #合并dataframe

        print('打开' + file)
        if '吊弦' in file:
            fi = dotense_old(file)
            orcsv = orcsv.append(fi[0]) 
            mode = 'tense'
            ctcsv = ctcsv.append(fi[1])
        else:
            fi = dolabel(file)
            orcsv = orcsv.append(fi)
            mode = 'label'

    if mode == 'tense':
        orcsv.to_csv(oneloc+'\\'+'吊弦标签.txt', index=None, header=None)
        ctcsv.to_csv(oneloc+'\\'+'吊弦统计.csv', index=None, header=None, mode = 'a')
    else:
        orcsv.to_csv(oneloc + '\\' + '腕臂标签.txt', index=None, header=None)
    print('ALL OVER')

def dotense_old(file):
    filepath, filename = os.path.split(file)
    ten = pd.read_excel(file, skiprows=2)
    length = int(len(ten))
    head = ten.columns.values.tolist()
    a = []
    count = 0
    for i in range(0, length, 5):
        for j in range(3, 18):
            if ten.iloc[i+4, j] > 0:
                a.append([ten.iloc[i, 0],ten.iloc[i+4,j], head[j]])
                count += 1
    s = pd.DataFrame(a)
    filesplit = place(filename)
    s[3] = filesplit[0]
    s[4] = filesplit[1]
    s = s.round(3)
    county = pd.DataFrame([[filesplit[0],filesplit[1],count]])
    return [s,county]

def dotense(file):
    filepath, filename = os.path.split(file)
    ten = pd.read_excel(file, skiprows=3, header = None)
    a = []
    stn = ''
    count = 0
    for i in range(0,ten.shape[0]):
        if pd.isna(ten.iloc[i,0]) == False:
            stn = ten.iloc[i,0]
        elif pd.isna(ten.iloc[i,1]) == False:
            a.append([stn,ten.iloc[i,5],ten.iloc[i,1]])
            count += 1
    s = pd.DataFrame(a)
    filesplit = place(filename)
    s[3] = filesplit[0]
    s[4] = filesplit[1]
    s = s.round(3)
    county = pd.DataFrame([[filesplit[0],filesplit[1],count]])
    return [s,county]

def place(filename):
    if filename[0:3] == '枣阳至' or  filename[0:3] == '随县至':
        if '锚' in filename[7:11]:
            return [filename[0:5], filename[7:10]]
        else:
            return [filename[0:5], filename[7:11]]

    elif filename[0:3] == '随县站' or filename[0:3] == '枣阳站':
        if '锚' in filename[3:7]:
            return [filename[0:3], filename[3:6]]
        else:
            return [filename[0:3], filename[3:7]]
    else:
        if '锚' in filename[8:12]:
            return [filename[0:6], filename[8:11]]
        else:
            return [filename[0:6], filename[8:12]]

def dolabel(file):
    filepath, filename = os.path.split(file)
    exc = pd.read_excel(file, header=None, skiprows=3, skipfooter=9)
    filesplit = place(filename)
    exc['qjh'] = filesplit[0]
    exc['mdh'] = filesplit[1]
    exc['diff'] = exc.apply(lambda row: diff(row), axis=1)
    exc['bar'] = exc.apply(lambda row: barcode(row), axis=1)
    qte = exc.dropna(how = 'any',thresh = 12)
    op = qte[[0, 'qjh', 'mdh',14,'diff','bar']]
    op = op.round(3)
    return op

def dozc(file):
    filepath, filename = os.path.split(file)
    exc = pd.read_excel(file, header =None, skiprows = 3, skipfooter = 9)
    filesplit = place(filename)
    exc['qjh'] = filesplit[0]
    exc['mdh'] = filesplit[1]
    exc = exc.dropna(how = 'any',thresh = 12)
    exc = exc.loc[abs(exc[15] - exc[2]) > 0.0015]
    exc[14] = exc[14] - 0.175
    op = exc[['qjh','mdh',0,14]]
    return op
    


def barcode(row):
    SN = str(row[0])
    SNN = SN.zfill(4)
    if '枣阳至襄阳' in row['qjh']:
        return 'ZY-XY'+ row['mdh']+'/'+SNN+'#'
    elif '随县至枣阳' in row['qjh']:
        return 'SX-ZY' + row['mdh']+'/'+SNN+'#'
    elif '随州南至随县' in row['qjh']:
        return 'SZN-SX' + row['mdh']+'/'+SNN+'#'
    elif '枣阳站' in row['qjh']:
        return 'ZYHCZ' + row['mdh']+'/'+SNN+'#'
    else:
        return 'SXHCZ' + row['mdh']+'/'+SNN+'#'

def diff(row):
    if abs(row[15] - row[2]) > 0.0015:
        return '支撑'
    else:
        return '吊线'


def quality(filename, loc, ori):  # 质检表
    quaname = filename.replace('装配数据', '质量检查')  # 替换名称
    mycopyfile('E:\Pyxel\汉十铁路智能化预配车间腕臂预配成品质量检查记录表.xlsx', loc + '\\' + quaname)
    qubook = xw.Book(loc + '\\' + quaname)
    i, j = 4, 8
    qua = qubook.sheets.active
    ora = ori.sheets.active
    quasplit = place(quaname)
    qua.range('M2').value = quasplit[0]
    qua.range('R2').value = quasplit[1]
    while i < 50:
        if ora.range('C' + str(i)).value != None:  # 判断C列是否为空
            if ora.range('Z' + str(i)).value != None:  # 判断是否需要平管帽
                qua.range('A' + str(j)).value = '*' + str(ora.range('A' + str(i)).value)
            else:
                qua.range('A' + str(j)).value = ora.range('A' + str(i)).value
            qua.range('N' + str(j)).value = ora.range('C' + str(i) + ':' + 'J' + str(i)).value
            qua.range('C' + str(j)).value = qua.range('C7:M7').value
            
            j += 1
        i += 1
    qubook.save() 
    qubook.close() 


def cut(filename, loc, ori):
    cutname = filename.replace('装配数据', '长度参照')
    mycopyfile('枣阳至襄阳区间2-28锚段装配数据10.26.xlsx', loc + '\\' + cutname)
    cutbook = xw.Book(loc + '\\' + cutname)
    cuta = cutbook.sheets.active  # 复制模板建立表格
    ora = ori.sheets.active
    cuta.range('A1').value = ora.range('A1').value
    i, j = 4, 4
    while i < 50:
        if ora.range('C' + str(i)).value != None:
            cuta.range('A' + str(j)).value = ora.range('A' + str(i) + ':' + 'F' + str(i)).value
            cuta.range('G' + str(j)).value = ora.range('G' + str(i)).value - 0.062  # 定位座
            cuta.range('H' + str(j)).value = ora.range('H' + str(i) + ':' + 'I' + str(i)).value
            cuta.range('J' + str(j)).value =float(ora.range('J' + str(i)).value[0:5]) - 0.175  # 定位管
            cuta.range('K' + str(j)).value = ora.range('L' + str(i)).value  # 腕臂支撑
            cuta.range('L' + str(j)).value = ora.range('M' + str(i) + ':' + 'N' + str(i)).value
            if abs(ora.range('P' + str(i)).value - ora.range('C' + str(i)).value) < 0.0015:  # 拉线和支撑
                cuta.range('N' + str(j)).value = ora.range('O' + str(i)).value
            else:
                cuta.range('O' + str(j)).value = ora.range('O' + str(i)).value - 0.175
            cuta.range('P' + str(j)).value = ora.range('P' + str(i)).value
            cuta.range('Q' + str(j)).value = ora.range('Q' + str(i)).value - 0.062  # 定位置
            cuta.range('R' + str(j)).value = ora.range('R' + str(i) + ':' + 'W' + str(i)).value
            cuta.range('X'+str(j)).value =ora.range('AA'+str(i)).value
            j += 1
        i += 1
    cutbook.save()
    cutbook.close()


def docsv(file, loc):  # csv
    filepath, filename = os.path.split(file)
    one = pd.read_excel(file, skiprows=3, skipfooter=9, header=None)
    norow = one.shape[0]
    cos = one.iloc[:, 0:24].dropna(how='any', thresh=12)
    split = place(filename)
    if '枣阳至襄阳' in filename:
        cos['qjh'] = 'ZY-XY'
    elif '随县至枣阳' in filename:
        cos['qjh'] = 'SX-ZY'
    elif '随州南' in filename:
        cos['qjh'] = 'SZN-SX'
    elif '枣阳站' in filename:
        cos['qjh'] = 'ZYHCZ'
    elif '地铁' in filename:
        cos['qjh'] = 'DTTCC'
    else:
        cos['qjh'] = 'SXHCZ'
    cos['mdh'] = split[1]
    cos['a'] = 'P181112-01' #套管座
    cos['b'] = 'P181115-01' #承力索座
    cos['c'] = 'P181115-46' #双耳管接头
    cos['d'] = 'P181112-13' #套管单耳
    cos['countall'] = norow
    cos['count'] = cos.shape[0]
    cos['alllabel'] = ','.join(one[0].astype(str))
    cos['label'] = ','.join(cos[0].astype(str))
    cos.to_csv(loc + '\\' + '导入.csv',header = None, index =None, mode = 'a')


def machine(filename, loc, ori):
    maname = filename.replace("装配数据", "机器数据")
    mycopyfile('腕臂预配机器数据.xlsx', loc + '\\' + maname)
    mabook = xw.Book(loc + '\\' + maname)  # 复制样板
    hubook = xw.Book(huloc)
    maa = mabook.sheets.active
    ora = ori.sheets.active
    hua = hubook.sheets.active
    maa.range('A1').value = ora.range('A1').value
    i, j = 4, 4
    cf = configparser.ConfigParser()
    cf.read('ex.ini')
    while i < 50:
        if ora.range('C' + str(i)).value != None:
            count = cf.get('dir', 'count')  # 取空行值，下方判断是否需要人工
            if float(ora.range('F' + str(i)).value[0:5]) < 0.19 or ora.range('H' + str(i)).value < 1.9 or ora.range(
                    'I' + str(i)).value < 2.0 or ora.range('N' + str(i)).value < 0.328 or ora.range(
                    'H' + str(i)).value > 3.582 or ora.range('I' + str(i)).value > 3.482 :
                if ora.range('Z' + str(i)).value != None:
                    hua.range('A' + count).value = ora.range('A1').value + '双管帽'
                else:
                    hua.range('A' + count).value = ora.range('A1').value
                hua.range('A' + str(int(count) + 1)).value = ora.range('A' + str(i) + ':' + 'K' + str(i)).value
                hua.range('L' + str(int(count) + 1)).value = ora.range('M' + str(i) + ':' + 'X' + str(i)).value
                ct = int(count) + 2
                os.chdir('E:\Pyxel')
                cf.set('dir', 'count', str(ct))
                with open('ex.ini', 'w+') as f:
                    cf.write(f)
            else:
                maa.range('A' + str(j)).value = ora.range('A' + str(i) + ':' + 'K' + str(i)).value
                maa.range('L' + str(j)).value = ora.range('M' + str(i) + ':' + 'X' + str(i)).value
                j += 1
        i += 1
    mabook.save()
    hubook.save()
    mabook.close()
    hubook.close()


qualoc = r'D:\OneDrive\腕臂质量校验'
cutloc = r'D:\OneDrive\腕臂长度参照'
oneloc = r'D:\OneDrive'
macloc = r'D:\OneDrive\机器生产信息'
huloc = r'D:\OneDrive\人工预配.xlsx'

def choosefile():
    os.chdir('E:\Pyxel')
    root = tkinter.Tk()
    filez = filedialog.askopenfilenames(parent=root, title='Choose a file',
                                        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    for file in root.tk.splitlist(filez):
        print('打开' + file)
        # xlfly(file)
        docsv(file,oneloc)
    print('ALL OVER')


def xlfly(file):

    ori = xw.Book(file)
    filepath, filename = os.path.split(file)
    quality(filename,qualoc,ori)
    cut(filename, cutloc, ori)
    machine(filename,macloc,ori)
    ori.app.kill()


if __name__ == "__main__":
    box()
