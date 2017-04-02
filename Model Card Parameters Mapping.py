# Update Date:  2016.11.15
# Author:       Grothendieck
# Version:      1.1.1 : Map any type of model card with any parameters
#               1.0.0 : Map Diode parameters to Cjgate model card
import os
import sys
import shutil
import xlrd
import re

def main(From,To,Table):

    def find(para,char):
        return re.search(r'\W' + para + r'\s*(=\s*-?\d+(\.\d+)?((e|E)(-|\+)\d+)?)', char)

    def replace(para,value,char):
        if '+' == find(para,char).group()[0]:
            return re.sub(r'\W' + para + r'\s*=\s*-?\d+(\.\d+)?((e|E)(-|\+)\d+)? *', ('+' + para).ljust(front,' ') + value.ljust(back,' '), char)
        else:
            return re.sub(r'\W' + para + r'\s*=\s*-?\d+(\.\d+)?((e|E)(-|\+)\d+)? *', (' ' + para).ljust(front,' ') + value.ljust(back,' '), char)

    def interval(mdl):
        m = re.search(r'(\+.*?)(=.*?=)',mdl)
        front = len(m.groups()[0])
        back = len(m.groups()[1]) - front
        return front,back

    def flatten(seq):
        s = str(seq).replace('[','').replace(']','')
        return [eval(x) for x in s.split(',') if x.strip()]

    def display(paraList,file):

        k,non = 0,[]

        for p in paraList:
            exist = False

            for i,line in enumerate(file):
                m = find(p,line)
                if m is not None:
                    exist = True
                    k += 1
                    print(' ',k,' >\t',p.ljust(10),'\t在该行被找到： ',i+1)
            if exist == False:
                print('  该参数未能找到: ',p)
                # collect the non-exist parameters
                non.append(p)
        # delete the non-exist parameters:
        for x in non: paraList.remove(x)
        return k, paraList

    #------------------------------------------------ Start ---------------------------------------------------------------------
    fFrom = open(From,'r').readlines()
    fTo = open(To,'r').readlines()
    fTable = wb.sheet_by_name(sheetName)
    # first column is source parameters:
    sourceP = fTable.col_values(0); mapCount = len(sourceP)
    # third column is target parameters:
    targetP = fTable.col_values(2)
    print('\n 从 Mapping Table 中显示你有 %d 对参数需要映射.\n'%mapCount)

    for i,p in enumerate(sourceP):
        if ',' in p:
            # e.g. convert 'is,js' to [is,js]:
            sourceP[i] = p.split(',')
    # e.g. [cj,mj,[is,js],n] -> [cj,mj,is,js,n]:
    sourceP_ = flatten(sourceP)

    show = display(sourceP_,fFrom) 
    k, sourceP_ = show[0], show[1] 
    if k != mapCount:
        print('\n Warning: 在源头 Model Card 中只找到了 %d 个参数, 请重新检查!'%k)
        input('\n press Enter to exit');exit()
    elif k == mapCount:
        print('\n 源头 Model Card 中所有的 %d 个参数已经找到。继续中......\n'%k)

    show = display(targetP,fTo)
    k = show[0]
    if k != mapCount:
        print('\n Warning: 在目标 Model Card 中只找到了 %d 个参数，请重新检查!'%k)
        input('\n press Enter to exit');exit()
    elif k == mapCount:
        print('\n 目标 Model Card 中所有的 %d 个参数已经找到。继续中......\n'%k)

    fsourceAll = open(From,'r').read()
    ftargetAll = open(To,'r').read()
    # interval is different between all kinds of format:
    front,back = interval(ftargetAll)[0],interval(ftargetAll)[1]
    for i,p in enumerate(sourceP_):
        m = find(p,fsourceAll)
        value = m.groups()[0]
        ftargetAll = replace(targetP[i], value, ftargetAll)
        print(' ',i+1,' >\t',m.group().lstrip('+| ').ljust(26,' '),'-------> ' + targetP[i].ljust(9, ' ') + value)
    #-----------------------------------------------------------------------------------------------------------------------------
    fnew = open(os.path.splitext(To)[0] + '_mapped' + os.path.splitext(To)[1],'w')
    fnew.write(ftargetAll)
    fnew.close()
    print("\n 映射完毕，已在此文件夹内生成一个新的 Model Card ，请查看。")
#------------------------------------------main function end--------------------------------------------------------------

if 1:
    print('''
        *********************************************
               Model Card Parameters Mapping
        *********************************************
        ''') 

    path = os.getcwd()
    #path = sys.path[0]
    
    #-------------------------------------------- Mapping Table File -----------------------------------------------------
    remoteDisk = r'\\tdtbde\TBDE\CompactModel2\NHD_Model2\Tools\Smic.go Repository'
    MapTable = 'Mapping Table.xlsx'

    def download():
        shutil.copyfile(remoteDisk + '/' + MapTable, path + '/' + MapTable)
        print('\n Mapping Table 已经下载到当前文件夹内，请检查。\n')

    AlreadyExist = os.path.exists(path + '/' + MapTable)
    NoneExist  = not AlreadyExist

    if NoneExist:
        ans = input(' 是否下载 Mapping Table ? \n\n y or n ? : ')
        download() if ans == 'y' else exit()
    if AlreadyExist:
        ans = input('\n 检测到你已经有 Mappint Table, 是否需要重新下载一份？\n\n y or n ? : ')
        download() if ans == 'y' else print('\n')

    ans = input(' 可以对 Mapping Table 进行修改，若无需修改，或已修改完毕，则按 y 继续。 \n\n 是否继续? : ')
    wb = xlrd.open_workbook(path + '/Mapping Table.xlsx') if ans == 'y' else exit()
    #--------------------------------------------------------------------------------------------------------------------

    #-------------------------------------------- Select Sheet ----------------------------------------------------------
    sheet_list = [line.name for line in wb.sheets()];print('\n')
    for i,sheet in enumerate(sheet_list):
        print(' ',i+1,'>',sheet)
    ans = input('\n 请选择需要做映射操作的参数所在的sheet名称 : ')
    sheetName = sheet_list[int(ans)-1]
    print('\n Your choice :\t\t%s \n'%sheetName)
    #--------------------------------------------------------------------------------------------------------------------

    #----------------------------------------- Select Model Card --------------------------------------------------------
    flist = os.listdir(path)
    dic = {str(i+1):f for i,f in enumerate(flist)}

    for i,f in enumerate(os.listdir(path)):
        print(' ',i+1,'>',f)

    def choice(file):
        ans = input('\n 请选择 %s : '%file)
        while ans not in dic:
            ans = input(' Please re-select : ')
        print('\n Your choice :\t\t%s'%dic[ans]);return dic[ans]

    ans1 = choice('源头 Model Card')
    ans2 = choice('目标 Model Card')
    #--------------------------------------------------------------------------------------------------------------------
    From,To,Table = path+'/'+ans1, path+'/'+ans2, path+'/Mapping Table.xlsx'
    main(From,To,Table)

if 0:
    if __name__=='__main__':
        path = sys.path[0]
        From = path+'/nmos_template.pm'
        To   = path+'/nmos_template.pm'
        Table = path+'/Mapping Table.xlsx'
        wb = xlrd.open_workbook(Table)
        sheetName = r'Cjgate Source to Drain'
        main(From,To,Table)
