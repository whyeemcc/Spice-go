# Update Date:  2016.11.14
# Author:       Grothendieck
# Version:      1.0
import os
import sys
import re
import shutil
import time

def main(folder):

    Cri = 'Diode_Corner_Parameter_Criterion.ini'
    remoteDisk = r'\\tdtbde\TBDE\CompactModel2\NHD_Model2\Tools\Smic.go Repository'
    FileOnServer = remoteDisk + '/' + Cri
    FileOnLocal  = folder + '/' + Cri
    AlreadyExist = os.path.exists(FileOnLocal)
    NoneExist   = not AlreadyExist

    def download():
        shutil.copyfile(FileOnServer, FileOnLocal)
        print('\n',open(FileOnLocal).read())
        print('''
 Corner参数的配置文件'Diode_Corner_Parameter_Criterion.ini'已下载，如上所示。
 若要修改，请更改此配置文件并保存，再继续。
 若不更改，请继续。''')

    if NoneExist:
        download()

    if AlreadyExist:
        print('\n',open(FileOnLocal).read())
        ans = input('''
 检测到当前文件夹已有'Diode_Corner_Parameter_Criterion.ini'配置文件，如上所示。
 是否需要重新下载？\n\n y or n ? : ''')
        download() if ans == 'y' else print('\n')

    ans = input('\n 确认是否开始合并 Model Card:\n\n y or n ? : ')
    print('\n') if ans == 'y' else exit()

#---------------------------------------- extract corner parameter and value from '.ini' file---------------------------------------
    f = open(FileOnLocal)
    Par,FF,SS = [],[],[]
    for line in f.readlines():
        # find the corner parameters' name: dxx:
        m = re.search('d\w+',line)
        if m is not None:
            Par.append(m.group())
            # find the corner parameters' value:
            n = re.findall('((-|\+)?\d+%)',line)
            FF.append(float(n[0][0].rstrip('%'))/100)
            SS.append(float(n[1][0].rstrip('%'))/100)
#-----------------------------------------------------------------------------------------------------------------------------------
    def find(para,char):
        return re.search(r'\W' + para + r'\s*=\s*(-?\d+(\.\d+)?((e|E)(-|\+)\d+)?)', char)

    def replace(para,corStr,char):
        # add corner string
        if '+' in find(para,char).group():
            return re.sub(r'\W' + para + r'(\s*=\s*-?\d+(\.\d+)?((e|E)(-|\+)\d+)?) *', ('+' + para).ljust(9,' ') + corStr, char)
        else:
            return re.sub(r'\W' + para + r'(\s*=\s*-?\d+(\.\d+)?((e|E)(-|\+)\d+)?) *', (' ' + para).ljust(9,' ') + corStr, char)

    mocalist = os.listdir(folder)
    # find out all '.pm' files
    mocalist = [x for x in mocalist if x[-3:] == '.pm']

    mdlAll = []
    namAll = []
    corAll = {}

    for i,moca in enumerate(mocalist):

        content = open(folder + '/' + moca,'r').read()
        # extrac the model name from each model card:
        moName = re.search('\.model.*d',content)
        moName = moName.group()
        moName = moName.lstrip('.model').rstrip('d')
        moName = moName.strip(' ')

        TTbox,FFbox,SSbox = {},{},{}

        for j,p in enumerate(Par):
            p = p.lstrip('d')
            # find each parameter in the model card
            m = find(p,content)

            if m is not None:
                pass
            else:
                # in BsimProPlus, cjp = cjsw, if there is no 'cjsw', then search 'cjp'
                m = find('cjp',content)
                # substitute 'cjp' to 'cjsw'
                content = re.sub('cjp ','cjsw',content)
                print(" 该model中参数名 'cjp' 已在合并的model中替换为 'cjsw' : %s" %moca)
            # extract the TT value: like (4e-007) from (+is      = 4e-007):
            V = m.groups()[0]
            # like: (= '4e-007+dis_ndio33') from (+is      = '4e-007+dis_ndio33'):
            corStr = "= '%s+d%s_%s'     "%(V,p,moName)
            # a new content waiting for merge:
            content = replace(p,corStr,content)
            # FF/SS ratio
            FFratio = FF[j]
            SSratio = SS[j]
            # TT/FF/SS value difference
            TTdiff = '0'
            FFdiff = format(eval(V) * FFratio,'.3e')
            SSdiff = format(eval(V) * SSratio,'.3e')

            TTbox['d'+p] = TTdiff
            FFbox['d'+p] = FFdiff
            SSbox['d'+p] = SSdiff

        mdlAll.append(content)
        mdlAll.append('\n')
        namAll.append(moName)
        # parameters prepare for .lib file:
        corAll[moName] = {'tt':TTbox,'ff':FFbox,'ss':SSbox}
        # corAll include several model_cards, each model_card include 3 corner, each corner_box include all corner_parameters

    mergeName = os.path.basename(folder)
    date = time.strftime("%Y-%m-%d-%H-%M",time.localtime())
    newfolder = folder + '/Model ' + date
    if os.path.exists(newfolder):
        pass
    else:
        os.mkdir(newfolder)
    # create Hspice .mdl file
    f_mdl = open(newfolder + '/' + mergeName + '.mdl','w')
    # create Hspice .lib file
    f_lib = open(newfolder + '/' + mergeName + '.lib','w')
    # create Spectre .mdl file
    f_mdl_spe = open(newfolder + '/' + mergeName + '_spe.mdl','w')
    # create Spectre .lib file
    f_lib_spe = open(newfolder + '/' + mergeName + '_spe.lib','w')


    #-----------------------------write in Hspice .mdl file----------------------------------------
    f_mdl.writelines(mdlAll)
    f_mdl.close()
    #-----------------------------------------end--------------------------------------------------


    #-----------------------------write in Hspice .lib file----------------------------------------
    f_lib.write('''***********************
*  diode corner model *
***********************

''')
    # three corners
    for cor in ['tt','ff','ss']:
        f_lib.write('''.lib dio_%s
.option scale=1 gmindc=1e-14
.param

'''%cor)
        # model cards
        for mod in namAll:
            # like: * dnwdio33(tt corner)
            f_lib.write('* %s(%s corner)\n'%(mod,cor))
            # parameters
            for par in Par:
                # like: +dis_dnwdio33 = 0
                f_lib.write('+%s_%s = %s\n'%(par,mod,corAll[mod][cor][par]))
            f_lib.write('\n')
        # like: .inc '45nor_diode.mdl':
        f_lib.write(".inc '%s.mdl'\n"%mergeName)
        # like: .endl dio_tt:
        f_lib.write('.endl dio_%s\n'%cor)
        f_lib.write('\n')
    f_lib.close()
    #------------------------------------------end----------------------------------------------------



    #-------------------------------write in Spectre .lib file----------------------------------------
    f_lib_spe.write('''// * Spectre Format by smic.go
simulator lang=spectre  insensitive=yes

// ***********************
// *  diode corner model *
// ***********************



''')

    # like: library 45nor_diode
    f_lib_spe.write('library %s\n\n\n'%mergeName)
    # three corners
    for cor in ['tt','ff','ss']:
        f_lib_spe.write('''section dio_%s
setoption1 options scale=1 gmin=1e-14\n'''%cor)

        for mod in namAll:
            for par in Par:
                # like: parameters dis_ndio33 = 0
                f_lib_spe.write('parameters %s_%s = %s\n'%(par,mod,corAll[mod][cor][par]))
        # like: include "45nor_diode_spe.mdl":
        f_lib_spe.write('include "%s_spe.mdl"\n'%mergeName)
        # like: endsection dio_tt:
        f_lib_spe.write('endsection dio_%s\n'%cor)
        f_lib_spe.write('\n\n\n')
    # like: endlibrary 45nor_diode:
    f_lib_spe.write('endlibrary %s'%mergeName)
    f_lib_spe.close()

    #----------------------------------------------end---------------------------------------------------




    #-----------------------------------write in Spectre .mdl file---------------------------------------
    #          Spectre : Hspice
    SpeToHsp = {'perim': 'pj',\
                'tnom' : 'tref',\
                'isw'  : 'jsw',\
                'ns'   : 'n',\
                'vj'   : 'pb',\
                'vjsw' : 'php',\
                'pta'  : 'tpb',\
                'ptp'  : 'tphp'}

    def writeSpe(Pspe):

        if Pspe == '\n':
            # add a \n
            f_mdl_spe.write('\n')
        else:
            # the parameters' name difference between Hspice & Spectre
            if Pspe.lstrip('+') in SpeToHsp:
                Phsp = SpeToHsp[Pspe.lstrip('+')]
            else:
                # Except the differ parameters, the name is equal
                Phsp = Pspe.lstrip('+')

            ThisIsCorner = 'd'+Phsp in Par
            ThisIsNotCorner = not ThisIsCorner

            if ThisIsCorner:

                m = re.search(r'\W' + Phsp + r"\s*=\s*('.*d%s_%s')"%(Phsp,mod),content)
                # like: (4e-007+dis_ndio33) from (+is      = '4e-007+dis_ndio33')
                V = m.groups()[0].lstrip("'").rstrip("'")
                # like: (+is           = 4e-007+dis_ndio33)
                f_mdl_spe.write(Pspe.ljust(14,' ') + '= %s'%V)

            if ThisIsNotCorner:
                m = find(Phsp,content)
                V = m.groups()[0]
                # level in spectre is fixed in 1
                if Phsp == 'level': V = '1'
                # tcv in hspice and spectre is reversed
                elif Phsp == 'tcv': V = str(-1 * eval(V))

                if '+' in Pspe:
                    # like: (+vb           = 12.76            )
                    f_mdl_spe.write(Pspe.ljust(14,' ') + ('= %s'%V).ljust(19,' '))
                if '+' not in Pspe:
                    # like: ( fc           = 0                )
                    f_mdl_spe.write((' '+Pspe).ljust(14,' ') + ('= %s'%V).ljust(19,' '))



    f_mdl_spe.write('''// * Spectre Format by smic.go
simulator lang=spectre  insensitive=yes

// **


''')

    for i,mod in enumerate(namAll):

        content = mdlAll[2*i]
        # like: model ndio33 diode:
        f_mdl_spe.write('model %s diode\n'%mod)
        f_mdl_spe.write('+allow_scaling = yes dskip = no imax = 1e20  minr = 1e-6\n')
        f_mdl_spe.write('''// ***************************************************************************
// *general parameters
// ***************************************************************************\n''')
        writeSpe('+level'); writeSpe('area'); writeSpe('perim'); writeSpe('\n')
        writeSpe('+tnom'); writeSpe('\n')
        f_mdl_spe.write('''// ***************************************************************************
// *dc parameters
// ***************************************************************************\n''')
        writeSpe('+ibv'); writeSpe('\n')
        writeSpe('+is'); writeSpe('\n')
        writeSpe('+isw'); writeSpe('\n')
        writeSpe('+n'); writeSpe('\n')
        writeSpe('+ns'); writeSpe('\n')
        writeSpe('+rs'); writeSpe('\n')
        writeSpe('+vb'); writeSpe('jtun'); writeSpe('jtunsw'); writeSpe('\n')
        writeSpe('+ntun'); writeSpe('\n')
        f_mdl_spe.write('''// ***************************************************************************
// *capacitance parameters
// ***************************************************************************\n''')
        writeSpe('+cj'); writeSpe('\n')
        writeSpe('+cjsw'); writeSpe('\n')
        writeSpe('+vj'); writeSpe('vjsw'); writeSpe('mj'); writeSpe('\n')
        writeSpe('+mjsw'); writeSpe('fc'); writeSpe('fcs'); writeSpe('\n')
        f_mdl_spe.write('''// ***************************************************************************
// *temperature parameters
// ***************************************************************************\n''')
        writeSpe('+tlev'); writeSpe('tlevc'); writeSpe('eg'); writeSpe('\n')
        writeSpe('+xti'); writeSpe('cta'); writeSpe('ctp'); writeSpe('\n')
        writeSpe('+tcv'); writeSpe('pta'); writeSpe('ptp'); writeSpe('\n')
        writeSpe('+trs'); writeSpe('xtitun'); writeSpe('\n')

        writeSpe('\n')
        writeSpe('\n')
        writeSpe('\n')
        writeSpe('\n')
    f_mdl_spe.close()
    #----------------------------------------------end----------------------------------------------------
    print('\n Completed :\t %s.lib'%mergeName)
    print('\n Completed :\t %s.mdl'%mergeName)
    print('\n Completed :\t %s_spe.lib'%mergeName)
    print('\n Completed :\t %s_spe.lib'%mergeName) 
#--------------------------------------------main function end--------------------------------------------

if 1:
    print('''
        ****************************************
          Diode Model Merge to Hspice & Spectre
        ****************************************
        ''') 

    path = os.getcwd()
    #path = sys.path[0]

    flist = [x for x in os.listdir(path) if x[-3:]=='.pm']
    mdlNum = len(flist)

    ans = input('''\n 是否将此目录下的 %d 个 Model Card 合并？

 y or n ? : \t'''%mdlNum)

    main(path) if ans == 'y' else exit()

if 0:
    if __name__=='__main__':
        path = sys.path[0]
        folder = r'D:\一Python Script\smic.go\Diode Model Merge to Hspice & Spectre\新建文件夹\28hk_diode'
        main(folder)
    