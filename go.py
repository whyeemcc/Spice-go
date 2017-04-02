# Update Date:  2017.03.30
# Author:       Grothendieck
# Version:      1.12
import os
import sys

# where this script is running, pathNow is none value in cmd
pathNow = sys.path[0]

RunInServer = r'\tool new\Smic.Go' in pathNow
RunInLocal  = not RunInServer

#----------------------download----------------------
if RunInServer:
    # running in the server
    
    import re
    for dir in sys.path:
        m = re.search(r'site-packages$',dir)
        if m is not None:
            smicDir = dir + '\smic'
            break
    
    AlreadyExist = os.path.exists(smicDir)
    NoneExist = not AlreadyExist
    
    def installSmicGo():
        import shutil
        install = 'double-click to intall me.py'
        shutil.copyfile(pathNow + '/' + install, smicDir + '/go.py')
        
    def quit(str):
        input('\n ' + str);exit()
    
    def DoskeyRegedit():
        # command mapping
        doskey = smicDir + '/Doskey.bat'
        f = open(doskey,'w')
        f.write('@doskey smic.go=python -m smic.go'); f.close()
        # Regidit append
        os.system(r'REG ADD "HKCU\Software\Microsoft\Command Processor" /v AutoRun /t REG_SZ /d %s'%doskey)
    
    def installRecord():
        import time,socket
        time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        pcName = socket.gethostname()
        register = r'\\tdtbde\TBDE\CompactModel2\NHD_Model2\Tools\Smic.go Repository/Who have installed smic.go.ini'
        f = open(register,'a')
        f.write(pcName + ' ' + time + '\n')
        f.close()
        
    if NoneExist:
        # download firstly
        os.mkdir(smicDir)
        #----------------------smic.go----------------------
        installSmicGo()
        #------------------smic.go command------------------
        DoskeyRegedit()
        #----------------------register---------------------
        installRecord()
        #---------------------------------------------------
        quit('Install completed.')
        
    if AlreadyExist:
        # update stater: smic/go.py
        ver_loc = open(smicDir + '/go.py').readline()
        ver_ser = open(pathNow + '/double-click to intall me.py').readline()
        if ver_loc < ver_ser:
            print('\n This is a new version for SMIC.GO stater, updating...')
            installSmicGo()
            installRecord()
            quit('Update completed.')
        else:
            quit('This is the latest version. No need to update. press Enter to exit.')

if RunInLocal:
    # running in local disk
    #----------------------check update----------------------
    remoteDisk = r'\\tdtbde\TBDE\CompactModel2\NHD_Model2\Tools\Smic.go Repository'
    smicDir = os.path.split(os.path.realpath(__file__))[0]
    
    def Listset(folder):
        Listset = os.listdir(folder)
        Listset = {x for x in Listset if x[-3:] == '.py'}
        if 'go.py' in Listset: Listset.remove('go.py')	
        return Listset
        
    def download(module):
        import shutil
        shutil.copyfile(remoteDisk + '/' + module, smicDir + '/' + module)
        print('\n The new tool has been downloaded :\t %s'%module[:-3])
    
    def delete(module):
        os.remove(smicDir + '/' + module)
        print('\n The old tool has been deleted :\t %s'%module[:-3])
    
    def update(module):
        import shutil
        shutil.copyfile(remoteDisk + '/' + module, smicDir + '/' + module)
        print('\n The tool has been updated : \t %s'%module[:-3])
        
    smicList = Listset(smicDir)
    remoteList = Listset(remoteDisk) 
    
    signal = False
    try:
        if len(remoteList - smicList) is not 0:
            # download new tools
            signal = True
            newTool = remoteList - smicList
            for module in newTool: download(module)                
        if len(smicList - remoteList) is not 0:
            # delete old tools
            oldTool = smicList - remoteList
            for module in oldTool: delete(module)
    except:
        print("\n Oh! Can't download, please contact the maintenance personnel.");exit()

    smicList = Listset(smicDir)
    try:
        for module in smicList:
            # check version
            ver_loc = open(smicDir + '/' + module,'rb').readline()
            ver_ser = open(remoteDisk +'/' + module,'rb').readline()
            
            if ver_loc < ver_ser:
                signal = True
                update(module)
                
        display = {False:'No need to update',True:'Updated completely'}
        print('\n ******************\n %s \n ******************\n'%display[signal])
    except:
        print("\n Oh! Can't check update, please contact the maintenance personnel.")
    #----------------------update end-----------------------

    #----------------------select tool----------------------
    print('''
                SPICE TOOL KIT
    Choose a tool you want to use:
    ''')

    smicList = list(Listset(smicDir))
    smicList.sort()
    for i,module in enumerate(smicList):
        print(' ',i + 1,' > ',module[:-3])
    dic = {str(i+1):module for i,module in enumerate(smicList)}

    ans = input('\n select : ')
    while ans not in dic:
        ans = input(' re-select : ')

    f = smicDir + '/' + dic[ans]
    exec(open(f,'rb').read())