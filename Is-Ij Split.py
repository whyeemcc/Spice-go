# Update Date:  2016.09.01
# Version 1.0
import matplotlib.pyplot as pl
import numpy as np
import os,sys

def main(folder):

    datalist = os.listdir(folder)

    for file in datalist:

        color=['ro','bo']
        fig=pl.figure()

        data=open(folder + '/' + file,'r').readlines()

        lens=len(data)

        I1 = data[16].split(',')[1][:2]
        I2 = data[16].split(',')[2][:2]

        newdata = []
        for line in data[17:lens]:
            obj = list(map(float,line.split(',')))
            newdata.append(obj)

        x = [x[0] for x in newdata]
        y1 = [abs(y[1]) for y in newdata]
        y2 = [abs(y[2]) for y in newdata]

        pl.plot(x, y1, 'ro', label = I1)
        pl.plot(x, y2, 'bo', label = I2)

        pl.title(file.rstrip('.dat'))
        pl.xlabel('Vj',fontsize = 15)
        pl.ylabel('Is , Ij',fontsize = 15)
        pl.legend(loc = 'upper left')
        pl.yscale('log')
        pl.show()

        ans = input('Which I do you want to save?\n\n   1: %s   2: %s\n\n  '%(I1,I2))
        ans = int(ans)


        newFolder = os.path.dirname(folder) + '/' + os.path.basename(folder) + '__split'
        if os.path.exists(newFolder):
            pass
        else:
            os.mkdir(newFolder)

        fnew = open(newFolder + '/' + file,'w')
        fnew.writelines(data[:16])
        fnew.write('[Vj,%s(Vs=0.),Ref_Vs=0.]\n'%[I1,I2][ans-1])

        for line in data[17:lens]:
            V = line.split(',')[0]
            I = line.split(',')[ans]
            if ans == 2:
                I = I.rstrip('\n')
            fnew.write(V + ',' + I +'\n')
        fnew.close()


#***************************main function end**************************************
if 1:
    path = os.getcwd()
    #path = sys.path[0]

    print('''
        *********************************************
           Is-Ij Split (BsimProPlus format only)
        *********************************************
        ''')

    print("\n Select the folder that contain '.dat' files with Is-Ij both exist.\n") 

    flist = os.listdir(path) 
    dic = {str(i+1):f for i,f in enumerate(flist)}

    for i,f in enumerate(os.listdir(path)):
        print(' ',i+1,':',f)

    ans = input('\n select : ')
    while ans not in dic:
        ans = input(' Please re-select : ')
    print('\n Your choice :\t\t%s'%dic[ans])

    folder = path + '/' + dic[ans]

    main(folder) 

if 0:
    if __name__=='__main__':
        path = sys.path[0]
        folder = path + '/' + 'IV_25'
        main(folder)