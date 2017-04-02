# Update Date:  2016.12.23
# Author:       Grothendieck
# Version:      1.0.0   :   Basic function completed.
import os
import sys
import re

def main(folder):
    from matplotlib.backends.backend_pdf import PdfPages
    pdfName = os.path.basename(folder)
    pdf	= PdfPages(folder + '/' + pdfName + '_report.pdf')

    filelist = os.listdir(folder)
    filelist.remove(pdfName + '_report.pdf')
    
    for f in filelist:
        content = open(folder + '/' + f,'r').read()
        
        Instance = re.search('Instance{(.*)}',content).group()
        W = re.search('W=(\d+(\.\d+)?)',Instance).groups()[0]
        L = re.search('L=(\d+(\.\d+)?)',Instance).groups()[0]
        N = str(float(L)/float(W)).strip('.0')
        
        input = re.search('Input{(.*)}',content).groups()[0]                   # only one input
        output = re.search('Output{(.*)}',content).groups()[0].split(',')      # possible output more than one
        
        content = open(folder + '/' + f,'r').readlines()
        dic = {}
        x = input
        for y in output:
            for i,line in enumerate(content):
                if x+','+y in line:
                    dic[i] = y; break                                           # find the line number

        segment = sorted(list(dic.keys()))
        segment.append(len(content))
        
        import matplotlib.pyplot as plt
        fig	= plt.figure()	
        for i,loc in enumerate(segment[:-1]):
            begin = loc + 1
            end = segment[i+1]
            x,y = [],[]
            for line in content[begin : end]:
                obj = list(map(float,line.split(',')))
                x.append(obj[0])
                y.append(obj[1])

            pic = plt.subplot(2,2,i+1)
            pic.plot(x,y,'r.')
            plt.xlabel(input,fontsize = 14)
            plt.ylabel(dic[loc],fontsize = 14)
            plt.xticks(fontsize = 10)
            plt.yticks(fontsize = 10)
            plt.grid(True) 	
        fig.suptitle(f.rstrip('.dat') + '   W = %s, L = %s, N = %s'%(W,L,N))
        pdf.savefig(fig)
        plt.close(fig)
        
        print(' Done : \t',f)
    pdf.close()
    
if 1:
    path = os.getcwd()
    #path = sys.path[0]
        
    print('''
        **************************************
               Resistor Raw Data Plot
        **************************************
        ''')
        
    ans = input('''\n 是否将此目录下所有文件画出电阻的测量数据曲线？

 y or n ? : \t''')
    
    main(path) if ans == 'y' else exit()
    
if 0:
    if __name__=='__main__':
        path = sys.path[0]
        folder = sys.path[0] + '/' + 'RNPOSAB'
        main(folder)