# Update Date:  2017.03.29
# Author:       Grothendieck
# Version:      1.0
import os
import sys
import random

def main(path,table):
    colors = '''#FFB6C1 #FFC0CB #DC143C #DB7093 #FF69B4 #FF1493 #C71585 #DA70D6 #D8BFD8 #DDA0DD #EE82EE
                #8B008B #800080 #BA55D3 #9400D3 #9932CC #4B0082 #8A2BE2 #9370DB #7B68EE #6A5ACD #483D8B 
                #0000FF #0000CD #191970 #00008B #000080 #4169E1 #6495ED #B0C4DE #778899 #708090 #1E90FF
                #87CEFA #87CEEB #00BFFF #ADD8E6 #B0E0E6 #5F9EA0 #AFEEEE #00FFFF #00FFFF #00CED1 #2F4F4F
                #008B8B #008080 #48D1CC #20B2AA #40E0D0 #7FFFAA #00FA9A #F5FFFA #3CB371 #2E8B57 #90EE90
                #98FB98 #8FBC8F #32CD32 #00FF00 #228B22 #008000 #006400 #7FFF00 #7CFC00 #ADFF2F #556B2F
                #FFFF00 #808000 #BDB76B #EEE8AA #F0E68C #FFD700 #FFF8DC #DAA520 #FF00FF #4682B4 #FA8072
                #F5DEB3 #FFE4B5 #FFA500 #FFEBCD #FFDEAD #D2B48C #DEB887 #FFE4C4 #FF8C00 #CD853F #800000
                #FFDAB9 #F4A460 #D2691E #8B4513 #A0522D #FFA07A #FF7F50 #FF4500 #E9967A #FF6347 #8B0000
                #F08080 #BC8F8F #CD5C5C #FF0000 #A52A2A #B22222 #D3D3D3'''
    colors = colors.split(' ')
    while '' in colors:
        colors.remove('')
    colors = [x.rstrip('\n') for x in colors]

    # user difined style
    # AA: red,strongly left-hand hatched dense
    # CT: Lime,cross hatched
    # GT: blue,strongly right-hand hatched dense
    # NW: purple,hollow
    # SN: Cyan,lightly left-hand hatched
    # SP: yellow,lightly left-hand hatched
    # M1: LightGrey,light unordered
    # M2: MediumPurple,light unordered
    dic = {
            'AA':('#FF0000','I6'),
            'CT':('#00FF00','I12'),
            'GT':('#0000FF','I10'),
            'NW':('#800080','I1'),
            'SN':('#00FFFF','I5'),
            'SP':('#FFFF00','I5'),
            'M1':('#D3D3D3','I31'),
            'M2':('#9370DB','I31')
          }

    for k,v in dic.items():
        colors.remove(v[0])

    def bl(x):
        return str(x).rstrip(' ').lstrip(' ')

    def style(layer):
        if layer in dic: 
            return dic[layer]
        else:
            color = colors[random.randint(0,len(colors)-1)]
            index = random.randint(2,46)
            return (color,'I'+str(index))    

    import xlrd
    table = path + '/' + table
    wb = xlrd.open_workbook(table)
    ws = wb.sheet_by_name('GDS')

    gdsName = [bl(x) for x in ws.col_values(0)[1:]]
    gdsNo = [bl(int(x)) for x in ws.col_values(1)[1:]]
    dataType = [bl(int(x)) for x in ws.col_values(2)[1:]]

    lyp = open(path + '/LayerProperties.lyp','w')
    lyp.write('<?xml version="1.0" encoding="utf-8"?>\n')
    lyp.write('<layer-properties>')
    for i,layer in enumerate(gdsName):
        NoType = gdsNo[i]+'/'+dataType[i]
        styles = style(layer)
        color,pattern = styles[0],styles[1]
        block = """
 <properties>
  <frame-color>%s</frame-color>
  <fill-color>%s</fill-color>
  <frame-brightness>0</frame-brightness>
  <fill-brightness>0</fill-brightness>
  <dither-pattern>%s</dither-pattern>
  <valid>true</valid>
  <visible>true</visible>
  <transparent>false</transparent>
  <width>1</width>
  <marked>false</marked>
  <animation>0</animation>
  <name>%s %s@1</name>
  <source>%s@1</source>
 </properties>"""%(color,color,pattern,layer,NoType,NoType)
        lyp.write(block)
    lyp.write('\n</layer-properties>')   
    lyp.close()

    print('\n .lyp文件已生成完毕。')

if 1:	
    print('''
        *********************************************
             Klayout Layer Properties Generator
        *********************************************
        ''')
    
    path = os.getcwd()
    #path = sys.path[0]

    #-------------------------------------------- Table -----------------------------------------------------
    remoteDisk = r'\\tdtbde\TBDE\CompactModel2\NHD_Model2\Tools\Smic.go Repository'	
    table = 'GDS Layer Mapping.xlsx'

    def download(f):
        import shutil
        shutil.copyfile(remoteDisk + '/' + f, path + '/' + f)
    
    AlreadyExist = os.path.exists(path + '/' + table)
    NoneExist = not AlreadyExist
    
    if NoneExist:
        ans = input('\n 是否下载 GDS Layer Mapping 表格 ? \n\n y or n ? : ')
        download(table) if ans == 'y' else exit()
    elif AlreadyExist:
        ans = input('\n 检测到你已经存在 GDS Layer Mapping 表格, 是否需要重新下载一份？\n\n y or n ? : ')
        download(table) if ans == 'y' else print('\n')
    
    ans = input('\n 请在表格内填入相关信息并保存，若无需修改，或已修改完毕，则按 y 继续。 \n\n 是否继续? : ')
    main(path,table) if ans == 'y' else exit()
    #--------------------------------------------------------------------------------------------------------------------

if 0:
    if __name__ == '__main__':
        pass