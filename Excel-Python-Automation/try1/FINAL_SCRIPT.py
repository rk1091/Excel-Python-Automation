import openpyxl
import glob

def adddata():
    a=[]
    b=[]
    a.append(worksheet['C13'].value)

    #CASE1 

    #conditions for revenues
    if ( (worksheet['E19'].value>=1649 and worksheet['E19'].value<=1651) and (worksheet['F19'].value>=1814 and worksheet['F19'].value<=1816) and (worksheet['G19'].value>=1995.5 and worksheet['G19'].value<=1997.5) and (worksheet['H19'].value>=2195.15 and worksheet['H19'].value<=2197.15) and (worksheet['I19'].value>=2414.77 and worksheet['I19'].value<=2416.77) and (worksheet['J19'].value>=2656.34 and worksheet['J19'].value<=2658.34) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for total costs
    if ( (worksheet['E20'].value>=974 and worksheet['E20'].value<=976) and (worksheet['F20'].value>=846.5 and worksheet['F20'].value<=848.5) and (worksheet['G20'].value>=738.13 and worksheet['G20'].value<=740.13) and (worksheet['H20'].value>=646.01 and worksheet['H20'].value<=648.01) and (worksheet['I20'].value>=567.71 and worksheet['I20'].value<=569.71) and (worksheet['J20'].value>=501.15 and worksheet['J20'].value<=503.15) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for taxes
    if ( (worksheet['E21'].value>=-114.75 and worksheet['E21'].value<=-112.75) and (worksheet['F21'].value>=-12.37 and worksheet['F21'].value<= -10.37) and (worksheet['G21'].value>=89.08 and worksheet['G21'].value<=91.08) and (worksheet['H21'].value>=191.2 and worksheet['H21'].value<=193.2) and (worksheet['I21'].value>=295.47 and worksheet['I21'].value<=297.47) and (worksheet['J21'].value>=403.32 and worksheet['J21'].value<=405.32) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for change in NWC
    if ( (worksheet['D22'].value>=-101 and worksheet['D22'].value<=-99) and (worksheet['E22'].value>=-148.5 and worksheet['E22'].value<=-146.5) and (worksheet['F22'].value>=-25.75 and worksheet['F22'].value<=-23.75) and (worksheet['G22'].value>=-28.23 and worksheet['G22'].value<=-26.23) and (worksheet['H22'].value>=-30.95 and worksheet['H22'].value<=-28.95) and (worksheet['I22'].value>=-33.94 and worksheet['I22'].value<=-31.94) and (worksheet['J22'].value>=-37.24 and worksheet['J22'].value<=-35.24) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for operating cash flows
    if ( (worksheet['E23'].value>=787.75 and worksheet['E23'].value<=789.75) and (worksheet['F23'].value>=977.88 and worksheet['F23'].value<=979.88) and (worksheet['G23'].value>=1166.29 and worksheet['G23'].value<=1168.29) and (worksheet['H23'].value>=1355.94 and worksheet['H23'].value<=1357.94) and (worksheet['I23'].value>=1549.59 and worksheet['I23'].value<=1551.59) and (worksheet['J23'].value>=1749.87 and worksheet['J23'].value<=1751.87) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for non-operating cash flows
    if ( (worksheet['D24'].value>=-6101 and worksheet['D24'].value<=-6099) and (worksheet['E24'].value>=-148.5 and worksheet['E24'].value<=-146.5) and (worksheet['F24'].value>=-25.75 and worksheet['F24'].value<=-23.75) and (worksheet['G24'].value>=-28.23 and worksheet['G24'].value<=-26.23) and (worksheet['H24'].value>=-30.95 and worksheet['H24'].value<=-28.95) and (worksheet['I24'].value>=-33.94 and worksheet['I24'].value<=-31.94) and (worksheet['J24'].value>=523.86 and worksheet['J24'].value<=525.86) ):
        a.append(1)
    else:
        a.append(0)

    #condition for IRR
    if (worksheet['D25'].value>=0.064 and worksheet['D25'].value<=0.065):
        a.append(1)
    else:
        a.append(0)

    #condition for accept/reject decision
    if (worksheet['D26'].value==2) :
        a.append(1)
    else:
        a.append(0)


    #CASE2

    #condition for IRR
    if (worksheet['D31'].value>=0.075 and worksheet['D31'].value<=0.076):
        a.append(1)
    else:
        a.append(0)

    #condition for accept/reject decision
    if (worksheet['D32'].value==2) :
        a.append(1)
    else:
        a.append(0)


    #CASE3 

    #conditions for revenues
    if ( (worksheet['E38'].value>=1649 and worksheet['E38'].value<=1651) and (worksheet['F38'].value>=1904.75 and worksheet['F38'].value<=1906.75) and (worksheet['G38'].value>=2200.14 and worksheet['G38'].value<=2202.14) and (worksheet['H38'].value>=2541.32 and worksheet['H38'].value<=2543.32) and (worksheet['I38'].value>=2935.38 and worksheet['I38'].value<=2937.38) and (worksheet['J38'].value>=3390.52 and worksheet['J38'].value<=3392.52) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for total costs
    if ( (worksheet['E39'].value>=1049 and worksheet['E39'].value<=1051) and (worksheet['F39'].value>=957.63 and worksheet['F39'].value<=959.63) and (worksheet['G39'].value>=995.56 and worksheet['G39'].value<=997.56) and (worksheet['H39'].value>=1035.38 and worksheet['H39'].value<=1037.38) and (worksheet['I39'].value>=1077.2 and worksheet['I39'].value<=1079.2) and (worksheet['J39'].value>=1121.11 and worksheet['J39'].value<=1123.11) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for taxes
    if ( (worksheet['E40'].value>=-85 and worksheet['E40'].value<=-83) and (worksheet['F40'].value>=-12.1 and worksheet['F40'].value<= -10.1) and (worksheet['G40'].value>=41.96 and worksheet['G40'].value<=43.96) and (worksheet['H40'].value>=105.25 and worksheet['H40'].value<=107.25) and (worksheet['I40'].value>=179.22 and worksheet['I40'].value<=181.22) and (worksheet['J40'].value>=265.57 and worksheet['J40'].value<=267.57) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for change in NWC
    if ( (worksheet['D41'].value>=-101 and worksheet['D41'].value<=-99) and (worksheet['E41'].value>=-148.5 and worksheet['E41'].value<=-146.5) and (worksheet['F41'].value>=-39.36 and worksheet['F41'].value<=-37.36) and (worksheet['G41'].value>=-45.31 and worksheet['G41'].value<=-43.31) and (worksheet['H41'].value>=-52.18 and worksheet['H41'].value<=-50.18) and (worksheet['I41'].value>=-60.11 and worksheet['I41'].value<=-58.11) and (worksheet['J41'].value>=439.46 and worksheet['J41'].value<=441.46) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for operating cash flows
    if ( (worksheet['E42'].value>=683 and worksheet['E42'].value<=685) and (worksheet['F42'].value>=957.23 and worksheet['F42'].value<=959.23) and (worksheet['G42'].value>=1160.62 and worksheet['G42'].value<=1162.62) and (worksheet['H42'].value>=1398.69 and worksheet['H42'].value<=1400.69) and (worksheet['I42'].value>=1676.96 and worksheet['I42'].value<=1678.96) and (worksheet['J42'].value>=2001.83 and worksheet['J42'].value<=2003.83) ):
        a.append(1)
    else:
        a.append(0)

    #conditions for non-operating cash flows
    if ( (worksheet['D43'].value>=-6101 and worksheet['D43'].value<=-6099) and (worksheet['E43'].value>=-148.5 and worksheet['E43'].value<=-146.5) and (worksheet['F43'].value>=-39.36 and worksheet['F43'].value<=-37.36) and (worksheet['G43'].value>=-45.31 and worksheet['G43'].value<=-43.31) and (worksheet['H43'].value>=-52.18 and worksheet['H43'].value<=-50.18) and (worksheet['I43'].value>=-60.11 and worksheet['I43'].value<=-58.11) and (worksheet['J43'].value>=439.46 and worksheet['J43'].value<=441.46) ):
        a.append(1)
    else:
        a.append(0)

    #condition for IRR
    if (worksheet['D44'].value>=0.066 and worksheet['D44'].value<=0.067):
        a.append(1)
    else:
        a.append(0)

    #condition for accept/reject decision
    if (worksheet['D45'].value==2) :
        a.append(1)
    else:
        a.append(0)


    #CASE4

    if (worksheet['D50'].value==2) :
        a.append(1)
    else:
        a.append(0)
    if (worksheet['F50'].value==2) :
        a.append(1)
    else:
        a.append(0)
    if (worksheet['H50'].value==2) :
        a.append(1)
    else:
        a.append(0)


    #BONUS QUES

    if (worksheet['D56'].value>=-126.85 and worksheet['D56'].value<=-124.85):
        a.append(1)
    else:
        a.append(0)
    if (worksheet['E56'].value>=130.35 and worksheet['E56'].value<=132.35):
        a.append(1)
    else:
        a.append(0)
    if (worksheet['F56'].value>=-73.04 and worksheet['F56'].value<=-71.04):
        a.append(1)
    else:
        a.append(0)
    if (worksheet['G56'].value>=339 and worksheet['G56'].value<=341):
        a.append(1)
    else:
        a.append(0)
    if (worksheet['H56'].value>=625.63 and worksheet['H56'].value<=627.63):
        a.append(1)
    else:
        a.append(0)
    if (worksheet['I56'].value>=410.53 and worksheet['I56'].value<=412.53):
        a.append(1)
    else:
        a.append(0)

    b.append(a)
    return b
    

try:
    c=[]
    for file in glob.glob('*.xlsx'):
        if (file!='OUTPUT SHEET.xlsx'):
            wb=openpyxl.load_workbook(file)
            worksheet= wb["Sheet1"]
            b=adddata()
            c=c+b 
    
    w=openpyxl.load_workbook('OUTPUT SHEET.xlsx')
    worksheet2= w["Sheet1"]
    for x in c:
        worksheet2.append(x)
    w.save('OUTPUT SHEET.xlsx')
    w.close()
    wb.close()


except FileNotFoundError:
    print('Please enter path of the outfile file correctly.')

except TypeError:
    print('Please make sure that the input excel files have data according to the correct format (i.e apart from name, input only numbers elsewhere).')

except NameError:
   print('Please make sure that the input excel files are in the same folder as this python script file and no extra output excel files ARE present in this folder.')

except:
   print('Please close the outfile file.')

