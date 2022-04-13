# -*- coding: utf-8 -*-
"""
Created on Mon Oct 14 10:36:13 2019

@author: PC
"""

import xlsxwriter  
import pandas as pd
workbook = xlsxwriter.Workbook('Bahman_ekhtesasi.xlsx')  
df = pd.read_excel (r'C:\Users\PC\Desktop\EPG Bahman.xlsx', sheet_name='اختصاصی')
worksheet = workbook.add_worksheet() 
###########################################################

#esteghlal=pd.DataFrame()
#tva=pd.DataFrame()
tva_sport=pd.DataFrame()
tva_sport_two=pd.DataFrame()
tva_kodak=pd.DataFrame()
digiton=pd.DataFrame()
lenz_sport=pd.DataFrame()
lenz_sport_plus=pd.DataFrame()
sarbaz_maher=pd.DataFrame()
shaparak=pd.DataFrame()
#tva_avand=pd.DataFrame()
#tva_two=pd.DataFrame()
#tva_film=pd.DataFrame()
#tva_nava=pd.DataFrame()
#tva_one=pd.DataFrame()
#mahfel=pd.DataFrame()
#shetab=pd.DataFrame()
#KarvanEshgh2=pd.DataFrame()
#konsertReyvandi=pd.DataFrame()
#perspolis=pd.DataFrame()

p1=0
p2=0
p3=0
p4=0
p5=0
p6=0
p7=0
p8=0
p9=0
p10=0
p11=0
p12=0
p13=0
p14=0
p15=0
p16=0
p17=0
p18=0
p19=0
p20=0

df5=df.groupby(['عنوان برنامه','نام شبکه']).sum().reset_index()
t=len(df5)
for i in range(0,t):
    f=df5.loc[i,'نام شبکه']
    
#####################################################################
######################### channels data #############################
#####################################################################

############################# شبکه 1 #################################
#    if f=='استقلال':
#        p1=p1+1  
#        esteghlal.loc[p1,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        esteghlal.loc[p1,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        esteghlal.loc[p1,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شبکه 2 #################################
#    if f=='تیوا': 
#        p2=p2+1 
#        tva.loc[p2,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva.loc[p2,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva.loc[p2,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################## شبکه 3 #################################
    if f=='تیوا اسپورت':
        p3=p3+1  
        tva_sport.loc[p3,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        tva_sport.loc[p3,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        tva_sport.loc[p3,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################### شبکه 4 #################################
    if f=='تیوا اسپورت دو':
        p4=p4+1 
        tva_sport_two.loc[p4,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        tva_sport_two.loc[p4,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        tva_sport_two.loc[p4,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################### شبکه 5 #################################
    if f=='تیوا کودک':
        p5=p5+1  
        tva_kodak.loc[p5,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        tva_kodak.loc[p5,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        tva_kodak.loc[p5,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################## خبر #################################
    if f=='کودک دیجیتون':  
        p6=p6+1 
        digiton.loc[p6,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        digiton.loc[p6,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        digiton.loc[p6,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################## افق #################################
    if f=='لنزاسپورت':
        p7=p7+1  
        lenz_sport.loc[p7,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        lenz_sport.loc[p7,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        lenz_sport.loc[p7,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################## افق #################################
    if f=='لنز اسپورت پلاس':
        p8=p8+1  
        lenz_sport_plus.loc[p8,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        lenz_sport_plus.loc[p8,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        lenz_sport_plus.loc[p8,'مدت بازدید']=df5.loc[i,'مدت بازدید']
        ############################## افق #################################
#    if f=='سرباز ماهر':
#        p9=p9+1  
#        sarbaz_maher.loc[p9,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        sarbaz_maher.loc[p9,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        sarbaz_maher.loc[p9,'مدت بازدید']=df5.loc[i,'مدت بازدید']
        ############################## افق #################################
    if f=='شاپرک':
        p10=p10+1  
        shaparak.loc[p10,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        shaparak.loc[p10,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        shaparak.loc[p10,'مدت بازدید']=df5.loc[i,'مدت بازدید']
         ############################## افق #################################
#    if f=='تیوا آوند':
#        p11=p11+1  
#        tva_avand.loc[p11,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva_avand.loc[p11,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva_avand.loc[p11,'مدت بازدید']=df5.loc[i,'مدت بازدید']
#         ############################## افق #################################
#    if f=='تیوا دو':
#        p12=p12+1  
#        tva_two.loc[p12,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva_two.loc[p12,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva_two.loc[p12,'مدت بازدید']=df5.loc[i,'مدت بازدید']
#         ############################## افق #################################
#    if f=='تیوا فیلم':
#        p13=p13+1  
#        tva_film.loc[p13,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva_film.loc[p13,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva_film.loc[p13,'مدت بازدید']=df5.loc[i,'مدت بازدید']
#         ############################## افق #################################
#    if f=='تیوا نوا':
#        p14=p14+1  
#        tva_nava.loc[p14,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva_nava.loc[p14,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva_nava.loc[p14,'مدت بازدید']=df5.loc[i,'مدت بازدید']
#         ############################## افق #################################
#    if f=='تیوا یک':
#        p15=p15+1  
#        tva_one.loc[p15,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        tva_one.loc[p15,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        tva_one.loc[p15,'مدت بازدید']=df5.loc[i,'مدت بازدید']
##         ############################## افق #################################
#    if f=='محفل':
#        p16=p16+1  
#        mahfel.loc[p16,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        mahfel.loc[p16,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        mahfel.loc[p16,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################## شبکه 1 #################################
#    if f=='پرسپولیس':
#        p17=p17+1 
#        perspolis.loc[p17,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        perspolis.loc[p17,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        perspolis.loc[p17,'مدت بازدید']=df5.loc[i,'مدت بازدید']      

############################# شبکه 1 #################################
#    if f=='شتاب':
#        p18=p18+1  
#        shetab.loc[p18,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        shetab.loc[p18,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        shetab.loc[p18,'مدت بازدید']=df5.loc[i,'مدت بازدید']      

############################# شبکه 1 #################################
#    if f=='کاروان عشق ۲':
#        p19=p19+1  
#        KarvanEshgh2.loc[p19,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        KarvanEshgh2.loc[p19,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        KarvanEshgh2.loc[p19,'مدت بازدید']=df5.loc[i,'مدت بازدید']      

############################# شبکه 1 #################################
#    if f=='کنسرت خنده حسن ریوندی':
#        p20=p20+1  
#        konsertReyvandi.loc[p20,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        konsertReyvandi.loc[p20,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        konsertReyvandi.loc[p20,'مدت بازدید']=df5.loc[i,'مدت بازدید']      
#####################################################################
######################### append data ###############################
#####################################################################
        
############################# شبکه 1 #################################
#esteghlal1=[]
#esteghlal2=[]
#esteghlal3=[]
#esteghlal4=[]
#esteghlal1=esteghlal["عنوان برنامه"].tolist()
#esteghlal4.append(esteghlal1)
#esteghlal2=esteghlal["تعداد بازدید"].tolist()
#esteghlal4.append(esteghlal2)
#esteghlal3=esteghlal["مدت بازدید"].tolist()
#esteghlal4.append(esteghlal3)
############################## شبکه 2 #################################
#tva1=[]
#tva2=[]
#tva3=[]
#tva4=[]
#tva1=tva["عنوان برنامه"].tolist()
#tva4.append(tva1)
#tva2=tva["تعداد بازدید"].tolist()
#tva4.append(tva2)
#tva3=tva["مدت بازدید"].tolist()
#tva4.append(tva3)
############################## شبکه 3 #################################
tva_sport1=[]
tva_sport2=[]
tva_sport3=[]
tva_sport4=[]
tva_sport1=tva_sport["عنوان برنامه"].tolist()
tva_sport4.append(tva_sport1)
tva_sport2=tva_sport["تعداد بازدید"].tolist()
tva_sport4.append(tva_sport2)
tva_sport3=tva_sport["مدت بازدید"].tolist()
tva_sport4.append(tva_sport3)
############################### شبکه 4 #################################
tva_sport_two1=[]
tva_sport_two2=[]
tva_sport_two3=[]
tva_sport_two4=[]
tva_sport_two1=tva_sport_two["عنوان برنامه"].tolist()
tva_sport_two4.append(tva_sport_two1)
tva_sport_two2=tva_sport_two["تعداد بازدید"].tolist()
tva_sport_two4.append(tva_sport_two2)
tva_sport_two3=tva_sport_two["مدت بازدید"].tolist()
tva_sport_two4.append(tva_sport_two3)
############################### شبکه 5 #################################
tva_kodak1=[]
tva_kodak2=[]
tva_kodak3=[]
tva_kodak4=[]
tva_kodak1=tva_kodak["عنوان برنامه"].tolist()
tva_kodak4.append(tva_kodak1)
tva_kodak2=tva_kodak["تعداد بازدید"].tolist()
tva_kodak4.append(tva_kodak2)
tva_kodak3=tva_kodak["مدت بازدید"].tolist()
tva_kodak4.append(tva_kodak3)
############################## خبر #################################
digiton1=[]
digiton2=[]
digiton3=[]
digiton4=[]
digiton1=digiton["عنوان برنامه"].tolist()
digiton4.append(digiton1)
digiton2=digiton["تعداد بازدید"].tolist()
digiton4.append(digiton2)
digiton3=digiton["مدت بازدید"].tolist()
digiton4.append(digiton3)
############################## افق #################################
lenz_sport1=[]
lenz_sport2=[]
lenz_sport3=[]
lenz_sport4=[]
lenz_sport1=lenz_sport["عنوان برنامه"].tolist()
lenz_sport4.append(lenz_sport1)
lenz_sport2=lenz_sport["تعداد بازدید"].tolist()
lenz_sport4.append(lenz_sport2)
lenz_sport3=lenz_sport["مدت بازدید"].tolist()
lenz_sport4.append(lenz_sport3)
############################## پویا #################################
lenz_sport_plus1=[]
lenz_sport_plus2=[]
lenz_sport_plus3=[]
lenz_sport_plus4=[]
lenz_sport_plus1=lenz_sport_plus["عنوان برنامه"].tolist()
lenz_sport_plus4.append(lenz_sport_plus1)
lenz_sport_plus2=lenz_sport_plus["تعداد بازدید"].tolist()
lenz_sport_plus4.append(lenz_sport_plus2)
lenz_sport_plus3=lenz_sport_plus["مدت بازدید"].tolist()
lenz_sport_plus4.append(lenz_sport_plus3)
############################## پویا #################################
#sarbaz_maher1=[]
#sarbaz_maher2=[]
#sarbaz_maher3=[]
#sarbaz_maher4=[]
#sarbaz_maher1=sarbaz_maher["عنوان برنامه"].tolist()
#sarbaz_maher4.append(sarbaz_maher1)
#sarbaz_maher2=sarbaz_maher["تعداد بازدید"].tolist()
#sarbaz_maher4.append(sarbaz_maher2)
#sarbaz_maher3=sarbaz_maher["مدت بازدید"].tolist()
#sarbaz_maher4.append(sarbaz_maher3)
############################## پویا #################################
shaparak1=[]
shaparak2=[]
shaparak3=[]
shaparak4=[]
shaparak1=shaparak["عنوان برنامه"].tolist()
shaparak4.append(shaparak1)
shaparak2=shaparak["تعداد بازدید"].tolist()
shaparak4.append(shaparak2)
shaparak3=shaparak["مدت بازدید"].tolist()
shaparak4.append(shaparak3)
############################## پویا #################################
#tva_avand1=[]
#tva_avand2=[]
#tva_avand3=[]
#tva_avand4=[]
#tva_avand1=tva_avand["عنوان برنامه"].tolist()
#tva_avand4.append(tva_avand1)
#tva_avand2=tva_avand["تعداد بازدید"].tolist()
#tva_avand4.append(tva_avand2)
#tva_avand3=tva_avand["مدت بازدید"].tolist()
#tva_avand4.append(tva_avand3)
############################### پویا #################################
#tva_two1=[]
#tva_two2=[]
#tva_two3=[]
#tva_two4=[]
#tva_two1=tva_two["عنوان برنامه"].tolist()
#tva_two4.append(tva_two1)
#tva_two2=tva_two["تعداد بازدید"].tolist()
#tva_two4.append(tva_two2)
#tva_two3=tva_two["مدت بازدید"].tolist()
#tva_two4.append(tva_two3)
############################### پویا #################################
#tva_film1=[]
#tva_film2=[]
#tva_film3=[]
#tva_film4=[]
#tva_film1=tva_film["عنوان برنامه"].tolist()
#tva_film4.append(tva_film1)
#tva_film2=tva_film["تعداد بازدید"].tolist()
#tva_film4.append(tva_film2)
#tva_film3=tva_film["مدت بازدید"].tolist()
#tva_film4.append(tva_film3)
############################### پویا #################################
#tva_nava1=[]
#tva_nava2=[]
#tva_nava3=[]
#tva_nava4=[]
#tva_nava1=tva_nava["عنوان برنامه"].tolist()
#tva_nava4.append(tva_nava1)
#tva_nava2=tva_nava["تعداد بازدید"].tolist()
#tva_nava4.append(tva_nava2)
#tva_nava3=tva_nava["مدت بازدید"].tolist()
#tva_nava4.append(tva_nava3)
############################### پویا #################################
#tva_one1=[]
#tva_one2=[]
#tva_one3=[]
#tva_one4=[]
#tva_one1=tva_one["عنوان برنامه"].tolist()
#tva_one4.append(tva_one1)
#tva_one2=tva_one["تعداد بازدید"].tolist()
#tva_one4.append(tva_one2)
#tva_one3=tva_one["مدت بازدید"].tolist()
#tva_one4.append(tva_one3)
############################## پویا #################################
#mahfel1=[]
#mahfel2=[]
#mahfel3=[]
#mahfel4=[]
#mahfel1=mahfel["عنوان برنامه"].tolist()
#mahfel4.append(mahfel1)
#mahfel2=mahfel["تعداد بازدید"].tolist()
#mahfel4.append(mahfel2)
#mahfel3=mahfel["مدت بازدید"].tolist()
#mahfel4.append(mahfel3)
############################## پویا #################################
#perspolis1=[]
#perspolis2=[]
#perspolis3=[]
#perspolis4=[]
#perspolis1=perspolis["عنوان برنامه"].tolist()
#perspolis4.append(perspolis1)
#perspolis2=perspolis["تعداد بازدید"].tolist()
#perspolis4.append(perspolis2)
#perspolis3=perspolis["مدت بازدید"].tolist()
#perspolis4.append(perspolis3)
############################## پویا #################################
#shetab1=[]
#shetab2=[]
#shetab3=[]
#shetab4=[]
#shetab1=shetab["عنوان برنامه"].tolist()
#shetab4.append(shetab1)
#shetab2=shetab["تعداد بازدید"].tolist()
#shetab4.append(shetab2)
#shetab3=shetab["مدت بازدید"].tolist()
#shetab4.append(shetab3)
############################### پویا #################################
#KarvanEshgh21=[]
#KarvanEshgh22=[]
#KarvanEshgh23=[]
#KarvanEshgh24=[]
#KarvanEshgh21=KarvanEshgh2["عنوان برنامه"].tolist()
#KarvanEshgh24.append(KarvanEshgh21)
#KarvanEshgh22=KarvanEshgh2["تعداد بازدید"].tolist()
#KarvanEshgh24.append(KarvanEshgh22)
#KarvanEshgh23=KarvanEshgh2["مدت بازدید"].tolist()
#KarvanEshgh24.append(KarvanEshgh23)
############################### پویا #################################
#konsertReyvandi1=[]
#konsertReyvandi2=[]
#konsertReyvandi3=[]
#konsertReyvandi4=[]
#konsertReyvandi1=konsertReyvandi["عنوان برنامه"].tolist()
#konsertReyvandi4.append(konsertReyvandi1)
#konsertReyvandi2=konsertReyvandi["تعداد بازدید"].tolist()
#konsertReyvandi4.append(konsertReyvandi2)
#konsertReyvandi3=konsertReyvandi["مدت بازدید"].tolist()
#konsertReyvandi4.append(konsertReyvandi3)
##################################################################
#ofogh4=ofogh4.sort_values(["تعداد بازدید افق" , "شبکه افق بازدید"], ascending=[True,False])
bold = workbook.add_format({'bold': 1})  
headings = ['استقلال بازدید', 'تعداد بازدید استقلال','استقلال (زمان)', 'زمان بازدید استقلال'
            ,'تیوا بازدید', 'تعداد بازدید تیوا','تیوا (زمان)', 'زمان بازدید تیوا',
            'تیوا اسپرت بازدید', 'تعداد بازدید تیوا اسپرت','تیوا اسپرت (زمان)', 'زمان بازدید تیوا اسپرت',
            'تیوا اسپرت 2 بازدید', 'تعداد بازدید تیوا اسپرت 2','تیوا اسپرت 2 (زمان)', 'زمان بازدید تیوا اسپرت 2',
            'تیوا کودک بازدید', 'تعداد بازدید تیوا کودک','تیوا کودک (زمان)', 'زمان بازدید تیوا کودک',
            'دیجیتون بازدید', 'تعداد بازدید دیجیتون','دیجیتون (زمان)', 'زمان بازدید دیجیتون',
            'لنز اسپرت بازدید', 'تعداد بازدید لنز اسپرت','لنز اسپرت (زمان)', 'زمان بازدید لنز اسپرت',
            'لنز اسپرت پلاس بازدید', 'تعداد بازدید لنز اسپرت پلاس','لنز اسپرت پلاس (زمان)', 'زمان بازدید لنز اسپرت پلاس',
            'سرباز ماهر بازدید', 'تعداد بازدید سرباز ماهر','سرباز ماهر (زمان)', 'زمان بازدید سرباز ماهر',
            'شاپرک بازدید', 'تعداد بازدید شاپرک','شاپرک (زمان)', 'زمان بازدید شاپرک',
            'تیوا آوند بازدید', 'تعداد بازدید تیوا آوند','تیوا آوند (زمان)', 'زمان بازدید تیوا آوند',
            'تیوا 2 بازدید', 'تعداد بازدید تیوا 2','تیوا 2 (زمان)', 'زمان بازدید تیوا 2',
            'تیوا فیلم بازدید', 'تعداد بازدید تیوا فیلم','تیوا فیلم (زمان)', 'زمان بازدید تیوا فیلم',
            'تیوا نوا بازدید', 'تعداد بازدید تیوا نوا','تیوا نوا (زمان)', 'زمان بازدید تیوا نوا',
            'تیوا 1 بازدید', 'تعداد بازدید تیوا 1','تیوا 1 (زمان)', 'زمان بازدید تیوا 1',
            'محفل بازدید', 'تعداد بازدید محفل','محفل (زمان)', 'زمان بازدید محفل',
            'پرسپولیس بازدید', 'تعداد بازدید پرسپولیس','پرسپولیس (زمان)', 'زمان بازدید پرسپولیس',
            'شتاب بازدید', 'تعداد بازدید شتاب','شتاب (زمان)', 'زمان بازدید شتاب',
            'کاروان عشق بازدید', 'تعداد بازدید کاروان عشق','کاروان عشق (زمان)', 'زمان بازدید کاروان عشق',
            'کنسرت خنده حسن ریوندی بازدید', 'تعداد بازدید کنسرت خنده حسن ریوندی','کنسرت خنده حسن ریوندی (زمان)', 'زمان بازدید کنسرت خنده حسن ریوندی']
  
worksheet.write_row('A1', headings, bold)  

#####################################################################
######################### write columns #############################
#####################################################################

############################# استقلال #################################
#worksheet.write_column('A2', esteghlal4[0])  
#worksheet.write_column('B2', esteghlal4[1]) 
#worksheet.write_column('C2', esteghlal4[0])  
#worksheet.write_column('D2', esteghlal4[2]) 
#
############################## تیوا #################################
#worksheet.write_column('E2', tva4[0])  
#worksheet.write_column('F2', tva4[1]) 
#worksheet.write_column('G2', tva4[0])  
#worksheet.write_column('H2', tva4[2]) 
############################## تیوا اسپرت #################################
worksheet.write_column('I2', tva_sport4[0])  
worksheet.write_column('J2', tva_sport4[1]) 
worksheet.write_column('K2', tva_sport4[0])  
worksheet.write_column('L2', tva_sport4[2]) 
############################### تیوا اسپرت 2 #################################
worksheet.write_column('M2', tva_sport_two4[0])  
worksheet.write_column('N2', tva_sport_two4[1]) 
worksheet.write_column('O2', tva_sport_two4[0])  
worksheet.write_column('P2', tva_sport_two4[2]) 
########################### تیوا کودک #################################
worksheet.write_column('Q2', tva_kodak4[0])  
worksheet.write_column('R2', tva_kodak4[1]) 
worksheet.write_column('S2', tva_kodak4[0])  
worksheet.write_column('T2', tva_kodak4[2]) 
############################## دیجیتون #################################
worksheet.write_column('U2', digiton4[0])  
worksheet.write_column('V2', digiton4[1]) 
worksheet.write_column('W2', digiton4[0])  
worksheet.write_column('X2', digiton4[2]) 
############################## لنز اسپرت #################################
worksheet.write_column('Y2', lenz_sport4[0])  
worksheet.write_column('Z2', lenz_sport4[1]) 
worksheet.write_column('AA2', lenz_sport4[0])  
worksheet.write_column('AB2', lenz_sport4[2]) 
############################## لنز اسپرت پلاس #################################
worksheet.write_column('AC2', lenz_sport_plus4[0])  
worksheet.write_column('AD2', lenz_sport_plus4[1]) 
worksheet.write_column('AE2', lenz_sport_plus4[0])  
worksheet.write_column('AF2', lenz_sport_plus4[2]) 
############################## سرباز ماهر #################################
#worksheet.write_column('AG2', sarbaz_maher4[0])  
#worksheet.write_column('AH2', sarbaz_maher4[1]) 
#worksheet.write_column('AI2', sarbaz_maher4[0])  
#worksheet.write_column('AJ2', sarbaz_maher4[2]) 
############################## شاپرک #################################
worksheet.write_column('AK2', shaparak4[0])  
worksheet.write_column('AL2', shaparak4[1]) 
worksheet.write_column('AM2', shaparak4[0])  
worksheet.write_column('AN2', shaparak4[2])
############################## شاپرک #################################
#worksheet.write_column('AO2', tva_avand4[0])  
#worksheet.write_column('AP2', tva_avand4[1]) 
#worksheet.write_column('AQ2', tva_avand4[0])  
#worksheet.write_column('AR2', tva_avand4[2])
############################### شاپرک #################################
#worksheet.write_column('AS2', tva_two4[0])  
#worksheet.write_column('AT2', tva_two4[1]) 
#worksheet.write_column('AU2', tva_two4[0])  
#worksheet.write_column('AV2', tva_two4[2])
############################### شاپرک #################################
#worksheet.write_column('AW2', tva_film4[0])  
#worksheet.write_column('AX2', tva_film4[1]) 
#worksheet.write_column('AY2', tva_film4[0])  
#worksheet.write_column('AZ2', tva_film4[2])
############################### شاپرک #################################
#worksheet.write_column('BA2', tva_nava4[0])  
#worksheet.write_column('BB2', tva_nava4[1]) 
#worksheet.write_column('BC2', tva_nava4[0])  
#worksheet.write_column('BD2', tva_nava4[2])
############################### شاپرک #################################
#worksheet.write_column('BE2', tva_one4[0])  
#worksheet.write_column('BF2', tva_one4[1]) 
#worksheet.write_column('BG2', tva_one4[0])  
#worksheet.write_column('BH2', tva_one4[2])
############################### شاپرک #################################
#worksheet.write_column('BI2', mahfel4[0])  
#worksheet.write_column('BJ2', mahfel4[1]) 
#worksheet.write_column('BK2', mahfel4[0])  
#worksheet.write_column('BL2', mahfel4[2]) 
############################## شاپرک #################################
#worksheet.write_column('BM2', perspolis4[0])  
#worksheet.write_column('BN2', perspolis4[1]) 
#worksheet.write_column('BO2', perspolis4[0])  
#worksheet.write_column('BP2', perspolis4[2])
############################## شاپرک #################################
#worksheet.write_column('BQ2', shetab4[0])  
#worksheet.write_column('BR2', shetab4[1]) 
#worksheet.write_column('BS2', shetab4[0])  
#worksheet.write_column('BT2', shetab4[2])
############################### شاپرک #################################
#worksheet.write_column('BU2', KarvanEshgh24[0])  
#worksheet.write_column('BV2', KarvanEshgh24[1]) 
#worksheet.write_column('BW2', KarvanEshgh24[0])  
#worksheet.write_column('BX2', KarvanEshgh24[2])
############################### شاپرک #################################
#worksheet.write_column('BY2', konsertReyvandi4[0])  
#worksheet.write_column('BZ2', konsertReyvandi4[1]) 
#worksheet.write_column('CA2', konsertReyvandi4[0])  
#worksheet.write_column('CB2', konsertReyvandi4[2])
workbook.close()
