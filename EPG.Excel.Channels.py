# -*- coding: utf-8 -*-
"""
Created on Mon Oct 14 10:36:13 2019

@author: PC
"""

import xlsxwriter  
import pandas as pd
workbook = xlsxwriter.Workbook('Bahman.xlsx')  
df = pd.read_excel (r'C:\Users\PC\Desktop\EPG Bahman.xlsx', sheet_name='کل')
worksheet = workbook.add_worksheet() 
###########################################################
ch_one=pd.DataFrame()
ch_two=pd.DataFrame()
ch_three=pd.DataFrame()
ch_four=pd.DataFrame()
ch_five=pd.DataFrame()
khabar=pd.DataFrame()
ofogh=pd.DataFrame()
pooya=pd.DataFrame()
omid=pd.DataFrame()
ifilm=pd.DataFrame()
namayesh=pd.DataFrame()
tamasha=pd.DataFrame()
mostanad=pd.DataFrame()
shoma=pd.DataFrame()
amozesh=pd.DataFrame()
varzesh=pd.DataFrame()
nasim=pd.DataFrame()
qoran=pd.DataFrame()
salamat=pd.DataFrame()
irankala=pd.DataFrame()
alalam=pd.DataFrame()
alkosar=pd.DataFrame()
press=pd.DataFrame()
#sepehr=pd.DataFrame()

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
p21=0
p22=0
p23=0
#p24=0

df5=df.groupby(['عنوان برنامه','نام شبکه']).sum().reset_index()
t=len(df5)
for i in range(0,t):
    f=df5.loc[i,'نام شبکه']
    
#####################################################################
######################### channels data #############################
#####################################################################

############################# شبکه 1 #################################
    if f=='شبکه 1':
        p1=p1+1  
        ch_one.loc[p1,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ch_one.loc[p1,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ch_one.loc[p1,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شبکه 2 #################################
    if f=='شبکه 2': 
        p2=p2+1 
        ch_two.loc[p2,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ch_two.loc[p2,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ch_two.loc[p2,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شبکه 3 #################################
    if f=='شبکه 3':
        p3=p3+1  
        ch_three.loc[p3,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ch_three.loc[p3,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ch_three.loc[p3,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شبکه 4 #################################
    if f=='شبکه 4':
        p4=p4+1 
        ch_four.loc[p4,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ch_four.loc[p4,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ch_four.loc[p4,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شبکه 5 #################################
    if f=='شبکه 5':
        p5=p5+1  
        ch_five.loc[p5,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ch_five.loc[p5,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ch_five.loc[p5,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# خبر #################################
    if f=='خبر':  
        p6=p6+1 
        khabar.loc[p6,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        khabar.loc[p6,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        khabar.loc[p6,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# افق #################################
    if f=='افق':
        p7=p7+1  
        ofogh.loc[p7,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ofogh.loc[p7,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ofogh.loc[p7,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# پویا #################################
    if f=='پویا': 
        p8=p8+1 
        pooya.loc[p8,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        pooya.loc[p8,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        pooya.loc[p8,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# امید #################################
    if f=='امید':
        p9=p9+1  
        omid.loc[p9,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        omid.loc[p9,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        omid.loc[p9,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# آی فیلم #################################
    if f=='آی فیلم': 
        p10=p10+1 
        ifilm.loc[p10,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        ifilm.loc[p10,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        ifilm.loc[p10,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# نمایش #################################
    if f=='نمایش':
        p11=p11+1  
        namayesh.loc[p11,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        namayesh.loc[p11,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        namayesh.loc[p11,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# تماشا #################################
    if f=='تماشا': 
        p12=p12+1 
        tamasha.loc[p12,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        tamasha.loc[p12,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        tamasha.loc[p12,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# مستند #################################
    if f=='مستند':
        p13=p13+1  
        mostanad.loc[p13,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        mostanad.loc[p13,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        mostanad.loc[p13,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# شما #################################
    if f=='شما':  
        p14=p14+1 
        shoma.loc[p14,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        shoma.loc[p14,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        shoma.loc[p14,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# آموزش #################################
    if f=='آموزش':
        p15=p15+1  
        amozesh.loc[p15,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        amozesh.loc[p15,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        amozesh.loc[p15,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# ورزش #################################
    if f=='ورزش': 
        p16=p16+1 
        varzesh.loc[p16,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        varzesh.loc[p16,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        varzesh.loc[p16,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# نسیم #################################
    if f=='نسیم':
        p17=p17+1  
        nasim.loc[p17,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        nasim.loc[p17,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        nasim.loc[p17,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# قرآن #################################
    if f=='قرآن': 
        p18=p18+1 
        qoran.loc[p18,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        qoran.loc[p18,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        qoran.loc[p18,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# سلامت #################################
    if f=='سلامت':
        p19=p19+1  
        salamat.loc[p19,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        salamat.loc[p19,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        salamat.loc[p19,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# ایران کالا #################################
    if f=='ایران کالا':  
        p20=p20+1 
        irankala.loc[p20,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        irankala.loc[p20,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        irankala.loc[p20,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# العالم #################################
    if f=='العالم':
        p21=p21+1  
        alalam.loc[p21,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        alalam.loc[p21,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        alalam.loc[p21,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# الکوثر #################################
    if f=='الکوثر':  
        p22=p22+1 
        alkosar.loc[p22,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        alkosar.loc[p22,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        alkosar.loc[p22,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# پرس تی وی #################################
    if f=='پرس تی وی':
        p23=p23+1  
        press.loc[p23,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
        press.loc[p23,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
        press.loc[p23,'مدت بازدید']=df5.loc[i,'مدت بازدید']
############################# سپهر #################################
#    if f=='سپهر':  
#        p24=p24+1 
#        sepehr.loc[p24,'عنوان برنامه']=df5.loc[i,'عنوان برنامه']
#        sepehr.loc[p24,'تعداد بازدید']=df5.loc[i,'تعداد بازدید']
#        sepehr.loc[p24,'مدت بازدید']=df5.loc[i,'مدت بازدید']
        
#####################################################################
######################### append data ###############################
#####################################################################

############################# شبکه 1 #################################
ch_one1=[]
ch_one2=[]
ch_one3=[]
ch_one4=[]
ch_one5=[]
ch_one.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_one1=ch_one["عنوان برنامه"].tolist()
ch_one5.append(ch_one1)
ch_one2=ch_one["تعداد بازدید"].tolist()
ch_one5.append(ch_one2)
ch_one.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_one3=ch_one["عنوان برنامه"].tolist()
ch_one5.append(ch_one3)
ch_one4=ch_one["مدت بازدید"].tolist()
ch_one5.append(ch_one4)
############################# شبکه 2 #################################
ch_two1=[]
ch_two2=[]
ch_two3=[]
ch_two4=[]
ch_two5=[]
ch_two.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_two1=ch_two["عنوان برنامه"].tolist()
ch_two5.append(ch_two1)
ch_two2=ch_two["تعداد بازدید"].tolist()
ch_two5.append(ch_two2)
ch_two.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_two3=ch_two["عنوان برنامه"].tolist()
ch_two5.append(ch_two3)
ch_two4=ch_two["مدت بازدید"].tolist()
ch_two5.append(ch_two4)
############################# شبکه 3 #################################
ch_three1=[]
ch_three2=[]
ch_three3=[]
ch_three4=[]
ch_three5=[]
ch_three.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_three1=ch_three["عنوان برنامه"].tolist()
ch_three5.append(ch_three1)
ch_three2=ch_three["تعداد بازدید"].tolist()
ch_three5.append(ch_three2)
ch_three.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_three3=ch_three["عنوان برنامه"].tolist()
ch_three5.append(ch_three3)
ch_three4=ch_three["مدت بازدید"].tolist()
ch_three5.append(ch_three4)
############################# شبکه 4 #################################
ch_four1=[]
ch_four2=[]
ch_four3=[]
ch_four4=[]
ch_four5=[]
ch_four.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_four1=ch_four["عنوان برنامه"].tolist()
ch_four5.append(ch_four1)
ch_four2=ch_four["تعداد بازدید"].tolist()
ch_four5.append(ch_four2)
ch_four.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_four3=ch_four["عنوان برنامه"].tolist()
ch_four5.append(ch_four3)
ch_four4=ch_four["مدت بازدید"].tolist()
ch_four5.append(ch_four4)
############################# شبکه 5 #################################
ch_five1=[]
ch_five2=[]
ch_five3=[]
ch_five4=[]
ch_five5=[]
ch_five.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_five1=ch_five["عنوان برنامه"].tolist()
ch_five5.append(ch_five1)
ch_five2=ch_five["تعداد بازدید"].tolist()
ch_five5.append(ch_five2)
ch_five.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ch_five3=ch_five["عنوان برنامه"].tolist()
ch_five5.append(ch_five3)
ch_five4=ch_five["مدت بازدید"].tolist()
ch_five5.append(ch_five4)
############################# خبر #################################
khabar1=[]
khabar2=[]
khabar3=[]
khabar4=[]
khabar5=[]
khabar.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
khabar1=khabar["عنوان برنامه"].tolist()
khabar5.append(khabar1)
khabar2=khabar["تعداد بازدید"].tolist()
khabar5.append(khabar2)
khabar.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
khabar3=khabar["عنوان برنامه"].tolist()
khabar5.append(khabar3)
khabar4=khabar["مدت بازدید"].tolist()
khabar5.append(khabar4)
############################# افق #################################
ofogh1=[]
ofogh2=[]
ofogh3=[]
ofogh4=[]
ofogh5=[]
ofogh.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ofogh1=ofogh["عنوان برنامه"].tolist()
ofogh5.append(ofogh1)
ofogh2=ofogh["تعداد بازدید"].tolist()
ofogh5.append(ofogh2)
ofogh.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ofogh3=ofogh["عنوان برنامه"].tolist()
ofogh5.append(ofogh3)
ofogh4=ofogh["مدت بازدید"].tolist()
ofogh5.append(ofogh4)
############################# پویا #################################
pooya1=[]
pooya2=[]
pooya3=[]
pooya4=[]
pooya5=[]
pooya.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
pooya1=pooya["عنوان برنامه"].tolist()
pooya5.append(pooya1)
pooya2=pooya["تعداد بازدید"].tolist()
pooya5.append(pooya2)
pooya.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
pooya3=pooya["عنوان برنامه"].tolist()
pooya5.append(pooya3)
pooya4=pooya["مدت بازدید"].tolist()
pooya5.append(pooya4)
############################# امید #################################
omid1=[]
omid2=[]
omid3=[]
omid4=[]
omid5=[]
omid.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
omid1=omid["عنوان برنامه"].tolist()
omid5.append(omid1)
omid2=omid["تعداد بازدید"].tolist()
omid5.append(omid2)
omid.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
omid3=omid["عنوان برنامه"].tolist()
omid5.append(omid3)
omid4=omid["مدت بازدید"].tolist()
omid5.append(omid4)
############################# آی فیلم #################################
ifilm1=[]
ifilm2=[]
ifilm3=[]
ifilm4=[]
ifilm5=[]
ifilm.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ifilm1=ifilm["عنوان برنامه"].tolist()
ifilm5.append(ifilm1)
ifilm2=ifilm["تعداد بازدید"].tolist()
ifilm5.append(ifilm2)
ifilm.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ifilm3=ifilm["عنوان برنامه"].tolist()
ifilm5.append(ifilm3)
ifilm4=ifilm["مدت بازدید"].tolist()
ifilm5.append(ifilm4)
############################# نمایش #################################
namayesh1=[]
namayesh2=[]
namayesh3=[]
namayesh4=[]
namayesh5=[]
namayesh.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
namayesh1=namayesh["عنوان برنامه"].tolist()
namayesh5.append(namayesh1)
namayesh2=namayesh["تعداد بازدید"].tolist()
namayesh5.append(namayesh2)
namayesh.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
namayesh3=namayesh["عنوان برنامه"].tolist()
namayesh5.append(namayesh3)
namayesh4=namayesh["مدت بازدید"].tolist()
namayesh5.append(namayesh4)
############################# تماشا #################################
tamasha1=[]
tamasha2=[]
tamasha3=[]
tamasha4=[]
tamasha5=[]
tamasha.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
tamasha1=tamasha["عنوان برنامه"].tolist()
tamasha5.append(tamasha1)
tamasha2=tamasha["تعداد بازدید"].tolist()
tamasha5.append(tamasha2)
tamasha.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
tamasha3=tamasha["عنوان برنامه"].tolist()
tamasha5.append(tamasha3)
tamasha4=tamasha["مدت بازدید"].tolist()
tamasha5.append(tamasha4)
############################# مستند #################################
mostanad1=[]
mostanad2=[]
mostanad3=[]
mostanad4=[]
mostanad5=[]
mostanad.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
mostanad1=mostanad["عنوان برنامه"].tolist()
mostanad5.append(mostanad1)
mostanad2=mostanad["تعداد بازدید"].tolist()
mostanad5.append(mostanad2)
mostanad.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
mostanad3=mostanad["عنوان برنامه"].tolist()
mostanad5.append(mostanad3)
mostanad4=mostanad["مدت بازدید"].tolist()
mostanad5.append(mostanad4)
############################# شما #################################
shoma1=[]
shoma2=[]
shoma3=[]
shoma4=[]
shoma5=[]
shoma.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shoma1=shoma["عنوان برنامه"].tolist()
shoma5.append(shoma1)
shoma2=shoma["تعداد بازدید"].tolist()
shoma5.append(shoma2)
shoma.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shoma3=shoma["عنوان برنامه"].tolist()
shoma5.append(shoma3)
shoma4=shoma["مدت بازدید"].tolist()
shoma5.append(shoma4)
############################# آموزش #################################
amozesh1=[]
amozesh2=[]
amozesh3=[]
amozesh4=[]
amozesh5=[]
amozesh.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
amozesh1=amozesh["عنوان برنامه"].tolist()
amozesh5.append(amozesh1)
amozesh2=amozesh["تعداد بازدید"].tolist()
amozesh5.append(amozesh2)
amozesh.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
amozesh3=amozesh["عنوان برنامه"].tolist()
amozesh5.append(amozesh3)
amozesh4=amozesh["مدت بازدید"].tolist()
amozesh5.append(amozesh4)
############################# ورزش #################################
varzesh1=[]
varzesh2=[]
varzesh3=[]
varzesh4=[]
varzesh5=[]
varzesh.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
varzesh1=varzesh["عنوان برنامه"].tolist()
varzesh5.append(varzesh1)
varzesh2=varzesh["تعداد بازدید"].tolist()
varzesh5.append(varzesh2)
varzesh.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
varzesh3=varzesh["عنوان برنامه"].tolist()
varzesh5.append(varzesh3)
varzesh4=varzesh["مدت بازدید"].tolist()
varzesh5.append(varzesh4)
############################# نسیم #################################
nasim1=[]
nasim2=[]
nasim3=[]
nasim4=[]
nasim5=[]
nasim.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
nasim1=nasim["عنوان برنامه"].tolist()
nasim5.append(nasim1)
nasim2=nasim["تعداد بازدید"].tolist()
nasim5.append(nasim2)
nasim.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
nasim3=nasim["عنوان برنامه"].tolist()
nasim5.append(nasim3)
nasim4=nasim["مدت بازدید"].tolist()
nasim5.append(nasim4)
############################# قرآن #################################
qoran1=[]
qoran2=[]
qoran3=[]
qoran4=[]
qoran5=[]
qoran.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
qoran1=qoran["عنوان برنامه"].tolist()
qoran5.append(qoran1)
qoran2=qoran["تعداد بازدید"].tolist()
qoran5.append(qoran2)
qoran.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
qoran3=qoran["عنوان برنامه"].tolist()
qoran5.append(qoran3)
qoran4=qoran["مدت بازدید"].tolist()
qoran5.append(qoran4)
############################# سلامت #################################
salamat1=[]
salamat2=[]
salamat3=[]
salamat4=[]
salamat5=[]
salamat.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
salamat1=salamat["عنوان برنامه"].tolist()
salamat5.append(salamat1)
salamat2=salamat["تعداد بازدید"].tolist()
salamat5.append(salamat2)
salamat.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
salamat3=salamat["عنوان برنامه"].tolist()
salamat5.append(salamat3)
salamat4=salamat["مدت بازدید"].tolist()
salamat5.append(salamat4)
############################# ایران کالا #################################
irankala1=[]
irankala2=[]
irankala3=[]
irankala4=[]
irankala5=[]
irankala.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
irankala1=irankala["عنوان برنامه"].tolist()
irankala5.append(irankala1)
irankala2=irankala["تعداد بازدید"].tolist()
irankala5.append(irankala2)
irankala.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
irankala3=irankala["عنوان برنامه"].tolist()
irankala5.append(irankala3)
irankala4=irankala["مدت بازدید"].tolist()
irankala5.append(irankala4)
############################# العالم #################################
alalam1=[]
alalam2=[]
alalam3=[]
alalam4=[]
alalam5=[]
alalam.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
alalam1=alalam["عنوان برنامه"].tolist()
alalam5.append(alalam1)
alalam2=alalam["تعداد بازدید"].tolist()
alalam5.append(alalam2)
alalam.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
alalam3=alalam["عنوان برنامه"].tolist()
alalam5.append(alalam3)
alalam4=alalam["مدت بازدید"].tolist()
alalam5.append(alalam4)
############################# الکوثر #################################
alkosar1=[]
alkosar2=[]
alkosar3=[]
alkosar4=[]
alkosar5=[]
alkosar.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
alkosar1=alkosar["عنوان برنامه"].tolist()
alkosar5.append(alkosar1)
alkosar2=alkosar["تعداد بازدید"].tolist()
alkosar5.append(alkosar2)
alkosar.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
alkosar3=alkosar["عنوان برنامه"].tolist()
alkosar5.append(alkosar3)
alkosar4=alkosar["مدت بازدید"].tolist()
alkosar5.append(alkosar4)
############################# پرس تی وی #################################
#press1=[]
#press2=[]
#press3=[]
#press4=[]
#press5=[]
#press.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
#press1=press["عنوان برنامه"].tolist()
#press5.append(press1)
#press2=press["تعداد بازدید"].tolist()
#press5.append(press2)
#press.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
#press3=press["عنوان برنامه"].tolist()
#press5.append(press3)
#press4=press["مدت بازدید"].tolist()
#press5.append(press4)
############################# سپهر #################################
#sepehr1=[]
#sepehr2=[]
#sepehr3=[]
#sepehr4=[]
#sepehr5=[]
#sepehr.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
#sepehr1=sepehr["عنوان برنامه"].tolist()
#sepehr5.append(sepehr1)
#sepehr2=sepehr["تعداد بازدید"].tolist()
#sepehr5.append(sepehr2)
#sepehr.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
#sepehr3=sepehr["عنوان برنامه"].tolist()
#sepehr5.append(sepehr3)
#sepehr4=sepehr["مدت بازدید"].tolist()
#sepehr5.append(sepehr4)

##################################################################
#ofogh4=ofogh4.sort_values(["تعداد بازدید افق" , "شبکه افق بازدید"], ascending=[True,False])
bold = workbook.add_format({'bold': 1})  
headings = ['شبکه 1 بازدید', 'تعداد بازدید شبکه 1','شبکه 1 (زمان)', 'زمان بازدید شبکه 1'
            ,'شبکه 2 بازدید', 'تعداد بازدید شبکه 2','شبکه 2 (زمان)', 'زمان بازدید شبکه 2',
            'شبکه 3 بازدید', 'تعداد بازدید شبکه 3','شبکه 3 (زمان)', 'زمان بازدید شبکه 3',
            'شبکه 4 بازدید', 'تعداد بازدید شبکه 4','شبکه 4 (زمان)', 'زمان بازدید شبکه 4',
            'شبکه 5 بازدید', 'تعداد بازدید شبکه 5','شبکه 5 (زمان)', 'زمان بازدید شبکه 5',
            'شبکه خبر بازدید', 'تعداد بازدید شبکه خبر','شبکه خبر (زمان)', 'زمان بازدید شبکه خبر',
            'شبکه افق بازدید', 'تعداد بازدید شبکه افق','شبکه افق (زمان)', 'زمان بازدید شبکه افق',
            'شبکه پویا بازدید', 'تعداد بازدید شبکه پویا','شبکه پویا (زمان)', 'زمان بازدید شبکه پویا',
            'شبکه امید بازدید', 'تعداد بازدید شبکه امید','شبکه امید (زمان)', 'زمان بازدید شبکه امید',
            'شبکه آی فیلم بازدید', 'تعداد بازدید شبکه آی فیلم','شبکه آی فیلم (زمان)', 'زمان بازدید شبکه آی فیلم',
            'شبکه نمایش بازدید', 'تعداد بازدید شبکه نمایش','شبکه نمایش (زمان)', 'زمان بازدید شبکه نمایش',
            'شبکه تماشا بازدید', 'تعداد بازدید شبکه تماشا','شبکه تماشا (زمان)', 'زمان بازدید شبکه تماشا',
            'شبکه مستند بازدید', 'تعداد بازدید شبکه مستند','شبکه مستند (زمان)', 'زمان بازدید شبکه مستند',
            'شبکه شما بازدید', 'تعداد بازدید شبکه شما','شبکه شما (زمان)', 'زمان بازدید شبکه شما',
            'شبکه آموزش بازدید', 'تعداد بازدید شبکه آموزش','شبکه آموزش (زمان)', 'زمان بازدید شبکه آموزش',
            'شبکه ورزش بازدید', 'تعداد بازدید شبکه ورزش','شبکه ورزش (زمان)', 'زمان بازدید شبکه ورزش',
            'شبکه نسیم بازدید', 'تعداد بازدید شبکه نسیم','شبکه نسیم (زمان)', 'زمان بازدید شبکه نسیم',
            'شبکه قرآن بازدید', 'تعداد بازدید شبکه قرآن','شبکه قرآن (زمان)', 'زمان بازدید شبکه قرآن',
            'شبکه سلامت بازدید', 'تعداد بازدید شبکه سلامت','شبکه سلامت (زمان)', 'زمان بازدید شبکه سلامت',
            'شبکه ایران کالا بازدید', 'تعداد بازدید شبکه ایران کالا','شبکه ایران کالا (زمان)', 'زمان بازدید شبکه ایران کالا',
            'شبکه العالم بازدید', 'تعداد بازدید شبکه العالم','شبکه العالم (زمان)', 'زمان بازدید شبکه العالم',
            'شبکه الکوثر بازدید', 'تعداد بازدید شبکه الکوثر','شبکه الکوثر (زمان)', 'زمان بازدید شبکه الکوثر',
            'شبکه پرس تی وی بازدید', 'تعداد بازدید شبکه پرس تی وی','شبکه پرس تی وی (زمان)', 'زمان بازدید شبکه پرس تی وی']       
worksheet.write_row('A1', headings, bold)  

#####################################################################
######################### write columns #############################
#####################################################################

############################# شبکه 1 #################################
worksheet.write_column('A2', ch_one5[0])  
worksheet.write_column('B2', ch_one5[1]) 
worksheet.write_column('C2', ch_one5[2])  
worksheet.write_column('D2', ch_one5[3]) 
############################# شبکه 2 #################################
worksheet.write_column('E2', ch_two5[0])  
worksheet.write_column('F2', ch_two5[1]) 
worksheet.write_column('G2', ch_two5[2])  
worksheet.write_column('H2', ch_two5[3]) 
############################# شبکه 3 #################################
worksheet.write_column('I2', ch_three5[0])  
worksheet.write_column('J2', ch_three5[1]) 
worksheet.write_column('K2', ch_three5[2])  
worksheet.write_column('L2', ch_three5[3]) 
############################# شبکه 4 #################################
worksheet.write_column('M2', ch_four5[0])  
worksheet.write_column('N2', ch_four5[1]) 
worksheet.write_column('O2', ch_four5[2])  
worksheet.write_column('P2', ch_four5[3]) 
############################# شبکه 5 #################################
worksheet.write_column('Q2', ch_five5[0])  
worksheet.write_column('R2', ch_five5[1]) 
worksheet.write_column('S2', ch_five5[2])  
worksheet.write_column('T2', ch_five5[3]) 
############################# خبر #################################
worksheet.write_column('U2', khabar5[0])  
worksheet.write_column('V2', khabar5[1]) 
worksheet.write_column('W2', khabar5[2])  
worksheet.write_column('X2', khabar5[3]) 
############################# افق #################################
worksheet.write_column('Y2', ofogh5[0])  
worksheet.write_column('Z2', ofogh5[1]) 
worksheet.write_column('AA2', ofogh5[2])  
worksheet.write_column('AB2', ofogh5[3]) 
############################# پویا #################################
worksheet.write_column('AC2', pooya5[0])  
worksheet.write_column('AD2', pooya5[1]) 
worksheet.write_column('AE2', pooya5[2])  
worksheet.write_column('AF2', pooya5[3]) 
############################# امید #################################
worksheet.write_column('AG2', omid5[0])  
worksheet.write_column('AH2', omid5[1]) 
worksheet.write_column('AI2', omid5[2])  
worksheet.write_column('AJ2', omid5[3]) 
############################# آی فیلم #################################
worksheet.write_column('AK2', ifilm5[0])  
worksheet.write_column('AL2', ifilm5[1]) 
worksheet.write_column('AM2', ifilm5[2])  
worksheet.write_column('AN2', ifilm5[3]) 
############################# نمایش #################################
worksheet.write_column('AO2', namayesh5[0])  
worksheet.write_column('AP2', namayesh5[1]) 
worksheet.write_column('AQ2', namayesh5[2])  
worksheet.write_column('AR2', namayesh5[3]) 
############################# تماشا #################################
worksheet.write_column('AS2', tamasha5[0])  
worksheet.write_column('AT2', tamasha5[1]) 
worksheet.write_column('AU2', tamasha5[2])  
worksheet.write_column('AV2', tamasha5[3])
############################# مستند #################################
worksheet.write_column('AW2', mostanad5[0])  
worksheet.write_column('AX2', mostanad5[1]) 
worksheet.write_column('AY2', mostanad5[2])  
worksheet.write_column('AZ2', mostanad5[3]) 
############################# شما #################################
worksheet.write_column('BA2', shoma5[0])  
worksheet.write_column('BB2', shoma5[1]) 
worksheet.write_column('BC2', shoma5[2])  
worksheet.write_column('BD2', shoma5[3]) 
############################# آموزش #################################
worksheet.write_column('BE2', amozesh5[0])  
worksheet.write_column('BF2', amozesh5[1]) 
worksheet.write_column('BG2', amozesh5[2])  
worksheet.write_column('BH2', amozesh5[3]) 
############################# ورزش #################################
worksheet.write_column('BI2', varzesh5[0])  
worksheet.write_column('BJ2', varzesh5[1]) 
worksheet.write_column('BK2', varzesh5[2])  
worksheet.write_column('BL2', varzesh5[3]) 
############################# نسیم #################################
worksheet.write_column('BM2', nasim5[0])  
worksheet.write_column('BN2', nasim5[1]) 
worksheet.write_column('BO2', nasim5[2])  
worksheet.write_column('BP2', nasim5[3]) 
############################# قرآن #################################
worksheet.write_column('BQ2', qoran5[0])  
worksheet.write_column('BR2', qoran5[1]) 
worksheet.write_column('BS2', qoran5[2])  
worksheet.write_column('BT2', qoran5[3])
############################# سلامت #################################
worksheet.write_column('BU2', salamat5[0])  
worksheet.write_column('BV2', salamat5[1]) 
worksheet.write_column('BW2', salamat5[2])  
worksheet.write_column('BX2', salamat5[3]) 
############################# ایران کالا #################################
worksheet.write_column('BY2', irankala5[0])  
worksheet.write_column('BZ2', irankala5[1]) 
worksheet.write_column('CA2', irankala5[2])  
worksheet.write_column('CB2', irankala5[3]) 
############################# العالم #################################
worksheet.write_column('CC2', alalam5[0])  
worksheet.write_column('CD2', alalam5[1]) 
worksheet.write_column('CE2', alalam5[2])  
worksheet.write_column('CF2', alalam5[3]) 
############################# الکوثر #################################
worksheet.write_column('CG2', alkosar5[0])  
worksheet.write_column('CH2', alkosar5[1]) 
worksheet.write_column('CI2', alkosar5[2])  
worksheet.write_column('CJ2', alkosar5[3]) 
############################# پرس تی وی #################################
#worksheet.write_column('CK2', press5[0])  
#worksheet.write_column('CL2', press5[1]) 
#worksheet.write_column('CM2', press5[2])  
#worksheet.write_column('CN2', press5[3]) 
############################# سپهر #################################
#worksheet.write_column('CO2', sepehr5[0])  
#worksheet.write_column('CP2', sepehr5[1]) 
#worksheet.write_column('CQ2', sepehr5[0])  
#worksheet.write_column('CR2', sepehr5[2])


workbook.close()
