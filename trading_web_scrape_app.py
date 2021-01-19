from tkinter import *
import requests
from bs4 import BeautifulSoup as soup
from xlwt import Workbook
import datetime
import pandas as pd
import tkinter
from tkinter import messagebox
import xlwt
import time
from win32com.client import Dispatch
import sys,os


pd.set_option('max_columns',None)
running=True
minutes=1
dirname, filename = os.path.split(os.path.abspath(__file__))
print("running from", dirname)
data_dic={}

def interface():
    global entry,entry2,entry3,window,entryq,entry2q,entry3q,\
        entryw,entry2w,entry3w,entrye,entry2e,entry3e,entryr,entry2r,entry3r,window,button,entry_refresh,button_stop
    print(repr(sys.argv[0]))
    print(repr(os.getcwd()))

    window=tkinter.Tk()

    window.title('Stock Calculator')
    window.config(background='SlateGray2')
    window.geometry("800x560")

    label = Label(window, text='URL:', bg='SlateGray2')
    label2 = Label(window, text='Price range from:', bg='SlateGray2')
    label3 = Label(window, text='Price range to:', bg='SlateGray2')

    labelq = Label(window, text='URL:',  bg='SlateGray2')
    label2q = Label(window, text='Price range from:',  bg='SlateGray2')
    label3q = Label(window, text='Price range to:', bg='SlateGray2')

    labelw = Label(window, text='URL:',  bg='SlateGray2')
    label2w = Label(window, text='Price range from:', bg='SlateGray2')
    label3w = Label(window, text='Price range to:',  bg='SlateGray2')

    labele = Label(window, text='URL:',  bg='SlateGray2')
    label2e = Label(window, text='Price range from:',  bg='SlateGray2')
    label3e = Label(window, text='Price range to:',  bg='SlateGray2')

    labelr = Label(window, text='URL:', bg='SlateGray2')
    label2r = Label(window, text='Price range from:',  bg='SlateGray2')
    label3r = Label(window, text='Price range to:', bg='SlateGray2')

    label_refresh=Label(window, text='Refresh rate(minutes):',  bg='SlateGray3')
    entry_refresh = Entry(window, bd=4, bg='azure3')

    empty_label=Label(window,text='',bg='LightSteelBlue1')
    empty_label1 = Label(window, text='', bg='LightSteelBlue1')
    empty_label2 = Label(window, text='', bg='LightSteelBlue1')
    empty_label3 = Label(window, text='', bg='LightSteelBlue1')
    empty_label4 = Label(window, text='', bg='LightSteelBlue1')
    empty_label5 = Label(window, text='', bg='LightSteelBlue1')

    button=Button(window,text='SUBSCRIBE',command=Interface_entry,bd=5,bg='light blue')

    button_stop = Button(window, text='STOP!', command=stop, bd=5, bg='light blue')

    entry = Entry(window, bd=3, bg='lavender')
    entry2 = Entry(window, bd=3, bg='lavender')
    entry3 = Entry(window, bd=3, bg='lavender')

    entryq = Entry(window, bd=3, bg='lavender')
    entry2q = Entry(window, bd=3, bg='lavender')
    entry3q = Entry(window, bd=3, bg='lavender')

    entryw = Entry(window, bd=3, bg='lavender')
    entry2w = Entry(window, bd=3, bg='lavender')
    entry3w = Entry(window, bd=3, bg='lavender')

    entrye = Entry(window, bd=3, bg='lavender')
    entry2e = Entry(window, bd=3, bg='lavender')
    entry3e = Entry(window, bd=3, bg='lavender')

    entryr = Entry(window, bd=3, bg='lavender')
    entry2r = Entry(window, bd=3, bg='lavender')
    entry3r = Entry(window, bd=3, bg='lavender')

    label.grid(row=1,sticky='w')
    entry.grid(row=1,column=1,sticky='ew')
    label2.grid(row=2, sticky='w')
    entry2.grid(row=2, column=1,sticky='w')
    label3.grid(row=3, sticky='w')
    entry3.grid(row=3, column=1,sticky='w')

    empty_label.grid(row=4, column=0, columnspan=3, sticky='news')

    labelw.grid(row=5, sticky='w')
    entryw.grid(row=5, column=1, sticky='ew')
    label2w.grid(row=6, sticky='w')
    entry2w.grid(row=6, column=1,sticky='w')
    label3w.grid(row=7, sticky='w')
    entry3w.grid(row=7, column=1,sticky='w')

    empty_label1.grid(row=8, column=0, columnspan=3, sticky='news')

    labele.grid(row=9, sticky='w')
    entrye.grid(row=9, column=1, sticky='ew')
    label2e.grid(row=10, sticky='w')
    entry2e.grid(row=10, column=1,sticky='w')
    label3e.grid(row=11, sticky='w')
    entry3e.grid(row=11, column=1,sticky='w')

    empty_label2.grid(row=12, column=0, columnspan=3, sticky='news')

    labelr.grid(row=13, sticky='w')
    entryr.grid(row=13, column=1, sticky='ew')
    label2r.grid(row=14, sticky='w')
    entry2r.grid(row=14, column=1,sticky='w')
    label3r.grid(row=15, sticky='w')
    entry3r.grid(row=15, column=1,sticky='w')

    empty_label3.grid(row=16, column=0, columnspan=3, sticky='news')

    labelq.grid(row=17, sticky='w')
    entryq.grid(row=17, column=1, sticky='ew')
    label2q.grid(row=18, sticky='w')
    entry2q.grid(row=18, column=1,sticky='w')
    label3q.grid(row=19, sticky='w')
    entry3q.grid(row=19, column=1,sticky='w')

    empty_label4.grid(row=20, column=0, columnspan=3, sticky='news')
    empty_label5.grid(row=21, column=0, columnspan=3, sticky='news')

    label_refresh.grid(row=22,column=0,sticky='e')
    entry_refresh.grid(row=22,column=1,sticky='w')

    button.grid(row=23,column=0,sticky='ew')
    button_stop.grid(row=24, column=0, sticky='ew')

    window.grid_columnconfigure(1, weight=1)
    window.grid_rowconfigure(0,weight=1)

    window.after(1000,scanning)
    window.mainloop()

def scrape(my_url):
    global df,rounded_value_of_current_price,wb,now
    internet=False
    strike_price_list=[]
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    while internet==False:
        try:
            r = requests.get(my_url)
            internet=True
            break
        except:
            print('trying')
            time.sleep(5)

    page_soup = soup(r.text,"html.parser")
    tablica=page_soup.find('div',{'id':'wrapper_btm'})
    price_table=tablica.find('table')
    spans=price_table.find_all('span')

    splited=(spans[0].text).split(' ')
    now=datetime.datetime.now()
    symbol=splited[2]

    sheet1.write(0,0,str(now.strftime('%d-%m-%Y')),xlwt.easyxf("pattern: pattern solid, fore_color yellow; font: color black; align: horiz center"))
    sheet1.write(1,0,'last update: '+str(now.strftime("%H:%M:%S")),xlwt.easyxf("pattern: pattern solid, fore_color yellow; font: color black; align: horiz center"))
    sheet1.write(2,0,splited[2],xlwt.easyxf("pattern: pattern solid, fore_color yellow; font: color black; align: horiz center"))
    sheet1.write(3,0,splited[3],xlwt.easyxf("pattern: pattern solid, fore_color yellow; font: color black; align: horiz center"))
    sheet1.write(0,3,spans[0].text,xlwt.easyxf("pattern: pattern solid, fore_color bright_green; font: color black; align: horiz center"))
    sheet1.write(1,3,spans[1].text,xlwt.easyxf("pattern: pattern solid, fore_color bright_green; font: color black; align: horiz center"))
    sheet1.write(7,5,'CALLS',xlwt.easyxf("pattern: pattern solid, fore_color blue; font: color black; align: horiz center"))
    sheet1.write(7, 17, 'PUTS',xlwt.easyxf("pattern: pattern solid, fore_color blue; font: color black; align: horiz center"))
    first_col=sheet1.col(0)
    forth_col=sheet1.col(3)
    eleven_col=sheet1.col(11)
    first_col.width = 256 * 20
    forth_col.width = 9000
    eleven_col.width=4000

    table=page_soup.find('table',{'id':'octable'})
    header1=table.find('thead')
    header11=header1.find_all('tr')
    tbody=table.find_all('tr')

    final = header11[1]
    names = final.find_all('th')
    i = 0
    for x in names:
        sheet1.write(8, i, x.text, xlwt.easyxf("pattern: pattern solid, fore_color light_turquoise; font: color black; align: horiz center"))
        i += 1

    count=0
    for x in range(2,len(tbody)):
        first_row=tbody[x].find_all('td')
        count+=1
        ii=0
        for c in first_row:
            if ii==11:
                sheet1.write(8 + count, ii, c.text,xlwt.easyxf("pattern: pattern solid, fore_color light_turquoise; font: color black; align: horiz center"))
                strike_price_list.append(c.text)
            else:
                sheet1.write(8+count,ii,c.text)
            ii+=1
    for i in range(0, len(strike_price_list)):
        strike_price_list[i] = int(float(strike_price_list[i]))
    strike_price_list.append(int(float(splited[3])))
    strike_price_list.sort()
    nomer=strike_price_list.index((int(float(splited[3]))))
    rounded_value_of_current_price=strike_price_list[nomer+1]
    wb.save(str(symbol)+'__example.xls')
    xls_file = pd.ExcelFile(str(symbol)+'__example.xls')
    df = xls_file.parse('Sheet 1', skiprows=9, header=None,\
                        names=['chart','oi','chng in oi','volume','IV','LTP','Net Chng','Bid','Bid','Ask','Ask',	\
                               'Strike Price','Bid','Bid','Ask','Ask','Net Chng','LTP','IV','Volume','Chng in OI','OI','Chart'])
    return symbol


def calculationsMAX(range_from,range_to):
    global max_in_PE,max_in_CE,strike_price_CE,strike_price_PE,price_range_list,index_of_current_price,LTP_in_PE,LTP_in_CE

    OTM_PE=None
    OTM_CE=None
    LTP_in_CE=[]
    LTP_in_PE=[]
    strike_price_PE=[]
    strike_price_CE=[]
    max_in_PE=[]
    max_in_CE=[]
    price_range_list=[]

    for x in range(len(df)):
        if df['Strike Price'][x]>=range_from:
            print(df['Strike Price'][x])
            price_range_list.append(df['Strike Price'][x])
        if df['Strike Price'][x]==range_to:
            break

    index_of_current_price=price_range_list.index(rounded_value_of_current_price)
    print(index_of_current_price)
    print(rounded_value_of_current_price)

    for x in range(len(df)):
        if df['Strike Price'][x]==rounded_value_of_current_price:
            OTM_PE=df.loc[0:x-1]
            OTM_CE=df.loc[x::]


    OTM_PE_mirror=OTM_PE
    OTM_CE_mirror=OTM_CE
    OTM_CE_mirror.drop(OTM_CE_mirror.tail(1).index, inplace=True)
    OTM_PE_mirror = OTM_PE_mirror.replace(to_replace ="-", value ='0')
    OTM_CE_mirror = OTM_CE_mirror.replace(to_replace="-", value='0')

    for x in range(len(OTM_PE_mirror)):
        if int(OTM_PE_mirror['Strike Price'][x])>=range_from:
            print(int(OTM_PE_mirror['Strike Price'][x]))
            OTM_PE_mirror=OTM_PE_mirror.loc[x::]
            break

    OTM_PE_mirror=OTM_PE_mirror.reset_index()
    OTM_CE_mirror=OTM_CE_mirror.reset_index()

    for x in range(len(OTM_CE_mirror)):
        if int(OTM_CE_mirror['Strike Price'][x])>=range_to:
            print(int(OTM_CE_mirror['Strike Price'][x]))
            OTM_CE_mirror=OTM_CE_mirror.loc[0:x]
            break

    print(OTM_CE_mirror)
    print(OTM_PE_mirror)

    for x in range(len(OTM_PE_mirror)):
        string = OTM_PE_mirror['OI'][x]
        raw=string.split(',')
        string=''.join(raw)
        final=int(string)
        a = OTM_PE_mirror.index[x]
        OTM_PE_mirror.at[a, 'OI'] = final

    for x in range(len(OTM_CE_mirror)):
        string = OTM_CE_mirror['oi'][x]
        raw=string.split(',')
        string=''.join(raw)
        final=int(string)
        a = OTM_CE_mirror.index[x]
        OTM_CE_mirror.at[a, 'oi'] = final
    print(OTM_PE_mirror)
    try:
        for x in range(6):
            maxima=OTM_PE_mirror['OI'].max()
            max_in_PE.append(maxima)
            maxi = OTM_PE_mirror.index[OTM_PE_mirror['OI'] == maxima].tolist()
            strike_price=OTM_PE_mirror['Strike Price'][maxi]
            LTP=OTM_PE_mirror['LTP'][maxi]
            LTP = str(LTP)
            res = LTP.split('\\n')
            result = res[-2]
            print(result)
            if '-' in result:
                end = '-'
                print(end)
            else:
                final = result.split(',')
                end = float(''.join(final))

            LTP_in_PE.append(end)
            strike_price_PE.append(float(strike_price))
            OTM_PE_mirror = OTM_PE_mirror.drop(maxi)
    except:
        pass


    try:
        for x in range(6):
            maxima=OTM_CE_mirror['oi'].max()
            max_in_CE.append(maxima)
            maxi = OTM_CE_mirror.index[OTM_CE_mirror['oi'] == maxima].tolist()
            strike_price = OTM_CE_mirror['Strike Price'][maxi]
            LTP=OTM_CE_mirror['LTP.1'][maxi]
            LTP = str(LTP)
            res = LTP.split('\\n')
            result = res[-2]
            print(result)
            if '-' in result:
                end1='-'
                print(end1)
            else:
                final = result.split(',')
                end1 = float(''.join(final))

            LTP_in_CE.append(end1)
            strike_price_CE.append(float(strike_price))
            OTM_CE_mirror = OTM_CE_mirror.drop(maxi)
    except:
        pass

    #print(max_in_PE)
    print(LTP_in_PE)
    #print(max_in_CE)
    print(LTP_in_CE)
    #print(strike_price_CE)
    #print(price_range_list)


def Interface_entry():
    global data_dic,running,fon,fon_q,fon_w,fon_e,fon_r,refresh_rate,button,hours,minutes
    running=False
    data_dic={}

    link=entry.get()
    range_from=entry2.get()
    range_to  =entry3.get()
    data_dic[link]=[range_from,range_to]

    link_q = entryq.get()
    range_from_q = entry2q.get()
    range_to_q = entry3q.get()
    data_dic[link_q] = [range_from_q, range_to_q]

    link_w = entryw.get()
    range_from_w = entry2w.get()
    range_to_w = entry3w.get()
    data_dic[link_w] = [range_from_w, range_to_w]

    link_e = entrye.get()
    range_from_e = entry2e.get()
    range_to_e = entry3e.get()
    data_dic[link_e] = [range_from_e, range_to_e]

    link_r = entryr.get()
    range_from_r = entry2r.get()
    range_to_r = entry3r.get()
    data_dic[link_r] = [range_from_r, range_to_r]
    print(data_dic)

    if link=='':
        messagebox.showerror('error',message='Please enter URL!')

    refresh_rate=entry_refresh.get()
    if refresh_rate=='':
        messagebox.showerror('error', message='Please enter Refresh rate!')

    if refresh_rate!='':
        #try:
            refresh_rate=int(refresh_rate)
            minutes=refresh_rate % 60
            hours=refresh_rate // 6
            for x,y in data_dic.items():
                if x!= '' and y[0]!='' and y[1]!='':
                    symbol=scrape(x)
                    calculationsMAX(int(y[0]), int(y[1]))
                    sheet2(int(y[0]), int(y[1]), symbol,first=True)
                    button.config(bg='lime green', text='LIVE')
                    button_stop.config(bg='firebrick3')
                    running=True


    print(running)

def stop():
    global running,button
    if running==True:
        button_stop.config(bg='light blue')
    running=False
    button.config(bg='light blue',text='SUBSCRIBE')


def scanning():
    global running
    if running:
        currenttime = datetime.datetime.now()
        print(currenttime.minute % minutes)
        if (currenttime.second == 0 or currenttime.second ==1) and (currenttime.minute % minutes)==0 :
            for x, y in data_dic.items():
                if x != '' and y[0] != '' and y[1] != '':
                    symbol = scrape(x)
                    calculationsMAX(int(y[0]), int(y[1]))
                    sheet2(int(y[0]), int(y[1]), symbol, first=False)
    window.after(1000,scanning)

def sheet2(range_from,range_to,symbol,first):
    global string
    string = str(dirname + '\Results__' + str(symbol) + '.xls')

    wb = Workbook()

    sheet2 = wb.add_sheet('Sheet 2')
    sheet2.write_merge(0, 0,0,2,  'Market range:  '+str(range_from)+'-'+str(range_to),
                       xlwt.easyxf("pattern: pattern solid, fore_color 43; font: color black; align: horiz center"))
    sheet2.write_merge(0, 0, 3, 14, "Top CEs",
                       xlwt.easyxf("pattern: pattern solid, fore_color 7; font: color black; align: horiz center"))
    sheet2.write_merge(0, 0, 15, 26, "Top PEs",
                       xlwt.easyxf("pattern: pattern solid, fore_color 46; font: color black; align: horiz center"))
    sheet2.write_merge(1, 1,0,2, 'Current Price',xlwt.easyxf("pattern: pattern solid, fore_color 51; font: color black; align: horiz center"))



    try:
        i=0
        for x in range(1,13):
            if x%2==0:
                sheet2.write(1,14+x,'difference',xlwt.easyxf("pattern: pattern solid, fore_color 22; font: color black; align: horiz center"))
            else:
                sheet2.write(1,14+x,strike_price_PE[i],xlwt.easyxf("pattern: pattern solid, fore_color 51; font: color black; align: horiz center"))
                i+=1
    except:
        pass


    try:
        i=0
        for x in range(1, 13):
            if x % 2 == 0:
                sheet2.write(1, 2 + x, 'difference',xlwt.easyxf("pattern: pattern solid, fore_color 22; font: color black; align: horiz center"))
            else:
                sheet2.write(1, 2 + x, strike_price_CE[i],xlwt.easyxf("pattern: pattern solid, fore_color 51; font: color black; align: horiz center"))
                i += 1
    except:
        pass

    i=0
    for x in range(1,13):
        if x %2 ==0:
            pass

        else:
            sheet2.write(2 + index_of_current_price, 2 + x, LTP_in_CE[i],
                         xlwt.easyxf("pattern: pattern solid, fore_color 51; font: color black; align: horiz center"))
            i+=1

    i = 0
    for x in range(1, 13):
        if x % 2 == 0:
            pass

        else:
            sheet2.write(2 + index_of_current_price, 14 + x, LTP_in_PE[i],
                         xlwt.easyxf("pattern: pattern solid, fore_color 51; font: color black; align: horiz center"))
            i += 1


    try:
        i=0
        for x in range(len(price_range_list)):
            sheet2.write(2+i,1,price_range_list[x],xlwt.easyxf("pattern: pattern solid, fore_color 67; font: color black; align: horiz center"))
            i+=1
    except:
        pass

    i = 0
    for u in range(2, 26):
        if u % 2 == 0:
                pass
        else:
            try:
                for x in range(len(price_range_list)):
                    sheet2.write(2 + x, 2 + i, '', xlwt.easyxf("pattern: pattern solid, fore_color 67"))
            except:
                pass
        i += 1

    if first:
        wb.save(string)

    xl = Dispatch("Excel.Application")

    wz = xl.Workbooks.Open(dirname+'\Results__' + str(symbol)+'.xls')

    wz.Close(True,string)

    time.sleep(0.05)
    wb.save(string)

    xl.Workbooks.Open(dirname+'\Results__' + str(symbol)+'.xls')

    xl.Visible = True

def excelOpen(symbol):
    xl = Dispatch("Excel.Application")
    wz = xl.Workbooks.Open(dirname+'\Results__' + str(symbol) + '.xls')
    wz.Visible=True

    return wz
def excelClose(wz,symbol):
    wz.Close(True, dirname+'\Results__' + str(symbol) + '.xls')

interface()

