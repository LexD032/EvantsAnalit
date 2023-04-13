# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 15:19:20 2022

@author: Администратор
"""

import xml.dom.minidom as xm 
import datetime as dt
import shutil
import win32com.client
import dbf

#**************  Инфоаптека сформировать и прайс ******************************

def PriceWrite(ost,Pr):
#    outFile = r"\\\\SERV7\\MailBox\\netapt\\OUT\\2216_FROM_БРЯНСКАЯОБЛ.plt"
    outFile = r"d:\\Analiz\\2216_FROM_БРЯНСКАЯОБЛ.plt"

    if Pr == 'ГУП':
        FileoutBF=r'D:\\Analiz\\price_gup.xls'
#        NameFilFK=r'D:\\Analiz\\price_fk.dbf'
        tableFK = dbf.Table(filename = r'D:\\Analiz\\price_fk.dbf',
                          field_specs = 'GOODSCODE C(20); \
                          GOODSNAME N(19,0); \
                          PRODNAME C(255);\
                          UNIT C(15);\
                          QUANTITY N(15,0);\
                          COST N(8,2);\
                          PERIOD D;\
                          JVLS N(1,0)',
#                          SHK C(13)',
                          codepage='cp866')
        tableFK.open(mode=dbf.READ_WRITE)    
        

    if Pr == 'Прочие':
        FileoutBF=r'D:\\Analiz\\price_pr.xls'
        
    if Pr == 'Соц':
        FileoutBF=r'D:\\Analiz\\price_soc.xls'
        
    shutil.copyfile(r"D:\\Analiz\\pr_shbl.xls", FileoutBF)
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = True        # otherwise excel is hidden
    wb = xl.Workbooks.Open(FileoutBF)  
    sht = wb.Sheets("sheet1")
        
#	    NameFilFK=r'D:\\Analiz\\price_fk.dbf'
        
        
#  Копируем шаблон dbf
        
    if Pr == 'Аналит':
#	    NameFilAnalit=r'\\mark10\\608609\\Analit\\price.dbf"
        NameFilAnalit = r'D:\\Analiz\\price.dbf'
        table = dbf.Table(filename = NameFilAnalit,
                          field_specs = 'GOODSCODE C(20);\
                          GOODSNAME N(19,0); \
                          PRODNAME C(255);\
                          UNIT C(15);\
                          QUANTITY N(15,0);\
                          COST N(8,2);\
                          PERIOD D;\
                          JVLS N(1,0)',
 #                         SHK C(13)',
                          codepage='cp886')
        table.open(mode=dbf.READ_WRITE)    

    DOM_XML = xm.Document()
    ITEM__List = XML_OPEN(DOM_XML, Pr)

    dicItem = {}
#        тут копируем файл шаблона excel  
#        и открываем для вывода
    i=4
    for elem in ost:

        dicItem.clear()
        
        sht.Cells(i,9).value = elem.get('NDS')         #/* ставка ндс  товра */

        if elem.get('QTTY') > 0:
            
            sht.Cells(i,6).value = elem.get('QTTY')        #/* количество    */
            
            if (Pr=='ГУП' or Pr=='Прочие' or Pr=='Соц'):
#             тут записываем в файл шаблона excel 
                sht.Cells(i,1).value = elem.get('CODE')    
                sht.Cells(i,2).value = elem.get('NAME')    
                sht.Cells(i,3).value = elem.get('VENDOR')  
                sht.Cells(i,4).value = 'упак' 
                sht.Cells(i,8).value = elem.get('VALID_DATE')
               
                if (Pr=='ГУП'):
                    sht.Cells(i,5).value = elem.get('CenaOptBNDS')
                    
                if (Pr=='Прочие'):
                    sht.Cells(i,5).value = elem.get('CenaOptSNDS') #/* цена      */

                CodeVNR = elem.get('CODE_VENDOR')
                if (not CodeVNR):
                    CodeNomkl = elem.get('CODE')
                else:
                    CodeNomkl = elem.get('CODE')+CodeVNR
                if (Pr=='ГУП' or Pr=='Соц'):
                    dicItem = {'CODE': CodeNomkl}
                else:
                    dicItem = {'CODE': elem.get('CODE')}

                dicItem.update({'NAME': elem.get('NAME')})
                dicItem.update({'VENDOR': elem.get('VENDOR')})
                dicItem.update({'COUNTRY': elem.get('COUNTRY')})
                dicItem.update({'VENDORBARCODE': elem.get('VENDORBARCODE')})
                dicItem.update({'VALID_DATE': elem.get('VALID_DATE')})
                dicItem.update({'CenaOptSNDS': elem.get('CenaOptBNDS')})
                getStringToXML(DOM_XML, ITEM__List, dicItem)

            if (Pr=='Аналит'):
                
                if not ((elem.get('SECTOR') == 100) or 
                        (elem.get('SECTOR') == 25)):

                    ba = ()    
                    ba = (elem.get('CODE'),
                    elem.get('NAME'),
                    elem.get('VENDOR'),
                    elem.get('COUNTRY'),
                    elem.get('VENDORBARCODE'),
                    elem.get('QTTY'),
                    elem.get('VALID_DATE'),
                    elem.get('CenaOptSNDS'))
                    
                    table.append(ba)

            if (Pr=='ГУП'):
                b = ()
                b = (elem.get('CODE'),
                     elem.get('NAME'),
                     elem.get('VENDOR'),
                     elem.get('COUNTRY'),
                     elem.get('VENDORBARCODE'),
                     elem.get('QTTY'),
                     elem.get('VALID_DATE'),
                     elem.get('CenaOptSNDS'))
                
                tableFK.append(b)


    i=i+1
# Конец цикла по записям
        
    if (Pr =='ГУП'):
#		ИмяФайлаВывода="\\Mark10\608609\Analit\price_гуп.xls";  
#		ДБФ.ЗакрытьФайл();
#	    print("Файл сформирован Прайс_Для_ФармКомплита")			
        wb.Close()
        xl.Quit()
        print('Сформирован Прайс_ГУП')
#		ИмяФИА="\\Serv7\MailBox\netall\OUT\2216_FROM_БрянскаяОбл.plt"
        NameFilIA="d:\Analiz\2216_FROM_БрянскаяОбл.plt"
        ITEM__List.Save(NameFilIA)
        print("Сформирован Прайс_ГУП Для Инфоаптека")  
        tableFK.close()
        print("Файл сформирован Прайс для ФармКомплита")

    elif (Pr =='Прочие'):
        wb.Close()
        xl.Quit()
        print('Файл сформирован Прайс_ПРОЧИЕ')
        
    elif (Pr =='ОС'):
#		ИмяФайлаВывода="d:\Analiz\price_ос.xls";
#//		ИмяФайлаВывода="\\mark10\608609\Analit\price_ос.xls";
	    print("Файл сформирован Прайс_Особые")
        
    elif (Pr =='Аналит'):               
        table.close()
        print("Файл сформирован Прайс_Для_Аналита");	

    elif (Pr =='Соц'):
#		ИмяФайлаВывода="D:\Analiz\price_соц.xls";  
#//		ИмяФайлаВывода="\\mark10\608609\Analit\price_соц.xls";  
   	    print("Сформирован Прайс_ГУП Соц")	
#        NmaeFilSoc=
#               "\\Serv7\MailBox\netall\OUT\22161_FROM_БрянскаяОбл.plt"
#        ITEM__List.Save(NameFilSoc)
#	    print("Сформирован и скопирован Соц Прайс_ГУП Для Инфоаптека")
    return
              

              
#****************  Заголовок xml файла ****************************************
def XML_OPEN(ObjectXML,Pr):
    
    Header = ObjectXML.createProcessingInstruction("xml","version=""1.0"""\
                                                 "encoding=""windows-1251"" " )   
    ObjectXML.appendChild(Header)
    TageMassage = ObjectXML.createElement("PACKET")
    TageMassage.setAttribute("TYPE", "10") 
    TageMassage.setAttribute("NAME", "Прайс-лист")
    TageMassage.setAttribute("FROM", "ГУПБрянскфармация")
    ObjectXML.appendChild(TageMassage) 
	
    PRICELIST = ObjectXML.createElement("PRICELIST")
    dtt=dt.datetime.now()
    dtst=dtt.strftime( '%Y-%m-%d %H:%M:%S')
    PRICELIST.setAttribute("DATE", ""+dtst+"")
    if Pr=="Соц":
        PRICELIST.setAttribute("NAME", "СоцГУПБрянскфармация")
    else:
        PRICELIST.setAttribute("NAME", "ГУПБрянскфармация")
        TageMassage.appendChild(PRICELIST)
        
    return(PRICELIST)

#****************  Записи по позициям xml файла ****************************************
def getStringToXML(objXML,Item__List,dcList):
        ITEM = objXML.createElement("ITEM")
        Item__List.appendChild(ITEM)
        
        CODE = objXML.createElement("CODE")
        ITEM.appendChild(CODE)
        CODE_=objXML.createTextNode(dcList.get("CODE"))
        CODE.appendChild(CODE_)

        NAME = objXML.createElement("NAME")
        ITEM.appendChild(NAME)
        NAME_=objXML.createTextNode(dcList.get("NAME"))
        NAME.appendChild(NAME_) 
		
        VENDOR = objXML.createElement("VENDOR")
        ITEM.appendChild(VENDOR)
        VENDOR_=objXML.createTextNode(dcList.get("VENDOR"))
        VENDOR.appendChild(VENDOR_)

        COUNTRY = objXML.createElement("COUNTRY")
        ITEM.appendChild(COUNTRY)
        COUNTRY_=objXML.createTextNode(dcList.get("COUNTRY"))
        COUNTRY.appendChild(COUNTRY_)

        VENDORBARCODE = objXML.createElement("VENDORBARCODE")
        ITEM.appendChild(VENDORBARCODE)
        VENDORBARCODE_=objXML.createTextNode(dcList.get("VENDORBARCODE"))
        VENDORBARCODE.appendChild(VENDORBARCODE_)

        QTTY = objXML.createElement("QTTY")
        ITEM.appendChild(QTTY);
        QTTY_=objXML.createTextNode(str(dcList.get("QTTY")))
        QTTY.appendChild(QTTY_)

        PRICES = objXML.createElement("PRICES")
        ITEM.appendChild(PRICES)
		
        VALID_DATE = objXML.createElement("VALID_DATE")
        ITEM.appendChild(VALID_DATE)
        strdate =  dcList.get("VALID_DATE").strftime("%d %H:%M")
        VALID_DATE_=objXML.createTextNode(strdate)
        VALID_DATE.appendChild(VALID_DATE_)

        Cena = objXML.createElement("Отстрочка_0")
        PRICES.appendChild(Cena)
        if (dcList.get("Otsrochka_0")):
            Cena_ = objXML.createTextNode(str(dcList.get("Otsrochka_0")))
        else: 
            Cena_ = objXML.createTextNode('0.00')
        Cena.appendChild(Cena_)
    
        return    

