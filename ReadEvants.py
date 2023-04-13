#from win32com.client import Dispatch
import win32com.client 


def readZagKol(Exl):
   shl = []
   tm=''
   j = 1
   tm=Exl.Cells(1, j).Value
   while  (tm) :
 #      str_tmp = Exl.Cells(1, j).Address
 #      str_arr = str_tmp.split("$")
       shl.append(Exl.Cells(1, j).Value)
 #      shl.append(str_arr[1])
       j=j+1
       tm=Exl.Cells(1, j).Value
   return(shl)


def getZagAdrr(shl, NameKol):
   d = 0
   try:
      d = shl.index(NameKol)+1
   except:
      d = 0
   return(d)


def getCenaSoc(CnOpt):
    if (CnOpt >= 500):
       CenaProd = CnOpt*1.05
    elif (CnOpt >= 100) and (CnOpt < 500):
       CenaProd = CnOpt*1.1
    elif (CnOpt < 100):
       CenaProd = CnOpt*1.15
    return(CenaProd)


def getCena(CnOpt):
    if (CnOpt > 500):
       CenaProd = CnOpt*1.15
    else:
       CenaProd = CnOpt*1.1
    return(CenaProd)



def ReadEvantsToList(FllNm, Pr):
   xl = win32com.client.Dispatch("Excel.Application")
   xl.Visible = True        # otherwise excel is hidden

# newest excel does not accept forward slash in path
   wb = xl.Workbooks.Open(FllNm)  # непонял про r- только для чтения?
   sht = wb.Sheets("EXPORTED_DATA")

   zagkol = []
   zagkol = readZagKol(sht)
   ost_in = []
   ost_str = {}
   i = 2
   while (sht.Cells(i, 1).Value):
       ost_str.clear()
#  Определяем сектор хранения
       sector = 0
       adrr = 0
       adrr = getZagAdrr(zagkol, "сектор_хранения")
       sektor_str = sht.Cells(i,adrr).value
       tmp = []
       tmp = sektor_str.split("=")
       sector_str = tmp[0].replace('М','')
       sector = int(sector_str.strip())
       
       ost_str.update({'SECTOR': sector})

       KodTovara = sht.Range("D"+str(i)).value
       ost_str.update({'CODE': KodTovara})
       NameTovara = sht.Range("E"+str(i)).value
       ost_str.update({'NAME': NameTovara})

       adrr = 0
       adrr = getZagAdrr(zagkol, "штрих_код")
       shk = sht.Cells(i,adrr).value
       ost_str.update({'VENDORBARCODE': shk})

       ZhV = sht.Range("AD"+str(i)).value
       ost_str.update({'ZhV':ZhV})
       
       adrr = 0
       adrr = getZagAdrr(zagkol, "срок_годности")
       try:
          GodenDo = sht.Cells(i,adrr).value
       except:
          GodenDo = ""
       ost_str.update({'VALID_DATE': GodenDo})

       nds_str = sht.Range("Q"+str(i)).value
       tmp = []
       tmp = nds_str.split("=")
       NDS = int(tmp[0])
       ost_str.update({'NDS':NDS})

       adrr = 0
       adrr = getZagAdrr(zagkol, "изготовитель")
       Izgot = sht.Cells(i,adrr).value
       tmp = []
       tmp = Izgot.split("==")

       ost_str.update({'VENDOR': tmp[1]})
       ost_str.update({'COD_VENDOR': tmp[0]})
       

#       adrr = 0
#       adrr = getZagAdrr(zagkol, "цена_учетная_остатка")
#       CenaPost = float(sht.Range(str(adrr)+str(i)).value)

#       adrr = 0
#       adrr = getZagAdrr(zagkol, "розничная_цена_остатка")
#       CenaRee = float(sht.Range(str(adrr)+str(i)).value)

#       adrr = 0
#       adrr = getZagAdrr(zagkol, "цена_изготовителя")
#       CenaIzg = float(sht.Range(str(adrr)+str(i)).value)
#      if (CenaIzg == 0):
#           adrr = 0
#           adrr = getZagAdrr(zagkol, "розничная_цена_остатка")
#           CenaIzg = float(sht.Range(str(adrr)+str(i)).value)
       if (Pr == 'Прочие'):      # Прочие :-))
           
          if ((sector == 100) or (sector == 25) or (sector == 12)):
             adrr = 0
             adrr = getZagAdrr(zagkol, "оптовая_цена_остатка")
             CenaOptBNDS = float(sht.Cells(i,adrr).value)
          elif ((sector == 1) or (sector == 9) or (sector == 13)):
             adrr = 0
             adrr = getZagAdrr(zagkol, "оптовая_цена_остатка")
             CenaOptBNDS = float(sht.Cells(i,adrr).value)
             CenaOptBNDS = CenaOptBNDS * 1.1
          else  : 
             adrr = 0
             adrr = getZagAdrr(zagkol, "цена_учетная_остатка")
             CenaOptBNDS = float(sht.Cells(i,adrr).value)
             
       else :                # Аналиту, ГУП, особые
          
          if (Pr == 'Соц'):
              
             adrr = 0
             adrr = getZagAdrr(zagkol, "цена_учетная_остатка")
             CenaOptBNDS = float(sht.Cells(i,adrr).value)
             CenaOptBNDS = getCenaSoc(CenaOptBNDS) 
             
          else:
              
             if ((sector == 100) or (sector == 25) or (sector == 13)):
                adrr = 0
                adrr = getZagAdrr(zagkol, "оптовая_цена_остатка")
                CenaOptBNDS = float(sht.Cells(i,adrr).value)
             else :
                adrr = 0
                adrr = getZagAdrr(zagkol, "цена_учетная_остатка")
                CenaOptBNDS = float(sht.Cells(i,adrr).value)
                
       ost_str.update({'CenaOptBNDS': CenaOptBNDS})
       CenaOptSNDS = CenaOptBNDS * 1.1
       ost_str.update({'CenaOptSNDS':CenaOptSNDS})
       KolVo = sht.Range("B"+str(i)).value         
       ost_str.update({'QTTY':KolVo})
       adrr = 0
       adrr = getZagAdrr(zagkol, "страна_происхождения_1")
       Country = sht.Cells(i,adrr).value
       ost_str.update({'COUNTRY':Country})
       ost_in.append(ost_str)

       i=i+1
   
   wb.Close()
   xl.Quit()

   return(ost_in)
