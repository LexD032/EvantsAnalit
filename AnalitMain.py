import datetime as dt
import Price as ia
import ReadEvants as re


path_Evants = r"d:\\Analiz\\"
path_outFile = r'd:\\Analiz\\'
File_Evants = 'exported_data.xls'

TypePrice = ('ГУП','Прочие','ОС','Аналит','Соц')
        # список с остатками склада

#***************  ГУП  ****************************#
Ost = []
Ost.clear()
FullNameEv = path_Evants+r"\\"+File_Evants
Ost = re.ReadEvantsToList(FullNameEv, TypePrice[0])
ia.PriceWrite(Ost,TypePrice[0])

#************** Прочие ****************************#
#Ost.clear()
#Ost = re.ReadEvantsToList(FullNameEv, TypePrice[2])
#ia.PriceWrite(Ost,TypePrice[2])

#************* Особый *****************************#
#Ost.clear()
#Ost = re.ReadEvantsToList(FullNameEv, TypePrice[3])
#path_outFile = ""
#FileNameIA  = ""
#FullNameIA  = path_outFile+FileNameIA
#BFpr_Price(path_outFile)

#************* Аналит ****************************#
#Ost.clear()
#Ost = re.ReadEvantsToList(FullNameEv, TypePrice[4])
#ia.PriceWrite(Ost,TypePrice[4])

#************ Социальные ************************#
#Ost.clear()
#Ost = re.ReadEvantsToList(FullNameEv, TypePrice[5])
#ia.PriceWrite(Ost,TypePrice[5])
