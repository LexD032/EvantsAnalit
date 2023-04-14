# EvantsAnalit
конвертор EvantsAnalit Программа читает файл exported_data.xls и создает
несколько файлов: 2216_FROM_БРЯНСКАЯОБЛ.plt - прайс для ИнфоАптеки
                  price_gup.xls - прайс для Аналита
                  price_fk.dbf  - прайс для ФармКомплита
Для работы требуется установка модулей
     pywin32
     dbf
     xml
expPrice.yml среда Anсaonda.
Ошибка: win32api.dll - не найден файл решается копированием из
<Папка с установленной Anaconda>.\Lib\site-packages\pywin32_system32\
файлы pywintypes39.dll pythoncom39.dll в папку c:\windows\system32
тудаже скопировать win32api.dll 
