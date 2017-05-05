import win32com.client as win32
import Wort as WBL
import Dicts as DC

#Open Excel Worterbuch
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open('Worterbuch.xlsx')
#ws = wb.Worksheets("Deutsch")

#Define Variables
Worter = {} # GlobalDictionary
Dicts = DC.listws(wb) # Each Worksheet is a Dictionary
Ubersetzung = [] # List of Translation

dic = DC.Menu(wb, Dicts)
# dic = DC.Wahlen(choose, Dicts, wb)
excel.Application.Quit()

while True:
    Wahl = WBL.menu()
    wb = excel.Workbooks.Open('Worterbuch.xlsx')
    excel.Visible = False
    if Wahl == 1:
        ws = wb.Worksheets(dic)
        length = WBL.lenWB(ws)
        Worter = WBL.readWB(ws, Worter, length)
        print(Worter)
    elif Wahl == 2:
        ws = wb.Worksheets(dic)
        length = WBL.lenWB(ws)
        a = raw_input("Italienisch Wort: \n")
        i = WBL.iswort(ws, a, length)
        if i == -1:
            b = raw_input("Ubersetzung (teil mit Komma auf, wenn du mehr Worter hast): \n")
            Ubersetzung = ([x.strip() for x in b.split(',')])  # list of Word
            WBL.hinzu(Worter, a, Ubersetzung)
            WBL.hinzuWB(ws, a, Worter, length)
        else:
            ant = raw_input("Willst du neue Ubersetzungen hinzufugen? Y/N")
            if ant == "Y":
                b = raw_input("Ubersetzung (teil mit Komma auf, wenn du mehr Worter hast): \n")
                Ubersetzung = []
                j = 2
                while ws.Cells(i, j).Value is not None:
                    Ubersetzung.append(ws.Cells(i, j).Value)
                    j += 1
                Ubersetzung.append([x.strip() for x in b.split(',')])  # list of Word
                WBL.hinzu(Worter, a, Ubersetzung)
                WBL.hinzuWB(ws, a, Worter, length)
        wb.Save()
    elif Wahl == 3:
        ws = wb.Worksheets(dic)
        Wortweg = raw_input("Was willst du wegraumen?")
        length = WBL.lenWB(ws)
        WBL.weg(ws, Wortweg, length)
        wb.Save()
    elif Wahl == 4:
        ws = wb.Worksheets(dic)
        Wortsuc = raw_input("Was willst du suchen?")
        lenght = WBL.lenWB(ws)
        WBL.suche(ws, Wortsuc, lenght)
        wb.Save()
    elif Wahl == 5:
        # wb = excel.Workbooks.Open('Worterbuch.xlsx')
        W = DC.Menu(wb, Dicts)
    elif Wahl == 6:
        break
    excel.Application.Quit()
