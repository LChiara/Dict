def Menu(wb, Dicts):
    print("Was willst du machen? \n\
        0. Neues Worterbuch \n\
         . Existiert Worterbuch\n")
    i = 1
    for d in Dicts:
        print("%d. %s \n" %(i, Dicts[i]))
        i += 1
    W = input("--> ")
    if W == 0:
        neue = raw_input("Welche neue Sprache wurdest du hinzufugen? ")
        if isws(Dicts, neue) == True:
            print("Schon existiert die Sprache in dem Worterbuch! ")
        else:
            addws(wb, Dicts, neue)
            print("Du hast ein neue Worterbuch erzeugt! Gut!")
            dic = neue
            wb.Save()
    else:
        dic = Dicts[W]
    return(dic)

def isws(dict, neue):
    for ws in dict:
        if neue == dict[ws]:
            return True
        else:
            return False


def addws(wb, dict, neue):
    ws = wb.Worksheets.Add()
    ws.Name = neue
    dict[len(dict)] = neue


def listws (wb):
    dict = {}
    count = wb.Sheets.Count
    for i in range(1, count + 1):
        ws = wb.Sheets(i)
        dict[i] = ws.Name
    return(dict)

