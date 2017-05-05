def menu():
    Wahl = input("Was willst du machen? \n\
          1. Auflisten \n\
          2. Hinzufugen \n\
          3. Wegraumen\n\
          4. Suche \n\
          5. Change Dictionary \n\
          6. Exit")
    return (Wahl)

def hinzu(Worterbuch, NeuesWort, Ubersetzung):
    Worterbuch[NeuesWort] = Ubersetzung

def lenWB(ws):
    lungh=0
    while True:
        if ws.Cells(lungh + 1, 1).Value is not None:
            lungh += 1
        else:
            break
    return (lungh+1)

def readWB(ws, Worter, lunghezza):
    #lunghezza = lenWB(ws)
    for i in range(1, lunghezza+1):
        Ubersetzung = []
        j = 2  # Index to read Rows = how many translation there is
        while ws.Cells(i, j).Value is not None:
            Ubersetzung.append(ws.Cells(i, j).Value)
            j += 1
        Worter[ws.Cells(i, 1).Value] = Ubersetzung
    return(Worter)

def hinzuWB(ws, NeuesWort, Worter, l):
    # l = lenWB(ws)
    ws.Cells(l + 1, 1).Value = NeuesWort
    ws.Cells(l + 1, 1).Font.Bold = True
    usnum = 2
    for us in Worter[NeuesWort]:
        ws.Cells(l + 1, usnum).Value = us
        usnum += 1

def iswort (ws, Wort, lenght):
    for rowI in range(1, lenght):
        if ws.Cells(rowI, 1).Value == Wort:
            index = rowI
            break
        elif ws.Cells(rowI, 1).Value is None:
            index = -1
            break
        elif ws.Cells(rowI, 1).Value is not Wort:
            index = -1
            continue
    return (index)

def weg(ws, WortWeg, lenght):
    i = iswort(ws, WortWeg, lenght)
    if i == -1:
        print("No %s found" % WortWeg)
    else:
        ws.Cells(i, 1).EntireRow.Delete()

def suche (ws, WortSuc, lenght):
    i = iswort(ws, WortSuc, lenght)
    if i == -1:
        print("No %s found" % WortSuc)
    else:
        Ubersetzung = []
        j = 2
        while ws.Cells(i, j).Value is not None:
            Ubersetzung.append(ws.Cells(i, j).Value)
            j += 1
        print("%s : %s" % (WortSuc, Ubersetzung))