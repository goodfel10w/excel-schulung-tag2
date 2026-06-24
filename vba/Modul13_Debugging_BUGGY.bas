Attribute VB_Name = "Modul13_Debugging_BUGGY"
Option Explicit

' Set 13 - Debugging-Werkstatt (Aufgabe)
'
' Ein Kollege hat dieses Modul hinterlassen. ReportFuellen soll aus dem Blatt
' "Rohdaten" (Kunde, Menge, Preis) das Blatt "Report" (Kunde, Gesamt) befuellen:
' je Zeile Gesamt = Menge * Preis. Es laeuft aber nicht sauber - mal falsche
' Zahlen, mal ein Absturz. Jagen Sie die Fehler NUR mit den Debugging-Werkzeugen
' (F8, F9, Direktfenster Strg+G, Lokal-Fenster, Ueberwachung) und beheben Sie sie.
'
' WICHTIG: Code zuerst importieren (VBA-Editor: Datei -> Datei importieren).

' --- Teil A: Aufwaermen ------------------------------------------------------
' Mini-Makro zum gemeinsamen Durchsteppen. Setzen Sie einen Haltepunkt auf die
' erste Zeile in der Schleife (F9), laufen Sie mit F8 und beobachten Sie summe
' und r im Lokal-Fenster. Pruefen Sie im Direktfenster (Strg+G):
'     ?Cells(r, 1).Value
' Was steht in r, solange die Schleife laeuft?
Sub AufwaermenDurchsteppen()
    Dim r As Long
    Dim summe As Double
    summe = 0
    For r = 2 To 5
        summe = summe + Cells(r, 2).Value
    Next r
    MsgBox "Summe der ersten Mengen: " & summe
End Sub

' --- Teil B: Drei Bugs jagen -------------------------------------------------
' ReportFuellen schreibt fuer jede Rohdaten-Zeile den Kunden und das Produkt aus
' Menge und Preis in das Blatt "Report". Drei Dinge gehen schief. Dokumentieren
' Sie je Fehler: "gefunden in Zeile X mit Werkzeug Y" und beheben Sie ihn.
Sub ReportFuellen()
    Dim wsRoh As Worksheet, wsRep As Worksheet
    Dim r As Long
    Dim repZeile As Long
    Dim gesamt As Double

    Set wsRoh = Worksheets("Rohdaten")
    Set wsRep = Worksheets("Report")

    ' Report-Kopf schreiben und alten Inhalt darunter loeschen
    wsRep.Cells(1, 1).Value = "Kunde"
    wsRep.Cells(1, 2).Value = "Gesamt"
    wsRep.Range("A2:B1000").ClearContents

    repZeile = 2
    r = 1
    Do While wsRoh.Cells(r, 1).Value <> ""
        gesamt = wsRoh.Cells(r, 2).Value * wsRoh.Cells(r, 4).Value
        wsRep.Cells(repZeile, 1).Value = wsRoh.Cells(r, 1).Value
        wsRep.Cells(repZeile, 2).Value = gesamt
        repZeile = repZeile + 1
        r = r + 1
    Loop

    MsgBox "Report erstellt."
End Sub
