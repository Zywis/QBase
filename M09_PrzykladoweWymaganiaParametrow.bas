Attribute VB_Name = "M09_PrzykladoweWymaganiaParametrow"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M09_PrzykladoweWymaganiaParametrow
' Wstawienie przyk³adowych wymagañ (20 parametrów)
'==========================================================

Private Sub ExecSQL(SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d ExecSQL (M09): " & Err.Description, vbCritical
End Sub

Public Sub Przyklad_Wymagania_20()
    On Error GoTo ErrH
    Dim specID As Long, warstwaID As Long, paramID As Long, sitoID As Long
    
    ' Przyk³ad: Nasyp @ D-02.03.01 (?b, wopt)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-02.03.01'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='NASYP'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='?b'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',1.8,2.2)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='wopt'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',8,12)"
    
    ' Przyk³ad: AC11S @ D-05.03.05B (?b, Pb, Pb_sol)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.05B'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC11S'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='?b'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',2.3,2.6)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Pb'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',4.5,6)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Pb_sol'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',4,6)"
    
    ' Przyk³ad: SMA11 @ D-05.03.13 (ITS, ITSR)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.13'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='SMA11'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='ITS'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',0.9)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='ITSR'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',80)"
    
    ' Uziarnienie – przyk³adowe envelope dla AC11S
    For Each sitoID In CurrentDb.OpenRecordset("SELECT SitoID FROM Sita WHERE Rozmiar_mm In(11.2,8,4,2,0.063)")
        ExecSQL "INSERT INTO UziarnienieWymaganie(SpecyfikacjaID, WarstwaID, SitoID, MinProc, MaxProc) " & _
                "VALUES(" & specID & "," & warstwaID & "," & sitoID!SitoID & ",5,90)"
    Next
    
    MsgBox "Wstawiono przyk³adowe wymagania (20).", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d Przyklad_Wymagania_20: " & Err.Description, vbCritical
End Sub
