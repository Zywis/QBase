Attribute VB_Name = "M09_PrzykladoweWymaganiaParametrow"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M09_PrzykladoweWymaganiaParametrow
' Wstawienie przykadowych wymaga (>=20 parametrw)
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
    Dim rs As DAO.Recordset

    ' Przyklad: Nasyp @ D-02.03.01 (?b, wopt, Is, E2)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-02.03.01'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='NASYP'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='?b'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',1.8,2.2)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='wopt'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',8,12)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Is'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',0.98)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='E2'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',45)"

    ' Przyklad: AC11S @ D-05.03.05B (?b, Pb, Pb_sol)
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

    ' Przyklad: SMA11 @ D-05.03.13 (ITS, ITSR)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.13'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='SMA11'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='ITS'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',0.9)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='ITSR'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',80)"

    ' Przyklad: MROZOO @ D-04.02.02 (CBR, P063, SE)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-04.02.02'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='MROZOO'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='CBR'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',15)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='P063'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'<=Max',8)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='SE'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',35)"

    ' Przyklad: KNN-PB @ D-04.04.02 (Is, E2, LA)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-04.04.02'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='KNN-PB'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Is'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',1)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='E2'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',120)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='LA'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'<=Max',30)"

    ' Przyklad: STAB-C @ D-04.05.01 (C_cem, E2, F_ubytek)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-04.05.01'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='STAB-C'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='C_cem'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',4,6)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='E2'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',150)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='F_ubytek'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'<=Max',10)"

    ' Przyklad: AC-PB @ D-04.07.01 (Vv, VMA, VFB, Gmm)
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-04.07.01'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC-PB'")
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Vv'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',3,8)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='VMA'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc, MaxWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'Zakres',12,18)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='VFB'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, MinWartosc) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'>=Min',65)"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Gmm'")
    ExecSQL "INSERT INTO WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID, TypKryterium, Nominal, TolMinus, TolPlus) " & _
            "VALUES(" & specID & "," & warstwaID & "," & paramID & ",'+/-',2.45,0.02,0.03)"

    ' Uziarnienie  przykladowe envelope dla AC11S
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.05B'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC11S'")
    Set rs = CurrentDb.OpenRecordset("SELECT SitoID FROM Sita WHERE Rozmiar_mm In(11.2,8,4,2,0.063)", dbOpenSnapshot)
    Do While Not rs.EOF
        sitoID = rs!SitoID
        ExecSQL "INSERT INTO UziarnienieWymaganie(SpecyfikacjaID, WarstwaID, SitoID, MinProc, MaxProc) " & _
                "VALUES(" & specID & "," & warstwaID & "," & sitoID & ",5,90)"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' Uziarnienie  przykladowe envelope dla AC16W
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.05A'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC16W'")
    Set rs = CurrentDb.OpenRecordset("SELECT SitoID FROM Sita WHERE Rozmiar_mm In(16,11.2,8,4,0.063)", dbOpenSnapshot)
    Do While Not rs.EOF
        sitoID = rs!SitoID
        ExecSQL "INSERT INTO UziarnienieWymaganie(SpecyfikacjaID, WarstwaID, SitoID, MinProc, MaxProc) " & _
                "VALUES(" & specID & "," & warstwaID & "," & sitoID & ",10,85)"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    MsgBox "Wstawiono przykladowe wymagania (22).", vbInformation
    Exit Sub
ErrH:
    MsgBox "Bd Przyklad_Wymagania_20: " & Err.Description, vbCritical
End Sub
