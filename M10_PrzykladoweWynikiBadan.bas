Attribute VB_Name = "M10_PrzykladoweWynikiBadan"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M10_PrzykladoweWynikiBadan
' Wstawienie przyk³adowych partii, badañ i wyników' Modul: M10_PrzykladoweWynikiBadan
'==========================================================

Private Sub ExecSQL(SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "Blad ExecSQL (M10): " & Err.Description, vbCritical
End Sub

Public Sub Przyklad_Wyniki_20()
    On Error GoTo ErrH
    Dim obID As Long, warstwaID As Long, specID As Long, partiaID As Long
	Dim badID1 As Long, badID2 As Long
    Dim planID As Long, oczID As Long
    Dim rodzajID As Long
    Dim rsSita As DAO.Recordset
    Dim curveValues As Variant
    Dim curveIndex As Long
    Dim valueStr As String
    Dim definicje As Variant
    Dim paramInfo As Variant
    Dim paramID As Long
    Dim paramIDVar As Variant
    Dim rodzajVar As Variant
    Dim planVar As Variant
    Dim oczVar As Variant
    Dim sitoID As Long
    Dim i As Long 
	
    obID = DLookup("ObiektID", "Obiekty", "KodObiektu='OB-001'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC11S'")
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.05B'")
    
    ' Partia
    ExecSQL "INSERT INTO Partie(ObiektID, WarstwaID, SpecyfikacjaID, KM_Start_m, KM_End_m, DataWbudowania, Ilosc, Jednostka, Wykonawca, Status) " & _
            "VALUES(" & obID & "," & warstwaID & "," & specID & ",0,100,Date(),200,'t','Wykonawca X','Aktywna')"
    partiaID = DLookup("PartiaID", "Partie", "ObiektID=" & obID & " AND WarstwaID=" & warstwaID)
   
    rodzajVar = DLookup("RodzajBadaniaID", "RodzajeBadania", "Nazwa='Gsto objtociowa'")
    If IsNull(rodzajVar) Then Err.Raise vbObjectError + 101, "Przyklad_Wyniki_20", "Brak rodzaju badania 'Gsto objtociowa'."
    rodzajID = CLng(rodzajVar)

    planVar = DLookup("PlanID", "PlanPoboru", "WarstwaID=" & warstwaID & " AND SpecyfikacjaID=" & specID & " AND RodzajBadaniaID=" & rodzajID)
    If IsNull(planVar) Then
        ExecSQL "INSERT INTO PlanPoboru(WarstwaID, SpecyfikacjaID, RodzajBadaniaID, ParametrID, Interwal, Jednostka, MinimalnaLiczba, Aktywne) " & _
                "VALUES(" & warstwaID & "," & specID & "," & rodzajID & ",Null,Null,'m',1,True)"
        planVar = DMax("PlanID", "PlanPoboru", "WarstwaID=" & warstwaID & " AND SpecyfikacjaID=" & specID & " AND RodzajBadaniaID=" & rodzajID)
    End If
    planID = CLng(planVar)

    oczVar = DLookup("ExpID", "OczekiwaneBadanie", "PartiaID=" & partiaID & " AND PlanID=" & planID)
    If IsNull(oczVar) Then
        ExecSQL "INSERT INTO OczekiwaneBadanie(PartiaID, PlanID, RodzajBadaniaID, ParametrID, Termin, Status) " & _
                "VALUES(" & partiaID & "," & planID & "," & rodzajID & ",Null,Date(),'DoWykonania')"
        oczVar = DMax("ExpID", "OczekiwaneBadanie", "PartiaID=" & partiaID & " AND PlanID=" & planID)
    End If
    oczID = CLng(oczVar)

    ' Badanie 1  uziarnienie
    ExecSQL "INSERT INTO Badania(ObiektID,LokalizacjaID,WarstwaID,SpecyfikacjaID,PartiaID,RodzajBadaniaID,DataBadania,ProtokolNr,ProbkaKod,MiejscePobrania,DataWprowadzenia,Wprowadzil) " & _
            "SELECT " & obID & ",L.LokalizacjaID," & warstwaID & "," & specID & "," & partiaID & ",RB.RodzajBadaniaID,Date(),'P-001','PR-001','plac budowy',Date(),'Tester' " & _
            "FROM Lokalizacje L, RodzajeBadania RB WHERE RB.Nazwa='Uziarnienie'"
    badID1 = DLookup("BadanieID", "Badania", "ProtokolNr='P-001'")
    
    ' Badanie 2 – gêstoœæ
    ExecSQL "INSERT INTO Badania(ObiektID,LokalizacjaID,WarstwaID,SpecyfikacjaID,PartiaID,RodzajBadaniaID,DataBadania,ProtokolNr,ProbkaKod,MiejscePobrania,DataWprowadzenia,Wprowadzil) " & _
            "SELECT " & obID & ",L.LokalizacjaID," & warstwaID & "," & specID & "," & partiaID & ",RB.RodzajBadaniaID,Date(),'P-002','PR-002','plac budowy',Date(),'Tester' " & _
            "FROM Lokalizacje L, RodzajeBadania RB WHERE RB.Nazwa='Gsto objtociowa'"
    badID2 = DLookup("BadanieID", "Badania", "ProtokolNr='P-002'")
    

    ' Wyniki dla badania 1
    definicje = Array( _
        "wopt|4.8|%", _
        "Vv|4.2|%", _
        "VMA|15.8|%", _
        "VFB|73.4|%", _
        "Gmm|2.47|g/cm3", _
        "ITS|0.84|MPa", _
        "ITSR|91|%", _
        "P063|6.3|%", _
        "E2|1300|MPa", _
        "E1|850|MPa")
    For i = LBound(definicje) To UBound(definicje)
        paramInfo = Split(definicje(i), "|")
        paramIDVar = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='" & paramInfo(0) & "'")
        If IsNull(paramIDVar) Then Err.Raise vbObjectError + 102, "Przyklad_Wyniki_20", "Brak parametru: " & paramInfo(0)
        paramID = CLng(paramIDVar)
        ExecSQL "INSERT INTO WynikiBadania(BadanieID,ParametrID,NrProbki,Wartosc,Jednostka) VALUES(" & badID1 & "," & paramID & ",1," & Replace(paramInfo(1), ",", ".") & ",'" & Replace(paramInfo(2), "'", "''") & "')"
    Next i

    ' Wyniki dla badania 2
    definicje = Array( _
        "?b|2.35|g/cm3", _
        "Pb|5.2|%", _
        "Pb_sol|5.0|%", _
        "SE|68|%", _
        "LA|18|%", _
        "F_ubytek|0.4|%", _
        "C_cem|0|%", _
        "E2E1|1.53|-", _
        "CBR|96|%", _
        "Is|99.2|%")
    For i = LBound(definicje) To UBound(definicje)
        paramInfo = Split(definicje(i), "|")
        paramIDVar = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='" & paramInfo(0) & "'")
        If IsNull(paramIDVar) Then Err.Raise vbObjectError + 103, "Przyklad_Wyniki_20", "Brak parametru: " & paramInfo(0)
        paramID = CLng(paramIDVar)
        ExecSQL "INSERT INTO WynikiBadania(BadanieID,ParametrID,NrProbki,Wartosc,Jednostka) VALUES(" & badID2 & "," & paramID & ",1," & Replace(paramInfo(1), ",", ".") & ",'" & Replace(paramInfo(2), "'", "''") & "')"
    Next i

    ' Wyniki uziarnienia - pelna krzywa
    curveValues = Array(100, 98, 90, 82, 74, 66, 52, 38, 28, 20, 12, 7, 4)
    Set rsSita = CurrentDb.OpenRecordset("SELECT SitoID FROM Sita ORDER BY Kolejnosc", dbOpenSnapshot)
    curveIndex = LBound(curveValues)
    Do While Not rsSita.EOF
        If curveIndex > UBound(curveValues) Then Err.Raise vbObjectError + 104, "Przyklad_Wyniki_20", "Brak wartosci krzywej dla dodatkowych sit."
        sitoID = rsSita!SitoID
        valueStr = Replace(CStr(curveValues(curveIndex)), ",", ".")
        ExecSQL "INSERT INTO UziarnienieWynik(BadanieID,SitoID,ProcPrzechodzenia) VALUES(" & badID1 & "," & sitoID & "," & valueStr & ")"
        curveIndex = curveIndex + 1
        rsSita.MoveNext
    Loop
    rsSita.Close
    Set rsSita = Nothing

    ' Powiazanie badania z oczekiwanym
    If oczID <> 0 Then
        ExecSQL "UPDATE OczekiwaneBadanie SET BadanieID=" & badID2 & ", Status='Zrealizowane' WHERE ExpID=" & oczID
    End If

    MsgBox "Dodano przykladowe badania i wyniki.", vbInformation
    Exit Sub
ErrH:
    If Not rsSita Is Nothing Then
        On Error Resume Next
        rsSita.Close
        Set rsSita = Nothing
        On Error GoTo 0
    End If
    MsgBox "Blad Przyklad_Wyniki_20: " & Err.Description, vbCritical
End Sub
