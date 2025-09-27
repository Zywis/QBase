Attribute VB_Name = "M10_PrzykladoweWynikiBadan"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M10_PrzykladoweWynikiBadan
' Wstawienie przyk³adowych partii, badañ i wyników
'==========================================================

Private Sub ExecSQL(SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d ExecSQL (M10): " & Err.Description, vbCritical
End Sub

Public Sub Przyklad_Wyniki_20()
    On Error GoTo ErrH
    Dim obID As Long, warstwaID As Long, specID As Long, partiaID As Long
    Dim badID1 As Long, badID2 As Long, paramID As Long, sitoID As Long
    
    obID = DLookup("ObiektID", "Obiekty", "KodObiektu='OB-001'")
    warstwaID = DLookup("WarstwaID", "Warstwy", "KodWarstwy='AC11S'")
    specID = DLookup("SpecyfikacjaID", "Specyfikacje", "Kod='D-05.03.05B'")
    
    ' Partia
    ExecSQL "INSERT INTO Partie(ObiektID, WarstwaID, SpecyfikacjaID, KM_Start_m, KM_End_m, DataWbudowania, Ilosc, Jednostka, Wykonawca, Status) " & _
            "VALUES(" & obID & "," & warstwaID & "," & specID & ",0,100,Date(),200,'t','Wykonawca X','Aktywna')"
    partiaID = DLookup("PartiaID", "Partie", "ObiektID=" & obID & " AND WarstwaID=" & warstwaID)
    
    ' Badanie 1 – uziarnienie
    ExecSQL "INSERT INTO Badania(ObiektID,LokalizacjaID,WarstwaID,SpecyfikacjaID,PartiaID,RodzajBadaniaID,DataBadania,ProtokolNr,ProbkaKod,MiejscePobrania,DataWprowadzenia,Wprowadzil) " & _
            "SELECT " & obID & ",L.LokalizacjaID," & warstwaID & "," & specID & "," & partiaID & ",RB.RodzajBadaniaID,Date(),'P-001','PR-001','plac budowy',Date(),'Tester' " & _
            "FROM Lokalizacje L, RodzajeBadania RB WHERE RB.Nazwa='Uziarnienie'"
    badID1 = DLookup("BadanieID", "Badania", "ProtokolNr='P-001'")
    
    ' Badanie 2 – gêstoœæ
    ExecSQL "INSERT INTO Badania(ObiektID,LokalizacjaID,WarstwaID,SpecyfikacjaID,PartiaID,RodzajBadaniaID,DataBadania,ProtokolNr,ProbkaKod,MiejscePobrania,DataWprowadzenia,Wprowadzil) " & _
            "SELECT " & obID & ",L.LokalizacjaID," & warstwaID & "," & specID & "," & partiaID & ",RB.RodzajBadaniaID,Date(),'P-002','PR-002','plac budowy',Date(),'Tester' " & _
            "FROM Lokalizacje L, RodzajeBadania RB WHERE RB.Nazwa='Gêstoœæ objêtoœciowa'"
    badID2 = DLookup("BadanieID", "Badania", "ProtokolNr='P-002'")
    
    ' Wyniki – badanie 2
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='?b'")
    ExecSQL "INSERT INTO WynikiBadania(BadanieID,ParametrID,NrProbki,Wartosc,Jednostka) VALUES(" & badID2 & "," & paramID & ",1,2.35,'g/cm3')"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Pb'")
    ExecSQL "INSERT INTO WynikiBadania(BadanieID,ParametrID,NrProbki,Wartosc,Jednostka) VALUES(" & badID2 & "," & paramID & ",1,5.2,'%')"
    paramID = DLookup("ParametrID", "ParametryJakosci", "SymbolParametru='Pb_sol'")
    ExecSQL "INSERT INTO WynikiBadania(BadanieID,ParametrID,NrProbki,Wartosc,Jednostka) VALUES(" & badID2 & "," & paramID & ",1,5.0,'%')"
    
    ' Wyniki – uziarnienie
    For Each sitoID In CurrentDb.OpenRecordset("SELECT SitoID FROM Sita")
        ExecSQL "INSERT INTO UziarnienieWynik(BadanieID,SitoID,ProcPrzechodzenia) VALUES(" & badID1 & "," & sitoID!SitoID & "," & Rnd * 100 & ")"
    Next
    
    ' Powi¹zanie badania z oczekiwanym
    ExecSQL "UPDATE TOP 1 OczekiwaneBadanie SET BadanieID=" & badID2 & " WHERE PartiaID=" & partiaID
    
    MsgBox "Dodano przyk³adowe badania i wyniki.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d Przyklad_Wyniki_20: " & Err.Description, vbCritical
End Sub
