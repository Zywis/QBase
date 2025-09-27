Attribute VB_Name = "M08_PartieNCR"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M08_PartieNCR
' Partie, Oczekiwane badania, NCR, FE/BE
'==========================================================

Private Sub ExecSQL(ByVal SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d ExecSQL: " & Err.Description, vbCritical
End Sub

Private Function NowyForm(ByVal Nazwa As String, ByVal RS As String, ByVal Continuous As Boolean) As Form
    On Error GoTo ErrH
    Dim f As Form
    Set f = CreateForm
    f.RecordSource = RS
    f.DefaultView = IIf(Continuous, 1, 0)
    f.Caption = Nazwa
    DoCmd.Save acForm, f.Name
    DoCmd.Rename Nazwa, acForm, f.Name
    Set NowyForm = Forms(Nazwa)
    Exit Function
ErrH:
    MsgBox "B³¹d NowyForm: " & Err.Description, vbCritical
    Set NowyForm = Nothing
End Function

Public Sub UtworzFormularze_PartieNCR()
    On Error GoTo ErrH
    Dim fPartia As Form, fPlan As Form, fExp As Form, fNCR As Form
    Dim sfExp As Form, sfBadania As Form
    
    ' Podformularze
    Set sfExp = NowyForm("SF_Oczekiwane", "OczekiwaneBadanie", True)
    Set sfBadania = NowyForm("SF_BadaniaLotu", "Badania", True)
    
    ' F_Partia
    Set fPartia = NowyForm("F_Partia", "Partie", False)
    CreateControl "F_Partia", acSubform, acDetail, , "SF_Oczekiwane", 3000, 300, 5000, 3000
    CreateControl "F_Partia", acSubform, acDetail, , "SF_BadaniaLotu", 3000, 3400, 5000, 3000
    DoCmd.Save acForm, "F_Partia"
    
    ' F_PlanPoboru
    Set fPlan = NowyForm("F_PlanPoboru", "PlanPoboru", True)
    DoCmd.Save acForm, "F_PlanPoboru"
    
    ' F_OczekiwaneBadanie
    Set fExp = NowyForm("F_OczekiwaneBadanie", "OczekiwaneBadanie", True)
    DoCmd.Save acForm, "F_OczekiwaneBadanie"
    
    ' F_NCR
    Set fNCR = NowyForm("F_NCR", "NCR", True)
    DoCmd.Save acForm, "F_NCR"
    
    ' Proste formularze s³ownikowe (opcjonalnie)
    On Error Resume Next
    NowyForm "F_Specyfikacje", "Specyfikacje", True: DoCmd.Save acForm, "F_Specyfikacje"
    NowyForm "F_Warstwy", "Warstwy", True: DoCmd.Save acForm, "F_Warstwy"
    NowyForm "F_Materialy", "Materialy", True: DoCmd.Save acForm, "F_Materialy"
    NowyForm "F_Recepty", "ReceptyMieszanek", True: DoCmd.Save acForm, "F_Recepty"
    NowyForm "F_Sita", "Sita", True: DoCmd.Save acForm, "F_Sita"
    NowyForm "F_Parametry", "ParametryJakosci", True: DoCmd.Save acForm, "F_Parametry"
    NowyForm "F_Obiekty", "Obiekty", True: DoCmd.Save acForm, "F_Obiekty"
    NowyForm "F_Lokalizacje", "Lokalizacje", True: DoCmd.Save acForm, "F_Lokalizacje"
    NowyForm "F_Sprzet", "SprzetPomiarowy", True: DoCmd.Save acForm, "F_Sprzet"
    NowyForm "F_Personel", "Personel", True: DoCmd.Save acForm, "F_Personel"
    NowyForm "F_Probka", "Probki", True: DoCmd.Save acForm, "F_Probka"
    NowyForm "F_ChainOfCustody", "ChainOfCustody", True: DoCmd.Save acForm, "F_ChainOfCustody"
    NowyForm "F_Laboratoria", "Laboratoria", True: DoCmd.Save acForm, "F_Laboratoria"
    On Error GoTo 0
    
    MsgBox "Utworzono formularze Partia/NCR i s³ownikowe.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d UtworzFormularze_PartieNCR: " & Err.Description, vbCritical
End Sub

Public Sub GenerujOczekiwaneBadania(ByVal PartiaID As Long)
    On Error GoTo ErrH
    Dim rsP As DAO.Recordset, rsPlan As DAO.Recordset
    Dim warstwaID As Long, specID As Long
    Dim dlugosc_m As Long
    Dim sql As String
    Dim req As Long, istnieje As Long
    Dim interwal As Variant, jedn As String
    Dim minLiczba As Long, i As Long
    Dim planID As Long, rodzajID As Variant, parametrID As Variant
    Dim termin As Date
    
    sql = "SELECT WarstwaID, SpecyfikacjaID, KM_Start_m, KM_End_m, Nz(DataWbudowania, Date()) AS DW " & _
          "FROM Partie WHERE PartiaID=" & PartiaID
    Set rsP = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rsP.EOF Then
        rsP.Close: Set rsP = Nothing
        MsgBox "Nie znaleziono Partii.", vbExclamation
        Exit Sub
    End If
    warstwaID = Nz(rsP!WarstwaID, 0)
    specID = Nz(rsP!SpecyfikacjaID, 0)
    dlugosc_m = Nz(rsP!KM_End_m, 0) - Nz(rsP!KM_Start_m, 0)
    termin = Nz(rsP!DW, Date)
    rsP.Close: Set rsP = Nothing
    
    If warstwaID = 0 Or specID = 0 Then
        MsgBox "Partia musi mieæ przypisan¹ Warstwê i Specyfikacjê.", vbExclamation
        Exit Sub
    End If
    If dlugosc_m < 0 Then
        MsgBox "KM_End_m musi byæ ? KM_Start_m.", vbExclamation
        Exit Sub
    End If
    
    sql = "SELECT PlanID, RodzajBadaniaID, ParametrID, Interwal, Jednostka, MinimalnaLiczba " & _
          "FROM PlanPoboru WHERE Aktywne=True AND WarstwaID=" & warstwaID & " AND SpecyfikacjaID=" & specID
    Set rsPlan = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rsPlan.EOF Then
        rsPlan.Close: Set rsPlan = Nothing
        MsgBox "Brak aktywnego planu poboru dla tej Partii.", vbInformation
        Exit Sub
    End If
    
    Do While Not rsPlan.EOF
        planID = rsPlan!PlanID
        rodzajID = rsPlan!RodzajBadaniaID
        parametrID = rsPlan!ParametrID
        interwal = rsPlan!Interwal
        jedn = Nz(rsPlan!Jednostka, "m")
        minLiczba = Nz(rsPlan!MinimalnaLiczba, 0)
        
        If LCase$(jedn) = "m" Then
            If Nz(interwal, 0) > 0 Then
                req = WorksheetCeil(dlugosc_m, CDbl(interwal))
                If req < minLiczba Then req = minLiczba
            Else
                req = minLiczba
            End If
        Else
            req = minLiczba
        End If
        
        ' Ile ju¿ istnieje?
        istnieje = DCount("*", "OczekiwaneBadanie", "PartiaID=" & PartiaID & " AND PlanID=" & planID)
        
        For i = 1 To (req - istnieje)
            ExecSQL "INSERT INTO OczekiwaneBadanie(PartiaID, PlanID, RodzajBadaniaID, ParametrID, Termin, Status) " & _
                    "VALUES(" & PartiaID & "," & planID & "," & Nz(rodzajID, "Null") & "," & Nz(parametrID, "Null") & ",#" & Format$(termin, "yyyy-mm-dd") & "#,'DoWykonania')"
        Next i
        
        rsPlan.MoveNext
    Loop
    rsPlan.Close: Set rsPlan = Nothing
    
    MsgBox "Wygenerowano oczekiwane badania dla Partii " & PartiaID & ".", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d GenerujOczekiwaneBadania: " & Err.Description, vbCritical
End Sub

Private Function WorksheetCeil(ByVal dl_m As Long, ByVal interwal As Double) As Long
    On Error GoTo ErrH
    If interwal <= 0 Then
        WorksheetCeil = 0
    Else
        WorksheetCeil = CLng(Int((dl_m + interwal - 1) / interwal))
    End If
    Exit Function
ErrH:
    WorksheetCeil = 0
End Function

Public Sub AktualizujStatusOczekiwanych()
    On Error GoTo ErrH
    ' Zrealizowane: ma BadanieID
    ExecSQL "UPDATE OczekiwaneBadanie SET Status='Zrealizowane' WHERE BadanieID Is Not Null"
    ' Po terminie: brak BadanieID i Termin < Date
    ExecSQL "UPDATE OczekiwaneBadanie SET Status='PoTermin' WHERE BadanieID Is Null AND Termin<Date()"
    ' DoWykonania: brak BadanieID i Termin ? Date
    ExecSQL "UPDATE OczekiwaneBadanie SET Status='DoWykonania' WHERE BadanieID Is Null AND Termin>=Date()"
    MsgBox "Zaktualizowano statusy oczekiwanych badañ.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d AktualizujStatusOczekiwanych: " & Err.Description, vbCritical
End Sub

Public Sub PokrycieBadan(ByVal PartiaID As Long)
    On Error GoTo ErrH
    Dim rs As DAO.Recordset
    Dim sql As String, info As String
    
    sql = "SELECT PlanID, Count(*) AS Wymagana " & _
          "FROM OczekiwaneBadanie WHERE PartiaID=" & PartiaID & " GROUP BY PlanID"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    info = "Pokrycie badañ dla Partii " & PartiaID & ":" & vbCrLf
    Do While Not rs.EOF
        Dim planID As Long, wykonane As Long
        planID = rs!PlanID
        wykonane = DCount("*", "OczekiwaneBadanie", "PartiaID=" & PartiaID & " AND PlanID=" & planID & " AND BadanieID Is Not Null")
        info = info & "- PlanID=" & planID & ": " & wykonane & " / " & rs!Wymagana & vbCrLf
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    MsgBox info, vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d PokrycieBadan: " & Err.Description, vbCritical
End Sub

Public Sub UtworzIZlinkujBE(ByVal SciezkaBE As String)
    On Error GoTo ErrH
    Dim dbBE As DAO.Database
    Dim td As DAO.TableDef
    Dim srcDB As DAO.Database
    Dim tName As String
    
    ' Utwórz pusty plik BE
    Set dbBE = DBEngine.CreateDatabase(SciezkaBE, dbLangGeneral, dbVersion120)
    dbBE.Close
    Set dbBE = Nothing
    
    Set srcDB = CurrentDb
    
    ' Eksport wszystkich tabel (bez MSys)
    For Each td In srcDB.TableDefs
        tName = td.Name
        If Left$(tName, 4) <> "MSys" Then
            DoCmd.TransferDatabase acExport, "Microsoft Access", SciezkaBE, acTable, tName, tName, False
        End If
    Next td
    
    ' Usuniêcie lokalnych tabel i podlinkowanie
    For Each td In srcDB.TableDefs
        tName = td.Name
        If Left$(tName, 4) <> "MSys" Then
            DoCmd.DeleteObject acTable, tName
            DoCmd.TransferDatabase acLink, "Microsoft Access", SciezkaBE, acTable, tName, tName, False
        End If
    Next td
    
    MsgBox "Utworzono backend i podlinkowano tabele.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d UtworzIZlinkujBE: " & Err.Description, vbCritical
End Sub

Public Sub BackupBE(ByVal SciezkaBE As String, ByVal FolderBackup As String)
    On Error GoTo ErrH
    Dim dest As String
    If Right$(FolderBackup, 1) = "\" Or Right$(FolderBackup, 1) = "/" Then
        dest = FolderBackup & "BE_backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".accdb"
    Else
        dest = FolderBackup & "\BE_backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".accdb"
    End If
    FileCopy SciezkaBE, dest
    MsgBox "Wykonano kopiê BE: " & dest, vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d BackupBE: " & Err.Description, vbCritical
End Sub
