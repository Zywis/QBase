Attribute VB_Name = "M06_Raporty"
Option Compare Database
Option Explicit

'==========================================================
' Modul: M06_Raporty
' Tworzenie raportow podstawowych
'==========================================================

Private Sub DropReportIfExists(Nazwa As String)
    On Error Resume Next
    DoCmd.DeleteObject acReport, Nazwa
    On Error GoTo 0
End Sub

Private Function NowyRaport(Nazwa As String, Tabela As String) As Report
    On Error GoTo ErrH
    Dim rpt As Report
	DropReportIfExists Nazwa
    Set rpt = CreateReport
    rpt.RecordSource = Tabela
    rpt.Caption = Nazwa
    DoCmd.Save acReport, rpt.Name
    DoCmd.Rename Nazwa, acReport, rpt.Name
    Set NowyRaport = Reports(Nazwa)
    Exit Function
ErrH:
    MsgBox "Blad NowyRaport: " & Err.Description, vbCritical
    Set NowyRaport = Nothing
End Function

Public Sub UtworzRaporty()
    On Error GoTo ErrH
    
    '------------------------------------------------------
    ' R_ProtokolBadania
    '------------------------------------------------------
    Dim rptProt As Report
    Set rptProt = NowyRaport("R_ProtokolBadania", "Badania")
    KonfigurujRaportProtokol rptProt

    '------------------------------------------------------
    ' R_ZestawienieJakosci
    '------------------------------------------------------
    Dim rptZest As Report
    Set rptZest = NowyRaport("R_ZestawienieJakosci", "Q_Wyniki_Z_Ocena")
    KonfigurujRaportZestawienie rptZest

    '------------------------------------------------------
    ' R_Niespelnienia
    '------------------------------------------------------
    Dim rptNiesp As Report
    Set rptNiesp = NowyRaport("R_Niespelnienia", "Q_Niespelnienia_Param")
    KonfigurujRaportNiespelnienia rptNiesp

    '------------------------------------------------------
    ' R_PokrycieBadan
    '------------------------------------------------------
    Dim rptPokrycie As Report
    Set rptPokrycie = NowyRaport("R_PokrycieBadan", "Q_Loty_Pokrycie_Summary")
    KonfigurujRaportPokrycie rptPokrycie
	
    '------------------------------------------------------
    ' R_RejestrNCR
    '------------------------------------------------------
    Dim rptNCR As Report
    Set rptNCR = NowyRaport("R_RejestrNCR", "NCR")
    KonfigurujRaportNCR rptNCR

    Exit Sub
ErrH:
    MsgBox "Blad UtworzRaporty: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportProtokol(rpt As Report)
    On Error GoTo ErrH

    Dim lblTitle As Control
    Set lblTitle = CreateReportControl(rpt.Name, acLabel, acReportHeader, , "Protokol badania", 200, 100, 5000, 400)
    lblTitle.FontSize = 16
    lblTitle.FontBold = True

    rpt.Section(acReportHeader).Height = 600
    rpt.Section(acPageHeader).Height = 500
    rpt.Section(acDetail).Height = 4200

    DodajPoleRaportu rpt, "BadanieID", "ID badania", 200, 2000
    DodajPoleRaportu rpt, "DataBadania", "Data badania", 2200, 2000
    DodajPoleRaportu rpt, "ProtokolNr", "Nr protokolu", 4200, 2000
    DodajPoleRaportu rpt, "ProbkaKod", "Kod probki", 6200, 2000

    Dim lblChart As Control
    Set lblChart = CreateReportControl(rpt.Name, acLabel, acDetail, , "Krzywa uziarnienia", 200, 800, 3000, 300)
    lblChart.FontBold = True

    Dim ctlChart As Control
    Set ctlChart = CreateReportControl(rpt.Name, acGraph, acDetail, , , 200, 1200, 6000, 2800)
    ctlChart.Name = "grKrzywaUziarnienia"
    ctlChart.RowSourceType = "Table/Query"
    ctlChart.RowSource = "SELECT S.Kolejnosc, S.Rozmiar_mm, UW.ProcPrzechodzenia " & _
                         "FROM (UziarnienieWynik AS UW INNER JOIN Sita AS S ON UW.SitoID=S.SitoID) " & _
                         "WHERE UW.BadanieID=[BadanieID] ORDER BY S.Kolejnosc"
    On Error Resume Next
    ctlChart.Object.ChartType = 17 ' XY scatter with lines
    On Error GoTo ErrH

    KonfigurujRaportMinMax

    Dim lblTabela As Control
    Set lblTabela = CreateReportControl(rpt.Name, acLabel, acDetail, , "Envelope (Min/Max)", 6400, 800, 3000, 300)
    lblTabela.FontBold = True

    Dim ctlSub As Control
    Set ctlSub = CreateReportControl(rpt.Name, acSubreport, acDetail, , , 6400, 1200, 3200, 2800)
    ctlSub.Name = "srpEnvelope"
    ctlSub.SourceObject = "Report.R_ProtokolBadania_MinMax"
    ctlSub.LinkMasterFields = "BadanieID"
    ctlSub.LinkChildFields = "BadanieID"

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
    MsgBox "Blad KonfigurujRaportProtokol: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportMinMax()
    On Error GoTo ErrH

    Dim rpt As Report
    Set rpt = NowyRaport("R_ProtokolBadania_MinMax", "Q_Uziarnienie_Ocena")

    rpt.Section(acReportHeader).Height = 0
    rpt.Section(acPageHeader).Height = 400
    rpt.Section(acDetail).Height = 360

    DodajKolumne rpt, "SitoID", "Sito", 200, 900
    DodajKolumne rpt, "ProcPrzechodzenia", "Proc przech.", 1100, 1200
    DodajKolumne rpt, "MinProc", "Min proc", 2400, 900
    DodajKolumne rpt, "MaxProc", "Max proc", 3400, 900

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
   MsgBox "Blad KonfigurujRaportMinMax: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportZestawienie(rpt As Report)
    On Error GoTo ErrH

    DodajNaglowekZTytulem rpt, "Zestawienie jakosci"
    rpt.Section(acDetail).Height = 360

    DodajKolumne rpt, "BadanieID", "Badanie", 200, 900
    DodajKolumne rpt, "ParametrID", "Parametr", 1100, 1200
    DodajKolumne rpt, "Wartosc", "Wartosc", 2400, 1200
    DodajKolumne rpt, "TypKryterium", "Kryterium", 3700, 1200
    DodajKolumne rpt, "MinWartosc", "Min", 5000, 900
    DodajKolumne rpt, "MaxWartosc", "Max", 5900, 900
    DodajKolumne rpt, "Spelnia", "Spelnia", 6800, 900

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
    MsgBox "Blad KonfigurujRaportZestawienie: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportNiespelnienia(rpt As Report)
    On Error GoTo ErrH

    DodajNaglowekZTytulem rpt, "Niespelnione parametry"
    rpt.Section(acDetail).Height = 360

    DodajKolumne rpt, "BadanieID", "Badanie", 200, 900
    DodajKolumne rpt, "ParametrID", "Parametr", 1100, 1200
    DodajKolumne rpt, "Wartosc", "Wartosc", 2400, 1200
    DodajKolumne rpt, "MinWartosc", "Min", 3700, 900
    DodajKolumne rpt, "MaxWartosc", "Max", 4600, 900
    DodajKolumne rpt, "Spelnia", "Spelnia", 5500, 900

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
    MsgBox "Blad KonfigurujRaportNiespelnienia: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportPokrycie(rpt As Report)
    On Error GoTo ErrH

    DodajNaglowekZTytulem rpt, "Pokrycie badan"
    rpt.Section(acDetail).Height = 360

    DodajKolumne rpt, "PartiaID", "Partia", 200, 900
    DodajKolumne rpt, "Typ", "Typ", 1100, 900
    DodajKolumne rpt, "ID", "ID", 2000, 900
    DodajKolumne rpt, "Ilosc", "Ilosc", 2900, 900

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
    MsgBox "Blad KonfigurujRaportPokrycie: " & Err.Description, vbCritical
End Sub

Private Sub KonfigurujRaportNCR(rpt As Report)
    On Error GoTo ErrH

    DodajNaglowekZTytulem rpt, "Rejestr NCR"
    rpt.Section(acDetail).Height = 360

    DodajKolumne rpt, "NCRID", "NCR", 200, 900
    DodajKolumne rpt, "PartiaID", "Partia", 1100, 900
    DodajKolumne rpt, "BadanieID", "Badanie", 2000, 900
    DodajKolumne rpt, "DataZgloszenia", "Data zgl.", 2900, 1200
    DodajKolumne rpt, "Klasyfikacja", "Klasyfikacja", 4200, 1500
    DodajKolumne rpt, "Status", "Status", 5800, 1200
    DodajKolumne rpt, "Termin", "Termin", 7100, 1200

    DoCmd.Save acReport, rpt.Name
    Exit Sub
ErrH:
    MsgBox "Blad KonfigurujRaportNCR: " & Err.Description, vbCritical
End Sub

Private Sub DodajPoleRaportu(rpt As Report, Pole As String, Naglowek As String, LeftPos As Long, Szerokosc As Long)
    Dim txt As Control
    Set txt = CreateReportControl(rpt.Name, acTextBox, acDetail, , Pole, LeftPos, 200, Szerokosc, 300)
    txt.Name = "txt" & Pole
    If InStr(1, Pole, "Data", vbTextCompare) > 0 Then
        txt.Format = "Short Date"
    End If

    Dim lbl As Control
    Set lbl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , Naglowek, LeftPos, 100, Szerokosc, 300)
    lbl.Name = "lbl" & Pole
End Sub

Private Sub DodajKolumne(rpt As Report, ControlSource As String, Caption As String, LeftPos As Long, Szerokosc As Long)
    Dim txt As Control
    Set txt = CreateReportControl(rpt.Name, acTextBox, acDetail, , ControlSource, LeftPos, 60, Szerokosc, 260)
    txt.Name = "txt" & Replace(ControlSource, " ", "")
    If StrComp(ControlSource, "Spelnia", vbTextCompare) = 0 Then
        txt.Format = "Yes/No"
    ElseIf InStr(1, ControlSource, "Data", vbTextCompare) > 0 Then
        txt.Format = "Short Date"
    End If

    Dim lbl As Control
    Set lbl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , Caption, LeftPos, 20, Szerokosc, 260)
    lbl.Name = "lbl" & Replace(ControlSource, " ", "")
End Sub

Private Sub DodajNaglowekZTytulem(rpt As Report, Tytul As String)
    Dim lbl As Control
    Set lbl = CreateReportControl(rpt.Name, acLabel, acReportHeader, , Tytul, 200, 100, 5000, 400)
    lbl.FontSize = 14
    lbl.FontBold = True
    rpt.Section(acReportHeader).Height = 600
    rpt.Section(acPageHeader).Height = 360
End Sub
