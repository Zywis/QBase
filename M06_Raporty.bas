Attribute VB_Name = "M06_Raporty"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M06_Raporty
' Tworzenie raportów podstawowych
'==========================================================

Private Function NowyRaport(Nazwa As String, Tabela As String) As Report
    On Error GoTo ErrH
    Dim rpt As Report
    Set rpt = CreateReport
    rpt.RecordSource = Tabela
    rpt.Caption = Nazwa
    DoCmd.Save acReport, rpt.Name
    DoCmd.Rename Nazwa, acReport, rpt.Name
    Set NowyRaport = Reports(Nazwa)
    Exit Function
ErrH:
    MsgBox "B³¹d NowyRaport: " & Err.Description, vbCritical
    Set NowyRaport = Nothing
End Function

Public Sub UtworzRaporty()
    On Error GoTo ErrH
    
    '------------------------------------------------------
    ' R_ProtokolBadania (bez wykresu, tylko tabela)
    '------------------------------------------------------
    Dim rptProt As Report
    Set rptProt = NowyRaport("R_ProtokolBadania", "Badania")
    
    '------------------------------------------------------
    ' R_ZestawienieJakosci
    '------------------------------------------------------
    Dim rptZest As Report
    Set rptZest = NowyRaport("R_ZestawienieJakosci", "Q_Wyniki_Z_Ocena")
    
    '------------------------------------------------------
    ' R_Niespelnienia
    '------------------------------------------------------
    Dim rptNiesp As Report
    Set rptNiesp = NowyRaport("R_Niespelnienia", "Q_Niespelnienia_Param")
    
    '------------------------------------------------------
    ' R_PokrycieBadan
    '------------------------------------------------------
    Dim rptPokrycie As Report
    Set rptPokrycie = NowyRaport("R_PokrycieBadan", "Q_Loty_Pokrycie_Summary")
    
    '------------------------------------------------------
    ' R_RejestrNCR
    '------------------------------------------------------
    Dim rptNCR As Report
    Set rptNCR = NowyRaport("R_RejestrNCR", "NCR")
    
    Exit Sub
ErrH:
    MsgBox "B³¹d UtworzRaporty: " & Err.Description, vbCritical
End Sub
