Attribute VB_Name = "M05_Formularze"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M05_Formularze
' Tworzenie formularzy podstawowych (F_Badanie, SF_Uziarnienie,
' SF_Wyniki, F_WarstwaSpecDefault)
'==========================================================

Private Function NowyFormularz(Nazwa As String, Tabela As String, TrybCi¹g³y As Boolean) As Form
    On Error GoTo ErrH
    Dim frm As Form
    Set frm = CreateForm
    frm.RecordSource = Tabela
    frm.DefaultView = IIf(TrybCi¹g³y, 1, 0) ' 1=Continuous, 0=Single
    frm.Caption = Nazwa
    DoCmd.Save acForm, frm.Name
    DoCmd.Rename Nazwa, acForm, frm.Name
    Set NowyFormularz = Forms(Nazwa)
    Exit Function
ErrH:
    MsgBox "B³¹d NowyFormularz: " & Err.Description, vbCritical
    Set NowyFormularz = Nothing
End Function

Public Sub UtworzFormularze()
    On Error GoTo ErrH
    
    '------------------------------------------------------
    ' Podformularze: SF_Uziarnienie, SF_Wyniki
    '------------------------------------------------------
    Dim frmUzi As Form, frmWyn As Form
    Set frmUzi = NowyFormularz("SF_Uziarnienie", "UziarnienieWynik", True)
    Set frmWyn = NowyFormularz("SF_Wyniki", "WynikiBadania", True)
    
    '------------------------------------------------------
    ' Formularz F_WarstwaSpecDefault
    '------------------------------------------------------
    Dim frmWSD As Form
    Set frmWSD = NowyFormularz("F_WarstwaSpecDefault", "WarstwaSpecDefault", True)
    
    '------------------------------------------------------
    ' Formularz g³ówny: F_Badanie
    '------------------------------------------------------
    Dim frmBad As Form
    Set frmBad = NowyFormularz("F_Badanie", "Badania", False)
    
    ' Pola g³ówne
    CreateControl "F_Badanie", acComboBox, acDetail, , "ObiektID", 100, 100, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "LokalizacjaID", 100, 500, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "WarstwaID", 100, 900, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "SpecyfikacjaID", 100, 1300, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "MaterialID", 100, 1700, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "ReceptaID", 100, 2100, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "PartiaID", 100, 2500, 2000, 300
    CreateControl "F_Badanie", acComboBox, acDetail, , "RodzajBadaniaID", 100, 2900, 2000, 300
    CreateControl "F_Badanie", acTextBox, acDetail, , "DataBadania", 100, 3300, 2000, 300
    
    ' Przyciski
    Dim btnKrzywa As Control, btnParam As Control, btnPowiaz As Control
    Set btnKrzywa = CreateControl("F_Badanie", acCommandButton, acDetail, , , 2500, 100, 2000, 300)
    btnKrzywa.Caption = "Za³aduj krzyw¹ (sita)"
    btnKrzywa.OnClick = "=M07_LogikaOceny.WstawSitaDlaBadania([BadanieID])"
    
    Set btnParam = CreateControl("F_Badanie", acCommandButton, acDetail, , , 2500, 500, 2000, 300)
    btnParam.Caption = "Dodaj parametry"
    btnParam.OnClick = "=M07_LogikaOceny.DodajParametryDomyslne([BadanieID])"
    
    Set btnPowiaz = CreateControl("F_Badanie", acCommandButton, acDetail, , , 2500, 900, 2000, 300)
    btnPowiaz.Caption = "Powi¹¿ z oczekiwanym"
    btnPowiaz.OnClick = "=MsgBox(""Powi¹zanie z oczekiwanym - logika w M08"")"
    
    ' Podformularze
    Dim sfUzi As SubForm, sfWyn As SubForm
    Set sfUzi = CreateControl("F_Badanie", acSubform, acDetail, , "SF_Uziarnienie", 100, 3800, 4000, 2000)
    Set sfWyn = CreateControl("F_Badanie", acSubform, acDetail, , "SF_Wyniki", 100, 6000, 4000, 2000)
    
    DoCmd.Save acForm, "F_Badanie"
    
    Exit Sub
ErrH:
    MsgBox "B³¹d UtworzFormularze: " & Err.Description, vbCritical
End Sub
