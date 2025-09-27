Attribute VB_Name = "M07_LogikaOceny"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M07_LogikaOceny
' Logika ocen, walidacje, self-test, funkcje pomocnicze
' (bez logowania do LogZdarzen – zgodnie z ustaleniami)
'==========================================================

Private Sub ExecSQL(ByVal SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d ExecSQL: " & Err.Description, vbCritical
End Sub

Public Function FormatKM(ByVal m As Long) As String
    On Error GoTo ErrH
    Dim km As Long
    Dim plusy As Long
    If m < 0 Then
        FormatKM = "km ?"
        Exit Function
    End If
    km = m \ 1000
    plusy = m Mod 1000
    FormatKM = "km " & km & "+" & Format(plusy, "000")
    Exit Function
ErrH:
    MsgBox "B³¹d FormatKM: " & Err.Description, vbCritical
    FormatKM = ""
End Function

Public Function OcenaParametru( _
    ByVal TypKryterium As String, _
    ByVal MinWartosc As Variant, _
    ByVal MaxWartosc As Variant, _
    ByVal Nominal As Variant, _
    ByVal TolMinus As Variant, _
    ByVal TolPlus As Variant, _
    ByVal Wartosc As Variant) As Variant
    On Error GoTo ErrH
    If IsNull(Wartosc) Then
        OcenaParametru = Null
        Exit Function
    End If
    
    Select Case UCase$(TypKryterium)
        Case "ZAKRES"
            If IsNull(MinWartosc) Or IsNull(MaxWartosc) Then
                OcenaParametru = Null
            Else
                OcenaParametru = (Wartosc >= MinWartosc And Wartosc <= MaxWartosc)
            End If
        Case ">=MIN", ">= MIN"
            If IsNull(MinWartosc) Then
                OcenaParametru = Null
            Else
                OcenaParametru = (Wartosc >= MinWartosc)
            End If
        Case "<=MAX", "<= MAX"
            If IsNull(MaxWartosc) Then
                OcenaParametru = Null
            Else
                OcenaParametru = (Wartosc <= MaxWartosc)
            End If
        Case "+/-", "+-"
            If IsNull(Nominal) Or (IsNull(TolMinus) And IsNull(TolPlus)) Then
                OcenaParametru = Null
            Else
                Dim lo As Double, hi As Double
                lo = Nz(Nominal - Nz(TolMinus, 0), 0)
                hi = Nz(Nominal + Nz(TolPlus, 0), 0)
                OcenaParametru = (Wartosc >= lo And Wartosc <= hi)
            End If
        Case Else
            OcenaParametru = Null
    End Select
    Exit Function
ErrH:
    MsgBox "B³¹d OcenaParametru: " & Err.Description, vbCritical
    OcenaParametru = Null
End Function

Public Function OcenaUziarnienia(ByVal BadanieID As Long) As Variant
    On Error GoTo ErrH
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' SprawdŸ czy istnieje envelope dla pary Warstwa+Spec
    sql = "SELECT Count(*) AS Cnt " & _
          "FROM (Badania B INNER JOIN UziarnienieWymaganie UW ON B.SpecyfikacjaID=UW.SpecyfikacjaID " & _
          "AND B.WarstwaID=UW.WarstwaID) " & _
          "WHERE B.BadanieID=" & BadanieID
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Or rs!Cnt = 0 Then
        rs.Close: Set rs = Nothing
        OcenaUziarnienia = Null   ' brak envelope › brak oceny
        Exit Function
    End If
    rs.Close: Set rs = Nothing
    
    ' Ocena: wszystkie sita w zakresie
    sql = "SELECT Min(IIf(UWyn.ProcPrzechodzenia Between UWym.MinProc And UWym.MaxProc, 1, 0)) AS OkAll " & _
          "FROM (UziarnienieWynik UWyn INNER JOIN Badania B ON UWyn.BadanieID=B.BadanieID) " & _
          "INNER JOIN UziarnienieWymaganie UWym ON B.SpecyfikacjaID=UWym.SpecyfikacjaID " & _
          "AND B.WarstwaID=UWym.WarstwaID AND UWyn.SitoID=UWym.SitoID " & _
          "WHERE UWyn.BadanieID=" & BadanieID
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        OcenaUziarnienia = Null
    Else
        If IsNull(rs!OkAll) Then
            OcenaUziarnienia = Null
        Else
            OcenaUziarnienia = (rs!OkAll = 1)
        End If
    End If
    rs.Close: Set rs = Nothing
    Exit Function
ErrH:
    MsgBox "B³¹d OcenaUziarnienia: " & Err.Description, vbCritical
    OcenaUziarnienia = Null
End Function

Public Function PobierzDomyslnaSpec(ByVal WarstwaID As Long, Optional ByVal ObiektID As Variant) As Long
    On Error GoTo ErrH
    Dim rs As DAO.Recordset
    Dim sql As String
    
    If Not IsMissing(ObiektID) And Not IsNull(ObiektID) Then
        sql = "SELECT TOP 1 SpecyfikacjaID FROM ObiektWarstwaSpec " & _
              "WHERE ObiektID=" & CLng(ObiektID) & " AND WarstwaID=" & WarstwaID
        Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        If Not rs.EOF Then
            PobierzDomyslnaSpec = Nz(rs!SpecyfikacjaID, 0)
            rs.Close: Set rs = Nothing
            Exit Function
        End If
        rs.Close: Set rs = Nothing
    End If
    
    sql = "SELECT TOP 1 SpecyfikacjaID FROM WarstwaSpecDefault WHERE WarstwaID=" & WarstwaID
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        PobierzDomyslnaSpec = Nz(rs!SpecyfikacjaID, 0)
    Else
        PobierzDomyslnaSpec = 0
    End If
    rs.Close: Set rs = Nothing
    Exit Function
ErrH:
    MsgBox "B³¹d PobierzDomyslnaSpec: " & Err.Description, vbCritical
    PobierzDomyslnaSpec = 0
End Function

Public Sub DodajParametryDomyslne(ByVal BadanieID As Long)
    On Error GoTo ErrH
    Dim rsB As DAO.Recordset
    Dim sql As String, typWarstwy As String
    Dim potrzebne As Variant, i As Long
    Dim parID As Long, jedn As String
    
    sql = "SELECT W.TypWarstwy FROM Badania B INNER JOIN Warstwy W ON B.WarstwaID=W.WarstwaID " & _
          "WHERE B.BadanieID=" & BadanieID
    Set rsB = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rsB.EOF Then
        rsB.Close: Set rsB = Nothing
        MsgBox "Nie znaleziono badania.", vbExclamation
        Exit Sub
    End If
    typWarstwy = Nz(rsB!TypWarstwy, "")
    rsB.Close: Set rsB = Nothing
    
    ' Zestawy parametrów wg typu warstwy
    If UCase$(typWarstwy) = "NASYP" Or UCase$(typWarstwy) = "KNN" Or _
       UCase$(typWarstwy) = "STABILIZACJA" Or UCase$(typWarstwy) = "MROZO" Then
        potrzebne = Split("?b|wopt", "|")
    Else
        ' MMA/SMA/AC i inne mieszanki asfaltowe
        potrzebne = Split("?b|Pb|Pb_sol", "|")
    End If
    
    For i = LBound(potrzebne) To UBound(potrzebne)
        parID = ParametrID_BySymbol(CStr(potrzebne(i)), jedn)
        If parID <> 0 Then
            ' Wstaw, jeœli brak
            sql = "SELECT Count(*) AS Cnt FROM WynikiBadania " & _
                  "WHERE BadanieID=" & BadanieID & " AND ParametrID=" & parID
            If DCount("*", "WynikiBadania", "BadanieID=" & BadanieID & " AND ParametrID=" & parID) = 0 Then
                ExecSQL "INSERT INTO WynikiBadania(BadanieID, ParametrID, Jednostka) " & _
                        "VALUES(" & BadanieID & "," & parID & ",'" & Replace(Nz(jedn, ""), "'", "''") & "')"
            End If
        End If
    Next i
    
    MsgBox "Dodano domyœlne parametry.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d DodajParametryDomyslne: " & Err.Description, vbCritical
End Sub

Private Function ParametrID_BySymbol(ByVal Symbol As String, ByRef JednostkaOut As String) As Long
    On Error GoTo ErrH
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ParametrID, Jednostka FROM ParametryJakosci " & _
        "WHERE SymbolParametru='" & Replace(Symbol, "'", "''") & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        ParametrID_BySymbol = rs!ParametrID
        JednostkaOut = Nz(rs!Jednostka, "")
    Else
        ParametrID_BySymbol = 0
        JednostkaOut = ""
    End If
    rs.Close: Set rs = Nothing
    Exit Function
ErrH:
    MsgBox "B³¹d ParametrID_BySymbol: " & Err.Description, vbCritical
    ParametrID_BySymbol = 0
    JednostkaOut = ""
End Function

Public Sub WstawSitaDlaBadania(ByVal BadanieID As Long)
    On Error GoTo ErrH
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim specID As Long, warstwaID As Long
    
    sql = "SELECT SpecyfikacjaID, WarstwaID FROM Badania WHERE BadanieID=" & BadanieID
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        MsgBox "Nie znaleziono badania.", vbExclamation
        Exit Sub
    End If
    specID = Nz(rs!SpecyfikacjaID, 0)
    warstwaID = Nz(rs!WarstwaID, 0)
    rs.Close: Set rs = Nothing
    
    If specID = 0 Or warstwaID = 0 Then
        MsgBox "Badanie musi mieæ przypisan¹ Specyfikacjê i Warstwê.", vbExclamation
        Exit Sub
    End If
    
    ' Walidacja: envelope istnieje?
    If DCount("*", "UziarnienieWymaganie", _
              "SpecyfikacjaID=" & specID & " AND WarstwaID=" & warstwaID) = 0 Then
        MsgBox "Brak envelope uziarnienia dla danej Warstwy i Specyfikacji. Uzupe³nij wymagania.", vbExclamation
        Exit Sub
    End If
    
    ' Wstaw brakuj¹ce sita tylko te, które s¹ w envelope dla pary (Spec, Warstwa)
    sql = "INSERT INTO UziarnienieWynik(BadanieID, SitoID) " & _
          "SELECT " & BadanieID & ", UW.SitoID " & _
          "FROM UziarnienieWymaganie UW " & _
          "WHERE UW.SpecyfikacjaID=" & specID & " AND UW.WarstwaID=" & warstwaID & " " & _
          "AND NOT EXISTS (SELECT 1 FROM UziarnienieWynik W " & _
          "WHERE W.BadanieID=" & BadanieID & " AND W.SitoID=UW.SitoID)"
    ExecSQL sql
    MsgBox "Za³adowano pozycje sit dla badania.", vbInformation
    Exit Sub
ErrH:
    MsgBox "B³¹d WstawSitaDlaBadania: " & Err.Description, vbCritical
End Sub

Public Function SelfTest_Kompletnosc() As Boolean
    On Error GoTo ErrH
    Dim brak As String
    brak = ""
    
    ' Tabele (wybór kluczowych z Sekcji 16)
    SprawdzTabele Split( _
        "Specyfikacje|Warstwy|WarstwaSpecDefault|Materialy|WarstwaMaterial|ReceptyMieszanek|" & _
        "Obiekty|Lokalizacje|ParametryJakosci|WymaganiaParametru|Sita|UziarnienieWymaganie|" & _
        "UziarnienieWynik|RodzajeBadania|Badania|WynikiBadania|Partie|PlanPoboru|" & _
        "OczekiwaneBadanie|NCR|SprzetPomiarowy|Personel|Probki|ChainOfCustody|Laboratoria|" & _
        "LogZdarzen|Zalaczniki|IntegracjeODBC", "|"), brak
    
    ' Kwerendy
    SprawdzKwerendy Split( _
        "Q_Wyniki_Z_Ocena|Q_Niespelnienia_Param|Q_Uziarnienie_Ocena|" & _
        "Q_Uziarnienie_Krzywa_Spelnia|Q_Uziarnienie_Crosstab|Q_Insert_Uziarnienie_Sita|" & _
        "Q_Partie_Combo_RS|Q_SF_Oczekiwane_RS|Q_SF_BadaniaLotu_RS|" & _
        "Q_Loty_Pokrycie_Rodzaje|Q_Loty_Pokrycie_Parametry|Q_Loty_Pokrycie|Q_Loty_Pokrycie_Summary|" & _
        "Q_Badania_DoWykonania", "|"), brak
    
    ' Formularze
    SprawdzFormy Split( _
        "F_Badanie|SF_Uziarnienie|SF_Wyniki|F_WarstwaSpecDefault|" & _
        "F_Partia|SF_Oczekiwane|SF_BadaniaLotu|F_PlanPoboru|F_OczekiwaneBadanie|F_NCR|" & _
        "F_Specyfikacje|F_Warstwy|F_Materialy|F_Recepty|F_Sita|F_Parametry|F_Obiekty|F_Lokalizacje|" & _
        "F_Sprzet|F_Personel|F_Probka|F_ChainOfCustody|F_Laboratoria", "|"), brak
    
    ' Raporty
    SprawdzRaporty Split( _
        "R_ProtokolBadania|R_ZestawienieJakosci|R_Niespelnienia|R_PokrycieBadan|R_RejestrNCR", "|"), brak
    
    ' Indeksy unikalne wg M02 (sprawdzamy po nazwach)
    SprawdzIndeks "Specyfikacje", "UQ_Spec_Kod", brak
    SprawdzIndeks "WarstwaSpecDefault", "UQ_WSD_Warstwa", brak
    SprawdzIndeks "WarstwaMaterial", "UQ_WM_WarstwaMaterial", brak
    SprawdzIndeks "ReceptyMieszanek", "UQ_Rec_WarSpecKod", brak
    SprawdzIndeks "WymaganiaParametru", "UQ_Wym_WarSpecParam", brak
    SprawdzIndeks "UziarnienieWymaganie", "UQ_UWym_WarSpecSito", brak
    SprawdzIndeks "UziarnienieWynik", "UQ_UWyn_BadSito", brak
    SprawdzIndeks "WynikiBadania", "UQ_Wyn_BadParamProb", brak
    
    If Len(brak) = 0 Then
        MsgBox "SelfTest OK – kompletnoœæ potwierdzona.", vbInformation
        SelfTest_Kompletnosc = True
    Else
        MsgBox "SelfTest – wykryto braki:" & vbCrLf & brak, vbExclamation
        SelfTest_Kompletnosc = False
    End If
    Exit Function
ErrH:
    MsgBox "B³¹d SelfTest_Kompletnosc: " & Err.Description, vbCritical
    SelfTest_Kompletnosc = False
End Function

Private Sub SprawdzTabele(ByVal Nazwy As Variant, ByRef Braki As String)
    Dim i As Long
    For i = LBound(Nazwy) To UBound(Nazwy)
        If Not IstniejeTabela(CStr(Nazwy(i))) Then
            Braki = Braki & "- Brak tabeli: " & Nazwy(i) & vbCrLf
        End If
    Next i
End Sub

Private Sub SprawdzKwerendy(ByVal Nazwy As Variant, ByRef Braki As String)
    Dim i As Long
    For i = LBound(Nazwy) To UBound(Nazwy)
        If Not IstniejeKwerenda(CStr(Nazwy(i))) Then
            Braki = Braki & "- Brak kwerendy: " & Nazwy(i) & vbCrLf
        End If
    Next i
End Sub

Private Sub SprawdzFormy(ByVal Nazwy As Variant, ByRef Braki As String)
    Dim i As Long
    For i = LBound(Nazwy) To UBound(Nazwy)
        If Len(CStr(Nazwy(i))) = 0 Then GoTo NextI
        If Not IstniejeForma(CStr(Nazwy(i))) Then
            ' Czêœæ z nich (s³ownikowe) nie s¹ wymagane jeœli nie tworzone – oznacz jako ostrze¿enie
            Braki = Braki & "- (Ostrze¿enie) Brak formularza: " & Nazwy(i) & vbCrLf
        End If
NextI:
    Next i
End Sub

Private Sub SprawdzRaporty(ByVal Nazwy As Variant, ByRef Braki As String)
    Dim i As Long
    For i = LBound(Nazwy) To UBound(Nazwy)
        If Not IstniejeRaport(CStr(Nazwy(i))) Then
            Braki = Braki & "- (Ostrze¿enie) Brak raportu: " & Nazwy(i) & vbCrLf
        End If
    Next i
End Sub

Private Sub SprawdzIndeks(ByVal Tabela As String, ByVal Indeks As String, ByRef Braki As String)
    On Error GoTo ErrH
    Dim td As DAO.TableDef, ix As DAO.Index
    Set td = CurrentDb.TableDefs(Tabela)
    For Each ix In td.Indexes
        If ix.Name = Indeks Then Exit Sub
    Next ix
    Braki = Braki & "- Brak indeksu: " & Indeks & " w tabeli " & Tabela & vbCrLf
    Exit Sub
ErrH:
    Braki = Braki & "- Brak tabeli do weryfikacji indeksu: " & Tabela & vbCrLf
End Sub

Private Function IstniejeTabela(ByVal Nazwa As String) As Boolean
    On Error GoTo ErrH
    Dim t As DAO.TableDef
    For Each t In CurrentDb.TableDefs
        If t.Name = Nazwa Then
            IstniejeTabela = True
            Exit Function
        End If
    Next
    IstniejeTabela = False
    Exit Function
ErrH:
    IstniejeTabela = False
End Function

Private Function IstniejeKwerenda(ByVal Nazwa As String) As Boolean
    On Error Resume Next
    Dim q As DAO.QueryDef
    Set q = CurrentDb.QueryDefs(Nazwa)
    IstniejeKwerenda = (Err.Number = 0)
    Err.Clear
End Function

Private Function IstniejeForma(ByVal Nazwa As String) As Boolean
    On Error Resume Next
    Dim accObj As AccessObject
    Set accObj = CurrentProject.AllForms(Nazwa)
    IstniejeForma = (Err.Number = 0)
    Err.Clear
End Function

Private Function IstniejeRaport(ByVal Nazwa As String) As Boolean
    On Error Resume Next
    Dim accObj As AccessObject
    Set accObj = CurrentProject.AllReports(Nazwa)
    IstniejeRaport = (Err.Number = 0)
    Err.Clear
End Function
