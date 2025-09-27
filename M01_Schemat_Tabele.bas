Attribute VB_Name = "M01_Schemat_Tabele"
Option Compare Database
Option Explicit

' M01: Schemat tabel  utworzenie bez uycia polecenia CREATE TABLE

'==================== Narzêdzia ====================
Private Sub ExecSQL(ByVal s As String)
    On Error GoTo ErrH
    CurrentDb.Execute s, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "Bd SQL: " & Err.Number & " - " & Err.Description & vbCrLf & Left$(s, 1024), vbExclamation, "ExecSQL"
End Sub

Private Function TableExists(ByVal name As String) As Boolean
    On Error Resume Next
    TableExists = Not CurrentDb.TableDefs(name) Is Nothing
End Function

Private Sub DropTableIfExists(ByVal name As String)
    On Error Resume Next
    CurrentDb.TableDefs.Delete name
    On Error GoTo 0
End Sub

Private Sub CreateTableFromSpec(ByVal tableName As String, ByVal fieldsSpec As String, _
                                 ByVal pkName As String, ByVal pkFields As String)
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim dbTbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim fldSpec As Variant
    Dim parts As Variant
    Dim fldName As String
    Dim fldType As Integer
    Dim sizeVal As Long
    Dim flags As String
    Dim isAuto As Boolean
    Dim idx As DAO.Index
    Dim pkField As Variant

    On Error GoTo ErrHandler

    Set db = CurrentDb
    Set tbl = db.CreateTableDef(tableName)

    For Each fldSpec In Split(fieldsSpec, ";")
        fldSpec = Trim$(CStr(fldSpec))
        If Len(fldSpec) > 0 Then
            parts = Split(fldSpec, "|")
            fldName = Trim$(parts(0))
            If UBound(parts) >= 1 Then
                fldType = ResolveFieldType(parts(1), isAuto)
            Else
                Err.Raise vbObjectError + 701, "CreateTableFromSpec", "Brak typu dla pola: " & fldName
            End If
            sizeVal = 0
            If UBound(parts) >= 2 Then
                sizeVal = Val(parts(2))
            End If
            flags = ""
            If UBound(parts) >= 3 Then
                flags = Trim$(parts(3))
            End If

            Set fld = tbl.CreateField(fldName, fldType)
            If fldType = DAO.DataTypeEnum.dbText And sizeVal > 0 Then
                fld.Size = sizeVal
            End If
            If isAuto Then
                fld.Attributes = fld.Attributes Or dbAutoIncrField
            End If
            If InStr(1, flags, "REQ", vbTextCompare) > 0 Then
                fld.Required = True
            End If
            If InStr(1, flags, "ALLOWZERO", vbTextCompare) > 0 Then
                On Error Resume Next
                fld.AllowZeroLength = True
                On Error GoTo ErrHandler
            End If
            tbl.Fields.Append fld
        End If
    Next fldSpec

    db.TableDefs.Append tbl
    db.TableDefs.Refresh

    If Len(Trim$(pkFields)) > 0 Then
        Set dbTbl = db.TableDefs(tableName)
        Set idx = dbTbl.CreateIndex(pkName)
        idx.Primary = True
        idx.Unique = True
        For Each pkField In Split(pkFields, ",")
            pkField = Trim$(CStr(pkField))
            If Len(pkField) > 0 Then
                idx.Fields.Append idx.CreateField(pkField)
            End If
        Next pkField
        dbTbl.Indexes.Append idx
        dbTbl.Indexes.Refresh
    End If

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ResolveFieldType(ByVal token As String, ByRef isAuto As Boolean) As Integer
    Dim clean As String
    clean = UCase$(Trim$(token))
    isAuto = False
    Select Case clean
        Case "AUTONUM", "AUTOINCREMENT", "AUTONUMBER"
            ResolveFieldType = DAO.DataTypeEnum.dbLong
            isAuto = True
        Case "LONG"
            ResolveFieldType = DAO.DataTypeEnum.dbLong
        Case "TEXT"
            ResolveFieldType = DAO.DataTypeEnum.dbText
        Case "LONGTEXT", "MEMO"
            ResolveFieldType = DAO.DataTypeEnum.dbMemo
        Case "DATE"
            ResolveFieldType = DAO.DataTypeEnum.dbDate
        Case "DOUBLE"
            ResolveFieldType = DAO.DataTypeEnum.dbDouble
        Case "SHORT", "INT", "INTEGER"
            ResolveFieldType = DAO.DataTypeEnum.dbInteger
        Case "YESNO", "BOOLEAN"
            ResolveFieldType = DAO.DataTypeEnum.dbBoolean
        Case "ATTACHMENT"
            ResolveFieldType = DAO.DataTypeEnum.dbAttachment
        Case Else
            Err.Raise vbObjectError + 702, "ResolveFieldType", "Nieznany typ pola: " & token
    End Select
End Function


Private Sub EnsureFieldDefaults()
       'Miejsce na ewentualne ustawienia domylnych wartoci przez ALTER TABLE
End Sub

'==================== Public API ====================
Public Sub UtworzTabele()
    On Error GoTo ErrH

    '--- Usuwanie (idempotencja) – tylko zale¿ne, od koñca grafu ---
    Dim tbls As String
	Dim t As Variant
		
    tbls = "IntegracjeODBC|Zalaczniki|ChainOfCustody|Probki|Personel|SprzetPomiarowy|Laboratoria|" & _
           "NCR|OczekiwaneBadanie|PlanPoboru|Partie|" & _
           "WynikiBadania|UziarnienieWynik|Badania|RodzajeBadania|" & _
           "UziarnienieWymaganie|Sita|WymaganiaParametru|ParametryJakosci|" & _
           "Lokalizacje|Obiekty|" & _
           "ReceptyMieszanek|WarstwaMaterial|Materialy|ObiektWarstwaSpec|WarstwaSpecDefault|Warstwy|Specyfikacje|LogZdarzen"

    For Each t In Split(tbls, "|")
        DropTableIfExists CStr(t)

 Next t

    '--- Tworzenie tabel sownikowych/konfiguracyjnych ---
    CreateTableFromSpec "Specyfikacje", _
        "SpecyfikacjaID|AUTONUM;" & _
        "Kod|TEXT|50|REQ;" & _
        "Nazwa|TEXT|255|REQ;" & _
        "Wersja|TEXT|20;" & _
        "DataSpec|DATE;" & _
        "Uwagi|LONGTEXT", _
        "PK_Spec", "SpecyfikacjaID"

    CreateTableFromSpec "Warstwy", _
        "WarstwaID|AUTONUM;" & _
        "KodWarstwy|TEXT|20|REQ;" & _
        "NazwaWarstwy|TEXT|100|REQ;" & _
        "TypWarstwy|TEXT|30|REQ;" & _
        "MaterialDomyslny|LONG;" & _
        "Uwagi|LONGTEXT", _
        "PK_Warstwy", "WarstwaID"

    CreateTableFromSpec "WarstwaSpecDefault", _
        "WSDID|AUTONUM;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ", _
        "PK_WSD", "WSDID"

    CreateTableFromSpec "ObiektWarstwaSpec", _
        "OWSID|AUTONUM;" & _
        "ObiektID|LONG|0|REQ;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ", _
        "PK_OWS", "OWSID"

    CreateTableFromSpec "Materialy", _
        "MaterialID|AUTONUM;" & _
        "NazwaMaterialu|TEXT|100|REQ;" & _
        "TypMaterialu|TEXT|50;" & _
        "Producent|TEXT|100;" & _
        "Zrodlo|TEXT|100;" & _
        "Uwagi|LONGTEXT", _
        "PK_Mat", "MaterialID"

    CreateTableFromSpec "WarstwaMaterial", _
        "WMID|AUTONUM;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "MaterialID|LONG|0|REQ", _
        "PK_WM", "WMID"

    CreateTableFromSpec "ReceptyMieszanek", _
        "ReceptaID|AUTONUM;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "MaterialID|LONG;" & _
        "KodRecepty|TEXT|50|REQ;" & _
        "Opis|TEXT|255;" & _
        "DataUtw|DATE", _
        "PK_Rec", "ReceptaID"

    '--- Kontrakt i lokalizacja ---
        CreateTableFromSpec "Obiekty", _
        "ObiektID|AUTONUM;" & _
        "KodObiektu|TEXT|50|REQ;" & _
        "NazwaObiektu|TEXT|255|REQ;" & _
        "KM_Poczatek_m|LONG;" & _
        "KM_Koniec_m|LONG;" & _
        "Inwestor|TEXT|100;" & _
        "Kontrakt|TEXT|100;" & _
        "Uwagi|LONGTEXT", _
        "PK_Ob", "ObiektID"

    CreateTableFromSpec "Lokalizacje", _
        "LokalizacjaID|AUTONUM;" & _
        "ObiektID|LONG|0|REQ;" & _
        "KM_Start_m|LONG|0|REQ;" & _
        "KM_End_m|LONG|0|REQ;" & _
        "Pas|TEXT|20;" & _
        "Szerokosc_m|DOUBLE;" & _
        "Uwagi|LONGTEXT", _
        "PK_Lok", "LokalizacjaID"

    '--- Parametry i wymagania ---
    CreateTableFromSpec "ParametryJakosci", _
        "ParametrID|AUTONUM;" & _
        "NazwaParametru|TEXT|100|REQ;" & _
        "SymbolParametru|TEXT|20|REQ;" & _
        "Jednostka|TEXT|20;" & _
        "Opis|LONGTEXT", _
        "PK_Param", "ParametrID"

    CreateTableFromSpec "WymaganiaParametru", _
        "WymaganieID|AUTONUM;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "ParametrID|LONG|0|REQ;" & _
        "TypKryterium|TEXT|10|REQ;" & _
        "MinWartosc|DOUBLE;" & _
        "MaxWartosc|DOUBLE;" & _
        "Nominal|DOUBLE;" & _
        "TolMinus|DOUBLE;" & _
        "TolPlus|DOUBLE;" & _
        "NormaMetoda|TEXT|100;" & _
        "Uwagi|LONGTEXT", _
        "PK_WymPar", "WymaganieID"

    '--- Uziarnienie ---
     CreateTableFromSpec "Sita", _
        "SitoID|AUTONUM;" & _
        "Rozmiar_mm|DOUBLE|0|REQ;" & _
        "Opis|TEXT|50;" & _
        "Kolejnosc|SHORT|0|REQ", _
        "PK_Sita", "SitoID"

    CreateTableFromSpec "UziarnienieWymaganie", _
        "UziWymagID|AUTONUM;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SitoID|LONG|0|REQ;" & _
        "MinProc|DOUBLE;" & _
        "MaxProc|DOUBLE;" & _
        "NormaID|LONG;" & _
        "Uwagi|LONGTEXT", _
        "PK_UWym", "UziWymagID"

    CreateTableFromSpec "UziarnienieWynik", _
        "UziWynikID|AUTONUM;" & _
        "BadanieID|LONG|0|REQ;" & _
        "SitoID|LONG|0|REQ;" & _
        "ProcPrzechodzenia|DOUBLE;" & _
        "Uwagi|LONGTEXT", _
        "PK_UWyn", "UziWynikID"

    '--- Badania ---
    CreateTableFromSpec "RodzajeBadania", _
        "RodzajBadaniaID|AUTONUM;" & _
        "Nazwa|TEXT|100|REQ;" & _
        "NormaID|LONG", _
        "PK_Rodz", "RodzajBadaniaID"

    CreateTableFromSpec "Partie", _
        "PartiaID|AUTONUM;" & _
        "ObiektID|LONG|0|REQ;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "KM_Start_m|LONG|0|REQ;" & _
        "KM_End_m|LONG|0|REQ;" & _
        "DataWbudowania|DATE;" & _
        "Ilosc|DOUBLE;" & _
        "Jednostka|TEXT|10;" & _
        "Wykonawca|TEXT|100;" & _
        "[Status]|TEXT|30;" & _
        "Uwagi|LONGTEXT;" & _
        "DataUtworzenia|DATE;" & _
        "Utworzyl|TEXT|50", _
        "PK_Partie", "PartiaID"

    CreateTableFromSpec "NCR", _
        "NCRID|AUTONUM;" & _
        "PartiaID|LONG;" & _
        "BadanieID|LONG;" & _
        "DataZgloszenia|DATE;" & _
        "Klasyfikacja|TEXT|50;" & _
        "Opis|LONGTEXT;" & _
        "Przyczyna|LONGTEXT;" & _
        "Dzialania|LONGTEXT;" & _
        "Odpowiedzialny|TEXT|100;" & _
        "Termin|DATE;" & _
        "[Status]|TEXT|30;" & _
        "DataZamkniecia|DATE;" & _
        "Uwagi|LONGTEXT", _
        "PK_NCR", "NCRID"

    CreateTableFromSpec "PlanPoboru", _
        "PlanID|AUTONUM;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "RodzajBadaniaID|LONG;" & _
        "ParametrID|LONG;" & _
        "Interwal|DOUBLE;" & _
        "Jednostka|TEXT|10|REQ;" & _
        "MinimalnaLiczba|SHORT|0|REQ;" & _
        "Aktywne|YESNO;" & _
        "Uwagi|LONGTEXT", _
        "PK_Plan", "PlanID"

    CreateTableFromSpec "OczekiwaneBadanie", _
        "ExpID|AUTONUM;" & _
        "PartiaID|LONG|0|REQ;" & _
        "PlanID|LONG|0|REQ;" & _
        "RodzajBadaniaID|LONG;" & _
        "ParametrID|LONG;" & _
        "Termin|DATE;" & _
        "[Status]|TEXT|20;" & _
        "BadanieID|LONG;" & _
        "KM_m|LONG;" & _
        "Uwagi|LONGTEXT", _
        "PK_Exp", "ExpID"

    CreateTableFromSpec "Badania", _
        "BadanieID|AUTONUM;" & _
        "ObiektID|LONG|0|REQ;" & _
        "LokalizacjaID|LONG|0|REQ;" & _
        "WarstwaID|LONG|0|REQ;" & _
        "SpecyfikacjaID|LONG|0|REQ;" & _
        "MaterialID|LONG;" & _
        "ReceptaID|LONG;" & _
        "PartiaID|LONG;" & _
        "RodzajBadaniaID|LONG|0|REQ;" & _
        "DataBadania|DATE;" & _
        "ProtokolNr|TEXT|50;" & _
        "ProbkaKod|TEXT|50;" & _
        "MiejscePobrania|TEXT|100;" & _
        "DataPobrania|DATE;" & _
        "Uwagi|LONGTEXT;" & _
        "DataWprowadzenia|DATE;" & _
        "Wprowadzil|TEXT|50", _
        "PK_Bad", "BadanieID"

    CreateTableFromSpec "WynikiBadania", _
        "WynikID|AUTONUM;" & _
        "BadanieID|LONG|0|REQ;" & _
        "ParametrID|LONG|0|REQ;" & _
        "NrProbki|SHORT;" & _
        "Wartosc|DOUBLE;" & _
        "Jednostka|TEXT|20;" & _
        "Uwagi|LONGTEXT", _
        "PK_Wyn", "WynikID"

    '--- Priorytet 2 ---
    CreateTableFromSpec "SprzetPomiarowy", _
        "SprzetID|AUTONUM;" & _
        "Nazwa|TEXT|100|REQ;" & _
        "Typ|TEXT|50;" & _
        "NrSeryjny|TEXT|50;" & _
        "LaboratoriumID|LONG|0|REQ;" & _
        "DataKalibracji|DATE;" & _
        "DataWaznosci|DATE;" & _
        "Dokument|ATTACHMENT;" & _
        "Sciezka|TEXT|255;" & _
        "Uwagi|LONGTEXT", _
        "PK_Sprz", "SprzetID"

    CreateTableFromSpec "Personel", _
        "OsobaID|AUTONUM;" & _
        "ImieNazwisko|TEXT|100|REQ;" & _
        "Rola|TEXT|50;" & _
        "Uprawnienia|TEXT|100;" & _
        "DataWaznosci|DATE;" & _
        "LaboratoriumID|LONG|0|REQ;" & _
        "Uwagi|LONGTEXT", _
        "PK_Os", "OsobaID"

    CreateTableFromSpec "Probki", _
        "ProbkaID|AUTONUM;" & _
        "BadanieID|LONG|0|REQ;" & _
        "KodQR|TEXT|100;" & _
        "DataPobrania|DATE;" & _
        "Miejsce|TEXT|100;" & _
        "Warunki|TEXT|100;" & _
        "Przechowywanie|TEXT|100;" & _
        "Uwagi|LONGTEXT", _
        "PK_Prob", "ProbkaID"

    CreateTableFromSpec "ChainOfCustody", _
        "CoCID|AUTONUM;" & _
        "ProbkaID|LONG|0|REQ;" & _
        "OdKogo|TEXT|100;" & _
        "DoKogo|TEXT|100;" & _
        "Data|DATE;" & _
        "Adnotacje|LONGTEXT", _
        "PK_CoC", "CoCID"

    CreateTableFromSpec "Laboratoria", _
        "LaboratoriumID|AUTONUM;" & _
        "Nazwa|TEXT|100|REQ;" & _
        "NrAkredytacji|TEXT|50;" & _
        "Kontakt|TEXT|100;" & _
        "Uwagi|LONGTEXT", _
        "PK_Lab", "LaboratoriumID"

    '--- Priorytet 3 ---
    CreateTableFromSpec "LogZdarzen", _
        "LogID|AUTONUM;" & _
        "Encja|TEXT|50;" & _
        "EncjaID|LONG;" & _
        "Akcja|TEXT|50;" & _
        "Uzytkownik|TEXT|50;" & _
        "DataCzas|DATE;" & _
        "Szczegoly|LONGTEXT", _
        "PK_Log", "LogID"

    CreateTableFromSpec "Zalaczniki", _
        "FileID|AUTONUM;" & _
        "Encja|TEXT|20|REQ;" & _
        "EncjaID|LONG|0|REQ;" & _
        "Typ|TEXT|20;" & _
        "Opis|TEXT|100;" & _
        "Plik|ATTACHMENT;" & _
        "Sciezka|TEXT|255", _
        "PK_File", "FileID"

    CreateTableFromSpec "IntegracjeODBC", _
        "IntegracjaID|AUTONUM;" & _
        "Nazwa|TEXT|100|REQ;" & _
        "DSN|TEXT|100;" & _
        "TabelaDocelowa|TEXT|100;" & _
        "Mapowanie|LONGTEXT;" & _
        "Aktywne|YESNO;" & _
        "Uwagi|LONGTEXT", _
        "PK_Int", "IntegracjaID"

    EnsureFieldDefaults
    MsgBox "M01: Tabele utworzone.", vbInformation

    Exit Sub
ErrH:
    MsgBox "M01.UtworzTabele  b³¹d " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub Start_Build()
    On Error GoTo ErrH
    'Kolejno: M01M02M03M04M05M06M08M09M10, potem SelfTest
    Call UtworzTabele
    Call M02_IndeksyOgraniczenia.UtworzIndeksyIRelacje
    Call M03_DaneStartowe.ZaladujDaneStartowe
    Call M04_Kwerendy.UtworzKwerendy
    Call M05_Formularze.UtworzFormularze
    Call M06_Raporty.UtworzRaporty
    Call M08_PartieNCR.UtworzFormularze_PartieNCR
    Call M09_PrzykladoweWymaganiaParametrow.Przyklad_Wymagania_20
    Call M10_PrzykladoweWynikiBadan.Przyklad_Wyniki_20

    Dim ok As Boolean
    ok = M07_LogikaOceny.SelfTest_Kompletnosc()
    If ok Then
        MsgBox "SelfTest OK  kompletnoœæ potwierdzona.", vbInformation
    Else
        MsgBox "SelfTest wykry³ braki sprawdŸ komunikaty.", vbExclamation
    End If
    Exit Sub
ErrH:
    MsgBox "Start_Build  bd " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
