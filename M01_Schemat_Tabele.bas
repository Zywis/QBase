Attribute VB_Name = "M01_Schemat_Tabele"
Option Compare Database
Option Explicit

' M01: Schemat tabel – CREATE TABLE (bez FK), narzêdzia DDL i Start_Build

'==================== Narzêdzia ====================
Private Sub ExecSQL(ByVal s As String)
    On Error GoTo ErrH
    CurrentDb.Execute s, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d SQL: " & Err.Number & " - " & Err.Description & vbCrLf & Left$(s, 1024), vbExclamation, "ExecSQL"
End Sub

Private Function TableExists(ByVal name As String) As Boolean
    On Error Resume Next
    TableExists = Not CurrentDb.TableDefs(name) Is Nothing
End Function

Private Sub DropTableIfExists(ByVal name As String)
    On Error Resume Next
    If TableExists(name) Then CurrentDb.Execute "DROP TABLE [" & name & "]"
End Sub

Private Sub EnsureFieldDefaults()
    'Miejsce na ewentualne ALTER TABLE SET DEFAULT jeœli konieczne
End Sub

'==================== Public API ====================
Public Sub UtworzTabele()
    On Error GoTo ErrH

    '--- Usuwanie (idempotencja) – tylko zale¿ne, od koñca grafu ---
    Dim tbls As String
    tbls = "IntegracjeODBC|Zalaczniki|ChainOfCustody|Probki|Personel|SprzetPomiarowy|Laboratoria|" & _
           "NCR|OczekiwaneBadanie|PlanPoboru|Partie|" & _
           "WynikiBadania|UziarnienieWynik|Badania|RodzajeBadania|" & _
           "UziarnienieWymaganie|Sita|WymaganiaParametru|ParametryJakosci|" & _
           "Lokalizacje|Obiekty|" & _
           "ReceptyMieszanek|WarstwaMaterial|Materialy|ObiektWarstwaSpec|WarstwaSpecDefault|Warstwy|Specyfikacje|LogZdarzen"
    Dim t As Variant
    For Each t In Split(tbls, "|")
        DropTableIfExists CStr(t)
    Next

    '--- Tworzenie tabel s³ownikowych/konfiguracyjnych ---
    ExecSQL "CREATE TABLE Specyfikacje (" & _
            " SpecyfikacjaID AUTOINCREMENT CONSTRAINT PK_Spec PRIMARY KEY," & _
            " Kod TEXT(50) NOT NULL," & _
            " Nazwa TEXT(255) NOT NULL," & _
            " Wersja TEXT(20)," & _
            " DataSpec DATE," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Warstwy (" & _
            " WarstwaID AUTOINCREMENT CONSTRAINT PK_Warstwy PRIMARY KEY," & _
            " KodWarstwy TEXT(20) NOT NULL," & _
            " NazwaWarstwy TEXT(100) NOT NULL," & _
            " TypWarstwy TEXT(30) NOT NULL," & _
            " MaterialDomyslny LONG," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE WarstwaSpecDefault (" & _
            " WSDID AUTOINCREMENT CONSTRAINT PK_WSD PRIMARY KEY," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL" & _
            ")"

    ExecSQL "CREATE TABLE ObiektWarstwaSpec (" & _
            " OWSID AUTOINCREMENT CONSTRAINT PK_OWS PRIMARY KEY," & _
            " ObiektID LONG NOT NULL," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL" & _
            ")"

    ExecSQL "CREATE TABLE Materialy (" & _
            " MaterialID AUTOINCREMENT CONSTRAINT PK_Mat PRIMARY KEY," & _
            " NazwaMaterialu TEXT(100) NOT NULL," & _
            " TypMaterialu TEXT(50)," & _
            " Producent TEXT(100)," & _
            " Zrodlo TEXT(100)," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE WarstwaMaterial (" & _
            " WMID AUTOINCREMENT CONSTRAINT PK_WM PRIMARY KEY," & _
            " WarstwaID LONG NOT NULL," & _
            " MaterialID LONG NOT NULL" & _
            ")"

    ExecSQL "CREATE TABLE ReceptyMieszanek (" & _
            " ReceptaID AUTOINCREMENT CONSTRAINT PK_Rec PRIMARY KEY," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " MaterialID LONG," & _
            " KodRecepty TEXT(50) NOT NULL," & _
            " Opis TEXT(255)," & _
            " DataUtw DATE" & _
            ")"

    '--- Kontrakt i lokalizacja ---
    ExecSQL "CREATE TABLE Obiekty (" & _
            " ObiektID AUTOINCREMENT CONSTRAINT PK_Ob PRIMARY KEY," & _
            " KodObiektu TEXT(50) NOT NULL," & _
            " NazwaObiektu TEXT(255) NOT NULL," & _
            " KM_Poczatek_m LONG," & _
            " KM_Koniec_m LONG," & _
            " Inwestor TEXT(100)," & _
            " Kontrakt TEXT(100)," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Lokalizacje (" & _
            " LokalizacjaID AUTOINCREMENT CONSTRAINT PK_Lok PRIMARY KEY," & _
            " ObiektID LONG NOT NULL," & _
            " KM_Start_m LONG NOT NULL," & _
            " KM_End_m LONG NOT NULL," & _
            " Pas TEXT(20)," & _
            " Szerokosc_m DOUBLE," & _
            " Uwagi LONGTEXT" & _
            ")"

    '--- Parametry i wymagania ---
    ExecSQL "CREATE TABLE ParametryJakosci (" & _
            " ParametrID AUTOINCREMENT CONSTRAINT PK_Param PRIMARY KEY," & _
            " NazwaParametru TEXT(100) NOT NULL," & _
            " SymbolParametru TEXT(20) NOT NULL," & _
            " Jednostka TEXT(20)," & _
            " Opis LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE WymaganiaParametru (" & _
            " WymaganieID AUTOINCREMENT CONSTRAINT PK_WymPar PRIMARY KEY," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " WarstwaID LONG NOT NULL," & _
            " ParametrID LONG NOT NULL," & _
            " TypKryterium TEXT(10) NOT NULL," & _
            " MinWartosc DOUBLE," & _
            " MaxWartosc DOUBLE," & _
            " Nominal DOUBLE," & _
            " TolMinus DOUBLE," & _
            " TolPlus DOUBLE," & _
            " NormaMetoda TEXT(100)," & _
            " Uwagi LONGTEXT" & _
            ")"

    '--- Uziarnienie ---
    ExecSQL "CREATE TABLE Sita (" & _
            " SitoID AUTOINCREMENT CONSTRAINT PK_Sita PRIMARY KEY," & _
            " Rozmiar_mm DOUBLE NOT NULL," & _
            " Opis TEXT(50)," & _
            " Kolejnosc SHORT NOT NULL" & _
            ")"

    ExecSQL "CREATE TABLE UziarnienieWymaganie (" & _
            " UziWymagID AUTOINCREMENT CONSTRAINT PK_UWym PRIMARY KEY," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " WarstwaID LONG NOT NULL," & _
            " SitoID LONG NOT NULL," & _
            " MinProc DOUBLE," & _
            " MaxProc DOUBLE," & _
            " NormaID LONG," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE UziarnienieWynik (" & _
            " UziWynikID AUTOINCREMENT CONSTRAINT PK_UWyn PRIMARY KEY," & _
            " BadanieID LONG NOT NULL," & _
            " SitoID LONG NOT NULL," & _
            " ProcPrzechodzenia DOUBLE," & _
            " Uwagi LONGTEXT" & _
            ")"

    '--- Badania ---
    ExecSQL "CREATE TABLE RodzajeBadania (" & _
            " RodzajBadaniaID AUTOINCREMENT CONSTRAINT PK_Rodz PRIMARY KEY," & _
            " Nazwa TEXT(100) NOT NULL," & _
            " NormaID LONG" & _
            ")"

    ExecSQL "CREATE TABLE Partie (" & _
            " PartiaID AUTOINCREMENT CONSTRAINT PK_Partie PRIMARY KEY," & _
            " ObiektID LONG NOT NULL," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " KM_Start_m LONG NOT NULL," & _
            " KM_End_m LONG NOT NULL," & _
            " DataWbudowania DATE," & _
            " Ilosc DOUBLE," & _
            " Jednostka TEXT(10)," & _
            " Wykonawca TEXT(100)," & _
            " Status TEXT(30)," & _
            " Uwagi LONGTEXT," & _
            " DataUtworzenia DATE," & _
            " Utworzyl TEXT(50)" & _
            ")"

    ExecSQL "CREATE TABLE PlanPoboru (" & _
            " PlanID AUTOINCREMENT CONSTRAINT PK_Plan PRIMARY KEY," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " RodzajBadaniaID LONG," & _
            " ParametrID LONG," & _
            " Interwal DOUBLE," & _
            " Jednostka TEXT(10) NOT NULL," & _
            " MinimalnaLiczba SHORT NOT NULL," & _
            " Aktywne YESNO," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE OczekiwaneBadanie (" & _
            " ExpID AUTOINCREMENT CONSTRAINT PK_Exp PRIMARY KEY," & _
            " PartiaID LONG NOT NULL," & _
            " PlanID LONG NOT NULL," & _
            " RodzajBadaniaID LONG," & _
            " ParametrID LONG," & _
            " Termin DATE," & _
            " Status TEXT(20)," & _
            " BadanieID LONG," & _
            " KM_m LONG," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Badania (" & _
            " BadanieID AUTOINCREMENT CONSTRAINT PK_Bad PRIMARY KEY," & _
            " ObiektID LONG NOT NULL," & _
            " LokalizacjaID LONG NOT NULL," & _
            " WarstwaID LONG NOT NULL," & _
            " SpecyfikacjaID LONG NOT NULL," & _
            " MaterialID LONG," & _
            " ReceptaID LONG," & _
            " PartiaID LONG," & _
            " RodzajBadaniaID LONG NOT NULL," & _
            " DataBadania DATE," & _
            " ProtokolNr TEXT(50)," & _
            " ProbkaKod TEXT(50)," & _
            " MiejscePobrania TEXT(100)," & _
            " DataPobrania DATE," & _
            " Uwagi LONGTEXT," & _
            " DataWprowadzenia DATE," & _
            " Wprowadzil TEXT(50)" & _
            ")"

    ExecSQL "CREATE TABLE WynikiBadania (" & _
            " WynikID AUTOINCREMENT CONSTRAINT PK_Wyn PRIMARY KEY," & _
            " BadanieID LONG NOT NULL," & _
            " ParametrID LONG NOT NULL," & _
            " NrProbki SHORT," & _
            " Wartosc DOUBLE," & _
            " Jednostka TEXT(20)," & _
            " Uwagi LONGTEXT" & _
            ")"

    '--- Priorytet 2 ---
    ExecSQL "CREATE TABLE SprzetPomiarowy (" & _
            " SprzetID AUTOINCREMENT CONSTRAINT PK_Sprz PRIMARY KEY," & _
            " Nazwa TEXT(100) NOT NULL," & _
            " Typ TEXT(50)," & _
            " NrSeryjny TEXT(50)," & _
            " LaboratoriumID LONG NOT NULL," & _
            " DataKalibracji DATE," & _
            " DataWaznosci DATE," & _
            " Dokument ATTACHMENT," & _
            " Sciezka TEXT(255)," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Personel (" & _
            " OsobaID AUTOINCREMENT CONSTRAINT PK_Os PRIMARY KEY," & _
            " ImieNazwisko TEXT(100) NOT NULL," & _
            " Rola TEXT(50)," & _
            " Uprawnienia TEXT(100)," & _
            " DataWaznosci DATE," & _
            " LaboratoriumID LONG NOT NULL," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Probki (" & _
            " ProbkaID AUTOINCREMENT CONSTRAINT PK_Prob PRIMARY KEY," & _
            " BadanieID LONG NOT NULL," & _
            " KodQR TEXT(100)," & _
            " DataPobrania DATE," & _
            " Miejsce TEXT(100)," & _
            " Warunki TEXT(100)," & _
            " Przechowywanie TEXT(100)," & _
            " Uwagi LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE ChainOfCustody (" & _
            " CoCID AUTOINCREMENT CONSTRAINT PK_CoC PRIMARY KEY," & _
            " ProbkaID LONG NOT NULL," & _
            " OdKogo TEXT(100)," & _
            " DoKogo TEXT(100)," & _
            " Data DATE," & _
            " Adnotacje LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Laboratoria (" & _
            " LaboratoriumID AUTOINCREMENT CONSTRAINT PK_Lab PRIMARY KEY," & _
            " Nazwa TEXT(100) NOT NULL," & _
            " NrAkredytacji TEXT(50)," & _
            " Kontakt TEXT(100)," & _
            " Uwagi LONGTEXT" & _
            ")"

    '--- Priorytet 3 ---
    ExecSQL "CREATE TABLE LogZdarzen (" & _
            " LogID AUTOINCREMENT CONSTRAINT PK_Log PRIMARY KEY," & _
            " Encja TEXT(50)," & _
            " EncjaID LONG," & _
            " Akcja TEXT(50)," & _
            " Uzytkownik TEXT(50)," & _
            " DataCzas DATE," & _
            " Szczegoly LONGTEXT" & _
            ")"

    ExecSQL "CREATE TABLE Zalaczniki (" & _
            " FileID AUTOINCREMENT CONSTRAINT PK_File PRIMARY KEY," & _
            " Encja TEXT(20) NOT NULL," & _
            " EncjaID LONG NOT NULL," & _
            " Typ TEXT(20)," & _
            " Opis TEXT(100)," & _
            " Plik ATTACHMENT," & _
            " Sciezka TEXT(255)" & _
            ")"

    ExecSQL "CREATE TABLE IntegracjeODBC (" & _
            " IntegracjaID AUTOINCREMENT CONSTRAINT PK_Int PRIMARY KEY," & _
            " Nazwa TEXT(100) NOT NULL," & _
            " DSN TEXT(100)," & _
            " TabelaDocelowa TEXT(100)," & _
            " Mapowanie LONGTEXT," & _
            " Aktywne YESNO," & _
            " Uwagi LONGTEXT" & _
            ")"

    EnsureFieldDefaults
    MsgBox "M01: Tabele utworzone.", vbInformation

    Exit Sub
ErrH:
    MsgBox "M01.UtworzTabele – b³¹d " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub Start_Build()
    On Error GoTo ErrH
    'Kolejnoœæ: M01›M02›M03›M04›M05›M06›M08›M09›M10, potem SelfTest
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
    ok = M07_LogikaOceny.SelfTest_Kompletnosc
    If ok Then
        MsgBox "SelfTest OK – kompletnoœæ potwierdzona.", vbInformation
    Else
        MsgBox "SelfTest wykry³ braki – sprawdŸ komunikaty.", vbExclamation
    End If
    Exit Sub
ErrH:
    MsgBox "Start_Build – b³¹d " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
