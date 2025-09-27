Attribute VB_Name = "M02_IndeksyOgraniczenia"
Option Compare Database
Option Explicit

' M02: Indeksy, unikalności i relacje (FK) z nazwami oraz ON DELETE CASCADE ograniczone

Private Sub ExecSQL(ByVal s As String)
    On Error GoTo ErrH
    CurrentDb.Execute s, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "Błąd SQL: " & Err.Number & " - " & Err.Description & vbCrLf & Left$(s, 1024), vbExclamation, "ExecSQL(M02)"
End Sub

Private Function IndexExists(tbl As String, idx As String) As Boolean
    On Error Resume Next
    Dim tdef As DAO.TableDef, i As DAO.Index
    Set tdef = CurrentDb.TableDefs(tbl)
    For Each i In tdef.Indexes
        If StrComp(i.Name, idx, vbTextCompare) = 0 Then IndexExists = True: Exit Function
    Next
End Function

Private Sub DropIndexIfExists(tbl As String, idx As String)
    On Error Resume Next
    If IndexExists(tbl, idx) Then CurrentDb.Execute "DROP INDEX [" & idx & "] ON [" & tbl & "]"
End Sub

Public Sub UtworzIndeksyIRelacje()
    On Error GoTo ErrH

    '================ Unikalności (sekcja 6) ================
    DropIndexIfExists "Specyfikacje", "UQ_Spec_Kod"
    ExecSQL "CREATE UNIQUE INDEX UQ_Spec_Kod ON Specyfikacje(Kod)"

    DropIndexIfExists "WarstwaSpecDefault", "UQ_WSD_Warstwa"
    ExecSQL "CREATE UNIQUE INDEX UQ_WSD_Warstwa ON WarstwaSpecDefault(WarstwaID)"

    DropIndexIfExists "WarstwaMaterial", "UQ_WM_WarstwaMaterial"
    ExecSQL "CREATE UNIQUE INDEX UQ_WM_WarstwaMaterial ON WarstwaMaterial(WarstwaID, MaterialID)"

    DropIndexIfExists "ReceptyMieszanek", "UQ_Rec_WarSpecKod"
    ExecSQL "CREATE UNIQUE INDEX UQ_Rec_WarSpecKod ON ReceptyMieszanek(WarstwaID, SpecyfikacjaID, KodRecepty)"

    DropIndexIfExists "WymaganiaParametru", "UQ_Wym_WarSpecParam"
    ExecSQL "CREATE UNIQUE INDEX UQ_Wym_WarSpecParam ON WymaganiaParametru(SpecyfikacjaID, WarstwaID, ParametrID)"

    DropIndexIfExists "UziarnienieWymaganie", "UQ_UWym_WarSpecSito"
    ExecSQL "CREATE UNIQUE INDEX UQ_UWym_WarSpecSito ON UziarnienieWymaganie(SpecyfikacjaID, WarstwaID, SitoID)"

    DropIndexIfExists "UziarnienieWynik", "UQ_UWyn_BadSito"
    ExecSQL "CREATE UNIQUE INDEX UQ_UWyn_BadSito ON UziarnienieWynik(BadanieID, SitoID)"

    DropIndexIfExists "WynikiBadania", "UQ_Wyn_BadParamProb"
    ExecSQL "CREATE UNIQUE INDEX UQ_Wyn_BadParamProb ON WynikiBadania(BadanieID, ParametrID, NrProbki)"

    DropIndexIfExists "OczekiwaneBadanie", "IX_Exp_Partia"
    ExecSQL "CREATE INDEX IX_Exp_Partia ON OczekiwaneBadanie(PartiaID)"
    DropIndexIfExists "OczekiwaneBadanie", "IX_Exp_Badanie"
    ExecSQL "CREATE INDEX IX_Exp_Badanie ON OczekiwaneBadanie(BadanieID)"

    DropIndexIfExists "Lokalizacje", "IX_Lok_KM"
    ExecSQL "CREATE INDEX IX_Lok_KM ON Lokalizacje(KM_Start_m, KM_End_m)"
    DropIndexIfExists "Sita", "IX_Sita_Kolejnosc"
    ExecSQL "CREATE INDEX IX_Sita_Kolejnosc ON Sita(Kolejnosc)"
    DropIndexIfExists "UziarnienieWymaganie", "IX_UWym_WarSpec"
    ExecSQL "CREATE INDEX IX_UWym_WarSpec ON UziarnienieWymaganie(SpecyfikacjaID, WarstwaID)"
    DropIndexIfExists "Badania", "IX_Badania_Data"
    ExecSQL "CREATE INDEX IX_Badania_Data ON Badania(DataBadania)"

    'Indeksy na wszystkich FK
    Dim fkIdx As String, fkList As String
    fkList = "WarstwaSpecDefault.WarstwaID|WarstwaSpecDefault.SpecyfikacjaID|" & _
             "ObiektWarstwaSpec.ObiektID|ObiektWarstwaSpec.WarstwaID|ObiektWarstwaSpec.SpecyfikacjaID|" & _
             "WarstwaMaterial.WarstwaID|WarstwaMaterial.MaterialID|" & _
             "ReceptyMieszanek.WarstwaID|ReceptyMieszanek.SpecyfikacjaID|ReceptyMieszanek.MaterialID|" & _
             "Lokalizacje.ObiektID|" & _
             "WymaganiaParametru.SpecyfikacjaID|WymaganiaParametru.WarstwaID|WymaganiaParametru.ParametrID|" & _
             "UziarnienieWymaganie.SpecyfikacjaID|UziarnienieWymaganie.WarstwaID|UziarnienieWymaganie.SitoID|" & _
             "UziarnienieWynik.BadanieID|UziarnienieWynik.SitoID|" & _
             "Partie.ObiektID|Partie.WarstwaID|Partie.SpecyfikacjaID|" & _
             "PlanPoboru.WarstwaID|PlanPoboru.SpecyfikacjaID|PlanPoboru.RodzajBadaniaID|PlanPoboru.ParametrID|" & _
             "OczekiwaneBadanie.PartiaID|OczekiwaneBadanie.PlanID|OczekiwaneBadanie.RodzajBadaniaID|OczekiwaneBadanie.ParametrID|OczekiwaneBadanie.BadanieID|" & _
             "Badania.ObiektID|Badania.LokalizacjaID|Badania.WarstwaID|Badania.SpecyfikacjaID|Badania.MaterialID|Badania.ReceptaID|Badania.PartiaID|Badania.RodzajBadaniaID|" & _
             "WynikiBadania.BadanieID|WynikiBadania.ParametrID|" & _
             "SprzetPomiarowy.LaboratoriumID|Personel.LaboratoriumID|" & _
             "Probki.BadanieID|ChainOfCustody.ProbkaID"
    Dim p As Variant, tbl As String, col As String
    For Each p In Split(fkList, "|")
        tbl = Split(CStr(p), ".")(0): col = Split(CStr(p), ".")(1)
        fkIdx = "IX_" & Replace(tbl, " ", "") & "_" & Replace(col, " ", "")
        DropIndexIfExists tbl, fkIdx
        ExecSQL "CREATE INDEX " & fkIdx & " ON [" & tbl & "](" & col & ")"
    Next

    '================ Relacje z nazwami (FK) przez DAO ================
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Set db = CurrentDb

    ' Helper: FK Add - relacja z nazwą, z CASCADE, bez CASCADE
    ' relName, tbl, col, refTbl, refCol, deleteCascade
    AddFK "FK_WSD_Warstwa", "WarstwaSpecDefault", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_WSD_Spec", "WarstwaSpecDefault", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False

    AddFK "FK_OWS_Obiekt", "ObiektWarstwaSpec", "ObiektID", "Obiekty", "ObiektID", False
    AddFK "FK_OWS_Warstwa", "ObiektWarstwaSpec", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_OWS_Spec", "ObiektWarstwaSpec", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False

    AddFK "FK_WM_Warstwa", "WarstwaMaterial", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_WM_Material", "WarstwaMaterial", "MaterialID", "Materialy", "MaterialID", False

    AddFK "FK_Rec_Warstwa", "ReceptyMieszanek", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_Rec_Spec", "ReceptyMieszanek", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False
    AddFK "FK_Rec_Mat", "ReceptyMieszanek", "MaterialID", "Materialy", "MaterialID", False

    AddFK "FK_Lok_Obiekt", "Lokalizacje", "ObiektID", "Obiekty", "ObiektID", False

    AddFK "FK_Wym_Spec", "WymaganiaParametru", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False
    AddFK "FK_Wym_Warstwa", "WymaganiaParametru", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_Wym_Param", "WymaganiaParametru", "ParametrID", "ParametryJakosci", "ParametrID", False

    AddFK "FK_UWym_Spec", "UziarnienieWymaganie", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False
    AddFK "FK_UWym_Warstwa", "UziarnienieWymaganie", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_UWym_Sito", "UziarnienieWymaganie", "SitoID", "Sita", "SitoID", False

    AddFK "FK_Partie_Obiekt", "Partie", "ObiektID", "Obiekty", "ObiektID", False
    AddFK "FK_Partie_Warstwa", "Partie", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_Partie_Spec", "Partie", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False

    AddFK "FK_Plan_Warstwa", "PlanPoboru", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_Plan_Spec", "PlanPoboru", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False
    AddFK "FK_Plan_Rodzaj", "PlanPoboru", "RodzajBadaniaID", "RodzajeBadania", "RodzajBadaniaID", False
    AddFK "FK_Plan_Param", "PlanPoboru", "ParametrID", "ParametryJakosci", "ParametrID", False

    AddFK "FK_Exp_Partia", "OczekiwaneBadanie", "PartiaID", "Partie", "PartiaID", True
    AddFK "FK_Exp_Plan", "OczekiwaneBadanie", "PlanID", "PlanPoboru", "PlanID", False
    AddFK "FK_Exp_Rodzaj", "OczekiwaneBadanie", "RodzajBadaniaID", "RodzajeBadania", "RodzajBadaniaID", False
    AddFK "FK_Exp_Param", "OczekiwaneBadanie", "ParametrID", "ParametryJakosci", "ParametrID", False
    AddFK "FK_Exp_Badanie", "OczekiwaneBadanie", "BadanieID", "Badania", "BadanieID", False

    AddFK "FK_Bad_Obiekt", "Badania", "ObiektID", "Obiekty", "ObiektID", False
    AddFK "FK_Bad_Lok", "Badania", "LokalizacjaID", "Lokalizacje", "LokalizacjaID", False
    AddFK "FK_Bad_Warstwa", "Badania", "WarstwaID", "Warstwy", "WarstwaID", False
    AddFK "FK_Bad_Spec", "Badania", "SpecyfikacjaID", "Specyfikacje", "SpecyfikacjaID", False
    AddFK "FK_Bad_Mat", "Badania", "MaterialID", "Materialy", "MaterialID", False
    AddFK "FK_Bad_Rec", "Badania", "ReceptaID", "ReceptyMieszanek", "ReceptaID", False
    AddFK "FK_Bad_Partia", "Badania", "PartiaID", "Partie", "PartiaID", False
    AddFK "FK_Bad_Rodzaj", "Badania", "RodzajBadaniaID", "RodzajeBadania", "RodzajBadaniaID", False

    AddFK "FK_Wyn_Bad", "WynikiBadania", "BadanieID", "Badania", "BadanieID", True
    AddFK "FK_Wyn_Param", "WynikiBadania", "ParametrID", "ParametryJakosci", "ParametrID", False

    AddFK "FK_UWyn_Bad", "UziarnienieWynik", "BadanieID", "Badania", "BadanieID", True
    AddFK "FK_UWyn_Sito", "UziarnienieWynik", "SitoID", "Sita", "SitoID", False

    AddFK "FK_Sprz_Lab", "SprzetPomiarowy", "LaboratoriumID", "Laboratoria", "LaboratoriumID", False
    AddFK "FK_Os_Lab", "Personel", "LaboratoriumID", "Laboratoria", "LaboratoriumID", False
    AddFK "FK_Prob_Bad", "Probki", "BadanieID", "Badania", "BadanieID", True
    AddFK "FK_CoC_Prob", "ChainOfCustody", "ProbkaID", "Probki", "ProbkaID", True

    MsgBox "M02: Indeksy i relacje utworzone.", vbInformation
    Exit Sub
ErrH:
    MsgBox "M02.UtworzIndeksyIRelacje – błąd " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

' Helper: Tworzenie relacji FK przez DAO z nazwą i opcją ON DELETE CASCADE
Private Sub AddFK(relName As String, tbl As String, col As String, refTbl As String, refCol As String, deleteCascade As Boolean)
    On Error Resume Next
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Set db = CurrentDb
    ' Usuń starą relację jeśli istnieje
    For Each rel In db.Relations
        If rel.Name = relName Then db.Relations.Delete relName: Exit For
    Next
    Set rel = db.CreateRelation(relName, tbl, refTbl, IIf(deleteCascade, dbRelationDeleteCascade, 0))
    Set fld = rel.CreateField(col)
    fld.ForeignName = refCol
    rel.Fields.Append fld
    db.Relations.Append rel
    Set rel = Nothing
    Set fld = Nothing
End Sub