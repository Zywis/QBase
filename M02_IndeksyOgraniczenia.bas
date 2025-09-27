Attribute VB_Name = "M02_IndeksyOgraniczenia"
Option Compare Database
Option Explicit

' M02: Indeksy, unikalnoœci i relacje (FK) z ON UPDATE CASCADE; ON DELETE CASCADE ograniczone

Private Sub ExecSQL(ByVal s As String)
    On Error GoTo ErrH
    CurrentDb.Execute s, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d SQL: " & Err.Number & " - " & Err.Description & vbCrLf & Left$(s, 1024), vbExclamation, "ExecSQL(M02)"
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

    '================ Unikalnoœci (sekcja 6) ================
    'Najpierw usuñ potencjalne istniej¹ce
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

    'OczekiwaneBadanie: PK ju¿ jest, dodatkowe indeksy:
    DropIndexIfExists "OczekiwaneBadanie", "IX_Exp_Partia"
    ExecSQL "CREATE INDEX IX_Exp_Partia ON OczekiwaneBadanie(PartiaID)"
    DropIndexIfExists "OczekiwaneBadanie", "IX_Exp_Badanie"
    ExecSQL "CREATE INDEX IX_Exp_Badanie ON OczekiwaneBadanie(BadanieID)"

    'Dodatkowe indeksy wydajnoœci:
    DropIndexIfExists "Lokalizacje", "IX_Lok_KM"
    ExecSQL "CREATE INDEX IX_Lok_KM ON Lokalizacje(KM_Start_m, KM_End_m)"
    DropIndexIfExists "Sita", "IX_Sita_Kolejnosc"
    ExecSQL "CREATE INDEX IX_Sita_Kolejnosc ON Sita(Kolejnosc)"
    DropIndexIfExists "UziarnienieWymaganie", "IX_UWym_WarSpec"
    ExecSQL "CREATE INDEX IX_UWym_WarSpec ON UziarnienieWymaganie(SpecyfikacjaID, WarstwaID)"
    DropIndexIfExists "Badania", "IX_Badania_Data"
    ExecSQL "CREATE INDEX IX_Badania_Data ON Badania(DataBadania)"

    'Indeksy na wszystkich FK – utworz¹ siê po relacjach, ale jawnie dodajemy
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

    '================ Relacje (FK) ================
    'Uwaga: ON DELETE CASCADE tylko w uzgodnionych miejscach

    'S³owniki
    ExecSQL "ALTER TABLE WarstwaSpecDefault ADD CONSTRAINT FK_WSD_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE WarstwaSpecDefault ADD CONSTRAINT FK_WSD_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE ObiektWarstwaSpec ADD CONSTRAINT FK_OWS_Obiekt FOREIGN KEY (ObiektID) REFERENCES Obiekty(ObiektID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE ObiektWarstwaSpec ADD CONSTRAINT FK_OWS_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE ObiektWarstwaSpec ADD CONSTRAINT FK_OWS_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE WarstwaMaterial ADD CONSTRAINT FK_WM_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE WarstwaMaterial ADD CONSTRAINT FK_WM_Material FOREIGN KEY (MaterialID) REFERENCES Materialy(MaterialID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE ReceptyMieszanek ADD CONSTRAINT FK_Rec_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE ReceptyMieszanek ADD CONSTRAINT FK_Rec_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE ReceptyMieszanek ADD CONSTRAINT FK_Rec_Mat FOREIGN KEY (MaterialID) REFERENCES Materialy(MaterialID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE Lokalizacje ADD CONSTRAINT FK_Lok_Obiekt FOREIGN KEY (ObiektID) REFERENCES Obiekty(ObiektID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE WymaganiaParametru ADD CONSTRAINT FK_Wym_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE WymaganiaParametru ADD CONSTRAINT FK_Wym_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE WymaganiaParametru ADD CONSTRAINT FK_Wym_Param FOREIGN KEY (ParametrID) REFERENCES ParametryJakosci(ParametrID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE UziarnienieWymaganie ADD CONSTRAINT FK_UWym_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE UziarnienieWymaganie ADD CONSTRAINT FK_UWym_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE UziarnienieWymaganie ADD CONSTRAINT FK_UWym_Sito FOREIGN KEY (SitoID) REFERENCES Sita(SitoID) ON UPDATE CASCADE"

    'Partie/Plan/Oczekiwane
    ExecSQL "ALTER TABLE Partie ADD CONSTRAINT FK_Partie_Obiekt FOREIGN KEY (ObiektID) REFERENCES Obiekty(ObiektID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Partie ADD CONSTRAINT FK_Partie_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Partie ADD CONSTRAINT FK_Partie_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE PlanPoboru ADD CONSTRAINT FK_Plan_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE PlanPoboru ADD CONSTRAINT FK_Plan_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE PlanPoboru ADD CONSTRAINT FK_Plan_Rodzaj FOREIGN KEY (RodzajBadaniaID) REFERENCES RodzajeBadania(RodzajBadaniaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE PlanPoboru ADD CONSTRAINT FK_Plan_Param FOREIGN KEY (ParametrID) REFERENCES ParametryJakosci(ParametrID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE OczekiwaneBadanie ADD CONSTRAINT FK_Exp_Partia FOREIGN KEY (PartiaID) REFERENCES Partie(PartiaID) ON UPDATE CASCADE ON DELETE CASCADE"
    ExecSQL "ALTER TABLE OczekiwaneBadanie ADD CONSTRAINT FK_Exp_Plan FOREIGN KEY (PlanID) REFERENCES PlanPoboru(PlanID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE OczekiwaneBadanie ADD CONSTRAINT FK_Exp_Rodzaj FOREIGN KEY (RodzajBadaniaID) REFERENCES RodzajeBadania(RodzajBadaniaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE OczekiwaneBadanie ADD CONSTRAINT FK_Exp_Param FOREIGN KEY (ParametrID) REFERENCES ParametryJakosci(ParametrID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE OczekiwaneBadanie ADD CONSTRAINT FK_Exp_Badanie FOREIGN KEY (BadanieID) REFERENCES Badania(BadanieID) ON UPDATE CASCADE"

    'Badania / Wyniki / Uziarnienie
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Obiekt FOREIGN KEY (ObiektID) REFERENCES Obiekty(ObiektID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Lok FOREIGN KEY (LokalizacjaID) REFERENCES Lokalizacje(LokalizacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Warstwa FOREIGN KEY (WarstwaID) REFERENCES Warstwy(WarstwaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Spec FOREIGN KEY (SpecyfikacjaID) REFERENCES Specyfikacje(SpecyfikacjaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Mat FOREIGN KEY (MaterialID) REFERENCES Materialy(MaterialID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Rec FOREIGN KEY (ReceptaID) REFERENCES ReceptyMieszanek(ReceptaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Partia FOREIGN KEY (PartiaID) REFERENCES Partie(PartiaID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Badania ADD CONSTRAINT FK_Bad_Rodzaj FOREIGN KEY (RodzajBadaniaID) REFERENCES RodzajeBadania(RodzajBadaniaID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE WynikiBadania ADD CONSTRAINT FK_Wyn_Bad FOREIGN KEY (BadanieID) REFERENCES Badania(BadanieID) ON UPDATE CASCADE ON DELETE CASCADE"
    ExecSQL "ALTER TABLE WynikiBadania ADD CONSTRAINT FK_Wyn_Param FOREIGN KEY (ParametrID) REFERENCES ParametryJakosci(ParametrID) ON UPDATE CASCADE"

    ExecSQL "ALTER TABLE UziarnienieWynik ADD CONSTRAINT FK_UWyn_Bad FOREIGN KEY (BadanieID) REFERENCES Badania(BadanieID) ON UPDATE CASCADE ON DELETE CASCADE"
    ExecSQL "ALTER TABLE UziarnienieWynik ADD CONSTRAINT FK_UWyn_Sito FOREIGN KEY (SitoID) REFERENCES Sita(SitoID) ON UPDATE CASCADE"

    'Priorytet 2
    ExecSQL "ALTER TABLE SprzetPomiarowy ADD CONSTRAINT FK_Sprz_Lab FOREIGN KEY (LaboratoriumID) REFERENCES Laboratoria(LaboratoriumID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Personel ADD CONSTRAINT FK_Os_Lab FOREIGN KEY (LaboratoriumID) REFERENCES Laboratoria(LaboratoriumID) ON UPDATE CASCADE"
    ExecSQL "ALTER TABLE Probki ADD CONSTRAINT FK_Prob_Bad FOREIGN KEY (BadanieID) REFERENCES Badania(BadanieID) ON UPDATE CASCADE ON DELETE CASCADE"
    ExecSQL "ALTER TABLE ChainOfCustody ADD CONSTRAINT FK_CoC_Prob FOREIGN KEY (ProbkaID) REFERENCES Probki(ProbkaID) ON UPDATE CASCADE ON DELETE CASCADE"

    'Walidacja TypKryterium – w VBA (Access DDL CHECK nieobs³ugiwane)

    MsgBox "M02: Indeksy i relacje utworzone.", vbInformation
    Exit Sub
ErrH:
    MsgBox "M02.UtworzIndeksyIRelacje – b³¹d " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
