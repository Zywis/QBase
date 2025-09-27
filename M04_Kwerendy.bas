Attribute VB_Name = "M04_Kwerendy"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M04_Kwerendy
' Tworzenie kwerend (QueryDefs)
'==========================================================

Private Sub DropQueryIfExists(Nazwa As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete Nazwa
    On Error GoTo 0
End Sub

Public Sub UtworzKwerendy()
    On Error GoTo ErrH
    
    Dim q As DAO.QueryDef
    
    '------------------------------------------------------
    ' Q_Wyniki_Z_Ocena
    '------------------------------------------------------
    DropQueryIfExists "Q_Wyniki_Z_Ocena"
    Set q = CurrentDb.CreateQueryDef("Q_Wyniki_Z_Ocena", _
        "SELECT WB.WynikID, WB.BadanieID, WB.ParametrID, WB.Wartosc, Wym.TypKryterium, Wym.MinWartosc, Wym.MaxWartosc, Wym.Nominal, Wym.TolMinus, Wym.TolPlus, " & _
        "IIf(WB.Wartosc Is Null, Null, " & _
        "IIf(Wym.TypKryterium='Zakres', WB.Wartosc Between Wym.MinWartosc And Wym.MaxWartosc, " & _
        "IIf(Wym.TypKryterium='>=Min', WB.Wartosc>=Wym.MinWartosc, " & _
        "IIf(Wym.TypKryterium='<=Max', WB.Wartosc<=Wym.MaxWartosc, " & _
        "IIf(Wym.TypKryterium='+/-', WB.Wartosc Between (Wym.Nominal-Wym.TolMinus) And (Wym.Nominal+Wym.TolPlus), Null))))) AS Spelnia " & _
        "FROM WynikiBadania WB INNER JOIN WymaganiaParametru Wym ON WB.ParametrID=Wym.ParametrID AND WB.BadanieID IN (SELECT BadanieID FROM Badania WHERE SpecyfikacjaID=Wym.SpecyfikacjaID AND WarstwaID=Wym.WarstwaID)")
    
    '------------------------------------------------------
    ' Q_Niespelnienia_Param
    '------------------------------------------------------
    DropQueryIfExists "Q_Niespelnienia_Param"
    Set q = CurrentDb.CreateQueryDef("Q_Niespelnienia_Param", _
        "SELECT * FROM Q_Wyniki_Z_Ocena WHERE Spelnia=False OR Spelnia Is Null")
    
    '------------------------------------------------------
    ' Q_Uziarnienie_Ocena
    '------------------------------------------------------
    DropQueryIfExists "Q_Uziarnienie_Ocena"
    Set q = CurrentDb.CreateQueryDef("Q_Uziarnienie_Ocena", _
        "SELECT UW.UziWynikID, UW.BadanieID, UW.SitoID, UW.ProcPrzechodzenia, UWym.MinProc, UWym.MaxProc, " & _
        "IIf(UW.ProcPrzechodzenia Between UWym.MinProc And UWym.MaxProc, True, False) AS SpelniaSito " & _
        "FROM (UziarnienieWynik UW INNER JOIN Badania B ON UW.BadanieID=B.BadanieID) " & _
        "INNER JOIN UziarnienieWymaganie UWym ON B.SpecyfikacjaID=UWym.SpecyfikacjaID AND B.WarstwaID=UWym.WarstwaID AND UW.SitoID=UWym.SitoID")
    
    '------------------------------------------------------
    ' Q_Uziarnienie_Krzywa_Spelnia
    '------------------------------------------------------
    DropQueryIfExists "Q_Uziarnienie_Krzywa_Spelnia"
    Set q = CurrentDb.CreateQueryDef("Q_Uziarnienie_Krzywa_Spelnia", _
        "SELECT BadanieID, Min(SpelniaSito) AS KrzywaSpelnia FROM Q_Uziarnienie_Ocena GROUP BY BadanieID")
    
    '------------------------------------------------------
    ' Q_Uziarnienie_Crosstab
    '------------------------------------------------------
    DropQueryIfExists "Q_Uziarnienie_Crosstab"
    Set q = CurrentDb.CreateQueryDef("Q_Uziarnienie_Crosstab", _
        "TRANSFORM First(ProcPrzechodzenia) SELECT BadanieID FROM UziarnienieWynik GROUP BY BadanieID PIVOT SitoID")
    
    '------------------------------------------------------
    ' Q_Insert_Uziarnienie_Sita
    '------------------------------------------------------
    DropQueryIfExists "Q_Insert_Uziarnienie_Sita"
    Set q = CurrentDb.CreateQueryDef("Q_Insert_Uziarnienie_Sita", _
        "INSERT INTO UziarnienieWynik(BadanieID, SitoID) " & _
        "SELECT B.BadanieID, S.SitoID FROM Badania B, Sita S " & _
        "WHERE NOT EXISTS(SELECT 1 FROM UziarnienieWynik UW WHERE UW.BadanieID=B.BadanieID AND UW.SitoID=S.SitoID)")
    
    '------------------------------------------------------
    ' Q_Partie_Combo_RS
    '------------------------------------------------------
    DropQueryIfExists "Q_Partie_Combo_RS"
    Set q = CurrentDb.CreateQueryDef("Q_Partie_Combo_RS", _
        "SELECT PartiaID, ObiektID, WarstwaID, SpecyfikacjaID FROM Partie")
    
    '------------------------------------------------------
    ' Q_SF_Oczekiwane_RS
    '------------------------------------------------------
    DropQueryIfExists "Q_SF_Oczekiwane_RS"
    Set q = CurrentDb.CreateQueryDef("Q_SF_Oczekiwane_RS", _
        "SELECT * FROM OczekiwaneBadanie")
    
    '------------------------------------------------------
    ' Q_SF_BadaniaLotu_RS
    '------------------------------------------------------
    DropQueryIfExists "Q_SF_BadaniaLotu_RS"
    Set q = CurrentDb.CreateQueryDef("Q_SF_BadaniaLotu_RS", _
        "SELECT * FROM Badania WHERE PartiaID=Forms!F_Partia!PartiaID")
    
    '------------------------------------------------------
    ' Q_Loty_Pokrycie_Rodzaje, Q_Loty_Pokrycie_Parametry
    '------------------------------------------------------
    DropQueryIfExists "Q_Loty_Pokrycie_Rodzaje"
    Set q = CurrentDb.CreateQueryDef("Q_Loty_Pokrycie_Rodzaje", _
        "SELECT PartiaID, RodzajBadaniaID, Count(*) AS Ilosc FROM Badania GROUP BY PartiaID, RodzajBadaniaID")
    
    DropQueryIfExists "Q_Loty_Pokrycie_Parametry"
    Set q = CurrentDb.CreateQueryDef("Q_Loty_Pokrycie_Parametry", _
        "SELECT B.PartiaID, WB.ParametrID, Count(WB.WynikID) AS Ilosc FROM WynikiBadania WB INNER JOIN Badania B ON WB.BadanieID=B.BadanieID GROUP BY B.PartiaID, WB.ParametrID")
    
    DropQueryIfExists "Q_Loty_Pokrycie"
    Set q = CurrentDb.CreateQueryDef("Q_Loty_Pokrycie", _
        "SELECT PartiaID, 'Rodzaj' AS Typ, RodzajBadaniaID AS ID, Ilosc FROM Q_Loty_Pokrycie_Rodzaje " & _
        "UNION SELECT PartiaID, 'Parametr' AS Typ, ParametrID AS ID, Ilosc FROM Q_Loty_Pokrycie_Parametry")
    
    DropQueryIfExists "Q_Loty_Pokrycie_Summary"
    Set q = CurrentDb.CreateQueryDef("Q_Loty_Pokrycie_Summary", _
        "SELECT PartiaID, Typ, ID, Sum(Ilosc) AS Ilosc FROM Q_Loty_Pokrycie GROUP BY PartiaID, Typ, ID")
    
    '------------------------------------------------------
    ' Q_Badania_DoWykonania
    '------------------------------------------------------
    DropQueryIfExists "Q_Badania_DoWykonania"
    Set q = CurrentDb.CreateQueryDef("Q_Badania_DoWykonania", _
        "SELECT * FROM OczekiwaneBadanie WHERE Status In('DoWykonania','PoTermin')")
    
    Exit Sub
ErrH:
    MsgBox "B³¹d UtworzKwerendy: " & Err.Description, vbCritical
End Sub
