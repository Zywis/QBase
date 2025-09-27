Attribute VB_Name = "M03_DaneStartowe"
Option Compare Database
Option Explicit

'==========================================================
' Modu³: M03_DaneStartowe
' Wczytanie danych startowych (specyfikacje, sita, parametry,
' warstwy, materia³y, recepty, obiekt testowy)
'==========================================================

Private Sub ExecSQL(SQL As String)
    On Error GoTo ErrH
    CurrentDb.Execute SQL, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "B³¹d ExecSQL: " & Err.Description, vbCritical
End Sub

Public Sub ZaladujDaneStartowe()
    On Error GoTo ErrH
    
    '------------------------------------------------------
    ' Sita
    '------------------------------------------------------
    Dim sita As Variant
    sita = Split("63|31.5|22.4|16|11.2|8|4|2|1|0.5|0.25|0.125|0.063", "|")
    Dim i As Integer
    For i = LBound(sita) To UBound(sita)
      Dim rozmiarSita As Double
        rozmiarSita = CDbl(sita(i))
        ExecSQL "INSERT INTO Sita(Rozmiar_mm, Opis, Kolejnosc) VALUES(" & rozmiarSita & ", 'Sito " & rozmiarSita & " mm'," & (i + 1) & ")"
    Next i
    
    '------------------------------------------------------
    ' Rodzaje badañ
    '------------------------------------------------------
    Dim rodzaje As Variant
    rodzaje = Split("Uziarnienie|Gêstoœæ objêtoœciowa|Wilgotnoœæ optymalna|Zawartoœæ asfaltu|Zawartoœæ asfaltu rozpuszczalnego|" & _
                    "Gêstoœæ maksymalna teoretyczna (Gmm)|Wytrzyma³oœæ na rozci¹ganie poœrednie (ITS)|Odpornoœæ na dzia³anie wody (ITSR)|" & _
                    "WskaŸnik piaskowy (SE)|Los Angeles (LA)|Mrozoodpornoœæ – ubytek masy|Zawartoœæ py³ów <0,063 mm|" & _
                    "Modu³ odkszta³cenia E2|Modu³ odkszta³cenia E1|Stosunek E2/E1|Noœnoœæ CBR|WskaŸnik zagêszczenia Is|" & _
                    "Zawartoœæ cementu (stabilizacja)|Zawartoœæ lepiszcza odzyskana (Pb_sol)|Temperatura mieszanki przy wbudowaniu", "|")
    Dim r As Integer
    For r = LBound(rodzaje) To UBound(rodzaje)
        ExecSQL "INSERT INTO RodzajeBadania(Nazwa) VALUES('" & rodzaje(r) & "')"
    Next r
    
    '------------------------------------------------------
    ' Parametry jakoœci
    '------------------------------------------------------
    Dim parametry As Variant
    parametry = Split("Gestosc objetosciowa|?b|g/cm3|" & _
        "Wilgotnosc optymalna|wopt|%|" & _
        "Zawartosc asfaltu|Pb|%|" & _
        "Zawartosc asfaltu rozpuszczalnego|Pb_sol|%|" & _
        "Wolne przestrzenie w mieszance|Vv|%|" & _
        "Pusta przestrzen miedzy ziarnami|VMA|%|" & _
        "Wypelnienie wolnych przestrzeni lepiszczem|VFB|%|" & _
        "Gestosc maksymalna teoretyczna|Gmm|g/cm3|" & _
        "Wytrzymalosc ITS|ITS|MPa|" & _
        "Odpornosc ITSR|ITSR|%|" & _
        "Wskaznik piaskowy|SE|%|" & _
        "Los Angeles|LA|%|" & _
        "Ubytek po mrozoodpornosci|F_ubytek|%|" & _
        "Zawartosc pylow <0,063|P063|%|" & _
        "Zawartosc cementu|C_cem|%|" & _
        "Modul odksztalcenia E2|E2|MPa|" & _
        "Modul odksztalcenia E1|E1|MPa|" & _
        "Stosunek E2/E1|E2E1|-|" & _
        "Nosnosc CBR|CBR|%|" & _
        "Wskaznik zageszczenia|Is|%", "|")
    Dim p As Integer
    For p = LBound(parametry) To UBound(parametry) Step 3
        ExecSQL "INSERT INTO ParametryJakosci(NazwaParametru, SymbolParametru, Jednostka) VALUES('" & parametry(p) & "','" & parametry(p + 1) & "','" & parametry(p + 2) & "')"
    Next p
    
    '------------------------------------------------------
    ' Specyfikacje (pe³na lista z sekcji 13)
    '------------------------------------------------------
    Dim specyfikacje As Variant
    specyfikacje = Split("D-M-00.00.00;Wymagania Ogólne|" & _
        "D-01.00.00;ROBOTY PRZYGOTOWAWCZE|" & _
        "D-01.01.01;Odtworzenie trasy i punktów wysokoœciowych|" & _
        "D-01.02.01;Usuniêcie drzew, zagajników i krzewów|" & _
        "D-01.02.01A;Zabezpieczenie istniej¹cych drzew i krzewów na okres wykonywania robót|" & _
        "D-01.02.02;Zdjêcie warstwy humusu|" & _
        "D-01.02.03;Wyburzenie obiektów budowlanych|" & _
        "D-01.02.04;Rozbiórki elementów dróg i ulic|" & _
        "D-02.00.01;Roboty ziemne. Wymagania ogólne|" & _
        "D-02.01.01;Roboty ziemne. Wykonanie wykopów|" & _
        "D-02.01.01A;Platformy robocze dla ciê¿kiego sprzêtu budowlanego|" & _
        "D-02.01.01B;Wzmocnienie pod³o¿a gruntowego. Wymiana gruntów|" & _
        "D-02.01.01C;Wzmocnienie pod³o¿a gruntowego. Materace geosyntetyczne|" & _
        "D-02.01.01D;Wzmocnienie pod³o¿a gruntowego. Metoda drenów pionowych i nasypu przeci¹¿aj¹cego|" & _
        "D-02.01.01E;Wzmocnienie pod³o¿a gruntowego. Kolumny DSM|" & _
        "D-02.01.01F;Wzmocnienie pod³o¿a gruntowego. Metoda iniekcji strumieniowej Jet Grouting|" & _
        "D-02.01.01G;Wzmocnienie pod³o¿a gruntowego. Kolumny ¿wirowe|" & _
        "D-02.01.01H;Wzmocnienie pod³o¿a gruntowego. Kolumny betonowo-¿wirowe|" & _
        "D-02.01.01I;Wzmocnienie pod³o¿a gruntowego. Prefabrykowane pale ¿elbetowe|" & _
        "D-02.01.01J;Wzmocnienie pod³o¿a gruntowego. Pale wiercone typu CFA|" & _
        "D-02.03.01;Roboty ziemne. Wykonanie nasypów|" & _
        "D-03.00.00;ODWODNIENIE KORPUSU DROGOWEGO|" & _
        "D-03.01.01;Przepusty pod koron¹ drogi|" & _
        "D-03.03.01;S¹czki pod³u¿ne|" & _
        "D-04.00.00;PODBUDOWY|" & _
        "D-04.02.01;Warstwa odcinaj¹ca|" & _
        "D-04.02.02;Warstwa mrozoochronna/ods¹czaj¹ca|" & _
        "D-04.03.01;Oczyszczenie i skropienie warstw konstrukcyjnych|" & _
        "D-04.04.02;Podbudowa pomocnicza i zasadnicza z mieszanki niezwi¹zanej|" & _
        "D-04.05.00;Warstwa ulepszonego pod³o¿a z gruntu stabilizowanego spoiwem hydraulicznym lub wapnem|" & _
        "D-04.05.01;Podbudowa i warstwa mrozoochronna z mieszanki zwi¹zanej cementem|" & _
        "D-04.07.01;Warstwa podbudowy z AC|" & _
        "D-04.10.01;Podbudowa z mieszanki mineralno-cementowo-emulsyjnej|" & _
        "D-05.00.00;NAWIERZCHNIE|" & _
        "D-05.02.01;Nawierzchnia z kruszywa niezwi¹zanego|" & _
        "D-05.03.01;Nawierzchnia z kostki kamiennej|" & _
        "D-05.03.05A;Nawierzchnia z betonu asfaltowego warstwa wi¹¿¹ca|" & _
        "D-05.03.05B;Nawierzchnia z betonu asfaltowego warstwa œcieralna|" & _
        "D-05.03.05C;Warstwa podbudowy i wi¹¿¹ca z WMS|" & _
        "D-05.03.13;Nawierzchnia z mieszanki mastyksowo-grysowej (SMA)|" & _
        "D-05.03.23;Nawierzchnia i chodniki z kostki brukowej|" & _
        "D-06.00.00;ROBOTY WYKOÑCZENIOWE|" & _
        "D-06.01.01A;Umocnienie skarp, rowów i pasa dziel¹cego przez humusowanie i obsiew|" & _
        "D-06.01.01B;Umocnienie skarp i rowów elementami prefabrykowanymi|" & _
        "D-06.02.01;Przepusty pod zjazdami|" & _
        "D-06.03.01;Umocnienie poboczy|" & _
        "D-07.00.00;URZ¥DZENIA BEZPIECZEÑSTWA RUCHU|" & _
        "D-07.01.01;Oznakowanie poziome|" & _
        "D-07.02.01;Oznakowanie pionowe|" & _
        "D-07.02.02;S³upki prowadz¹ce i krawêdziowe oraz znaki kilometrowe i hektometrowe|" & _
        "D-07.05.01;Bariery ochronne|" & _
        "D-07.06.01;Ogrodzenia dróg|" & _
        "D-07.06.02;Urz¹dzenia zabezpieczaj¹ce ruch pieszy i rowerowy|" & _
        "D-07.08.01;Ekrany akustyczne. Ekrany przeciwolœnieniowe. Ekrany ekologiczne|" & _
        "D-07.09.01;Drogowy ekran przeciwolœnieniowy|" & _
        "D-08.00.00;ELEMENTY ULIC|" & _
        "D-08.01.01;Krawê¿niki betonowe|" & _
        "D-08.01.02;Krawê¿niki kamienne|" & _
        "D-08.03.01;Obrze¿a betonowe|" & _
        "D-08.05.01;Œcieki|" & _
        "D-09.00.00;ZIELEÑ DROGOWA|" & _
        "D-09.01.01;Zieleñ drogowa|" & _
        "M-11.00.00;Fundamentowanie|" & _
        "M-12.00.00;Zbrojenie|" & _
        "M-13.00.00;Beton|" & _
        "M-14.00.00;Konstrukcje stalowe|" & _
        "M-15.00.00;Izolacje i nawierzchnie|" & _
        "M-16.00.00;Odwodnienie|" & _
        "M-17.00.00;£o¿yska|" & _
        "M-18.00.00;Urz¹dzenia dylatacyjne|" & _
        "M-19.00.00;Elementy zabezpieczaj¹ce|" & _
        "M-20.01.00;Inne roboty mostowe. Roboty przyobiektowe|" & _
        "M-20.02.00;Inne roboty mostowe. Roboty na obiekcie|" & _
        "U-01.03.01;Przebudowa napowietrznych linii elektroenergetycznych|" & _
        "U-01.03.02;Przebudowa kablowych linii energetycznych|" & _
        "U-01.03.03;Przebudowa napowietrznych linii telekomunikacyjnych|" & _
        "U-01.03.04;Przebudowa kablowych linii telekomunikacyjnych|" & _
        "U-01.03.05;Przebudowa i budowa podziemnych sieci wodoci¹gowych|" & _
        "U-01.03.06;Przebudowa podziemnych sieci gazowych|" & _
        "U-01.03.07;Przebudowa urz¹dzeñ melioracyjnych|" & _
        "U-03.02.01;Kanalizacja deszczowa i sanitarna|" & _
        "U-03.05.01b;Zbiorniki retencyjne i retencyjno-infiltracyjne|" & _
        "U-05.03.01;Kanalizacja teletechniczna|" & _
        "U-07.07.01;Oœwietlenie drogowe|" & _
        "U-10.15.01;£¹cznoœæ drogi ekspresowej", "|")
    Dim s As Integer
    For s = LBound(specyfikacje) To UBound(specyfikacje)
        Dim kv As Variant
        kv = Split(specyfikacje(s), ";")
        ExecSQL "INSERT INTO Specyfikacje(Kod, Nazwa) VALUES('" & kv(0) & "','" & kv(1) & "')"
    Next s
    
    '------------------------------------------------------
    ' Warstwy + domyœlne spec
    '------------------------------------------------------
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('NASYP','Nasyp','Nasyp')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='NASYP' AND Kod='D-02.03.01'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('MROZOO','Warstwa mrozoochronna','Mrozo')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='MROZOO' AND Kod='D-04.02.02'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('KNN-PB','Podbudowa KNN','KNN')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='KNN-PB' AND Kod='D-04.04.02'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('STAB-C','Stabilizacja cementem','Stabilizacja')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='STAB-C' AND Kod='D-04.05.01'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('AC-PB','Podbudowa AC','MMA')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='AC-PB' AND Kod='D-04.07.01'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('AC16W','AC16W','MMA')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='AC16W' AND Kod='D-05.03.05A'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('AC11S','AC11S','MMA')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='AC11S' AND Kod='D-05.03.05B'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('SMA11','SMA11','MMA')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='SMA11' AND Kod='D-05.03.13'"
    
    ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('KNN-N','Nawierzchnia KNN','KNN')"
    ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje WHERE KodWarstwy='KNN-N' AND Kod='D-05.02.01'"
    
    '------------------------------------------------------
    ' Materia³y + przypisania
    '------------------------------------------------------
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('Grunt piaszczysty','Grunt')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy='NASYP' AND NazwaMaterialu='Grunt piaszczysty'"
    
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('Kruszywo 0/31,5','Kruszywo')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy IN('KNN-PB','KNN-N') AND NazwaMaterialu='Kruszywo 0/31,5'"
    
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('AC 22P 50/70','Asfalt')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy='AC-PB' AND NazwaMaterialu='AC 22P 50/70'"
    
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('AC 16W 50/70','Asfalt')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy='AC16W' AND NazwaMaterialu='AC 16W 50/70'"
    
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('AC 11S 50/70','Asfalt')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy='AC11S' AND NazwaMaterialu='AC 11S 50/70'"
    
    ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) VALUES('SMA 11 50/70','Asfalt')"
    ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) SELECT WarstwaID, MaterialID FROM Warstwy, Materialy WHERE KodWarstwy='SMA11' AND NazwaMaterialu='SMA 11 50/70'"
    
    '------------------------------------------------------
    ' Recepty mieszanek
    '------------------------------------------------------
    ExecSQL "INSERT INTO ReceptyMieszanek(WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) " & _
            "SELECT W.WarstwaID, S.SpecyfikacjaID, M.MaterialID, 'AC11S-50/70','Recepta AC11S', Date() " & _
            "FROM Warstwy W, Specyfikacje S, Materialy M WHERE W.KodWarstwy='AC11S' AND S.Kod='D-05.03.05B' AND M.NazwaMaterialu='AC 11S 50/70'"
    ExecSQL "INSERT INTO ReceptyMieszanek(WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) " & _
            "SELECT W.WarstwaID, S.SpecyfikacjaID, M.MaterialID, 'AC16W-50/70','Recepta AC16W', Date() " & _
            "FROM Warstwy W, Specyfikacje S, Materialy M WHERE W.KodWarstwy='AC16W' AND S.Kod='D-05.03.05A' AND M.NazwaMaterialu='AC 16W 50/70'"
    ExecSQL "INSERT INTO ReceptyMieszanek(WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) " & _
            "SELECT W.WarstwaID, S.SpecyfikacjaID, M.MaterialID, 'AC22P-50/70','Recepta AC22P', Date() " & _
            "FROM Warstwy W, Specyfikacje S, Materialy M WHERE W.KodWarstwy='AC-PB' AND S.Kod='D-04.07.01' AND M.NazwaMaterialu='AC 22P 50/70'"
    ExecSQL "INSERT INTO ReceptyMieszanek(WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) " & _
            "SELECT W.WarstwaID, S.SpecyfikacjaID, M.MaterialID, 'SMA11-50/70','Recepta SMA11', Date() " & _
            "FROM Warstwy W, Specyfikacje S, Materialy M WHERE W.KodWarstwy='SMA11' AND S.Kod='D-05.03.13' AND M.NazwaMaterialu='SMA 11 50/70'"
    
    '------------------------------------------------------
    ' Obiekt testowy i lokalizacja
    '------------------------------------------------------
    ExecSQL "INSERT INTO Obiekty(KodObiektu, NazwaObiektu, KM_Poczatek_m, KM_Koniec_m, Inwestor, Kontrakt) " & _
            "VALUES('OB-001','Odcinek testowy Sx',0,500,'Test','Test-K')"
    ExecSQL "INSERT INTO Lokalizacje(ObiektID, KM_Start_m, KM_End_m, Pas, Szerokosc_m) " & _
            "SELECT ObiektID, 0, 500, 'lewy', 3.5 FROM Obiekty WHERE KodObiektu='OB-001'"
    
    Exit Sub
ErrH:
    MsgBox "B³¹d ZaladujDaneStartowe: " & Err.Description, vbCritical
End Sub
