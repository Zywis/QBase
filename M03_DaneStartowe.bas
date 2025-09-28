' Version: 1.0
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

Private Sub WyczyscDane()
    On Error Resume Next
    ExecSQL "DELETE FROM Lokalizacje"
    ExecSQL "DELETE FROM Obiekty"
    ExecSQL "DELETE FROM ReceptyMieszanek"
    ExecSQL "DELETE FROM WarstwaMaterial"
    ExecSQL "DELETE FROM WarstwaSpecDefault"
    ExecSQL "DELETE FROM Materialy"
    ExecSQL "DELETE FROM Warstwy"
    ExecSQL "DELETE FROM Specyfikacje"
    ExecSQL "DELETE FROM ParametryJakosci"
    ExecSQL "DELETE FROM RodzajeBadania"
    ExecSQL "DELETE FROM Sita"
    On Error GoTo 0
End Sub

Public Sub ZaladujDaneStartowe()
    On Error GoTo ErrH
    
    ' Czyszczenie istniej¹cych danych (idempotentnoœæ)
    WyczyscDane
    
    '------------------------------------------------------
    ' Sita
    '------------------------------------------------------
    Dim i As Integer
    Dim rozmiary As Variant
    rozmiary = Array(0.063, 0.125, 0.25, 0.5, 1, 2, 4, 8, 11.2, 16, 22.4, 31.5)
    For i = LBound(rozmiary) To UBound(rozmiary)
        Dim rozmiarSita As Double
        rozmiarSita = rozmiary(i)
        ExecSQL "INSERT INTO Sita(Rozmiar_mm, Opis, Kolejnosc) VALUES(" & rozmiarSita & ", 'Sito " & rozmiarSita & " mm'," & (i + 1) & ")"
    Next i
    
    '------------------------------------------------------
    ' Rodzaje badañ
    '------------------------------------------------------
    Dim rodzaje As Variant
    Dim strRodzajeList As String
    strRodzajeList = "Uziarnienie|Gêstoœæ objêtoœciowa|Wilg...ymalna|Zawartoœæ asfaltu|Zawartoœæ asfaltu rozpuszczalnego|"
    strRodzajeList = strRodzajeList & "Gêstoœæ maksymalna teoretyczna (Gmm)|...i¹ganie poœrednie (ITS)|Odpornoœæ na dzia³anie wody (ITSR)|"
    strRodzajeList = strRodzajeList & "WskaŸnik piaskowy (SE)|Los Angeles (LA)|Mrozoodpornoœæ – ubytek masy|Zawartoœæ py³ów <0,063 mm|"
    strRodzajeList = strRodzajeList & "Modu³ odkszta³cenia E2|Modu³ odkszta³...nia E1|Stosunek E2/E1|Noœnoœæ CBR|WskaŸnik zagêszczenia Is|"
    strRodzajeList = strRodzajeList & "Zawartoœæ cementu (stabilizacja)|Zawa... odzyskana (Pb_sol)|Temperatura mieszanki przy wbudowaniu"
    rodzaje = Split(strRodzajeList, "|")
    Dim r As Integer
    For r = LBound(rodzaje) To UBound(rodzaje)
        ExecSQL "INSERT INTO RodzajeBadania(Nazwa) VALUES('" & rodzaje(r) & "')"
    Next r
    
    '------------------------------------------------------
    ' Parametry jakoœci
    '------------------------------------------------------
    Dim parametry As Variant
    Dim strParametryList As String
    strParametryList = "Gestosc objetosciowa|?b|g/cm3|"
    strParametryList = strParametryList & "Wilgotnosc optymalna|wopt|%|"
    strParametryList = strParametryList & "Zawartosc asfaltu|Pb|%|"
    strParametryList = strParametryList & "Zawartosc asfaltu rozpuszczalnego|Pb_sol|%|"
    strParametryList = strParametryList & "Wolne przestrzenie w mieszance|Vv|%|"
    strParametryList = strParametryList & "Pusta przestrzen miedzy ziarnami|VMA|%|"
    strParametryList = strParametryList & "Wypelnienie wolnych przestrzeni lepiszczem|VFB|%|"
    strParametryList = strParametryList & "Gestosc maksymalna teoretyczna|Gmm|g/cm3|"
    strParametryList = strParametryList & "Gestosc obj mieszanek mineralno-asfaltowych|?m|g/cm3|"
    strParametryList = strParametryList & "Wolne przestrzenie w ziarnach grubych|VCA|%|"
    strParametryList = strParametryList & "Wytrzymalosc na rozciaganie poœrednie|ITS|MPa|"
    strParametryList = strParametryList & "Odpornoœæ na dzia³anie wody|ITSR|%|"
    strParametryList = strParametryList & "Ubytek masy po 10 cyklach mrozowych|F10|%|"
    strParametryList = strParametryList & "Zawartosc pylow|P<063|%|"
    strParametryList = strParametryList & "Modul odksztalcenia E2|E2|MPa|"
    strParametryList = strParametryList & "Modul odksztalcenia E1|E1|MPa|"
    strParametryList = strParametryList & "Stosunek E2/E1|E2E1|-|"
    strParametryList = strParametryList & "Nosnosc CBR|CBR|%|"
    strParametryList = strParametryList & "Wskaznik zageszczenia|Is|%"
    parametry = Split(strParametryList, "|")
    Dim p As Integer
    For p = LBound(parametry) To UBound(parametry) Step 3
        ExecSQL "INSERT INTO ParametryJakosci(NazwaParametru, KodParametru, Jednostka) VALUES('" & parametry(p) & "','" & parametry(p + 1) & "','" & parametry(p + 2) & "')"
    Next p
    
    '------------------------------------------------------
    ' Specyfikacje (pe³na lista z sekcji 13)
    '------------------------------------------------------
    Private Sub WstawSpecyfikacje_PelnaLista()
    On Error GoTo ErrH

    ' Uwaga: ka¿da specyfikacja jako osobny INSERT (unikamy kontynuacji linii)
    Dim K As String, N As String

    ' Helper:
    Dim SubAdd As Object
    ' Nic – u¿yj lokalnej procedury:
    ' Insert: if not exists by Kod

    ' --- Wymagania Ogólne i D-01... ---
    AddSpec "D-M-00.00.00", "Wymagania Ogólne"
    AddSpec "D-01.00.00", "ROBOTY PRZYGOTOWAWCZE"
    AddSpec "D-01.01.01", "Odtworzenie trasy i punktów wysokoœciowych"
    AddSpec "D-01.02.01", "Usuniêcie drzew, zagajników i krzewów"
    AddSpec "D-01.02.01A", "Zabezpieczenie istniej¹cych drzew i krzewów na okres wykonywania robót"
    AddSpec "D-01.02.02", "Zdjêcie warstwy humusu"
    AddSpec "D-01.02.03", "Wyburzenie obiektów budowlanych"
    AddSpec "D-01.02.04", "Rozbiórki elementów dróg i ulic"

    ' --- D-02... Roboty ziemne ---
    AddSpec "D-02.00.01", "Roboty ziemne. Wymagania ogólne"
    AddSpec "D-02.01.01", "Roboty ziemne. Wykonanie wykopów"
    AddSpec "D-02.01.01A", "Platformy robocze dla ciê¿kiego sprzêtu budowlanego"
    AddSpec "D-02.01.01B", "Wzmocnienie pod³o¿a gruntowego. Wymiana gruntów"
    AddSpec "D-02.01.01C", "Wzmocnienie pod³o¿a gruntowego. Materace geosyntetyczne"
    AddSpec "D-02.01.01D", "Wzmocnienie pod³o¿a gruntowego. Metoda drenów pionowych i nasypu przeci¹¿aj¹cego"
    AddSpec "D-02.01.01E", "Wzmocnienie pod³o¿a gruntowego. Kolumny DSM"
    AddSpec "D-02.01.01F", "Wzmocnienie pod³o¿a gruntowego. Metoda iniekcji strumieniowej Jet Grouting"
    AddSpec "D-02.01.01G", "Wzmocnienie pod³o¿a gruntowego. Kolumny ¿wirowe"
    AddSpec "D-02.01.01H", "Wzmocnienie pod³o¿a gruntowego. Kolumny betonowo-¿wirowe"
    AddSpec "D-02.01.01I", "Wzmocnienie pod³o¿a gruntowego. Prefabrykowane pale ¿elbetowe"
    AddSpec "D-02.01.01J", "Wzmocnienie pod³o¿a gruntowego. Pale wiercone typu CFA"
    AddSpec "D-02.03.01", "Roboty ziemne. Wykonanie nasypów"

    ' --- D-03... Odwodnienie korpusu ---
    AddSpec "D-03.00.00", "ODWODNIENIE KORPUSU DROGOWEGO"
    AddSpec "D-03.01.01", "Przepusty pod koron¹ drogi"
    AddSpec "D-03.03.01", "S¹czki pod³u¿ne"

    ' --- D-04... Podbudowy ---
    AddSpec "D-04.00.00", "PODBUDOWY"
    AddSpec "D-04.02.01", "Warstwa odcinaj¹ca"
    AddSpec "D-04.02.02", "Warstwa mrozoochronna/ods¹czaj¹ca"
    AddSpec "D-04.03.01", "Oczyszczenie i skropienie warstw konstrukcyjnych"
    AddSpec "D-04.04.02", "Podbudowa pomocnicza i zasadnicza z mieszanki niezwi¹zanej"
    AddSpec "D-04.05.00", "Warstwa ulepszonego pod³o¿a z gruntu stabilizowanego spoiwem hydraulicznym lub wapnem"
    AddSpec "D-04.05.01", "Podbudowa i warstwa mrozoochronna z mieszanki zwi¹zanej cementem"
    AddSpec "D-04.07.01", "Warstwa podbudowy z AC"
    AddSpec "D-04.10.01", "Podbudowa z mieszanki mineralno-cementowo-emulsyjnej"

    ' --- D-05... Nawierzchnie ---
    AddSpec "D-05.00.00", "NAWIERZCHNIE"
    AddSpec "D-05.02.01", "Nawierzchnia z kruszywa niezwi¹zanego"
    AddSpec "D-05.03.01", "Nawierzchnia z kostki kamiennej"
    AddSpec "D-05.03.05A", "Nawierzchnia z betonu asfaltowego warstwa wi¹¿¹ca"
    AddSpec "D-05.03.05B", "Nawierzchnia z betonu asfaltowego warstwa œcieralna"
    AddSpec "D-05.03.05C", "Warstwa podbudowy i wi¹¿¹ca z WMS"
    AddSpec "D-05.03.13", "Nawierzchnia z mieszanki mastyksowo-grysowej (SMA)"
    AddSpec "D-05.03.23", "Nawierzchnia i chodniki z kostki brukowej"

    ' --- D-06... Roboty wykoñczeniowe ---
    AddSpec "D-06.00.00", "ROBOTY WYKOÑCZENIOWE"
    AddSpec "D-06.01.01A", "Umocnienie skarp, rowów i pasa dziel¹cego przez humusowanie i obsiew"
    AddSpec "D-06.01.01B", "Umocnienie skarp i rowów elementami prefabrykowanymi"
    AddSpec "D-06.02.01", "Przepusty pod zjazdami"
    AddSpec "D-06.03.01", "Umocnienie poboczy"

    ' --- D-07... Urz¹dzenia BRD ---
    AddSpec "D-07.00.00", "URZ¥DZENIA BEZPIECZEÑSTWA RUCHU"
    AddSpec "D-07.01.01", "Oznakowanie poziome"
    AddSpec "D-07.02.01", "Oznakowanie pionowe"
    AddSpec "D-07.02.02", "S³upki prowadz¹ce i krawêdziowe oraz znaki kilometrowe i hektometrowe"
    AddSpec "D-07.05.01", "Bariery ochronne"
    AddSpec "D-07.06.01", "Ogrodzenia dróg"
    AddSpec "D-07.06.02", "Urz¹dzenia zabezpieczaj¹ce ruch pieszy i rowerowy"
    AddSpec "D-07.08.01", "Ekrany akustyczne. Ekrany przeciwolœnieniowe. Ekrany ekologiczne"
    AddSpec "D-07.09.01", "Drogowy ekran przeciwolœnieniowy"

    ' --- D-08... Elementy ulic ---
    AddSpec "D-08.00.00", "ELEMENTY ULIC"
    AddSpec "D-08.01.01", "Krawê¿niki betonowe"
    AddSpec "D-08.01.02", "Krawê¿niki kamienne"
    AddSpec "D-08.03.01", "Obrze¿a betonowe"
    AddSpec "D-08.05.01", "Œcieki"

    ' --- D-09... Zieleñ drogowa ---
    AddSpec "D-09.00.00", "ZIELEÑ DROGOWA"
    AddSpec "D-09.01.01", "Zieleñ drogowa"

    ' --- M-11... M-20... (mostowe) ---
    AddSpec "M-11.00.00", "Fundamentowanie"
    AddSpec "M-12.00.00", "Zbrojenie"
    AddSpec "M-13.00.00", "Beton"
    AddSpec "M-14.00.00", "Konstrukcje stalowe"
    AddSpec "M-15.00.00", "Izolacje i nawierzchnie"
    AddSpec "M-16.00.00", "Odwodnienie"
    AddSpec "M-17.00.00", "£o¿yska"
    AddSpec "M-18.00.00", "Urz¹dzenia dylatacyjne"
    AddSpec "M-19.00.00", "Elementy zabezpieczaj¹ce"
    AddSpec "M-20.01.00", "Inne roboty mostowe. Roboty przyobiektowe"
    AddSpec "M-20.02.00", "Inne roboty mostowe. Roboty na obiekcie"

    ' --- U-01... U-10... (uzbrojenie) ---
    AddSpec "U-01.03.01", "Przebudowa napowietrznych linii elektroenergetycznych"
    AddSpec "U-01.03.02", "Przebudowa kablowych linii energetycznych"
    AddSpec "U-01.03.03", "Przebudowa napowietrznych linii telekomunikacyjnych"
    AddSpec "U-01.03.04", "Przebudowa kablowych linii telekomunikacyjnych"
    AddSpec "U-01.03.05", "Przebudowa i budowa podziemnych sieci wodoci¹gowych"
    AddSpec "U-01.03.06", "Przebudowa podziemnych sieci gazowych"
    AddSpec "U-01.03.07", "Przebudowa urz¹dzeñ melioracyjnych"
    AddSpec "U-03.02.01", "Kanalizacja deszczowa i sanitarna"
    AddSpec "U-03.05.01b", "Zbiorniki retencyjne i retencyjno-infiltracyjne"
    AddSpec "U-05.03.01", "Kanalizacja teletechniczna"
    AddSpec "U-07.07.01", "Oœwietlenie drogowe"
    AddSpec "U-10.15.01", "£¹cznoœæ drogi ekspresowej"

    LogOK "Specyfikacje", 0, "Wstawiono pe³n¹ listê (Sekcja 13)."
    Exit Sub
ErrH:
    LogErr "Specyfikacje", 0, "WstawSpecyfikacje_PelnaLista: " & Err.Number & " - " & Err.Description
End Sub

Private Sub AddSpec(ByVal Kod As String, ByVal Nazwa As String)
    Dim sql As String
    If Nz(DCount("*", "Specyfikacje", "Kod='" & Replace(Kod, "'", "''") & "'"), 0) = 0 Then
        sql = "INSERT INTO Specyfikacje (Kod, Nazwa, Wersja, DataSpec) VALUES ('" & _
              Replace(Kod, "'", "''") & "', '" & Replace(Nazwa, "'", "''") & "', '1.0', Date())"
        InsertIfMissing_Simple "Specyfikacje", "Kod='" & Replace(Kod, "'", "''") & "'", sql
    End If
End Sub
    '------------------------------------------------------
    ' Warstwy + domyœlne spec
    '------------------------------------------------------
    Dim warstwyData As Variant
    Dim strWarstwyData As String
    strWarstwyData = "NASYP;Nasyp;Nasyp;D-02.03.01"
    strWarstwyData = strWarstwyData & "||MROZOO;Warstwa mrozoochronna;Mrozo;D-04.02.02"
    strWarstwyData = strWarstwyData & "||KNN-PB;Podbudowa KNN;KNN;D-04.04.02"
    strWarstwyData = strWarstwyData & "||STAB-C;Stabilizacja cementem;Stabilizacja;D-04.05.01"
    strWarstwyData = strWarstwyData & "||AC-PB;Podbudowa AC;MMA;D-04.07.01"
    strWarstwyData = strWarstwyData & "||AC16W;AC16W;MMA;D-05.03.05A"
    strWarstwyData = strWarstwyData & "||AC11S;AC11S;MMA;D-05.03.05B"
    strWarstwyData = strWarstwyData & "||SMA11;SMA11;MMA;D-05.03.13"
    strWarstwyData = strWarstwyData & "||KNN-N;Nawierzchnia KNN;KNN;D-05.02.01"
    warstwyData = Split(strWarstwyData, "||")
    
    Dim w As Integer
    For w = LBound(warstwyData) To UBound(warstwyData)
        Dim warstwaInfo As Variant
        warstwaInfo = Split(warstwyData(w), ";")
        ExecSQL "INSERT INTO Warstwy(KodWarstwy, NazwaWarstwy, TypWarstwy) VALUES('" & warstwaInfo(0) & "','" & warstwaInfo(1) & "','" & warstwaInfo(2) & "')"
        ExecSQL "INSERT INTO WarstwaSpecDefault(WarstwaID, SpecyfikacjaID) " & _
                "SELECT WarstwaID, SpecyfikacjaID FROM Warstwy, Specyfikacje " & _
                "WHERE KodWarstwy='" & warstwaInfo(0) & "' AND Kod='" & warstwaInfo(3) & "'"
    Next w
    
    '------------------------------------------------------
    ' Materia³y + przypisania
    '------------------------------------------------------
    Dim materialyData As Variant
    Dim strMaterialyData As String
    strMaterialyData = "Grunt piaszczysty;Grunt;NASYP"
    strMaterialyData = strMaterialyData & "||Kruszywo 0/31,5;Kruszywo;KNN-PB,KNN-N"
    strMaterialyData = strMaterialyData & "||AC 22P 50/70;Asfalt;AC-PB"
    strMaterialyData = strMaterialyData & "||AC 16W 50/70;Asfalt;AC16W"
    strMaterialyData = strMaterialyData & "||AC 11S 50/70;Asfalt;AC11S"
    strMaterialyData = strMaterialyData & "||SMA 11 50/70;Asfalt;SMA11"
    materialyData = Split(strMaterialyData, "||")
    
    Dim m As Integer
    For m = LBound(materialyData) To UBound(materialyData)
        Dim materialInfo As Variant
        materialInfo = Split(materialyData(m), ";")
        ExecSQL "INSERT INTO Materialy(NazwaMaterialu, TypMaterialu) " & _
                "VALUES('" & materialInfo(0) & "','" & materialInfo(1) & "')"
        Dim warstwyMat As Variant
        warstwyMat = Split(materialInfo(2), ",")
        Dim wm As Integer
        For wm = LBound(warstwyMat) To UBound(warstwyMat)
            ExecSQL "INSERT INTO WarstwaMaterial(WarstwaID, MaterialID) " & _
                    "SELECT WarstwaID, MaterialID FROM Warstwy, Materialy " & _
                    "WHERE KodWarstwy='" & warstwyMat(wm) & "' AND NazwaMaterialu='" & materialInfo(0) & "'"
        Next wm
    Next m
    
    '------------------------------------------------------
    ' Recepty mieszanek
    '------------------------------------------------------
    Dim receptyData As Variant
    Dim strReceptyData As String
    strReceptyData = "AC11S;D-05.03.05B;AC 11S 50/70;AC11S-50/70;Recepta AC11S"
    strReceptyData = strReceptyData & "||AC16W;D-05.03.05A;AC 16W 50/70;AC16W-50/70;Recepta AC16W"
    strReceptyData = strReceptyData & "||AC-PB;D-04.07.01;AC 22P 50/70;AC22P-50/70;Recepta AC22P"
    strReceptyData = strReceptyData & "||SMA11;D-05.03.13;SMA 11 50/70;SMA11-50/70;Recepta SMA11"
    receptyData = Split(strReceptyData, "||")
    
    Dim rec As Integer
    For rec = LBound(receptyData) To UBound(receptyData)
        Dim receptaInfo As Variant
        receptaInfo = Split(receptyData(rec), ";")
        Dim strSQLRecepty As String
        strSQLRecepty = "INSERT INTO ReceptyMieszanek(WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) "
        strSQLRecepty = strSQLRecepty & "SELECT W.WarstwaID, S.SpecyfikacjaID, M.M... '" & receptaInfo(3) & "','" & receptaInfo(4) & "', Date() "
        strSQLRecepty = strSQLRecepty & "FROM Warstwy W, Specyfikacje S, Materialy M "
        strSQLRecepty = strSQLRecepty & "WHERE W.KodWarstwy='" & receptaInfo(0) & "...ceptaInfo(1) & "' AND M.NazwaMaterialu='" & receptaInfo(2) & "'"
        ExecSQL strSQLRecepty
    Next rec
    
    '------------------------------------------------------
    ' Obiekt testowy i lokalizacja
    '------------------------------------------------------
    ExecSQL "INSERT INTO Obiekty(KodObiektu, NazwaObiektu, KM_Poczatek_m, KM_Koniec_m, Inwestor, Kontrakt) " & _
            "VALUES('OB-001','Odcinek testowy Sx',0,500,'Test','Test-K')"
    
    ExecSQL "INSERT INTO Lokalizacje(ObiektID, KM_Start_m, KM_End_m, Pas, Szerokosc_m) " & _
            "SELECT ObiektID, 0, 500, 'lewy', 3.5 FROM Obiekty WHERE KodObiektu='OB-001'"
    
    MsgBox "Dane startowe zosta³y za³adowane pomyœlnie.", vbInformation
    Exit Sub
    
ErrH:
    MsgBox "B³¹d ZaladujDaneStartowe: " & Err.Description, vbCritical
End Sub
