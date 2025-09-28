Attribute VB_Name = "M03_DaneStartowe"
Option Compare Database
Option Explicit

'==========================================================
' Moduł: M03_DaneStartowe
' Wczytanie danych startowych (specyfikacje, sita, parametry,
' warstwy, materiały, recepty, obiekt testowy)
'==========================================================

'==================== Narzędzia ====================
Private Sub ExecSQL(ByVal sql As String)
    On Error GoTo ErrH
    CurrentDb.Execute sql, dbFailOnError
    Exit Sub
ErrH:
    MsgBox "Błąd ExecSQL: " & Err.Number & " - " & Err.Description & vbCrLf & Left$(sql, 255), _
           vbExclamation, "M03.ExecSQL"
    Err.Raise Err.Number, "M03.ExecSQL", Err.Description
End Sub

Private Sub InsertIfMissing(ByVal tableName As String, ByVal criteria As String, ByVal insertSQL As String)
    On Error GoTo ErrH
    If Nz(DCount("*", tableName, criteria), 0) = 0 Then
        ExecSQL insertSQL
    End If
    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.InsertIfMissing", Err.Description
End Sub

Private Function SqlQuote(ByVal value As String) As String
    SqlQuote = Replace(value, "'", "''")
End Function

Private Sub WyczyscDane()
    On Error GoTo ErrH

    Dim tables As Variant
    Dim t As Variant

    tables = Split("Lokalizacje|Obiekty|ReceptyMieszanek|WarstwaMaterial|WarstwaSpecDefault|" & _
                   "Materialy|Warstwy|Specyfikacje|ParametryJakosci|RodzajeBadania|Sita", "|")

    On Error Resume Next
    For Each t In tables
        CurrentDb.Execute "DELETE FROM " & CStr(t)
    Next t
    On Error GoTo ErrH

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WyczyscDane", Err.Description
End Sub

Private Sub WstawSita()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim idx As Long
    Dim roz As String
    Dim sql As String

    raw = Split("0.063|0.125|0.25|0.5|1|2|4|8|11.2|16|22.4|31.5", "|")

    For idx = LBound(raw) To UBound(raw)
        roz = Trim$(CStr(raw(idx)))
        sql = "INSERT INTO Sita (Rozmiar_mm, Opis, Kolejnosc) " & _
              "SELECT " & roz & ", 'Sito " & roz & " mm', " & (idx + 1) & " " & _
              "WHERE NOT EXISTS (SELECT 1 FROM Sita WHERE Rozmiar_mm=" & roz & ")"
        ExecSQL sql
    Next idx

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawSita", Err.Description
End Sub

Private Sub WstawRodzajeBadania()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim nazwa As Variant
    Dim sql As String

    raw = Split("Uziarnienie|Gęstość objętościowa|Wilgotność optymalna|Zawartość asfaltu|" & _
                "Zawartość asfaltu rozpuszczalnego|Gęstość maksymalna teoretyczna (Gmm)|" & _
                "Rozciąganie pośrednie (ITS)|Odporność na działanie wody (ITSR)|" & _
                "Wskaźnik piaskowy (SE)|Los Angeles (LA)|Mrozoodporność - ubytek masy|" & _
                "Zawartość pyłów <0,063 mm|Moduł odkształcenia E2|Moduł odkształcenia E1|" & _
                "Stosunek E2/E1|Nośność CBR|Wskaźnik zagęszczenia Is|Zawartość cementu (stabilizacja)|" & _
                "Zawartość asfaltu odzyskanego (Pb_sol)|Temperatura mieszanki przy wbudowaniu", "|")

    For Each nazwa In raw
        nazwa = Trim$(CStr(nazwa))
        If Len(nazwa) > 0 Then
            sql = "INSERT INTO RodzajeBadania (Nazwa) " & _
                  "SELECT '" & SqlQuote(CStr(nazwa)) & "' " & _
                  "WHERE NOT EXISTS (SELECT 1 FROM RodzajeBadania WHERE Nazwa='" & SqlQuote(CStr(nazwa)) & "')"
            ExecSQL sql
        End If
    Next nazwa

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawRodzajeBadania", Err.Description
End Sub

Private Sub WstawParametryJakosci()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim sql As String

    raw = Split(_
        "Gęstość objętościowa;?b;g/cm3|" & _
        "Wilgotność optymalna;wopt;%|" & _
        "Zawartość asfaltu;Pb;%|" & _
        "Zawartość asfaltu rozpuszczalnego;Pb_sol;%|" & _
        "Wolne przestrzenie w mieszance;Vv;%|" & _
        "Pusta przestrzeń między ziarnami;VMA;%|" & _
        "Wypełnienie wolnych przestrzeni lepiszczem;VFB;%|" & _
        "Gęstość maksymalna teoretyczna;Gmm;g/cm3|" & _
        "Gęstość objętościowa MMA;?m;g/cm3|" & _
        "Wolne przestrzenie w ziarnach grubych;VCA;%|" & _
        "Wytrzymałość na rozciąganie pośrednie;ITS;MPa|" & _
        "Odporność na działanie wody;ITSR;%|" & _
        "Wskaźnik piaskowy;SE;%|" & _
        "Los Angeles;LA;%|" & _
        "Mrozoodporność - ubytek masy;F_ubytek;%|" & _
        "Zawartość pyłów <0,063 mm;P063;%|" & _
        "Nośność CBR;CBR;%|" & _
        "Moduł odkształcenia E2;E2;MPa|" & _
        "Moduł odkształcenia E1;E1;MPa|" & _
        "Stosunek E2/E1;E2E1;-|" & _
        "Wskaźnik zagęszczenia;Is;%|" & _
        "Zawartość cementu (stabilizacja);C_cem;%|" & _
        "Temperatura mieszanki przy wbudowaniu;T_mix;°C", "|")

    For Each entry In raw
        entry = Trim$(CStr(entry))
        If Len(entry) > 0 Then
            parts = Split(entry, ";")
            If UBound(parts) >= 2 Then
                sql = "INSERT INTO ParametryJakosci (NazwaParametru, KodParametru, Jednostka) " & _
                      "SELECT '" & SqlQuote(Trim$(parts(0))) & "', '" & SqlQuote(Trim$(parts(1))) & "', '" & SqlQuote(Trim$(parts(2))) & "' " & _
                      "WHERE NOT EXISTS (SELECT 1 FROM ParametryJakosci WHERE KodParametru='" & SqlQuote(Trim$(parts(1))) & "')"
                ExecSQL sql
            End If
        End If
    Next entry

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawParametryJakosci", Err.Description
End Sub

Private Sub AddSpec(ByVal kod As String, ByVal nazwa As String)
    Dim sql As String
    sql = "INSERT INTO Specyfikacje (Kod, Nazwa, Wersja, DataSpec) VALUES ('" & SqlQuote(kod) & "', '" & SqlQuote(nazwa) & "', '1.0', Date())"
    InsertIfMissing "Specyfikacje", "Kod='" & SqlQuote(kod) & "'", sql
End Sub

Private Sub WstawSpecyfikacje_PelnaLista()
    On Error GoTo ErrH

    ' --- Wymagania Ogólne i D-01... ---
    AddSpec "D-M-00.00.00", "Wymagania Ogólne"
    AddSpec "D-01.00.00", "ROBOTY PRZYGOTOWAWCZE"
    AddSpec "D-01.01.01", "Odtworzenie trasy i punktów wysokościowych"
    AddSpec "D-01.02.01", "Usunięcie drzew, zagajników i krzewów"
    AddSpec "D-01.02.01A", "Zabezpieczenie istniejących drzew i krzewów na okres wykonywania robót"
    AddSpec "D-01.02.02", "Zdjęcie warstwy humusu"
    AddSpec "D-01.02.03", "Wyburzenie obiektów budowlanych"
    AddSpec "D-01.02.04", "Rozbiórki elementów dróg i ulic"

    ' --- D-02... Roboty ziemne ---
    AddSpec "D-02.00.01", "Roboty ziemne. Wymagania ogólne"
    AddSpec "D-02.01.01", "Roboty ziemne. Wykonanie wykopów"
    AddSpec "D-02.01.01A", "Platformy robocze dla ciężkiego sprzętu budowlanego"
    AddSpec "D-02.01.01B", "Wzmocnienie podłoża gruntowego. Wymiana gruntów"
    AddSpec "D-02.01.01C", "Wzmocnienie podłoża gruntowego. Materace geosyntetyczne"
    AddSpec "D-02.01.01D", "Wzmocnienie podłoża gruntowego. Metoda drenów pionowych i nasypu przeciążającego"
    AddSpec "D-02.01.01E", "Wzmocnienie podłoża gruntowego. Kolumny DSM"
    AddSpec "D-02.01.01F", "Wzmocnienie podłoża gruntowego. Metoda iniekcji strumieniowej Jet Grouting"
    AddSpec "D-02.01.01G", "Wzmocnienie podłoża gruntowego. Kolumny wiercone"
    AddSpec "D-02.01.01H", "Wzmocnienie podłoża gruntowego. Kolumny betonowo-wiercone"
    AddSpec "D-02.01.01I", "Wzmocnienie podłoża gruntowego. Prefabrykowane pale żelbetowe"
    AddSpec "D-02.01.01J", "Wzmocnienie podłoża gruntowego. Pale wiercone typu CFA"
    AddSpec "D-02.03.01", "Roboty ziemne. Wykonanie nasypów"

    ' --- D-03... Odwodnienie korpusu ---
    AddSpec "D-03.00.00", "ODWODNIENIE KORPUSU DROGOWEGO"
    AddSpec "D-03.01.01", "Przepusty pod koroną drogi"
    AddSpec "D-03.03.01", "Sączki podłużne"

    ' --- D-04... Podbudowy ---
    AddSpec "D-04.00.00", "PODBUDOWY"
    AddSpec "D-04.02.01", "Warstwa odcinająca"
    AddSpec "D-04.02.02", "Warstwa mrozoochronna/odsączająca"
    AddSpec "D-04.03.01", "Oczyszczenie i skropienie warstw konstrukcyjnych"
    AddSpec "D-04.04.02", "Podbudowa pomocnicza i zasadnicza z mieszanki nie związanej"
    AddSpec "D-04.05.00", "Warstwa ulepszonego podłoża z gruntu stabilizowanego spoiwem hydraulicznym lub wapnem"
    AddSpec "D-04.05.01", "Podbudowa i warstwa mrozoochronna z mieszanki związanej cementem"
    AddSpec "D-04.07.01", "Warstwa podbudowy z AC"
    AddSpec "D-04.10.01", "Podbudowa z mieszanki mineralno-cementowo-emulsyjnej"

    ' --- D-05... Nawierzchnie ---
    AddSpec "D-05.00.00", "NAWIERZCHNIE"
    AddSpec "D-05.02.01", "Nawierzchnia z kruszywa niezwiązanego"
    AddSpec "D-05.03.01", "Nawierzchnia z kostki kamiennej"
    AddSpec "D-05.03.05A", "Nawierzchnia z betonu asfaltowego warstwa wiążąca"
    AddSpec "D-05.03.05B", "Nawierzchnia z betonu asfaltowego warstwa ścieralna"
    AddSpec "D-05.03.05C", "Warstwa podbudowy i wiążąca z WMS"
    AddSpec "D-05.03.13", "Nawierzchnia z mieszanki mastyksowo-grysowej (SMA)"
    AddSpec "D-05.03.23", "Nawierzchnia i chodniki z kostki brukowej"

    ' --- D-06... Roboty wykończeniowe ---
    AddSpec "D-06.00.00", "ROBOTY WYKOŃCZENIOWE"
    AddSpec "D-06.01.01A", "Umocnienie skarp, rowów i pasa dzielącego przez humusowanie i obsiew"
    AddSpec "D-06.01.01B", "Umocnienie skarp i rowów elementami prefabrykowanymi"
    AddSpec "D-06.02.01", "Przepusty pod zjazdami"
    AddSpec "D-06.03.01", "Umocnienie poboczy"

    ' --- D-07... Urządzenia BRD ---
    AddSpec "D-07.00.00", "URZĄDZENIA BEZPIECZEŃSTWA RUCHU"
    AddSpec "D-07.01.01", "Oznakowanie poziome"
    AddSpec "D-07.02.01", "Oznakowanie pionowe"
    AddSpec "D-07.02.02", "Słupki prowadzące i krawędziowe oraz znaki kilometrowe i hektometrowe"
    AddSpec "D-07.05.01", "Bariery ochronne"
    AddSpec "D-07.06.01", "Ogrodzenia dróg"
    AddSpec "D-07.06.02", "Urządzenia zabezpieczające ruch pieszy i rowerowy"
    AddSpec "D-07.08.01", "Ekrany akustyczne. Ekrany przeciwolśnieniowe. Ekrany ekologiczne"
    AddSpec "D-07.09.01", "Drogowy ekran przeciwolśnieniowy"

    ' --- D-08... Elementy ulic ---
    AddSpec "D-08.00.00", "ELEMENTY ULIC"
    AddSpec "D-08.01.01", "Krawężniki betonowe"
    AddSpec "D-08.01.02", "Krawężniki kamienne"
    AddSpec "D-08.03.01", "Obrzeża betonowe"
    AddSpec "D-08.05.01", "Ścieki"

    ' --- D-09... Zieleń drogowa ---
    AddSpec "D-09.00.00", "ZIELEŃ DROGOWA"
    AddSpec "D-09.01.01", "Zieleń drogowa"

    ' --- M-11... M-20... (mostowe) ---
    AddSpec "M-11.00.00", "Fundamentowanie"
    AddSpec "M-12.00.00", "Zbrojenie"
    AddSpec "M-13.00.00", "Beton"
    AddSpec "M-14.00.00", "Konstrukcje stalowe"
    AddSpec "M-15.00.00", "Izolacje i nawierzchnie"
    AddSpec "M-16.00.00", "Odwodnienie"
    AddSpec "M-17.00.00", "Łożyska"
    AddSpec "M-18.00.00", "Urządzenia dylatacyjne"
    AddSpec "M-19.00.00", "Elementy zabezpieczające"
    AddSpec "M-20.01.00", "Inne roboty mostowe. Roboty przyobiektowe"
    AddSpec "M-20.02.00", "Inne roboty mostowe. Roboty na obiekcie"

    ' --- U-01... U-10... (uzbrojenie) ---
    AddSpec "U-01.03.01", "Przebudowa napowietrznych linii elektroenergetycznych"
    AddSpec "U-01.03.02", "Przebudowa kablowych linii energetycznych"
    AddSpec "U-01.03.03", "Przebudowa napowietrznych linii telekomunikacyjnych"
    AddSpec "U-01.03.04", "Przebudowa kablowych linii telekomunikacyjnych"
    AddSpec "U-01.03.05", "Przebudowa i budowa podziemnych sieci wodociągowych"
    AddSpec "U-01.03.06", "Przebudowa podziemnych sieci gazowych"
    AddSpec "U-01.03.07", "Przebudowa urządzeń melioracyjnych"
    AddSpec "U-03.02.01", "Kanalizacja deszczowa i sanitarna"
    AddSpec "U-03.05.01b", "Zbiorniki retencyjne i retencyjno-infiltracyjne"
    AddSpec "U-05.03.01", "Kanalizacja teletechniczna"
    AddSpec "U-07.07.01", "Oświetlenie drogowe"
    AddSpec "U-10.15.01", "Łączenie drogi ekspresowej"

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawSpecyfikacje_PelnaLista", Err.Description
End Sub

Private Sub WstawWarstwyIDomyslneSpec()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim sql As String

    raw = Split(_
        "NASYP;Nasyp;Nasyp;D-02.03.01|" & _
        "MROZOO;Warstwa mrozoochronna;Mrozo;D-04.02.02|" & _
        "KNN-PB;Podbudowa KNN;KNN;D-04.04.02|" & _
        "STAB-C;Stabilizacja cementem;Stabilizacja;D-04.05.01|" & _
        "AC-PB;Podbudowa AC;MMA;D-04.07.01|" & _
        "AC16W;AC16W;MMA;D-05.03.05A|" & _
        "AC11S;AC11S;MMA;D-05.03.05B|" & _
        "SMA11;SMA11;MMA;D-05.03.13|" & _
        "KNN-N;Nawierzchnia KNN;KNN;D-05.02.01", "|")

    For Each entry In raw
        entry = Trim$(CStr(entry))
        If Len(entry) > 0 Then
            parts = Split(entry, ";")
            If UBound(parts) >= 3 Then
                sql = "INSERT INTO Warstwy (KodWarstwy, NazwaWarstwy, TypWarstwy) " & _
                      "SELECT '" & SqlQuote(Trim$(parts(0))) & "', '" & SqlQuote(Trim$(parts(1))) & "', '" & SqlQuote(Trim$(parts(2))) & "' " & _
                      "WHERE NOT EXISTS (SELECT 1 FROM Warstwy WHERE KodWarstwy='" & SqlQuote(Trim$(parts(0))) & "')"
                ExecSQL sql

                sql = "INSERT INTO WarstwaSpecDefault (WarstwaID, SpecyfikacjaID) " & _
                      "SELECT W.WarstwaID, S.SpecyfikacjaID FROM Warstwy W, Specyfikacje S " & _
                      "WHERE W.KodWarstwy='" & SqlQuote(Trim$(parts(0))) & "' AND S.Kod='" & SqlQuote(Trim$(parts(3))) & "' " & _
                      "AND NOT EXISTS (SELECT 1 FROM WarstwaSpecDefault WHERE WarstwaID=W.WarstwaID)"
                ExecSQL sql
            End If
        End If
    Next entry

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawWarstwyIDomyslneSpec", Err.Description
End Sub

Private Sub WstawMaterialy()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim sql As String
    Dim warstwy As Variant
    Dim w As Variant

    raw = Split(_
        "Grunt piaszczysty;Grunt;NASYP|" & _
        "Kruszywo 0/31,5;Kruszywo;KNN-PB,KNN-N|" & _
        "AC 22P 50/70;Asfalt;AC-PB|" & _
        "AC 16W 50/70;Asfalt;AC16W|" & _
        "AC 11S 50/70;Asfalt;AC11S|" & _
        "SMA 11 50/70;Asfalt;SMA11", "|")

    For Each entry In raw
        entry = Trim$(CStr(entry))
        If Len(entry) > 0 Then
            parts = Split(entry, ";")
            If UBound(parts) >= 2 Then
                sql = "INSERT INTO Materialy (NazwaMaterialu, TypMaterialu) " & _
                      "SELECT '" & SqlQuote(Trim$(parts(0))) & "', '" & SqlQuote(Trim$(parts(1))) & "' " & _
                      "WHERE NOT EXISTS (SELECT 1 FROM Materialy WHERE NazwaMaterialu='" & SqlQuote(Trim$(parts(0))) & "')"
                ExecSQL sql

                warstwy = Split(parts(2), ",")
                For Each w In warstwy
                    w = Trim$(CStr(w))
                    If Len(w) > 0 Then
                        sql = "INSERT INTO WarstwaMaterial (WarstwaID, MaterialID) " & _
                              "SELECT W.WarstwaID, M.MaterialID FROM Warstwy W, Materialy M " & _
                              "WHERE W.KodWarstwy='" & SqlQuote(CStr(w)) & "' AND M.NazwaMaterialu='" & SqlQuote(Trim$(parts(0))) & "' " & _
                              "AND NOT EXISTS (SELECT 1 FROM WarstwaMaterial WHERE WarstwaID=W.WarstwaID AND MaterialID=M.MaterialID)"
                        ExecSQL sql
                    End If
                Next w
            End If
        End If
    Next entry

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawMaterialy", Err.Description
End Sub

Private Sub WstawRecepty()
    On Error GoTo ErrH

    Dim raw As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim sql As String

    raw = Split(_
        "AC11S;D-05.03.05B;AC 11S 50/70;AC11S-50/70;Recepta AC11S|" & _
        "AC16W;D-05.03.05A;AC 16W 50/70;AC16W-50/70;Recepta AC16W|" & _
        "AC-PB;D-04.07.01;AC 22P 50/70;AC22P-50/70;Recepta AC22P|" & _
        "SMA11;D-05.03.13;SMA 11 50/70;SMA11-50/70;Recepta SMA11", "|")

    For Each entry In raw
        entry = Trim$(CStr(entry))
        If Len(entry) > 0 Then
            parts = Split(entry, ";")
            If UBound(parts) >= 4 Then
                sql = "INSERT INTO ReceptyMieszanek (WarstwaID, SpecyfikacjaID, MaterialID, KodRecepty, Opis, DataUtw) " & _
                      "SELECT W.WarstwaID, S.SpecyfikacjaID, M.MaterialID, '" & SqlQuote(Trim$(parts(3))) & "', '" & SqlQuote(Trim$(parts(4))) & "', Date() " & _
                      "FROM Warstwy W, Specyfikacje S, Materialy M " & _
                      "WHERE W.KodWarstwy='" & SqlQuote(Trim$(parts(0))) & "' AND S.Kod='" & SqlQuote(Trim$(parts(1))) & "' " & _
                      "AND M.NazwaMaterialu='" & SqlQuote(Trim$(parts(2))) & "' " & _
                      "AND NOT EXISTS (SELECT 1 FROM ReceptyMieszanek WHERE WarstwaID=W.WarstwaID AND SpecyfikacjaID=S.SpecyfikacjaID " & _
                      "AND KodRecepty='" & SqlQuote(Trim$(parts(3))) & "')"
                ExecSQL sql
            End If
        End If
    Next entry

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawRecepty", Err.Description
End Sub

Private Sub WstawObiektILokalizacje()
    On Error GoTo ErrH

    Dim sql As String

    sql = "INSERT INTO Obiekty (KodObiektu, NazwaObiektu, KM_Poczatek_m, KM_Koniec_m, Inwestor, Kontrakt) " & _
          "SELECT 'OB-001', 'Odcinek testowy Sx', 0, 500, 'Test', 'Test-K' " & _
          "WHERE NOT EXISTS (SELECT 1 FROM Obiekty WHERE KodObiektu='OB-001')"
    ExecSQL sql

    sql = "INSERT INTO Lokalizacje (ObiektID, KM_Start_m, KM_End_m, Pas, Szerokosc_m) " & _
          "SELECT O.ObiektID, 0, 500, 'lewy', 3.5 FROM Obiekty O " & _
          "WHERE O.KodObiektu='OB-001' AND NOT EXISTS (" & _
          "SELECT 1 FROM Lokalizacje L WHERE L.ObiektID=O.ObiektID AND L.KM_Start_m=0 AND L.KM_End_m=500 AND L.Pas='lewy')"
    ExecSQL sql

    Exit Sub
ErrH:
    Err.Raise Err.Number, "M03.WstawObiektILokalizacje", Err.Description
End Sub

'==================== Public API ====================
Public Sub ZaladujDaneStartowe()
    On Error GoTo ErrH

    WyczyscDane
    WstawSita
    WstawRodzajeBadania
    WstawParametryJakosci
    WstawSpecyfikacje_PelnaLista
    WstawWarstwyIDomyslneSpec
    WstawMaterialy
    WstawRecepty
    WstawObiektILokalizacje

    MsgBox "Dane startowe zostały załadowane pomyślnie.", vbInformation, "M03"
    Exit Sub

ErrH:
    MsgBox "Błąd ZaladujDaneStartowe: " & Err.Number & " - " & Err.Description, vbCritical, "M03"
End Sub
