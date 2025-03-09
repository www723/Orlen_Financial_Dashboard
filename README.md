# Orlen Dashboard
 
![orlen-dash](https://github.com/user-attachments/assets/c9bdf5c6-68ce-4b1f-88e0-eafeebf58609)


## Wstęp

Dashboard został stworzony,aby pomóc inwestorom w podjęciu świadomej decyzji inwestycyjnej poprzez zwizualizowanie danych liczbowych z Rachunku Zysków i Strat, Bilansu, Rachunku Przepływów Pieniężnych i policzonych wskaźników finansowych.

### Plik dashboardu


### Użyte umiejętności z Exela.
Poniższe umiejętności z Excela zostały użyte w tym dashboardzie:

- **💪🏻 Power Query** 
- **👨‍💻 VBA**
- **📉 Wykresy**
- **🧮 Formuły i funkcje**
- **❎ Data Validation**


### Dane użyte w  dashboardzie

Dane z:

- **💰 Rachunku Zysków i Strat**
- **📝 Bilansu**
- **🌊 Rachunku Przepływów Pieniężnych**

Zostały pozyskane ze strony https://www.biznesradar.pl/ i stanowiły fundamenty mojego dashboardu.

Z pozyskanych danych obliczyłem:
- **🎯 Wskaźniki finansowe**

## Budowa dashboardu

### 🔍 Umiejętność:**💪🏻 Power Query**


#### 📥 Pozyskanie danych

   - Pierwsze użyłem **💪🏻 Power Query**,aby pozyskać dane ze strony https://www.biznesradar.pl/ i stworzyłem 3 zapytania
   - 🗃️ Pierwsze z danymi RZiS.
   - 🔧 Drugie z danymi Bilansu.
   - 🏆 Trzecie z danymi RPP.

#### 🧹 Oszczyszczanie danych

- 😱 Dane zastałem w takim stanie ![obraz](https://github.com/user-attachments/assets/d164e5e8-519c-4d55-9639-77aa2199c248)

- 🐾 Kroki,które uczyniłem

 ![obraz](https://github.com/user-attachments/assets/0e29654b-0a73-47b7-a595-d104e9a5cdd1)

Wraz z kodem
```
let
    Source = Web.BrowserContents("https://www.biznesradar.pl/raporty-finansowe-rachunek-zyskow-i-strat/ORLEN"),
    #"Pozyskanie danych ze strony" = Html.Table(Source, 
        List.Transform({1..23}, each { "Column" & Text.From(_), "TABLE.report-table > * > TR > :nth-child(" & Text.From(_) & ")" }),
        [RowSelector="TABLE.report-table > * > TR"]
    ),
    #"Zmiana typu danych" = Table.TransformColumnTypes(#"Pozyskanie danych ze strony", List.Transform(Table.ColumnNames(#"Pozyskanie danych ze strony"), each {_, type text})),
    #"Oczyszczenie kolumn" = Table.TransformColumns(#"Zmiana typu danych",  
        List.Transform(Table.ColumnNames(#"Zmiana typu danych"),  
            each {_, each Text.BeforeDelimiter(_, "r/r"), type text}  
        )  
    ),
    #"Usunięcie odstępów" = Table.ReplaceValue(#"Oczyszczenie kolumn"," ","",Replacer.ReplaceText,{"Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22"}),
    #"Zmiana typu danych na walute" = Table.TransformColumnTypes(#"Usunięcie odstępów",{{"Column2", Currency.Type}, {"Column3", Currency.Type}, {"Column4", Currency.Type}, {"Column5", Currency.Type}, {"Column6", Currency.Type}, {"Column7", Currency.Type}, {"Column8", Currency.Type}, {"Column9", Currency.Type}, {"Column10", Currency.Type}, {"Column11", Currency.Type}, {"Column12", Currency.Type}, {"Column13", Currency.Type}, {"Column14", Currency.Type}, {"Column15", Currency.Type}, {"Column16", Currency.Type}, {"Column17", Currency.Type}, {"Column18", Currency.Type}, {"Column19", Currency.Type}, {"Column20", Currency.Type}, {"Column21", Currency.Type}, {"Column22", Currency.Type}}),
    #"Transponowanie tabeli" = Table.Transpose(#"Zmiana typu danych na walute")
in
    #"Transponowanie tabeli"
```

- 🏁 Rezultat
![obraz](https://github.com/user-attachments/assets/bbdb8b2a-ea69-4430-a86b-37ac629cd666)


### 🔍 Umiejętność:**👨‍💻 VBA**

Użyłem języka programowania **VBA** ,aby
   - ⚙️ Zautomatyzować zmianę koloru.
   - 🚀 Usprawnić prezentacje danych.
   - 🎨 Zapewnić spójność wizualną.


```
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Ustawienia pól tekstowych
    Dim txtBox1 As Object
    Dim txtBox2 As Object
    Dim txtBox3 As Object
    Dim txtBox4 As Object
    Dim txtBox5 As Object
    Dim txtBox6 As Object
    Dim wszrodlo As Worksheet

    ' Ustawienie TextBoxów
    Set txtBox1 = ActiveSheet.Shapes("%change_z")
    Set txtBox2 = ActiveSheet.Shapes("ocena_odchylenia_z")
    Set txtBox3 = ActiveSheet.Shapes("%change_a")
    Set txtBox4 = ActiveSheet.Shapes("ocena_odchylenia_a")
    Set txtBox5 = ActiveSheet.Shapes("%change_p")
    Set txtBox6 = ActiveSheet.Shapes("ocena_odchylenia_p")

    Set wszrodlo = ThisWorkbook.Sheets("Data Validation")

    ' Ustaw tlo wszystkich TextBoxów na biale
    txtBox1.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox2.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox3.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox4.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox5.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox6.Fill.ForeColor.RGB = RGB(255, 255, 255)

    ' Warunki dla pierwszego TextBoxa (txtBox1)
    If txtBox1.TextFrame2.TextRange.Text = "Brak danych" Then
        txtBox1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf wszrodlo.Range("procentz").Value >= 0.01 Then
        txtBox1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla drugiego TextBoxa (txtBox2)
    If txtBox2.TextFrame2.TextRange.Text = "W granicach normy" Then
        txtBox2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla trzeciego TextBoxa (txtBox3)
    If txtBox3.TextFrame2.TextRange.Text = "Brak danych" Then
        txtBox3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf wszrodlo.Range("procenta").Value >= 0.01 Then
        txtBox3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla czwartego TextBoxa (txtBox4)
    If txtBox4.TextFrame2.TextRange.Text = "W granicach normy" Then
        txtBox4.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox4.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla piatego TextBoxa (txtBox5)
    If txtBox5.TextFrame2.TextRange.Text = "Brak danych" Then
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf wszrodlo.Range("procentpp").Value >= 0.01 Then
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla szostego TextBoxa (txtBox6)
    If txtBox6.TextFrame2.TextRange.Text = "W granicach normy" Then
        txtBox6.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox6.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

End Sub
```
![obraz](https://github.com/user-attachments/assets/a29f1b1a-e697-4bd7-8c72-5f162cf35053)



### 🔍 Umiejętność:**📉 Wykresy**

#### 📊 Aktywa razem,Zysk/Strata netto,Przepływy pieniężne z działalności operacyjnej wykres kolumnowy
![obraz](https://github.com/user-attachments/assets/3b774283-2ba5-4345-8573-4f4606bf7a5b)

- 🛠️ **Funkcja excela:** Wykorzystano wykres kolumnowy(z formatowaniem liczb na miliony i wyróżnianiem wybranego roku) dodatkowo zoptymalizowano układ dla przejrzystości.
- 🎨 **Wybór designu:** Pionowy wykres kolumnowy do porównania kluczowych danych finansowych.
- 📉 **Organizacja danych:** Zachowanie chronologii dat dla polepszenia czytelności.
- 💡 **Spostrzeżenia:** Wykres pozwala na szybkie spostrzeżenie jakie lata w firmie były najbardziej , a jakie najmniej np. zyskowne,jak zmieniały się aktywa w czasie czy przepływy pieniężne.


#### 📈 Wykres wskaźników zadłużenia,płynności finansowej,rentowności czy sprawności finansowej
![obraz](https://github.com/user-attachments/assets/dcaa1a5c-5610-4ed8-9131-c99dfc9594ac)

- 🛠️ **Funkcja excela:** Wykorzystano wykres liniowy ze wskaźnikami i wyróżnieniem wybranego roku,zoptymalizowano układ dla przejrzystości.
- 🎨 **Wybór designu:** Wykres liniowy ze wskaźnikami,został wybrany aby pokazać jak zmieniały się wskaźniki w czasie.
- 👁️ **Poprawki wizualne:** Dodano linie pionowe, które będą wskazywały konkretne daty i ułatwią odnalezienie odpowiednich punktów na wykresie
- 💡 **Spostrzeżenia:** Wykres daje możliwość na szczególowe przenalizowanie wybranego wskaźnika np.spadki płynności w 2008 roku.


#### 🚀 Wykresy struktury kosztów,podziału aktywów i pasywów.
![obraz](https://github.com/user-attachments/assets/ed8a1716-73af-4fef-966d-f87f7da0332f)

- 🛠️ **Funkcja excela:** Wykorzystano wykres kołowy.
- 🎨 **Wybór designu:** Wykres kołowy najlepiej obrazuje struktury danych czy podziały.
- 👁️ **Poprawki wizualne:** Zwiększono czcionki,zmieniono układ dla czytelności.
- 💡 **Spostrzeżenia:** Wykres pomaga w szybkim uchwyceniu najważniejszych kategorii i ich udziału w całości, co usprawnia podejmowanie decyzji.



### 🔍 Umiejętność:**🧮 Formuły i funkcje**

####  ⌕ Wyszukiwanie pożądanej wartości

##### **Prosta funkcja XLOOKUP**
```
=XLOOKUP(year,$Z$5:$Z$25,$AB$5:$AB$25)
```
- 🔢**Cel formuły:** Umożliwia zlokalizowanie wartości przypisanej do komórki year(rok).


##### 🥵 **Zaawansowana funkcja XLOOKUP**
```
=IF(ABS(XLOOKUP(rok, Z5:Z25, AA5:AA25) - AVERAGE(AA5:AA25)) <= STDEV.P(AA5:AA25), "W granicach normy", "Znaczna zmiana")
```
- 🔢 **Co robi?:**Formuła sprawdza, czy wartość, którą zwróciła funkcja XLOOKUP dla wybranego roku, mieści się w granicach „normalnej” zmienności, czyli w granicach odchylenia standardowego od średniej wartości w kolumnie AA5:AA25.
- 🛠️  **Jak działa?:** XLOOKUP znajduje wartość dla wybranego roku.Oblicza różnicę między tą wartością a średnią w kolumnie.Sprawdza, czy ta różnica mieści się w granicach odchylenia standardowego.Na tej podstawie zwraca odpowiedni wynik.
- 🎯 **Dodatkowo** Formuła używa: ```AVERAGE()```podaje średnią zakresu danych,```STDEV.P()```ukazuje odchylenie standardowe,```ABS()```zwraca wartość bezwzględną liczby

##### 💥**Choose w połączeniu z match**

```
=CHOOSE(MATCH('Orlen dashboard'!J4, {"Zadłużenie ogólne(%)","Zadłużenie do kapitału własnego(%)","Pokrycie odsetek zyskiem","Dług netto do EBITDA","Płynności bieżącej","Płynności szybkiej","Płynności natychmiastowej","Marża operacyjna EBIT(%)","ROA(Rentowność aktywów)(%)","ROE (Rentowność kapitału własnego)(%)","Marża netto(%)","Rotacja należności( w dniach)","Rotacja zobowiazań( w dniach)","Rotacja zapasów( w dniach)"}, 0),Wska_roczne8[[#All],[Zadłużenie ogólne(%)]],Wska_roczne8[[#All],[Zadłużenie do kapitału własnego(%)]],Wska_roczne8[[#All],[Pokrycie odsetek zyskiem]],Wska_roczne8[[#All],[Dług netto do EBITDA]],Wska_roczne8[[#All],[Płynności bieżącej]],Wska_roczne8[[#All],[Płynności szybkiej]],Wska_roczne8[[#All],[Płynności natychmiastowej]],Wska_roczne8[[#All],[Marża operacyjna EBIT(%)]],Wska_roczne8[[#All],[ROA(Rentowność aktywów)(%)]],Wska_roczne8[[#All],[ROE (Rentowność kapitału własnego)(%)]],Wska_roczne8[[#All],[Marża netto(%)]],Wska_roczne8[[#All],[Rotacja należności( w dniach)]],Wska_roczne8[[#All],[Rotacja zobowiazań( w dniach)]],Wska_roczne8[[#All],[Rotacja zapasów( w dniach)]])
```

- 🔢 **Co robi?:** Formuła wybiera odpowiednią kolumnę wskaźników finansowych na podstawie wskazanego roku lub wartości.
- 🛠️ **Jak działa?** Sprawdza, który wskaźnik użytkownik wybrał na dashboardzie, szuka go, a następnie zwraca odpowiadającą kolumnę.


### 🔍 Umiejętność:**❎ Data Validation**

#### 🔍 Lista do wyboru

- 🎯 Wprowadzanie danych przez użytkownika jest ograniczone do wstępnie zdefiniowanych sprawdzonych typów harmonogramów
- 🚫 Zapobiega się wprowadzaniu nieprawidłowych lub niespójnych danych
- 👥 Ogólna użyteczność dashboardu jest zwiększona

![datavalidationgif](https://github.com/user-attachments/assets/0afaea23-af39-42a3-8269-d0859ff3b2f9)









## Podsumowanie
