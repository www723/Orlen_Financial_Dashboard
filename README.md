# Orlen Dashboard
 
![dashboard-orlen](https://github.com/user-attachments/assets/e6216d56-00ec-486c-bb77-cb7e6053a8bb)

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

3 Boxy, przedstawiające Zysk/Strate netto,Aktywa razem,Przepływy pieniężne z działalności operacyjnej, zmiany procentowe względem poprzedniego roku oraz odchylenia standardowe.

![obraz](https://github.com/user-attachments/assets/30f90f88-eefa-49ed-ad19-0f5f9b8bddbf)

Podstawą pozyskiwania kwot,zmiany procentowej czy odchylenia standardowego była funkcja XLOOKUP.

W przypadku kwot czy zmiany procentowej
```
=XLOOKUP(rok,$Z$5:$Z$25,$AA$5:$AA$25,"No value")
```
lub bardziej rozbudowana ,jeśli chodzi o odchylenie standardowe
```
=IF(ABS(XLOOKUP(rok, Z5:Z25, AA5:AA25) - AVERAGE(AA5:AA25)) <= STDEV.P(AA5:AA25), "W granicach normy", "Znaczna zmiana")
```
Natomiast w przypadku zmiany kolorów czcionek w polach tekstowcyh użyłem VBA
```
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Ustawienie pól tekstowych
    Dim txtBox1 As Object
    Dim txtBox2 As Object
    Dim txtBox3 As Object
    Dim txtBox4 As Object
    Dim txtBox5 As Object
    Dim txtBox6 As Object
    Dim wsród³o As Worksheet

    ' Ustawienie TextBoxów
    Set txtBox1 = ActiveSheet.Shapes("%change_z") ' Nazwa pierwszego TextBoxa
    Set txtBox2 = ActiveSheet.Shapes("ocena_odchylenia_z") ' Nazwa drugiego TextBoxa
    Set txtBox3 = ActiveSheet.Shapes("%change_a") ' Nazwa trzeciego TextBoxa
    Set txtBox4 = ActiveSheet.Shapes("ocena_odchylenia_a") ' Nazwa czwartego TextBoxa
    Set txtBox5 = ActiveSheet.Shapes("%change_p") ' Nazwa pi¹tego TextBoxa
    Set txtBox6 = ActiveSheet.Shapes("ocena_odchylenia_p") ' Nazwa szóstego TextBoxa

    
    Set wszrod³o = ThisWorkbook.Sheets("Data Validation") '

    ' Ustawienie t³a wszystkich TextBoxów na bialy
    txtBox1.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox2.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox3.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox4.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox5.Fill.ForeColor.RGB = RGB(255, 255, 255)
    txtBox6.Fill.ForeColor.RGB = RGB(255, 255, 255)

    ' Warunki dla pierwszego TextBoxa (txtBox1)
    If txtBox1.TextFrame2.TextRange.Text = "Brak danych" Then
        txtBox1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf wsród³o.Range("procentz").Value >= 0.01 Then
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
    ElseIf wsród³o.Range("procenta").Value >= 0.01 Then
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

    ' Warunki dla pi¹tego TextBoxa (txtBox5)
    If txtBox5.TextFrame2.TextRange.Text = "Brak danych" Then
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf wsród³o.Range("procentpp").Value >= 0.01 Then
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If

    ' Warunki dla szóstego TextBoxa (txtBox6)
    If txtBox6.TextFrame2.TextRange.Text = "W granicach normy" Then
        txtBox6.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 205, 50)
    Else
        txtBox6.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End If
End Sub
```


## Podsumowanie
