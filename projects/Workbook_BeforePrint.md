## OPCJA 1 – Application.BeforePrint (najlepsza, oficjalna)

Excel ma event Workbook_BeforePrint.

Jak to zrobić:

W ThisWorkbook:
```
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    MsgBox "Uwaga! Zaraz drukujesz. Sprawdź dane.", vbExclamation, "Kontrola przed drukiem"
End Sub
```

💡 Efekt:

Zawsze przed drukowaniem (Ctrl+P, przycisk, makro – wszystko)

Możesz nawet zablokować druk:
```
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    If MsgBox("Czy na pewno chcesz drukować?", vbQuestion + vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub
```
## OPCJA 2 – przechwycenie przycisku Drukuj (Application.OnKey)

Jeśli chcesz być sprytny i łapać Ctrl+P:
```
Sub HookPrint()
    Application.OnKey "^p", "MyPrintHandler"
End Sub

Sub MyPrintHandler()
    MsgBox "Drukowanie przechwycone 😎"
    Application.CommandBars.ExecuteMso "Print"
End Sub
```

⚠️ Minus:

działa tylko na skrót

jak ktoś kliknie z menu – ominie

## OPCJA 3 – własny przycisk „Drukuj” + makro

Najprostsze „korporacyjne” obejście:
```
Sub MyPrint()
    MsgBox "Sprawdź numer partii i datę!"
    ActiveWindow.SelectedSheets.PrintOut
End Sub
```

I przypinasz to pod przycisk.

👉 Moja rekomendacja bierz OPCJĘ 1 – Workbook_BeforePrint

Jest:
* czysta
* stabilna
* nie do obejścia
* audytor-friendly 😉


## Real life exaple
```
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    MsgBox "Uwaga! Musisz jeszcze da" & ChrW(263) & " zna" & ChrW(263) & " zespo" & ChrW(322) & "owi cyfryzacji o tym, " & ChrW(380) & "e trzeba zaktualizowa" & ChrW(263) & " ten plik w cyforwej produkcji", vbExclamation, "Nie zapomnij powiadomi" & ChrW(263) & " o aktualizacji"
End Sub
```
