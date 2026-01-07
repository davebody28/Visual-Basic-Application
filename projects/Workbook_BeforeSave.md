Perfekcyjnie – Excel ma event Workbook_BeforeSave, który działa analogicznie do BeforePrint. Możesz w nim zrobić komunikat, walidację albo nawet zablokować zapis.

W ThisWorkbook:

Minimalna wersja – tylko komunikat:
```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    MsgBox "Uwaga! Sprawdź dane przed zapisaniem.", vbExclamation, "Kontrola przed zapisem"
End Sub
```
Wersja z potwierdzeniem (możesz zablokować zapis):
```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim decyzja As VbMsgBoxResult

    decyzja = MsgBox("Czy na pewno chcesz zapisać ten dokument?" & vbCrLf & _
                     "Sprawdź wersję, datę i numer partii.", _
                     vbQuestion + vbYesNo, "Kontrola przed zapisem")

    If decyzja = vbNo Then
        Cancel = True ' <-- zatrzymuje zapis
    End If
End Sub
```
