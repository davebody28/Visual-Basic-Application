# Sub email template

v1
``` vba
Sub Email_Reminder()
    Dim output0, output1, output2 As String
    Dim x, y As Double

    items = Range(Cells(3, 2), Cells(3, 2).End(xlDown)).count + 2
    
    'przygotowanie output jako html tabela
    'pobranie nagłówków
    output0 = "<style>table, th, td{border: 1px solid black; border-collapse: collapse; margin: 3px; padding: 3px; text-align: center; font-family: Exo; font-size: 14px;} th {background: blue;}</style><table><thead><tr><th>" & Cells(2, 1).Value & "</th><th>" & Cells(2, 2).Value & "</th><th>" & Cells(2, 3).Value & "</th><th>" & Cells(2, 13).Value & "</th><th>" & Cells(2, 14).Value & "</th><th>" & Cells(2, 15).Value & "</th>" _
        & "<th>" & Cells(2, 16).Value & "</th><th>" & Cells(2, 17).Value & "</th><th>" & Cells(2, 18).Value & "</th><th>" & Cells(2, 19).Value & "</th><th>" & Cells(2, 20).Value & "</th><th>" & Cells(2, 21).Value & "</th>" _
        & "<th>" & Cells(2, 22).Value & "</th><th>" & Cells(2, 23).Value & "</th><th>" & Cells(2, 24).Value & "</th><th>" & Cells(2, 25).Value & "</th><th>" & Cells(2, 26).Value & "</th><th>" & Cells(2, 27).Value & "</th><th>" & Cells(2, 28).Value & "</th><th>" & Cells(2, 29).Value & "</th><th>" & Cells(2, 30).Value & "</th></thead><tbody>"
    
    'pobieranie zawartości tabeli
    For y = 3 To items
        output1 = output1 + "<tr><td>" & Cells(y, 1).Value & "</td><td>" & Cells(y, 2).Value & "</td><td>" & Cells(y, 3).Value & "</td><td>" & Cells(y, 13).Value & "</td><td>" & Cells(y, 14).Value & "</td><td>" & Cells(y, 15).Value & "</td><td>" _
            & Cells(y, 16).Value & "</td><td>" & Cells(y, 17).Value & "</td><td>" & Cells(y, 18).Value & "</td><td>" & Cells(y, 19).Value & "</td><td>" & Cells(y, 20).Value & "</td><td>" & Cells(y, 21).Value & "</td><td>" _
            & Cells(y, 22).Value & "</td><td>" & Cells(y, 23).Value & "</td><td>" & Cells(y, 24).Value & "</td><td>" & Cells(y, 25).Value & "</td><td>" & Cells(y, 26).Value & "</td><td>" & Cells(y, 27).Value & "</td><td>" & Cells(y, 28).Value & "</td><td>" & Cells(y, 29).Value & "</td><td>" & Cells(y, 30).Value & "</td></tr>"
    Next y
    
    'zamknięcie tagu
    output2 = "</tbody></table>"
    
    
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim odbiorcy As String
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    For x = 1 To 30
        odbiorcy = odbiorcy & Worksheets("Email").Cells(x, 1).Value
    Next x
    
    With outlookMail
        .To = odbiorcy 'Arkusz "Email" A1:A30
        .Subject = "Listy kontrolne " & ActiveSheet.Name & " - brakuj" & ChrW(261) & "ce wpisy"
        .HTMLBody = "<font style=""font-family: Exo;"" />" _
            & "Witam," _
            & "<br><br>" _
            & "poni" & ChrW(380) & "ej lista z listami kontrolnymi, do uzupe" & ChrW(322) & "nienia. <br>" _
            & "Prosz" & ChrW(281) & " o uzupe" & ChrW(322) & "nienie brakuj" & ChrW(261) & "cych list. <br><br>" _
            & output0 & output1 & output2
        .ReadReceiptRequested = False
        .display
    End With
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub
```

v2
``` vba
Option Explicit

Sub SendeMail(actualrow As Integer, newproject As Boolean)
    
    Dim newprojecttext As String

    If newproject = True Then
        newprojecttext = "<tr><td style=""color:orange; font-weight:bold;""><b>Nowy projekt:</td><td style=""color:orange; font-weight:bold;"">" & Cells(actualrow, 5) & "</b></td></tr>"
    ElseIf newproject = False Then
        newprojecttext = "<tr><td>Nowy projekt:</td><td>" & Cells(actualrow, 5) & "</td></tr>"
    End If

    Dim olApp As Object    'Outlook.Application
    Dim olMail As Object    'Outlook.MailItem
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .to = ""
        .CC = ""
        .bcc = ""
        .Subject = "Zlecenie detal: " & Cells(actualrow, 1)
        .HTMLBody = "<font style=""font-family: Exo;"" />" _
        & "Zlecenie przygotowania przyrzadow oraz programów pomiarowych: " _
        & "<br><br><style>table, th, td{border: 1px solid black; border-collapse: collapse; margin: 3px; padding: 3px; text-align: left; font-family: Exo; font-size: 14px;} th {background: blue;}</style><table><tbody>" _
        & "<tr><td>Detal:</td><td>" & Cells(actualrow, 1) & "</td></tr>" _
        & "<tr><td>Index:</td><td>" & Cells(actualrow, 2) & "</td></tr>" _
        & "<tr><td>Maszyna:</td><td>" & Cells(actualrow, 3) & " " & Cells(actualrow, 4) & "</td></tr>" _
        & newprojecttext _
        & "<tr><td>Bezpapierowa produkcja:</td><td>" & Cells(actualrow, 6) & "</td></tr>" _
        & "<tr><td>Planowany start produkcji:</td><td>" & Cells(actualrow, 7) & "</td></tr>" _
        & "<tr><td>Termin przygotowania przyrządów:</td><td>" & Cells(actualrow, 8) & "</td></tr>" _
        & "<tr><td>Termin przygotowania programów:</td><td>" & Cells(actualrow, 9) & "</td></tr>" _
        & "<tr><td colspan=2 style=""height:20px;""></td></tr>" _
        & "<tr><td>Zlecajacy:</td><td>" & Cells(actualrow, 11) & " " & Cells(actualrow, 10) & "</td></tr>" _
        & "</tbody></table>"
        .Display
    End With
    
    'olMail.Send

End Sub
```
