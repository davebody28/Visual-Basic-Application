# Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
with some if statements and parameters passthru in call function

Templates to use:
``` vba
actualrow = Target.Row
entrydate = Format(Now(), "DD.MM.YY HH:MM")
entryperson = Workbooks.Application.UserName
If ActiveSheet.Name = "Lista" Then
If Target.Row > 3 Then
If Target.Column <= 9 Then
Call Info1.Show
Call eMailSender.SendeMail(actualrow, newproject)
```

Example code
``` vba
Option Explicit

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Dim actualrow As Integer, cellcounter As Integer
    Dim entrydate As String, entryperson As String
    
    actualrow = Target.Row
    cellcounter = 0
    entrydate = Format(Now(), "DD.MM.YY HH:MM")
    entryperson = Workbooks.Application.UserName
    
    Dim planedproductionstart As Date, setequipmentdate As Date, setprogramsdate As Date, requestingdate As Date
    Dim checkcounter As Integer, newproject As Boolean
    
    If ActiveSheet.Name = "Lista" Then
        If Target.Row > 3 Then
            planedproductionstart = Cells(actualrow, 7).Value
            setequipmentdate = Cells(actualrow, 8).Value
            setprogramsdate = Cells(actualrow, 9).Value
            requestingdate = Cells(actualrow, 10).Value
            checkcounter = 0
            
            If Target.Column <= 9 Then
                If Not Cells(actualrow, 1) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 2) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 3) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 4) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 5) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Cells(actualrow, 5) = "TAK" Then
                    newproject = True
                Else
                    newproject = False
                End If
                If Not Cells(actualrow, 6) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 7) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 8) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                If Not Cells(actualrow, 9) = 0 Then
                    cellcounter = cellcounter + 1
                End If
                
                
                If cellcounter = 9 Then
                    Cells(actualrow, 10) = entrydate
                    Cells(actualrow, 11) = entryperson
                    
                    If planedproductionstart - setequipmentdate < 0.16 Then '0.16 dnia to 4h
                        Call Info1.Show
                    Else
                        checkcounter = checkcounter + 1
                    End If
                    If planedproductionstart - setprogramsdate < 0.33 Then '0.33 dnia to 8h
                        Call Info2.Show
                    Else
                        checkcounter = checkcounter + 1
                    End If
                    If checkcounter = 2 Then
                        Call eMailSender.SendeMail(actualrow, newproject)
                    End If
                End If
                
            End If
        End If
    End If
End Sub
```
