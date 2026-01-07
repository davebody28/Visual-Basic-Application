# chrW() function


## Polish characters table

|Number|Uppercase Letter|Number|Lowercase Letter|
|:-:|:-:|:-:|:-:|
|211|Ó|243|ó|
|260|Ą|261|ą|
|262|Ć|263|ć|
|280|Ę|281|ę|
|321|Ł|322|ł|
|323|Ń|324|ń|
|377|Ź|378|ź|
|379|Ż|380|ż|
|216|Ø|248|Ø|

### VBA Code

``` vba
Option Explicit

Sub charw_loop()
    Dim x As Integer
    
    Cells(1, 1).Select
    For x = 1 To 1000
        ActiveCell.Value = x
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ChrW(x)
        ActiveCell.Offset(1, -1).Activate
    Next x
End Sub
```
