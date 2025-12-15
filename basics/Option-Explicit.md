# Option Explicit
[Explained](https://learn.microsoft.com/pl-pl/dotnet/visual-basic/language-reference/statements/option-explicit-statement)

When Option Explicit On or Option Explicit appears in a file, you must explicitly declare all variables by using the Dim or ReDim statements. If you try to use an undeclared variable name, an error occurs at compile time. The Option Explicit Off statement allows implicit declaration of variables.
If used, the Option Explicit statement must appear in a file before any other source code statements.
Setting Option Explicit to Off is generally not a good practice. You could misspell a variable name in one or more locations, which would cause unexpected results when the program is run.

``` vba
Option Explicit

Sub xyz()
  'do something
End Sub
```
