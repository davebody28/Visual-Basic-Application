# Option Base statement
[Explained](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement)

Because the default base is 0, the Option Base statement is never required. If used, the statement must appear in a module before any procedures. Option Base can appear only once in a module and must precede array declarations that include dimensions.
The To clause in the Dim, Private, Public, ReDim, and Static statements provides a more flexible way to control the range of an array's subscripts. However, if you don't explicitly set the lower bound with a To clause, you can use Option Base to change the default lower bound to 1. The base of an array created with the ParamArray keyword is zero; Option Base does not affect ParamArray (or the Array function, when qualified with the name of its type library, for example VBA.Array).
The Option Base statement only affects the lower bound of arrays in the module where the statement is located.

Example:
``` vba
Option Base 1

Sub xyz()
  'do something
End Sub
```

``` vba
Sub xyz()
  Option Base 1
  'do something
End Sub
```
