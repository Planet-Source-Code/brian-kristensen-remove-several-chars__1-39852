<div align="center">

## Remove several chars\!


</div>

### Description

It removes a set of chars from a string, by using ParamArray, it is very simple! It was submitted as a alternative to another submission : http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39842&lngWId=1
 
### More Info
 
Text, char1, char2 .... etc.

It is simple and fast

It returns the string without the chars that is defined.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Kristensen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-kristensen.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-kristensen-remove-several-chars__1-39852/archive/master.zip)





### Source Code

```
Public Sub TestIt()
 Dim sText As String
 sText = InputBox("Please write something you want to remove 'e;s;n;t' from!", "Remove Char test", "It doesent work!")
 MsgBox RemoveChar(sText, "e", "s", "n", "t")
End Sub
Public Function RemoveChar(ByVal sText As String, ParamArray sChar()) As String
Dim lngdo As Long
 For lngdo = LBound(sChar) To UBound(sChar)
  sText = Replace(sText, sChar(lngdo), "")
 Next lngdo
 RemoveChar = sText
 MsgBox "Removed: " & Join(sChar, ";")
End Function
```

