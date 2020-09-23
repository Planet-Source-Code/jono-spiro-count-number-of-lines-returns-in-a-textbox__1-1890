<div align="center">

## Count number of lines/returns in a textbox


</div>

### Description

Simply counts the number of lines in a textbox (the textbox should be multiline=true, otherwise it is pretty useless). Put this in a module so it can be reused.
 
### More Info
 
USAGE--countLines(the textbox)

EXAMPLE--countLines(text1)

Returns the number of lines in a textbox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jono Spiro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jono-spiro.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jono-spiro-count-number-of-lines-returns-in-a-textbox__1-1890/archive/master.zip)





### Source Code

```
Public Function countLines(textBox As textBox) As Long
 Dim A%, B$
 A% = 1
 B$ = textBox.text
 Do While InStr(B$, Chr$(13))
  A% = A% + 1
  B$ = Mid$(B$, InStr(B$, Chr$(13)) + 1)
 Loop
 countLines = CStr(A%)
End Function
```

