<div align="center">

## Easily Replace a character in a string


</div>

### Description

Replaces a character in a string with another character. Pretty simple to understand. Uses a For...Next loop.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Berardi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-berardi.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-berardi-easily-replace-a-character-in-a-string__1-5980/archive/master.zip)





### Source Code

```
Function ReplaceCharacter(stringToChange$, charToReplace$, replaceWith$) As String
'Replaces a specified character in a string with another
'character that you specify
  Dim ln, n As Long
  Dim NextLetter As String
  Dim FinalString As String
  Dim txt, char, rep As String
  txt = stringToChange$ 'store all arguments in
  char = charToReplace$ 'new variables
  rep = replaceWith$
  ln = Len(txt)
  For n = 1 To ln Step 1
    NextLetter = Mid(txt, n, 1)
    If NextLetter = char Then
      NextLetter = rep
    End If
    FinalString = FinalString & NextLetter
  Next n
  Replace_Character = FinalString
End Function
```

