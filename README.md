<div align="center">

## Auto\-Select Text on Focus


</div>

### Description

Nice feature to help users edit textboxes. The code selects all the existing text when the users focus on the textbox control. This will definitely make your application more user friendly. Please Vote!!!
 
### More Info
 
Textbox Name

Copy this sub on your code and then use the textbox name as a parameter when you call the sub. When I use it, I usually call the SUB on the GotFocus event.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Delio Castillo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/delio-castillo.md)
**Level**          |Beginner
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/delio-castillo-auto-select-text-on-focus__1-11257/archive/master.zip)





### Source Code

```
Public Sub Select_Text(TextBoxName As Variant)
  TextBoxName.SelStart = 0
  TextBoxName.SelLength = Len(TextBoxName.Text)
End Sub
```

