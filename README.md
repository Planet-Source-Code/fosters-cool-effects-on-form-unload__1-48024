<div align="center">

## Cool effects on form unload


</div>

### Description

Paste into a fresh project's form1. on unloading the form, you can specify combination effects on how the form will disappear. This code is *not* mine, I found it on allapi.net. Simple effects that lend a touch of professionalism to an application. hope u like.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fosters](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fosters.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fosters-cool-effects-on-form-unload__1-48024/archive/master.zip)





### Source Code

```
Const AW_HOR_POSITIVE = &H1 'Animates the window from left to right. This flag can be used with roll or slide animation.
Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.
Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Const AW_ACTIVATE = &H20000 'Activates the window.
Const AW_SLIDE = &H40000 'Uses slide animation. By default, roll animation is used.
Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Private Sub Form_Load()
  'Credit to: http://www.allapi.net/
  Me.AutoRedraw = True
  Me.Print "Unload me"
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'Animate the window
  AnimateWindow Me.hwnd, 300, AW_BLEND Or AW_HIDE
  'Unload our form completely
  Set Form1 = Nothing
End Sub
```

