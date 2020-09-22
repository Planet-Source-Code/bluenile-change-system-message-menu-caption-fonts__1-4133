<div align="center">

## Change System \(Message, Menu, Caption\) Fonts


</div>

### Description

' Heres a very simple code to change the system

' NONCLIENTMETRICS like the the window title font,

' the message font,menu font using VB. You can also change

' other elements like status font etc

' in your window only or all the open windows

' like PLUS! or display settings (appearance)

' also it is possible to underline, strikethru fonts in

' your window with this code. This code is very useful

' if you are coding a multi-lingual software.

' For more info and more free code send e-mail.

' code by - NILESH P KURHADE

' email - bluenile5@hotmail.com
 
### More Info
 
ADD A COMBO BOX

Add a Combo box.

Changes the Message box font and Windows Caption Font (Title Font).

None that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[bluenile](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bluenile.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bluenile-change-system-message-menu-caption-fonts__1-4133/archive/master.zip)

### API Declarations

```
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As NONCLIENTMETRICS, ByVal fuWinIni As Long) As Long
```


### Source Code

```
Private Sub Combo1_Click()
Dim ncm As NONCLIENTMETRICS 'NONCLIENTMETRICS to change
Dim Orincm As NONCLIENTMETRICS 'NONCLIENTMETRICS to replace original
Dim Returned As Long
Dim i As Integer
ncm.cbSize = Len(ncm)
Returned = SystemParametersInfo(41, 0, ncm, 0) 'get the system NONCLIENTMETRICS
Orincm = ncm 'store the value of system NONCLIENTMETRICS to use later
'now to change the font name
'other functions can be used to change the font name
'but for simplicity i have used asc() & mid()
For i = 1 To Len(Combo1.Text) 'use ncm.lfMenuFont.lfFacename(i) to change menu font
  ncm.lfMessageFont.lfFaceName(i) = Asc(Mid(Combo1.Text, i, 1))
  ncm.lfCaptionFont.lfFaceName(i) = Asc(Mid(Combo1.Text, i, 1))
Next i
ncm.lfMessageFont.lfFaceName(i) = 0 'add null at the end of font name
ncm.lfCaptionFont.lfFaceName(i) = 0
Returned = SystemParametersInfo(42, 0, ncm, &H1 Or &H2) 'remove &H2 if you don't want to affect all the open windows
MsgBox "Message & Caption Font Changed to " & Combo1.Text, vbOKOnly, "NILESH"
Returned = SystemParametersInfo(42, 0, Orincm, &H1 Or &H2) 'replace original font
MsgBox "Message & Caption Font Replaced to " & StrConv(Orincm.lfCaptionFont.lfFaceName, vbUnicode), vbOKOnly, "NILESH"
End Sub
Private Sub Form_Load()
' Heres a very simple code to change the system
' NONCLIENTMETRICS like the the window title font,
' the message font,menu font using VB. You can also change
' other elements like status font etc
' in your window only or all the open windows
' like PLUS! or display settings (appearance)
' also it is possible to underline, strikethru fonts in
' your window with this code. This code is very useful
' if you are coding a multi-lingual software.
' For more info and more free code send e-mail.
' code by - NILESH P KURHADE
' email - bluenile5@hotmail.com
Dim i As Integer
Show
' to flood the combo box with first 10 fonts
For i = 1 To 10 ' or use For i = 1 To Screen.FontCount to flood all the fonts in your pc
  Combo1.AddItem Screen.Fonts(i)
Next i
End Sub
```

