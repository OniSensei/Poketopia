Attribute VB_Name = "modPokemon"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302
Public Sub AddPicture(RTB As RichTextBox, pic As String)
On Error Resume Next
    ' copy into the clipboard
    ' Copy the picture into the clipboard.
    
    Clipboard.Clear
    Clipboard.SetData LoadPicture(App.Path & "\Data Files\graphics\" & pic)
    ' paste into the RichTextBox control
    SendMessage RTB.hwnd, WM_PASTE, 0, 0
End Sub
Sub ReadText(ByVal FileName As String, textbox As VB.textbox)
Open FileName For Input As #1
textbox.text = Input$(LOF(1), #1)
Close #1
End Sub

Public Function ReadTextToString(ByVal FileName As String) As String
Open FileName For Input As #1
TextToString = Input$(LOF(1), #1)
Close #1
End Function

Sub BattleInfo(picturebox As Object)
picturebox.Visible = True
Dim tick As Long
Dim lastgettickcount As Long
Do While picturebox.Visible = True
tick = GetTickCount
If tick > 6000 Then
picturebox.Visible = False
Exit Sub
End If
Loop
End Sub

Public Sub PlayClick()
PlaySound ("Choose.wav")
End Sub
Sub WriteText(ByVal FileName As String, ByVal text As String)
Open FileName For Output As #1
Print #1, text
Close #1
End Sub

Public Sub CropImage(theImage As AlphaImgCtl, X As Long, Y As Long, _
                        Width As Long, Height As Long)
    
    ' the AlignCenter property should already be set to False for better runtime visual change
    
    With theImage
        .SetRedraw = False
        .AlignCenter = False
        .AutoSize = lvicNoAutoSize
        ' ensure X, Y are pixel scalemode
        .SetOffsets -X, -Y
        ' ensure X, Y, Width, Height are appropriate container scalemode
        '.move .Left + 16, .Top + 0, Width, Height
        .SetRedraw = True
    End With
        
End Sub

