Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMainGame.Font = Font
    frmMainGame.FontSize = Size - 4
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal text As String, color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, 0)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, x + 1, y + 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, x, y, text, Len(text))
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim color As Long
    If inBattle Then Exit Sub
    If Player(Index).notVisible And Index <> MyIndex Then Exit Sub
     If FlashLight = True Then
If Player(Index).x > Player(MyIndex).x + 3 Or Player(Index).x < Player(MyIndex).x - 3 Or Player(Index).y > Player(MyIndex).y + 3 Or Player(Index).y < Player(MyIndex).y - 3 Then Exit Sub
End If
   
      If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(GetPlayerName(Index))))
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - 32
    Else
        ' Determine location for text
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(GetPlayerName(Index))))
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight) / 4 + 28
    End If

        Select Case GetPlayerAccess(Index)
            Case 0
                color = QBColor(White)
                 ' Draw name
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(Index), color)
            Case 1
                color = QBColor(Green)
                Call DrawText(TexthDC, TextX - 15, TextY, "[CM]" & GetPlayerName(Index), color)
            Case 2
                color = QBColor(BrightBlue)
                Call DrawText(TexthDC, TextX - 20, TextY, "[MOD]" & GetPlayerName(Index), color)
            Case 3
                color = QBColor(Yellow)
                Call DrawText(TexthDC, TextX - 30, TextY, "[ADMIN]" & GetPlayerName(Index), color)
            Case 4
                color = QBColor(BrightRed)
                Call DrawText(TexthDC, TextX - 20, TextY, "[DEV]" & GetPlayerName(Index), color)
        End Select
      
    
End Sub

Public Sub DrawNpcName(ByVal numnpc As Long)
 Dim TextX As Long
    Dim TextY As Long
    Dim color As Long
    If inBattle Then Exit Sub
   If FlashLight = True Then
If MapNpc(numnpc).x > Player(MyIndex).x + 3 Or MapNpc(numnpc).x < Player(MyIndex).x - 3 Or MapNpc(numnpc).y > Player(MyIndex).y + 3 Or MapNpc(numnpc).y < Player(MyIndex).y - 3 Then Exit Sub
End If
color = QBColor(White)


     If NPC(MapNpc(numnpc).num).Sprite < 1 Or NPC(MapNpc(numnpc).num).Sprite > NumCharacters Then
        TextX = ConvertMapX(MapNpc(numnpc).x * PIC_X) + MapNpc(numnpc).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(NPC(MapNpc(numnpc).num).Name)))
        TextY = ConvertMapY(MapNpc(numnpc).y * PIC_Y) + MapNpc(numnpc).YOffset - 64
    Else
        ' Determine location for text
        TextX = ConvertMapX(MapNpc(numnpc).x * PIC_X) + MapNpc(numnpc).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(NPC(MapNpc(numnpc).num).Name)))
        TextY = ConvertMapY(MapNpc(numnpc).y * PIC_Y) + MapNpc(numnpc).YOffset - (DDSD_Character(NPC(MapNpc(numnpc).num).Sprite).lHeight) / 4 + 55
    End If

Call DrawText(TexthDC, TextX, TextY, NPC(MapNpc(numnpc).num).Name, color)
End Sub

Public Function BltMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long

    If frmEditor_Map.optAttribs.Value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    With map.Tile(x, y)
                        tX = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, tX, tY, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, tX, tY, "I", QBColor(White)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, tX, tY, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, tX, tY, "K", QBColor(White)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, tX, tY, "O", QBColor(White)
                            Case TILE_TYPE_RESOURCE
                                DrawText TexthDC, tX, tY, "O", QBColor(Green)
                            Case TILE_TYPE_DOOR
                                DrawText TexthDC, tX, tY, "D", QBColor(Brown)
                            Case TILE_TYPE_NPCSPAWN
                                DrawText TexthDC, tX, tY + 32, "S", QBColor(Yellow)
                            Case TILE_TYPE_SHOP
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                            Case TILE_TYPE_BATTLE
                                DrawText TexthDC, tX, tY, "B", QBColor(Blue)
                            Case TILE_TYPE_HEAL
                                DrawText TexthDC, tX, tY, "H", QBColor(Green)
                            Case TILE_TYPE_SPAWN
                                DrawText TexthDC, tX, tY, "S", QBColor(Pink)
                            Case TILE_TYPE_STORAGE
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                                Case TILE_TYPE_BANK
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightBlue)
                                Case TILE_TYPE_GYMBLOCK
                                DrawText TexthDC, tX, tY, "G", QBColor(BrightRed)
                                Case TILE_TYPE_CUSTOMSCRIPT
                                DrawText TexthDC, tX, tY, "CS", QBColor(Yellow)
                        End Select
                    End With
                End If
            Next
        Next
    End If

End Function

Sub BltActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, time As Long

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMainGame.picScreen.width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + time Then
        Call DrawText(TexthDC, x, y, ActionMsg(Index).message, QBColor(ActionMsg(Index).color))
    Else
        ClearActionMsg Index
    End If

End Sub

Public Function getWidth(ByVal DC As Long, ByVal text As String) As Long
    getWidth = frmMainGame.TextWidth(text) \ 2
End Function

    Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim S As String

   
    
    S = vbNewLine & Msg
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text)
    frmChat.txtChat.SelColor = QBColor(color)
    frmChat.txtChat.SelText = S
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text) - 1
    
    'Evilbunnie's DrawnChat System
    ReOrderChat Msg, QBColor(color)
    
    ' Error handler
Exit Sub
End Sub


Public Sub AddBattleText(ByVal Msg As String, ByVal color As Integer)
    If color = Black Then
    color = White
    End If
    Dim S As String
    S = vbNewLine & Msg
    frmMainGame.txtBtlLog.SelStart = Len(frmMainGame.txtBtlLog.text)
    frmMainGame.txtBtlLog.SelColor = QBColor(color)
    frmMainGame.txtBtlLog.SelText = S
    frmMainGame.txtBtlLog.SelStart = Len(frmMainGame.txtBtlLog.text) - 1
End Sub

'Evilbunnie's DrawnChat system
Public Sub DrawChat()

End Sub

'Evilbunnie's DrawChat system
Public Sub ReOrderChat(ByVal nText As String, nColour As Long)

End Sub
