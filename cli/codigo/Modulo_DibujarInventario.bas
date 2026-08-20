Attribute VB_Name = "DibujarInventario"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
' Modulo de inventario (DX8)
' Reemplazado el blit DirectDraw7 por DrawGrhtoHdc (stub DX8 en TileEngine.bas).
' La logica de seleccion (ItemElegido, ItemClick, ClicEnItemElegido) se conserva intacta.

Option Explicit

Public Const XCantItems = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Function ClicEnItemElegido(X As Integer, Y As Integer) As Boolean
bInvMod = True
mx = X \ 32 + 1
my = Y \ 32 + 1
If ItemElegido = 0 Or FLAGORO Then
    ClicEnItemElegido = False
Else
    ClicEnItemElegido = (UserInventory(ItemElegido).OBJIndex > 0) And (ItemElegido = (mx + (my - 1) * 5) + OffsetDelInv)
End If
End Function

Sub ItemClick(X As Integer, Y As Integer)
Dim lPreItem As Long

bInvMod = False
mx = X \ 32 + 1
my = Y \ 32 + 1

lPreItem = (mx + (my - 1) * 5) + OffsetDelInv

If lPreItem <= MAX_INVENTORY_SLOTS Then _
If UserInventory(lPreItem).GrhIndex > 0 Then _
    ItemElegido = lPreItem: bInvMod = True
End Sub

Sub DibujarInv()
Dim iX As Integer
Dim rSource As RECT
Dim rDest As RECT
Dim itemCount As Integer

frmMain.picInv.Cls

With rDest
    .Top = 0
    .Left = 0
    .Right = 32
    .Bottom = 32
End With

itemCount = 0

For iX = OffsetDelInv + 1 To UBound(UserInventory)
    If UserInventory(iX).GrhIndex > 0 Then
    
        itemCount = itemCount + 1
        Dim Grh As Long
        Grh = UserInventory(iX).GrhIndex
        If Grh > 0 And Grh <= GrhCount And GrhData(Grh).Active Then
            Dim frameGrh As Long
            frameGrh = GrhData(Grh).Frames(1)
            If LogsEnabled And Log_Inventario Then
                LogError "INV: Slot=" & iX & _
                    " ObjGrh=" & UserInventory(iX).GrhIndex & _
                    " FrameGrh=" & frameGrh & _
                    " sx=" & GrhData(frameGrh).sx & _
                    " sy=" & GrhData(frameGrh).sy & _
                    " FileNum=" & GrhData(frameGrh).FileNum
            End If
            If frameGrh > 0 And frameGrh <= GrhCount And GrhData(frameGrh).Active Then
                With rSource
                    .Left = GrhData(frameGrh).sx
                    .Top = GrhData(frameGrh).sy
                    .Right = .Left + GrhData(frameGrh).pixelWidth
                    .Bottom = .Top + GrhData(frameGrh).pixelHeight
                End With
            End If
        End If
    '===========================
    ' LOG DE INVENTARIO / OBJETOS
    '===========================
    If LogsEnabled And Log_Inventario Then

    Dim indLine As String
    indLine = "Grh" & frameGrh & "=" & _
              GrhData(frameGrh).NumFrames & "-" & _
              GrhData(frameGrh).FileNum & "-" & _
              GrhData(frameGrh).sx & "-" & _
              GrhData(frameGrh).sy & "-" & _
              GrhData(frameGrh).pixelWidth & "-" & _
              GrhData(frameGrh).pixelHeight

    Dim logMsg As String
    logMsg = "Slot=" & iX & _
             " Grh=" & UserInventory(iX).GrhIndex & _
             " Cant=" & UserInventory(iX).amount & _
             " Equipped=" & UserInventory(iX).Equipped & _
             " FrameGrh=" & frameGrh & _
             " sx=" & rSource.Left & _
             " sy=" & rSource.Top & _
             " W=" & (rSource.Right - rSource.Left) & _
             " H=" & (rSource.Bottom - rSource.Top) & _
             " | IND=" & indLine

    Call LogLimited("Inventario.log", logMsg, 20)
End If
    '===========================
        If LogsEnabled And Log_Inventario Then
            LogError "INV: DRAWING Slot=" & iX & _
                     " Source=(" & rSource.Left & "," & rSource.Top & ")-(" & rSource.Right & "," & rSource.Bottom & ")" & _
                     " Dest=(" & rDest.Left & "," & rDest.Top & ")-(" & rDest.Right & "," & rDest.Bottom & ")"
        End If
        Call DrawGrhtoHdc(frmMain.picInv.hwnd, frmMain.picInv.Hdc, UserInventory(iX).GrhIndex, rSource, rDest)

        frmMain.picInv.CurrentX = rDest.Left
        frmMain.picInv.CurrentY = rDest.Top
        frmMain.picInv.Print UserInventory(iX).amount

        If UserInventory(iX).Equipped = 1 Then
            frmMain.picInv.CurrentX = rDest.Left + 20
            frmMain.picInv.CurrentY = rDest.Top + 20
            frmMain.picInv.ForeColor = vbYellow
            frmMain.picInv.Print "+"
        End If
    End If

    rDest.Left = rDest.Left + 32
    rDest.Right = rDest.Right + 32
    If rDest.Left >= 160 Then
        rDest.Left = 0
        rDest.Right = 32
        rDest.Top = rDest.Top + 32
        rDest.Bottom = rDest.Bottom + 32
    End If

Next iX

frmMain.picInv.Refresh

bInvMod = False

If ItemElegido = 0 Then _
    Call ItemClick(2, 2)

End Sub
