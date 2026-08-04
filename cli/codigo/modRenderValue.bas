Attribute VB_Name = "modRenderValue"
Option Explicit

' GS-Zone AO - Sistema de render de valores en pantalla (dano, oro, etc.)

Public Enum RVType
    eGold = 0
    eDamage = 1
    eHeal = 2
    eMana = 3
    eExp = 4
End Enum

Public Const RENDER_TIME As Integer = 120

Type RVList
    Activated As Boolean
    RenderVal As Long
    ColorRGB As Long
    TimeRendered As Integer
    Downloading As Byte
    RenderType As RVType
End Type

Public Sub Draw(ByVal X As Integer, ByVal Y As Integer, ByVal PixelX As Integer, ByVal PixelY As Integer)
    Dim tRV As RVList
    With MapData(X, Y).RenderValue
        If .Activated Then
            If .TimeRendered < RENDER_TIME Then
                .TimeRendered = .TimeRendered + 1
                If (.TimeRendered / 2) > 0 Then
                    .Downloading = (.TimeRendered / 8)
                End If
            Else
                .Activated = False
                .TimeRendered = 0
                Exit Sub
            End If
        End If
    End With
End Sub

Public Sub SetRenderValue(ByVal X As Integer, ByVal Y As Integer, ByVal rValue As Long, ByVal eMode As Byte)
    With MapData(X, Y).RenderValue
        .Activated = True
        .RenderType = eMode
        .RenderVal = rValue
        .TimeRendered = 0
        .Downloading = 0
        .ColorRGB = ModifyColour(eMode)
    End With
End Sub

Private Function ModifyColour(ByVal RenderType As RVType) As Long
    Select Case RenderType
        Case RVType.eGold
            ModifyColour = RGB(1, 240, 255)
        Case RVType.eDamage
            ModifyColour = RGB(255, 0, 0)
        Case RVType.eHeal
            ModifyColour = RGB(0, 255, 0)
        Case RVType.eMana
            ModifyColour = RGB(0, 0, 255)
        Case Else
            ModifyColour = RGB(255, 255, 255)
    End Select
End Function
