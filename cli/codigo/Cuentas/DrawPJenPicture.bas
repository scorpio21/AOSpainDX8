Attribute VB_Name = "DrawPJenPicture"

Private Type tPJCuentas
    Body As Long
    Head As Long
    Casco As Long
    Shield As Long
    Weapon As Long
    Baned As Long
    nombre As String
    LVL As Integer
    Clase As String
    Muerto As Integer
    Active As Boolean
    
    ' Animacion
    FrameIndex As Integer
End Type

Private PJs(0 To 9) As tPJCuentas

Sub DibujaPJ(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Index As Integer)
On Error Resume Next
Dim iGrhIndex As Long
If Grh.GrhIndex <= 0 Then Exit Sub

' Aseguramos que el frame sea valido
If Grh.FrameCounter <= 0 Then Grh.FrameCounter = 1
If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then Grh.FrameCounter = 1

iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

Call GrhRenderToHdc(iGrhIndex, frmCuent.PJ(Index).hDC, X, Y, True)

frmCuent.PJ(Index).Refresh

End Sub

Sub RenderizarPJsCuentas()
    ' [CODE] - Renderizado centralizado para AoSpain
    On Error Resume Next
    Dim i As Integer
    Dim rColor As Long

    ' Color de resaltado (Dorado para seleccion)
    rColor = RGB(255, 215, 0)

    ' Recorremos los 10 slots (Estandarizado Fase 2)
    For i = 0 To 9
        ' Limpiamos siempre el fondo para evitar rastro
        frmCuent.PJ(i).Cls

        If PJs(i).Active Then
            ' Incrementamos animacion
            PJs(i).FrameIndex = PJs(i).FrameIndex + 1

            ' Dibujamos el personaje (Cuerpo + Cabeza + Equipo)
            Call ActualizarDibujoPJ(i)

            ' Si este personaje es el seleccionado, borde dorado
            If frmCuent.nombre(i).Caption = PJClickeado And PJClickeado <> "" Then
                frmCuent.PJ(i).Line (0, 0)-(frmCuent.PJ(i).ScaleWidth - 1, frmCuent.PJ(i).ScaleHeight - 1), rColor, B
                frmCuent.PJ(i).Line (1, 1)-(frmCuent.PJ(i).ScaleWidth - 2, frmCuent.PJ(i).ScaleHeight - 2), rColor, B
            End If
        Else
            ' Slot vacio: mostrar "Crear Personaje"
            frmCuent.CP(i).Visible = True
        End If
    Next i
End Sub

Public Sub LimpiarPJsCuentas()
    ' [CODE] - Limpieza centralizada de slots (10 slots)
    Dim i As Integer
    For i = 0 To 9
        PJs(i).Active = False
        PJs(i).nombre = ""
        PJs(i).LVL = 0
        PJs(i).Body = 0
        PJs(i).Head = 0

        ' Limpiamos labels y pictures del formulario
        frmCuent.nombre(i).Caption = "Nada"
        frmCuent.nombre(i).Visible = False
        frmCuent.Label2(i).Caption = "Nivel: 0"
        frmCuent.Label2(i).Visible = False
        frmCuent.CP(i).Visible = True
        frmCuent.GM(i).Visible = False
        frmCuent.PJ(i).Cls
    Next i

    PJClickeado = ""
End Sub

Private Sub ActualizarDibujoPJ(ByVal Index As Integer)
    Dim Body As Long, Head As Long, Casco As Long, Shield As Long, Weapon As Long
    Dim Muerto As Integer, Baned As Long
    Dim Grh As Grh
    Dim Pos As Integer
    Dim YBody As Long, YYY As Integer, XBody As Long, BBody As Long
    
    Body = PJs(Index).Body
    Head = PJs(Index).Head
    Casco = PJs(Index).Casco
    Shield = PJs(Index).Shield
    Weapon = PJs(Index).Weapon
    Muerto = PJs(Index).Muerto
    Baned = PJs(Index).Baned
    
    XBody = 29
    YBody = 20
    BBody = 35

    If Muerto = 1 Then
        Body = 8
        Head = 500
        XBody = 28
        YBody = 38
        BBody = 35
    End If

    ' Cuerpo (heading 3 = Sur, orientacion por defecto)
    If Body > 0 And Body <= UBound(BodyData) Then
        Grh = BodyData(Body).Walk(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = (PJs(Index).FrameIndex Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            Call DibujaPJ(Grh, XBody, YBody, Index)
        End If
    End If

    If Muerto = 0 And Body > 0 And Body <= UBound(BodyData) Then
        YYY = BodyData(Body).HeadOffset.Y
    End If
    If Muerto = 1 Then YYY = -9

    Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)).pixelHeight
    
    ' Cabeza
    If Head > 0 And Head <= UBound(HeadData) Then
        Grh = HeadData(Head).Head(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = 1
            If Baned = 1 Then
                Call dibujaban(Index, vbBlack)
                Call dibujaban(Index, vbRed)
            End If
            Call DibujaPJ(Grh, BBody, Pos + 2, Index)
        End If
    End If
        
    ' Casco
    If Casco > 0 And Casco <> 2 And Casco <= UBound(CascoAnimData) Then
        Grh = CascoAnimData(Casco).Head(3)
        If Grh.GrhIndex > 0 Then Call DibujaPJ(Grh, BBody, Pos + 2, Index)
    End If

    ' Arma
    If Weapon > 0 And Weapon <> 2 And Weapon <= UBound(WeaponAnimData) Then
        Grh = WeaponAnimData(Weapon).WeaponWalk(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = (PJs(Index).FrameIndex Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            Call DibujaPJ(Grh, XBody, YBody, Index)
        End If
    End If

    ' Escudo
    If Shield > 0 And Shield <> 2 And Shield <= UBound(ShieldAnimData) Then
        Grh = ShieldAnimData(Shield).ShieldWalk(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = (PJs(Index).FrameIndex Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            Call DibujaPJ(Grh, XBody + 4, BBody - 13, Index)
        End If
    End If
End Sub

Public Sub DibujarTodo(ByVal Index As Integer, Body As Long, Head As Long, Casco As Long, Shield As Long, Weapon As Long, Baned As Long, nombre As String, LVL As Integer, Clase As String, Muerto As Integer)

    With PJs(Index)
        .Body = Body
        .Head = Head
        .Casco = Casco
        .Shield = Shield
        .Weapon = Weapon
        .Baned = Baned
        .nombre = nombre
        .LVL = LVL
        .Clase = Clase
        .Muerto = Muerto
        .Active = True
        .FrameIndex = 1
    End With

    ' Actualizamos labels del formulario
    frmCuent.nombre(Index).Caption = nombre
    frmCuent.CP(Index).Visible = False
    frmCuent.nombre(Index).Visible = True
    frmCuent.Label2(Index).Visible = True
    
    If LVL > 50 Then
        frmCuent.Label2(Index).Caption = "Nivel: 50 + " & LVL - 50
    Else
        frmCuent.Label2(Index).Caption = "Nivel: " & LVL
    End If
    
    ' Dibujamos el primer frame inmediatamente
    ActualizarDibujoPJ Index

End Sub

Sub dibujaban(Index As Integer, Color As Long)
    ' Dibuja un rectangulo de color sobre la cabeza (baneado)
    frmCuent.PJ(Index).Line (27, 0)-(55, 50), Color, BF
End Sub