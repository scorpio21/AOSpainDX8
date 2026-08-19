Attribute VB_Name = "DrawPJenPicture"

Private Type tPJCuentas
    Body As Long
    Head As Long
    Casco As Long
    Shield As Long
    Weapon As Long
    Baned As Long
    Nombre As String
    LVL As Integer
    Clase As String
    Muerto As Integer
    GM As Long
    Active As Boolean
    
    ' Animacion
    FrameIndex As Integer
End Type

Private PJs(0 To 9) As tPJCuentas

Sub DibujaPJ(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Index As Integer)

    On Error GoTo ErrorHandler

    Dim iGrhIndex As Long

    If Grh.GrhIndex <= 0 Then
        LogError "DibujaPJ: Grh.GrhIndex INVALIDO=" & Grh.GrhIndex
        Exit Sub
    End If

    If Grh.GrhIndex > GrhCount Then
        LogError "DibujaPJ: Grh fuera de rango=" & Grh.GrhIndex
        Exit Sub
    End If

    If Not GrhData(Grh.GrhIndex).Active Then
        LogError "DibujaPJ: Grh INACTIVO=" & Grh.GrhIndex
        Exit Sub
    End If

    If Grh.FrameCounter <= 0 Then
        Grh.FrameCounter = 1
    End If

    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
        Grh.FrameCounter = 1
    End If

    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    If iGrhIndex <= 0 Or iGrhIndex > GrhCount Then
        LogError "DibujaPJ: Frame INVALIDO. Grh=" & Grh.GrhIndex & _
                 " Frame=" & Grh.FrameCounter & _
                 " iGrh=" & iGrhIndex
        Exit Sub
    End If

    If Not GrhData(iGrhIndex).Active Then
        LogError "DibujaPJ: Frame INACTIVO=" & iGrhIndex
        Exit Sub
    End If

    LogError "DibujaPJ: DIBUJANDO Index=" & Index & _
             " Grh=" & Grh.GrhIndex & _
             " Frame=" & Grh.FrameCounter & _
             " iGrh=" & iGrhIndex & _
             " X=" & X & _
             " Y=" & Y & _
             " W=" & GrhData(iGrhIndex).pixelWidth & _
             " H=" & GrhData(iGrhIndex).pixelHeight

    Call GrhRenderToHdc( _
        iGrhIndex, _
        frmCuent.PJ(Index).Hdc, _
        X - (GrhData(iGrhIndex).pixelWidth \ 2), _
        Y, _
        True)

    Exit Sub

ErrorHandler:

    LogError "DibujaPJ ERROR: Index=" & Index & _
             " Grh=" & Grh.GrhIndex & _
             " Frame=" & Grh.FrameCounter & _
             " iGrh=" & iGrhIndex & _
             " Err=" & Err.Number & _
             " Desc=" & Err.Description

End Sub
Sub RenderizarPJsCuentas()

    On Error Resume Next

    Dim i As Integer
    Dim rColor As Long

    Static ContadorAnim As Integer

    rColor = RGB(255, 215, 0)

    ' Reducimos la velocidad de la animación
    ContadorAnim = ContadorAnim + 1

    For i = 0 To 9

        frmCuent.PJ(i).Cls

            ' Solo avanzamos el frame cada 4 ciclos
            If ContadorAnim >= 4 Then
                PJs(i).FrameIndex = PJs(i).FrameIndex + 2
            End If
            If PJs(i).Active Then
            Call ActualizarDibujoPJ(i)
            

            If frmCuent.Nombre(i).Caption = PJClickeado And PJClickeado <> "" Then
                frmCuent.PJ(i).Line (0, 0)- _
                    (frmCuent.PJ(i).ScaleWidth - 1, frmCuent.PJ(i).ScaleHeight - 1), _
                    rColor, B

                frmCuent.PJ(i).Line (1, 1)- _
                    (frmCuent.PJ(i).ScaleWidth - 2, frmCuent.PJ(i).ScaleHeight - 2), _
                    rColor, B
            End If

        Else
            frmCuent.CP(i).Visible = True
        End If

    Next i

    If ContadorAnim >= 4 Then
        ContadorAnim = 0
    End If

End Sub

Public Sub LimpiarPJsCuentas()
    ' [CODE] - Limpieza centralizada de slots (10 slots)
    Dim i As Integer
    For i = 0 To 9
        PJs(i).Active = False
        PJs(i).Nombre = ""
        PJs(i).LVL = 0
        PJs(i).Body = 0
        PJs(i).Head = 0

        ' Limpiamos labels y pictures del formulario
        frmCuent.Nombre(i).Caption = "Nada"
        frmCuent.Nombre(i).Visible = False
        frmCuent.Label2(i).Caption = "Nivel: 0"
        frmCuent.Label2(i).Visible = False
        frmCuent.CP(i).Visible = True
        frmCuent.GM(i).Visible = False
        frmCuent.PJ(i).Cls
    Next i

    PJClickeado = ""
End Sub

Private Sub ActualizarDibujoPJ(ByVal Index As Integer)
    
    LogError "PJ TAMAÑO Index=" & Index & _
         " Width=" & frmCuent.PJ(Index).ScaleWidth & _
         " Height=" & frmCuent.PJ(Index).ScaleHeight
    
    Dim Body As Long, Head As Long, Casco As Long, Shield As Long, Weapon As Long
    Dim Muerto As Integer, Baned As Long
    Dim Grh As Grh
    Dim Pos As Integer
    Dim YBody As Long, YYY As Integer, XBody As Long, BBody As Long, YHead As Long
    
    Body = PJs(Index).Body
    Head = PJs(Index).Head
    Casco = PJs(Index).Casco
    Shield = PJs(Index).Shield
    Weapon = PJs(Index).Weapon
    Muerto = PJs(Index).Muerto
    Baned = PJs(Index).Baned
    
    XBody = 40
    YBody = 26
    BBody = 40
    

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
        PJs(Index).FrameIndex = PJs(Index).FrameIndex + 1

    If PJs(Index).FrameIndex > 6 Then
        PJs(Index).FrameIndex = 1
    End If
        Grh.FrameCounter = ((PJs(Index).FrameIndex - 1) Mod _
                            GrhData(Grh.GrhIndex).NumFrames) + 1

        Call DibujaPJ(Grh, XBody, YBody, Index)

    End If
End If

    If Muerto = 0 And Body > 0 And Body <= UBound(BodyData) Then
        YYY = BodyData(Body).HeadOffset.Y
    End If
    If Muerto = 1 Then YYY = -9

    Dim BodyGrhIndex As Long
Dim BodyFrameIndex As Long
Dim BodyHeight As Long

BodyGrhIndex = Grh.GrhIndex
BodyFrameIndex = Grh.FrameCounter

If BodyGrhIndex > 0 And BodyGrhIndex <= GrhCount Then
    If BodyFrameIndex > 0 And BodyFrameIndex <= GrhData(BodyGrhIndex).NumFrames Then
        
        BodyFrameIndex = GrhData(BodyGrhIndex).Frames(BodyFrameIndex)
        
        If BodyFrameIndex > 0 And BodyFrameIndex <= GrhCount Then
            BodyHeight = GrhData(BodyFrameIndex).pixelHeight
        End If
        
    End If
End If

Pos = YYY + BodyHeight
    YHead = Pos + 8
    ' Cabeza
    If Head > 0 And Head <= UBound(HeadData) Then
        Grh = HeadData(Head).Head(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = 1
            If Baned = 1 Then
                Call dibujaban(Index, vbBlack)
                Call dibujaban(Index, vbRed)
            End If
            Call DibujaPJ(Grh, BBody, YHead, Index)
            Debug.Print Pos + 8
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
            Grh.FrameCounter = ((PJs(Index).FrameIndex - 1) Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            Call DibujaPJ(Grh, XBody, YBody, Index)
        End If
    End If

    ' Escudo
    If Shield > 0 And Shield <> 2 And Shield <= UBound(ShieldAnimData) Then
        Grh = ShieldAnimData(Shield).ShieldWalk(3)
        If Grh.GrhIndex > 0 Then
            Grh.FrameCounter = ((PJs(Index).FrameIndex - 1) Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            Call DibujaPJ(Grh, XBody + 4, BBody - 13, Index)
        End If
    End If
End Sub

Public Sub DibujarTodo(ByVal Index As Integer, Body As Long, Head As Long, Casco As Long, Shield As Long, Weapon As Long, Baned As Long, Nombre As String, LVL As Integer, Clase As String, Muerto As Integer, GM As Long)

    With PJs(Index)
        .Body = Body
        .Head = Head
        .Casco = Casco
        .Shield = Shield
        .Weapon = Weapon
        .Baned = Baned
        .Nombre = Nombre
        .LVL = LVL
        .Clase = Clase
        .Muerto = Muerto
        .GM = GM
        .Active = True
        .FrameIndex = 1
    End With

    ' Actualizamos labels del formulario
frmCuent.Nombre(Index).Caption = Nombre
frmCuent.CP(Index).Visible = False
frmCuent.Nombre(Index).Visible = True
frmCuent.Label2(Index).Visible = True

    If GM = 1 Then
        frmCuent.GM(Index).Visible = True
    Else
        frmCuent.GM(Index).Visible = False
    End If
    
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
