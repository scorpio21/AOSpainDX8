Attribute VB_Name = "CargasInit"
'==========================================================
' CargasInit.bas - Carga e inicializacion de datos DX8
' Separado de TileEngine.bas para mayor modularidad
'==========================================================
Option Explicit

' === Declaraciones API GDI32 - necesarias para DrawGrhtoHdc y GrhRenderToHdc ===
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY As Long = &HCC0020

' === Cache de bitmaps para GrhRenderToHdc (evita re-leer archivo cada frame) ===
Private GrhHdcCache_Count As Long
Private GrhHdcCache_FileNum() As Integer
Private GrhHdcCache_Picture() As Object

Public Sub DrawGrhtoHdc(ByVal hwnd As Long, ByVal Hdc As Long, ByVal GrhIdx As Integer, _
                        SourceRect As RECT, destRect As RECT)
    On Error GoTo ErrHandler
    
    Static callCount As Long
    callCount = callCount + 1
    
    If GrhIdx <= 0 Or GrhIdx > GrhCount Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL GrhIdx=" & GrhIdx & " GrhCount=" & GrhCount
        Exit Sub
    End If
    If Not GrhData(GrhIdx).Active Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL NotActive GrhIdx=" & GrhIdx
        Exit Sub
    End If
    
    Dim GrhIndex As Long
    GrhIndex = GrhData(GrhIdx).Frames(1)
    If GrhIndex <= 0 Or GrhIndex > GrhCount Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL FrameIdx=" & GrhIndex
        Exit Sub
    End If
    If Not GrhData(GrhIndex).Active Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL FrameNotActive idx=" & GrhIndex
        Exit Sub
    End If
    
    Dim FileNum As Integer
    Dim FileSize As Long
    Dim Buffer() As Byte
    Dim BmpFile As String
    
    FileNum = GrhData(GrhIndex).FileNum
    BmpFile = App.Path & "\Graficos\" & FileNum & ".bmp"
    
    If Dir(BmpFile) = "" Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL NoFile " & BmpFile
        Exit Sub
    End If
    
    Dim fNum As Integer
    fNum = FreeFile
    Open BmpFile For Binary Access Read As #fNum
    FileSize = LOF(fNum)
    If FileSize <= 0 Then Close #fNum: Exit Sub
    
    ReDim Buffer(0 To FileSize - 1)
    Get #fNum, , Buffer
    Close #fNum
    
    Dim BiWidth As Long
    Dim BiHeight As Long
    Dim BiBitCount As Integer
    
    CopyMemory BiWidth, Buffer(18), 4
    CopyMemory BiHeight, Buffer(22), 4
    CopyMemory BiBitCount, Buffer(28), 2
    
    Dim sx As Long, sy As Long, pw As Long, ph As Long
    sx = GrhData(GrhIndex).SX
    sy = GrhData(GrhIndex).SY
    pw = GrhData(GrhIndex).pixelWidth
    ph = GrhData(GrhIndex).pixelHeight
    
    If sx + pw > BiWidth Then pw = BiWidth - sx
    If sy + ph > BiHeight Then ph = BiHeight - sy
    If pw <= 0 Or ph <= 0 Then
        If callCount <= 5 Then LogError "DrawGrhtoHdc: FAIL BadDim pw=" & pw & " ph=" & ph
        Exit Sub
    End If
    
    Dim dstW As Long, dstH As Long
    dstW = destRect.Right - destRect.Left
    dstH = destRect.bottom - destRect.Top
    If dstW <= 0 Then dstW = pw
    If dstH <= 0 Then dstH = ph
    
    Dim PixelDataOffset As Long
    CopyMemory PixelDataOffset, Buffer(10), 4
    
    ' BITMAPINFOHEADER starts at offset 14 in BMP file, not offset 2
    ' BITMAPINFO = BITMAPINFOHEADER (40) + ColorTable (up to 1024 for 8-bit)
    Dim colorTableSize As Long
    If BiBitCount <= 8 Then
        colorTableSize = CLng(4) * (CLng(2) ^ BiBitCount)
    Else
        colorTableSize = 0
    End If
    
    Dim BmInfo() As Byte
    ReDim BmInfo(0 To 40 + colorTableSize - 1)
    CopyMemory BmInfo(0), Buffer(14), 40
    If colorTableSize > 0 Then
        CopyMemory BmInfo(40), Buffer(54), colorTableSize
    End If
    
    If callCount <= 5 Then
        LogError "DrawGrhtoHdc: OK GrhIdx=" & GrhIdx & " FileNum=" & FileNum & " sx=" & sx & " sy=" & sy & " pw=" & pw & " ph=" & ph & " dstW=" & dstW & " dstH=" & dstH & " BmpW=" & BiWidth & " BmpH=" & BiHeight & " Bits=" & BiBitCount & " hdc=" & Hdc
    End If
    
    Dim result As Long
    result = StretchDIBits(Hdc, destRect.Left, destRect.Top, dstW, dstH, _
                  sx, sy, pw, ph, _
                  ByVal VarPtr(Buffer(PixelDataOffset)), ByVal VarPtr(BmInfo(0)), 0, vbSrcCopy)
    
    If callCount <= 5 Then
        LogError "DrawGrhtoHdc: StretchDIBits result=" & result & " GrhIdx=" & GrhIdx
    End If
    
Exit Sub
ErrHandler:
    LogError "DrawGrhtoHdc: Err=" & Err.Number & " Desc=" & Err.Description & " GrhIdx=" & GrhIdx
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Loaders migrados (los datos .ind del cliente usan el formato historico:
'cabecera tCabecera + campos Integer/Int16; los arrays de destino son los del
'motor DX8, por eso los campos se leen en variables temporales).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub LoadGrhData()
'*****************************************************************
'Carga Graficos.ind: MiCabecera + 5 x Integer + registros Int16
'terminados en un indice <= 0. Pasa 1: halla el indice maximo para
'dimensionar GrhData. Pasa 2: lee los registros.
'*****************************************************************
On Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim maxGrh As Long
    Dim grhIdx As Integer
    Dim nf As Integer
    Dim fileNum As Integer
    Dim sx As Integer
    Dim sy As Integer
    Dim pw As Integer
    Dim ph As Integer
    Dim sp As Integer
    Dim fr As Integer
    Dim Handle As Integer

    Handle = FreeFile
    Open IniPath & "Graficos.ind" For Binary Access Read As #Handle

    'Pasa 1: determinar el indice de grh maximo
    Get #Handle, , MiCabecera
    For i = 1 To 5
        Get #Handle, , tempint
    Next i

    maxGrh = 0
    Do
        Get #Handle, , grhIdx
        If grhIdx <= 0 Then Exit Do
        If grhIdx > maxGrh Then maxGrh = grhIdx
        Get #Handle, , nf
        If nf > 1 Then
            For Frame = 1 To nf
                Get #Handle, , fr
            Next Frame
            Get #Handle, , sp
        Else
            Get #Handle, , fileNum
            Get #Handle, , sx
            Get #Handle, , sy
            Get #Handle, , pw
            Get #Handle, , ph
        End If
    Loop

    If maxGrh <= 0 Then
        Close #Handle
        Exit Sub
    End If

    ReDim GrhData(1 To maxGrh) As GrhData

    'Pasa 2: leer los registros
    Seek #Handle, 1
    Get #Handle, , MiCabecera
    For i = 1 To 5
        Get #Handle, , tempint
    Next i

    Do
        Get #Handle, , grhIdx
        If grhIdx <= 0 Then Exit Do
        GrhData(grhIdx).Active = True
        Get #Handle, , nf
        If nf < 1 Then nf = 1
        GrhData(grhIdx).NumFrames = nf
        ReDim GrhData(grhIdx).Frames(1 To nf)
        If nf > 1 Then
            For Frame = 1 To nf
                Get #Handle, , fr
                GrhData(grhIdx).Frames(Frame) = fr
            Next Frame
            Get #Handle, , sp
            GrhData(grhIdx).Speed = sp
            GrhData(grhIdx).pixelHeight = GrhData(GrhData(grhIdx).Frames(1)).pixelHeight
            GrhData(grhIdx).pixelWidth = GrhData(GrhData(grhIdx).Frames(1)).pixelWidth
            GrhData(grhIdx).TileWidth = GrhData(GrhData(grhIdx).Frames(1)).TileWidth
            GrhData(grhIdx).TileHeight = GrhData(GrhData(grhIdx).Frames(1)).TileHeight
        Else
            Get #Handle, , fileNum
            GrhData(grhIdx).FileNum = fileNum
            Get #Handle, , sx
            GrhData(grhIdx).SX = sx
            Get #Handle, , sy
            GrhData(grhIdx).SY = sy
            Get #Handle, , pw
            GrhData(grhIdx).pixelWidth = pw
            Get #Handle, , ph
            GrhData(grhIdx).pixelHeight = ph
            GrhData(grhIdx).TileWidth = pw / TilePixelHeight
            GrhData(grhIdx).TileHeight = ph / TilePixelWidth
            GrhData(grhIdx).Frames(1) = grhIdx
        End If
    Loop

    Close #Handle

    GrhCount = maxGrh
    LogError "LoadGrhData: maxGrh=" & maxGrh & " GrhCount=" & GrhCount
    If maxGrh <= 0 Then LogError "LoadGrhData: WARNING - no GRH entries!"
    
    ' Diagnostic: dump first 10 GRH records
    Dim diag As Long
    For diag = 1 To 10
        If GrhData(diag).Active Then
            LogError "GRH " & diag & ": Active=" & GrhData(diag).Active & " FileNum=" & GrhData(diag).FileNum & " SX=" & GrhData(diag).SX & " SY=" & GrhData(diag).SY & " PW=" & GrhData(diag).pixelWidth & " PH=" & GrhData(diag).pixelHeight & " Frames=" & GrhData(diag).NumFrames & " Speed=" & GrhData(diag).Speed
        Else
            LogError "GRH " & diag & ": INACTIVE"
        End If
    Next diag
End Sub

Sub CargarCuerpos()
'*****************************************************************
'Carga Personajes.ind: MiCabecera + NumCuerpos + registros
'tIndiceCuerpo (4 grh de caminata + offset de cabeza).
'*****************************************************************
On Error Resume Next
    Dim n As Integer
    Dim i As Long
    Dim j As Byte
    Dim MisCuerpos() As tIndiceCuerpo

    n = FreeFile
    Open IniPath & "Personajes.ind" For Binary Access Read As #n

    Get #n, , MiCabecera
    Get #n, , NumCuerpos

    ReDim BodyData(0 To NumCuerpos + 1) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

    LogError "CargarCuerpos: NumCuerpos=" & NumCuerpos

    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        For j = 1 To 4
            Call InitGrh(BodyData(i).Walk(j), MisCuerpos(i).Body(j), 0)
        Next j
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        If i <= 5 Then
            LogError "Body " & i & ": grhIdx=" & MisCuerpos(i).Body(1) & " HeadOffsetX=" & MisCuerpos(i).HeadOffsetX & " HeadOffsetY=" & MisCuerpos(i).HeadOffsetY
        End If
    Next i

    Close #n
End Sub

Sub CargarCabezas()
'*****************************************************************
'Carga Cabezas.ind: MiCabecera + NumHeads + registros tIndiceCabeza.
'Ademas deriva heads() (textura/coords) para DrawHead desde el
'primer grh de cada cabeza.
'*****************************************************************
On Error Resume Next
    Dim n As Integer
    Dim i As Long
    Dim j As Byte
    Dim grhIndex As Long
    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open IniPath & "Cabezas.ind" For Binary Access Read As #n

    Get #n, , MiCabecera
    Get #n, , NumHeads

    ReDim HeadData(0 To NumHeads + 1) As HeadData
    ReDim Miscabezas(0 To NumHeads + 1) As tIndiceCabeza
    ReDim heads(0 To NumHeads + 1) As tHead

    LogError "CargarCabezas: NumHeads=" & NumHeads

    For i = 1 To NumHeads
        Get #n, , Miscabezas(i)
        For j = 1 To 4
            Call InitGrh(HeadData(i).Head(j), Miscabezas(i).Head(j), 0)
        Next j
        grhIndex = HeadData(i).Head(1).GrhIndex
        If grhIndex > 0 And grhIndex <= UBound(GrhData) Then
            heads(i).Texture = GrhData(grhIndex).FileNum
            heads(i).startX = GrhData(grhIndex).SX
            heads(i).startY = GrhData(grhIndex).SY
        End If
        If i <= 5 Then
            LogError "Head " & i & ": grhIdx=" & grhIndex & " FileNum=" & heads(i).Texture & " startX=" & heads(i).startX & " startY=" & heads(i).startY
        End If
    Next i

    Close #n
End Sub

Sub CargarCascos()
'*****************************************************************
'Carga Cascos.ind: MiCabecera + NumCascos + registros tIndiceCabeza.
'Ademas deriva Cascos() (textura/coords) para DrawHead.
'*****************************************************************
On Error Resume Next
    Dim n As Integer
    Dim i As Long
    Dim j As Byte
    Dim grhIndex As Long
    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open IniPath & "Cascos.ind" For Binary Access Read As #n

    Get #n, , MiCabecera
    Get #n, , NumCascos

    ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
    ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
    ReDim Cascos(0 To NumCascos + 1) As tHead

    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        For j = 1 To 4
            Call InitGrh(CascoAnimData(i).Head(j), Miscabezas(i).Head(j), 0)
        Next j
        grhIndex = CascoAnimData(i).Head(1).GrhIndex
        If grhIndex > 0 And grhIndex <= UBound(GrhData) Then
            Cascos(i).Texture = GrhData(grhIndex).FileNum
            Cascos(i).startX = GrhData(grhIndex).SX
            Cascos(i).startY = GrhData(grhIndex).SY
        End If
    Next i

    Close #n
End Sub

Sub CargarFxs()
'*****************************************************************
'Carga Fxs.ind: MiCabecera + NumFxs + registros tIndiceFx.
'*****************************************************************
On Error Resume Next
    Dim n As Integer
    Dim i As Long

    n = FreeFile
    Open IniPath & "Fxs.ind" For Binary Access Read As #n

    Get #n, , MiCabecera
    Get #n, , NumFxs

    ReDim FxData(1 To NumFxs) As tIndiceFx

    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i

    Close #n
End Sub

Sub CargarAtaques()
'*****************************************************************
'Carga Ataques.ind si existe (no presente en este cliente): MiCabecera
'+ NumAtaques + registros tIndiceAtaque. Si el archivo falta, deja
'AtaqueData dimensionado a 0 para que MakeChar con Ataque=0 no falle.
'*****************************************************************
On Error Resume Next
    Dim n As Integer
    Dim i As Long
    Dim j As Byte
    Dim MisAtaques() As tIndiceAtaque

    n = FreeFile
    Open IniPath & "Ataques.ind" For Binary Access Read As #n

    Get #n, , MiCabecera
    Get #n, , NumAtaques

    If NumAtaques < 1 Then NumAtaques = 0
    ReDim AtaqueData(0 To NumAtaques) As AtaqueAnimData
    ReDim MisAtaques(0 To NumAtaques) As tIndiceAtaque

    For i = 1 To NumAtaques
        Get #n, , MisAtaques(i)
        If MisAtaques(i).Body(1) Then
            For j = 1 To 4
                Call InitGrh(AtaqueData(i).AtaqueWalk(j), MisAtaques(i).Body(j), 0)
            Next j
            AtaqueData(i).HeadOffset.X = MisAtaques(i).HeadOffsetX
            AtaqueData(i).HeadOffset.Y = MisAtaques(i).HeadOffsetY
        End If
    Next i

    Close #n
End Sub

Sub LoadMiniMap()
'*****************************************************************
'El cliente no dispone de minimap.dat; los colores del minimapa se
'obtienen de GrhData.MiniMap_color (que el render dibuja directamente).
'*****************************************************************
On Error Resume Next
End Sub

Sub CargarParticulas()
'*****************************************************************
'El cliente no dispone de particulas.ini; las particulas se crean en
'tiempo de ejecucion via Particle_Group_Create / Char_Particle_Group_Create.
'*****************************************************************
On Error Resume Next
End Sub

Private Function GetGrhPictureForHdc(ByVal FileNum As Integer) As Object
    Dim i As Long
    For i = 1 To GrhHdcCache_Count
        If GrhHdcCache_FileNum(i) = FileNum Then
            Set GetGrhPictureForHdc = GrhHdcCache_Picture(i)
            Exit Function
        End If
    Next i

    On Error Resume Next
    Dim Pic As Object
    Set Pic = LoadPicture(DirGraficos & FileNum & ".bmp")
    On Error GoTo 0
    If Pic Is Nothing Then Exit Function

    GrhHdcCache_Count = GrhHdcCache_Count + 1
    ReDim Preserve GrhHdcCache_FileNum(1 To GrhHdcCache_Count)
    ReDim Preserve GrhHdcCache_Picture(1 To GrhHdcCache_Count)
    GrhHdcCache_FileNum(GrhHdcCache_Count) = FileNum
    Set GrhHdcCache_Picture(GrhHdcCache_Count) = Pic
    Set GetGrhPictureForHdc = Pic
End Function

Public Sub GrhRenderToHdc(ByVal GrhIndex As Long, ByVal DesthDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal Transparent As Boolean)
    On Error Resume Next

    Dim FileNum As Integer
    Dim sX As Integer, sY As Integer
    Dim pixelWidth As Integer, pixelHeight As Integer
    Dim Pic As Object
    Dim MemDC As Long
    Dim OldBmp As Long

    If GrhIndex <= 0 Then Exit Sub

    ' Obtener datos del Grh
    With GrhData(GrhIndex)
        sX = .SX
        sY = .SY
        pixelWidth = .pixelWidth
        pixelHeight = .pixelHeight
        FileNum = .FileNum
    End With

    ' Validaciones de seguridad
    If FileNum <= 0 Then Exit Sub

    Set Pic = GetGrhPictureForHdc(FileNum)
    If Pic Is Nothing Then Exit Sub

    ' Crear un DC compatible en memoria y seleccionar el bitmap cargado
    MemDC = CreateCompatibleDC(DesthDC)
    OldBmp = SelectObject(MemDC, Pic.Handle)

    Call BitBlt(DesthDC, X, Y, pixelWidth, pixelHeight, MemDC, sX, sY, SRCCOPY)

    ' Restaurar y liberar el DC temporal (el bitmap en si queda vivo en cache)
    Call SelectObject(MemDC, OldBmp)
    Call DeleteDC(MemDC)
End Sub
