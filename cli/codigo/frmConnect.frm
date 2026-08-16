VERSION 5.00
Begin VB.Form frmConnect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   471.094
   ScaleMode       =   0  'User
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   450
      ItemData        =   "frmConnect.frx":639D3
      Left            =   10560
      List            =   "frmConnect.frx":639DA
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   9240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox DescTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "AOSpain Primario"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   8640
      TabIndex        =   5
      Text            =   "localhost"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   8640
      TabIndex        =   4
      Text            =   "7666"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   248
      IMEMode         =   3  'DISABLE
      Left            =   5055
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4700
      Width           =   1725
   End
   Begin VB.TextBox NameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   248
      Left            =   5055
      TabIndex        =   1
      Top             =   3401
      Width           =   1725
   End
   Begin VB.Image EliminarCuenta 
      Height          =   495
      Left            =   8160
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Image RecuperarCuenta 
      Height          =   255
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   1125
   End
   Begin VB.Image Conectar 
      Height          =   405
      Index           =   1
      Left            =   6240
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recordar cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5055
      TabIndex        =   9
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image CrearPersonaje 
      Height          =   315
      Left            =   4320
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   1005
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   191
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private bLoadingRecordar As Boolean

' --- CARGA DE SERVIDORES (Opcional si se usa lista) ---
Public Sub CargarLst()
    Dim i As Integer
    lst_servers.Clear
    For i = 1 To UBound(ServersLst)
        lst_servers.AddItem ServersLst(i).desc
    Next i
End Sub

' --- ACTIVACION DEL FORMULARIO ---
Private Sub Form_Activate()
    Dim nDirectorio As String
    
    ' Configuracion de IP/Puerto segun seleccion o configuraci�n inicial
    If CurServer <> 0 Then
        IPTxt = ServersLst(CurServer).Ip
        PortTxt = ServersLst(CurServer).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If

    Call CargarLst
    DescTxt.Text = ServersLst(CurServer).Desc

    ' Aseguramos que el directorio Web existe para logs locales si fuera necesario
    nDirectorio = Dir(App.Path & "\Web", vbDirectory)
    If nDirectorio <> "Web" Then MkDir (App.Path & "\Web")
End Sub

' --- CERRAR JUEGO CON ESC ---
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        DeinitTileEngine
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
    End If
End Sub

' --- ATAJO PARA VER IP/PUERTO (DEBUG) ---
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyI And Shift = vbCtrlMask Then
        PortTxt.Visible = Not PortTxt.Visible
        IPTxt.Visible = Not IPTxt.Visible
        KeyCode = 0
    End If
End Sub

' --- CARGA INICIAL ---
Private Sub Form_Load()
    EngineRun = False
    
    ' Seteamos el puerto por defecto
    PortTxt.Text = Config_Inicio.Puerto
    
    ' Recordar Cuenta By scorpio
    Dim sFile As String
    sFile = App.Path & "\LIBS\Recordar.dat"
    
    bLoadingRecordar = True
    
    If Dir(sFile) <> "" Then
        If Val(GetVar(sFile, "Recordar", "Check")) = 1 Then
            NameTxt.Text = GetVar(sFile, "Recordar", "Nombre")
            PasswordTxt.Text = DecryptINI$(GetVar(sFile, "Recordar", "Password"), "aopass")
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    Else
        Check1.Value = 0
    End If
    
    bLoadingRecordar = False
End Sub

' --- BOTON: CREAR CUENTA ---
Private Sub CrearPersonaje_Click()
    Call PlayWaveDS(SND_CLICK)
    
    ' [FIX] Conectarse al servidor antes de mostrar el form de creacion.
    ' El form envia NACCNT via SendData -- si no hay conexion activa
    ' da error 24038 (Invalid socket descriptor).
    If Not frmMain.Socket1.Connected Then
        EstadoLogin = CrearAccount
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        DoEvents
    End If
    
    frmCrearAccount.Show vbModal, Me
End Sub

' --- BOTON PRINCIPAL: CONECTAR A CUENTA ---
Private Sub Conectar_Click(Index As Integer)
    ' Solo procesamos el bot�n Index 1 (Conectar)
    If Index <> 1 Then Exit Sub

    Call PlayWaveDS(SND_CLICK)

    ' Validaci�n de campos
    If NameTxt.Text = "" Or PasswordTxt.Text = "" Then
        MsgBox "Ingrese Nombre de Cuenta y Contrase�a.", vbExclamation
        Exit Sub
    End If

    ' Limpieza de socket previo
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    ' Guardamos datos de sesi�n
    nombrecuent = NameTxt.Text
    passcuent = PasswordTxt.Text
    
    ' Si "Recordar cuenta" est� activo, guardamos los datos
    If Check1.Value = 1 Then
        Dim sFile As String
        sFile = App.Path & "\LIBS\Recordar.dat"
        Call WriteVar(sFile, "Recordar", "Nombre", nombrecuent)
        Call WriteVar(sFile, "Recordar", "Password", EncryptINI$(passcuent, "aopass"))
        Call WriteVar(sFile, "Recordar", "Check", "1")
    End If
    
    ' El servidor guarda la password con MD5, por lo que enviamos el hash
    UserPassword = MD5String(passcuent)
    MD5HushYo = UserPassword

    ' Intentamos Loguear la Cuenta
    EstadoLogin = LoginAccount
    frmMain.Socket1.HostAddress = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
End Sub

' --- BOTON: RECUPERAR CUENTA ---
Private Sub RecuperarCuenta_Click()
    Call PlayWaveDS(SND_CLICK)
    frmRecuperarCuenta.Show vbModal, Me
End Sub

' --- BOTON: ELIMINAR CUENTA (SEGURIDAD CRITICA) ---
Private Sub EliminarCuenta_Click()
    Call PlayWaveDS(SND_CLICK)
    
    Dim sAccount As String
    sAccount = InputBox("Ingrese el nombre de la CUENTA que desea eliminar permanentemente:", "ELIMINAR CUENTA")
    
    If sAccount = "" Then Exit Sub
    
    Dim sPass As String
    sPass = InputBox("Para borrar la cuenta " & sAccount & " escriba su CONTRASE�A:", "VALIDACION DE SEGURIDAD")
    If sPass = "" Then Exit Sub
    
    Dim sMail As String
    sMail = InputBox("Escriba el CORREO registrado en la cuenta " & sAccount & ":", "VALIDACION DE SEGURIDAD")
    If sMail = "" Then Exit Sub
    
    Dim sRespuesta As String
    sRespuesta = InputBox("Escriba la RESPUESTA SECRETA de la cuenta " & sAccount & ":", "VALIDACION DE SEGURIDAD")
    If sRespuesta = "" Then Exit Sub
    
    ' Si el socket no est� conectado, intentamos conectar para enviar el paquete
    If Not frmMain.Socket1.Connected Then
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        
        ' Esperamos un momento a que conecte (max 3 segundos)
        Dim lWait As Long
        lWait = GetTickCount
        Do While Not frmMain.Socket1.Connected
            DoEvents
            If GetTickCount - lWait > 3000 Then
                MsgBox "No se ha podido establecer conexion con el servidor.", vbCritical
                Exit Sub
            End If
        Loop
    End If
    
    ' Enviamos la solicitud: BRCU [Cuenta],[Nombre],[Pass],[Mail],[Respuesta]
    ' El servidor valida: campo Nombre contra GetVar("Nombre") (vacio), la pass
    ' comparada con MD5String(bPass) contra la guardada, mail y respuesta en claro.
    Call SendData("BRCU" & sAccount & ",," & sPass & "," & sMail & "," & sRespuesta)
End Sub

' --- CHECK: RECORDAR CUENTA ---
Private Sub Check1_Click()
    If bLoadingRecordar Then Exit Sub
    
    Dim sFile As String
    sFile = App.Path & "\LIBS\Recordar.dat"
    
    If Check1.Value = 0 Then
        ' Desactivar recordar - borrar archivo
        If Dir(sFile) <> "" Then Kill sFile
        NameTxt.Text = ""
        PasswordTxt.Text = ""
    Else
        ' Activar recordar - guardar lo que haya en los campos
        Call WriteVar(sFile, "Recordar", "Nombre", NameTxt.Text)
        Call WriteVar(sFile, "Recordar", "Password", EncryptINI$(PasswordTxt.Text, "aopass"))
        Call WriteVar(sFile, "Recordar", "Check", "1")
    End If
End Sub

' --- LOGICA DE LISTA DE SERVIDORES (SI SE USA) ---
Private Sub lst_servers_Click()
    CurServer = lst_servers.ListIndex + 1
    DescTxt = ServersLst(CurServer).desc
    IPTxt = ServersLst(CurServer).Ip
    PortTxt = ServersLst(CurServer).Puerto
End Sub
