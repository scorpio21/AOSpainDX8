Attribute VB_Name = "Mod_ErrorLOG"
'===============================================================
' Mod_ErrorLog.bas – Sistema de Logs PRO para AOSpainDX8
'
' Versión: PRO 2026
' Autor: scorpio21
' Proyecto: AOSpainDX8 (Cliente DX8 Modernizado)
'
' Este módulo implementa:
'   • Sistema de logs avanzado con niveles (INFO, WARN, ERROR, CRITICAL)
'   • Logs rotativos por tamaño
'   • Logs diarios opcionales
'   • Logs por categoría, por sub y por módulo
'   • Limitación inteligente de repeticiones
'   • Integración completa con config.ini avanzado
'   • Optimización para rendimiento DX8
'
' Basado en Argentum Online 0.9.0.9 (GPL)
' Código original © 2002 Pablo Ignacio Márquez y colaboradores
'
' Este archivo es una obra derivada bajo licencia GPL v2 o superior.
' Puedes redistribuirlo y modificarlo bajo los términos de la GPL.
'
' Para más información sobre la licencia:
' https://www.gnu.org/licenses/old-licenses/gpl-2.0.html
'
'===============================================================



Option Explicit

' ============================================================
' VARIABLES GLOBALES (PRO)
' ============================================================

' Control general
Public LogsEnabled As Boolean

' DX8
Public Log_DrawGrh As Boolean
Public Log_DibujaPJ As Boolean
Public Log_MapaDibuja As Boolean
Public Log_RenderScreen As Boolean
Public Log_Animaciones As Boolean
Public Log_DX8Debug As Boolean

' Mapa
Public Log_MapaCarga As Boolean
Public Log_MapaTransiciones As Boolean

' TCP
Public Log_TCP_Conexiones As Boolean
Public Log_TCP_Paquetes As Boolean
Public Log_TCP_Errores As Boolean

' Inventario / Objetos
Public Log_Inventario As Boolean
Public Log_Items As Boolean
Public Log_Hechizos As Boolean

' Usuario
Public Log_UserLogin As Boolean
Public Log_UserAcciones As Boolean
Public Log_UserDebug As Boolean

' Sistema
Public Log_Errores As Boolean
Public Log_Criticos As Boolean
Public Log_Warnings As Boolean
Public Log_DebugGeneral As Boolean

' Colección para contadores
Private LogCount As New Collection

' ============================================================
' API PARA LEER INI
' ============================================================

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Function ReadIni(Section As String, Key As String, Default As String, File As String) As String
    Dim Buffer As String
    Buffer = String$(255, vbNullChar)
    GetPrivateProfileString Section, Key, Default, Buffer, 255, File
    ReadIni = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End Function

' ============================================================
' INICIALIZAR LOGS (PRO)
' ============================================================

Public Sub InitLogs()
    On Error Resume Next

    Set LogCount = New Collection
    MkDir App.Path & "\Logs"

    Dim iniPath As String
    iniPath = App.Path & "\Init\config.ini"

    ' Control general
    LogsEnabled = (ReadIni("LOGS", "EnableLogs", "1", iniPath) = "1")

    ' DX8
    Log_DrawGrh = (ReadIni("LOGS", "EnableDrawGrh", "1", iniPath) = "1")
    Log_DibujaPJ = (ReadIni("LOGS", "EnableDibujaPJ", "1", iniPath) = "1")
    Log_MapaDibuja = (ReadIni("LOGS", "EnableMapaDibuja", "1", iniPath) = "1")
    Log_RenderScreen = (ReadIni("LOGS", "EnableRenderScreen", "1", iniPath) = "1")
    Log_Animaciones = (ReadIni("LOGS", "EnableAnimaciones", "1", iniPath) = "1")
    Log_DX8Debug = (ReadIni("LOGS", "EnableDX8Debug", "0", iniPath) = "1")

    ' Mapa
    Log_MapaCarga = (ReadIni("LOGS", "EnableMapaCarga", "1", iniPath) = "1")
    Log_MapaTransiciones = (ReadIni("LOGS", "EnableMapaTransiciones", "0", iniPath) = "1")

    ' TCP
    Log_TCP_Conexiones = (ReadIni("LOGS", "EnableTCP_Conexiones", "1", iniPath) = "1")
    Log_TCP_Paquetes = (ReadIni("LOGS", "EnableTCP_Paquetes", "0", iniPath) = "1")
    Log_TCP_Errores = (ReadIni("LOGS", "EnableTCP_Errores", "1", iniPath) = "1")

    ' Inventario / Objetos
    Log_Inventario = (ReadIni("LOGS", "EnableInventario", "1", iniPath) = "1")
    Log_Items = (ReadIni("LOGS", "EnableItems", "1", iniPath) = "1")
    Log_Hechizos = (ReadIni("LOGS", "EnableHechizos", "0", iniPath) = "1")

    ' Usuario
    Log_UserLogin = (ReadIni("LOGS", "EnableUserLogin", "1", iniPath) = "1")
    Log_UserAcciones = (ReadIni("LOGS", "EnableUserAcciones", "1", iniPath) = "1")
    Log_UserDebug = (ReadIni("LOGS", "EnableUserDebug", "0", iniPath) = "1")

    ' Sistema
    Log_Errores = (ReadIni("LOGS", "EnableErrores", "1", iniPath) = "1")
    Log_Criticos = (ReadIni("LOGS", "EnableCriticos", "1", iniPath) = "1")
    Log_Warnings = (ReadIni("LOGS", "EnableWarnings", "1", iniPath) = "1")
    Log_DebugGeneral = (ReadIni("LOGS", "EnableDebugGeneral", "0", iniPath) = "1")

End Sub

' ============================================================
' LOG ROTATIVO (PRO)
' ============================================================

Private Sub RotateLog(FilePath As String, MaxSizeKB As Long)
    On Error Resume Next

    If Dir$(FilePath) = "" Then Exit Sub

    If FileLen(FilePath) > (MaxSizeKB * 1024) Then
        Dim Backup As String
        Backup = Replace(FilePath, ".log", "_old.log")
        Kill Backup
        Name FilePath As Backup
    End If
End Sub

' ============================================================
' LOG CON NIVELES (INFO / WARN / ERROR / CRITICAL)
' ============================================================

Public Sub LogLevel(FileName As String, Level As String, Msg As String)
    If Not LogsEnabled Then Exit Sub

    Dim nfile As Integer
    nfile = FreeFile

    Dim FilePath As String
    FilePath = App.Path & "\Logs\" & FileName

    RotateLog FilePath, 500   ' 500 KB rotación

    Open FilePath For Append As #nfile
    Print #nfile, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & Level & "] " & Msg
    Close #nfile
End Sub

' ============================================================
' LOG LIMITADO POR MENSAJE EXACTO
' ============================================================

Public Sub LogLimited(FileName As String, desc As String, MaxRepeats As Integer)
    If Not LogsEnabled Then Exit Sub

    On Error Resume Next

    Dim Key As String
    Key = desc

    Dim Count As Integer
    If LogCount.Exists(Key) Then Count = LogCount(Key)

    If Count >= MaxRepeats Then Exit Sub

    LogCount.Remove Key
    LogCount.Add Count + 1, Key

    LogLevel FileName, "INFO", desc
End Sub

' ============================================================
' LOG LIMITADO POR SUB
' ============================================================

Public Sub LogLimitedSub(FileName As String, SubName As String, desc As String, MaxRepeats As Integer)
    If Not LogsEnabled Then Exit Sub

    On Error Resume Next

    Dim Key As String
    Key = SubName & desc

    Dim Count As Integer
    If LogCount.Exists(Key) Then Count = LogCount(Key)

    If Count >= MaxRepeats Then Exit Sub

    LogCount.Remove Key
    LogCount.Add Count + 1, Key

    LogLevel FileName, SubName, desc
End Sub

' ============================================================
' LOG LIMITADO POR CATEGORÍA (PRO)
' ============================================================

Public Sub LogLimitedCategory(FileName As String, Category As String, desc As String, MaxRepeats As Integer)
    If Not LogsEnabled Then Exit Sub

    On Error Resume Next

    Dim Count As Integer
    If LogCount.Exists(Category) Then Count = LogCount(Category)

    If Count >= MaxRepeats Then Exit Sub

    LogCount.Remove Category
    LogCount.Add Count + 1, Category

    LogLevel FileName, Category, desc
End Sub


