Attribute VB_Name = "MoD_MIDI"
Option Explicit

' ==========================================================
' MoDuLo_MIDI.bas - Reemplazado por clsSoundEngine (DX8)
' Las llamadas a DirectMusicPerformance DX7 se delegan al
' motor de audio DX8 (Sound As clsSoundEngine)
' ==========================================================

Public CurMidi As String      ' Musica actual
Public MIdi_Inicio As String  ' Musica de inicio

' -------------------------------------------------------
' CargarMIDI - Carga y prepara un archivo de musica
' En DX8: establece la proxima musica en clsSoundEngine
' -------------------------------------------------------
Public Sub CargarMIDI(Archivo As String)
    On Error Resume Next
    If Sound Is Nothing Then Exit Sub
    If LenB(Archivo) = 0 Then Exit Sub
    
    'DX8: el motor extrae "<id>.mp3" desde los recursos, por lo que solo usamos el nombre base
    Dim f As String
    Dim i As Long
    f = Archivo
    i = InStrRev(f, "\")
    If i > 0 Then f = Mid$(f, i + 1)
    i = InStrRev(f, ".")
    If i > 0 Then f = Left$(f, i - 1)
    
    CurMidi = f
    Sound.NextMusic = f
    Sound.Fading = 350
End Sub

' -------------------------------------------------------
' Play_Midi - Reproduce la musica cargada
' -------------------------------------------------------
Public Sub Play_Midi()
    On Error Resume Next
    If Sound Is Nothing Then Exit Sub
    Sound.Music_Play
End Sub

' -------------------------------------------------------
' Stop_Midi - Detiene la musica
' -------------------------------------------------------
Public Sub Stop_Midi()
    On Error Resume Next
    If Sound Is Nothing Then Exit Sub
    Sound.Music_Stop
End Sub

' -------------------------------------------------------
' Sonando - Devuelve True si hay musica reproduciendose
' -------------------------------------------------------
Function Sonando() As Boolean
    On Error Resume Next
    If Sound Is Nothing Then Sonando = False: Exit Function
    Sonando = Sound.Music_GetLoop
End Function