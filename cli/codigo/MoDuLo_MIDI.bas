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
    CurMidi = Archivo
    ' En el motor DX8, la musica se carga via Sound.NextMusic
    ' El archivo se procesa en Sound_Render
    Sound.NextMusic = Archivo
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