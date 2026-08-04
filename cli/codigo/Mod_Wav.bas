Attribute VB_Name = "Mod_Wav"
Option Explicit

' =====================================================================
' Mod_Wav - Motor de audio WAV (DX8 compatible)
' Reemplazado: DirectSound DX7 -> sndPlaySound (Win32) + clsSoundEngine
' =====================================================================

Public Const SND_SYNC       = &H0   ' SINCRONO
Public Const SND_ASYNC      = &H1   ' ASINCRONO
Public Const SND_NODEFAULT  = &H2   ' silence not default, if sound not found
Public Const SND_LOOP       = &H8   ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP     = &H10  ' don't stop any currently playing sound

Public Const SND_WAV_CLICK      = "click.Wav"
Public Const SND_WAV_PASOS1     = "23.Wav"
Public Const SND_WAV_PASOS2     = "24.Wav"
Public Const SND_WAV_NAVEGANDO  = "50.wav"
Public Const SND_WAV_OVER       = "click2.Wav"
Public Const SND_WAV_DICE       = "cupdice.Wav"

' -------------------------------------------------------
' PlayWaveDS - Reproduce un WAV usando Win32 sndPlaySound
' Compatible con todas las llamadas existentes del cliente
' -------------------------------------------------------
Sub PlayWaveDS(File As String)
    If Fx = 1 Then Exit Sub
    On Error Resume Next
    sndPlaySound DirSound & File, SND_ASYNC Or SND_NODEFAULT
End Sub