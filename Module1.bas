Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlaySound(strFileName As String)
    
    sndPlaySound strFileName, 2

End Sub
