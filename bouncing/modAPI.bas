Attribute VB_Name = "modAPI"

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal _
lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function WaveOutGetNumDevs Lib "winmm.dll" () As Long


Public Function FileExists(FullFileName) As Boolean
    FileExists = IIf(Dir(FullFileName) = "", False, IIf(FullFileName <> "", True, False))
End Function

Public Function CanPlayWAVs() As Boolean
    'This function determines if the machine can play WAV files
    CanPlayWAVs = WaveOutGetNumDevs()
End Function

Public Function PlayWAVFile(StrFileName As String, Optional blnAsync As Boolean) As Boolean
    'This function plays a wav file
    Dim lngFlags
    'Flag Values for Parameter
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_FILENAME = &H20000
    'Set the flags for PlaySound
    lngFlags = SND_NODEFAULT Or SND_FILENAME Or SND_SYNC
    If blnAsync Then lngFlags = lngFlags Or SND_ASYNC
    'Play the WAV file
    PlayWAVFile = PlaySound(StrFileName, 0&, lngFlags)
End Function



