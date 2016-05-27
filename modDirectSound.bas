Attribute VB_Name = "modDirectSound"
#Const FinalMode = 1
' DirectX sound basic functions - by Blackd
Option Explicit
Public DirectX As DirectX7
Public DS As DirectSound
Public NumberOfSoundFiles As Integer
Public DSBuffer() As DirectSoundBuffer
Public HotkeysAreUsable As Boolean
Public SoundIsUsable As Boolean
Public SoundErrorWasThis As String
Public soundErrorLine As String
Public Function DirectX_Init(mainWindowHWND As Long, numberOfFiles As Long) As Boolean
  ' This function should be called only once, at the Load event of our main window
  On Error GoTo goterr
  soundErrorLine = "NumberOfSoundFiles = numberOfFiles"
  NumberOfSoundFiles = numberOfFiles
  soundErrorLine = "Set DS = DirectX.DirectSoundCreate("""")"
  Set DS = DirectX.DirectSoundCreate("")
  soundErrorLine = "DS.SetCooperativeLevel mainWindowHWND, DSSCL_PRIORITY"
  DS.SetCooperativeLevel mainWindowHWND, DSSCL_PRIORITY ' init with certain priority
  soundErrorLine = "NumberOfSoundFiles = numberOfFiles"
  NumberOfSoundFiles = numberOfFiles
  soundErrorLine = "ReDim DSBuffer(1 To NumberOfSoundFiles)"
  ReDim DSBuffer(1 To NumberOfSoundFiles)
  soundErrorLine = "SoundIsUsable = True"
  SoundIsUsable = True
  DirectX_Init = True ' initialized ok
  Exit Function
goterr:
  SoundIsUsable = False
  SoundErrorWasThis = "Executing: " & soundErrorLine & vbCrLf & "Got error number " & CStr(Err.Number) & " : " & Err.Description
  DirectX_Init = False ' failed to initialize
End Function
Public Sub DirectX_LoadSound(File As String, bufferSlot As Integer)
  ' copy a sound file from hard disk to RAM
  If SoundIsUsable = True Then
  Dim bufferDesc As DSBUFFERDESC
  Dim waveFormat As WAVEFORMATEX
  ' Let's tell it clear to directx : if we want to play a file,
  ' we want to hear the sound ALWAYS , doesn't matter if we have
  ' focus or not or whatever.
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or _
   DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC Or _
   DSBCAPS_STICKYFOCUS Or DSBCAPS_GLOBALFOCUS
  ' And now tell directx how to read the file :
  waveFormat.nFormatTag = WAVE_FORMAT_PCM ' .wav
  waveFormat.nChannels = 2 'stereo
  waveFormat.lSamplesPerSec = 22050 ' 22khz
  waveFormat.nBitsPerSample = 16 ' 16 bits
  waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
  waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
  ' Finally, lets copy file from hard disk to RAM :
  Set DSBuffer(bufferSlot) = DS.CreateSoundBufferFromFile(File, bufferDesc, waveFormat)
  End If
End Sub
  
Public Sub DirectX_PlaySound(bufferSlot As Integer)
  On Error GoTo gotError
  If SoundIsUsable = True Then
  ' Play one of the sounds that we have on memory
  DSBuffer(bufferSlot).Play DSBPLAY_DEFAULT ' only once: it just stop when it end
  End If
  Exit Sub
gotError:
  SoundIsUsable = False
  LogOnFile "errors.txt", "DirectX_PlaySound failed: " & Err.Number & " - " & Err.Description
End Sub

Public Sub DirectX_StopSound(bufferSlot As Integer)
  On Error GoTo gotError
  If SoundIsUsable = True Then
  ' (optional) Stop playing one of the sounds we have on Memory
  DSBuffer(bufferSlot).Stop
  End If
  Exit Sub
gotError:
  SoundIsUsable = False
  LogOnFile "errors.txt", "DirectX_StopSound failed: " & Err.Number & " - " & Err.Description
End Sub
  
Public Sub DirectX_SetVolume(vol As Long, bufferSlot As Integer)
  On Error GoTo gotError
  If SoundIsUsable = True Then
  ' (optional) Change the volume of 1 of our sounds that we have on memory
  DSBuffer(bufferSlot).SetVolume vol
  End If
  Exit Sub
gotError:
  SoundIsUsable = False
  LogOnFile "errors.txt", "DirectX_SetVolume failed: " & Err.Number & " - " & Err.Description
End Sub
