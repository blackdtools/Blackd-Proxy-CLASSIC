Attribute VB_Name = "modHotkeys"
#Const FinalMode = 1
Option Explicit
Public Type TypeHotkey
  key1 As Byte
  key2 As Byte
  command As String
  usable As Boolean
End Type

Public Hotkeys() As TypeHotkey
Public NumberOfHotkeys As Long
Public lastHotkeyCol As Long
Public lastHotkeyRow As Long
Public espectingHotkey As Boolean

Public dx As DirectX7
Public DI As DirectInput
Public DIV As DirectInputDevice
Public DID As DirectInputEnumDevices
Public DI_GUID As String
Public DII As DirectInputDeviceInstance
Public KeyB As DIKEYBOARDSTATE

Public debugdxError As String

Public Function InitDI() As String
  #If FinalMode Then
    On Error GoTo justend
  #End If
  Dim res As String
  res = ""
  HotkeysAreUsable = False
  If SoundIsUsable = True Then
  res = "Set DI = DX.DirectInputCreate"
         Set DI = dx.DirectInputCreate
  res = "Set DID = DI.GetDIEnumDevices(DIDEVTYPE_KEYBOARD, DIEDFL_ATTACHEDONLY)"
         Set DID = DI.GetDIEnumDevices(DIDEVTYPE_KEYBOARD, DIEDFL_ATTACHEDONLY)
  res = "Set DII = DID.GetItem(1)"
         Set DII = DID.GetItem(1)
  res = "DI_GUID = DII.GetGuidInstance"
         DI_GUID = DII.GetGuidInstance
  res = "Set DIV = DI.CreateDevice(DI_GUID)"
         Set DIV = DI.CreateDevice(DI_GUID)
  res = "DIV.SetCommonDataFormat DIFORMAT_KEYBOARD"
         DIV.SetCommonDataFormat DIFORMAT_KEYBOARD
  res = "DIV.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE"
         DIV.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  res = "DIV.Acquire"
         DIV.Acquire
  res = ""
  HotkeysAreUsable = True
  End If
justend:
  debugdxError = "Error number: " & Err.Number & " ; Error description: " & Err.Description
  InitDI = res
  Exit Function
End Function

Public Function TranslateHotkeyID(HotkeyID As Byte) As String
  Dim res As String
  Select Case HotkeyID
  Case 0
    res = "<NONE>"
  Case 1
    res = "ESCAPE"
  Case 2
    res = "1"
  Case 3
    res = "2"
  Case 4
    res = "3"
  Case 5
    res = "4"
  Case 6
    res = "5"
  Case 7
    res = "6"
  Case 8
    res = "7"
  Case 9
    res = "8"
  Case 10
    res = "9"
  Case 11
    res = "0"
  Case 14
    res = "BACKSPACE"
  Case 15
    res = "TAB"
  Case 16
    res = "Q"
  Case 17
    res = "W"
  Case 18
    res = "E"
  Case 19
    res = "R"
  Case 20
    res = "T"
  Case 21
    res = "Y"
  Case 22
    res = "U"
  Case 23
    res = "I"
  Case 24
    res = "O"
  Case 25
    res = "P"
  Case 28
    res = "ENTER"
  Case 29
    res = "L-CONTROL"
  Case 30
    res = "A"
  Case 31
    res = "S"
  Case 32
    res = "D"
  Case 33
    res = "F"
  Case 34
    res = "G"
  Case 35
    res = "H"
  Case 36
    res = "J"
  Case 37
    res = "K"
  Case 38
    res = "L"
  Case 42
    res = "L-SHIFT"
  Case 44
    res = "Z"
  Case 45
    res = "X"
  Case 46
    res = "C"
  Case 47
    res = "V"
  Case 48
    res = "B"
  Case 49
    res = "N"
  Case 50
    res = "M"
  Case 51
    res = ","
  Case 52
    res = "."
  Case 53
    res = "-"
  Case 54
    res = "R-SHIFT"
  Case 55
    res = "PAD *"
  Case 56
    res = "L-ALT"
  Case 57
    res = "SPACE"
  Case 58
    res = "CAPS"
  Case 59
    res = "F1"
  Case 60
    res = "F2"
  Case 61
    res = "F3"
  Case 62
    res = "F4"
  Case 63
    res = "F5"
  Case 64
    res = "F6"
  Case 65
    res = "F7"
  Case 66
    res = "F8"
  Case 67
    res = "F9"
  Case 68
    res = "F10"
  Case 69
    res = "PAD LOCK"
  Case 70
    res = "LOCK"
  Case 71
    res = "PAD 7"
  Case 72
    res = "PAD 8"
  Case 73
    res = "PAD 9"
  Case 74
    res = "PAD -"
  Case 75
    res = "PAD 4"
  Case 76
    res = "PAD 5"
  Case 77
    res = "PAD 6"
  Case 78
    res = "PAD +"
  Case 79
    res = "PAD 1"
  Case 80
    res = "PAD 2"
  Case 81
    res = "PAD 3"
  Case 82
    res = "PAD 0"
  Case 83
    res = "PAD ."
  Case 87
    res = "F11"
  Case 88
    res = "F12"
  Case 156
    res = "PAD ENTER"
  Case 157
    res = "R-CONTROL"
  Case 197
    res = "PAUSE"
  Case 181
    res = "PAD /"
  Case 183
    res = "PRINT"
  Case 184
    res = "R-ALT"
  Case 199
    res = "HOME"
  Case 200
    res = "UP ARROW"
  Case 201
    res = "PAG UP"
  Case 203
    res = "LEFT ARROW"
  Case 205
    res = "RIGHT ARROW"
  Case 207
    res = "END"
  Case 208
    res = "DOWN ARROW"
  Case 209
    res = "PAG DOWN"
  Case 210
    res = "INSERT"
  Case 211
    res = "DELETE"
  Case 219
    res = "WINDOWS"
  Case 221
    res = "MENU"
  Case Else
    res = "KEY #" & CStr(CInt(HotkeyID))
  End Select
  TranslateHotkeyID = res
End Function

Public Function HotkeyIDFixedLen(HotkeyID As Byte) As String
  Dim tmp As String
  Dim ltmp As Long
  Dim res As String
  tmp = CStr(CInt(HotkeyID))
  ltmp = Len(tmp)
  If ltmp = 1 Then
    res = "00" & tmp
  ElseIf ltmp = 2 Then
    res = "0" & tmp
  Else
    res = tmp
  End If
  HotkeyIDFixedLen = res
End Function

