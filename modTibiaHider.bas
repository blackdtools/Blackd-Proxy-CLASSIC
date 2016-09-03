Attribute VB_Name = "modTibiaHider"
Option Explicit
#Const FinalMode = 1





'Exactly one of the following flags specifying how to show the window:
Private Const SW_HIDE = 0
'Hide the window.
Private Const SW_MAXIMIZE = 3
'Maximize the window.
Private Const SW_MINIMIZE = 6
'Minimize the window.
Private Const SW_RESTORE = 9
'Restore the window (not maximized nor minimized).
Private Const SW_SHOW = 5
'Show the window.
Private Const SW_SHOWMAXIMIZED = 3
'Show the window maximized.
Private Const SW_SHOWMINIMIZED = 2
'Show the window minimized.
Private Const SW_SHOWMINNOACTIVE = 7
'Show the window minimized but do not activate it.
Private Const SW_SHOWNA = 8
'Show the window in its current state but do not activate it.
Private Const SW_SHOWNOACTIVATE = 4
'Show the window in its most recent size and position but do not activate it.
Private Const SW_SHOWNORMAL = 1
'Show the window and activate it (as usual).

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As _
Long, ByVal nCmdShow As Long) As Long

Private Const HideProccess As Long = 0
Private Const ShowProccess As Long = 5
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Declare Function FormatMessage Lib "kernel32" Alias _
   "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
   ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, _
   Arguments As Long) As Long

Public Function APIErrorDescription(ErrorCode As Long) As String



Dim sAns As String
Dim lRet As Long

'PURPOSE: Returns Human Readable Description of
'Error Code that occurs in API function

'PARAMETERS: ErrorCode: System Error Code

'Returns: Description of Error

'Example: After Calling API Function:
         'MsgBox (APIErrorDescription(Err.LastDllError))
 
sAns = Space(255)
lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, _
   ErrorCode, 0, sAns, 255, 0)

APIErrorDescription = StripNull(sAns)

End Function

Private Function StripNull(ByVal InString As String) As String

'Input: String containing null terminator (Chr(0))
'Returns: all character before the null terminator

Dim iNull As Integer
If Len(InString) > 0 Then
    iNull = InStr(InString, vbNullChar)
    Select Case iNull
    Case 0
        StripNull = Trim(InString)
    Case 1
        StripNull = ""
    Case Else
       StripNull = Left$(Trim(InString), iNull - 1)
   End Select
End If

End Function
   
   
   


Public Sub SetTibiaClientsVisible(ByVal blnVisible As Boolean)
  Dim intMode As Long
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim res As Boolean
  Dim tibiaclient As Long
  If blnVisible = True Then
    intMode = SW_SHOW
  Else
    intMode = SW_HIDE
  End If
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      res = ShowWindow(tibiaclient, intMode)
      'If res = False And Err.LastDllError <> 0 Then
      '  MsgBox "Dll error " & Err.LastDllError & ": " & APIErrorDescription(Err.LastDllError), vbOKOnly + vbCritical, "SetTibiaClientsVisible"
      '  End
     ' End If
    End If
  Loop
  Exit Sub
gotErr:
  intMode = -1
End Sub




