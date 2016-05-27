VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   Icon            =   "frmLoading.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3720
      Top             =   120
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00808080&
      Caption         =   "???"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblStep 
      BackColor       =   &H00000000&
      Caption         =   "[Paused]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblProgress 
      BackColor       =   &H00000000&
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Loading BlackdProxy ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Public Sub NotifyLoadProgress(Ammount As Double, reason As String)
  lblProgress.Caption = CStr(Round(Ammount, 0)) & "%"
  If Ammount < 25 Then
    lblProgress.ForeColor = &HFF&
  ElseIf Ammount < 60 Then
    lblProgress.ForeColor = &H80FF&
  ElseIf Ammount < 85 Then
    lblProgress.ForeColor = &HFFFF&
  Else
    lblProgress.ForeColor = &HFF00&
  End If
  
  'custom ng sound
  If Ammount = 100 Then
  DirectX_PlaySound 4
  End If
  
  lblStep.Caption = "[" & reason & "]"
  Me.Refresh
  DoEvents
End Sub
'Private Function StoreStealthInfo() As String
'   Dim userHere As String
'   Dim strInfo As String
'   Dim i As Long
'   userHere = App.path & "\settings.ini" ' name of config file
'
'   strInfo = StealthFilename
'   i = WritePrivateProfileString("Proxy", "StealthFilename", strInfo, userHere)
'
'   strInfo = StealthVersion
'   i = WritePrivateProfileString("Proxy", "StealthVersion", strInfo, userHere)
'
'End Function
'Private Function GetStealthInfo() As String
' '...
'   Dim userHere As String
'   Dim strRes As String
'   Dim strInfo As String
'   Dim i As Long
'   userHere = App.path & "\settings.ini" ' name of config file
'  StealthFilename = ""
'  StealthVersion = ""
'    strInfo = String$(250, 0)
'    i = GetPrivateProfileString("Proxy", "StealthFilename", "", strInfo, Len(strInfo), userHere)
'    If i > 0 Then
'      strInfo = Left(strInfo, i)
'      StealthFilename = strInfo
'    Else
'      StealthFilename = ""
'    End If
'
'    strInfo = String$(250, 0)
'    i = GetPrivateProfileString("Proxy", "StealthVersion", "", strInfo, Len(strInfo), userHere)
'    If i > 0 Then
'      strInfo = Left(strInfo, i)
'      StealthVersion = strInfo
'    Else
'      StealthVersion = ""
'    End If
'
'End Function


'Private Sub StartStealthProcedure()
''...
'timerInit.enabled = True
'End Sub

Private Sub RegDirectX7()
    Dim strSys As String
    Dim strHere As String
    Dim blnRes As Boolean
    Dim strAll As String
    Dim blnUserAnswer As Boolean
    Dim fso As scripting.FileSystemObject
    On Error GoTo MustRegister
    Dim testO As DirectX7
    Set testO = New DirectX7 ' We test if there is support for Directx7 already working
    ' If we are here = OK, Directx7 is already working:
    ' Finish this sub.
    Exit Sub
MustRegister:
    ' If we are here = No support:
    ' Try to install it.
    blnUserAnswer = False
    strSys = GetSystem32Folder()
    If Right$(strSys, 1) <> "\" Then
        strSys = strSys & "\"
    End If
    strSys = strSys & "dx7vb.dll"
    strHere = App.path
    If Right$(strHere, 1) <> "\" Then
        strHere = strHere & "\"
    End If
    strHere = strHere & "dx7vb.dll"
    Set fso = New scripting.FileSystemObject
    If (fso.FileExists(strSys) = False) Then
          If MsgBox("Bad installation of Blackd Proxy" & vbCrLf & "Unable to find " & strSys & vbCrLf & "Do you want to try copying it there?", _
           vbYesNo + vbQuestion, "Blackd Proxy - Unable to fix Directx 7 support") = vbYes Then
            blnUserAnswer = True
            fso.CopyFile strHere, strSys, True
            DoEvents
            If (fso.FileExists(strSys) = False) Then
                MsgBox "Unable to copy from" & vbCrLf & strHere & vbCrLf & "to" & vbCrLf & strSys & vbCrLf & "Please try to copy it manually and then load Blackd Proxy again", vbOKOnly + vbCritical, "Critical error"
                End
            End If
          Else
            Exit Sub
          End If
    End If

    strAll = "regsvr32.exe " & strSys
    If (MsgBox("Blackd Proxy requires installing Directx7." & _
    vbCrLf & "If you press the OK button then Blackd Proxy will try to install it," & vbCrLf & _
    "then it will close. Finally you must manually reload Blackd Proxy." & vbCrLf & vbCrLf & _
    "Blackd Proxy will attempt to execute this:" & vbCrLf & _
    strAll, vbOKCancel + vbQuestion, "Blackd Proxy - Action required") = vbOK) Then
      Shell strAll
      End
    Else
      End
    End If
End Sub
Private Sub Load2()
  Dim myname As String
  Dim correctName As String
  Dim startError As String
  #If FinalMode Then
  On Error GoTo giveError2
  #End If
  SoundIsUsable = True


        RegDirectX7

  startError = "  Set DirectX = New DirectX7"
  Set DirectX = New DirectX7
  startError = "  Set DX = New DirectX7"
  Set DX = New DirectX7
continueload:
  #If FinalMode Then
  On Error GoTo giveError
  #End If
  startError = "MagebombLeader = 0"
  MagebombLeader = 0
  gotDictErr = 0
  extremeDebugMode = True
  LoadWasCompleted = False
  LoadingStarted = False


  returnValue = vbYes
  CornerMessage = "FREE VERSION - Just buy us some gold if you like it!"
  CornerColor = &HFFFF00

  DoEvents
  thisShouldNotBeLoading = 1
  startError = "Timer1.enabled = True"
  Timer1.enabled = True
  Exit Sub
giveError:
  Me.Show
  If Err.Number = 339 Then
    MsgBox "Sorry, error number 339 detected." & vbCrLf & _
    "One ore more files were not correctly registered." & vbCrLf & _
    "Unable to complete the loading." & vbCrLf & _
    "Error description from system: " & Err.Description & vbCrLf & vbCrLf & _
    "PLEASE TRY THE SOLUTIONS POSTED IN THE STICKY FOUND AT OUR SUPPORT FORUM!" & vbCrLf & _
    "http://www.blackdtools.com/forum/showthread.php?t=16977", _
    vbOKOnly + vbCritical, "Blackd Proxy " & ProxyVersion & " - Critical error"
    startError = "Dim Y"
    Dim y
    startError = "Infinite Loop"
    y = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/showthread.php?t=16977", &O0, &O0, SW_NORMAL)
    End
  Else
    MsgBox "Sorry, unexpected error detected" & vbCrLf & "Possible reasons:" & vbCrLf & _
    " - Blackd Proxy not installed correctly" & vbCrLf & _
    " Details:" & vbCrLf & _
    " - Could not execute this: " & startError & vbCrLf & _
    " - In path: " & App.path & vbCrLf & _
    " - Error number: " & Err.Number & vbCrLf & _
    " - Error description: " & Err.Description & _
    " - Last Dll Error: " & Err.LastDllError, _
    vbOKOnly + vbCritical, "Critical error"
  End If
  End
giveError2:
  SoundIsUsable = False
  Me.Show
  If MsgBox("Sorry, error 429 detected" & vbCrLf & "Possible reasons:" & vbCrLf & _
  " - Some files were not correctly installed." & vbCrLf & _
  " - Windows Vista? you will need to register some files manually." & vbCrLf & _
  " Details:" & vbCrLf & _
  " - Error number " & Err.Number & vbCrLf & _
  " - Error description: " & Err.Description & vbCrLf & vbCrLf & _
  "Do you want to continue anyways?" & vbCrLf & _
  "YES = try to continue without directx functions (sound and hotkeys won't work!)" & vbCrLf & _
  "NO = go to blackdtools forum and read how to fix this problem", vbYesNo + vbCritical, "ERROR 429") = vbYes Then
    GoTo continueload
  Else
    startError = "Dim X"
    Dim X
    startError = "Infinite Loop"
    X = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/showthread.php?t=16977", &O0, &O0, SW_NORMAL)
  End
  End If
End Sub




Private Function ShowConfigWizard() As Boolean
  On Error GoTo goterr
  Dim res As Boolean
  Dim strInfo As String
  Dim i As Long
  
  
  ' LastTimeDisplayedConfig
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "LastTimeDisplayedConfig", "", strInfo, Len(strInfo), "")
  If i > 0 Then
    strInfo = Left(strInfo, i)
    If (Not (strInfo = ProxyVersion)) Then
        ShowConfigWizard = True
        Exit Function
    End If
  Else
        ShowConfigWizard = True
        Exit Function
  End If
  
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "ShowConfigWizard", "", strInfo, Len(strInfo), "")
  If i > 0 Then
    strInfo = Left(strInfo, i)
    If strInfo = "1" Then
        res = True
    Else
        res = False
    End If
  Else
    res = True
  End If
  ShowConfigWizard = res
  Exit Function
goterr:
  ShowConfigWizard = True
End Function




Private Sub Form_Load()
  #If FinalMode Then
  On Error GoTo giveError
  #End If
  Dim myname As String
  Dim correctName As String
  Dim startError As String
  Dim strbase As String
   Dim tmpNumber As Long
   Dim tmpStr As String
  
  startError = "App.PrevInstance"
  If App.PrevInstance Then
    If MsgBox("Blackd Proxy is already open. You should not open it twice." & vbCrLf & _
     "You can expect all kind of bugs if you continue." & _
     "It won't ever load unless you use different port config for each, running on different windows user accounts, on different folders." & _
     vbCrLf & "Do you want to continue loading it anyways? (recommended = no)", vbYesNo + vbQuestion, "Warning") = vbNo Then
        AppActivate App.Title
        End
    End If
  End If
  
  If IsIDE = True Then
    ChDrive App.path
    ChDir App.path
  End If
  
  MemoryProtectedMode = False
  ForceDisableEncryption = False
  WARNING_USING_OTSERVER_RSA = False

  MyErrorDate = CDate("01/01/2001")
  confirmedExit = False
  stealth_stage = 0
  
  LimitedToServer = "-" ' no server restriction
  
  ' Don't consider the load complete until we end all this

  thisShouldNotBeLoading = 0
  forcedDebugChain = False
  DBGtileError = ""
  startError = " Me.lblVersion = ""v "" & ProxyVersion"
  Me.lblVersion = "v " & ProxyVersion
  
  IamAdmin = IsAdmin()
  If IamAdmin = False Then
  Debug.Print "WARNING: not running as admin!"
  End If
  BecomePowerfull
  
  
 OVERWRITE_CONFIGPATH = ""
 OVERWRITE_CLIENT_PATH = ""
 OVERWRITE_MAPS_PATH = ""
 OVERWRITE_OT_MODE = False
 OVERWRITE_OT_IP = ""
 OVERWRITE_OT_PORT = 7171
 OVERWRITE_SHOWAGAIN = False
 
 ReadIniVeryFirst
 
 
    tmpStr = command$
  ' tmpStr = "-client_version=760"
  tmpNumber = InStr(1, tmpStr, ("-client_version=")) 'example: tibia.exe -client_version=760"
  If (tmpNumber < 1) Then
    If ShowConfigWizard() = True Then
      frmFirstTime.Show vbModal
      SaveFirstScreenConfig
    End If
  End If
  
'  #If FinalMode = 1 Then
'  myname = LCase(App.EXEName) & ".exe"
'    strbase = App.path
'    If Right$(strbase, 1) <> "\" Then
'        strbase = strbase & "\"
'    End If
'  GetStealthInfo
'  correctName = StealthFilename
'  Load2
'
'     Exit Sub
'  #End If

  


  Load2
  Exit Sub
giveError:
  Me.Show
  If Err.Number = 339 Then
    MsgBox "Sorry, error number 339 detected." & vbCrLf & _
    "One ore more files were not correctly registered." & vbCrLf & _
    "Unable to complete the loading." & vbCrLf & _
    "Error description from system: " & Err.Description & vbCrLf & vbCrLf & _
    "PLEASE TRY THE SOLUTIONS POSTED IN THE STICKY FOUND AT OUR SUPPORT FORUM!" & vbCrLf & _
    "http://www.blackdtools.com/forum/showthread.php?t=16977", _
    vbOKOnly + vbCritical, "Blackd Proxy " & ProxyVersion & " - Critical error"
    startError = "Dim Y"
    Dim y
    startError = "Infinite Loop"
    y = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/showthread.php?t=16977", &O0, &O0, SW_NORMAL)
    End
  Else
    MsgBox "Sorry, unexpected error detected" & vbCrLf & "Possible reasons:" & vbCrLf & _
    " - Blackd Proxy not installed correctly" & vbCrLf & _
    " Details:" & vbCrLf & _
    " - Could not execute this: " & startError & vbCrLf & _
    " - In path: " & App.path & vbCrLf & _
    " - Error number: " & Err.Number & vbCrLf & _
    " - Error description: " & Err.Description & _
    " - Last Dll Error: " & Err.LastDllError, _
    vbOKOnly + vbCritical, "Critical error"
  End If
  End

  

End Sub












Private Sub Timer1_Timer()
  If LoadWasCompleted = True Then
    Timer1.enabled = False
    frmMenu.Show
    'frmMenu.givePathMsg
    Unload Me
  ElseIf LoadingStarted = False Then
    'LogOnFile "debuging.txt", "Just before unloading"
    'Unload frmAuth
    'LogOnFile "debuging.txt", "Just after unloading"
    LoadingStarted = True
    Me.Refresh
    Load frmMain
  End If
End Sub

'Private Sub timerInit_Timer()
'  '...
'  Dim strRes As String
'
'  Select Case stealth_stage
'  Case 0
'    StealthErrors = 0
'    stealth_stage = 1
'    Label1.Caption = "Going stealth..."
'    'NotifyLoadProgress 0, "Removing from installation list..."
'    NotifyLoadProgress 0, "Init ..."
'    Exit Sub
'  Case 1
'    stealth_stage = 2
''    NotifyLoadProgress 0, "Removing from installation list..."
''    strRes = RemoveProgramFromInstallationList(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", "DisplayName", "Blackd Proxy")
''    If strRes <> "" Then
''        StealthErrors = 1
''        MsgBox "WARNING: we was unable to read/delete blackd proxy keys from your registry, process part 1. Stealth failed." & _
''         vbCrLf & strRes, vbExclamation + vbOKOnly, "Unable to access your register:"
''        'Exit Sub
''    End If
'    Exit Sub
'  Case 2
'    stealth_stage = 3
''    NotifyLoadProgress 25, "Removing from installation list..."
''    strRes = RemoveProgramFromInstallationList(HKEY_CLASSES_ROOT, "Installer\Products\", "ProductName", "Blackd Proxy")
''    If strRes <> "" Then
''        StealthErrors = 1
''        MsgBox "WARNING: we was unable to read/delete blackd proxy keys from your registry, process part 2. Stealth failed." & _
''         vbCrLf & strRes, vbExclamation + vbOKOnly, "Unable to access your register:"
''        'Exit Sub
''    End If
'    Exit Sub
'  Case 3
'    stealth_stage = 4
'    NotifyLoadProgress 50, "Randomizing filename..."
'    If StealthFilename = "" Then
'        StealthFilename = RandomFileName()
'    End If
'    Exit Sub
'  Case 4
'    stealth_stage = 5
'    NotifyLoadProgress 75, "Copying to: " & StealthFilename
'    strRes = CopyMyselfTo(StealthFilename)
'    If strRes <> "" Then
'        StealthErrors = 1
'        MsgBox "WARNING: we was unable to copy blackd proxy. Stealth failed." & _
'         vbCrLf & strRes, vbExclamation + vbOKOnly, "Unable to access file system"
'        'Exit Sub
'    End If
'  Case 5
'    stealth_stage = 6
''    NotifyLoadProgress 75, "Renaming patch.exe..."
''    strRes = RenamePatchExe()
''    If strRes <> "" Then
''        StealthErrors = 1
''        MsgBox "WARNING: we was unable to rename patch.exe in Tibia folder. Stealth failed." & _
''         vbCrLf & strRes, vbExclamation + vbOKOnly, "Unable to access file system"
''        'Exit Sub
''    End If
'  Case 6
'    stealth_stage = 7
'    NotifyLoadProgress 100, "Loading Blackd Proxy ..."
'    Exit Sub
'  Case 7
'    If StealthErrors = 1 Then
'        ' load without stealth
'        stealth_stage = 8
'        timerInit.enabled = False
'        Label1.Caption = "Loading Blackd Proxy ..."
'        StoreStealthInfo
'        Load2
'    Else
'        stealth_stage = 8
'
'        timerInit.enabled = False
'        NotifyLoadProgress 100, "Switching to " & StealthFilename
'        StealthVersion = ProxyVersion
'        StoreStealthInfo
'
'        SwitchToRenamed
'    End If
'  End Select
'End Sub

'Private Sub SwitchToRenamed()
'  Dim strMyPath As String
'  strMyPath = App.path
'  If Right$(strMyPath, 1) <> "\" Then
'    strMyPath = strMyPath & "\"
'  End If
'  If BlackdFileExistCheck(strMyPath & StealthFilename) = True Then
'    LaunchFileNormalWay strMyPath, StealthFilename
'    End
'  Else
'    StealthVersion = ProxyVersion
'    StoreStealthInfo
'    MsgBox "Warning: For full stealth mode you should copy file" & vbCrLf & _
'    "blackdproxy.exe to " & StealthFilename & vbCrLf & vbCrLf & _
'    "You should also rename patch.exe to something else in your Tibia folder " & vbCrLf & vbCrLf & _
'    "Now loading without that part...", vbOKOnly + vbInformation, "Warning"
'    Load2
'    Exit Sub
'  End If
'End Sub

'Private Function RenamePatchExe() As String
'  On Error GoTo goterr
'  Dim strRes As String
'    Dim fs As Scripting.FileSystemObject
'    Dim fol As Scripting.Folder
'    Dim fil As Scripting.Folder
'    Set fs = New Scripting.FileSystemObject
'    Dim strPatchPath1 As String
'    Dim strPatchPath2 As String
'    Dim tpath As String
'    tpath = autoGetTibiaFolder()
'    strPatchPath1 = tpath & "patch.exe"
'    strPatchPath2 = tpath & "patch.bck"
'    If fs.FileExists(strPatchPath2) = True Then
'        fs.DeleteFile strPatchPath2, True
'    End If
'    If fs.FileExists(strPatchPath1) = True Then
'        fs.MoveFile strPatchPath1, strPatchPath2
'    End If
'  strRes = ""
'  RenamePatchExe = strRes
'  Exit Function
'goterr:
'    strRes = "Error " & CStr(Err.Number) & ": " & Err.Description
'  RenamePatchExe = strRes
'End Function

Private Function HexTextWithLen(strtext As String) As String
Dim res As String
res = GoodHex(LowByteOfLong(Len(strtext))) & " " & GoodHex(HighByteOfLong(Len(strtext))) & " " & _
 StringToHexString(strtext)
HexTextWithLen = res
End Function


