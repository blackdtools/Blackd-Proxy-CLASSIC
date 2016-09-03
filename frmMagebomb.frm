VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMagebomb 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Magebomb"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMagebomb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReloginChar 
      Height          =   285
      Left            =   4080
      TabIndex        =   25
      Text            =   "2000"
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkResend 
      BackColor       =   &H00000000&
      Caption         =   "Relogin a char if he is not in after "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock clientLess 
      Index           =   0
      Left            =   3960
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridBomb 
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   0
      BorderStyle     =   0
   End
   Begin VB.CommandButton cmdDebug 
      BackColor       =   &H0080FF80&
      Caption         =   "Enable  magebomb DEBUG"
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CheckBox chkRecordLogins 
      BackColor       =   &H00000000&
      Caption         =   $"frmMagebomb.frx":0442
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Timer armageddonTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4920
      Top             =   2040
   End
   Begin VB.CommandButton cmdAddMagebomb 
      BackColor       =   &H0000C000&
      Caption         =   "Add character to memory with selected settings"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   5415
   End
   Begin VB.CommandButton cmdDeleteSelected 
      BackColor       =   &H0000C000&
      Caption         =   "Delete selected in memory"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdClearMemory 
      BackColor       =   &H0000C000&
      Caption         =   "Clear the all list in memory"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdLoadMagebomb 
      BackColor       =   &H0000C000&
      Caption         =   "LOAD"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveMagebomb 
      BackColor       =   &H0000C000&
      Caption         =   "SAVE"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "magebomb\list.txt"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtRetryTime 
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Text            =   "30000"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.ComboBox cmbOrderType 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Text            =   "type 5 : SD (battlelist)"
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   5880
      Width           =   3255
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdReloadFiles 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- Reload"
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Loads from given file"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "ms"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   3870
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   $"frmMagebomb.frx":04FD
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "WELCOME TO THE ULTIMATE DESTRUCTION WEAPON!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "  To shoot the mage bomb just go near your target(s) and type:                      EXIVA BOMB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   8040
      Width           =   5655
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   "Status: Preload your desired characters first."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Save / Load the list:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "CHOOSE TIME FOR RETRYS BEFORE CLOSE (in ms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "CHOOSE TARGET:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "CHOOSE ATTACK MODE:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "CHOOSE LOGIN .LOG:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LIST OF PRELOADED CHARACTERS FOR THE MAGEBOMB:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5295
   End
End
Attribute VB_Name = "frmMagebomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Public Sub ReloadMagebombFiles()
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  Dim path As String
  path = ""
  Dim fs As scripting.FileSystemObject
  Dim f As scripting.Folder
  Dim f1 As scripting.File
  Set fs = New scripting.FileSystemObject
  path = App.path & "\magebomb"
  Set f = fs.GetFolder(path)
  txtFile.Clear
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "log" Then
        txtFile.AddItem f1.name
    End If
  Next
  txtFile.Text = ""
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "ERROR WITH FILESYSTEM OBJECT at ReloadMagebombFiles (" & Err.Number & ") : " & Err.Description & " ; path=" & path
End Sub

Private Function convertThisStrToLong(strLong As String) As Long
  Dim lonResult As Long
  On Error GoTo gotErr
  lonResult = CLng(strLong)
  convertThisStrToLong = lonResult
  Exit Function
gotErr:
  convertThisStrToLong = -1
End Function


Public Function PreloadAbomb(strLoginLogFile As String, strAttackMode As String, _
  strTarget As String, strTimeRetry As String)
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim strMode As String
  Dim longTime As Long
  Dim res As Long
  Dim logpath As String
  Dim spCond As Boolean
  res = -1
  If (Len(strLoginLogFile) < 5) Then
    PreloadAbomb = -2 'invalid login log selected
    Exit Function
  End If
  Select Case strAttackMode
  Case "type 0 : SD (XYZ)"
    strMode = "0"
  Case "type 1 : HMM (XYZ)"
    strMode = "1"
  Case "type 2 : Explosion (XYZ)"
    strMode = "2"
  Case "type 3 : IH (XYZ)"
    strMode = "3"
  Case "type 4 : UH (XYZ)"
    strMode = "4"
  Case "type 5 : SD (battlelist)"
    strMode = "5"
  Case "type 6 : HMM (battlelist)"
    strMode = "6"
  Case "type 7 : Explosion (battlelist)"
    strMode = "7"
  Case "type 8 : IH (battlelist)"
    strMode = "8"
  Case "type 9 : UH (battlelist)"
    strMode = "9"
  Case "type B : fireball (battlelist)"
    strMode = "B"
  Case "type C : stalagmite (battlelist)"
    strMode = "C"
  Case "type D : icicle (battlelist)"
    strMode = "D"
 
  Case Else
    strMode = " "
  End Select
  If strMode = " " Then
    PreloadAbomb = -3 ' Invalid attack mode selected
    Exit Function
  End If
  If (strTarget = "") Then
    PreloadAbomb = -4 ' Invalid target selected
    Exit Function
  End If
  longTime = convertThisStrToLong(strTimeRetry)
  If longTime < 0 Then
    PreloadAbomb = -5 ' Invalid time selected
    Exit Function
  End If
  ' log file exist? ...
  res = -7 ' will return - 7 if FileSystemObject fails
  Dim fso As scripting.FileSystemObject
  Set fso = New scripting.FileSystemObject
  logpath = App.path & "\magebomb\" & strLoginLogFile
  If (fso.FileExists(logpath) = False) Then
    PreloadAbomb = -8 ' Doesn't exist
    Exit Function
  End If
  res = -9
  ' all ok , then load it ...
  Dim AddingCharname As String
  Dim AddingVersion As Long
  Dim AddingIP As String
  Dim AddingPort As Long
  Dim AddingRawKey As String
  Dim AddingRawLoginPacket As String
  Dim strTmp As String
  Dim ubLoginPacket As Long
  Dim fn As Integer
  fn = FreeFile
  Open logpath For Input As #fn
    Line Input #fn, AddingCharname
    Line Input #fn, strTmp
    AddingVersion = CLng(strTmp)
    Line Input #fn, AddingIP
    Line Input #fn, strTmp
    AddingPort = CLng(strTmp)
    Line Input #fn, AddingRawKey
    Line Input #fn, strTmp
    ubLoginPacket = CLng(strTmp)
    Line Input #fn, AddingRawLoginPacket
  Close #fn
  spCond = ExistMagebombCharInMemory(AddingCharname)
  AddToMagebombMemory strLoginLogFile, AddingCharname, AddingVersion, AddingIP, AddingPort, AddingRawKey, ubLoginPacket, AddingRawLoginPacket, strMode, strTarget, longTime
  DisplayMagebombMemory
  If spCond = True Then
    PreloadAbomb = -6
  Else
    PreloadAbomb = 0
  End If
  Exit Function
gotErr:
  PreloadAbomb = res
End Function

Public Sub ProcessArmageddon()
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim limM As Long
  Dim gtc As Long
  Dim anyNotReady As Boolean
  Dim anyShooted As Boolean
  Dim i As Long
  Dim aRes As Long
  Dim dRes As Long
  Dim sCheat As String
  Dim cCheat() As Byte
  Dim inRes As Integer
  Dim lngD1 As Long
  Dim lngD2 As Long
  Dim k As Long
  Dim errline As Long
  
  errline = 0
  anyShooted = False
  errline = 1
  anyNotReady = False
  errline = 2
  gtc = GetTickCount()
  errline = 3
  limM = MagebombsLoaded - 1
  
  
  
 ' lngD1 = Me.clientLess(0).State
 ' lngD2 = Me.clientLess(1).State
  
 ' dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & CStr(lngD1) & ";" & CStr(lngD2) & " in stage " & CStr(MagebombStage))
 ' DoEvents
  
  
  
  errline = 4
  If (MagebombStage < 2) Then
  errline = 5
  For i = 0 To limM
  errline = 6
    If clientLess(i).State > 7 Then
    errline = 7
        Magebombs(i).ConnectionStatus = 0
        errline = 8
        If clientLess(i).State = 9 Then
            errline = 9
            clientLess(i).Close
            errline = 10
        End If
    End If
    errline = 11
    Select Case Magebombs(i).ConnectionStatus
    Case 0 ' not started yet
      errline = 12
      anyNotReady = True
      errline = 13
      Magebombs(i).ConnectionStatus = 1
      errline = 14
      Magebombs(i).ConnectionTimeout = gtc + FIRSTCONNECTIONTIMEOUT_ms
      errline = 15
      clientLess(i).Close
      errline = 16
      DoEvents
      errline = 17
      clientLess(i).RemoteHost = Magebombs(i).IPstring
      errline = 18
      clientLess(i).RemotePort = Magebombs(i).PORTnumber
      errline = 19
      clientLess(i).Connect
      errline = 20
      DoEvents
      errline = 21
      If DebugingMagebomb = True Then
        errline = 22
        dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  connecting...")
        errline = 23
        DoEvents
      End If
      errline = 24
    Case 1 ' connecting
      errline = 25
      anyNotReady = True
      errline = 26
      If (gtc > Magebombs(i).ConnectionTimeout) Then
        errline = 27
        clientLess(i).Close
        errline = 28
        If DebugingMagebomb = True Then
          errline = 29
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  had connection timeout.")
          errline = 30
          DoEvents
        End If
      End If
      errline = 31
      If (clientLess(i).State <> sckConnected) And (clientLess(i).State <> sckConnecting) Then
        errline = 32
        Magebombs(i).ConnectionStatus = 1
        errline = 33
        Magebombs(i).ConnectionTimeout = gtc + FIRSTCONNECTIONTIMEOUT_ms
        errline = 34
        clientLess(i).Close
        errline = 35
        DoEvents
        errline = 36
        clientLess(i).RemoteHost = Magebombs(i).IPstring
        errline = 37
        clientLess(i).RemotePort = Magebombs(i).PORTnumber
        errline = 38
        clientLess(i).Connect
        errline = 39
        DoEvents
        If (DebugingMagebomb = True) Then
          errline = 40
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  retrying to connect...")
          errline = 41
          DoEvents
        End If
      End If
    End Select
  Next i
  errline = 42
  If (anyNotReady = False) Then
    errline = 43
    MagebombStage = 2
    errline = 44
    DoEvents
    Exit Sub
  End If
  End If
  errline = 45
  If (MagebombStage = 2) Then ' sending login
     errline = 46
     For i = 0 To limM
         errline = 47
         If clientLess(i).State > 7 Then
            errline = 48
            MagebombStage = 1
            errline = 49
            Magebombs(i).ConnectionStatus = 0
            errline = 50
            If clientLess(i).State = 9 Then
                errline = 51
                clientLess(i).Close
            End If
            DoEvents
            Exit Sub
        End If
     Next i
     errline = 52
     If DebugingMagebomb = True Then
          errline = 53
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : Stage 1 completed. Now sending logins...")
          DoEvents
     End If
     errline = 54
     For i = 0 To limM
       errline = 55
       Magebombs(i).ConnectionStatus = 3
       errline = 56
       Magebombs(i).ConnectionTimeout = gtc + SECONDCONNECTIONTIMEOUT_ms
       errline = 57
       Magebombs(i).nextSendLogin = gtc + CLng(frmMagebomb.txtReloginChar.Text)
       errline = 58
     Next i
     errline = 59
     For i = 0 To limM
       errline = 60
       SendClientless CInt(i), Magebombs(i).loginPacket, True
       errline = 61
       DoEvents
       DoEvents
        If DebugingMagebomb = True Then
          errline = 62
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  is now doing login...")
          errline = 63
          DoEvents
        End If
        errline = 64
       'sCheat = "04 00 A0 01 01 01"
       'inRes = GetCheatPacket(cCheat, sCheat)
       'SendClientless i, cCheat
       'DoEvents
     Next i
     errline = 65
     MagebombStage = 3
     errline = 66
     Exit Sub
  ElseIf (MagebombStage = 3) Then ' time to shot
    errline = 67
    For i = 0 To limM
      errline = 68
      If (gtc > Magebombs(i).ConnectionTimeout) Then ' retry time out ?
        errline = 69
        If ((DebugingMagebomb = True) And (Magebombs(i).ConnectionStatus <> 0)) Then
          errline = 70
          dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  attack timeout. Closing.")
          DoEvents
          errline = 71
        End If
        errline = 72
        Magebombs(i).ConnectionStatus = 0
        errline = 73
        clientLess(i).Close
      Else
        errline = 74
        If ((clientLess(i).State = sckConnected) And (Magebombs(i).ConnectionStatus = 4)) Then
          errline = 75
          anyShooted = True
          errline = 76
          SendClientless i, SafeModeOutPacket, False
          errline = 77
          DoEvents
          errline = 78
          SendClientless i, Magebombs(i).attackPacket, False
          errline = 79
          DoEvents
        Else
          errline = 80
          If frmMagebomb.chkResend.Value = 1 Then
            errline = 81
            If Magebombs(i).nextSendLogin < gtc Then
                errline = 82
                Magebombs(i).nextSendLogin = gtc + CLng(frmMagebomb.txtReloginChar.Text)
                errline = 83
                Magebombs(i).ConnectionStatus = 2
                errline = 84
                clientLess(i).Close
                errline = 85
                DoEvents
                errline = 86
                clientLess(i).RemoteHost = Magebombs(i).IPstring
                errline = 87
                clientLess(i).RemotePort = Magebombs(i).PORTnumber
                errline = 88
                clientLess(i).Connect
                errline = 89
                DoEvents
                If DebugingMagebomb = True Then
                  errline = 90
                  dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : " & Magebombs(i).CharacterName & "  connecting...")
                  DoEvents
                  errline = 91
                End If

            End If
          End If
          errline = 92
          anyShooted = True
          DoEvents
          errline = 93
        End If
      End If
    Next i
    errline = 94
    If (anyShooted = False) Then
      errline = 95
      For i = 0 To limM
        errline = 96
        clientLess(i).Close
        errline = 97
      Next i
      errline = 98
      MagebombStage = 0
      errline = 99
      armageddonTimer.enabled = False
      errline = 100
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(gtc - MagebombStartTime) & " ms : The magebomb attack secuence have finished.")
      errline = 101
      MagebombLeader = 0
      DoEvents
      errline = 102
    End If
    errline = 103
  End If
  errline = 104
  Exit Sub
gotErr:
  MagebombStage = 0
  armageddonTimer.enabled = False
  LogOnFile "errors.txt", "An error happened at line " & CStr(errline) & " in ProcessArmageddon() , with code " & CStr(Err.Number) & " and description " & Err.Description
End Sub
Private Sub armageddonTimer_Timer()
  ProcessArmageddon
End Sub

Public Sub SendClientless(ByVal Index As Long, ByRef packet() As Byte, ignoreEncryption As Boolean)
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim extrab As Long
  Dim i As Long
  Dim rnumber As Byte
  Dim totalLong As Long
  Dim goodPacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim LocalUseCrackd As Boolean
  Dim lngTest As Long
  Dim strTest As String
  Dim errline As Long
  Dim onlygood As Long
  Dim fourBytesCRC(3) As Byte
  Dim thedamnCRC As Long
  errline = 0
  If MagebombLeader = 0 Then
    errline = 1
    Exit Sub
  End If
  errline = 2
  ' use encryption?
  If (Magebombs(Index).LoginVersion <= 760) Or (ignoreEncryption = True) Then
    errline = 3
    LocalUseCrackd = False
  Else
    errline = 4
    LocalUseCrackd = True
  End If
  errline = 5
  ' connected?
  lngTest = clientLess(Index).State
  errline = 6
  strTest = Magebombs(Index).CharacterName
  errline = 7
  If clientLess(Index).State <> sckConnected Then
    errline = 8
    Magebombs(Index).ConnectionStatus = 0
    errline = 9
    clientLess(Index).Close
    errline = 10
    DoEvents
    Exit Sub
  End If
  errline = 11
  ' encryption mode
  If (LocalUseCrackd = True) Then
    errline = 12
    
    If TibiaVersionLong >= 830 Then
    
        totalLong = GetTheLong(packet(0), packet(1))
        onlygood = totalLong + 2
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
    
        ReDim goodPacket(totalLong + 7)
        hbytes = UBound(packet) + 1
        RtlMoveMemory goodPacket(6), packet(0), (onlygood)
        goodPacket(0) = LowByteOfLong(UBound(goodPacket) - 1)
        goodPacket(1) = HighByteOfLong(UBound(goodPacket) - 1)
        pres = EncipherTibiaProtectedSP(goodPacket(0), Magebombs(Index).key(0), UBound(goodPacket), UBound(Magebombs(Index).key))
        If (pres < 0) Then
          errline = 24
          frmMain.GiveCrackdDllErrorMessage pres, goodPacket, Magebombs(Index).key, UBound(goodPacket), UBound(Magebombs(Index).key), 22
          Exit Sub
        End If
        thedamnCRC = GetTibiaCRC(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
        longToBytes fourBytesCRC, thedamnCRC
        goodPacket(2) = fourBytesCRC(0)
        goodPacket(3) = fourBytesCRC(1)
        goodPacket(4) = fourBytesCRC(2)
        goodPacket(5) = fourBytesCRC(3)
        clientLess(Index).SendData goodPacket
        DoEvents
    Else
        totalLong = GetTheLong(packet(0), packet(1))
        errline = 13
        extrab = 8 - ((totalLong + 2) Mod 8)
        errline = 14
        If extrab < 8 Then
          errline = 15
          totalLong = totalLong + extrab
        End If
        errline = 16
        totalLong = totalLong + 2
        errline = 17
        ReDim goodPacket(totalLong + 1)
        errline = 18
        hbytes = UBound(packet) + 1
        errline = 19
        RtlMoveMemory goodPacket(2), packet(0), (totalLong)
        errline = 20
        goodPacket(0) = LowByteOfLong(totalLong)
        errline = 21
        goodPacket(1) = HighByteOfLong(totalLong)
        errline = 22
        pres = EncipherTibiaProtected(goodPacket(0), Magebombs(Index).key(0), UBound(goodPacket), UBound(Magebombs(Index).key))
        errline = 23
        If (pres < 0) Then
          errline = 24
          frmMain.GiveCrackdDllErrorMessage pres, goodPacket, Magebombs(Index).key, UBound(goodPacket), UBound(Magebombs(Index).key), 10
          Exit Sub
        End If
        errline = 25
        clientLess(Index).SendData goodPacket
        DoEvents
    End If
  Else ' old mode, for OTs
    errline = 26
    clientLess(Index).SendData packet
    DoEvents
    errline = 27
  End If
  errline = 28
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "An error happened in SendClientless() sub AT LINE " & CStr(errline) & " , with code " & CStr(Err.Number) & " and description " & Err.Description
End Sub


Private Sub chkRecordLogins_Click()
If chkRecordLogins.Value = 1 Then
  RecordLogin = True
Else
  RecordLogin = False
End If
End Sub





Private Sub clientLess_Close(Index As Integer)
  Dim aRes As Long
  Dim dRes As Long
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  If (Magebombs(Index).ConnectionStatus > 0) Then
    Magebombs(Index).ConnectionStatus = 0
    If DebugingMagebomb = True Then
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : " & Magebombs(Index).CharacterName & "  lost connection")
      DoEvents
    End If
  End If
  Exit Sub
gotErr:
  aRes = -1
End Sub

Private Sub clientLess_Connect(Index As Integer)
  Dim aRes As Long
  Dim dRes As Long
  Dim i As Long
  Dim okToContinue As Boolean
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  If (Magebombs(Index).ConnectionStatus <= 1) Then
    Magebombs(Index).ConnectionStatus = 2
    If DebugingMagebomb = True Then
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : " & Magebombs(Index).CharacterName & "  is now connected.")
      DoEvents
    End If
    okToContinue = True
    For i = 1 To (MagebombsLoaded - 1)
      If Magebombs(i).ConnectionStatus <> 2 Then
        okToContinue = False
      End If
    Next i
      If okToContinue = True Then
        DoEvents
        'ProcessArmageddon
      End If
  ElseIf Magebombs(Index).ConnectionStatus = 2 Then
       Magebombs(Index).ConnectionStatus = 3
       Magebombs(i).nextSendLogin = GetTickCount + CLng(frmMagebomb.txtReloginChar.Text)
       SendClientless CInt(i), Magebombs(i).loginPacket, True
       DoEvents
       If DebugingMagebomb = True Then
           dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : " & Magebombs(Index).CharacterName & "  is now sending login again")
           DoEvents
       End If
  Else
    clientLess(Index).Close
    Magebombs(Index).ConnectionStatus = 0
    If DebugingMagebomb = True Then
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : " & Magebombs(Index).CharacterName & "  did an unexpected connection, closing it.")
      DoEvents
    End If
  End If
  Exit Sub
gotErr:
  aRes = -1
End Sub



Private Sub clientLess_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  Dim dRes As Long
  Dim packet() As Byte 'a tibia packet is an array of bytes
  Dim res As Integer
  Dim pres As Long
  clientLess(Index).GetData packet, vbArray + vbByte
  If Magebombs(Index).ConnectionStatus < 4 Then
    Magebombs(Index).ConnectionStatus = 4
    Magebombs(Index).ConnectionTimeout = GetTickCount() + Magebombs(Index).RetryTime
    If DebugingMagebomb = True Then
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : " & Magebombs(Index).CharacterName & "  is now inside. Starting attack against " & Magebombs(Index).TargetToShot)
      DoEvents
    End If
    DoEvents
  End If
  Exit Sub
errclose:
  res = -1
End Sub



Private Sub clientLess_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      Dim dRes As Long
      Magebombs(Index).ConnectionStatus = 0
      If DebugingMagebomb = True Then
      dRes = SendLogSystemMessageToClient(MagebombLeader, CStr(GetTickCount() - MagebombStartTime) & " ms : Error " & CStr(Number) & " in " & Magebombs(Index).CharacterName & " : " & Description)
      DoEvents
      End If
End Sub

Private Sub cmdAddMagebomb_Click()
  Dim res As Long
  res = PreloadAbomb(txtFile.Text, cmbOrderType.Text, txtTarget.Text, txtRetryTime.Text)
  Select Case res
  Case 0
    lblStatus.Caption = "Status: Added character successfully!"
    lblStatus.ForeColor = &HC000&
  Case -1
    lblStatus.Caption = "Status: Unexpected error detected!"
    lblStatus.ForeColor = &HFF&
  Case -2
    lblStatus.Caption = "Status: Invalid login log selected"
    lblStatus.ForeColor = &HFF&
  Case -3
    lblStatus.Caption = "Status: Invalid attack mode selected"
    lblStatus.ForeColor = &HFF&
  Case -4
    lblStatus.Caption = "Status: Invalid target selected"
    lblStatus.ForeColor = &HFF&
  Case -5
    lblStatus.Caption = "Status: Invalid time selected"
    lblStatus.ForeColor = &HFF&
  Case -6
    lblStatus.Caption = "Status: Warning - You added the same character again."
    lblStatus.ForeColor = &HFFFF&
  Case -7
    lblStatus.Caption = "Status: Can't read files - missing libraries"
    lblStatus.ForeColor = &HFF&
  Case -8
    lblStatus.Caption = "Status: File doesn't exist"
    lblStatus.ForeColor = &HFF&
  Case -9
    lblStatus.Caption = "Status: An Error happened while reading the file"
    lblStatus.ForeColor = &HFF&
  Case Else
    lblStatus.Caption = "Status: Total failure detected in PreloadAbomb function!"
    lblStatus.ForeColor = &HFF&
  End Select
End Sub

Private Sub cmdClearMemory_Click()
  If MagebombsLoaded = 0 Then
    lblStatus.Caption = "Status: Warning - The magebomb memory was already empty."
    lblStatus.ForeColor = &HFFFF&
  Else
    lblStatus.Caption = "Status: Cleared the magebomb memory successfully"
    lblStatus.ForeColor = &HC000&
  End If
  DeleteAllMagebombMemory
  DisplayMagebombMemory
End Sub

Private Sub cmdDebug_Click()
If DebugingMagebomb = False Then
  DebugingMagebomb = True
  cmdDebug.Caption = "Disable magebomb DEBUG"
  lblStatus.Caption = "Magebomb debug mode is now ENABLED"
  lblStatus.ForeColor = &HC000&
Else
  DebugingMagebomb = False
  cmdDebug.Caption = "Enable  magebomb DEBUG"
  lblStatus.Caption = "Magebomb debug mode is now DISABLED"
  lblStatus.ForeColor = &HC000&
End If
End Sub

Private Sub cmdDeleteSelected_Click()
  Dim firstI As Long
  Dim lasti As Long
  Dim difR As Long
  Dim numofEv As Long
  Dim vrow As Long
  Dim vrowsel As Long
  Dim firstrow As Long
  Dim lastrow As Long
  Dim i As Long
  Dim gotanyerr As Boolean
  vrow = gridBomb.Row
  vrowsel = gridBomb.RowSel
  If vrow > vrowsel Then
    firstrow = vrowsel
    lastrow = vrow
  Else
    firstrow = vrow
    lastrow = vrowsel
  End If
  numofEv = MagebombsLoaded
  If lastrow > MagebombsLoaded Then
    lastrow = MagebombsLoaded
  End If
  If (firstrow > lastrow) Or (MagebombsLoaded = 0) Then
    lblStatus.Caption = "Status: Warning - Invalid selection"
    lblStatus.ForeColor = &HFFFF&
  Else
   firstI = firstrow
   lasti = lastrow
   difR = lasti - firstI + 1
   gotanyerr = False
   For i = 1 To difR
     If (DeleteMagebombMemory(firstI - 1)) = -1 Then
       gotanyerr = True
     End If
   Next i
   DisplayMagebombMemory
   If gotanyerr = True Then
    lblStatus.Caption = "Status: Internal function error. List reseted"
    lblStatus.ForeColor = &HFF&
   Else
    lblStatus.Caption = "Status: Deleted " & CStr(difR) & " items - from " & CStr(firstI) & " to " & CStr(lasti)
    lblStatus.ForeColor = &HC000&
   End If
  End If
End Sub

Private Sub cmdLoadMagebomb_Click()
  Dim res As Integer
  res = LoadMagebombList()
  DisplayMagebombMemory
  Select Case res
  Case 0
    lblStatus.Caption = "Status: Loaded list successfully"
    lblStatus.ForeColor = &HC000&
  Case -1
    lblStatus.Caption = "Status: Function failed. Unable to load"
    lblStatus.ForeColor = &HFF&
  Case -2
    lblStatus.Caption = "Status: List file not found"
    lblStatus.ForeColor = &HFF&
  Case -3
    lblStatus.Caption = "Status: Warning - One or more logs are missing"
    lblStatus.ForeColor = &HFFFF&
  Case Else
    lblStatus.Caption = "Status: Unexpected error. Unable to load"
    lblStatus.ForeColor = &HFF&
  End Select
End Sub

Private Sub cmdReloadFiles_Click()
  ReloadMagebombFiles
End Sub

Public Sub DisplayMagebombMemory()
  Dim i As Long
  gridBomb.Rows = MagebombsLoaded + 1
  For i = 1 To MagebombsLoaded
    With gridBomb
    .TextMatrix(i, 0) = Magebombs(i - 1).CharacterName
    .TextMatrix(i, 1) = Magebombs(i - 1).AttackMode
    .TextMatrix(i, 2) = Magebombs(i - 1).TargetToShot
    .TextMatrix(i, 3) = CStr(Magebombs(i - 1).RetryTime)
    .Row = i
    .Col = 0
    .CellAlignment = flexAlignCenterCenter
    .Col = 1
    .CellAlignment = flexAlignCenterCenter
    .Col = 2
    .CellAlignment = flexAlignCenterCenter
    .Col = 3
    .CellAlignment = flexAlignCenterCenter
    End With
  Next i
End Sub


Public Function SaveMagebombList() As Integer
  Dim fn As Integer
  Dim strLine As String
  Dim res As Integer
  Dim i As Integer
  #If FinalMode Then
  On Error GoTo justend
  #End If
  res = -1
  fn = FreeFile
  Open App.path & "\" & txtFileName.Text For Output As #fn
    Print #fn, CStr(MagebombsLoaded)
    For i = 0 To (MagebombsLoaded - 1)
      Print #fn, Magebombs(i).LogFileName
      Print #fn, Magebombs(i).AttackMode
      Print #fn, Magebombs(i).TargetToShot
      Print #fn, CStr(Magebombs(i).RetryTime)
    Next i
  Close #fn
  res = 0
justend:
  SaveMagebombList = res
End Function

Public Function LoadMagebombList() As Integer
  Dim fn As Integer
  Dim fn2 As Integer
  Dim strTmp As String
  Dim res As Integer
  Dim i As Long
  Dim numberToLoad As Long
  Dim TheLogName As String
  Dim TheAttackMode As String
  Dim TheTarget As String
  Dim TheTime As Long
  Dim AddingCharname As String
  Dim AddingVersion As Long
  Dim AddingIP As String
  Dim AddingPort As Long
  Dim AddingRawKey As String
  Dim ubLoginPacket As Long
  Dim AddingRawLoginPacket As String
  #If FinalMode Then
  On Error GoTo justend
  #End If
  res = -1
  Dim fso As scripting.FileSystemObject
  Set fso = New scripting.FileSystemObject
  DeleteAllMagebombMemory
  If (fso.FileExists(App.path & "\" & txtFileName.Text) = False) Then
    LoadMagebombList = -2
    Exit Function
  End If
  fn = FreeFile
  Open App.path & "\" & txtFileName.Text For Input As #fn
    Line Input #fn, strTmp
    numberToLoad = CLng(strTmp)
    For i = 1 To numberToLoad
      Line Input #fn, TheLogName
      Line Input #fn, strTmp
      TheAttackMode = strTmp
      Line Input #fn, TheTarget
      Line Input #fn, strTmp
      TheTime = CLng(strTmp)
      If (fso.FileExists(App.path & "\magebomb\" & TheLogName) = False) Then
        Close #fn
        LoadMagebombList = -3
        Exit Function
      Else
      fn2 = FreeFile
      Open App.path & "\magebomb\" & TheLogName For Input As #fn2
        Line Input #fn2, AddingCharname
        Line Input #fn2, strTmp
        AddingVersion = CLng(strTmp)
        Line Input #fn2, AddingIP
        Line Input #fn2, strTmp
        AddingPort = CLng(strTmp)
        Line Input #fn2, AddingRawKey
        Line Input #fn2, strTmp
        ubLoginPacket = CLng(strTmp)
        Line Input #fn2, AddingRawLoginPacket
      Close #fn2
      AddToMagebombMemory TheLogName, AddingCharname, AddingVersion, AddingIP, AddingPort, AddingRawKey, ubLoginPacket, AddingRawLoginPacket, TheAttackMode, TheTarget, TheTime
      End If
    Next i
  Close #fn
  LoadMagebombList = 0
  Exit Function
justend:
  DeleteAllMagebombMemory
  LoadMagebombList = -1
End Function


Private Sub cmdSaveMagebomb_Click()
  Dim res As Integer
  res = SaveMagebombList()
  If res = 0 Then
    lblStatus.Caption = "Status: Saved list successfully"
    lblStatus.ForeColor = &HC000&
  Else
    lblStatus.Caption = "Status: Failed to save list"
    lblStatus.ForeColor = &HFF&
  End If
End Sub

Private Sub Form_Load()
  cmbOrderType.Clear
  cmbOrderType.AddItem "type 5 : SD (battlelist)", 0
  cmbOrderType.AddItem "type 6 : HMM (battlelist)", 1
  cmbOrderType.AddItem "type 7 : Explosion (battlelist)", 2
  cmbOrderType.AddItem "type 8 : IH (battlelist)", 3
  cmbOrderType.AddItem "type 9 : UH (battlelist)", 4
  cmbOrderType.AddItem "type B : fireball (battlelist)", 5
  cmbOrderType.AddItem "type C : stalagmite (battlelist)", 6
  cmbOrderType.AddItem "type D : icicle (battlelist)", 7

  cmbOrderType.Text = "type 5 : SD (battlelist)"
  With gridBomb
  .ColWidth(0) = 2000
  .ColWidth(1) = 500
  .ColWidth(2) = 2000
  .ColWidth(3) = 800
  .TextMatrix(0, 0) = "Character"
  .TextMatrix(0, 1) = "Mode"
  .TextMatrix(0, 2) = "Target"
  .TextMatrix(0, 3) = "Time"
  .Row = 0
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignCenterCenter
  .Col = 3
  .CellAlignment = flexAlignCenterCenter
  End With
  ReloadMagebombFiles
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
  Me.Hide
  Cancel = BlockUnload
  If Cancel = False Then
    For i = 0 To clientLess.UBound
      clientLess(i).Close
      'Unload clientLess(i)
    Next i
  End If
End Sub



