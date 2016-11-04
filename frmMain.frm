VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection and logs"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerTibiaQ 
      Interval        =   500
      Left            =   5760
      Top             =   6480
   End
   Begin MSWinsockLib.Winsock SckHttpGet 
      Index           =   0
      Left            =   3240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkForceLoginServer 
      BackColor       =   &H00000000&
      Caption         =   "Force login server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   6720
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sckServerGame 
      Index           =   0
      Left            =   5520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClientGame 
      Index           =   0
      Left            =   5040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   4440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SckClient 
      Index           =   0
      Left            =   3960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkBlockRemote 
      BackColor       =   &H00000000&
      Caption         =   "Block remote connections"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   39
      Top             =   960
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Timer timeToSpam 
      Interval        =   2000
      Left            =   6000
      Top             =   6120
   End
   Begin VB.CommandButton cmbBrowse 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox txtTibiaPath 
      Height          =   375
      Left            =   1200
      TabIndex        =   36
      Text            =   "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->"
      Top             =   5640
      Width           =   5055
   End
   Begin VB.ComboBox cmbPrefered 
      Height          =   315
      Left            =   2520
      TabIndex        =   35
      Text            =   "server.tibia.com"
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   6360
   End
   Begin VB.CheckBox chkAutoHide 
      BackColor       =   &H00000000&
      Caption         =   "Hide 2nd logger when log packets is disabled"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6960
      TabIndex        =   21
      Top             =   6840
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtLogFile 
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Text            =   "log.txt"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.OptionButton TrueServer3 
      BackColor       =   &H00000000&
      Caption         =   "Forward to OT server"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox chckAlter 
      BackColor       =   &H00000000&
      Caption         =   "Change character list packets (use proxy for game connection)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   6480
      Value           =   1  'Checked
      Width           =   5295
   End
   Begin VB.Frame frLogger 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6840
      TabIndex        =   33
      Top             =   3600
      Width           =   1455
      Begin VB.OptionButton LogFull3 
         BackColor       =   &H00000000&
         Caption         =   "Log to file and clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton LogFull1 
         BackColor       =   &H00000000&
         Caption         =   "Clear (faster)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton LogFull2 
         BackColor       =   &H00000000&
         Caption         =   "Delete first line (slow)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAdvanced 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show advanced options"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox ForwardGameTo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      MaxLength       =   200
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.OptionButton TrueServer2 
      BackColor       =   &H00000000&
      Caption         =   "Forward to bouncer"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.OptionButton TrueServer1 
      BackColor       =   &H00000000&
      Caption         =   "Forward to true servers"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear logs"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtMaxLines 
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Text            =   "3000"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtMaxChar 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Text            =   "30000"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtServerGameP 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Text            =   "16000"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtServerLoginP 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Text            =   "15000"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtClientGameP 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Text            =   "16000"
      Top             =   7290
      Width           =   735
   End
   Begin VB.TextBox txtPackets 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmMain.frx":0442
      Top             =   2160
      Width           =   6615
   End
   Begin VB.TextBox txtClientLoginP 
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Text            =   "15000"
      Top             =   7050
      Width           =   735
   End
   Begin VB.CheckBox chckMemoryIP 
      BackColor       =   &H00000000&
      Caption         =   "Change all Tibia clients memory so they can login to this proxy"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Value           =   1  'Checked
      Width           =   5295
   End
   Begin VB.CheckBox chkSelect 
      BackColor       =   &H00000000&
      Caption         =   "Auto Select Hex <-> Ascii"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6960
      TabIndex        =   20
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkLogPackets 
      BackColor       =   &H00000000&
      Caption         =   "Log Packets"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid gridLog 
      Height          =   2055
      Left            =   120
      TabIndex        =   40
      Top             =   3480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      Cols            =   21
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.Label lblTibiaPath 
      BackColor       =   &H00000000&
      Caption         =   "Maps Path:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00000000&
      Caption         =   "Warning: don't use the same port for login and game !"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3000
      TabIndex        =   31
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblLogTo 
      BackColor       =   &H00000000&
      Caption         =   "Log file:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblWhenAloggerIsFull 
      BackColor       =   &H00000000&
      Caption         =   "When a logger is full, take this action:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6960
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblWarning2 
      BackColor       =   &H00000000&
      Caption         =   "Warning: this is ignored since Tibia 8.41 Using random ports"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   32
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label lblMaxHexLines 
      BackColor       =   &H00000000&
      Caption         =   "Max hex lines:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line LineAdv 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   8280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblMaxTextChar 
      BackColor       =   &H00000000&
      Caption         =   "Max text characters:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblAdvanced 
      BackColor       =   &H00000000&
      Caption         =   "Advanced options:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblListenGameServer 
      BackColor       =   &H00000000&
      Caption         =   "Listen game server connections in this port"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label lblGamePort 
      BackColor       =   &H00000000&
      Caption         =   "game port:"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblEnterOtherComputerIP 
      BackColor       =   &H00000000&
      Caption         =   "Enter other proxy IP ..."
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblListenLoginServer 
      BackColor       =   &H00000000&
      Caption         =   "Listen login server connections in this port"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   7080
      Width           =   3495
   End
   Begin VB.Label lblLoginPort 
      BackColor       =   &H00000000&
      Caption         =   "login port:"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
#Const BufferDebug = 0

Option Explicit


Private Const OptCte4 = 1 'internal size of a byte ( 1 byte )
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal ByValcbCopy As Long)
'Public LoginMethod As Integer

Public userHere As String

Private fastestconnect As Long

Private fastestLoginServerTime As Long

Private lastLoadLine As Long
Const IPPROTO_TCP = 6         ' Protocol constant for TCP.
Const TCP_NODELAY = &H1&      ' Turn off Nagel Algorithm.
Private Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Const errIndexOutOfRange = -1610350521

Dim HTTPGetConnected() As Boolean
Dim HTTPGetResponseBuffer() As String



'Private Function getFasterLoginServer() As String
'    Dim i As Long
'    Dim idLoginSP As Long
'    Dim maxgtc As Long
'    Dim firstGTC As Long
'    If AlreadyCheckingFasterLogin = True Then
'        While AlreadyCheckingFasterLogin = True
'            DoEvents
'        Wend
'        getFasterLoginServer = LastFasterLogin
'        Exit Function
'    End If
'    AlreadyCheckingFasterLogin = True
'    For i = 0 To NumberOfLoginServers - 1
'        If i > sckFasterLogin.UBound Then
'            Load sckFasterLogin(sckFasterLogin.UBound + 1)
'        End If
'    Next i
'    For i = 0 To NumberOfLoginServers - 1
'        idLoginSP = i + 1
'        If trueLoginPort(idLoginSP) = 7171 Then
'            sckFasterLogin(i).Close
'        End If
'    Next i
'    DoEvents
'    firstGTC = GetTickCount()
'    fastestconnect = -1
'    For i = 0 To NumberOfLoginServers - 1
'        idLoginSP = i + 1
'        If trueLoginPort(idLoginSP) = 7171 Then
'            sckFasterLogin(i).RemoteHost = trueLoginServer(idLoginSP)
'            sckFasterLogin(i).RemotePort = trueLoginPort(idLoginSP)
'            sckFasterLogin(i).Connect
'        End If
'    Next i
'    DoEvents
'    maxgtc = GetTickCount() + 10000
'    Do
'        DoEvents
'    Loop Until ((fastestconnect > -1) Or (GetTickCount() > maxgtc))
'    fastestLoginServerTime = GetTickCount() - firstGTC
'    For i = 0 To NumberOfLoginServers - 1
'        idLoginSP = i + 1
'        If trueLoginPort(idLoginSP) = 7171 Then
'            sckFasterLogin(i).Close
'        End If
'    Next i
'    DoEvents
'    If fastestconnect = -1 Then
'        fastestconnect = 0
'    End If
'
'    LastFasterLogin = trueLoginServer(fastestconnect + 1)
'    AlreadyCheckingFasterLogin = False
'    getFasterLoginServer = LastFasterLogin
'End Function

'Private Function getBlackdINI(ByRef par1 As String, ByRef par2 As String, _
' ByRef par3 As String, ByRef par4 As String, ByRef par5 As Long, ByRef par6 As String) As Long
'
'  If ((par1 = "MemoryAddresses") Or (par1 = "tileIDs") Or (par2 = "configPath")) Then
'    getBlackdINI = GetPrivateProfileString(par1, par2, par3, par4, par5, par6)
'  Else
'    getBlackdINI = GetPrivateProfileString(par1, par2, par3, par4, par5, App.path & "\settings.ini")
'  End If
'End Function

' BLAKCKDINI FUNCTIONS MOVED TO MODCODE

Private Sub cmbBrowse_Click()
  ConfigurePath Me.hwnd, True
End Sub

Public Sub GiveCrackdDllErrorMessage(pres As Long, ByRef packet1() As Byte, ByRef packet2() As Byte, ubound1 As Long, ubound2 As Long, p As Long)
  Dim errorstr As String
        Select Case pres
        Case -1
          errorstr = "ERROR -1 : Packet header is not multiplier of 8"
        Case -2
          errorstr = "ERROR -2 : Wrong size of key (ubound must be 15)"
        Case -3
          errorstr = "ERROR -3 : Header of packet doesn't match with real size of the packet"
        Case -4
          errorstr = "ERROR -4 : This is not a packet"
        Case Else
          errorstr = "ERROR " & CStr(pres) & " : Unknown error"
        End Select
  errorstr = errorstr & vbCrLf & "PARAMETERS:" & vbCrLf & _
    "Packet : " & showAsStr2(packet1, 2) & vbCrLf & _
    "Key : " & showAsStr2(packet2, 2) & vbCrLf & _
    "Ubound(Packet) : " & CStr(ubound1) & vbCrLf & _
    "Ubound(Key) : " & CStr(ubound2) & vbCrLf & _
    "Called at point : " & CStr(p)
  LogOnFile "errors.txt", errorstr
End Sub
Private Function InitSounds(thehwnd As Long) As Boolean
  #If FinalMode Then
  On Error GoTo gotserr
  #End If
  Dim bRes As Boolean
  Dim loadingThisSound As String
  soundErrorLine = "<nothing>"
  SoundErrorWasThis = "Executing: " & soundErrorLine & vbCrLf & "Got error number " & CStr(0) & " : " & "<no error>"
  bRes = DirectX_Init(thehwnd, 4)
  If bRes = True Then
    soundErrorLine = "loadingThisSound = App.Path & ""\player.wav"""
    loadingThisSound = App.Path & "\player.wav"
    soundErrorLine = "DirectX_LoadSound loadingThisSound, 1"
    DirectX_LoadSound loadingThisSound, 1
    soundErrorLine = "loadingThisSound = App.Path & ""\danger.wav"""
    loadingThisSound = App.Path & "\danger.wav"
    soundErrorLine = "DirectX_LoadSound loadingThisSound, 2"
    DirectX_LoadSound loadingThisSound, 2
    soundErrorLine = "loadingThisSound = App.Path & ""\ding.wav"""
    loadingThisSound = App.Path & "\ding.wav"
    soundErrorLine = "DirectX_LoadSound loadingThisSound, 3"
    DirectX_LoadSound loadingThisSound, 3
    'custom ng
    soundErrorLine = "loadingThisSound = App.Path & ""\newsound1.wav"""
    loadingThisSound = App.Path & "\newsound1.wav"
    soundErrorLine = "DirectX_LoadSound loadingThisSound, 4"
    DirectX_LoadSound loadingThisSound, 4
    InitSounds = True
  Else
    InitSounds = False
  End If
  Exit Function
gotserr:
  SoundErrorWasThis = "Executing: " & soundErrorLine & vbCrLf & "Got error number " & CStr(Err.Number) & " : " & Err.Description
  InitSounds = False
End Function
Private Sub Form_Load()
  ' HERE IS WHERE ALL START
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim res As Integer
  Dim resT As TypeTrial
  Dim tmp As Long
  Dim loadingThisSound As String
  Dim bRes As Boolean
  Dim str As String
  Dim str2 As String
  Dim a As Integer
  Dim dblAmLoaded As Double
  Dim strHint As String
  Dim tibiadathere As String
  Dim trythis As String
  Dim prevValue As String
  Dim lngTemp As Long
  Dim moreDetails As String
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  dblAmLoaded = 0
  lastLoadLine = 1
  strHint = ""
  If thisShouldNotBeLoading = 0 Then
    Unload Me
    Exit Sub
  End If
  For i = 1 To MAXCLIENTS
    Load sckClientGame(i)
    Load sckServerGame(i)
  Next i
  loadingThisSound = "nothing"
  ExivaExpPlace = "19 : white center"
  ' Don't unload child forms while father is alive
  BlockUnload = 1
  
  ' Init some memory protection vars.
  ' (This make the code confusing for asm readers)
  trialSafety1 = 1 ' 0 if trial version, 1 if full
  trialSafety2 = 2  ' 0 if trial version , 2 if full
  trialSafety300 = 300   ' 0 if trial version , 300 if full
  trialSafety4 = 4 ' 0 if trial version , 4 if full
  dblAmLoaded = 5
  frmLoading.NotifyLoadProgress dblAmLoaded, "Starting randomizer"
  DoEvents
  ' Init random number generator
  Randomize

  ' Load sounds in RAM memory
  If (SoundIsUsable = True) Then
  dblAmLoaded = 15
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading sound in RAM memory"
  If InitSounds(Me.hwnd) = False Then
    SoundIsUsable = False
    If MsgBox("Sorry, Blackd Proxy got an error while trying to initialize the Directx sound system" & vbCrLf & _
      SoundErrorWasThis & vbCrLf & vbCrLf & _
      "Possible reasons:" & vbCrLf & _
      "- system32\dx7vb.dll is missing or corrupted (then you will get a 429 error later)" & vbCrLf & _
      "- you miss a sound card or your sound card drivers are corrupted" & vbCrLf & _
     "Do you want to continue anyways? (you won't listen the sound of the alarms)", vbYesNo + vbInformation, "There is a problem with your sound card") = vbNo Then
      End
      Exit Sub
    End If
  Else
    SoundIsUsable = True
  End If
  End If
  dblAmLoaded = 20

  ' Removed most of the code here.
  ' (trial version have no sense in a free Blackd Proxy)
  TrialMode = 2
  TrialLimit_Day = -1
  lastLockReason = ""
  
  ' Do another anti memory / packet modifier check
  ' to avoid people cracking this program
'  If TrialVersion = True Then
'    If DoTheCheck() = True Then
'      'LogOnFile "debug.txt", "Terminated by protection functions (-1)"
'      End
'      Exit Sub
'    End If
'  End If
  
  ' Optional for versions given to work in only 1 computer:
'  If DoMAC_check = True Then
'    ' if computer MAC address doesn't match, close
'    If DoMACAddressCompare() = False Then
'      'LogOnFile "debug.txt", "Terminated by protection functions (-2)"
'      End
'      Exit Sub
'    End If
'  End If
  dblAmLoaded = 25
  frmLoading.NotifyLoadProgress dblAmLoaded, "Linking modules"
  ' Init some vars
  MagebombsLoaded = 0
  givenUFO = False ' used to debug truemap. Avoids giving certain strange error more than 1 time.
  runemakerIDselected = 0 ' current runemaker character selected
  'custom ng
  healingIDselected = 0
  extrasIDselected = 0
  persistentIDselected = 0
  aimbotIDselected = 0
  '
  PUSHDELAYTIMES = 9 ' ticks between 2 push in exiva push
  dblAmLoaded = 30
  frmLoading.NotifyLoadProgress dblAmLoaded, "Reading high priority info"
  ' Read high priority vars from config.ini
  If ReadIniThisFirst() = -1 Then
    End
  End If
  GotTrialLock = False ' No trial lock happened yet
  publicDebugMode = False ' Debugging = false . This can be changed by users.
 
  dblAmLoaded = 50
  frmLoading.NotifyLoadProgress dblAmLoaded, "Processing environment"
  ' Add some text in main window if this is a trial version:
  If TrialVersion = True Then
    If TrialMode = 1 Then
      txtPackets.Text = "SHORT TRIAL VERSION - " & txtPackets.Text
    Else
      txtPackets.Text = "MONTH TRIAL VERSION - " & txtPackets.Text
    End If
  End If
  dblAmLoaded = 55
  frmLoading.NotifyLoadProgress dblAmLoaded, "Reseting all globals"
  ResetCharServer
  bpIDselected = 0 ' current container ID selected
  
  LastFasterLogin = ""
  AlreadyCheckingFasterLogin = False
  
  ' this will lock events while changing checkboxes values
  avoidC = False
  lock_chkActivate = False
  lock_chkFood = False
  lock_chkManaFluid = False
  lock_chkLogoutDangerAny = False
  lock_chkLogoutDangerCurrent = False
  lock_chkLogoutOutRunes = False
  lock_chkWaste = False
  lock_chkmsgSound = False
  lock_chkmsgSound2 = False
  ChangePlayTheDangerSound False
  PlayMsgSound = False
  PlayMsgSound2 = False
  'custom ng
  PlayPMSound = False
  lock_chkMana = False
  lock_chkDanger = False
  lock_chkPM = False
  lock_chkEat = False
  lock_chkautoUtamo = False
  lock_chkautoHur = False
  lock_chkautogHur = False
  lock_chkAFK = False
  lock_chkGold = False
  lock_chkPlat = False
  lock_chkDash = False
  lock_chkMW = False
  lock_chkSSA = False
  lock_chkHouse = False
  lock_chkTitle = False
  lock_chkExiva1 = False
  lock_chkExiva2 = False
  lock_chkExiva3 = False
  lock_chkExiva4 = False
  lock_chkExiva5 = False
  lock_chkExiva6 = False
  lock_chkExiva7 = False
  lock_chkExiva8 = False
  lock_chkExiva9 = False
  lock_chkExiva10 = False
  lock_chkExiva11 = False
  lock_chkSDcombo = False
  lock_chkUEcombo = False
  '
  
  ResetSpamOrders ' Init list of requested actions that require spam (autoUH, autoPush)
  
  dblAmLoaded = 60
  lastLoadLine = 600
  frmLoading.NotifyLoadProgress dblAmLoaded, "Creating the first activex dictionary"
lastLoadLine = 601
  Set GameServerDictionary = New Scripting.Dictionary
  Set GameServerDictionaryDOMAIN = New Scripting.Dictionary
  Set mainTibiaHandle = New Scripting.Dictionary
  Set specialGMnames = New Scripting.Dictionary
  lastLoadLine = 602
  Set ValueOfUservar = New Scripting.Dictionary
  lastLoadLine = 603
  Set ProcessidIPrelations = New Scripting.Dictionary
  Set ProcessidAccountRelations = New Scripting.Dictionary
  Set IgnoredCreatures = New Scripting.Dictionary
  lastLoadLine = 604
  
 dblAmLoaded = 65
 lastLoadLine = 651
  frmLoading.NotifyLoadProgress dblAmLoaded, "Giving default values to all"
  'trainer defaults
  For i = 1 To MAXCLIENTS
    ResetInternalTrainerValues CInt(i)
  Next i
lastLoadLine = 652
SafeModeOutPacket(0) = &H4
SafeModeOutPacket(1) = &H0
SafeModeOutPacket(2) = &HA0
SafeModeOutPacket(3) = &H3
SafeModeOutPacket(4) = &H0
SafeModeOutPacket(5) = &H0
DebugingMagebomb = False
lastLoadLine = 660
  ' Init every connection
  For i = 1 To MAXCLIENTS
    'Init some vars to their empty value
    runeTurn(i) = randomNumberBetween(0, 29)
    CavebotHaveSpecials(i) = False
    CavebotLastSpecialMove(i) = 0
    StatusBits(i) = "0000000000000000"
    lastUsedChannelID(i) = "05 00"
    lastRecChannelID(i) = "05 00"
    makingRune(i) = False
    UHRetryCount(i) = 0
    runemakerMana1(i) = -1
    reconnectionRetryCount(i) = 0
    nextReconnectionRetry(i) = 0
    ResetEventList i
    ResetCondEventList i
    SelfDefenseID(i) = 0
    logoutAllowed(i) = 0
    ReconnectionStage(i) = 0
    ReconnectionPacket(i).numbytes = 0
    pauseStacking(i) = 0
    ResetCharList2 CInt(i)

    AllowUHpaused(i) = False
    doingTrade(i) = False
    doingTrade2(i) = False
    cavebotOnTrapGiveAlarm(i) = False
    GotKillOrderTargetID(i) = 0
    GotKillOrder(i) = False
    GotKillOrderTargetName(i) = ""
    lastAttackedIDstatus(i) = 0
    previousAttackedID(i) = 0
    posSpamActivated(i) = False
    posSpamChannelB1(i) = &HFF
    posSpamChannelB2(i) = &HFF
    executingCavebot(i) = False
    lastLoadLine = 670
    getSpamActivated(i) = False
    getSpamChannelB1(i) = &HFF
    getSpamChannelB2(i) = &HFF
    nextAllowedmsg(i) = 0
    DelayAttacks(i) = 0
    AvoidReAttacks(i) = True
    IgnoreServer(i) = False
    FirstCharInCharList(i) = ""
    NoHealingNextTurn(i) = False
    DropDelayerTurn(i) = 0
    var_expleft(i) = ""
    var_nextlevel(i) = ""
    var_exph(i) = ""
    var_timeleft(i) = ""
    var_played(i) = ""
    var_expgained(i) = ""
    var_lf(i) = vbLf
    var_lastsender(i) = ""
    var_lastmsg(i) = ""
    initialRuneBackpack(i) = &HFF
    RequiredMoveBuffer(i) = ""
    ReadyBuffer(i) = True
    lastLoadLine = 680
    ReDim ConnectionBuffer(i).packet(0)
    ReDim ConnectionBufferLogin(i).packet(0)
    
    CheatsPaused(i) = False
    
    LoginMsgCount(i) = 0
    lastHPchange(i) = 0
    cancelAllMove(i) = 0
    ConnectionBuffer(i).numbytes = 0
    ConnectionBufferLogin(i).numbytes = 0
    lastFloorTrap(i) = -1
    DoingMainLoop(i) = False
    DoingMainLoopLogin(i) = False
    nextForcedDepotDeployRetry(i) = 0
    nextLight(i) = "D7" ' default light colour (215 = D7 in hex)
    lastDestX(i) = 0
    lastDestY(i) = 0
    lastDestZ(i) = 0
    ignoreNext(i) = 0
    somethingChangedInBps(i) = False
    onDepotPhase(i) = 0
    CavebotChaoticMode(i) = 0
    TurnsWithRedSquareZero(i) = 0
    
    bLevelSpy(i) = False
    depotX(i) = 0
    depotY(i) = 0
    depotZ(i) = 0
    doneDepotChestOpen(i) = False
    depotTileID(i) = 0
    depotS(i) = 0
    lastDepotBPID(i) = 0
    receivedLogin(i) = False
    friendlyMode(i) = 0
    currTargetName(i) = ""
    currTargetID(i) = 0
    SendingSpecialOutfit(i) = False
    DangerGMname(i) = ""
    DangerPKname(i) = ""
    DangerPlayerName(i) = ""
    lootTimeExpire(i) = 0
    requestLootBp(i) = &HFF ' no container requested for being looted
    autoLoot(i) = False
    myLastCorpseX(i) = 0
    myLastCorpseY(i) = 0
    myLastCorpseZ(i) = 0
    myLastCorpseS(i) = 0
    lastIngameCheck(i) = ""
    lastIngameCheckTileID(i) = "00 00"
    myLastCorpseTileID(i) = 0
    lootWaiting(i) = False
    setFollowTarget(i) = True
    lastLoadLine = 690
    moveRetry(i) = 0
    lastX(i) = 0
    lastY(i) = 0
    lastZ(i) = 0
    lastAttackedID(i) = 0
    CavebotTimeWithSameTarget(i) = GetTickCount()
    CavebotTimeStart(i) = GetTickCount()
    maxAttackTime(i) = 40000
    ChaotizeNextMaxAttackTime i
    maxHit(i) = 10000
    previousAttackedID(i) = 0
    cavebotOnDanger(i) = -1
    cavebotOnGMclose(i) = False
    cavebotOnGMpause(i) = False
    cavebotOnPLAYERpause(i) = False
    DangerGM(i) = False
    DangerPK(i) = False
    DangerPlayer(i) = False
    LogoutTimeGM(i) = 0
    GMname(i) = ""
    Connected(i) = False
    GameConnected(i) = False
    MustCheckFirstClientPacket(i) = True
    If TibiaVersionLong >= 841 Then
      NeedToIgnoreFirstGamePacket(i) = True
    Else
      NeedToIgnoreFirstGamePacket(i) = False
    End If
    ConnectionBuffer(i).numbytes = 0
    sentFirstPacket(i) = False
    sentWelcome(i) = False
    IDstring(i) = ""
    myID(i) = 0
    CharacterName(i) = ""
    myX(i) = 0
    myY(i) = 0
    myZ(i) = 7
    LastHealTime(i) = 0
    timeToRetryOpenDepot(i) = 0
lastLootOrder(i) = 0
  
     ResetLooter i
    OldLootMode(i) = True
    ClientExecutingLongCommand(i) = False
    LootAll(i) = False
    PKwarnings(i) = True
    
    LastCavebotTime(i) = 0
    stealthLog(i) = ""
    myHP(i) = cte_initHP ' init HP as 10000, else autoheal might jump as start
    myMaxHP(i) = cte_initHP ' myMaxHP should not be 0, else % of current heal would get a divide by 0 at start
    myMaxMana(i) = cte_initMANA
    lastHPchange(i) = 0
    
    myNewStat(i) = 0
    myMana(i) = 0
    myCap(i) = 0
    myStamina(i) = 0
    myExp(i) = 0
    SpellKillHPlimit(i) = 0
    SpellKillMaxHPlimit(i) = 100
    AllowedLootDistance(i) = 3
    myInitialExp(i) = 0
    myInitialTickCount(i) = 0
    myLevel(i) = 50000000
    myMagLevel(i) = 0
    mySoulpoints(i) = 100
    lastLoadLine = 700
    For k = 1 To EQUIPMENT_SLOTS
      mySlot(i, k).t1 = &H0
      mySlot(i, k).t2 = &H0
      mySlot(i, k).t3 = &H0
    Next k
    savedItem(i).t1 = &H0
    savedItem(i).t2 = &H0
    savedItem(i).t2 = &H0
    AfterLoginLogoutReason(i) = ""
    ProcessID(i) = -1
    'exeLine(i) = 0
    updateExeLine i, 0, False, False
    fishCounter(i) = 0
    pushTarget(i) = 0
    pushDelay(i) = CInt(Int((PUSHDELAYTIMES * Rnd)))
    lastLoadLine = 710
    ' Init internal runemaker options:
    RuneMakerOptions(i).activated = RuneMakerOptions_activated_default
    RuneMakerOptions(i).autoEat = RuneMakerOptions_autoEat_default
    RuneMakerOptions(i).ManaFluid = RuneMakerOptions_ManaFluid_default
    RuneMakerOptions(i).autoLogoutAnyFloor = RuneMakerOptions_autoLogoutAnyFloor_default
    RuneMakerOptions(i).autoLogoutCurrentFloor = RuneMakerOptions_autoLogoutCurrentFloor_default
    RuneMakerOptions(i).autoLogoutOutOfRunes = RuneMakerOptions_autoLogoutOutOfRunes_default
    RuneMakerOptions(i).autoWaste = RuneMakerOptions_autoWaste_default
    RuneMakerOptions(i).msgSound = RuneMakerOptions_msgSound_default
    RuneMakerOptions(i).msgSound2 = RuneMakerOptions_msgSound2_default
    RuneMakerOptions(i).firstActionText = RuneMakerOptions_firstActionText_default
    RuneMakerOptions(i).firstActionMana = RuneMakerOptions_firstActionMana_default
    RuneMakerOptions(i).LowMana = RuneMakerOptions_LowMana_default
    RuneMakerOptions(i).secondActionText = RuneMakerOptions_secondActionText_default
    RuneMakerOptions(i).secondActionMana = RuneMakerOptions_secondActionMana_default
    RuneMakerOptions(i).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default

    'custom ng var
    healingCheatsOptions(i).sdmax = healingCheatsOptions_sdmax_default
    healingCheatsOptions(i).antipush = healingCheatsOptions_antipush_default
    healingCheatsOptions(i).pmax = healingCheatsOptions_pmax_default
    healingCheatsOptions(i).htarget = healingCheatsOptions_htarget_default
    healingCheatsOptions(i).exaustEat = healingCheatsOptions_exaustEat_default
    healingCheatsOptions(i).HouseX = healingCheatsOptions_HouseX_default
    healingCheatsOptions(i).HouseY = healingCheatsOptions_HouseY_default
    
    'custom ng
    'healing
    healingCheatsOptions(i).txtSpellhi = healingCheatsOptions_txtSpellhi_default
    healingCheatsOptions(i).txtSpelllo = healingCheatsOptions_txtSpelllo_default
    healingCheatsOptions(i).txtPot = healingCheatsOptions_txtPot_default
    healingCheatsOptions(i).txtMana = healingCheatsOptions_txtMana_default
    healingCheatsOptions(i).txtHealthhi = healingCheatsOptions_txtHealthhi_default
    healingCheatsOptions(i).txtHealthlo = healingCheatsOptions_txtHealthlo_default
    healingCheatsOptions(i).txtHealpot = healingCheatsOptions_txtHealpot_default
    healingCheatsOptions(i).txtManapot = healingCheatsOptions_txtManapot_default
    healingCheatsOptions(i).txtManahi = healingCheatsOptions_txtManahi_default
    healingCheatsOptions(i).txtManalo = healingCheatsOptions_txtManalo_default
    healingCheatsOptions(i).Combo1 = healingCheatsOptions_Combo1_default
    healingCheatsOptions(i).Combo2 = healingCheatsOptions_Combo2_default
    
    'custom ng
    'extras
    extrasOptions(i).txtSpell = extrasOptions_txtSpell_default
    extrasOptions(i).txtMana = extrasOptions_txtMana_default
    extrasOptions(i).txtSSA = extrasOptions_txtSSA_default
    extrasOptions(i).cmbHouse = extrasOptions_cmbHouse_default
    extrasOptions(i).chkMana = extrasOptions_chkMana_default
    extrasOptions(i).chkDanger = extrasOptions_chkDanger_default
    extrasOptions(i).chkPM = extrasOptions_chkPM_default
    extrasOptions(i).chkEat = extrasOptions_chkEat_default
    extrasOptions(i).chkautoUtamo = extrasOptions_chkautoUtamo_default
    extrasOptions(i).chkautoHur = extrasOptions_chkautoHur_default
    extrasOptions(i).chkautogHur = extrasOptions_chkautogHur_default
    extrasOptions(i).chkAFK = extrasOptions_chkAFK_default
    extrasOptions(i).chkGold = extrasOptions_chkGold_default
    extrasOptions(i).chkPlat = extrasOptions_chkPlat_default
    extrasOptions(i).chkDash = extrasOptions_chkDash_default
    extrasOptions(i).chkMW = extrasOptions_chkMW_default
    extrasOptions(i).chkSSA = extrasOptions_chkSSA_default
    extrasOptions(i).chkTitle = extrasOptions_chkTitle_default
    extrasOptions(i).chkHouse = extrasOptions_chkHouse_default
    
    'custom ng
    'persistent
    persistentOptions(i).txtHk1 = persistentOptions_txtHk1_default
    persistentOptions(i).txtHk2 = persistentOptions_txtHk2_default
    persistentOptions(i).txtHk3 = persistentOptions_txtHk3_default
    persistentOptions(i).txtHk4 = persistentOptions_txtHk4_default
    persistentOptions(i).txtHk5 = persistentOptions_txtHk5_default
    persistentOptions(i).txtHk6 = persistentOptions_txtHk6_default
    persistentOptions(i).txtHk7 = persistentOptions_txtHk7_default
    persistentOptions(i).txtHk8 = persistentOptions_txtHk8_default
    persistentOptions(i).txtHk9 = persistentOptions_txtHk9_default
    persistentOptions(i).txtHk10 = persistentOptions_txtHk10_default
    persistentOptions(i).txtHk11 = persistentOptions_txtHk11_default
    persistentOptions(i).txtExiva1 = persistentOptions_txtExiva1_default
    persistentOptions(i).txtExiva2 = persistentOptions_txtExiva2_default
    persistentOptions(i).txtExiva3 = persistentOptions_txtExiva3_default
    persistentOptions(i).txtExiva4 = persistentOptions_txtExiva4_default
    persistentOptions(i).txtExiva5 = persistentOptions_txtExiva5_default
    persistentOptions(i).txtExiva6 = persistentOptions_txtExiva6_default
    persistentOptions(i).txtExiva7 = persistentOptions_txtExiva7_default
    persistentOptions(i).txtExiva8 = persistentOptions_txtExiva8_default
    persistentOptions(i).txtExiva9 = persistentOptions_txtExiva9_default
    persistentOptions(i).txtExiva10 = persistentOptions_txtExiva10_default
    persistentOptions(i).txtExiva11 = persistentOptions_txtExiva11_default
    persistentOptions(i).chkExiva1 = persistentOptions_chkExiva1_default
    persistentOptions(i).chkExiva2 = persistentOptions_chkExiva2_default
    persistentOptions(i).chkExiva3 = persistentOptions_chkExiva3_default
    persistentOptions(i).chkExiva4 = persistentOptions_chkExiva4_default
    persistentOptions(i).chkExiva5 = persistentOptions_chkExiva5_default
    persistentOptions(i).chkExiva6 = persistentOptions_chkExiva6_default
    persistentOptions(i).chkExiva7 = persistentOptions_chkExiva7_default
    persistentOptions(i).chkExiva8 = persistentOptions_chkExiva8_default
    persistentOptions(i).chkExiva9 = persistentOptions_chkExiva9_default
    persistentOptions(i).chkExiva10 = persistentOptions_chkExiva10_default
    persistentOptions(i).chkExiva11 = persistentOptions_chkExiva11_default
    
    'custom ng
    'aimbot
    aimbotOptions(i).chkSDcombo = aimbotOptions_chkSDcombo_default
    aimbotOptions(i).chkUEcombo = aimbotOptions_chkUEcombo_default
    aimbotOptions(i).txtLeader = aimbotOptions_txtLeader_default
    aimbotOptions(i).txtCombo = aimbotOptions_txtCombo_default
    
    dblAmLoaded = 67
    frmLoading.NotifyLoadProgress dblAmLoaded, "Creating activex dictionaries"
    ' Init dictionary objects:
    lastLoadLine = 720
    Set cavebotScript(i) = New Scripting.Dictionary
    Set cavebotMelees(i) = New Scripting.Dictionary
    Set cavebotAvoid(i) = New Scripting.Dictionary
    Set cavebotExorivis(i) = New Scripting.Dictionary
    Set cavebotHMMs(i) = New Scripting.Dictionary
    Set DictSETUSEITEM(i) = New Scripting.Dictionary
    Set shotTypeDict(i) = New Scripting.Dictionary
    Set exoriTypeDict(i) = New Scripting.Dictionary
    Set cavebotGoodLoot(i) = New Scripting.Dictionary
    Set killPriorities(i) = New Scripting.Dictionary
    Set SpellKills_SpellName(i) = New Scripting.Dictionary
    Set SpellKills_Dist(i) = New Scripting.Dictionary
    Set NameOfID(i) = New Scripting.Dictionary

    Set HPOfID(i) = New Scripting.Dictionary
    Set DirectionOfID(i) = New Scripting.Dictionary
    Set BigMapNamesX = New Scripting.Dictionary
    Set BigMapNamesY = New Scripting.Dictionary
    Set BigMapNamesZ = New Scripting.Dictionary
    Set BigMapNamesC = New Scripting.Dictionary
    Set MapIDTranslator = New Scripting.Dictionary
    Set IgnoredCreatures(i) = New Scripting.Dictionary
    lastLoadLine = 730
    dblAmLoaded = 70
      frmLoading.NotifyLoadProgress dblAmLoaded, "Building the core"
    ' Init some more vars
    cavebotEnabled(i) = False ' cavebot disabled
    GotPacketWarning(i) = False ' safemode = off by default
    cavebotLenght(i) = 0 ' no scripts loaded for every cavebot
     
    ' init open containers for every client:
    For j = 0 To HIGHEST_BP_ID
      Backpack(i, j).open = False
      Backpack(i, j).cap = 0
      Backpack(i, j).used = 0
      Backpack(i, j).name = ""
    Next j
  Next i
  Set EnemyList = New Scripting.Dictionary
  lastLoadLine = 740
  ' blank lines to be writen in map matrix in a block while moving:
  For j = 0 To 10
    tmpStack.s(j).t1 = &H0
    tmpStack.s(j).t2 = &H0
    tmpStack.s(j).t3 = &H0
    tmpStack.s(j).t4 = &H0
    tmpStack.s(j).dblID = &H0
  Next j

  lastLoadLine = 741
  InitGridLog
  lastLoadLine = 742
  'load first server
  LastNumTibiaClients = 0
  ClosedBoard = True
  VisibleAdvancedOptions = False ' advanced options hidden by default
  HideAdvancedOptions
  dblAmLoaded = 73
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: conditional events"
  lastLoadLine = 743
  Load frmCondEvents
  frmCondEvents.timerCheck.enabled = True
  frmCondEvents.Hide
  dblAmLoaded = 74
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: tools"
    lastLoadLine = 744
  Load frmCheats
  frmCheats.Hide
  dblAmLoaded = 75
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: bigtext"
    lastLoadLine = 745
  Load frmBigText
  frmBigText.Hide
  dblAmLoaded = 76
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: advanced"
    lastLoadLine = 746
  Load frmAdvanced
  frmAdvanced.Hide
  dblAmLoaded = 77
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: cheats"
    lastLoadLine = 747
  Load frmHardcoreCheats
  frmHardcoreCheats.Hide
  dblAmLoaded = 78
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: backpacks"
    lastLoadLine = 748
  Load frmBackpacks
  frmBackpacks.UpdateBPlist
  frmBackpacks.Hide
  dblAmLoaded = 79
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: runemaker"
    lastLoadLine = 749
  '
  Load frmRunemaker
  runemakerIDselected = 0
  frmRunemaker.lstFriends.Clear
  frmRunemaker.UpdateValues
  frmRunemaker.Hide
  cavebotIDselected = 0
  dblAmLoaded = 80
  '
  'custom ng healing load
  Load frmHealing
  healingIDselected = 0
  frmHealing.UpdateValues
  frmHealing.Hide
  'custom ng extras load
  Load frmExtras
  extrasIDselected = 0
  frmExtras.UpdateValues
  frmExtras.Hide
  'custom ng persistent load
  Load frmPersistent
  persistentIDselected = 0
  frmPersistent.UpdateValues
  frmPersistent.Hide
  'custom ng aimbot load
  Load frmAimbot
  aimbotIDselected = 0
  frmAimbot.UpdateValues
  frmAimbot.Hide
  
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: hotkeys" 'antes puse cavebot por error
    lastLoadLine = 750
  Load frmHotkeys
  frmHotkeys.Hide
  dblAmLoaded = 81
    lastLoadLine = 751
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: cavebot"
      lastLoadLine = 752
  Load frmCavebot
      lastLoadLine = 753
  frmCavebot.UpdateValues
      lastLoadLine = 754
  frmCavebot.ReloadFiles
      lastLoadLine = 755
  frmCavebot.Hide
      lastLoadLine = 756
  dblAmLoaded = 82
      lastLoadLine = 757
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: events"
      lastLoadLine = 758
  Load frmEvents
  frmEvents.Hide
  dblAmLoaded = 83
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: trainer"
      lastLoadLine = 759
  Load frmTrainer
  frmTrainer.Hide
  dblAmLoaded = 84
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: warbot"
        lastLoadLine = 760
  Load frmWarbot
  frmWarbot.Hide
  dblAmLoaded = 85
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: magebomb"
        lastLoadLine = 761
  Load frmMagebomb
  frmMagebomb.Hide
  dblAmLoaded = 86
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: screenshots"
        lastLoadLine = 762
        
  broadcastIDselected = 0
  currentBroadcastIndex = -1
  Load frmBroadcast

  frmBroadcast.Hide
  dblAmLoaded = 87
  
  Load frmScreenshot

  frmScreenshot.Hide
  dblAmLoaded = 88
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: true map"
        lastLoadLine = 763

  Load frmTrueMap

  mapIDselected = 0
  
  mapFloorSelected = 7
  frmTrueMap.Hide
  LightIntesityHex = "0F"
  HighestConnectionID = 0
  
  
  Load frmConfirm
  frmConfirm.Hide
  
  
  Load frmMapReader
  frmMapReader.SetDefaultMapPosition 32097, 32219, 7
  frmMapReader.Hide
  LoadingAmap = False
  
  
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: news"
        lastLoadLine = 764
  Load frmNews
  frmNews.Hide
  dblAmLoaded = 88
  frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: stealth"
        lastLoadLine = 765
  Load frmStealth
  ToggleTopmost frmStealth.hwnd, True
  frmStealth.Hide
  dblAmLoaded = 89
  
  
  
  dblAmLoaded = 90
  'frmLoading.NotifyLoadProgress dblAmLoaded, "Loading modules: HP & mana"



  lastLoadLine = 800
  
  dblAmLoaded = 92
  frmLoading.NotifyLoadProgress dblAmLoaded, "Reading tibia.dat"
  If (TibiaVersionLong < 1100) Then
    If TibiaDatExists() = False Then
      MsgBox "tibia.dat missing/unreadable : " & vbCrLf & DBGtileError, vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
      End
    End If
  End If
  If ((TibiaVersionLong < 710) Or (TibiaVersionLong > highestTibiaVersionLong)) Then
    MsgBox "TibiaVersionLong is holding an unsupported version value (" & CStr(TibiaVersionLong) & ")" & vbCrLf & _
     "Solution (3 steps) :" & vbCrLf & _
     "1) close Blackd Proxy" & vbCrLf & _
     "2) delete settings.ini" & vbCrLf & _
     "3) reopen Blackd Proxy"
     End
  End If
'loadTibiaDatPath:
'  If configPath = "" Then
'    tibiadathere = App.path
'  Else
'    tibiadathere = App.path & "\" & configPath
'  End If
'  If ((TibiaVersionLong = highestTibiaVersionLong) And (UseRealTibiaDatInLatestTibiaVersion = True)) Then
'    If (TibiaExePath = "") Then
'        MsgBox "IMPORTANT WARNING:" & vbCrLf & _
'        "Please install real Tibia on default folder" & vbCrLf & _
'        "-or-" & vbCrLf & _
'        "define TibiaExePath in Blackd Proxy latest config.ini and reload it" & vbCrLf & _
'        "before trying to play in real servers." & vbCrLf & _
'        "Other way Blackd Proxy won't be able to detect" & vbCrLf & _
'        "ninja patchs of the tibia.dat file." & vbCrLf & _
'        "This would mean higher risk of autodetection for you", vbOKOnly + vbExclamation, "Warning"
'    Else
'        tibiadathere = TibiaExePath
'    End If
'  End If
 ' CurrentTibiaDatPath = tibiadathere
  


'  If ((Right$(CurrentTibiaDatPath, 1) = "\") Or (Right$(CurrentTibiaDatPath, 1) = "/")) Then
'     CurrentTibiaDatPath = CurrentTibiaDatPath & "Tibia.dat"
'  Else
'     CurrentTibiaDatPath = CurrentTibiaDatPath & "\Tibia.dat"
'  End If
  
  If (TibiaVersionLong >= 1100) Then
    
    GetInfoOfNewDatFile TibiaExePathWITHTIBIADAT, CurrentTibiaDatVERSION, CurrentTibiaDatFILE
    If (CurrentTibiaDatVERSION = -1) Then
       MsgBox "IMPORTANT WARNING:" & vbCrLf & _
       "Unable to read version of .dat file at" & vbCrLf & _
       TibiaExePathWITHTIBIADAT & vbCrLf & _
       "Blackd Proxy must close now." & vbCrLf & _
       vbCrLf & "DETAILS FOR DEBUG:" & vbCrLf & _
       dateErrDescription, vbOKOnly + vbCritical, "Critical error"
       SaveConfigWizard True
       End
    End If
  Else
    CurrentTibiaDatDATE = GetDATEOfFile(TibiaExePathWITHTIBIADAT)
    If (CurrentTibiaDatDATE = MyErrorDate) Then
       MsgBox "IMPORTANT WARNING:" & vbCrLf & _
       "Unable to read file:" & vbCrLf & _
       TibiaExePathWITHTIBIADAT & vbCrLf & _
       "Please ensure that you really installed Tibia there! Blackd Proxy must close now." & vbCrLf & _
       vbCrLf & "DETAILS FOR DEBUG:" & vbCrLf & _
       dateErrDescription, vbOKOnly + vbCritical, "Critical error"
       SaveConfigWizard True
       End
    End If
 End If
  
  res = UnifiedLoadDatFile(TibiaExePathWITHTIBIADAT)
  moreDetails = vbCrLf & "Trying to read Tibia " & TibiaVersion & " data here:" & vbCrLf & TibiaExePathWITHTIBIADAT & vbCrLf & vbCrLf & "Tibia client " & TibiaVersion & " is probably not installed there." & vbCrLf & "That folder probably contains a different Tibia version." & vbCrLf & "Update and run Blackd Proxy again and config everything correctly."
  If ((res = -1) Or (res = -2)) Then
    If MsgBox("Non compatible tibia.dat file , error " & CStr(res) & moreDetails & vbCrLf & _
    "You can resume loading Blackd Proxy, but it is strongly NOT RECOMMENDED. 99% of cheats will not work." & vbCrLf & _
    "- Do you want to resume loading anyways?", vbYesNo + vbExclamation, "Problem with config" & CStr(TibiaVersionLong)) = vbNo Then
        SaveConfigWizard True
        End
    End If
  End If
  If (res = -3) Then
    If MsgBox("Too many tiles found in tibia.dat , please increase MAXDATTILES in your config.ini" & CStr(res) & vbCrLf & _
    "You can resume loading Blackd Proxy, but it is strongly NOT RECOMMENDED. 99% of cheats will not work." & vbCrLf & _
    "- Do you want to resume loading anyways?", vbYesNo + vbExclamation, "Problem with config" & CStr(TibiaVersionLong)) = vbNo Then
        SaveConfigWizard True
        End
    End If
  End If
  If (res = -4) Then
    If MsgBox("Outstanding error -4 while reading tibia.dat: " & vbCrLf & DBGtileError & vbCrLf & _
    "You can resume loading Blackd Proxy, but it is strongly NOT RECOMMENDED. 99% of cheats will not work." & vbCrLf & _
    "- Do you want to resume loading anyways?", vbYesNo + vbExclamation, "Problem with config" & CStr(TibiaVersionLong)) = vbNo Then
        SaveConfigWizard True
        End
    End If
    
  End If
  If (res = -5) Then
    If MsgBox("Bug caught: " & vbCrLf & DBGtileError & vbCrLf & _
    "You can resume loading Blackd Proxy, but it is strongly NOT RECOMMENDED. 99% of cheats will not work." & vbCrLf & _
    "- Do you want to resume loading anyways?", vbYesNo + vbExclamation, "Problem with config" & CStr(TibiaVersionLong)) = vbNo Then
        SaveConfigWizard True
        End
    End If
  End If
  
  lastLoadLine = 850
  dblAmLoaded = 92
  frmLoading.NotifyLoadProgress dblAmLoaded, "Special gm names..."
  LoadSpecialGMnames
  dblAmLoaded = 93
  frmLoading.NotifyLoadProgress dblAmLoaded, "Reading user config"
  ' load ini
  ReadIni
  lastLoadLine = 860
  dblAmLoaded = 94
  frmLoading.NotifyLoadProgress dblAmLoaded, "Analyzing paths..."
retrypaths:
  givePathMsg frmLoading.hwnd
  lastLoadLine = 870
  str = TibiaPath
  If (Not (OVERWRITE_MAPS_PATH = "")) Then
  str = OVERWRITE_MAPS_PATH
  End If
  str2 = ValidateTibiaPath(str)
  lastLoadLine = 880
  If ((str2 = "") Or (str2 = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->")) Then
    If MsgBox("Sorry, it looks like the automap folder you selected is not valid " & vbCrLf & _
     "(" & TibiaPath & ")" & vbCrLf & vbCrLf & _
     "Picking the tibia automap folder is not mandatory. However it is recommended for optimal results at cavebot (rest of cheats will still work perfectly)" & vbCrLf & vbCrLf & _
     "Do you want to try again with other folder?" & vbCrLf & _
     "YES = I want best exp/h so let's try selecting a different folder." & vbCrLf & _
     "NO = I will fix this later.", vbYesNo + vbInformation, "Bad path") = vbYes Then
      TibiaPath = ""
      GoTo retrypaths
    Else
       If ReadHardiskMaps() = -1 Then
        ' ignore
       End If
       GoTo continuewithoutload
    End If
  End If
  
  dblAmLoaded = 96
  frmLoading.NotifyLoadProgress dblAmLoaded, "Reading maps..."
  
  lastLoadLine = 890
  ' read maps from harddisk
  lngTemp = ReadHardiskMaps()
  If lngTemp = -1 Then
     prevValue = TibiaPath
     TibiaPath = TryAutoPath()
     If TibiaPath <> "" Then
        TibiaPath = TryAutoPath()
        lngTemp = ReadHardiskMaps()
        If lngTemp = -1 Then
          TibiaPath = prevValue
        End If
     Else
        TibiaPath = prevValue
     End If
    If ((lngTemp = -1) And (TibiaVersionLong = highestTibiaVersionLong)) Then
        If MsgBox("We could not read maps in your automap folder" & vbCrLf & _
        "(" & TibiaPath & ")" & vbCrLf & _
        "This might be caused by corrupted maps or because you never executed Tibia before." & vbCrLf & vbCrLf & _
        "Cavebot pathing might get some problems with no maps." & vbCrLf & _
        "Do you want to continue anyways?", vbYesNo + vbQuestion, "Could not read maps") = vbNo Then
            End
            Exit Sub
        End If
    End If
  End If
continuewithoutload:
lastLoadLine = 895
  dblAmLoaded = 97
  frmLoading.NotifyLoadProgress dblAmLoaded, "Preloading character settings"
  PreloadAllCharSettingsFromHardDisk
  
  dblAmLoaded = 98
  
  
 ' If (TibiaVersionLong <= 970) Then
    MAXCHARACTERLEN = 30
 ' Else
 '   MAXCHARACTERLEN = 28
 ' End If
    frmLoading.NotifyLoadProgress dblAmLoaded, "Loading main menu"
  

       
       frmCavebot.TimerScript.enabled = True
       


  If LimitedLeader <> "-" Then
  frmMain.Caption = frmMain.Caption & "- LIMITED"
  frmMain.txtPackets.Text = "Special limited version for " & LimitedLeader & vbCrLf & vbCrLf & frmMain.txtPackets.Text
  End If
  
  dblAmLoaded = 100
  frmLoading.NotifyLoadProgress dblAmLoaded, "Done"

  Me.Hide
  txtTibiaPath.Text = TibiaPath
  Load frmMenu
  
  'custom ng
  If frmOld.chkOldMenu.value = 1 Then
    frmMenu.Hide
  End If
  
  
  For i = 1 To MAXSCHEDULED
    scheduledActions(i).pending = False
    scheduledActions(i).action = ""
    scheduledActions(i).clientID = 1
    scheduledActions(i).tickc = 0
  Next i

  frmEvents.timerScheduledActions.enabled = True
  frmTrainer.timerTrainer.enabled = True
  LoadWasCompleted = True
  Exit Sub
gotErr:
  'LogOnFile "debug.txt", "Terminated by critical error (-4)"
  Select Case lastLoadLine
  Case 749 ' first jwldbutton
    strHint = "Details: Unable to init buttons." & vbCrLf & _
     "This usually means that somehow JwldButn2b.ocx" & vbCrLf & _
     "was not correctly installed or registered." & vbCrLf & _
     "Did you really run the installer??." & vbCrLf & _
     "If installer doesn't register it then ..." & vbCrLf & _
     "Copy JwldButn2b.ocx and JwldButn2b.oca to your" & vbCrLf & _
     "dll folder, usually: windows\system32\" & vbCrLf & _
     "and register it using regsvr32" & vbCrLf & _
     "(please search a -how to register an ocx- in Google)"
  Case 601, 602, 603 ' first dictionary
    strHint = "Details: Unable to create the first dictionary object." & vbCrLf & _
     "This usually means that somehow scrrun.dll" & vbCrLf & _
     "was not correctly installed or registered." & vbCrLf & _
     "Did you really run the installer??." & vbCrLf & _
     "If installer doesn't register it then ..." & vbCrLf & _
     "Download scrrun.dll from Microsoft or google it," & vbCrLf & _
     "then save it on windows\system32\" & vbCrLf & _
     "and register it using regsvr32" & vbCrLf & _
     "(please search a -how to register a dll- in Google)"
  Case 651
    strHint = "Details: Outstanding debug in basic instructions, subpoint " & CStr(subdebug651)
  Case Else
    strHint = ""
  End Select
  
  MsgBox "Sorry, Blackd Proxy was not able to complete the loading." & vbCrLf & _
  " Debug mode activated. Please send the following details to daniel@blackdtools.com :" & vbCrLf & _
  " Details:" & vbCrLf & _
  " - Blackd Proxy Version: " & ProxyVersion & vbCrLf & _
  " - Tibia Version: " & CStr(TibiaVersionLong) & vbCrLf & _
  " - % sucesfully loaded: " & CStr(dblAmLoaded) & "%" & vbCrLf & _
  " - Last sucessfull debug waypoint: " & CStr(lastLoadLine) & vbCrLf & _
  " - Error number: " & Err.Number & vbCrLf & _
  " - Error description: " & Err.Description & vbCrLf & strHint, vbOKOnly + vbCritical, "Critical error"
  End
End Sub

Public Sub DoCloseActions(ByVal Index As Integer)
  ' Reset their vars to their initial states
 ' #If FinalMode Then
  On Error Resume Next
 ' #End If
  Dim j As Long
  Dim k As Long
  If Index > 0 Then
  If sckServerGame(Index).State <> sckClosed Then
    sckServerGame(Index).Close
  End If
  If sckClientGame(Index).State <> sckClosed Then
    sckClientGame(Index).Close
  End If
      AvoidReAttacks(Index) = True
  UHRetryCount(Index) = 0
  runemakerMana1(Index) = -1
    var_expleft(Index) = ""
    var_nextlevel(Index) = ""
    var_exph(Index) = ""
    var_timeleft(Index) = ""
    var_played(Index) = ""
    var_expgained(Index) = ""
    var_lf(Index) = vbLf
    var_lastsender(Index) = ""
    var_lastmsg(Index) = ""
  CavebotHaveSpecials(Index) = False
  CavebotLastSpecialMove(Index) = 0
  StatusBits(Index) = "0000000000000000"
  runeTurn(Index) = randomNumberBetween(0, 29)
  lastUsedChannelID(Index) = "05 00"
  lastRecChannelID(Index) = "05 00"
  reconnectionRetryCount(Index) = 0
  nextReconnectionRetry(Index) = 0
  SelfDefenseID(Index) = 0
  logoutAllowed(Index) = 0
  ReconnectionStage(Index) = 0
  IgnoreServer(Index) = False
  FirstCharInCharList(Index) = ""
  NoHealingNextTurn(Index) = False
  DropDelayerTurn(Index) = 0
  DelayAttacks(Index) = 0
  ReconnectionPacket(Index).numbytes = 0
  pauseStacking(Index) = 0
  nextAllowedmsg(Index) = 0
  AllowUHpaused(Index) = False
  doingTrade(Index) = False
  doingTrade2(Index) = False
  cavebotOnTrapGiveAlarm(Index) = False
  GotKillOrderTargetID(Index) = 0
  GotKillOrder(Index) = False
  GotKillOrderTargetName(Index) = ""
  lastAttackedIDstatus(Index) = 0
  previousAttackedID(Index) = 0
  initialRuneBackpack(Index) = &HFF
  DoingMainLoop(Index) = False
  RequiredMoveBuffer(Index) = ""
  ReadyBuffer(Index) = True
  frmMapReader.RemoveListItem CharacterName(Index)
  ReDim ConnectionBuffer(Index).packet(0)
  makingRune(Index) = False
  LoginMsgCount(Index) = 0
  lastHPchange(Index) = 0
  ConnectionBuffer(Index).numbytes = 0
  lastFloorTrap(Index) = -1
  givenUFO = False
  cancelAllMove(Index) = 0
  posSpamActivated(Index) = False
  posSpamChannelB1(Index) = &HFF
  posSpamChannelB2(Index) = &HFF
  
  getSpamActivated(Index) = False
  getSpamChannelB1(Index) = &HFF
  getSpamChannelB2(Index) = &HFF
  executingCavebot(Index) = False
  ResetEventList Index
  ResetCondEventList Index
  MustCheckFirstClientPacket(Index) = True
  
  If TibiaVersionLong >= 841 Then
     NeedToIgnoreFirstGamePacket(Index) = True
  Else
    NeedToIgnoreFirstGamePacket(Index) = False
  End If
  
  sentFirstPacket(Index) = True
  IDstring(Index) = ""
  myID(Index) = 0
  CharacterName(Index) = ""
  ConnectionBuffer(Index).numbytes = 0
  GameConnected(Index) = False
  onDepotPhase(Index) = 0
  CavebotChaoticMode(Index) = 0
  TurnsWithRedSquareZero(Index) = 0
  bLevelSpy(Index) = False
  depotX(Index) = 0
  nextForcedDepotDeployRetry(Index) = 0
  depotY(Index) = 0
  depotZ(Index) = 0
  doneDepotChestOpen(Index) = False
  depotTileID(Index) = 0
  depotS(Index) = 0
  lastDepotBPID(Index) = 0
  nextLight(Index) = "D7"
  NameOfID(Index).RemoveAll
  HPOfID(Index).RemoveAll
  DirectionOfID(Index).RemoveAll
  currTargetName(Index) = ""
  currTargetID(Index) = 0
  lootTimeExpire(Index) = 0
  CheatsPaused(Index) = False
  DangerGM(Index) = False
  DangerPK(Index) = False
  DangerPlayer(Index) = False
  LogoutTimeGM(Index) = 0
  GMname(Index) = ""
  cavebotOnDanger(Index) = -1
  cavebotOnGMclose(Index) = False
  cavebotOnGMpause(Index) = False
  lastAttackedID(Index) = 0
  CavebotTimeWithSameTarget(Index) = GetTickCount()
  CavebotTimeStart(Index) = GetTickCount()
  maxAttackTime(Index) = 40000
  ChaotizeNextMaxAttackTime Index
  maxHit(Index) = 10000
  previousAttackedID(Index) = 0
  DangerGMname(Index) = ""
  DangerPKname(Index) = ""
  DangerPlayerName(Index) = ""
  friendlyMode(Index) = 0
  cavebotLenght(Index) = 0
  cavebotEnabled(Index) = False
  EnableMaxAttackTime(Index) = False
  cavebotScript(Index).RemoveAll
  autoLoot(Index) = False
  myLastCorpseX(Index) = 0
  myLastCorpseY(Index) = 0
  myLastCorpseZ(Index) = 0
  myLastCorpseS(Index) = 0
  lastIngameCheck(Index) = ""
  lastIngameCheckTileID(Index) = "00 00"
  myLastCorpseTileID(Index) = 0
  lootWaiting(Index) = False
  requestLootBp(Index) = &HFF
  SendingSpecialOutfit(Index) = False
  moveRetry(Index) = 0
  lastX(Index) = 0
  lastY(Index) = 0
  lastZ(Index) = 0
  lastDestX(Index) = 0
  lastDestY(Index) = 0
  lastDestZ(Index) = 0
  receivedLogin(Index) = False
  setFollowTarget(Index) = True
  ignoreNext(Index) = 0
  GotPacketWarning(Index) = False
  LastHealTime(Index) = 0
  lastLootOrder(Index) = 0
  timeToRetryOpenDepot(Index) = 0
  ResetLooter Index
  OldLootMode(Index) = True
  ClientExecutingLongCommand(Index) = False
  LootAll(Index) = False
  PKwarnings(Index) = True
  LastCavebotTime(Index) = 0
  stealthLog(Index) = ""
  myHP(Index) = cte_initHP
  myMaxHP(Index) = cte_initHP
  myMaxMana(Index) = cte_initMANA
  lastHPchange(Index) = 0
  myNewStat(Index) = 0
  myMana(Index) = 0
  myCap(Index) = 0
  myStamina(Index) = 0
  somethingChangedInBps(Index) = False
  mySoulpoints(Index) = 100
  myExp(Index) = 0
  SpellKillHPlimit(Index) = 0
  SpellKillMaxHPlimit(Index) = 100
  AllowedLootDistance(Index) = 3
  myInitialExp(Index) = 0
  myInitialTickCount(Index) = 0
  myLevel(Index) = 50000000
  myMagLevel(Index) = 0
  For k = 1 To EQUIPMENT_SLOTS
    mySlot(Index, k).t1 = &H0
    mySlot(Index, k).t2 = &H0
    mySlot(Index, k).t3 = &H0
  Next k
  savedItem(Index).t1 = &H0
  savedItem(Index).t2 = &H0
  savedItem(Index).t2 = &H0
  pushDelay(Index) = CInt(Int((PUSHDELAYTIMES * Rnd)))
  'exeLine(Index) = 0
   updateExeLine Index, 0, False, False
  pushTarget(Index) = 0
  'ProcessID(index) = -1
  fishCounter(Index) = 0
  AfterLoginLogoutReason(Index) = ""
  RemoveAllMelee Index
  RemoveAllHMM Index
  
  RemoveAllSETUSEITEM Index

  
  RemoveAllAvoid Index
  RemoveAllShotType Index
  RemoveAllExorivis Index
  RemoveAllGoodLoot Index
  RemoveAllClientSpamOrders Index
  
  RuneMakerOptions(Index).activated = RuneMakerOptions_activated_default
  RuneMakerOptions(Index).autoEat = RuneMakerOptions_autoEat_default
  RuneMakerOptions(Index).ManaFluid = RuneMakerOptions_ManaFluid_default
  RuneMakerOptions(Index).autoLogoutAnyFloor = RuneMakerOptions_autoLogoutAnyFloor_default
  RuneMakerOptions(Index).autoLogoutCurrentFloor = RuneMakerOptions_autoLogoutCurrentFloor_default
  RuneMakerOptions(Index).autoLogoutOutOfRunes = RuneMakerOptions_autoLogoutOutOfRunes_default
  RuneMakerOptions(Index).autoWaste = RuneMakerOptions_autoWaste_default
  RuneMakerOptions(Index).msgSound = RuneMakerOptions_msgSound_default
  RuneMakerOptions(Index).msgSound2 = RuneMakerOptions_msgSound2_default
  RuneMakerOptions(Index).firstActionText = RuneMakerOptions_firstActionText_default
  RuneMakerOptions(Index).firstActionMana = RuneMakerOptions_firstActionMana_default
  RuneMakerOptions(Index).LowMana = RuneMakerOptions_LowMana_default
  RuneMakerOptions(Index).secondActionText = RuneMakerOptions_secondActionText_default
  RuneMakerOptions(Index).secondActionMana = RuneMakerOptions_secondActionMana_default
  RuneMakerOptions(Index).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default
  
  'custom ng var
  healingCheatsOptions(Index).sdmax = healingCheatsOptions_sdmax_default
  healingCheatsOptions(Index).antipush = healingCheatsOptions_antipush_default
  healingCheatsOptions(Index).pmax = healingCheatsOptions_pmax_default
  healingCheatsOptions(Index).htarget = healingCheatsOptions_htarget_default
  healingCheatsOptions(Index).exaustEat = healingCheatsOptions_exaustEat_default
  healingCheatsOptions(Index).HouseX = healingCheatsOptions_HouseX_default
  healingCheatsOptions(Index).HouseY = healingCheatsOptions_HouseY_default
  
  'custom ng healing
    healingCheatsOptions(Index).txtSpellhi = healingCheatsOptions_txtSpellhi_default
    healingCheatsOptions(Index).txtSpelllo = healingCheatsOptions_txtSpelllo_default
    healingCheatsOptions(Index).txtPot = healingCheatsOptions_txtPot_default
    healingCheatsOptions(Index).txtMana = healingCheatsOptions_txtMana_default
    healingCheatsOptions(Index).txtHealthhi = healingCheatsOptions_txtHealthhi_default
    healingCheatsOptions(Index).txtHealthlo = healingCheatsOptions_txtHealthlo_default
    healingCheatsOptions(Index).txtHealpot = healingCheatsOptions_txtHealpot_default
    healingCheatsOptions(Index).txtManapot = healingCheatsOptions_txtManapot_default
    healingCheatsOptions(Index).txtManahi = healingCheatsOptions_txtManahi_default
    healingCheatsOptions(Index).txtManalo = healingCheatsOptions_txtManalo_default
    healingCheatsOptions(Index).Combo1 = healingCheatsOptions_Combo1_default
    healingCheatsOptions(Index).Combo2 = healingCheatsOptions_Combo2_default
  
    'custom ng
    'extras
    extrasOptions(Index).txtSpell = extrasOptions_txtSpell_default
    extrasOptions(Index).txtMana = extrasOptions_txtMana_default
    extrasOptions(Index).txtSSA = extrasOptions_txtSSA_default
    extrasOptions(Index).cmbHouse = extrasOptions_cmbHouse_default
    extrasOptions(Index).chkMana = extrasOptions_chkMana_default
    extrasOptions(Index).chkDanger = extrasOptions_chkDanger_default
    extrasOptions(Index).chkPM = extrasOptions_chkPM_default
    extrasOptions(Index).chkEat = extrasOptions_chkEat_default
    extrasOptions(Index).chkautoUtamo = extrasOptions_chkautoUtamo_default
    extrasOptions(Index).chkautoHur = extrasOptions_chkautoHur_default
    extrasOptions(Index).chkautogHur = extrasOptions_chkautogHur_default
    extrasOptions(Index).chkAFK = extrasOptions_chkAFK_default
    extrasOptions(Index).chkGold = extrasOptions_chkGold_default
    extrasOptions(Index).chkPlat = extrasOptions_chkPlat_default
    extrasOptions(Index).chkDash = extrasOptions_chkDash_default
    extrasOptions(Index).chkMW = extrasOptions_chkMW_default
    extrasOptions(Index).chkSSA = extrasOptions_chkSSA_default
    extrasOptions(Index).chkTitle = extrasOptions_chkTitle_default
    extrasOptions(Index).chkHouse = extrasOptions_chkHouse_default

    'custom ng
    'persistent
    persistentOptions(Index).txtHk1 = persistentOptions_txtHk1_default
    persistentOptions(Index).txtHk2 = persistentOptions_txtHk2_default
    persistentOptions(Index).txtHk3 = persistentOptions_txtHk3_default
    persistentOptions(Index).txtHk4 = persistentOptions_txtHk4_default
    persistentOptions(Index).txtHk5 = persistentOptions_txtHk5_default
    persistentOptions(Index).txtHk6 = persistentOptions_txtHk6_default
    persistentOptions(Index).txtHk7 = persistentOptions_txtHk7_default
    persistentOptions(Index).txtHk8 = persistentOptions_txtHk8_default
    persistentOptions(Index).txtHk9 = persistentOptions_txtHk9_default
    persistentOptions(Index).txtHk10 = persistentOptions_txtHk10_default
    persistentOptions(Index).txtHk11 = persistentOptions_txtHk11_default
    persistentOptions(Index).txtExiva1 = persistentOptions_txtExiva1_default
    persistentOptions(Index).txtExiva2 = persistentOptions_txtExiva2_default
    persistentOptions(Index).txtExiva3 = persistentOptions_txtExiva3_default
    persistentOptions(Index).txtExiva4 = persistentOptions_txtExiva4_default
    persistentOptions(Index).txtExiva5 = persistentOptions_txtExiva5_default
    persistentOptions(Index).txtExiva6 = persistentOptions_txtExiva6_default
    persistentOptions(Index).txtExiva7 = persistentOptions_txtExiva7_default
    persistentOptions(Index).txtExiva8 = persistentOptions_txtExiva8_default
    persistentOptions(Index).txtExiva9 = persistentOptions_txtExiva9_default
    persistentOptions(Index).txtExiva10 = persistentOptions_txtExiva10_default
    persistentOptions(Index).txtExiva11 = persistentOptions_txtExiva11_default
    persistentOptions(Index).chkExiva1 = persistentOptions_chkExiva1_default
    persistentOptions(Index).chkExiva2 = persistentOptions_chkExiva2_default
    persistentOptions(Index).chkExiva3 = persistentOptions_chkExiva3_default
    persistentOptions(Index).chkExiva4 = persistentOptions_chkExiva4_default
    persistentOptions(Index).chkExiva5 = persistentOptions_chkExiva5_default
    persistentOptions(Index).chkExiva6 = persistentOptions_chkExiva6_default
    persistentOptions(Index).chkExiva7 = persistentOptions_chkExiva7_default
    persistentOptions(Index).chkExiva8 = persistentOptions_chkExiva8_default
    persistentOptions(Index).chkExiva9 = persistentOptions_chkExiva9_default
    persistentOptions(Index).chkExiva10 = persistentOptions_chkExiva10_default
    persistentOptions(Index).chkExiva11 = persistentOptions_chkExiva11_default
    
    'ng custom aimbot
    aimbotOptions(Index).chkSDcombo = aimbotOptions_chkSDcombo_default
    aimbotOptions(Index).chkUEcombo = aimbotOptions_chkUEcombo_default
    aimbotOptions(Index).txtLeader = aimbotOptions_txtLeader_default
    aimbotOptions(Index).txtCombo = aimbotOptions_txtCombo_default
    
    
  sentWelcome(Index) = False
  For j = 0 To HIGHEST_BP_ID
    Backpack(Index, j).open = False
    Backpack(Index, j).cap = 0
    Backpack(Index, j).used = 0
    Backpack(Index, j).name = ""
  Next j
  frmTrueMap.LoadChars
  frmRunemaker.LoadRuneChars
  frmHPmana.LoadHPmanaChars
  frmStealth.LoadStealthChars
  'custom ng
  frmHealing.LoadHealingChars
  frmExtras.LoadextrasChars
  frmPersistent.LoadPersistentChars
  frmAimbot.LoadAimbotChars
  
  
  frmEvents.LoadEventChars
  frmCondEvents.LoadCondEventChars
  ResetInternalTrainerValues Index
  frmTrainer.LoadTrainerChars
  frmCavebot.LoadCavebotChars
  frmBroadcast.LoadBroadcastChars
  End If
  If frmRunemaker.chkCloseSound.value = 1 Then
     frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "(Giving alarm because client " & CStr(Index) & " was closed)"
     ChangePlayTheDangerSound True
  End If
  Exit Sub
'gotErr:
'  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during DoCloseActions(" & index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub

Private Sub HideAdvancedOptions()
  ' Hide advanced options
  cmdAdvanced.Caption = "Show advanced options"
  chckMemoryIP.enabled = False
  lblListenLoginServer.enabled = False
  lblListenGameServer.enabled = False
  txtClientLoginP.enabled = False
  txtClientGameP.enabled = False
  lblWarning2.enabled = False
  lblAdvanced.enabled = False
  lblMaxTextChar.enabled = False
  lblMaxHexLines.enabled = False
  txtMaxChar.enabled = False
  txtMaxLines.enabled = False
  lblWhenAloggerIsFull.enabled = False
  LogFull1.enabled = False
  LogFull2.enabled = False
  LogFull3.enabled = False
  chkSelect.enabled = False
  chckAlter.enabled = False
  txtLogFile.enabled = False
  chkAutoHide.enabled = False
  cmbPrefered.enabled = False
  frmMain.Width = 6900
  frmMain.Height = 6550  ' Back to old default height
  'custom ng
 ' frmMain.Height = 2490
End Sub

Private Sub ShowAdvancedOptions()
  ' Show advanced options
  cmdAdvanced.Caption = "Hide advanced options"
  chckMemoryIP.enabled = True
  lblListenLoginServer.enabled = True
  cmbPrefered.enabled = True
  lblListenGameServer.enabled = True
  txtClientLoginP.enabled = True
  txtClientGameP.enabled = True
  lblWarning2.enabled = True
  lblAdvanced.enabled = True
  lblMaxTextChar.enabled = True
  lblMaxHexLines.enabled = True
  txtMaxChar.enabled = True
  txtMaxLines.enabled = True
  lblWhenAloggerIsFull.enabled = True
  LogFull1.enabled = True
  LogFull2.enabled = True
  LogFull3.enabled = True
  chkSelect.enabled = True
  chckAlter.enabled = True
  txtLogFile.enabled = True
  chkAutoHide.enabled = True
  frmMain.Width = 8490
  frmMain.Height = 8115
End Sub

Private Sub chkAutoHide_Click()
  ' change auto Hide option
  If chkAutoHide.value = 1 Then
    If chkLogPackets.value = 0 Then
      gridLog.Visible = False
      gridLog.enabled = False
      txtPackets.Height = 3495
    End If
  Else
    gridLog.Visible = True
    gridLog.enabled = True
    txtPackets.Height = 1215
  End If
End Sub

Private Sub chkLogPackets_Click()
  ' change log packets mode
  If chkLogPackets.value = 1 Then
    gridLog.Visible = True
    gridLog.enabled = True
    txtPackets.Height = 1215
  Else
    If chkAutoHide.value = 1 Then
      gridLog.Visible = False
      gridLog.enabled = False
      txtPackets.Height = 3495
    End If
  End If
End Sub

Private Sub cmbPrefered_Change()
  ' change prefered login server
  PREFEREDLOGINSERVER = cmbPrefered.Text
End Sub

Private Sub cmbPrefered_Click()
  Dim idLoginSP As Long
  ' change prefered login server by menu
  PREFEREDLOGINSERVER = cmbPrefered.Text
  For idLoginSP = 1 To NumberOfLoginServers
    If PREFEREDLOGINSERVER = trueLoginServer(idLoginSP) Then
        PREFEREDLOGINPORT = trueLoginPort(idLoginSP)
    End If
  Next idLoginSP
End Sub

Private Sub cmdAdvanced_Click()
  ' pressed Show advanced options / Hide advanced options
  If blnShowAdvancedOptions = False Then
    blnShowAdvancedOptions = True
    ShowAdvancedOptions
  Else
    blnShowAdvancedOptions = False
    HideAdvancedOptions
  End If
End Sub
Private Sub InitGridLog()
  ' init the grid log
  Dim i As Integer
  gridLog.Clear
  gridLog.Rows = 1
  For i = 0 To 20
    gridLog.ColWidth(i) = 300
  Next i
  ' gridLog head
  With gridLog
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, 1) = ""
    .TextMatrix(0, 2) = "("
    .TextMatrix(0, 3) = "H"
    .TextMatrix(0, 4) = "E"
    .TextMatrix(0, 5) = "X"
    .TextMatrix(0, 6) = ")"
    .TextMatrix(0, 7) = ""
    .TextMatrix(0, 8) = ""
    .TextMatrix(0, 9) = ""
    .TextMatrix(0, 10) = "#"
    .TextMatrix(0, 11) = ""
    .TextMatrix(0, 12) = ""
    .TextMatrix(0, 13) = "("
    .TextMatrix(0, 14) = "A"
    .TextMatrix(0, 15) = "S"
    .TextMatrix(0, 16) = "C"
    .TextMatrix(0, 17) = "I"
    .TextMatrix(0, 18) = "I"
    .TextMatrix(0, 19) = ")"
    .TextMatrix(0, 20) = ""
  End With
End Sub

Private Sub chckMemoryIP_Click()
  ' change memory IPS option
  If chckMemoryIP.value = 1 Then
   LastNumTibiaClients = 0 ' this will force change IPs now
  End If
End Sub





Private Sub cmdClear_Click()
  ' clear logs
  txtPackets.Text = ""
  InitGridLog
End Sub












Public Sub ReadTileIDFromIni(ByRef thing As Long, ByRef name As String, ByRef here As String, ByRef defaultV As String)
  ' read a tileID from ini
  Dim strInfo As String
  Dim lonInfo As Long
  Dim i As Integer
  strInfo = String$(50, 0)
  i = getBlackdINI("tileIDs", name, "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = GetTheLongFromFiveChr(strInfo)
    thing = lonInfo
  Else
    thing = GetTheLongFromFiveChr(defaultV)
  End If
End Sub

Public Sub FillTheListFromString(ByRef theList() As Long, ByRef theString As String)
  Dim remainingString As String
  Dim aTile As String
  Dim pos As Long
  Dim lonS As Long
  Dim listPos As Long
  Dim currChar As String
  On Error GoTo letsIgnoreIt
  lonS = Len(theString)
  pos = 1
  listPos = 0
  Do
    If pos > lonS Then
      theList(listPos) = 0
      Exit Do
    Else
      currChar = Mid$(theString, pos, 1)
      If (currChar = ",") Or (currChar = " ") Then
        pos = pos + 1
      Else
        If (pos + 5) <= (lonS + 1) Then
          aTile = Mid$(theString, pos, 5)
          theList(listPos) = GetTheLongFromFiveChr(aTile)
          listPos = listPos + 1
        End If
        pos = pos + 5
      End If
    End If
  Loop
  Exit Sub
letsIgnoreIt:
  theList(0) = 0
End Sub
Public Sub ReadTileIDListFromIni(ByRef thing() As Long, ByRef name As String, ByRef here As String, ByRef defaultV As String)
  ' read a tileID from ini
  Dim strInfo As String
  Dim i As Integer
  strInfo = String$(255, 0)
  i = getBlackdINI("tileIDs", name, "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    FillTheListFromString thing, strInfo
  Else
    FillTheListFromString thing, defaultV
  End If
End Sub

Public Sub WriteTileIDToIni(ByRef thing As Long, ByRef name As String, ByRef here As String)
  ' write a tileID to ini
  Dim strInfo As String
  Dim i As Integer
  strInfo = FiveChrLon(thing)
  i = setBlackdINI("tileIDs", name, strInfo, here)
End Sub

Public Sub WriteTileIDListToIni(ByRef thing() As Long, ByRef name As String, ByRef here As String)
  ' read a tileID from ini
  Dim strInfo As String
  Dim i As Integer
  strInfo = ""
  If thing(0) <> 0 Then
    strInfo = FiveChrLon(thing(0))
    For i = 1 To MAXTILEIDLISTSIZE
      If thing(i) <> 0 Then
        strInfo = strInfo & "," & FiveChrLon(thing(i))
      Else
        Exit For
      End If
    Next i
  End If
  i = setBlackdINI("tileIDs", name, strInfo, here)
End Sub


Public Function ReadIniThisFirst() As Long
  ' This function will read some important vars before the rest
  Dim i As Integer
  Dim strInfo As String
  Dim lonInfo As Long
  Dim here As String
  Dim tmpStr As String
  Dim res As Long
  Dim p1 As String
  Dim p2 As String
  Dim idLoginSP As Long
  Dim tmpNumber As Long
  Dim tmpVersion As String
  Dim debugPoint As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  res = -1
  debugPoint = 1
  userHere = App.Path
  debugPoint = 2
  If Right$(userHere, 1) = "\" Then
    userHere = userHere & "settings.ini"
  Else
    userHere = userHere & "\settings.ini"
  End If
  debugPoint = 3
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "configPath", "", strInfo, Len(strInfo), myMainConfigINIPath())
  If i > 0 Then
    strInfo = Left(strInfo, i)
    configPath = strInfo
  Else
    configPath = ""
  End If
  
  
  
  debugPoint = 4
  If (Not (OVERWRITE_CONFIGPATH = "")) Then
    configPath = OVERWRITE_CONFIGPATH
  Else
      ' new since Blackd Proxy 22.3
      configOverrideByCommand = False
      tmpStr = command$
      ' tmpStr = "-client_version=760"
      tmpNumber = InStr(1, tmpStr, ("-client_version=")) 'example: tibia.exe -client_version=760"
      If (tmpNumber > 0) Then
        configOverrideByCommand = True
        tmpVersion = Right$(tmpStr, Len(tmpStr) - 16)
        configPath = "config" & tmpVersion
      End If
    
'      If ((configOverrideByCommand = False) And (Not (configPath = ("config" & highestTibiaVersionLong)))) Then
'        If MsgBox("Do you want to load the config for Tibia " & TibiaVersionDefaultString & " instead?" & vbCrLf & _
'         "YES = Load config for latest Tibia (config" & highestTibiaVersionLong & ")" & vbCrLf & _
'         "NO = Keep loading current config (" & configPath & ")", vbQuestion + vbYesNo, "Warning: old config detected (" & configPath & ")") = vbYes Then
'         configPath = "config" & highestTibiaVersionLong
'        End If
'
'      End If
  End If
  If configPath = "" Then
    here = myMainConfigINIPath()
  Else
    here = App.Path & "\" & configPath & "\config.ini"
  End If
  Select Case configPath
  Case "config740"
    TibiaVersion = "7.72"
    TibiaVersionLong = 772
  Case "config760"
    TibiaVersion = "7.6"
    TibiaVersionLong = 760
  Case "config770"
    TibiaVersion = "7.7"
    TibiaVersionLong = 770
  Case "config772"
    TibiaVersion = "7.72"
    TibiaVersionLong = 772
  Case "config780"
    TibiaVersion = "7.8"
    TibiaVersionLong = 780
  Case "config781"
    TibiaVersion = "7.81"
    TibiaVersionLong = 781
  Case "configTEST"
    TibiaVersion = "8.0"
    TibiaVersionLong = 800
  Case "config790"
    TibiaVersion = "7.9"
    TibiaVersionLong = 790
  Case "config792"
    TibiaVersion = "7.92"
    TibiaVersionLong = 792
  Case "config800"
    TibiaVersion = "8.00"
    TibiaVersionLong = 800
  Case "config810"
    TibiaVersion = "8.1"
    TibiaVersionLong = 810
  Case "config811"
    TibiaVersion = "8.11"
    TibiaVersionLong = 811
  Case "config820"
    TibiaVersion = "8.2"
    TibiaVersionLong = 820
  Case "config821"
    TibiaVersion = "8.21"
    TibiaVersionLong = 821
  Case "config822"
    TibiaVersion = "8.22"
    TibiaVersionLong = 822
  Case "config830"
    TibiaVersion = "8.3"
    TibiaVersionLong = 830
  Case "config831"
    TibiaVersion = "8.31"
    TibiaVersionLong = 831
  Case "config840"
    TibiaVersion = "8.4"
    TibiaVersionLong = 840
  Case "config841"
    TibiaVersion = "8.41"
    TibiaVersionLong = 841
  Case "config842"
    TibiaVersion = "8.42"
    TibiaVersionLong = 842
  Case "config850"
    TibiaVersion = "8.5"
    TibiaVersionLong = 850
  Case "config852"
    TibiaVersion = "8.52"
    TibiaVersionLong = 852
  Case "config853"
    TibiaVersion = "8.53"
    TibiaVersionLong = 853
  Case "config854"
    TibiaVersion = "8.54"
    TibiaVersionLong = 854
  Case "config855"
    TibiaVersion = "8.55"
    TibiaVersionLong = 855
  Case "config856"
    TibiaVersion = "8.56"
    TibiaVersionLong = 856
  Case "config857"
    TibiaVersion = "8.57"
    TibiaVersionLong = 857
  Case "config860"
    TibiaVersion = "8.6"
    TibiaVersionLong = 860
  Case "config861"
    TibiaVersion = "8.61"
    TibiaVersionLong = 861
  Case "config862"
    TibiaVersion = "8.62"
    TibiaVersionLong = 862
  Case "config870"
    TibiaVersion = "8.70"
    TibiaVersionLong = 870
  Case "config871"
    TibiaVersion = "8.71"
    TibiaVersionLong = 871
  Case "config872"
    TibiaVersion = "8.72"
    TibiaVersionLong = 872
  Case "config873"
    TibiaVersion = "8.73"
    TibiaVersionLong = 873
  Case "config874"
    TibiaVersion = "8.74"
    TibiaVersionLong = 874
  Case "config900"
    TibiaVersion = "9.00"
    TibiaVersionLong = 900
  Case "config910"
    TibiaVersion = "9.1"
    TibiaVersionLong = 910
  Case "config920"
    TibiaVersion = "9.2"
    TibiaVersionLong = 920
  Case "config931"
    TibiaVersion = "9.31"
    TibiaVersionLong = 931
  Case "config940"
    TibiaVersion = "9.4"
    TibiaVersionLong = 940
  Case "config941"
    TibiaVersion = "9.41"
    TibiaVersionLong = 941
  Case "config942"
    TibiaVersion = "9.42"
    TibiaVersionLong = 942
  Case "config943"
    TibiaVersion = "9.43"
    TibiaVersionLong = 943
  Case "config944"
    TibiaVersion = "9.44"
    TibiaVersionLong = 944
  Case "config945"
    TibiaVersion = "9.45"
    TibiaVersionLong = 945
  Case "config946"
    TibiaVersion = "9.46"
    TibiaVersionLong = 946
  Case "config950"
    TibiaVersion = "9.5"
    TibiaVersionLong = 950
  Case "config951"
    TibiaVersion = "9.51"
    TibiaVersionLong = 951
  Case "config952"
    TibiaVersion = "9.52"
    TibiaVersionLong = 952
  Case "config953"
    TibiaVersion = "9.53"
    TibiaVersionLong = 953
  Case "config954"
    TibiaVersion = "9.54"
    TibiaVersionLong = 954
  Case "config960"
    TibiaVersion = "9.6"
    TibiaVersionLong = 960
  Case "config961"
    TibiaVersion = "9.61"
    TibiaVersionLong = 961
  Case "config962"
    TibiaVersion = "9.62"
    TibiaVersionLong = 962
  Case "config963"
    TibiaVersion = "9.63"
    TibiaVersionLong = 963
  Case "config970"
    TibiaVersion = "9.7"
    TibiaVersionLong = 970
  Case "config971"
    TibiaVersion = "9.71"
    TibiaVersionLong = 971
  Case "config980"
    TibiaVersion = "9.8"
    TibiaVersionLong = 980
  Case "config981"
    TibiaVersion = "9.81"
    TibiaVersionLong = 981
  Case "config982"
    TibiaVersion = "9.82"
    TibiaVersionLong = 982
  Case "config983"
    TibiaVersion = "9.83"
    TibiaVersionLong = 983
  Case "config984"
    TibiaVersion = "9.84"
    TibiaVersionLong = 984
  Case "config985"
    TibiaVersion = "9.85"
    TibiaVersionLong = 985
  Case "config986"
    TibiaVersion = "9.86"
    TibiaVersionLong = 986
  Case "config990"
    TibiaVersion = "9.9"
    TibiaVersionLong = 990
  Case "config991"
    TibiaVersion = "9.91"
    TibiaVersionLong = 991
  Case "config992"
    TibiaVersion = "9.92"
    TibiaVersionLong = 992
  Case "config1000"
    TibiaVersion = "10.0"
    TibiaVersionLong = 1000
  Case "config1001"
    TibiaVersion = "10.01"
    TibiaVersionLong = 1001
  Case "config1002"
    TibiaVersion = "10.02"
    TibiaVersionLong = 1002
  Case "config1010"
    TibiaVersion = "10.1"
    TibiaVersionLong = 1010
  Case "config1011"
    TibiaVersion = "10.11"
    TibiaVersionLong = 1011
  Case "config1012"
    TibiaVersion = "10.12"
    TibiaVersionLong = 1012
  Case "config1020"
    TibiaVersion = "10.2"
    TibiaVersionLong = 1020
  Case "config1021"
    TibiaVersion = "10.21"
    TibiaVersionLong = 1021
  Case "config1021preview"
    TibiaVersion = "10.21"
    TibiaVersionLong = 1021
  Case "config1022"
    TibiaVersion = "10.22"
    TibiaVersionLong = 1022
  Case "config1030"
    TibiaVersion = "10.3"
    TibiaVersionLong = 1030
  Case "config1031"
    TibiaVersion = "10.31"
    TibiaVersionLong = 1031
  Case "config1032"
    TibiaVersion = "10.32"
    TibiaVersionLong = 1032
  Case Else
    SetFutureTibiaVersion (configPath)
    If Tibia11allowed = False Then
        If (TibiaVersionLong >= 1100) Then
            End
        End If
    End If
  End Select
  
  If TibiaVersionLong < 820 Then
    oldmessage_H0 = &H0
    oldmessage_H1 = &H1
    oldmessage_H2 = &H2
    oldmessage_H3 = &H3
    oldmessage_H4 = &H4
    oldmessage_H5 = &H5
    oldmessage_H6 = &H6
    oldmessage_H7 = &H7
    oldmessage_H8 = &H8
    oldmessage_H9 = &H9
    oldmessage_HA = &HA
    oldmessage_HB = &HB
    oldmessage_HC = &HC
    oldmessage_HD = &HD
    oldmessage_HE = &HE
    oldmessage_HF = &HF
    oldmessage_H10 = &H10
    oldmessage_H11 = &H11
    oldmessage_H12 = &H12
    oldmessage_H13 = &H13
    oldmessage_H14 = &H14
    oldmessage_H15 = &H15
    newmessage_H8 = &HFF
    
    newchatmessage_H9 = &HFF
    newchatmessage_HA = &HFF
    verynewchatmessage_HB = &HFF
  ElseIf TibiaVersionLong < 840 Then
    oldmessage_H0 = &H2
    oldmessage_H1 = &H3
    oldmessage_H2 = &H4
    oldmessage_H3 = &H5
    oldmessage_H4 = &H6
    oldmessage_H5 = &H7
    oldmessage_H6 = &H8
    oldmessage_H7 = &H9
    oldmessage_H8 = &HA
    oldmessage_H9 = &HB
    oldmessage_HA = &HC
    oldmessage_HB = &HD
    oldmessage_HC = &HE
    oldmessage_HD = &HF
    oldmessage_HE = &H10
    oldmessage_HF = &H11
    oldmessage_H10 = &H12
    oldmessage_H11 = &H13
    oldmessage_H12 = &H14
    oldmessage_H13 = &H15
    oldmessage_H14 = &H0
    oldmessage_H15 = &H1
    newmessage_H8 = &HFF

    newchatmessage_H9 = &HFF
    newchatmessage_HA = &HFF
    verynewchatmessage_HB = &HFF
  ElseIf TibiaVersionLong <= 860 Then
    oldmessage_H0 = &H2
    oldmessage_H1 = &H3
    oldmessage_H2 = &H4
    oldmessage_H3 = &H5
    oldmessage_H4 = &H6
    oldmessage_H5 = &H7 ' channel OK
    
    newmessage_H8 = &H8
    
    oldmessage_H6 = &H9 ' ?
    oldmessage_H7 = &HA ' ?
    oldmessage_H8 = &HB ' ?
    
    oldmessage_H9 = &HC ' gm OK
    oldmessage_HA = &HD
    oldmessage_HB = &HE
    oldmessage_HC = &HF
    oldmessage_HD = &H10
    oldmessage_HE = &H11
    oldmessage_HF = &H12
    oldmessage_H10 = &H13
    oldmessage_H11 = &H14
    oldmessage_H12 = &H15
    oldmessage_H13 = &H16
    oldmessage_H14 = &H0
    oldmessage_H15 = &H1
    
    newchatmessage_H9 = &HFF
    newchatmessage_HA = &HFF
    verynewchatmessage_HB = &HFF
  ElseIf TibiaVersionLong <= 871 Then
    oldmessage_H0 = &H2
    oldmessage_H1 = &H3
    oldmessage_H2 = &H4
    oldmessage_H3 = &H5
    oldmessage_H4 = &H6
    oldmessage_H5 = &H7
    oldmessage_H6 = &HFF ' deleted?
    oldmessage_H7 = &H9 ' gm broadcast OK?
    newmessage_H8 = &H8 ' PARTY LOOT OK
    oldmessage_H8 = &H8 ' PARTY LOOT OK
    oldmessage_H9 = &HB ' ok, gm -1
    oldmessage_HA = &HA ' ok, gm tals to channel
    oldmessage_HB = &HFF ' UNSURE ' was duplicated!
    oldmessage_HC = &HC  ' UNSURE
    oldmessage_HD = &HFF ' deleted
    oldmessage_HE = &HFF ' deleted
    oldmessage_HF = &HFF ' deleted
    oldmessage_H10 = &HD   '  old 13-new 0D = -6 ok , monster talk (ex: cat meow)
    oldmessage_H11 = &HE ' logical move -6
    oldmessage_H12 = &HF ' logical move -6
    oldmessage_H13 = &HFF ' deleted
    oldmessage_H14 = &H0 ' ?
    oldmessage_H15 = &H1 ' ?

    newchatmessage_H9 = &HFF
    newchatmessage_HA = &HFF
    verynewchatmessage_HB = &HFF
  ElseIf TibiaVersionLong < 1036 Then
    oldmessage_H0 = &H0 ' -2 ok
    oldmessage_H1 = &H1 ' -2 ok
    oldmessage_H2 = &H2 ' -2 ok
    oldmessage_H3 = &H3 ' -2 ok
    oldmessage_H4 = &H4 ' -2 ok
    oldmessage_H5 = &H7 ' no change ?

    oldmessage_H7 = &H1D ' + &H14 ?
    newmessage_H8 = &H8 ' no change ?

    oldmessage_H9 = &HC ' +1 OK
    oldmessage_HA = &HFF ' unknown equivalent
    newchatmessage_HA = &HA
    newchatmessage_H9 = &H9
    
    oldmessage_HC = &HD ' logical move +1

    oldmessage_H10 = &H22 ' + &H14
    oldmessage_H11 = &H23 ' + &H14 OK
    oldmessage_H12 = &H24 ' + &H14

    oldmessage_H14 = &HFF ' no change OK
    oldmessage_H15 = &HFF ' no change OK
    

    
    oldmessage_H6 = &HFF
    oldmessage_H13 = &HFF
    oldmessage_HD = &HFF
    oldmessage_HE = &HFF
    oldmessage_HF = &HFF
    oldmessage_HB = &HFF ' was duplicated!
    oldmessage_H8 = &HFF ' not really used in the code!
    verynewchatmessage_HB = &HFF
  Else '10.36+
    oldmessage_H0 = &H0
    oldmessage_H1 = &H1
    oldmessage_H2 = &H2
    oldmessage_H3 = &H3
    oldmessage_H4 = &H4
    oldmessage_H5 = &H7
    newmessage_H8 = &H8
    newchatmessage_H9 = &H9
    newchatmessage_HA = &HA
    verynewchatmessage_HB = &HB
    oldmessage_H9 = &HD ' +1 OK
    oldmessage_HC = &HE ' +1
    oldmessage_H7 = &H1E ' +1
    oldmessage_H10 = &H23 ' +1
    oldmessage_H11 = &H24 ' +1
    oldmessage_H12 = &H25 ' +1
    oldmessage_H14 = &HFF
    oldmessage_H15 = &HFF
    oldmessage_H6 = &HFF
    oldmessage_H13 = &HFF
    oldmessage_HD = &HFF
    oldmessage_HE = &HFF
    oldmessage_HF = &HFF
    oldmessage_HB = &HFF
    oldmessage_H8 = &HFF
    
    oldmessage_HA = &HFF
    
  End If
  ' Note that all the default values are for Tibia 7.5
  
  'strInfo = String$(50, 0)
  'i = getBlackdINI("Proxy", "TibiaVersion", "", strInfo, Len(strInfo), here)
  'If i > 0 Then
  '  strInfo = Left(strInfo, i)
  '  TibiaVersion = strInfo
  '  If Len(TibiaVersion) = 3 Then
  '    TibiaVersionLong = (CLng(Left$(TibiaVersion, 1)) * 100) + (CLng(Right$(TibiaVersion, 1)) * 10)
  '  ElseIf Len(TibiaVersion) = 4 Then
  '    p1 = Left$(TibiaVersion, 1)
  '    p2 = Right$(TibiaVersion, 2)
  '    TibiaVersionLong = (CLng(p1) * 100) + (CLng(p2))
  '  Else
  '    LogOnFile "errors.txt", "Invalid TibiaVersion length (got " & CStr(Len(TibiaVersion)) & ")"
  '  End If
  'Else
  '  TibiaVersion = "7.5"
 '   TibiaVersionLong = 750
 ' End If
  
  'If TibiaVersionLong <= 760 Then
  '   LoginMethod = 0
  'Else
  '   LoginMethod = 1
  'End If

  debugPoint = 10
  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "MAXCLIENTS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MAXCLIENTS = lonInfo
  Else
    MAXCLIENTS = 5
  End If
  
  
  ' DefaultTibiaFolder
  strInfo = String$(250, 0)
  i = getBlackdINI("MemoryAddresses", "DefaultTibiaFolder", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     DefaultTibiaFolder = strInfo
  Else
     DefaultTibiaFolder = "Tibia"
  End If
  
  debugPoint = 11
If (OVERWRITE_CLIENT_PATH = "") Then
  strInfo = String$(250, 0)
  i = getBlackdINI("MemoryAddresses", "TibiaExePath", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     TibiaExePath = strInfo
  Else
     TibiaExePath = autoGetTibiaFolder()
  End If
Else
 TibiaExePath = OVERWRITE_CLIENT_PATH
End If
TibiaExePathWITHTIBIADAT = GetWITHTIBIADAT()
  
'  strInfo = String$(10, 0)
'  i = getBlackdINI("Proxy", "UseRealTibiaDatInLatestTibiaVersion", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'    lonInfo = CLng(strInfo)
'    If lonInfo = 1 Then
'       UseRealTibiaDatInLatestTibiaVersion = True
'    Else
'       UseRealTibiaDatInLatestTibiaVersion = False
'    End If
'  Else
'    UseRealTibiaDatInLatestTibiaVersion = True
'  End If
  

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LAST_BATTLELISTPOS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LAST_BATTLELISTPOS = lonInfo
  Else
    If TibiaVersionLong >= 873 Then
      LAST_BATTLELISTPOS = 1299
    Else
      LAST_BATTLELISTPOS = 147
    End If
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "useDynamicOffset", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    useDynamicOffset = strInfo
  Else
    If TibiaVersionLong >= 910 Then
      useDynamicOffset = "yes"
    Else
      useDynamicOffset = "no"
    End If
  End If
  If useDynamicOffset = "yes" Then
  useDynamicOffsetBool = True
  Else
  useDynamicOffsetBool = False
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "CloseLoginServerAfterCharList", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    If lonInfo = 1 Then
      CloseLoginServerAfterCharList = True
    Else
      CloseLoginServerAfterCharList = False
    End If
  Else
    CloseLoginServerAfterCharList = False
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "MemoryProtectedMode", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    If lonInfo = 0 Then
      MemoryProtectedMode = False
    Else
      MemoryProtectedMode = True
      SetAllPrivilegesForMe
    End If
  Else
    MemoryProtectedMode = True
    SetAllPrivilegesForMe
  End If
  

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "ForceDisableEncryption", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    If lonInfo = 1 Then
      ForceDisableEncryption = True
    Else
      ForceDisableEncryption = False
    End If
  Else
    ForceDisableEncryption = False
  End If

  
    strInfo = String$(50, 0)
    i = getBlackdINI("MemoryAddresses", "tibiaModuleRegionSize", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      tibiaModuleRegionSize = lonInfo
    Else
      tibiaModuleRegionSize = &H2C3000
    End If
    
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "MAXEVENTS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MAXEVENTS = lonInfo
  Else
    MAXEVENTS = 100
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "MAXCONDS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MAXCONDS = lonInfo
  Else
    MAXCONDS = 100
  End If
  
  
  
  strInfo = String$(50, 0)
  i = getBlackdINI("MemoryAddresses", "HIGHEST_BP_ID", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    HIGHEST_BP_ID = lonInfo
  Else
    HIGHEST_BP_ID = 15
  End If
  strInfo = String$(50, 0)
  i = getBlackdINI("MemoryAddresses", "MAXDATTILES", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MAXDATTILES = lonInfo
  Else
    MAXDATTILES = 10000
  End If

  
  strInfo = String$(50, 0)
  i = getBlackdINI("Proxy", "MAXTILEIDLISTSIZE", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MAXTILEIDLISTSIZE = lonInfo
  Else
    MAXTILEIDLISTSIZE = 50
  End If
  ReDim AditionalStairsToDownFloor(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalStairsToUpFloor(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalRequireRope(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalRequireShovel(0 To MAXTILEIDLISTSIZE)
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Proxy", "FirstExecute", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    If strInfo = "FALSE" Then
      FirstExecute = False
    Else
      FirstExecute = True
    End If
  Else
    FirstExecute = True
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "NextScreenshotNumber", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    lngNextScreenshotNumber = lonInfo
  Else
    lngNextScreenshotNumber = 1
  End If
  

  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "AlternativeBinding", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    AlternativeBinding = lonInfo
  Else
    AlternativeBinding = 0
  End If
  frmAdvanced.chkAlternativeBinding.value = AlternativeBinding
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "MyPriorityID", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    MyPriorityID = lonInfo
  Else
    MyPriorityID = 4
  End If
  frmAdvanced.LoadMyPriorityValue

  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "TibiaPriorityID", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TibiaPriorityID = lonInfo
  Else
    TibiaPriorityID = 2
  End If
  frmAdvanced.LoadTibiaPriorityValue
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "TOOSLOWLOGINSERVER_MS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TOOSLOWLOGINSERVER_MS = lonInfo
  Else
    TOOSLOWLOGINSERVER_MS = 500
  End If
  
  ' tibiaclassname
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "tibiaclassname", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    tibiaclassname = strInfo
  Else
    tibiaclassname = "tibiaclient"
  End If
  
    ' tibiaclassname
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "tibiamainname", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    tibiamainname = strInfo
  Else
    tibiamainname = "Tibia"
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "NumberOfLoginServers", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    NumberOfLoginServers = lonInfo
  Else
    If TibiaVersionLong >= 800 Then
        NumberOfLoginServers = 10
    Else
        NumberOfLoginServers = 5
    End If
  End If
  
  ReDim trueLoginServer(1 To NumberOfLoginServers)
  ReDim trueLoginPort(1 To NumberOfLoginServers)
  ReDim memLoginServer(1 To NumberOfLoginServers)
  ReDim MemPortLoginServer(1 To NumberOfLoginServers)
  
  If TibiaVersionLong < 1011 Then
  For idLoginSP = 1 To NumberOfLoginServers
    strInfo = String$(50, 0)
    i = getBlackdINI("MemoryAddresses", "MemLoginServer" & CStr(idLoginSP), "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      memLoginServer(idLoginSP) = lonInfo
    Else
      memLoginServer(idLoginSP) = &H5EB998
    End If
    
    strInfo = String$(50, 0)
    i = getBlackdINI("MemoryAddresses", "MemPortLoginServer" & CStr(idLoginSP), "", strInfo, Len(strInfo), here)
    If i > 0 Then
        strInfo = Left(strInfo, i)
        lonInfo = CLng(strInfo)
        MemPortLoginServer(idLoginSP) = lonInfo
    Else
        MemPortLoginServer(idLoginSP) = &H5EB9FC
    End If
    Next idLoginSP
  Else ' new system
    
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "LoginServerStartPointer", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      LoginServerStartPointer = lonInfo
    Else
      LoginServerStartPointer = &H80F31C
    End If
    
    
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "LoginServerStep", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      LoginServerStep = lonInfo
    Else
      LoginServerStep = &H32
    End If
    
    
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "HostnamePointerOffset", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      HostnamePointerOffset = lonInfo
    Else
      HostnamePointerOffset = &H4
    End If
    
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "IPAddressPointerOffset", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      IPAddressPointerOffset = lonInfo
    Else
      IPAddressPointerOffset = &H20
    End If
     
   
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "PortOffset", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      PortOffset = lonInfo
    Else
      PortOffset = &H30
    End If
    
    
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "proxyChecker", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      proxyChecker = lonInfo
    Else
      proxyChecker = &H0
    End If
    
    
    If TibiaVersionLong = 1011 Then
    strInfo = String$(250, 0)
    i = getBlackdINI("MemoryAddresses", "tibiaEntryServer", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      tibiaEntryServer = Trim$(strInfo)
    Else
      tibiaEntryServer = "game.tibia.ciproxy.com"
    End If
    End If
    
  End If
  
  
 






  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LEVELSPY_NOP", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LEVELSPY_NOP = lonInfo
  Else
    LEVELSPY_NOP = &H4D1680
  End If
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LEVELSPY_ABOVE", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LEVELSPY_ABOVE = lonInfo
  Else
    LEVELSPY_ABOVE = &H4D167C
  End If
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LEVELSPY_BELOW", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LEVELSPY_BELOW = lonInfo
  Else
    LEVELSPY_BELOW = &H4D1684
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LIGHT_NOP", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LIGHT_NOP = lonInfo
  Else
    LIGHT_NOP = &H4E51B9
  End If
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "LIGHT_AMOUNT", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LIGHT_AMOUNT = lonInfo
  Else
    LIGHT_AMOUNT = &H4E51BC
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "PLAYER_Z", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    PLAYER_Z = lonInfo
  Else
    PLAYER_Z = &H63BAD8
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "RedSquare", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    RedSquare = lonInfo
  Else
    RedSquare = 0 ' undefined
  End If













 
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrMulticlient", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrMulticlient = lonInfo
  Else
    adrMulticlient = &H502BB5
  End If
  strInfo = String$(10, 0)
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrRSA", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrRSA = lonInfo
  Else
    adrRSA = &H0
  End If
  strInfo = String$(10, 0)
  
  i = getBlackdINI("MemoryAddresses", "multiclientByte1", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    multiclientByte1 = CByte(lonInfo)
  Else
    multiclientByte1 = &H90
  End If
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "multiclientByte2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    multiclientByte2 = CByte(lonInfo)
  Else
    multiclientByte2 = &H90
  End If

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrXgo", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrXgo = lonInfo
  Else
    adrXgo = &H49D070
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrYgo", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrYgo = lonInfo
  Else
    adrYgo = &H49D06C
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrZgo", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrZgo = lonInfo
  Else
    adrZgo = &H49D068
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrGo", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrGo = lonInfo
  Else
    adrGo = &H49D0DC
  End If
  

  
'  If TibiaVersionLong >= 1100 Then
'    moveByMemoryEnabled = False ' temporally disabled
'  Else
'    moveByMemoryEnabled = True
'  End If
  moveByMemoryEnabled = True

  ' Tibia ANY version
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrConnectionKey", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrConnectionKey = ReadAddressPath(strInfo)
  Else
    If TibiaVersionLong >= 1100 Then
        adrConnectionKey = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 1DC > 308 > 90 > C8") ' 11.01
    Else
        adrConnectionKey = ReadAddressPath("&H6FA1A0") ' oldest value
    End If
  End If

  ' Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrNewRedSquare", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrNewRedSquare = ReadAddressPath(strInfo)
  Else
    adrNewRedSquare = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 1D8 > E4 > 1C")  ' 11.0
  End If
  
  
  ' FOR MINIMAP CLICK (1) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrGameRect_Width_Double", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrGameRect_Width_Double = ReadAddressPath(strInfo)
  Else
    adrGameRect_Width_Double = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 10C > 18 > 4 > 74 > 4 > 88")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (2) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapRect_Y_Double", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapRect_Y_Double = ReadAddressPath(strInfo)
  Else
    adrMiniMapRect_Y_Double = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 8 > 4 > 90")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (3) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapRect_Width_Double", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapRect_Width_Double = ReadAddressPath(strInfo)
  Else
    adrMiniMapRect_Width_Double = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 8 > 134 > 24 > 8 > 4 > 98")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (4) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapRect_Height_Double", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapRect_Height_Double = ReadAddressPath(strInfo)
  Else
    adrMiniMapRect_Height_Double = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 8 > 4 > A0")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (5) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_MinX", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_MinX = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_MinX = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 20")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (6) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_MinY", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_MinY = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_MinY = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 24")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (7) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_Z", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_Z = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_Z = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 28")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (8) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_SizeX", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_SizeX = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_SizeX = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 30")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (9) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_SizeY", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_SizeY = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_SizeY = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 34")  ' 11.0
  End If
  ' FOR MINIMAP CLICK (10) - Tibia 11 +
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrMiniMapDisplay_Zoom_PointSize1_Float", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrMiniMapDisplay_Zoom_PointSize1_Float = ReadAddressPath(strInfo)
  Else
    adrMiniMapDisplay_Zoom_PointSize1_Float = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 134 > 24 > 2c > 48")  ' 11.0
  End If

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "offSetSquare_ARGB_8bytes", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    offSetSquare_ARGB_8bytes = lonInfo
  Else
    offSetSquare_ARGB_8bytes = &H84 ' Tibia 11.00
  End If
  
'  ' Tibia 11 +
'  strInfo = String$(255, 0)
'  i = getBlackdINI("MemoryAddresses", "adrNewBlueSquare", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'    adrNewBlueSquare = ReadAddressPath(strInfo)
'  Else
'    adrNewBlueSquare = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 1D8 > E4 > 14")  ' 11.0
'  End If

  ' Tibia ANY version
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedCharIndex", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedCharIndex = ReadAddressPath(strInfo)
  Else
    'adrSelectedCharIndex = &H6FC9D8 '7.63
    adrSelectedCharIndex = ReadAddressPath("&H5F6CB0") '7.6
  End If
  
  
  ' Tibia 11
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedCharIndex", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedItem_height = ReadAddressPath(strInfo)
  Else
    adrSelectedItem_height = ReadAddressPath("""Qt5Widgets.dll"" + 00401DC4 > 70 > 4 > a8") ' 11.0
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedCharName", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedCharName = ReadAddressPath(strInfo)
  Else
    adrSelectedCharName = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > C > 4 > 18 > 38 > 28") ' 11.0
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrServerList_CollectionStart", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrServerList_CollectionStart = ReadAddressPath(strInfo)
  Else
    adrServerList_CollectionStart = ReadAddressPath("""Qt5Core.dll"" + 4555C8 > 8 > 168 > 54 > 18 > 2c") ' 11.0
  End If
 
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", " adrBattlelist_CollectionStart", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     adrBattlelist_CollectionStart = ReadAddressPath(strInfo)
  Else
     adrBattlelist_CollectionStart = ReadAddressPath("""Qt5Core.dll"" + 4555C8 > 8 > 1D8 > E4 > 8") ' 11.0
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedCharName_afterCharList", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedCharName_afterCharList = ReadAddressPath(strInfo)
  Else
    adrSelectedCharName_afterCharList = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 320 > 18 > 60 > 0") ' 11.0
  End If
   strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedServerURL_afterCharList", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedServerURL_afterCharList = ReadAddressPath(strInfo)
  Else
    adrSelectedServerURL_afterCharList = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 320 > 18 > 60 > 4") ' 11.0
  End If
    strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedServerPORT_afterCharList", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedServerPORT_afterCharList = ReadAddressPath(strInfo)
  Else
    adrSelectedServerPORT_afterCharList = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 320 > 18 > 60 > 8") ' 11.0
  End If
    strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "adrSelectedServerNAME_afterCharList", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    adrSelectedServerNAME_afterCharList = ReadAddressPath(strInfo)
  Else
    adrSelectedServerNAME_afterCharList = ReadAddressPath("""Qt5Core.dll"" + 004555C8 > 8 > 320 > 18 > 60 > C") ' 11.0
  End If
  

  
  
  
  
  
  
  
  
  
  
  
  
  
'  strInfo = String$(10, 0)
'  i = getBlackdINI("MemoryAddresses", "adrAccount", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'    lonInfo = CLng(strInfo)
'    adrAccount = lonInfo
'  Else
'    adrAccount = &H7893D4 '8.41
'  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrLastPacket", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrLastPacket = lonInfo
  Else
    'adrLastPacket = &H6F78BA '7.64
    adrLastPacket = &H5F3D98 '7.6
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrCharListPtr", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrCharListPtr = lonInfo
  Else
    'adrCharListPtr = &H6FA92C '7.64
    adrCharListPtr = &H5F6CB4  '7.6
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrCharListPtrEND", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrCharListPtrEND = lonInfo
  Else
    adrCharListPtrEND = &H951A4C  ' 10.11
  End If
  

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrNChar", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrNChar = lonInfo
  Else
    adrNChar = &H49D090
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "OutfitDist", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    OutfitDist = lonInfo
  Else
    OutfitDist = &H60
  End If
  
  If TibiaVersionLong >= 944 Then
    adrOutfit = adrNChar + OutfitDist
  Else
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "adrOutfit", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      adrOutfit = lonInfo
    Else
      adrOutfit = &H49D0F0
    End If
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "CharDist", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    CharDist = lonInfo
  Else
    CharDist = &H9C
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "NameDist", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    NameDist = lonInfo
  Else
    NameDist = &H4
  End If
  

  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "SpeedDist", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    SpeedDist = lonInfo
  Else
    SpeedDist = &H88
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrNum", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrNum = lonInfo
  Else
    adrNum = &H49D02C
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrConnected", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrConnected = lonInfo
  Else
    adrConnected = &H5F0380
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "maxStepsPerMovement", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    maxStepsPerMovement = lonInfo
  Else
    maxStepsPerMovement = 5
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrPointerToInternalFPSminusH5D", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrPointerToInternalFPSminusH5D = lonInfo
  Else
    adrPointerToInternalFPSminusH5D = &H7526BC 'default 7.81
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrNumberOfAttackClick", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrNumberOfAttackClicks = lonInfo
  Else
    adrNumberOfAttackClicks = &H0
  End If
  
  If adrNumberOfAttackClicks = &H0 Then
    ' try to read old var ( adrNumberOfAttackClicks )
    strInfo = String$(10, 0)
    i = getBlackdINI("MemoryAddresses", "adrNumberOfAttackClicks", "", strInfo, Len(strInfo), here)
    If i > 0 Then
      strInfo = Left(strInfo, i)
      lonInfo = CLng(strInfo)
      adrNumberOfAttackClicks = lonInfo
    Else
      adrNumberOfAttackClicks = &H0
    End If
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "adrInternalFPS", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    adrInternalFPS = lonInfo
  Else
    adrInternalFPS = &H5F6DF5 'default 7.6 (only usefull for 7.6)
  End If

  ' Meaning of memory addresses used:
  

  
  'adrXgo = &H49D070 ' goto this x
  'adrYgo = &H49D06C ' goto this y
  'adrZgo = &H49D068 ' goto this z
  'adrGo = &H49D0DC ' start goto process of first battlelist item
  
  'adrOutfit = &H49D0F0 ' first outfit byte of first battlelist item
  
  'adrNChar = &H49D090 ' ID at first battlelist item
  
  'CharDist = &H9C ' distance between 2 battlelist items
  'adrNum = &H49D02C  ' this always will contain your ID
  
  'adrConnected = &H5F0380 ' 0 if not connected / else it is connected
  
  
  ' Resize arrays to hold MAXCLIENTS clients :
    ResetOffsetCache MAXCLIENTS
    ReDim timeToRetryOpenDepot(1 To MAXCLIENTS)
    ReDim Looter(1 To MAXCLIENTS)
    ReDim SpellKillHPlimit(1 To MAXCLIENTS)
    ReDim SpellKillMaxHPlimit(1 To MAXCLIENTS)
    ReDim ClientExecutingLongCommand(1 To MAXCLIENTS)
    ReDim AllowRepositionAtStart(1 To MAXCLIENTS)
    ReDim AllowRepositionAtTrap(1 To MAXCLIENTS)
    ReDim NextLootStart(1 To MAXCLIENTS)
    ReDim MAXTIMEINLOOTQUEUE(1 To MAXCLIENTS)
    ReDim MINDELAYTOLOOT(1 To MAXCLIENTS)
    ReDim OldLootMode(1 To MAXCLIENTS)
    ReDim LootAll(1 To MAXCLIENTS)
    ReDim PKwarnings(1 To MAXCLIENTS)
    ReDim DoingNewLoot(1 To MAXCLIENTS)
ReDim DoingNewLootX(1 To MAXCLIENTS)
ReDim DoingNewLootY(1 To MAXCLIENTS)
ReDim DoingNewLootZ(1 To MAXCLIENTS)
ReDim MAXTIMETOREACHCORPSE(1 To MAXCLIENTS)
ReDim DoingNewLootMAXGTC(1 To MAXCLIENTS)

  ReDim Aux_LastLoadedCond(1 To MAXCLIENTS)
  ReDim stealthLog(1 To MAXCLIENTS)
  ReDim LastHealTime(1 To MAXCLIENTS)
  ReDim lastLootOrder(1 To MAXCLIENTS)
  ReDim LastCavebotTime(1 To MAXCLIENTS)
  ReDim CavebotStartTime(1 To MAXCLIENTS)
  ReDim IgnoreServer(1 To MAXCLIENTS)
  ReDim FirstCharInCharList(1 To MAXCLIENTS)
  ReDim NoHealingNextTurn(1 To MAXCLIENTS)
  ReDim DropDelayerTurn(1 To MAXCLIENTS)
  ReDim ConnectionSignal(1 To MAXCLIENTS)
  ReDim usingPriorities(1 To MAXCLIENTS)
  ReDim gotFirstLoginPacket(1 To MAXCLIENTS)
  ReDim packetKey(1 To MAXCLIENTS)
  ReDim loginPacketKey(1 To MAXCLIENTS)
  ReDim CharacterList2(1 To MAXCLIENTS)

  ReDim ConnectionBuffer(1 To MAXCLIENTS)
  ReDim ConnectionBufferLogin(1 To MAXCLIENTS)
  ReDim Connected(1 To MAXCLIENTS)
  ReDim GameConnected(1 To MAXCLIENTS)
  ReDim NeedToIgnoreFirstGamePacket(1 To MAXCLIENTS)
  ReDim MustCheckFirstClientPacket(1 To MAXCLIENTS)
  ReDim Backpack(1 To MAXCLIENTS, 0 To HIGHEST_BP_ID)
  ReDim IDstring(1 To MAXCLIENTS)
  ReDim myX(1 To MAXCLIENTS)
  ReDim myY(1 To MAXCLIENTS)
  ReDim myZ(1 To MAXCLIENTS)
  ReDim myHP(1 To MAXCLIENTS)
  ReDim myMaxHP(1 To MAXCLIENTS)
  ReDim myMaxMana(1 To MAXCLIENTS)
  ReDim myNewStat(1 To MAXCLIENTS)
  ReDim myMana(1 To MAXCLIENTS)
  ReDim myCap(1 To MAXCLIENTS)
  ReDim myStamina(1 To MAXCLIENTS)
  ReDim mySoulpoints(1 To MAXCLIENTS)
  ReDim mySlot(1 To MAXCLIENTS, 1 To EQUIPMENT_SLOTS)
  ReDim savedItem(1 To MAXCLIENTS)
  ReDim CharacterName(1 To MAXCLIENTS)
  ReDim sentFirstPacket(1 To MAXCLIENTS)
  ReDim sentWelcome(1 To MAXCLIENTS)
  ReDim Matrix(-6 To 7, -8 To 9, 0 To 15, 1 To MAXCLIENTS) ' y, x, z, idConnection
  ReDim LogoutReason(1 To MAXCLIENTS)
  ReDim NameOfID(1 To MAXCLIENTS)
  ReDim HPOfID(1 To MAXCLIENTS)
  ReDim DirectionOfID(1 To MAXCLIENTS)
  ReDim GotPacketWarning(1 To MAXCLIENTS)
  ReDim DatTiles(0 To MAXDATTILES)
  ReDim RuneMakerOptions(1 To MAXCLIENTS)
  'custom ng
  ReDim healingCheatsOptions(1 To MAXCLIENTS)
  ReDim extrasOptions(1 To MAXCLIENTS)
  ReDim persistentOptions(1 To MAXCLIENTS)
  ReDim aimbotOptions(1 To MAXCLIENTS)
  '
  ReDim AfterLoginLogoutReason(1 To MAXCLIENTS)
  ReDim myExp(1 To MAXCLIENTS)
  ReDim myLevel(1 To MAXCLIENTS)
  ReDim myMagLevel(1 To MAXCLIENTS)
  ReDim myInitialExp(1 To MAXCLIENTS)
  ReDim myInitialTickCount(1 To MAXCLIENTS)
  ReDim cavebotLenght(1 To MAXCLIENTS)
  ReDim cavebotScript(1 To MAXCLIENTS)
  ReDim debugPIDs(1 To MAXCLIENTS)
  ReDim AllowedLootDistance(1 To MAXCLIENTS)
  ReDim cavebotEnabled(1 To MAXCLIENTS)
  ReDim EnableMaxAttackTime(1 To MAXCLIENTS)
  ReDim myID(1 To MAXCLIENTS)
  ReDim exeLine(1 To MAXCLIENTS)
  ReDim ProcessID(1 To MAXCLIENTS)
  ReDim fishCounter(1 To MAXCLIENTS)
  ReDim waitCounter(1 To MAXCLIENTS)
  ReDim pushTarget(1 To MAXCLIENTS)
  ReDim pushDelay(1 To MAXCLIENTS)
  ReDim cavebotMelees(1 To MAXCLIENTS)
  ReDim cavebotExorivis(1 To MAXCLIENTS)
  ReDim cavebotAvoid(1 To MAXCLIENTS)
  ReDim cavebotHMMs(1 To MAXCLIENTS)
  ReDim DictSETUSEITEM(1 To MAXCLIENTS)
  ReDim DictSETUSEITEM_used(1 To MAXCLIENTS)
  ReDim SETUSEITEM_lastX(1 To MAXCLIENTS)
  ReDim SETUSEITEM_lastY(1 To MAXCLIENTS)
  ReDim shotTypeDict(1 To MAXCLIENTS)
  ReDim exoriTypeDict(1 To MAXCLIENTS)
  ReDim LogoutTimeGM(1 To MAXCLIENTS)
  ReDim DangerGM(1 To MAXCLIENTS)
  ReDim GMname(1 To MAXCLIENTS)
  ReDim DangerPK(1 To MAXCLIENTS)
  ReDim DangerPlayer(1 To MAXCLIENTS)
  ReDim cavebotOnDanger(1 To MAXCLIENTS)
  ReDim cavebotOnGMclose(1 To MAXCLIENTS)
  ReDim cavebotOnGMpause(1 To MAXCLIENTS)
  ReDim cavebotOnPLAYERpause(1 To MAXCLIENTS)
  ReDim lastAttackedID(1 To MAXCLIENTS)
  ReDim CavebotTimeWithSameTarget(1 To MAXCLIENTS)
  ReDim CavebotTimeStart(1 To MAXCLIENTS)
  ReDim maxAttackTime(1 To MAXCLIENTS)
  ReDim maxAttackTimeCHAOS(1 To MAXCLIENTS)
  ReDim maxHit(1 To MAXCLIENTS)
  ReDim previousAttackedID(1 To MAXCLIENTS)
  ReDim moveRetry(1 To MAXCLIENTS)
  ReDim lastX(1 To MAXCLIENTS)
  ReDim lastY(1 To MAXCLIENTS)
  ReDim lastZ(1 To MAXCLIENTS)
  ReDim setFollowTarget(1 To MAXCLIENTS)
  ReDim myLastCorpseX(1 To MAXCLIENTS)
  ReDim myLastCorpseY(1 To MAXCLIENTS)
  ReDim myLastCorpseZ(1 To MAXCLIENTS)
  ReDim myLastCorpseS(1 To MAXCLIENTS)
  ReDim lastIngameCheck(1 To MAXCLIENTS)
  ReDim lastIngameCheckTileID(1 To MAXCLIENTS)
  ReDim myLastCorpseTileID(1 To MAXCLIENTS)
  ReDim lootWaiting(1 To MAXCLIENTS)
  ReDim autoLoot(1 To MAXCLIENTS)
  ReDim requestLootBp(1 To MAXCLIENTS)
  ReDim lootTimeExpire(1 To MAXCLIENTS)
  ReDim cavebotGoodLoot(1 To MAXCLIENTS)
  ReDim killPriorities(1 To MAXCLIENTS)
  ReDim SpellKills_SpellName(1 To MAXCLIENTS)
  ReDim SpellKills_Dist(1 To MAXCLIENTS)
  ReDim DangerGMname(1 To MAXCLIENTS)
  ReDim DangerPlayerName(1 To MAXCLIENTS)
  ReDim DangerPKname(1 To MAXCLIENTS)
  ReDim SendingSpecialOutfit(1 To MAXCLIENTS)
  ReDim currTargetName(1 To MAXCLIENTS)
  ReDim currTargetID(1 To MAXCLIENTS)
  ReDim friendlyMode(1 To MAXCLIENTS)
  ReDim receivedLogin(1 To MAXCLIENTS)
  ReDim ignoreNext(1 To MAXCLIENTS)
  ReDim lastDestX(1 To MAXCLIENTS)
  ReDim lastDestY(1 To MAXCLIENTS)
  ReDim lastDestZ(1 To MAXCLIENTS)
  ReDim DoingMainLoop(1 To MAXCLIENTS)
  ReDim DoingMainLoopLogin(1 To MAXCLIENTS)
  ReDim lastFloorTrap(1 To MAXCLIENTS)
  ReDim nextLight(1 To MAXCLIENTS)
  ReDim onDepotPhase(1 To MAXCLIENTS)
  ReDim CavebotChaoticMode(1 To MAXCLIENTS)
  ReDim bLevelSpy(1 To MAXCLIENTS)
  ReDim depotX(1 To MAXCLIENTS)
  ReDim depotY(1 To MAXCLIENTS)
  ReDim depotZ(1 To MAXCLIENTS)
  ReDim depotS(1 To MAXCLIENTS)
  ReDim lastDepotBPID(1 To MAXCLIENTS)
  ReDim depotTileID(1 To MAXCLIENTS)
  ReDim doneDepotChestOpen(1 To MAXCLIENTS)
  ReDim somethingChangedInBps(1 To MAXCLIENTS)
  ReDim nextForcedDepotDeployRetry(1 To MAXCLIENTS)
  ReDim lastFloorChangeX(1 To MAXCLIENTS)
  ReDim lastFloorChangeY(1 To MAXCLIENTS)
  ReDim lastFloorChangeZ(1 To MAXCLIENTS)
  ReDim prevAttackState(1 To MAXCLIENTS)
  ReDim TurnsWithRedSquareZero(1 To MAXCLIENTS)
  ReDim cancelAllMove(1 To MAXCLIENTS)
  ReDim LoginMsgCount(1 To MAXCLIENTS)
  ReDim lastHPchange(1 To MAXCLIENTS)
  ReDim StatusBits(1 To MAXCLIENTS)
  ReDim CheatsPaused(1 To MAXCLIENTS)
  ReDim SpamAutoHeal(1 To MAXCLIENTS)
  ReDim SpamAutoMana(1 To MAXCLIENTS)

  ReDim SpamAutoFastHeal(1 To MAXCLIENTS)
  ReDim nextFastHeal(1 To MAXCLIENTS)
  ReDim SpamAutoPush(1 To MAXCLIENTS)
  ReDim RequiredMoveBuffer(1 To MAXCLIENTS)
  ReDim ReadyBuffer(1 To MAXCLIENTS)
  ReDim initialRuneBackpack(1 To MAXCLIENTS)
  ReDim posSpamActivated(1 To MAXCLIENTS)
  ReDim posSpamChannelB1(1 To MAXCLIENTS)
  ReDim posSpamChannelB2(1 To MAXCLIENTS)
  ReDim getSpamActivated(1 To MAXCLIENTS)
  ReDim getSpamChannelB1(1 To MAXCLIENTS)
  ReDim getSpamChannelB2(1 To MAXCLIENTS)
  ReDim lastAttackedIDstatus(1 To MAXCLIENTS)
  ReDim executingCavebot(1 To MAXCLIENTS)
  ReDim lastPing(1 To MAXCLIENTS)
  ReDim XYZCountdowns(1 To MAXCLIENTS, 1 To MAXCLIENTS) 'TODO: currently, we can only keep count of MAXCLIENTS walls per character. fix it so we can keep count of unlimited number of walls per character..
  ReDim doingTrade(1 To MAXCLIENTS)
  ReDim doingTrade2(1 To MAXCLIENTS)
  ReDim GotKillOrderTargetID(1 To MAXCLIENTS)
  ReDim GotKillOrder(1 To MAXCLIENTS)
  ReDim GotKillOrderTargetName(1 To MAXCLIENTS)
  ReDim cavebotOnTrapGiveAlarm(1 To MAXCLIENTS)
  ReDim cavebotCurrentTargetPriority(1 To MAXCLIENTS)
  ReDim AllowUHpaused(1 To MAXCLIENTS)
  ReDim pauseStacking(1 To MAXCLIENTS)
  ReDim ReconnectionStage(1 To MAXCLIENTS)
  ReDim runeTurn(1 To MAXCLIENTS)
  ReDim ReconnectionPacket(1 To MAXCLIENTS)
  ReDim logoutAllowed(1 To MAXCLIENTS)
  ReDim SelfDefenseID(1 To MAXCLIENTS)
  ReDim CustomEvents(1 To MAXCLIENTS)
  ReDim CustomCondEvents(1 To MAXCLIENTS)
  ReDim nextAllowedmsg(1 To MAXCLIENTS)
  ReDim var_expleft(1 To MAXCLIENTS)
  ReDim var_nextlevel(1 To MAXCLIENTS)
  ReDim var_exph(1 To MAXCLIENTS)
  ReDim var_timeleft(1 To MAXCLIENTS)
  ReDim var_played(1 To MAXCLIENTS)
  ReDim var_expgained(1 To MAXCLIENTS)
  ReDim var_lf(1 To MAXCLIENTS)
  ReDim var_lastsender(1 To MAXCLIENTS)
  ReDim var_lastmsg(1 To MAXCLIENTS)
  ReDim DelayAttacks(1 To MAXCLIENTS)
  ReDim AvoidReAttacks(1 To MAXCLIENTS)
  ReDim TrainerOptions(1 To MAXCLIENTS)
  ReDim reconnectionRetryCount(1 To MAXCLIENTS)
  ReDim nextReconnectionRetry(1 To MAXCLIENTS)
  ReDim UHRetryCount(1 To MAXCLIENTS)
  ReDim runemakerMana1(1 To MAXCLIENTS)
  ReDim makingRune(1 To MAXCLIENTS)
  ReDim lastUsedChannelID(1 To MAXCLIENTS)
  ReDim lastRecChannelID(1 To MAXCLIENTS)
  ReDim CavebotHaveSpecials(1 To MAXCLIENTS)
  ReDim CavebotLastSpecialMove(1 To MAXCLIENTS)

  For i = 1 To MAXCLIENTS
    ReDim CustomEvents(i).ev(1 To MAXEVENTS)
    ReDim CustomCondEvents(i).ev(1 To MAXCONDS)
  Next i

  ' Read some tile ID values from the ini :
  
  ' runes
  ReadTileIDListFromIni AditionalStairsToUpFloor, "AditionalStairsToUpFloor", here, "AC 07,AE 07,AA 07,94 08,96 08,90 08,92 08"
  ReadTileIDListFromIni AditionalStairsToDownFloor, "AditionalStairsToDownFloor", here, ""
  ReadTileIDListFromIni AditionalRequireRope, "AditionalRequireRope", here, ""
  ReadTileIDListFromIni AditionalRequireShovel, "AditionalRequireShovel", here, ""
  
  ReadTileIDFromIni tileID_Blank, "tileID_Blank", here, "0D 0C"

  ReadTileIDFromIni tileID_WallBugItem, "tileID_WallBugItem", here, "4E 10"
  
  blank1 = LowByteOfLong(tileID_Blank)
  blank2 = HighByteOfLong(tileID_Blank)
  
  ReadTileIDFromIni tileID_SD, "tileID_SD", here, "53 0C"
  ReadTileIDFromIni tileID_HMM, "tileID_HMM", here, "40 0C"
  ReadTileIDFromIni tileID_Explosion, "tileID_Explosion", here, "42 0C"
  ReadTileIDFromIni tileID_IH, "tileID_IH", here, "12 0C"
  ReadTileIDFromIni tileID_UH, "tileID_UH", here, "1A 0C"
  
  ReadTileIDFromIni tileID_fireball, "tileID_fireball", here, "75 0C"
  ReadTileIDFromIni tileID_stalagmite, "tileID_stalagmite", here, "6B 0C"
  ReadTileIDFromIni tileID_icicle, "tileID_icicle", here, "56 0C"
  
  ' items
  ReadTileIDFromIni tileID_Bag, "tileID_Bag", here, "E7 0A"
  ReadTileIDFromIni tileID_Backpack, "tileID_Backpack", here, "E8 0A"
  ReadTileIDFromIni tileID_Oracle, "tileID_Oracle", here, "DA 07"
  ReadTileIDFromIni tileID_FishingRod, "tileID_FishingRod", here, "5D 0D"
 
  ReadTileIDFromIni tileID_Rope, "tileID_Rope", here, "7D 0B"
  ReadTileIDFromIni tileID_LightRope, "tileID_LightRope", here, "86 02"
  ReadTileIDFromIni tileID_Shovel, "tileID_Shovel", here, "43 0D"
  ReadTileIDFromIni tileID_LightShovel, "tileID_LightShovel", here, "4E 16"

  ' water
  ReadTileIDFromIni tileID_waterEmpty, "tileID_waterEmpty", here, "5B 02"
  ReadTileIDFromIni tileID_waterWithFish, "tileID_waterWithFish", here, "59 02"
  
  ReadTileIDFromIni tileID_waterEmptyEnd, "tileID_waterEmptyEnd", here, "5B 02"
  ReadTileIDFromIni tileID_waterWithFishEnd, "tileID_waterWithFishEnd", here, "59 02"
  
  ' blocking table
  ReadTileIDFromIni tileID_blockingBox, "tileID_blockingBox", here, "A5 09"
  
  ' to UP floor
  ReadTileIDFromIni tileID_stairsToUp, "tileID_stairsToUp", here, "88 07"
  ReadTileIDFromIni tileID_woodenStairstoUp, "tileID_woodenStairstoUp", here, "93 07"
  
  ReadTileIDFromIni tileID_desertRamptoUp, "tileID_desertRamptoUp", here, "A8 07"
  
  ReadTileIDFromIni tileID_rampToNorth, "tileID_rampToNorth", here, "91 07"
  ReadTileIDFromIni tileID_rampToSouth, "tileID_rampToSouth", here, "8F 07"
 
  ReadTileIDFromIni tileID_rampToRightCycMountain, "tileID_rampToRightCycMountain", here, "8B 07"
  ReadTileIDFromIni tileID_rampToLeftCycMountain, "tileID_rampToLeftCycMountain", here, "8D 07"
  
  
  ReadTileIDFromIni tileID_jungleStairsToNorth, "tileID_jungleStairsToNorth", here, "B9 07"
  ReadTileIDFromIni tileID_jungleStairsToLeft, "tileID_jungleStairsToLeft", here, "BA 07"
  
  ' + requires rightClick
  ReadTileIDFromIni tileID_ladderToUp, "tileID_ladderToUp", here, "89 07"
  
  ' + requires rope
  ReadTileIDFromIni tileID_holeInCelling, "tileID_holeInCelling", here, "80 01"
  
  ' to DOWN
  ReadTileIDFromIni tileID_grassCouldBeHole, "tileID_grassCouldBeHole", here, "25 01"
  ReadTileIDFromIni tileID_pitfall, "tileID_pitfall", here, "26 01"

  ReadTileIDFromIni tileID_openHole, "tileID_openHole", here, "44 02"
  ReadTileIDFromIni tileID_OpenDesertLooseStonePile, "tileID_OpenDesertLooseStonePile", here, "51 02"
  
  
  ReadTileIDFromIni tileID_trapdoor, "tileID_trapdoor", here, "71 01"
  ReadTileIDFromIni tileID_down1, "tileID_down1", here, "72 01"
  
  ReadTileIDFromIni tileID_openHole2, "tileID_openHole2", here, "7F 01"
  
  ReadTileIDFromIni tileID_trapdoor2, "tileID_trapdoor2", here, "98 01"
  ReadTileIDFromIni tileID_down2, "tileID_down2", here, "99 01"
  ReadTileIDFromIni tileID_stairsToDownKazordoon, "tileID_stairsToDownKazordoon", here, "9A 01"
  ReadTileIDFromIni tileID_stairsToDownThais, "tileID_stairsToDownThais", here, "9B 01"
  
  ReadTileIDFromIni tileID_trapdoorKazordoon, "tileID_trapdoorKazordoon", here, "AB 01"
  ReadTileIDFromIni tileID_down3, "tileID_down3", here, "AC 01"
  ReadTileIDFromIni tileID_stairsToDown, "tileID_stairsToDown", here, "AD 01"
  
  ReadTileIDFromIni tileID_stairsToDown2, "tileID_stairsToDown2", here, "B0 01"
  ReadTileIDFromIni tileID_woodenStairstoDown, "tileID_woodenStairstoDown", here, "B1 01"
  
  ReadTileIDFromIni tileID_rampToDown, "tileID_rampToDown", here, "CB 01"

  ' + requires rightClick
  ReadTileIDFromIni tileID_sewerGate, "tileID_sewerGate", here, "AE 01"

  ' + requires shovel
  ReadTileIDFromIni tileID_closedHole, "tileID_closedHole", here, "43 02"
  ReadTileIDFromIni tileID_desertLooseStonePile, "tileID_desertLooseStonePile", here, "50 02"
  
  ' FOOD
  ReadTileIDFromIni tileID_firstFoodTileID, "tileID_firstFoodTileID", here, "BB 0D"
  ReadTileIDFromIni tileID_lastFoodTileID, "tileID_lastFoodTileID", here, "D9 0D"
  ReadTileIDFromIni tileID_firstMushroomTileID, "tileID_firstMushroomTileID", here, "4A 0E"
  ReadTileIDFromIni tileID_lastMushroomTileID, "tileID_lastMushroomTileID", here, "4E 0E"
  
  'FIELD RANGE1
  ReadTileIDFromIni tileID_firstFieldRangeStart, "tileID_firstFieldRangeStart", here, "31 08"
  ReadTileIDFromIni tileID_firstFieldRangeEnd, "tileID_firstFieldRangeEnd", here, "3A 08"
  ReadTileIDFromIni tileID_secondFieldRangeStart, "tileID_secondFieldRangeStart", here, "3E 08"
  ReadTileIDFromIni tileID_secondFieldRangeEnd, "tileID_secondFieldRangeEnd", here, "45 08"

  ReadTileIDFromIni tileID_campFire1, "tileID_campFire1", here, "20 20"
  ReadTileIDFromIni tileID_campFire2, "tileID_campFire2", here, "20 20"

  'WALKABLE FIELDS
  ReadTileIDFromIni tileID_walkableFire1, "tileID_walkableFire1", here, "33 08"
  ReadTileIDFromIni tileID_walkableFire2, "tileID_walkableFire2", here, "38 08"
  ReadTileIDFromIni tileID_walkableFire3, "tileID_walkableFire3", here, "40 08"
  
  ' Depot chest
  ReadTileIDFromIni tileID_depotChest, "tileID_depotChest", here, "70 0D"
  
  ' flasks - mana fluids
  ReadTileIDFromIni tileID_flask, "tileID_flask", here, "3A 0B"
  
  
  ReadTileIDFromIni tileID_health_potion, "tileID_health_potion", here, "0A 01"
  ReadTileIDFromIni tileID_strong_health_potion, "tileID_strong_health_potion", here, "EC 00"
  ReadTileIDFromIni tileID_great_health_potion, "tileID_great_health_potion", here, "EF 00"
  ReadTileIDFromIni tileID_small_health_potion, "tileID_small_health_potion", here, "C4 1E"
  ReadTileIDFromIni tileID_mana_potion, "tileID_mana_potion", here, "0C 01"
  ReadTileIDFromIni tileID_strong_mana_potion, "tileID_strong_mana_potion", here, "ED 00"
  ReadTileIDFromIni tileID_great_mana_potion, "tileID_great_mana_potion", here, "EE 00"
  
  ReadTileIDFromIni tileID_ultimate_health_potion, "tileID_ultimate_health_potion", here, "DB 1D"
  ReadTileIDFromIni tileID_great_spirit_potion, "tileID_great_spirit_potion", here, "DA 1D"
  
  strInfo = String$(10, 0)
  i = getBlackdINI("tileIDs", "byteNothing", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    byteNothing = lonInfo
  Else
    byteNothing = &H0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("tileIDs", "byteMana", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    byteMana = lonInfo
  Else
    byteMana = &H7
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("tileIDs", "byteLife", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    byteLife = lonInfo
  Else
    byteLife = &HB
  End If
  



  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "TrainerTimer1", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TrainerTimer1 = lonInfo
  Else
    TrainerTimer1 = 300
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "TrainerTimer2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TrainerTimer2 = lonInfo
  Else
    TrainerTimer2 = 1000
  End If
  
  
  
  
  
  res = 0
  Exit Function
gotErr:
  MsgBox "Sorry, Blackd Proxy was not able to read .ini files (start)" & vbCrLf & "Possible reasons:" & vbCrLf & _
  " - Corrupted settings.ini ?" & vbCrLf & _
  " - Corrupted config.ini ?" & vbCrLf & _
  " - Not enough privileges to read the required files ?" & vbCrLf & _
  " Details:" & vbCrLf & _
  " - Path settins.ini: " & userHere & vbCrLf & _
  " - Path config.ini: " & here & vbCrLf & _
  " - Debug Point: " & debugPoint & vbCrLf & _
  " - Error number: " & Err.Number & vbCrLf & _
  " - Error description: " & Err.Description, vbOKOnly + vbCritical, "Critical error"
  End
End Function

Public Sub ReadIni()
  ' Read the rest of vars from the ini
  Dim i As Integer
  Dim strInfo As String
  Dim lonInfo As Long
  Dim here As String
  Dim tmp As String
  Dim idLoginSP As Long
  Dim ibucle As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  userHere = App.Path & "\settings.ini" ' user config file name
  If configPath = "" Then
    here = myMainConfigINIPath()
  Else
    here = App.Path & "\" & configPath & "\config.ini"
  End If
  strInfo = String$(255, 0)
  i = getBlackdINI("Proxy", "ForwardGameTo", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  ForwardGameTo.Text = strInfo
  strInfo = String$(50, 0)
  i = getBlackdINI("Proxy", "txtServerLoginP", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  txtServerLoginP.Text = strInfo
  strInfo = String$(50, 0)
  i = getBlackdINI("Proxy", "txtServerGameP", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  txtServerGameP.Text = strInfo
  
  
  If (OVERWRITE_OT_MODE = True) Then
    ForwardGameTo.Text = OVERWRITE_OT_IP
    txtServerLoginP.Text = OVERWRITE_OT_PORT

  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "ForwardOption", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  

  If (Not (OVERWRITE_CONFIGPATH = "")) Then
    If OVERWRITE_OT_MODE = True Then
        strInfo = "3"
    Else
        strInfo = "1"
    End If
  End If
  
  
  If strInfo = "3" Then
    TrueServer1.value = False
    TrueServer2.value = False
    TrueServer3.value = True
    TrueServer3_Click
  ElseIf strInfo = "2" Then
    TrueServer1.value = False
    TrueServer2.value = True
    TrueServer3.value = False
    TrueServer2_Click
  Else
    TrueServer1.value = True
    TrueServer2.value = False
    TrueServer3.value = False
    TrueServer1_Click
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "MapClickOption", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmHardcoreCheats.ActionInspect.value = True
    frmHardcoreCheats.ActionMove.value = False
    frmHardcoreCheats.ActionNothing.value = False
    frmHardcoreCheats.ActionPath.value = False
  ElseIf strInfo = "3" Then
    frmHardcoreCheats.ActionInspect.value = False
    frmHardcoreCheats.ActionMove.value = False
    frmHardcoreCheats.ActionNothing.value = True
    frmHardcoreCheats.ActionPath.value = False
  ElseIf strInfo = "2" Then
    frmHardcoreCheats.ActionInspect.value = False
    frmHardcoreCheats.ActionMove.value = True
    frmHardcoreCheats.ActionNothing.value = False
    frmHardcoreCheats.ActionPath.value = False
  Else '4
    frmHardcoreCheats.ActionInspect.value = False
    frmHardcoreCheats.ActionMove.value = False
    frmHardcoreCheats.ActionNothing.value = False
    frmHardcoreCheats.ActionPath.value = True
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("MemoryAddresses", "serverLogoutMessage", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    If (Trim$(strInfo) = "") Then
      serverLogoutMessage = "You have been idle for 15 minutes. You will be disconnected in one minute if you are still idle then."
    Else
      serverLogoutMessage = Trim$(strInfo)
    End If
  Else
    serverLogoutMessage = "You have been idle for 15 minutes. You will be disconnected in one minute if you are still idle then."
  End If

  
  strInfo = String$(10, 0)
  i = getBlackdINI("Proxy", "ShowAdvancedOptions", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    If lonInfo = 0 Then
      blnShowAdvancedOptions = True
      cmdAdvanced_Click
    Else
      blnShowAdvancedOptions = False
      cmdAdvanced_Click
    End If
  Else
    blnShowAdvancedOptions = False
    cmdAdvanced_Click
  End If
  If ((TibiaVersionLong <= 760) Or (ForceDisableEncryption = True)) Then
    UseCrackd = False
  Else
    UseCrackd = True
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Log", "LogPacketsEnabled", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    chkLogPackets.value = lonInfo
    chkLogPackets_Click
  Else
    chkLogPackets.value = 0
    chkLogPackets_Click
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "LocalLoginUseProxy", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    chckMemoryIP.value = lonInfo
  Else
    chckMemoryIP.value = 1
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "LocalGameUseProxy", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    chckAlter.value = lonInfo
  Else
    chckAlter.value = 1
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "ListenLoginPort", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    txtClientLoginP.Text = strInfo
    txtClientLoginP_Validate False
  Else
    txtClientLoginP.Text = "15000"
    txtClientLoginP_Validate False
  End If
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "ListenGamePort", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    txtClientGameP.Text = strInfo
    txtClientGameP_Validate False
  Else
    txtClientGameP.Text = "16000"
    txtClientGameP_Validate False
  End If

  
  
  'CteMoveDelay
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "CteMoveDelay", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    CteMoveDelay = lonInfo
  Else
    CteMoveDelay = 700
  End If
  
  'TimeToGiveTrapAlarm
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "TimeToGiveTrapAlarm", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TimeToGiveTrapAlarm = lonInfo
  Else
    TimeToGiveTrapAlarm = 45000
  End If
  
  If (Not (OVERWRITE_MAPS_PATH = "")) Then
    TibiaPath = OVERWRITE_MAPS_PATH
    TibiaPath = ValidateTibiaPath(OVERWRITE_MAPS_PATH)
    If TibiaPath = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->" Then
      TibiaPath = ""
    End If
  Else
  
'  strInfo = String$(250, 0)
'  i = getBlackdINI("Proxy", "TibiaPath", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'    TibiaPath = strInfo
'    txtTibiaPath = strInfo
'    TibiaPath = ValidateTibiaPath(strInfo)
'    If TibiaPath = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->" Then
'      TibiaPath = ""
'    End If
'  Else
'    TibiaPath = "C:\Archivos de programa\Tibia"
'    txtTibiaPath = "C:\Archivos de programa\Tibia"
'    TibiaPath = ValidateTibiaPath("C:\Archivos de programa\Tibia")
'    If TibiaPath = "PATH NOT CONFIGURED! USE THIS BUTTON TO BROWSE -->" Then
'      TibiaPath = ""
'    End If
'  End If
  
  End If
  
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "GmStartWithThis3", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    tmp = Left(strInfo, Len(strInfo) - 1)
    strInfo = tmp
    If Len(strInfo) = 3 Then
      gmStart = strInfo
    Else
      gmStart = "gm "
    End If
  Else
    gmStart = "gm "
  End If
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "AltGmStartWithThis3", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    tmp = Left(strInfo, Len(strInfo) - 1)
    strInfo = tmp
    If Len(strInfo) = 3 Then
      gmStart2 = strInfo
    Else
      gmStart2 = "cm "
    End If
  Else
    gmStart2 = "cm "
  End If
  
  

  
  
  For idLoginSP = 1 To NumberOfLoginServers
  strInfo = String$(250, 0)
  i = getBlackdINI("MemoryAddresses", "trueLoginServer" & CStr(idLoginSP), "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    trueLoginServer(idLoginSP) = strInfo
  Else
    If TibiaVersionLong >= 800 Then
        Select Case idLoginSP
        Case 1
            trueLoginServer(idLoginSP) = "login01.tibia.com"
        Case 2
            trueLoginServer(idLoginSP) = "login02.tibia.com"
        Case 3
            trueLoginServer(idLoginSP) = "login03.tibia.com"
        Case 4
            trueLoginServer(idLoginSP) = "login04.tibia.com"
        Case 5
            trueLoginServer(idLoginSP) = "login05.tibia.com"
        Case 6
            trueLoginServer(idLoginSP) = "tibia01.cipsoft.com"
        Case 7
            trueLoginServer(idLoginSP) = "tibia02.cipsoft.com"
        Case 8
            trueLoginServer(idLoginSP) = "tibia03.cipsoft.com"
        Case 9
            trueLoginServer(idLoginSP) = "tibia04.cipsoft.com"
        Case 10
            trueLoginServer(idLoginSP) = "tibia05.cipsoft.com"
        End Select
    Else
        Select Case idLoginSP
        Case 1
            trueLoginServer(idLoginSP) = "server.tibia.com"
        Case 2
            trueLoginServer(idLoginSP) = "server2.tibia.com"
        Case 3
            trueLoginServer(idLoginSP) = "tibia1.cipsoft.com"
        Case 4
            trueLoginServer(idLoginSP) = "tibia2.cipsoft.com"
        Case 5
            trueLoginServer(idLoginSP) = "server2.tibia.com"
        End Select

    End If
    
  End If
  Next idLoginSP
  






  For idLoginSP = 1 To NumberOfLoginServers
  strInfo = String$(250, 0)
  i = getBlackdINI("MemoryAddresses", "trueLoginPort" & CStr(idLoginSP), "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    trueLoginPort(idLoginSP) = CStr(lonInfo)
  Else
    If TibiaVersionLong >= 800 Then
      
         trueLoginPort(idLoginSP) = "7171"
       
    Else
        If idLoginSP = 5 Then
            trueLoginPort(idLoginSP) = "7172"
        Else
            trueLoginPort(idLoginSP) = "7171"
        End If
    End If
  End If
  Next idLoginSP




  strInfo = String$(250, 0)
  i = getBlackdINI("MemoryAddresses", "PREFEREDLOGINSERVER", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    PREFEREDLOGINSERVER = strInfo
  Else
    If TibiaVersionLong >= 800 Then
        PREFEREDLOGINSERVER = "login01.tibia.com"
    Else
        PREFEREDLOGINSERVER = "tibia1.cipsoft.com"
    End If
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("MemoryAddresses", "PREFEREDLOGINPORT", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    PREFEREDLOGINPORT = lonInfo
  Else
    PREFEREDLOGINPORT = 7171
  End If
  

'  strInfo = String$(250, 0)
'  i = getBlackdINI("MemoryAddresses", "TibiaExePath", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'     TibiaExePath = strInfo
'  Else
'     TibiaExePath = autoGetTibiaFolder()
'  End If
  'MagebotPath = "" 'not used
  
'  strInfo = String$(250, 0)
'  i = getBlackdINI("MemoryAddresses", "MagebotPath", "", strInfo, Len(strInfo), here)
'  If i > 0 Then
'    strInfo = Left(strInfo, i)
'     MagebotPath = strInfo
'  Else
'     MagebotPath = autoGetMagebotFolder()
'  End If

  'MagebotExe = "" 'not used
  'MagebotExe = autoGetMagebotExe()



  strInfo = String$(10, 0)
  i = getBlackdINI("HPmana", "LimitRandomizator", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    LimitRandomizator = lonInfo
  Else
    LimitRandomizator = 10
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("HPmana", "HPmanaRECAST", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    HPmanaRECAST = lonInfo
  Else
    HPmanaRECAST = 300
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("HPmana", "HPmanaRECAST2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    HPmanaRECAST2 = lonInfo
  Else
    HPmanaRECAST2 = 700
  End If
  

  

  Load frmHPmana
  frmHPmana.Hide

  cmbPrefered.Clear
  For idLoginSP = 1 To NumberOfLoginServers
    cmbPrefered.AddItem trueLoginServer(idLoginSP)
  Next idLoginSP

  cmbPrefered.Text = PREFEREDLOGINSERVER
  
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Hotkeys", "HotkeysActivated", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHotkeys.chkHotkeysActivated.value = lonInfo
  Else
    frmHotkeys.chkHotkeysActivated.value = 1
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Hotkeys", "RepeatEnabled", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHotkeys.chkRepeat.value = lonInfo
  Else
    frmHotkeys.chkRepeat.value = 0
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Hotkeys", "RepeatDelay", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHotkeys.txtDelay.Text = strInfo
  Else
    frmHotkeys.txtDelay.Text = "500"
  End If
  
  
  

  
  
  
  
  
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "TimerInterval", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    CavebotRECAST = lonInfo
  Else
    CavebotRECAST = 300
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "TimerInterval2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    CavebotRECAST2 = lonInfo
  Else
    CavebotRECAST2 = 700
  End If
  frmCavebot.txtMs.Text = CStr(CavebotRECAST)
  frmCavebot.txtMs2.Text = CStr(CavebotRECAST2)
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "AutoChangePKHeal", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmCavebot.chkChangePkHeal.value = 0
  Else
    frmCavebot.chkChangePkHeal.value = 1
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "NewAutohealForPKattack", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmCavebot.scrollPkHeal.value = lonInfo
    frmCavebot.scrollPkHeal_Change
  Else
    frmCavebot.scrollPkHeal.value = 75
    frmCavebot.scrollPkHeal_Change
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "SafeHPforExoriVis", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmCavebot.scrollExorivis.value = lonInfo
    frmCavebot.scrollExorivis_Change
  Else
    frmCavebot.scrollExorivis.value = 50
    frmCavebot.scrollExorivis_Change
  End If
  

  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "SetVeryFriendly_NOATTACKTIMER_ms", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    SetVeryFriendly_NOATTACKTIMER_ms = lonInfo
  Else
    SetVeryFriendly_NOATTACKTIMER_ms = 10000
  End If
  
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Log", "MaxLogBuffer", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     txtMaxChar.Text = strInfo
  Else
    txtMaxChar.Text = "30000"
  End If

  strInfo = String$(50, 0)
  i = getBlackdINI("Log", "MaxHexLines", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     txtMaxLines.Text = strInfo
  Else
    txtMaxLines.Text = "3000"
  End If
  
  i = getBlackdINI("Log", "LogFullAction", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "3" Then
    LogFull1.value = False
    LogFull2.value = False
    LogFull3.value = True
  ElseIf strInfo = "2" Then
    LogFull1.value = False
    LogFull2.value = True
    LogFull3.value = False
  Else
    LogFull1.value = True
    LogFull2.value = False
    LogFull3.value = False
  End If
  
  
  

  

  
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Log", "LogFile", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     txtLogFile.Text = strInfo
  Else
    txtLogFile.Text = "log.txt"
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Log", "AutoSelectHexAscii", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    chkSelect.value = lonInfo
  Else
    chkSelect.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Log", "HideHexLogIfLogDisabled", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    chkAutoHide.value = lonInfo
    chkAutoHide_Click
  Else
    chkAutoHide.value = 1
    chkAutoHide_Click
  End If
  
  
   strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LogoutIfDangerAtStart", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmHardcoreCheats.chkLogoutIfDanger.value = 1
  Else
    frmHardcoreCheats.chkLogoutIfDanger.value = 0
  End If
  
  
   strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "chkStealthCommands", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmStealth.chkStealthCommands.value = 0
  Else
    frmStealth.chkStealthCommands.value = 1
  End If
  
   strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "chkStealthMessages", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmStealth.chkStealthMessages.value = 0
  Else
    frmStealth.chkStealthMessages.value = 1
  End If
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "chkStealthExp", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmStealth.chkStealthExp.value = 0
  Else
    frmStealth.chkStealthExp.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "chkAvoidChat", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmStealth.chkAvoidChat.value = 0
  Else
    frmStealth.chkAvoidChat.value = 1
  End If
  
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "RevealInvis", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkReveal.value = 0
  Else
    frmHardcoreCheats.chkReveal.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "ChangeLight", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkLight.value = 0
  Else
    frmHardcoreCheats.chkLight.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "AutoHealEnabled", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkAutoHeal.value = 0
  Else
    frmHardcoreCheats.chkAutoHeal.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "UHalarmEnabled", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkRuneAlarm.value = 0
  Else
    frmHardcoreCheats.chkRuneAlarm.value = 1
  End If
  


  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "AutoVitaEnabled", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmHardcoreCheats.chkAutoVita.value = 1
  Else
    frmHardcoreCheats.chkAutoVita.value = 0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LightPower", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHardcoreCheats.scrollLight.value = lonInfo
    frmHardcoreCheats.scrollLight_Change
  Else
    frmHardcoreCheats.scrollLight.value = 15
    frmHardcoreCheats.scrollLight_Change
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "AcceptSDorder", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmHardcoreCheats.chkAcceptSDorder.value = 1
  Else
    frmHardcoreCheats.chkAcceptSDorder.value = 0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Broadcast", "BroadcastMC", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    BroadcastMC = 1
  Else
    BroadcastMC = 0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Broadcast", "BroadcastDelay1", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    BroadcastDelay1 = lonInfo
  Else
    BroadcastDelay1 = 20000
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Broadcast", "BroadcastDelay2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    BroadcastDelay2 = lonInfo
  Else
    BroadcastDelay2 = 30000
  End If
  
  'custom ng dont load with this
  'frmBroadcast.chkMC.Value = BroadcastMC
  frmBroadcast.txtBroadcastDelay1.Text = BroadcastDelay1
  frmBroadcast.txtBroadcastDelay2.Text = BroadcastDelay2
  
  
  
  
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "SDorder", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtOrder = strInfo
  Else
    frmHardcoreCheats.txtOrder = "firenow"
  End If





  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "SDleader", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtRemoteLeader = strInfo
  Else
    frmHardcoreCheats.txtRemoteLeader = ""
  End If
  
  'ExivaExpPlace
  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "ExivaExpPlace", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    ExivaExpPlace = strInfo
  Else
    ExivaExpPlace = "19 : white center"
  End If
  frmHardcoreCheats.cmbWhere.Text = ExivaExpPlace
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "txtExivaExpFormat", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtExivaExpFormat = strInfo
  Else
    frmHardcoreCheats.txtExivaExpFormat = "You need $expleft$ exp$lf$for level $nextlevel$$lf$Average session speed:$lf$$exph$ exp/h$lf$Estimated time left for level up:$lf$$timeleft$$lf$Played this session:$lf$$played$$lf$Gained this session:$lf$$expgained$ exp"
  End If
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "tibiaTittleFormat", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.tibiaTittleFormat = strInfo
  Else
    frmHardcoreCheats.tibiaTittleFormat = "$charactername$ - $expleft$ exp to lv $nextlevel$ - $exph$ exp/h"
  End If
  
  'custom ng
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Cheats", "ColorEffects", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "1" Then
  '  frmHardcoreCheats.chkColorEffects.Value = 1
  'Else
  '  frmHardcoreCheats.chkColorEffects.Value = 0
  'End If
  
  'chkCaptionExp
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "TitleExp", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkCaptionExp.value = 0
  Else
    frmHardcoreCheats.chkCaptionExp.value = 1
  End If
  
  'custom ng
  'chkAutoGratz
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Cheats", "chkAutoGratz", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "1" Then
  '  frmHardcoreCheats.chkAutoGratz.Value = 1
  'Else
  '  frmHardcoreCheats.chkAutoGratz.Value = 0
  'End If
  
  'chkAutorelog
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "Antibanmode", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    Antibanmode = 0
  Else
    Antibanmode = 1
  End If
  
  'chkAutorelog
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "chkAutorelog", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmHardcoreCheats.chkAutorelog.value = 1
  Else
    frmHardcoreCheats.chkAutorelog.value = 0
  End If
  
  If ((Antibanmode = 1) Or (TibiaVersionLong >= 841)) Then
    'frmHardcoreCheats.chkAutorelog.Value = 0
    'frmHardcoreCheats.chkAutorelog.enabled = False
    If TibiaVersionLong >= 841 Then
       frmHardcoreCheats.chkAutorelog.Caption = "Autorelog disabled since 8.41"
       frmHardcoreCheats.chkAutorelog.value = 0
       frmHardcoreCheats.chkAutorelog.enabled = False
    Else
    frmHardcoreCheats.chkAutorelog.Caption = "Autorelog. WARNING: do not use during server save!"
    frmHardcoreCheats.chkAutorelog.ForeColor = vbYellow
    'frmHardcoreCheats.txtRelogBackpacks.enabled = False
    'frmHardcoreCheats.lblBackpacks.enabled = False
    End If
    frmAdvanced.chkWantBypass.value = 0
    frmAdvanced.chkWantBypass.Caption = "Bypass disabled (antiban mode)"
    frmAdvanced.chkWantBypass.enabled = False
  End If

  
  'custom ng
  'chkProtectedShots
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Cheats", "chkProtectedShots", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "0" Then
  '  frmHardcoreCheats.chkProtectedShots.Value = 0
  'Else
  '  frmHardcoreCheats.chkProtectedShots.Value = 1
  'End If
  
  'custom ng
  'chkGmMessagesPauseAll
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Cheats", "chkGmMessagesPauseAll", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "0" Then
  '  frmHardcoreCheats.chkGmMessagesPauseAll.Value = 0
  'Else
  '  frmHardcoreCheats.chkGmMessagesPauseAll.Value = 1
  'End If

  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "txtExuraVita", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtExuraVita.Text = strInfo
  Else
    frmHardcoreCheats.txtExuraVita.Text = "exura vita"
  End If
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Cheats", "txtExuraVitaMana", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtExuraVitaMana.Text = strInfo
  Else
    frmHardcoreCheats.txtExuraVitaMana.Text = "160"
  End If
  

  
  'BlueAuraDelay
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "BlueAuraDelay", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    BlueAuraDelay = lonInfo
    frmHardcoreCheats.txtBlueauraDelay.Text = CStr(BlueAuraDelay)
  Else
    BlueAuraDelay = 300
    frmHardcoreCheats.txtBlueauraDelay.Text = CStr(BlueAuraDelay)
  End If
  
  
  
  
  
  
  
   strInfo = String$(10, 0)
  i = getBlackdINI("Magebomb", "ConnectEventTIMEOUT_ms", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    FIRSTCONNECTIONTIMEOUT_ms = lonInfo

  Else
    FIRSTCONNECTIONTIMEOUT_ms = 5000

  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Magebomb", "ReceiveServerAnswerTIMEOUT_ms", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    SECONDCONNECTIONTIMEOUT_ms = lonInfo
  Else
    SECONDCONNECTIONTIMEOUT_ms = 10000
  End If

  

  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "RelogBackpacks", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHardcoreCheats.txtRelogBackpacks.Text = lonInfo
  Else
    frmHardcoreCheats.txtRelogBackpacks.Text = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LowHPforAutoHeal", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    ChangeGLOBAL_RUNEHEAL_HP lonInfo
  Else
    ChangeGLOBAL_RUNEHEAL_HP 63
  End If
  
  
  GLOBAL_FRIENDSLOWLIMIT_HP = 63
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "FRIENDSLOWLIMIT_HP", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    ChangeGLOBAL_FRIENDSLOWLIMIT_HP lonInfo
  Else
    ChangeGLOBAL_FRIENDSLOWLIMIT_HP 63
  End If
  
  GLOBAL_MYSAFELIMIT_HP = 80
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "MYSAFELIMIT_HP", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    ChangeGLOBAL_MYSAFELIMIT_HP lonInfo
  Else
    ChangeGLOBAL_MYSAFELIMIT_HP 80
  End If

  'wargroups\autoheal.txt
    strInfo = String$(250, 0)
  i = getBlackdINI("Warbot", "AutoHealFileName", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmWarbot.txtFileName.Text = strInfo
  Else
    frmWarbot.txtFileName.Text = "wargroups\autoheal.txt"
  End If
  frmWarbot.ReloadAutohealFile
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "AutoHealDelay", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmWarbot.txtAutohealDelay.Text = strInfo
    frmWarbot.timerFriendHealer.Interval = CLng(strInfo)
  Else
    frmWarbot.txtAutohealDelay.Text = "300"
    frmWarbot.timerFriendHealer.Interval = 300
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "AutoHealDelay2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmWarbot.txtAutohealDelay2.Text = strInfo

  Else
    frmWarbot.txtAutohealDelay2.Text = "700"

  End If
  
  ' GLOBAL_AUTOFRIENDHEAL_MODE
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "AUTOFRIENDHEAL_MODE", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "2" Then
    frmWarbot.AutoHealOption1.value = False
    frmWarbot.AutoHealOption3.value = False
    frmWarbot.AutoHealOption2.value = True
    GLOBAL_AUTOFRIENDHEAL_MODE = 2
  ElseIf strInfo = "3" Then
    frmWarbot.AutoHealOption1.value = False
    frmWarbot.AutoHealOption2.value = False
    frmWarbot.AutoHealOption3.value = True
    GLOBAL_AUTOFRIENDHEAL_MODE = 3
  Else
    frmWarbot.AutoHealOption2.value = False
    frmWarbot.AutoHealOption3.value = False
    frmWarbot.AutoHealOption1.value = True
    GLOBAL_AUTOFRIENDHEAL_MODE = 1
  End If
  
  
  'chkAutoHealFriendEnabled
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "AutoHealFriendEnabled", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmWarbot.chkAutoHealFriendEnabled.value = 1
  Else
    frmWarbot.chkAutoHealFriendEnabled.value = 0
  End If
  
  'chkRecordLogins
  strInfo = String$(10, 0)
  i = getBlackdINI("Warbot", "RecordLogins", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmMagebomb.chkRecordLogins.value = 1
  Else
    frmMagebomb.chkRecordLogins.value = 0
  End If
  
  'UHretry
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "UHretry", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHardcoreCheats.timerSpam.Interval = lonInfo
  Else
    frmHardcoreCheats.timerSpam.Interval = 100
  End If
  
  

  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LowHPforAutoVita", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHardcoreCheats.scrollHP2.value = lonInfo
    frmHardcoreCheats.scrollHP2_Change
  Else
    frmHardcoreCheats.scrollHP2.value = 70
    frmHardcoreCheats.scrollHP2_Change
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LowUHforAlarm", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmHardcoreCheats.txtAlarmUHs = strInfo
  Else
    frmHardcoreCheats.txtAlarmUHs = 5
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "OrderType", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    frmHardcoreCheats.cmbOrderType.ListIndex = lonInfo
    frmHardcoreCheats.cmbOrderType.Text = frmHardcoreCheats.cmbOrderType.List(lonInfo)
  Else
    frmHardcoreCheats.cmbOrderType.ListIndex = 5
    frmHardcoreCheats.cmbOrderType.Text = frmHardcoreCheats.cmbOrderType.List(5)
  End If

  strInfo = String$(10, 0)
  i = getBlackdINI("Tools", "InspectTileIDs", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmCheats.chkInspectTileID.value = 1
  Else
    frmCheats.chkInspectTileID.value = 0
  End If
  
    For ibucle = 1 To MAXCLIENTS
        AllowRepositionAtStart(ibucle) = 1
    Next ibucle

  
    For ibucle = 1 To MAXCLIENTS
        AllowRepositionAtTrap(ibucle) = 1
    Next ibucle

  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "LootProtection", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmCavebot.chkLootProtection.value = 0
  Else
    frmCavebot.chkLootProtection.value = 1
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cavebot", "BlockOption", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "2" Then
    frmCavebot.Option1.value = False
    frmCavebot.Option2.value = True
  Else
    frmCavebot.Option1.value = True
    frmCavebot.Option2.value = False
  End If
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Cavebot", "MAX_LOCKWAIT", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     frmCavebot.txtBlockSec.Text = strInfo
     MAX_LOCKWAIT = CLng(strInfo)
  Else
     frmCavebot.txtBlockSec.Text = "30000"
     MAX_LOCKWAIT = 30000
  End If
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Cavebot", "EXORIVIS_COST", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     EXORIVIS_COST = CLng(strInfo)
  Else
     EXORIVIS_COST = 20
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("Cavebot", "EXORIVIS_SPELL", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     EXORIVIS_SPELL = strInfo
  Else
     EXORIVIS_SPELL = "exori vis"
  End If
  
  
  
  
  
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Cavebot", "EXORIMORT_COST", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     EXORIMORT_COST = CLng(strInfo)
  Else
     EXORIMORT_COST = 20
  End If
  
  strInfo = String$(255, 0)
  i = getBlackdINI("Cavebot", "EXORIMORT_SPELL", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     EXORIMORT_SPELL = strInfo
  Else
     EXORIMORT_SPELL = "exori mort"
  End If
  
  
  
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "LockOnMyFloor", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkLockOnMyFloor.value = 0
  Else
    frmHardcoreCheats.chkLockOnMyFloor.value = 1
  End If


  strInfo = String$(10, 0)
  i = getBlackdINI("Runemaker", "RunemakerChaos", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    RunemakerChaos = lonInfo
  Else
    RunemakerChaos = 600
  End If
  
    strInfo = String$(10, 0)
  i = getBlackdINI("Runemaker", "RunemakerChaos2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    RunemakerChaos2 = lonInfo
  Else
    RunemakerChaos2 = 10000
  End If

frmRunemaker.txrRunemakerChaos.Text = CStr(RunemakerChaos)
frmRunemaker.txrRunemakerChaos2.Text = CStr(RunemakerChaos2)

'custom ng dont load with this
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Runemaker", "ChkDangerSound", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "0" Then
  '  frmRunemaker.ChkDangerSound.Value = 0
  'Else
  '  frmRunemaker.ChkDangerSound.Value = 1
  'End If

'custom ng dont load with this
  'strInfo = String$(10, 0)
  'i = getBlackdINI("Runemaker", "ChkCloseSound", "", strInfo, Len(strInfo), here)
  'strInfo = Left(strInfo, i)
  'If strInfo = "0" Then
  '  frmRunemaker.chkCloseSound.Value = 0
  'Else
  '  frmRunemaker.chkCloseSound.Value = 1
  'End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Runemaker", "ChkOnDangerSS2", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmRunemaker.chkOnDangerSS.value = 1
  Else
    frmRunemaker.chkOnDangerSS.value = 0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "ForceLoginServer", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmMain.chkForceLoginServer.value = 1
  Else
    frmMain.chkForceLoginServer.value = 0
  End If
  
  strInfo = String$(10, 0)
  i = getBlackdINI("AdvancedProxyOptions", "WantBypass", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    bypass_def1 = 1
  Else
    bypass_def1 = 0
  End If
  frmAdvanced.chkWantBypass.value = bypass_def1
  
  strInfo = String$(250, 0)
  i = getBlackdINI("AdvancedProxyOptions", "BypassLoginCharacter", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    bypass_def2 = strInfo
  Else
    bypass_def2 = ""
  End If
  frmAdvanced.txtLoginCharacter.Text = bypass_def2
  
  strInfo = String$(250, 0)
  i = getBlackdINI("AdvancedProxyOptions", "BypassGameserver", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    bypass_def3 = strInfo
  Else
    bypass_def3 = ""
  End If
  frmAdvanced.cmbTibiaServers.Text = bypass_def3
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "MapOnTop", "", strInfo, Len(strInfo), here)
  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkOnTop.value = 0
    ToggleTopmost frmTrueMap.hwnd, False
    ToggleTopmost frmMapReader.hwnd, False
    MapWantedOnTop = False
  Else
    frmHardcoreCheats.chkOnTop.value = 1
    ToggleTopmost frmTrueMap.hwnd, True
    ToggleTopmost frmMapReader.hwnd, True
    MapWantedOnTop = True
  End If
  
  strInfo = String$(50, 0)
  i = getBlackdINI("Cheats", "MapUpdateIntervalInMs", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
     frmHardcoreCheats.cmdMs.Text = strInfo
     frmHardcoreCheats.cmdMs_Change
  Else
     frmHardcoreCheats.cmdMs.Text = "1000"
     frmHardcoreCheats.cmdMs_Change
  End If
  
  
  'TimerConditionTick
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "TimerConditionTick", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TimerConditionTick = lonInfo
  Else
    TimerConditionTick = 300
  End If
  frmCondEvents.timerCheck.Interval = TimerConditionTick
  frmCondEvents.txtMs.Text = CStr(TimerConditionTick)
  
  
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "TimerConditionTick2", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(strInfo)
    TimerConditionTick2 = lonInfo
  Else
    TimerConditionTick2 = 700
  End If
  frmCondEvents.txtMs2.Text = CStr(TimerConditionTick2)
  
  'custom ng save old menu
  strInfo = String$(10, 0)
  i = getBlackdINI("OldMenu", "chkOldMenu", "", strInfo, Len(strInfo), here)
  
  strInfo = Left(strInfo, i)
  If strInfo = "1" Then
    frmOld.chkOldMenu.value = 1
    'frmOld.chkOldMenu_Click
  Else
    frmOld.chkOldMenu.value = 0
  End If
  
  
  ' CheatsEnabled and Version SHOULD BE LAST THINGS TO READ
  strInfo = String$(10, 0)
  i = getBlackdINI("Cheats", "CheatsEnabled", "", strInfo, Len(strInfo), here)

  strInfo = Left(strInfo, i)
  If strInfo = "0" Then
    frmHardcoreCheats.chkApplyCheats.value = 0
    frmHardcoreCheats.chkApplyCheats_Click
  Else
    frmHardcoreCheats.chkApplyCheats.value = 1
  End If
  strInfo = String$(50, 0)
  i = getBlackdINI("Proxy", "Version", "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    frmMain.Caption = strInfo
    If TrialVersion = True Then
      frmMain.Caption = "Blackd Proxy " & ProxyVersion & " TRIAL"
    End If
  Else
   frmMain.Caption = "Blackd Proxy " & ProxyVersion
    If TrialVersion = True Then
      frmMain.Caption = "Blackd Proxy " & ProxyVersion & " TRIAL"
    End If
  End If

   frmMain.Caption = frmMain.Caption & " for Tibia " & TibiaVersion

  Exit Sub
gotErr:
  MsgBox "Sorry, Blackd Proxy was not able to read ini files (end)" & vbCrLf & "Possible reasons:" & vbCrLf & _
  " - Corrupted config.ini" & vbCrLf & _
  " - Corrupted settings.ini" & vbCrLf & _
  " Details:" & vbCrLf & _
  " - Error number " & Err.Number & vbCrLf & _
  " - Error description: " & Err.Description, vbOKOnly + vbCritical, "Critical error"
  End
End Sub

Public Sub WriteIni()
  ' write ini file
  Dim i As Integer
  Dim strInfo As String
  Dim here As String
  Dim idLoginSP As Long
  userHere = App.Path & "\settings.ini"
  
  If configPath = "" Then
    here = myMainConfigINIPath()
  Else
    here = App.Path & "\" & configPath & "\config.ini"
  End If
  
  strInfo = CStr(configPath)
  i = setBlackdINI("Proxy", "configPath", strInfo, myMainConfigINIPath())

  strInfo = CStr(MAXCLIENTS)
  i = setBlackdINI("Proxy", "MAXCLIENTS", strInfo, here)
  
  strInfo = CStr(MAXEVENTS)
  i = setBlackdINI("Proxy", "MAXEVENTS", strInfo, here)
    
  strInfo = CStr(MAXCONDS)
  i = setBlackdINI("Proxy", "MAXCONDS", strInfo, here)
    
  
  
  strInfo = CStr(HIGHEST_BP_ID)
  i = setBlackdINI("MemoryAddresses", "HIGHEST_BP_ID", strInfo, here)
  strInfo = CStr(MAXDATTILES)
  i = setBlackdINI("MemoryAddresses", "MAXDATTILES", strInfo, here)
  strInfo = CStr(MAXTILEIDLISTSIZE)
  i = setBlackdINI("Proxy", "MAXTILEIDLISTSIZE", strInfo, here)
  strInfo = ForwardGameTo.Text
  i = setBlackdINI("Proxy", "ForwardGameTo", strInfo, here)
  strInfo = txtServerLoginP.Text
  i = setBlackdINI("Proxy", "txtServerLoginP", strInfo, here)
  strInfo = txtServerGameP.Text
  i = setBlackdINI("Proxy", "txtServerGameP", strInfo, here)
  strInfo = "FALSE"
  i = setBlackdINI("Proxy", "FirstExecute", strInfo, here)
  
  strInfo = CStr(lngNextScreenshotNumber)
  i = setBlackdINI("Proxy", "NextScreenshotNumber", strInfo, here)
  
  
  strInfo = TibiaPath
  i = setBlackdINI("Proxy", "TibiaPath", strInfo, here)
  
  strInfo = gmStart & "-"
  i = setBlackdINI("Proxy", "GmStartWithThis3", strInfo, here)
  
  strInfo = gmStart2 & "-"
  i = setBlackdINI("Proxy", "AltGmStartWithThis3", strInfo, here)
  
  If TrueServer1.value = True Then
    strInfo = "1"
  ElseIf TrueServer2.value = True Then
    strInfo = "2"
  Else
    strInfo = "3"
  End If
  i = setBlackdINI("Proxy", "ForwardOption", strInfo, here)
  
  If frmHardcoreCheats.ActionPath.value = True Then
    strInfo = "4"
  ElseIf frmHardcoreCheats.ActionInspect.value = True Then
    strInfo = "1"
  ElseIf frmHardcoreCheats.ActionMove.value = True Then
    strInfo = "2"
  Else
    strInfo = "3"
  End If
  i = setBlackdINI("Cheats", "MapClickOption", strInfo, here)
  
  strInfo = serverLogoutMessage
  i = setBlackdINI("MemoryAddresses", "serverLogoutMessage", strInfo, here)
  
  ' blnShowAdvancedOptions
  If blnShowAdvancedOptions = True Then
    strInfo = "1"
  Else
    strInfo = "0"
  End If
  
  i = setBlackdINI("Proxy", "ShowAdvancedOptions", strInfo, here)
  
    
  
  
  strInfo = CStr(chckMemoryIP.value)
  i = setBlackdINI("AdvancedProxyOptions", "LocalLoginUseProxy", strInfo, here)
  
   strInfo = CStr(chckAlter.value)
  i = setBlackdINI("AdvancedProxyOptions", "LocalGameUseProxy", strInfo, here)
  
  ' txtClientLoginP
  strInfo = txtClientLoginP.Text
  i = setBlackdINI("AdvancedProxyOptions", "ListenLoginPort", strInfo, here)
  ' txtClientGameP
  strInfo = txtClientGameP.Text
  i = setBlackdINI("AdvancedProxyOptions", "ListenGamePort", strInfo, here)
  
  strInfo = CStr(frmMain.chkForceLoginServer.value)
  i = setBlackdINI("AdvancedProxyOptions", "ForceLoginServer", strInfo, here)
  strInfo = CStr(frmAdvanced.chkWantBypass.value)
  i = setBlackdINI("AdvancedProxyOptions", "WantBypass", strInfo, here)
  strInfo = frmAdvanced.txtLoginCharacter.Text
  i = setBlackdINI("AdvancedProxyOptions", "BypassLoginCharacter", strInfo, here)
  strInfo = frmAdvanced.cmbTibiaServers.Text
  i = setBlackdINI("AdvancedProxyOptions", "BypassGameserver", strInfo, here)
  
  ' txtMaxChar
  strInfo = txtMaxChar.Text
  i = setBlackdINI("Log", "MaxLogBuffer", strInfo, here)
  ' txtMaxLines
  strInfo = txtMaxLines.Text
  i = setBlackdINI("Log", "MaxHexLines", strInfo, here)
  
  If LogFull1.value = True Then
    strInfo = "1"
  ElseIf LogFull2.value = True Then
    strInfo = "2"
  Else
    strInfo = "3"
  End If
  i = setBlackdINI("Log", "LogFullAction", strInfo, here)
  ' txtLogFile
  strInfo = txtLogFile.Text
  i = setBlackdINI("Log", "LogFile", strInfo, here)
  

  strInfo = tibiaclassname
  i = setBlackdINI("MemoryAddresses", "tibiaclassname", strInfo, here)
  
  If TibiaVersionLong < 1011 Then
  For idLoginSP = 1 To NumberOfLoginServers
    strInfo = "&H" & Hex(memLoginServer(idLoginSP))
    i = setBlackdINI("MemoryAddresses", "MemLoginServer" & CStr(idLoginSP), strInfo, here)
    
    strInfo = "&H" & Hex(MemPortLoginServer(idLoginSP))
    i = setBlackdINI("MemoryAddresses", "MemPortLoginServer" & CStr(idLoginSP), strInfo, here)
  Next idLoginSP
 

  Else
    strInfo = "&H" & Hex(LoginServerStartPointer)
    i = setBlackdINI("MemoryAddresses", "LoginServerStartPointer", strInfo, here)
    strInfo = "&H" & Hex(LoginServerStep)
    i = setBlackdINI("MemoryAddresses", "LoginServerStep", strInfo, here)
    strInfo = "&H" & Hex(HostnamePointerOffset)
    i = setBlackdINI("MemoryAddresses", "HostnamePointerOffset", strInfo, here)
    strInfo = "&H" & Hex(IPAddressPointerOffset)
    i = setBlackdINI("MemoryAddresses", "IPAddressPointerOffset", strInfo, here)
    strInfo = "&H" & Hex(PortOffset)
    i = setBlackdINI("MemoryAddresses", "PortOffset", strInfo, here)
    strInfo = "&H" & Hex(proxyChecker)
    i = setBlackdINI("MemoryAddresses", "proxyChecker", strInfo, here)
    
    If TibiaVersionLong = 1011 Then
    strInfo = tibiaEntryServer
    i = setBlackdINI("MemoryAddresses", "tibiaEntryServer", strInfo, here)
    End If
  End If


    strInfo = "&H" & Hex(LEVELSPY_NOP)
    i = setBlackdINI("MemoryAddresses", "LEVELSPY_NOP", strInfo, here)
    strInfo = "&H" & Hex(LEVELSPY_ABOVE)
    i = setBlackdINI("MemoryAddresses", "LEVELSPY_ABOVE", strInfo, here)
    strInfo = "&H" & Hex(LEVELSPY_BELOW)
    i = setBlackdINI("MemoryAddresses", "LEVELSPY_BELOW", strInfo, here)
    
    strInfo = "&H" & Hex(LIGHT_NOP)
    i = setBlackdINI("MemoryAddresses", "LIGHT_NOP", strInfo, here)
    strInfo = "&H" & Hex(LIGHT_AMOUNT)
    i = setBlackdINI("MemoryAddresses", "LIGHT_AMOUNT", strInfo, here)
    
    strInfo = "&H" & Hex(PLAYER_Z)
    i = setBlackdINI("MemoryAddresses", "PLAYER_Z", strInfo, here)


    strInfo = "&H" & Hex(RedSquare)
    i = setBlackdINI("MemoryAddresses", "RedSquare", strInfo, here)



  strInfo = "&H" & Hex(adrRSA)
  i = setBlackdINI("MemoryAddresses", "adrRSA", strInfo, here)
  

  strInfo = "&H" & Hex(adrMulticlient)
  i = setBlackdINI("MemoryAddresses", "adrMulticlient", strInfo, here)
  strInfo = "&H" & Hex(multiclientByte1)
  i = setBlackdINI("MemoryAddresses", "multiclientByte1", strInfo, here)
  strInfo = "&H" & Hex(multiclientByte2)
  i = setBlackdINI("MemoryAddresses", "multiclientByte2", strInfo, here)



  strInfo = "&H" & Hex(adrXgo)
  i = setBlackdINI("MemoryAddresses", "adrXgo", strInfo, here)
  strInfo = "&H" & Hex(adrYgo)
  i = setBlackdINI("MemoryAddresses", "adrYgo", strInfo, here)
  strInfo = "&H" & Hex(adrZgo)
  i = setBlackdINI("MemoryAddresses", "adrZgo", strInfo, here)
  strInfo = "&H" & Hex(adrGo)
  i = setBlackdINI("MemoryAddresses", "adrGo", strInfo, here)
  strInfo = "&H" & Hex(adrOutfit)
  i = setBlackdINI("MemoryAddresses", "adrOutfit", strInfo, here)
  
  'strInfo = "&H" & Hex(adrConnectionKey)
  'i = setBlackdINI("MemoryAddresses", "adrConnectionKey", strInfo, here)
  
  'strInfo = "&H" & Hex(adrSelectedCharIndex)
  'i = setBlackdINI("MemoryAddresses", "adrSelectedCharIndex", strInfo, here)
  
'  strInfo = "&H" & Hex(adrAccount)
'  i = setBlackdINI("MemoryAddresses", "adrAccount", strInfo, here)
  
  strInfo = "&H" & Hex(adrLastPacket)
  i = setBlackdINI("MemoryAddresses", "adrLastPacket", strInfo, here)
  
  strInfo = "&H" & Hex(adrCharListPtr)
  i = setBlackdINI("MemoryAddresses", "adrCharListPtr", strInfo, here)
  
  strInfo = "&H" & Hex(adrCharListPtrEND)
  i = setBlackdINI("MemoryAddresses", "adrCharListPtrEND", strInfo, here)
  
  
  strInfo = "&H" & Hex(adrNChar)
  i = setBlackdINI("MemoryAddresses", "adrNChar", strInfo, here)
  
  strInfo = "&H" & Hex(CharDist)
  i = setBlackdINI("MemoryAddresses", "CharDist", strInfo, here)
  
  strInfo = "&H" & Hex(NameDist)
  i = setBlackdINI("MemoryAddresses", "NameDist", strInfo, here)
  
  strInfo = "&H" & Hex(OutfitDist)
  i = setBlackdINI("MemoryAddresses", "OutfitDist", strInfo, here)
  
  strInfo = "&H" & Hex(SpeedDist)
  i = setBlackdINI("MemoryAddresses", "SpeedDist", strInfo, here)
  
  strInfo = "&H" & Hex(adrNum)
  i = setBlackdINI("MemoryAddresses", "adrNum", strInfo, here)
  
  strInfo = "&H" & Hex(adrConnected)
  i = setBlackdINI("MemoryAddresses", "adrConnected", strInfo, here)
  
  strInfo = CStr(maxStepsPerMovement)
  i = setBlackdINI("MemoryAddresses", "maxStepsPerMovement", strInfo, here)
  
  strInfo = "&H" & Hex(adrPointerToInternalFPSminusH5D)
  i = setBlackdINI("MemoryAddresses", "adrPointerToInternalFPSminusH5D", strInfo, here)
   
  strInfo = "&H" & Hex(adrInternalFPS)
  i = setBlackdINI("MemoryAddresses", "adrInternalFPS", strInfo, here)
  
  'frmHotkeys.chkHotkeysActivated.Value
  strInfo = CStr(frmHotkeys.chkHotkeysActivated.value)
  i = setBlackdINI("Hotkeys", "HotkeysActivated", strInfo, here)
  'frmHotkeys.chkRepeat.Value
  strInfo = CStr(frmHotkeys.chkRepeat.value)
  i = setBlackdINI("Hotkeys", "RepeatEnabled", strInfo, here)
  'frmHotkeys.txtDelay.Text
  strInfo = frmHotkeys.txtDelay.Text
  i = setBlackdINI("Hotkeys", "RepeatDelay", strInfo, here)
  
  
  
  
  'chkSelect
  strInfo = CStr(chkSelect.value)
  i = setBlackdINI("Log", "AutoSelectHexAscii", strInfo, here)
  'chkAutoHide
  strInfo = CStr(chkAutoHide.value)
  i = setBlackdINI("Log", "HideHexLogIfLogDisabled", strInfo, here)
  
 'chkLogoutIfDanger
  strInfo = CStr(frmHardcoreCheats.chkLogoutIfDanger.value)
  i = setBlackdINI("Cheats", "LogoutIfDangerAtStart", strInfo, here)
  
  
  

  strInfo = CStr(frmStealth.chkStealthMessages.value)
  i = setBlackdINI("Cheats", "chkStealthMessages", strInfo, here)
  
  strInfo = CStr(frmStealth.chkStealthCommands.value)
  i = setBlackdINI("Cheats", "chkStealthCommands", strInfo, here)
  
  strInfo = CStr(frmStealth.chkStealthExp.value)
  i = setBlackdINI("Cheats", "chkStealthExp", strInfo, here)
  
  strInfo = CStr(frmStealth.chkAvoidChat.value)
  i = setBlackdINI("Cheats", "chkAvoidChat", strInfo, here)
  
  'chkReveal
  strInfo = CStr(frmHardcoreCheats.chkReveal.value)
  i = setBlackdINI("Cheats", "RevealInvis", strInfo, here)
  
  'chkLight
  strInfo = CStr(frmHardcoreCheats.chkLight.value)
  i = setBlackdINI("Cheats", "ChangeLight", strInfo, here)
   
  'chkAutoHeal
  strInfo = CStr(frmHardcoreCheats.chkAutoHeal.value)
  i = setBlackdINI("Cheats", "AutoHealEnabled", strInfo, here)

  strInfo = CStr(frmHardcoreCheats.chkRuneAlarm.value)
  i = setBlackdINI("Cheats", "UHalarmEnabled", strInfo, here)
  

  'chkCaptionExp
  strInfo = CStr(frmHardcoreCheats.chkCaptionExp.value)
  i = setBlackdINI("Cheats", "TitleExp", strInfo, here)

  strInfo = CStr(frmHardcoreCheats.chkAutoGratz.value)
  i = setBlackdINI("Cheats", "chkAutoGratz", strInfo, here)
  
  strInfo = CStr(frmHardcoreCheats.chkAutorelog.value)
  i = setBlackdINI("Cheats", "chkAutorelog", strInfo, here)
  
  strInfo = CStr(Antibanmode)
  i = setBlackdINI("Cheats", "Antibanmode", strInfo, here)
  
  strInfo = CStr(frmHardcoreCheats.chkProtectedShots.value)
  i = setBlackdINI("Cheats", "chkProtectedShots", strInfo, here)
  
  strInfo = CStr(frmHardcoreCheats.chkGmMessagesPauseAll.value)
  i = setBlackdINI("Cheats", "chkGmMessagesPauseAll", strInfo, here)
  
  strInfo = frmHardcoreCheats.txtExuraVita.Text
  i = setBlackdINI("Cheats", "txtExuraVita", strInfo, here)
  
  strInfo = frmHardcoreCheats.txtExuraVitaMana.Text
  i = setBlackdINI("Cheats", "txtExuraVitaMana", strInfo, here)
  
  strInfo = CStr(frmCavebot.chkChangePkHeal.value)
  i = setBlackdINI("Cavebot", "AutoChangePKHeal", strInfo, here)


  strInfo = CStr(frmCavebot.scrollPkHeal.value)
  i = setBlackdINI("Cavebot", "NewAutohealForPKattack", strInfo, here) ' fixed in 9.42
  'SafeHPforExoriVis
  strInfo = CStr(frmCavebot.scrollExorivis.value)
  i = setBlackdINI("Cavebot", "SafeHPforExoriVis", strInfo, here)
  
  strInfo = CStr(SetVeryFriendly_NOATTACKTIMER_ms)
  i = setBlackdINI("Cavebot", "SetVeryFriendly_NOATTACKTIMER_ms", strInfo, here)
   
  'chkAutoVita
  strInfo = CStr(frmHardcoreCheats.chkAutoVita.value)
  i = setBlackdINI("Cheats", "AutoVitaEnabled", strInfo, here)
  
  'chkAcceptSDorder
  strInfo = CStr(frmHardcoreCheats.chkAcceptSDorder.value)
  i = setBlackdINI("Cheats", "AcceptSDorder", strInfo, here)
  
  
  
  strInfo = CStr(frmBroadcast.chkMC.value)
  i = setBlackdINI("Broadcast", "BroadcastMC", strInfo, here)
  
  
  strInfo = CStr(frmBroadcast.txtBroadcastDelay1)
  i = setBlackdINI("Broadcast", "BroadcastDelay1", strInfo, here)
  
  strInfo = CStr(frmBroadcast.txtBroadcastDelay2)
  i = setBlackdINI("Broadcast", "BroadcastDelay2", strInfo, here)
   
  'txtOrder
  strInfo = frmHardcoreCheats.txtOrder
  i = setBlackdINI("Cheats", "SDorder", strInfo, here)


  strInfo = CStr(LimitRandomizator)
  i = setBlackdINI("HPmana", "LimitRandomizator", strInfo, here)
  
  strInfo = CStr(HPmanaRECAST)
  i = setBlackdINI("HPmana", "HPmanaRECAST", strInfo, here)

  strInfo = CStr(HPmanaRECAST2)
  i = setBlackdINI("HPmana", "HPmanaRECAST2", strInfo, here)
  
  'txtRemoteLeader
  strInfo = frmHardcoreCheats.txtRemoteLeader
  i = setBlackdINI("Cheats", "SDleader", strInfo, here)
  
  'ExivaExpPlace
  strInfo = frmHardcoreCheats.cmbWhere.Text
  i = setBlackdINI("Cheats", "ExivaExpPlace", strInfo, here)
  
  
  'txtExivaExpFormat
  strInfo = frmHardcoreCheats.txtExivaExpFormat.Text
  i = setBlackdINI("Cheats", "txtExivaExpFormat", strInfo, here)
  
  'tibiaTittleFormat
  strInfo = frmHardcoreCheats.tibiaTittleFormat.Text
  i = setBlackdINI("Cheats", "tibiaTittleFormat", strInfo, here)
   
 
  'chkColorEffects
  strInfo = CStr(frmHardcoreCheats.chkColorEffects.value)
  i = setBlackdINI("Cheats", "ColorEffects", strInfo, here)
  
  'scrollLight
  strInfo = CStr(frmHardcoreCheats.scrollLight.value)
  i = setBlackdINI("Cheats", "LightPower", strInfo, here)
  
  
  'scrollHP
  strInfo = CStr(BlueAuraDelay)
  i = setBlackdINI("Cheats", "BlueAuraDelay", strInfo, here)
  
  
  strInfo = CStr(frmHardcoreCheats.txtRelogBackpacks.Text)
  i = setBlackdINI("Cheats", "RelogBackpacks", strInfo, here)
  
  'scrollHP
  strInfo = CStr(GLOBAL_RUNEHEAL_HP)
  i = setBlackdINI("Cheats", "LowHPforAutoHeal", strInfo, here)


  'scrollHP2
   strInfo = CStr(frmHardcoreCheats.scrollHP2.value)
  i = setBlackdINI("Cheats", "LowHPforAutoVita", strInfo, here)

  'warbot
  strInfo = CStr(GLOBAL_FRIENDSLOWLIMIT_HP)
  i = setBlackdINI("Warbot", "FRIENDSLOWLIMIT_HP", strInfo, here)
  
  strInfo = CStr(GLOBAL_MYSAFELIMIT_HP)
  i = setBlackdINI("Warbot", "MYSAFELIMIT_HP", strInfo, here)

 'AutoHealFileName
  strInfo = CStr(frmWarbot.txtFileName.Text)
  i = setBlackdINI("Warbot", "AutoHealFileName", strInfo, here)
  
  'txtAutohealDelay
  strInfo = CStr(frmWarbot.txtAutohealDelay.Text)
  i = setBlackdINI("Warbot", "AutoHealDelay", strInfo, here)
  
  'txtAutohealDelay2
  strInfo = CStr(frmWarbot.txtAutohealDelay2.Text)
  i = setBlackdINI("Warbot", "AutoHealDelay2", strInfo, here)
  
  'AUTOFRIENDHEAL_MODE
  strInfo = CStr(GLOBAL_AUTOFRIENDHEAL_MODE)
  i = setBlackdINI("Warbot", "AUTOFRIENDHEAL_MODE", strInfo, here)
  
  'AutoHealFriendEnabled
  strInfo = CStr(frmWarbot.chkAutoHealFriendEnabled.value)
  i = setBlackdINI("Warbot", "AutoHealFriendEnabled", strInfo, here)
  
  'chkRecordLogins
  strInfo = CStr(frmMagebomb.chkRecordLogins.value)
  i = setBlackdINI("Warbot", "RecordLogins", strInfo, here)
  
  'txtAlarmUHs
   strInfo = CStr(frmHardcoreCheats.txtAlarmUHs)
  i = setBlackdINI("Cheats", "LowUHforAlarm", strInfo, here)
  
  'UHretry
  strInfo = CStr(frmHardcoreCheats.timerSpam.Interval)
  i = setBlackdINI("Cheats", "UHretry", strInfo, here)
   
  'chkLockOnMyFloor
   strInfo = CStr(frmHardcoreCheats.chkLockOnMyFloor.value)
  i = setBlackdINI("Cheats", "LockOnMyFloor", strInfo, here)
  
  'chkOnTop
  strInfo = CStr(frmHardcoreCheats.chkOnTop.value)
  i = setBlackdINI("Cheats", "MapOnTop", strInfo, here)
  
  
  'cmdMs
  strInfo = frmHardcoreCheats.cmdMs.Text
  i = setBlackdINI("Cheats", "MapUpdateIntervalInMs", strInfo, here)
  
  

  strInfo = CStr(RunemakerChaos)
  i = setBlackdINI("Runemaker", "RunemakerChaos", strInfo, here)
  
  strInfo = CStr(RunemakerChaos2)
  i = setBlackdINI("Runemaker", "RunemakerChaos2", strInfo, here)
  
  'ChkDangerSound
  strInfo = CStr(frmRunemaker.ChkDangerSound.value)
  i = setBlackdINI("Runemaker", "ChkDangerSound", strInfo, here)
  
  'ChkCloseSound
  strInfo = CStr(frmRunemaker.chkCloseSound.value)
  i = setBlackdINI("Runemaker", "ChkCloseSound", strInfo, here)
  
  'ChkCloseSound
  strInfo = CStr(frmRunemaker.chkOnDangerSS.value)
  i = setBlackdINI("Runemaker", "ChkOnDangerSS2", strInfo, here)
  
  'chkApplyCheats
  strInfo = CStr(frmHardcoreCheats.chkApplyCheats.value)
  i = setBlackdINI("Cheats", "CheatsEnabled", strInfo, here)
  
  'custom ng save old menu
  strInfo = CStr(frmOld.chkOldMenu.value)
  i = setBlackdINI("OldMenu", "chkOldMenu", strInfo, here)
  
  
    'cmbOrderType
    
  strInfo = CStr(frmHardcoreCheats.cmbOrderType.ListIndex)
  i = setBlackdINI("Cheats", "OrderType", strInfo, here)
  

  
  strInfo = CStr(TimerConditionTick)
  i = setBlackdINI("Cheats", "TimerConditionTick", strInfo, here)
  
  strInfo = CStr(TimerConditionTick2)
  i = setBlackdINI("Cheats", "TimerConditionTick2", strInfo, here)
  
  'chkInspectTileID
  strInfo = CStr(frmCheats.chkInspectTileID.value)
  i = setBlackdINI("Tools", "InspectTileIDs", strInfo, here)
  
  
  

  strInfo = CStr(TrainerTimer1)
  i = setBlackdINI("Cavebot", "TrainerTimer1", strInfo, here)
  
  strInfo = CStr(TrainerTimer2)
  i = setBlackdINI("Cavebot", "TrainerTimer2", strInfo, here)
  
  
  
  strInfo = CStr(CavebotRECAST)
  i = setBlackdINI("Cavebot", "TimerInterval", strInfo, here)
  
  strInfo = CStr(CavebotRECAST2)
  i = setBlackdINI("Cavebot", "TimerInterval2", strInfo, here)
  
  strInfo = CStr(CteMoveDelay)
  i = setBlackdINI("Cavebot", "CteMoveDelay", strInfo, here)
  
  
  'strInfo = CStr(frmCavebot.chkAllowRepositionAtStart.Value)
  'i = setBlackdINI("Cavebot", "AllowRepositionAtStart", strInfo, here)
  
  strInfo = CStr(frmCavebot.chkLootProtection.value)
  i = setBlackdINI("Cavebot", "LootProtection", strInfo, here)
    
  If frmCavebot.Option1.value = True Then
    strInfo = "1"
  Else
    strInfo = "2"
  End If
  i = setBlackdINI("Cavebot", "BlockOption", strInfo, here)
    
  strInfo = CStr(MAX_LOCKWAIT)
  i = setBlackdINI("Cavebot", "MAX_LOCKWAIT", strInfo, here)
  
  strInfo = CStr(EXORIVIS_COST)
  i = setBlackdINI("Cavebot", "EXORIVIS_COST", strInfo, here)
  
  strInfo = EXORIVIS_SPELL
  i = setBlackdINI("Cavebot", "EXORIVIS_SPELL", strInfo, here)
  
  
  strInfo = CStr(EXORIMORT_COST)
  i = setBlackdINI("Cavebot", "EXORIMORT_COST", strInfo, here)
  
  strInfo = EXORIMORT_SPELL
  i = setBlackdINI("Cavebot", "EXORIMORT_SPELL", strInfo, here)
  
  strInfo = CStr(TimeToGiveTrapAlarm)
  i = setBlackdINI("Cavebot", "TimeToGiveTrapAlarm", strInfo, here)
  
  
  strInfo = CStr(NumberOfLoginServers)
  i = setBlackdINI("MemoryAddresses", "NumberOfLoginServers", strInfo, here)
  
  For idLoginSP = 1 To NumberOfLoginServers
    strInfo = cmbPrefered.List(idLoginSP - 1)
    i = setBlackdINI("MemoryAddresses", "trueLoginServer" & CStr(idLoginSP), strInfo, here)
  Next idLoginSP
    

  
  strInfo = cmbPrefered.Text
  i = setBlackdINI("MemoryAddresses", "PREFEREDLOGINSERVER", strInfo, here)
              
  For idLoginSP = 1 To NumberOfLoginServers
    strInfo = trueLoginPort(idLoginSP)
    i = setBlackdINI("MemoryAddresses", "trueLoginPort" & CStr(idLoginSP), strInfo, here)
  Next idLoginSP
   
  strInfo = PREFEREDLOGINPORT
  i = setBlackdINI("MemoryAddresses", "PREFEREDLOGINPORT", strInfo, here)
  
  
  
  strInfo = TibiaExePath
  i = setBlackdINI("MemoryAddresses", "TibiaExePath", strInfo, here)
  
  'strInfo = MagebotPath
  'i = setBlackdINI("MemoryAddresses", "MagebotPath", strInfo, here)
  
  
  
  
  strInfo = CStr(frmAdvanced.chkAlternativeBinding.value)
  i = setBlackdINI("AdvancedProxyOptions", "AlternativeBinding", strInfo, here)
   
   
  strInfo = CStr(MyPriorityID)
  i = setBlackdINI("AdvancedProxyOptions", "MyPriorityID", strInfo, here)
   
  strInfo = CStr(TibiaPriorityID)
  i = setBlackdINI("AdvancedProxyOptions", "TibiaPriorityID", strInfo, here)
      
  strInfo = CStr(TOOSLOWLOGINSERVER_MS)
  i = setBlackdINI("AdvancedProxyOptions", "TOOSLOWLOGINSERVER_MS", strInfo, here)
  
  strInfo = CStr(FIRSTCONNECTIONTIMEOUT_ms)
  i = setBlackdINI("Magebomb", "ConnectEventTIMEOUT_ms", strInfo, here)
   
  strInfo = CStr(SECONDCONNECTIONTIMEOUT_ms)
  i = setBlackdINI("Magebomb", "ReceiveServerAnswerTIMEOUT_ms", strInfo, here)
  

  
  ' runes
  WriteTileIDListToIni AditionalStairsToUpFloor, "AditionalStairsToUpFloor", here
  WriteTileIDListToIni AditionalStairsToDownFloor, "AditionalStairsToDownFloor", here
  WriteTileIDListToIni AditionalRequireRope, "AditionalRequireRope", here
  WriteTileIDListToIni AditionalRequireShovel, "AditionalRequireShovel", here
      
  WriteTileIDToIni tileID_Blank, "tileID_Blank", here

  WriteTileIDToIni tileID_WallBugItem, "tileID_WallBugItem", here
  
  WriteTileIDToIni tileID_SD, "tileID_SD", here
  WriteTileIDToIni tileID_HMM, "tileID_HMM", here
  WriteTileIDToIni tileID_Explosion, "tileID_Explosion", here
  WriteTileIDToIni tileID_IH, "tileID_IH", here
  WriteTileIDToIni tileID_UH, "tileID_UH", here
  
  WriteTileIDToIni tileID_fireball, "tileID_fireball", here
  WriteTileIDToIni tileID_stalagmite, "tileID_stalagmite", here
  WriteTileIDToIni tileID_icicle, "tileID_icicle", here
  
  ' items
  WriteTileIDToIni tileID_Bag, "tileID_Bag", here
  WriteTileIDToIni tileID_Backpack, "tileID_Backpack", here
  WriteTileIDToIni tileID_Oracle, "tileID_Oracle", here
  WriteTileIDToIni tileID_FishingRod, "tileID_FishingRod", here
 
  WriteTileIDToIni tileID_Rope, "tileID_Rope", here
  WriteTileIDToIni tileID_LightRope, "tileID_LightRope", here
  WriteTileIDToIni tileID_Shovel, "tileID_Shovel", here
  WriteTileIDToIni tileID_LightShovel, "tileID_LightShovel", here
  
  ' water
  WriteTileIDToIni tileID_waterEmpty, "tileID_waterEmpty", here
  WriteTileIDToIni tileID_waterWithFish, "tileID_waterWithFish", here
 
  WriteTileIDToIni tileID_waterEmptyEnd, "tileID_waterEmptyEnd", here
  WriteTileIDToIni tileID_waterWithFishEnd, "tileID_waterWithFishEnd", here
  
  ' blocking table
  WriteTileIDToIni tileID_blockingBox, "tileID_blockingBox", here
   
  ' to UP floor
  WriteTileIDToIni tileID_stairsToUp, "tileID_stairsToUp", here
  WriteTileIDToIni tileID_woodenStairstoUp, "tileID_woodenStairstoUp", here
  WriteTileIDToIni tileID_rampToNorth, "tileID_rampToNorth", here
  WriteTileIDToIni tileID_rampToSouth, "tileID_rampToSouth", here
  
  
  WriteTileIDToIni tileID_desertRamptoUp, "tileID_desertRamptoUp", here
  
  WriteTileIDToIni tileID_rampToRightCycMountain, "tileID_rampToRightCycMountain", here
  WriteTileIDToIni tileID_rampToLeftCycMountain, "tileID_rampToLeftCycMountain", here
  
  WriteTileIDToIni tileID_jungleStairsToNorth, "tileID_jungleStairsToNorth", here
  WriteTileIDToIni tileID_jungleStairsToLeft, "tileID_jungleStairsToLeft", here
  
  ' + requires rightClick
  WriteTileIDToIni tileID_ladderToUp, "tileID_ladderToUp", here
  
  ' + requires rope
  WriteTileIDToIni tileID_holeInCelling, "tileID_holeInCelling", here
  

  ' to DOWN
  WriteTileIDToIni tileID_grassCouldBeHole, "tileID_grassCouldBeHole", here
  WriteTileIDToIni tileID_pitfall, "tileID_pitfall", here

  WriteTileIDToIni tileID_openHole, "tileID_openHole", here
  WriteTileIDToIni tileID_OpenDesertLooseStonePile, "tileID_OpenDesertLooseStonePile", here
  
  WriteTileIDToIni tileID_trapdoor, "tileID_trapdoor", here
  WriteTileIDToIni tileID_down1, "tileID_down1", here
  
  WriteTileIDToIni tileID_openHole2, "tileID_openHole2", here
  
  WriteTileIDToIni tileID_trapdoor2, "tileID_trapdoor2", here
  WriteTileIDToIni tileID_down2, "tileID_down2", here
  WriteTileIDToIni tileID_stairsToDownKazordoon, "tileID_stairsToDownKazordoon", here
  WriteTileIDToIni tileID_stairsToDownThais, "tileID_stairsToDownThais", here
  
  WriteTileIDToIni tileID_trapdoorKazordoon, "tileID_trapdoorKazordoon", here
  WriteTileIDToIni tileID_down3, "tileID_down3", here
  WriteTileIDToIni tileID_stairsToDown, "tileID_stairsToDown", here
  
  WriteTileIDToIni tileID_stairsToDown2, "tileID_stairsToDown2", here
  WriteTileIDToIni tileID_woodenStairstoDown, "tileID_woodenStairstoDown", here
  
  WriteTileIDToIni tileID_rampToDown, "tileID_rampToDown", here
    
  ' + requires rightClick
  WriteTileIDToIni tileID_sewerGate, "tileID_sewerGate", here
  WriteTileIDToIni tileID_trapdoor, "tileID_trapdoor", here
  WriteTileIDToIni tileID_trapdoor2, "tileID_trapdoor2", here
  
  ' + requires shovel
  WriteTileIDToIni tileID_closedHole, "tileID_closedHole", here
  WriteTileIDToIni tileID_desertLooseStonePile, "tileID_desertLooseStonePile", here
  WriteTileIDToIni tileID_OpenDesertLooseStonePile, "tileID_OpenDesertLooseStonePile", here
  
   ' FOOD
  WriteTileIDToIni tileID_firstFoodTileID, "tileID_firstFoodTileID", here
  WriteTileIDToIni tileID_lastFoodTileID, "tileID_lastFoodTileID", here
  WriteTileIDToIni tileID_firstMushroomTileID, "tileID_firstMushroomTileID", here
  WriteTileIDToIni tileID_lastMushroomTileID, "tileID_lastMushroomTileID", here
  
  
  'FIELD RANGE1
  WriteTileIDToIni tileID_firstFieldRangeStart, "tileID_firstFieldRangeStart", here
  WriteTileIDToIni tileID_firstFieldRangeEnd, "tileID_firstFieldRangeEnd", here
  WriteTileIDToIni tileID_secondFieldRangeStart, "tileID_secondFieldRangeStart", here
  WriteTileIDToIni tileID_secondFieldRangeEnd, "tileID_secondFieldRangeEnd", here

  WriteTileIDToIni tileID_campFire1, "tileID_campFire1", here
  WriteTileIDToIni tileID_campFire2, "tileID_campFire2", here

  'WALKABLE FIELDS
  WriteTileIDToIni tileID_walkableFire1, "tileID_walkableFire1", here
  WriteTileIDToIni tileID_walkableFire2, "tileID_walkableFire2", here
  WriteTileIDToIni tileID_walkableFire3, "tileID_walkableFire3", here

  ' DEPOT CHEST
  WriteTileIDToIni tileID_depotChest, "tileID_depotChest", here
  

  WriteTileIDToIni tileID_health_potion, "tileID_health_potion", here
  WriteTileIDToIni tileID_strong_health_potion, "tileID_strong_health_potion", here
  WriteTileIDToIni tileID_great_health_potion, "tileID_great_health_potion", here
  WriteTileIDToIni tileID_small_health_potion, "tileID_small_health_potion", here
  WriteTileIDToIni tileID_mana_potion, "tileID_mana_potion", here
  WriteTileIDToIni tileID_strong_mana_potion, "tileID_strong_mana_potion", here
  WriteTileIDToIni tileID_great_mana_potion, "tileID_great_mana_potion", here
  
  WriteTileIDToIni tileID_ultimate_health_potion, "tileID_ultimate_health_potion", here
  WriteTileIDToIni tileID_great_spirit_potion, "tileID_great_spirit_potion", here
  
  ' flasks - mana fluids
  WriteTileIDToIni tileID_flask, "tileID_flask", here
  strInfo = "&H" & Hex(byteNothing)
  i = setBlackdINI("tileIDs", "byteNothing", strInfo, here)
  strInfo = "&H" & Hex(byteMana)
  i = setBlackdINI("tileIDs", "byteMana", strInfo, here)
  strInfo = "&H" & Hex(byteLife)
  i = setBlackdINI("tileIDs", "byteLife", strInfo, here)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
  If thisShouldNotBeLoading = 0 Then
    Exit Sub
  End If
  Me.Hide
  Cancel = BlockUnload
  If Cancel = False Then
  For i = 0 To SckClient.UBound
    SckClient(i).Close
    'Unload SckClient(i)
  Next i
  For i = 0 To sckServer.UBound
    sckServer(i).Close
    'Unload SckServer(i)
  Next i
  For i = 0 To Me.sckClientGame.UBound
    sckClientGame(i).Close
    'Unload SckClientGame(i)
  Next i
  For i = 0 To Me.sckServerGame.UBound
    sckServerGame(i).Close
    'Unload SckServerGame(i)
  Next i
'  For i = 0 To Me.sckFasterLogin.UBound
'    sckFasterLogin(i).Close
'  Next i
  End If
End Sub







Private Sub ForwardGameTo_Change()
  ModifyTibiaIPs
End Sub

Private Sub ForwardGameTo_Validate(Cancel As Boolean)
  ModifyTibiaIPs
End Sub

Private Sub gridLog_Click()
  ' user clicked in the gridLog
  If chkSelect.value = 1 Then
    'flash on the equivalent cell
    gridLog.Row = gridLog.RowSel
    If gridLog.ColSel < 10 Then
      gridLog.Col = gridLog.ColSel + 11
    ElseIf gridLog.ColSel > 10 Then
      gridLog.Col = gridLog.ColSel - 11
    End If
  End If
End Sub

' [COMMON WITH FREE PROXY]


Private Sub SckClient_Close(Index As Integer)
  ' client closes
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  SckClient(Index).Close
  'SckServer(index).Close 'close his brother server
  If Connected(Index) = True Then
   Connected(Index) = False
   txtPackets.Text = txtPackets.Text & vbCrLf & "#client" & Index & " closed#"
  End If
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckClient_Close(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub

Private Sub SckClient_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    ' client connects
    Dim i As Integer
    Dim num As Integer
    Dim useID As Long
    Dim lngUseServerID As Long
    Dim strFrom As String
    Dim idLoginSP As Long
    Dim strTmp2 As String
    Dim firstGTC As Long
    Dim blnDoingBest As Boolean
    Dim ntries As Long
    Dim blnTimeout As Boolean
    Dim lngVirtualTooSlow As Long

    #If FinalMode Then
    On Error GoTo gotErr
    #End If
    useID = 0
    For i = 1 To MAXCLIENTS
        If Connected(i) = False Then
            useID = i
            Exit For
        End If
    Next i
    If useID > SckClient.UBound Then
        num = SckClient.UBound + 1
        Load SckClient(num)
        Load sckServer(num)
    End If
    If useID > 0 Then
        SckClient(useID).Close
        SckClient(useID).Accept requestID
        DoEvents
        strFrom = SckClient(useID).RemoteHostIP
        If frmMain.chkBlockRemote.value = 1 Then
            If Left$(strFrom, 8) <> "127.0.0." Then
                txtPackets.Text = txtPackets.Text & vbCrLf & "ALARM: Remote connection detected from " & SckClient(useID).RemoteHostIP & " ( " & SckClient(useID).RemoteHost & " ) to port " & CStr(SckClient(useID).RemotePort) & " It have been blocked. Please block your Blackd Proxy ports in your router or firewall for higher security."
                ChangePlayTheDangerSound True
                SckClient(useID).Close
                DoEvents
                Exit Sub
            End If
        End If
        If TrueServer1.value = False Then
            PREFEREDLOGINSERVER = ForwardGameTo.Text
            PREFEREDLOGINPORT = CLng(txtServerLoginP.Text)
            blnDoingBest = False
        ElseIf ((chkForceLoginServer.value = 1) Or (TrueServer1.value = False)) Then
            PREFEREDLOGINSERVER = cmbPrefered.Text
            blnDoingBest = False
        Else
            blnDoingBest = True
            PREFEREDLOGINSERVER = cmbPrefered.Text
            'PREFEREDLOGINSERVER = getFasterLoginServer()
            'txtPackets.Text = txtPackets.Text & vbCrLf & ">> Checked fastest login server = " & PREFEREDLOGINSERVER & " ( " & CStr(fastestLoginServerTime) & " ms )"
        End If
        DoEvents
        For idLoginSP = 1 To NumberOfLoginServers
            If PREFEREDLOGINSERVER = trueLoginServer(idLoginSP) Then
               PREFEREDLOGINPORT = trueLoginPort(idLoginSP)
            End If
        Next idLoginSP
    
        gotFirstLoginPacket(useID) = False
        Connected(useID) = True
        If frmAdvanced.chkWantBypass.value = 0 Then ' new in 9.38
            If TrueServer1.value = True Then
                txtPackets.Text = txtPackets.Text & vbCrLf & "#client" & useID & " connected (IP " & _
                SckClient(useID).RemoteHostIP & ") , forwarding to " & PREFEREDLOGINSERVER & ":" & _
                 CStr(PREFEREDLOGINPORT) & " #"
                sckServer(useID).Close
                sckServer(useID).RemoteHost = PREFEREDLOGINSERVER
                sckServer(useID).RemotePort = PREFEREDLOGINPORT
            Else
                txtPackets.Text = txtPackets.Text & vbCrLf & "#client" & useID & " connected (IP " & _
                SckClient(useID).RemoteHostIP & ") , forwarding to " & ForwardGameTo.Text & ":" & _
                 txtServerLoginP.Text & " #"
                sckServer(useID).Close
                On Error GoTo gotHostErr
                If ForwardGameTo.Text = "" Then
                  GoTo gotHostErr
                End If
                sckServer(useID).RemoteHost = ForwardGameTo.Text
                On Error GoTo gotPortErr
                sckServer(useID).RemotePort = CLng(txtServerLoginP.Text)
            End If
            On Error GoTo gotErr
            If blnDoingBest = False Then
                sckServer(useID).Connect
                ' ot servers
                'Debug.Print "connected to " & sckServer(useID).RemoteHost & ":" & sckServer(useID).RemotePort
            Else
                ntries = 0
                lngVirtualTooSlow = TOOSLOWLOGINSERVER_MS
                Do
                    ntries = ntries + 1
                    sckServer(useID).Close
                    DoEvents
                    ConnectionSignal(useID) = False
                    firstGTC = GetTickCount()
                    blnTimeout = False
                    sckServer(useID).RemoteHost = PREFEREDLOGINSERVER
                    sckServer(useID).RemotePort = PREFEREDLOGINPORT
                    sckServer(useID).Connect
                    
                    Do
                        DoEvents
                        If SckClient(useID).State = sckClosed Then
                            txtPackets.Text = txtPackets.Text & vbCrLf & "Client #" & CStr(useID) & " closed connection"
                            Exit Sub
                        End If
                        If GetTickCount() > (firstGTC + lngVirtualTooSlow) Then
                            blnTimeout = True
                        End If
                    Loop Until ((ConnectionSignal(useID) = True) Or (blnTimeout = True))
                    If ConnectionSignal(useID) = True Then
                        txtPackets.Text = txtPackets.Text & vbCrLf & "Good login server found ( " & CStr(GetTickCount() - firstGTC) & " ms ) : " & PREFEREDLOGINSERVER
                        cmbPrefered.Text = PREFEREDLOGINSERVER
                    Else
                        sckServer(useID).Close
                        DoEvents
                        txtPackets.Text = txtPackets.Text & vbCrLf & "Login server too slow ( >" & CStr(lngVirtualTooSlow) & " ms ) : " & PREFEREDLOGINSERVER
                        For idLoginSP = 1 To NumberOfLoginServers
                            If PREFEREDLOGINSERVER = trueLoginServer(idLoginSP) Then
                                If idLoginSP = NumberOfLoginServers Then
                                    PREFEREDLOGINSERVER = trueLoginServer(1)
                                    PREFEREDLOGINPORT = trueLoginPort(1)
                                Else
                                    PREFEREDLOGINSERVER = trueLoginServer(idLoginSP + 1)
                                    PREFEREDLOGINPORT = trueLoginPort(idLoginSP + 1)
                                    cmbPrefered.Text = PREFEREDLOGINSERVER
                                End If
                                Exit For
                            End If
                        Next idLoginSP
                    End If
                    If (ntries = NumberOfLoginServers) Then
                        lngVirtualTooSlow = TOOSLOWLOGINSERVER_MS * 10
                    End If
                Loop Until ((ConnectionSignal(useID) = True) Or (ntries >= (NumberOfLoginServers * 2)))
            End If
        End If
    End If
    Exit Sub
gotErr:
    frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckClient_ConnectionRequest(" & Index & "," & requestID & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
    Exit Sub
gotHostErr:
    frmMain.Show
    frmMain.WindowState = vbNormal
    frmMain.SetFocus
    frmMain.Refresh
    MsgBox "Please enter a valid server IP", vbOKOnly + vbExclamation, "Error"
    Exit Sub
gotPortErr:
    
    frmMain.Show
    frmMain.WindowState = vbNormal
    frmMain.SetFocus
    frmMain.Refresh
    MsgBox "Please enter a valid server Port", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub SckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  ' data arrives from client
  Dim timeOut As Long
  Dim packet() As Byte 'a tibia packet is an array of bytes
  Dim i As Integer
  Dim res As Long
  Dim msg As String
  Dim mypid As Long
  Dim strIP As String
  Dim strip2 As String
  Dim gtcnow As Long
  'Dim strAccount As String
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  
  'If LoginMethod = 1 Then
  
  SckClient(Index).GetData packet, vbArray + vbByte
  ' PAUSE /  UNCOMMENT MSGBOX TO FIND ADDRESS OF LAST PACKET SENT,adrLastPacket (matching testing.txt , with tsearch)
  'OverwriteOnFile "testing.txt", frmMain.showAsStr2(packet, 0)
  'MsgBox "debug - continue"
  strIP = SckClient(Index).LocalIP
  strip2 = SckClient(Index).RemoteHostIP
  mypid = GiveProcessIDbyLastPacket(packet, strIP, strip2, "LOGIN1")
  
'    If TibiaVersionLong >= 841 Then
'        'pendingggg
'        strAccount = GetMemoryString(myPID, adrAccount)
'        AddProcessIdAccountRelation strAccount, myPID
'    End If
'
  If (UseCrackd = True) Then
    If (gotFirstLoginPacket(Index) = False) Then
      res = readLoginTibiaKeyAtPID(Index, mypid)
      If res < 0 Then
        Connected(Index) = False
        SckClient(Index).Close
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & _
         "WARNING: readLoginTibiaKeyAtPID failed! (this is a debug message that might be ignored)"
        Exit Sub
      End If
      gotFirstLoginPacket(Index) = True

      #If BufferDebug = 1 Then
      LogOnFile "bufferLog.txt", "USING DECIPHER KEY1 = " & _
       GoodHex(loginPacketKey(Index).Key(0)) & " " & _
       GoodHex(loginPacketKey(Index).Key(1)) & " " & _
       GoodHex(loginPacketKey(Index).Key(2)) & " " & _
       GoodHex(loginPacketKey(Index).Key(3)) & " " & _
       GoodHex(loginPacketKey(Index).Key(4)) & " " & _
       GoodHex(loginPacketKey(Index).Key(5)) & " " & _
       GoodHex(loginPacketKey(Index).Key(6)) & " " & _
       GoodHex(loginPacketKey(Index).Key(7)) & " " & _
       GoodHex(loginPacketKey(Index).Key(8)) & " " & _
       GoodHex(loginPacketKey(Index).Key(9)) & " " & _
       GoodHex(loginPacketKey(Index).Key(10)) & " " & _
       GoodHex(loginPacketKey(Index).Key(11)) & " " & _
       GoodHex(loginPacketKey(Index).Key(12)) & " " & _
       GoodHex(loginPacketKey(Index).Key(13)) & " " & _
       GoodHex(loginPacketKey(Index).Key(14)) & " " & _
       GoodHex(loginPacketKey(Index).Key(15)) & vbCrLf
      #End If

      If chkLogPackets.value = 1 Then
      txtPackets.Text = txtPackets.Text & vbCrLf & "USING DECIPHER KEY1 = " & _
       GoodHex(loginPacketKey(Index).Key(0)) & " " & _
       GoodHex(loginPacketKey(Index).Key(1)) & " " & _
       GoodHex(loginPacketKey(Index).Key(2)) & " " & _
       GoodHex(loginPacketKey(Index).Key(3)) & " " & _
       GoodHex(loginPacketKey(Index).Key(4)) & " " & _
       GoodHex(loginPacketKey(Index).Key(5)) & " " & _
       GoodHex(loginPacketKey(Index).Key(6)) & " " & _
       GoodHex(loginPacketKey(Index).Key(7)) & " " & _
       GoodHex(loginPacketKey(Index).Key(8)) & " " & _
       GoodHex(loginPacketKey(Index).Key(9)) & " " & _
       GoodHex(loginPacketKey(Index).Key(10)) & " " & _
       GoodHex(loginPacketKey(Index).Key(11)) & " " & _
       GoodHex(loginPacketKey(Index).Key(12)) & " " & _
       GoodHex(loginPacketKey(Index).Key(13)) & " " & _
       GoodHex(loginPacketKey(Index).Key(14)) & " " & _
       GoodHex(loginPacketKey(Index).Key(15)) & vbCrLf
       End If
    End If
  Else
      gotFirstLoginPacket(Index) = False
  End If
  If chkLogPackets.value = 1 Then
    LogLine "CLIENT" & Index & ":"
    LogPacket packet
    txtPackets.Text = txtPackets.Text & vbCrLf & "CLIENT" & Index & ">" & showAsStr2(packet, 0)
    txtPackets.SelStart = Len(txtPackets.Text)
  End If
  
    If frmAdvanced.chkWantBypass.value = 1 Then ' new in 9.38
       BypassLoginServer Index
       Exit Sub
    End If
    
  timeOut = GetTickCount() + 30000
  ' Debug.Print "Connected = " & Connected(Index) & " ; state = " & sckServer(Index).State
  While ((Connected(Index) = True) And (sckServer(Index).State <> sckConnected))
      gtcnow = GetTickCount()
      If gtcnow >= timeOut Then
          'frmMain.DoCloseActions index
          'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TIMEOUT(at loginclient) for ID " & CStr(index)
          Exit Sub
        End If
    'Debug.Print sckServer(Index).State
    If sckServer(Index).State = sckClosed Then
      Connected(Index) = False
    End If
    DoEvents 'wait
  Wend
  If ((Connected(Index) = True) And (sckServer(Index).State = sckConnected)) Then
    sckServer(Index).SendData packet
  End If
  
  
  'Else 'LOGIN METHOD= 0
  
 
  
  
  'End If
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & Index & " lost connection at SckClient_DataArrival #"
  Connected(Index) = False
  DoEvents
End Sub

Private Sub SckClientGame_Close(Index As Integer)
  ' game client closes
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  If frmRunemaker.chkCloseSound.value = 1 Then
     ChangePlayTheDangerSound True
  End If
  If TibiaVersionLong >= 841 Then
    GameConnected(Index) = False
  End If
  FirstCharInCharList(Index) = ""

  
  sckClientGame(Index).Close
  sckServerGame(Index).Close 'close his brother server
  txtPackets.Text = txtPackets.Text & vbCrLf & "#gameclient" & Index & " closed ( by client )#"
  DoCloseActions Index
  Exit Sub
gotErr:
 frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckClientGame_Close(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
 DoCloseActions Index
End Sub



Private Sub sckClientGame_Connect(Index As Integer)
'Debug.Print "clientgame connect:" & Index
 setsockopt sckClientGame(Index).SocketHandle, IPPROTO_TCP, TCP_NODELAY, 1, 4
End Sub

Private Sub closeAllTibiaClientsExcept(ByVal mypid As Long)
Dim tibiaclient As Long
Dim pid As Long
Dim bRes As Boolean
If (mypid > 0) Then
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      If (tibiaclient <> mypid) Then
        bRes = ProcessTerminate(, tibiaclient)
      End If
    End If
  Loop
End If
End Sub

Private Sub SckClientGame_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  ' game client gets connection request
  Dim i As Integer
  Dim num As Integer
  Dim useID As Long
  Dim strRemoteIP As String
  Dim listPos As Integer
  Dim pres As Long
  Dim selName As String
  Dim tmpID As Integer
  Dim UpdatedDATE As Date
  Dim res As Long
  Dim tmpPORT As Long
  Dim strToDomain As String
  Dim tmpStr As String
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  
  If cteDebugConEvents = True Then
     LogConEvent "SckClientGame_ConnectionRequest(" & CStr(Index) & "," & CStr(requestID) & ")"
  End If
  useID = 0
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = False Then
      useID = i
      Exit For
    End If
  Next i
  If useID > sckClientGame.UBound Then
    num = sckClientGame.UBound + 1
    Load sckClientGame(num)
    Load sckServerGame(num)
  End If
  If useID > 0 Then
    If useID > HighestConnectionID Then
      HighestConnectionID = useID
    End If
    sckClientGame(useID).Close
    sckClientGame(useID).Accept requestID
    If frmMain.chkBlockRemote.value = 1 Then
      strRemoteIP = sckClientGame(useID).RemoteHostIP
      If Left$(strRemoteIP, 8) <> "127.0.0." Then
        txtPackets.Text = txtPackets.Text & vbCrLf & "ALARM: Remote connection detected from " & sckClientGame(useID).RemoteHostIP & " ( " & sckClientGame(useID).RemoteHost & " ) to port " & CStr(sckClientGame(useID).RemotePort) & " It have been blocked. Please block your Blackd Proxy ports in your router or firewall for higher security."
        ChangePlayTheDangerSound True
        sckClientGame(useID).Close
        DoEvents
        Exit Sub
      End If
    End If
    If TibiaVersionLong <= 840 Then
      GameConnected(useID) = True
    End If
    sentFirstPacket(useID) = False
    lastPing(useID) = GetTickCount()
    txtPackets.Text = txtPackets.Text & vbCrLf & "#gameclient" & useID & " connected (IP " & _
     sckClientGame(useID).RemoteHostIP & ") #"
    If TibiaVersionLong >= 841 Then
      'now gameserver must send something first
      tmpID = CInt(useID)
      ProcessID(tmpID) = GetProcessIdByAdrConnected()
      If cteDebugConEvents = True Then
        LogConEvent "Connection request at mc slot " & CStr(tmpID) & " from PID = " & ProcessID(tmpID)
      End If
      'ProcessID(tmpID) = GetProcessIdByManualDebug()
      If ProcessID(tmpID) = -1 Then
        tmpStr = "#critical error 4 on connection " & tmpID & " , closing it#"
        txtPackets.Text = txtPackets.Text & vbCrLf & tmpStr
        If cteDebugConEvents = True Then
          LogConEvent tmpStr
        End If
        sckClientGame(tmpID).Close
        sckServerGame(tmpID).Close
        GameConnected(tmpID) = False
        DoCloseActions tmpID
        Exit Sub
      End If
      If ProcessID(tmpID) = -2 Then
        tmpStr = "#critical error 4 on connection " & tmpID & " , closing it#"
        txtPackets.Text = txtPackets.Text & vbCrLf & tmpStr
        If cteDebugConEvents = True Then
          LogConEvent tmpStr
        End If
        sckClientGame(tmpID).Close
        sckServerGame(tmpID).Close
        GameConnected(tmpID) = False
        DoCloseActions tmpID
        Exit Sub
      End If
      If (TibiaVersionLong >= 1100) Then ' Tibia 11.00
        If Tibia11allowed = False Then
            End
        End If
        Dim strSelectedChar As String
        strSelectedChar = QMemory_ReadStringP(ProcessID(tmpID), adrSelectedCharName_afterCharList)
        'Debug.Print ("SELECTED CHAR=" & strSelectedChar)
        ' TODO: Complete connection system in 11.0
        ' build char list
        ' set listpos
        ' set selName
        BuildCharListForTibiaQ tmpID, selName, listPos
        If listPos = -1 Then ' unexpected packet
          tmpStr = "#critical error 1 on connection " & tmpID & " , closing it#"
          txtPackets.Text = txtPackets.Text & vbCrLf & tmpStr
          If cteDebugConEvents = True Then
            LogConEvent tmpStr
          End If
          sckClientGame(tmpID).Close
          sckServerGame(tmpID).Close
          GameConnected(tmpID) = False
          DoCloseActions tmpID
          Exit Sub
        End If
      Else ' old
        listPos = GetCharListPositionPre(tmpID, selName)
        ' important problem in tibia 9.71 !
        pres = UpdateCharListFromMemory(tmpID, listPos)
        ' NEW ANTIBAN FEATURE:
        ' SINCE BLACKD PROXY 22.2, WE NOW RELOAD TIBIA.DAT AT NINJA PATCHES:
        ' If ((TibiaVersionLong = highestTibiaVersionLong) And (UseRealTibiaDatInLatestTibiaVersion = True)) Then
        UpdatedDATE = GetDATEOfFile(TibiaExePathWITHTIBIADAT)
        If (Not (CurrentTibiaDatDATE = UpdatedDATE)) Then
          CurrentTibiaDatDATE = UpdatedDATE ' fix 22.3
          ' close the rest of clients because they would be using outdated tibia.dat other way
          closeAllTibiaClientsExcept ProcessID(tmpID)
          ' reload tibia.dat
          res = UnifiedLoadDatFile(TibiaExePathWITHTIBIADAT)
          If ((res = -1) Or (res = -2)) Then
            MsgBox "Non compatible tibia.dat file , error " & CStr(res), vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
            End
          End If
          If (res = -3) Then
            MsgBox "Too many tiles found in tibia.dat , please increase MAXDATTILES in your settings.ini" & CStr(res), vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
            End
          End If
          If (res = -4) Then
            MsgBox "Outstanding error -4 while reading tibia.dat: " & vbCrLf & DBGtileError, vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
            End
          End If
          If (res = -5) Then
            MsgBox "Bug caught: " & vbCrLf & DBGtileError, vbOKOnly, "Debug report"
            End
          End If
          frmMenu.Caption = "Updated Tibia.dat : " & CStr(UpdatedDATE)
        End If
        listPos = GetCharListPosition2(tmpID, selName)
        If listPos = -1 Then ' unexpected packet
          txtPackets.Text = txtPackets.Text & vbCrLf & "#critical error 1 on connection " & tmpID & " , closing it#"
          sckClientGame(tmpID).Close
          sckServerGame(tmpID).Close
          GameConnected(tmpID) = False
          DoCloseActions tmpID
          Exit Sub
        End If
      End If
      CharacterName(tmpID) = selName
      '     frmTrueMap.LoadChars
      '     frmRunemaker.LoadRuneChars
      '     frmStealth.LoadStealthChars
      '     frmHPmana.LoadHPmanaChars
      '     frmEvents.LoadEventChars
      '     frmCondEvents.LoadCondEventChars
      '     frmTrainer.LoadTrainerChars
      '     frmCavebot.LoadCavebotChars
      '     MustCheckFirstClientPacket(tmpID) = False
      MustCheckFirstClientPacket(tmpID) = True
      NeedToIgnoreFirstGamePacket(tmpID) = True
      If cteDebugConEvents = True Then
         LogConEvent "M = True ; N = True"
      End If
      If TrueServer2.value = True Then
        txtPackets.Text = txtPackets.Text & vbCrLf & "# the client ID " & tmpID & " selected the character " & _
         selName & " - forwarding connection to " & _
         ForwardGameTo.Text & _
         CStr(txtServerGameP.Text) & " #"
        sckServerGame(tmpID).Close
        sckServerGame(tmpID).RemoteHost = ForwardGameTo.Text
        sckServerGame(tmpID).RemotePort = CLng(txtServerGameP.Text)
        sckServerGame(tmpID).Connect
      Else ' real servers
        If (LimitedToServer <> "-") Then
          If (LimitedToServer <> CharacterList2(tmpID).item(listPos).ServerName) Then
            txtPackets.Text = txtPackets.Text & vbCrLf & "#the client ID " & tmpID & " have been closed: You are only allowed to connect to " & LimitedToServer & " with this friend account"
            LogOnFile "errors.txt", "You are only allowed to connect to " & LimitedToServer & " with this friend account"
            frmMain.DoCloseActions tmpID
            Exit Sub
          End If
        End If
        logoutAllowed(tmpID) = 20000 + GetTickCount() ' disable reconnection 20 sec
        If TibiaVersionLong >= 1011 Then
          If (TibiaVersionLong = 1011) Then
            strToDomain = tibiaEntryServer
          Else
            strToDomain = CharacterList2(tmpID).item(listPos).serverDOMAIN
          End If
          txtPackets.Text = txtPackets.Text & vbCrLf & "#the client ID " & tmpID & " selected the character " & _
           CharacterList2(tmpID).item(listPos).CharacterName & " - forwarding connection to " & strToDomain & ":" & _
           CStr(CharacterList2(tmpID).item(listPos).serverPort) & " ( " & _
          CharacterList2(tmpID).item(listPos).ServerName & " ) #"
          sckServerGame(tmpID).Close
          If TibiaVersionLong = 1011 Then
            sckServerGame(tmpID).RemoteHost = tibiaEntryServer
          Else
            sckServerGame(tmpID).RemoteHost = CharacterList2(tmpID).item(listPos).serverDOMAIN
          End If
          tmpPORT = CLng(CharacterList2(tmpID).item(listPos).serverPort)
          sckServerGame(tmpID).RemotePort = tmpPORT
          sckServerGame(tmpID).Connect
          If cteDebugConEvents = True Then
            LogConEvent "Connecting to " & CharacterList2(tmpID).item(listPos).ServerName & " = " & CharacterList2(tmpID).item(listPos).serverDOMAIN & ":" & CharacterList2(tmpID).item(listPos).serverPort
          End If
        Else ' old method
          txtPackets.Text = txtPackets.Text & vbCrLf & "#the client ID " & tmpID & " selected the character " & _
           CharacterList2(tmpID).item(listPos).CharacterName & " - forwarding connection to " & _
           CharacterList2(tmpID).item(listPos).serverIP1 & "." & _
           CharacterList2(tmpID).item(listPos).serverIP2 & "." & _
           CharacterList2(tmpID).item(listPos).serverIP3 & "." & _
           CharacterList2(tmpID).item(listPos).serverIP4 & ":" & _
           CStr(CharacterList2(tmpID).item(listPos).serverPort) & " ( " & _
           CharacterList2(tmpID).item(listPos).ServerName & " ) #"
          sckServerGame(tmpID).Close
          sckServerGame(tmpID).RemoteHost = _
          buildIPstring(CInt(CharacterList2(tmpID).item(listPos).serverIP1), _
           CInt(CharacterList2(tmpID).item(listPos).serverIP2), _
           CInt(CharacterList2(tmpID).item(listPos).serverIP3), _
           CInt(CharacterList2(tmpID).item(listPos).serverIP4))
          sckServerGame(tmpID).RemotePort = CLng(CharacterList2(tmpID).item(listPos).serverPort)
          sckServerGame(tmpID).Connect
        End If ' old method
      End If ' else - real servers
    End If  ' tibia > 8.41
  End If ' If useID > 0 Then
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckClientGame_ConnectionRequest(" & Index & "," & requestID & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub
Private Sub HandleSckClientGame_Data(ByRef Index As Integer, ByRef MyCodingIsLazyPacket() As Byte, ByRef bytesTotal As Long)
  ' gameclient gets data
  Dim packet() As Byte 'a tibia packet is an array of bytes
  Dim listPos As Integer
  Dim selName As String
  Dim res As Integer
  Dim aRes As Long
  Dim timeOut As Long
  Dim realRawPacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim processIt As Boolean
  Dim SPpos As Long
  Dim SPlim As Long
  Dim SPlen As Long
  Dim SPpacket() As Byte
  Dim strIP As String
  Dim strip2 As String
  Dim tmpLen As Long

  #If FinalMode Then
  On Error GoTo errclose
  #End If
  
  
  
  If Index > 0 Then
  processIt = True
  If (UseCrackd = True) And (MustCheckFirstClientPacket(Index) = False) Then
    'sckClientGame(Index).GetData realRawPacket, vbArray + vbByte
    ReDim realRawPacket(bytesTotal - 1)
    RtlMoveMemory realRawPacket(0), MyCodingIsLazyPacket(0), bytesTotal
    SPpos = 0
    'Exit Sub 'borrame
    SPlim = UBound(realRawPacket)
    If SPlim < 0 Then
     Debug.Print "Warning: connection probably lost"
     Exit Sub
    End If
'    Debug.Print "received packet with " & CStr(SPlim + 1) & " bytes" & ":" & showAsStr(realRawPacket, True)
    Do ' NEW improved loop since Blackdproxy 8.77
        ' deals with 'stickied' packets from client
        If TibiaVersionLong < 830 Then
            SPlen = GetTheLong(realRawPacket(SPpos), realRawPacket(SPpos + 1))
            ReDim SPpacket(SPlen + 1)
            RtlMoveMemory SPpacket(0), realRawPacket(SPpos), (SPlen + 2)
            pres = DecipherTibiaProtected(SPpacket(0), packetKey(Index).Key(0), UBound(SPpacket), UBound(packetKey(Index).Key))
        Else ' skip CRC
            SPlen = GetTheLong(realRawPacket(SPpos), realRawPacket(SPpos + 1))
            ReDim SPpacket(SPlen + 1)
            RtlMoveMemory SPpacket(0), realRawPacket(SPpos), (SPlen + 2)
            pres = DecipherTibiaProtectedSP(SPpacket(0), packetKey(Index).Key(0), UBound(SPpacket), UBound(packetKey(Index).Key))
        End If
        
        If (pres = 0) Then
            If TibiaVersionLong < 830 Then
                hbytes = GetTheLong(SPpacket(2), SPpacket(3))
                'Debug.Print showAsStr(SPpacket, True)
                ReDim packet(hbytes + 1)
                RtlMoveMemory packet(0), SPpacket(2), (hbytes + 2)
            Else
                hbytes = GetTheLong(SPpacket(6), SPpacket(7))
                ReDim packet(hbytes + 1)
                RtlMoveMemory packet(0), SPpacket(6), (hbytes + 2)
            End If
        Else
            If pres = -1 Then
              ' somehow a login packet arrived here
              ReDim packet(UBound(realRawPacket))
              RtlMoveMemory packet(0), realRawPacket(0), UBound(realRawPacket) + 1
              MustCheckFirstClientPacket(Index) = True
              If cteDebugConEvents = True Then
                LogConEvent "M = True. Starting workaround for rare error"
              End If
              GoTo workAroundForRareError
            Else
              GiveCrackdDllErrorMessage pres, SPpacket, packetKey(Index).Key, UBound(SPpacket), UBound(packetKey(Index).Key), 1
              Exit Sub
            End If
        End If
        res = ApplyHardcoreCheats(packet, Index)
        If res <> 1 Then
            UnifiedSendToServerGame Index, packet, True
        End If
        SPpos = 2 + SPpos + SPlen
        If SPpos < SPlim Then
          res = 0 'set a debug here for tests. looks like not the problem
        End If
    Loop While (SPpos < SPlim)
    Exit Sub
  Else
    'sckClientGame(Index).GetData packet, vbArray + vbByte
    ReDim packet(bytesTotal - 1)
    RtlMoveMemory packet(0), MyCodingIsLazyPacket(0), bytesTotal
    
  End If
workAroundForRareError:
  'MustCheckFirstClientPacket(Index) = False ' tibia 8.41 !
  If MustCheckFirstClientPacket(Index) = True Then
       'store connection packet to allow reconnection later
       res = UBound(packet)
       ReDim ReconnectionPacket(Index).packet(res)
       ReconnectionPacket(Index).numbytes = res + 1
       RtlMoveMemory ReconnectionPacket(Index).packet(0), packet(0), ReconnectionPacket(Index).numbytes

      'LogOnFile "lastp.txt", showAsStr2(packet, 0) & vbCrLf
     'If AlternativeBinding = 0 Then
     '   UpdateProcessIDbyLastPacket Index, packet
     'Else
        strIP = sckClientGame(Index).LocalIP
        strip2 = sckClientGame(Index).RemoteHostIP
        If TibiaVersionLong >= 841 Then
            ' processid(index) already holds a valid value
            
            'ProcessID(Index) = GetProcessIDfromCharList3(Index)
        Else
            ProcessID(Index) = GiveProcessIDbyLastPacket(packet, strIP, strip2, "GAMESERVERLOGIN")
        End If
     'End If
     If ProcessID(Index) <= 0 Then
       DoCloseActions Index
       Exit Sub
     End If

     If TibiaVersionLong <= 840 Then
     listPos = GetCharListPositionPre(Index, selName)
     pres = UpdateCharListFromMemory(Index, listPos)
     End If
     If (UseCrackd = True) Then
      ' Debug.Print frmMain.showAsStr(packet, True)
       
       pres = readTibiaKeyAtPID(Index, ProcessID(Index))
       If (pres = -1) Then
         DoCloseActions Index
         Exit Sub
       End If
       #If BufferDebug = 1 Then
         LogOnFile "bufferLog.txt", "USING DECIPHER KEY2 = " & _
       GoodHex(packetKey(Index).Key(0)) & " " & _
       GoodHex(packetKey(Index).Key(1)) & " " & _
       GoodHex(packetKey(Index).Key(2)) & " " & _
       GoodHex(packetKey(Index).Key(3)) & " " & _
       GoodHex(packetKey(Index).Key(4)) & " " & _
       GoodHex(packetKey(Index).Key(5)) & " " & _
       GoodHex(packetKey(Index).Key(6)) & " " & _
       GoodHex(packetKey(Index).Key(7)) & " " & _
       GoodHex(packetKey(Index).Key(8)) & " " & _
       GoodHex(packetKey(Index).Key(9)) & " " & _
       GoodHex(packetKey(Index).Key(10)) & " " & _
       GoodHex(packetKey(Index).Key(11)) & " " & _
       GoodHex(packetKey(Index).Key(12)) & " " & _
       GoodHex(packetKey(Index).Key(13)) & " " & _
       GoodHex(packetKey(Index).Key(14)) & " " & _
       GoodHex(packetKey(Index).Key(15)) & vbCrLf
       #End If
       If chkLogPackets.value = 1 Then
      txtPackets.Text = txtPackets.Text & vbCrLf & "USING DECIPHER KEY2 = " & _
       GoodHex(packetKey(Index).Key(0)) & " " & _
       GoodHex(packetKey(Index).Key(1)) & " " & _
       GoodHex(packetKey(Index).Key(2)) & " " & _
       GoodHex(packetKey(Index).Key(3)) & " " & _
       GoodHex(packetKey(Index).Key(4)) & " " & _
       GoodHex(packetKey(Index).Key(5)) & " " & _
       GoodHex(packetKey(Index).Key(6)) & " " & _
       GoodHex(packetKey(Index).Key(7)) & " " & _
       GoodHex(packetKey(Index).Key(8)) & " " & _
       GoodHex(packetKey(Index).Key(9)) & " " & _
       GoodHex(packetKey(Index).Key(10)) & " " & _
       GoodHex(packetKey(Index).Key(11)) & " " & _
       GoodHex(packetKey(Index).Key(12)) & " " & _
       GoodHex(packetKey(Index).Key(13)) & " " & _
       GoodHex(packetKey(Index).Key(14)) & " " & _
       GoodHex(packetKey(Index).Key(15)) & vbCrLf
       End If
     End If 'usecrackd
     processIt = False

     listPos = GetCharListPosition2(Index, selName)

     If listPos = -1 Then ' unexpected packet
       txtPackets.Text = txtPackets.Text & vbCrLf & "#critical error 2 on connection " & Index & " , closing it#"
       sckClientGame(Index).Close
       sckServerGame(Index).Close
       GameConnected(Index) = False
       DoCloseActions Index
       Exit Sub
     End If

     CharacterName(Index) = selName
     If TibiaVersionLong >= 841 Then
        GameConnected(Index) = True
     End If
     ' events that happens when a char complete the login stage
    frmTrueMap.LoadChars
    frmRunemaker.LoadRuneChars
    'custom ng
    frmHealing.LoadHealingChars
    frmExtras.LoadextrasChars
    frmPersistent.LoadPersistentChars
    frmAimbot.LoadAimbotChars
    '
    frmStealth.LoadStealthChars
    frmHPmana.LoadHPmanaChars
    frmEvents.LoadEventChars
    frmCondEvents.LoadCondEventChars
    frmTrainer.LoadTrainerChars
    frmCavebot.LoadCavebotChars
    frmBroadcast.LoadBroadcastChars
    LoadCharSettings Index
    
    
     MustCheckFirstClientPacket(Index) = False
     If cteDebugConEvents = True Then
        LogConEvent "M = False"
     End If
     If TibiaVersionLong <= 840 Then
     If TrueServer2.value = True Then
       
       txtPackets.Text = txtPackets.Text & vbCrLf & "# the client ID " & Index & " selected the character " & _
         selName & " - forwarding connection to " & _
         ForwardGameTo.Text & _
         CStr(txtServerGameP.Text) & " #"
       sckServerGame(Index).Close
       sckServerGame(Index).RemoteHost = ForwardGameTo.Text
       sckServerGame(Index).RemotePort = CLng(txtServerGameP.Text)
       sckServerGame(Index).Connect
     Else
       If (LimitedToServer <> "-") Then
         If (LimitedToServer <> CharacterList2(Index).item(listPos).ServerName) Then
             txtPackets.Text = txtPackets.Text & vbCrLf & "#the client ID " & Index & " have been closed: You are only allowed to connect to " & LimitedToServer & " with this friend account"
             LogOnFile "errors.txt", "You are only allowed to connect to " & LimitedToServer & " with this friend account"
             frmMain.DoCloseActions Index
             Exit Sub
         End If
       End If
              logoutAllowed(Index) = 20000 + GetTickCount() ' disable reconnection 20 sec
       
      
       
       RecordLoginOnFile CharacterList2(Index).item(listPos).CharacterName, buildIPstring(CInt(CharacterList2(Index).item(listPos).serverIP1), _
         CInt(CharacterList2(Index).item(listPos).serverIP2), _
         CInt(CharacterList2(Index).item(listPos).serverIP3), _
         CInt(CharacterList2(Index).item(listPos).serverIP4)), CLng(CharacterList2(Index).item(listPos).serverPort), Index
       
       txtPackets.Text = txtPackets.Text & vbCrLf & "#the client ID " & Index & " selected the character " & _
         CharacterList2(Index).item(listPos).CharacterName & " - forwarding connection to " & _
         CharacterList2(Index).item(listPos).serverIP1 & "." & _
         CharacterList2(Index).item(listPos).serverIP2 & "." & _
         CharacterList2(Index).item(listPos).serverIP3 & "." & _
         CharacterList2(Index).item(listPos).serverIP4 & ":" & _
         CStr(CharacterList2(Index).item(listPos).serverPort) & " ( " & _
         CharacterList2(Index).item(listPos).ServerName & " ) #"
       sckServerGame(Index).Close
       sckServerGame(Index).RemoteHost = _
         buildIPstring(CInt(CharacterList2(Index).item(listPos).serverIP1), _
         CInt(CharacterList2(Index).item(listPos).serverIP2), _
         CInt(CharacterList2(Index).item(listPos).serverIP3), _
         CInt(CharacterList2(Index).item(listPos).serverIP4))
       sckServerGame(Index).RemotePort = CLng(CharacterList2(Index).item(listPos).serverPort)
       

       
       sckServerGame(Index).Connect
     End If
     
     End If
  End If 'first packet
  'If chkLogPackets.Value = 1 Then
  '  LogLine "GAMECLIENT" & Index & ":"
  '  LogPacket packet
  '  txtPackets.Text = txtPackets.Text & vbCrLf & "GAMECLIENT" & Index & ">" & showAsStr2(packet, 0)
  '  txtPackets.SelStart = Len(txtPackets.Text)
  'End If
  ' apply hardcore cheats
  If (processIt = True) Then
  If frmHardcoreCheats.chkApplyCheats.value = 1 Then
     res = ApplyHardcoreCheats(packet, Index)
     If res = 1 Then ' Hardcore cheats require skiping this packet
        Exit Sub
     End If
  End If
  End If
  timeOut = GetTickCount() + 30000
  While ((GameConnected(Index) = True) And (sckServerGame(Index).State <> sckConnected))
      If GetTickCount() >= timeOut Then
          'frmMain.DoCloseActions index
          'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TIMEOUT(at gameclient) for ID " & CStr(index)
          Exit Sub
        End If
    If sckClientGame(Index).State = sckClosed Then
      GameConnected(Index) = False
    End If
    DoEvents 'wait
  Wend
  If GameConnected(Index) = True And sckServerGame(Index).State = sckConnected Then
      If (processIt = True) Then
        UnifiedSendToServerGame Index, packet, True
      Else
        If chkLogPackets.value = 1 Then
          LogLine "GAMECLIENT" & Index & ":"
          LogPacket packet
          txtPackets.Text = txtPackets.Text & vbCrLf & "GAMECLIENT" & Index & ">" & showAsStr2(packet, 0)
          txtPackets.SelStart = Len(txtPackets.Text)
        End If
        sckServerGame(Index).SendData packet
      End If
  End If
  End If
  
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & Index & " lost connection at HandleSckClientGame_Data #"
  frmMain.DoCloseActions Index
  DoEvents
End Sub

Private Sub SckClientGame_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim FullPacket() As Byte
    Dim SubPacket() As Byte
    Dim Offset As Long
    Dim SubPacketLen As Long

    sckClientGame(Index).GetData FullPacket, vbArray + vbByte
    Offset = 0
    While Offset < bytesTotal    '
        If UBound(FullPacket) < Offset + 2 Then    'should never happen. means coding error
            Err.Raise _
                    Number:=errIndexOutOfRange, _
                    Description:="Failed to read all packets in SckClientGame_DataArrival (after reading 'all' packets, offset was not equal to bytesTotal)", _
                    Source:="SckClientGame_DataArrival"

        End If

        SubPacketLen = GetTheLong(FullPacket(Offset), FullPacket(Offset + 1)) + 2
        ReDim SubPacket(SubPacketLen - 1)
        RtlMoveMemory SubPacket(0), FullPacket(Offset), SubPacketLen
        'If (UBound(SubPacket) <> UBound(FullPacket)) Or SubPacketLen <> bytesTotal Then
        '    Err.Raise _
             '        Number:=1, _
             '        Description:="Sum Ting Wong", _
             '        Source:="SckClientGame_DataArrival"
        '
        'End If

        HandleSckClientGame_Data Index, SubPacket, SubPacketLen
        Offset = Offset + SubPacketLen
    Wend

    If Offset <> bytesTotal Then    'this should never happen.. means coding error
        Err.Raise _
                Number:=errIndexOutOfRange, _
                Description:="Failed to read all packets in SckClientGame_DataArrival (after reading 'all' packets, offset was not equal to bytesTotal)", _
                Source:="SckClientGame_DataArrival"
    End If


End Sub





Public Sub UnifiedSendToClientGame(ByVal Index As Integer, ByRef packet() As Byte, Optional forceOldMode As Boolean = False)
  Dim extrab As Long
  Dim i As Long
  Dim rnumber As Byte
  Dim totalLong As Long
  Dim goodPacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim lngwsck As Long
  Dim thedamnCRC As Long
  Dim fourBytesCRC(3) As Byte
'  Dim thedamnCRC2 As Long
'  Dim fourBytesCRC2(3) As Byte
  Dim onlygood As Long
  Dim dbg1 As Long
  Dim dbg2 As Long
  Dim dbact As Boolean
  dbact = False
  
  If GameConnected(Index) = True Or ((TibiaVersionLong >= 841) And (forceOldMode = True)) Then
  If ((UseCrackd = True) And (forceOldMode = False)) Then
    If TibiaVersionLong < 830 Then
        totalLong = GetTheLong(packet(0), packet(1))
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
        totalLong = totalLong + 2
        ReDim goodPacket(totalLong + 1)
        hbytes = UBound(packet) + 1
        RtlMoveMemory goodPacket(2), packet(0), (totalLong)
        goodPacket(0) = LowByteOfLong(totalLong)
        goodPacket(1) = HighByteOfLong(totalLong)
        pres = EncipherTibiaProtected(goodPacket(0), packetKey(Index).Key(0), UBound(goodPacket), UBound(packetKey(Index).Key))
    Else

        totalLong = GetTheLong(packet(0), packet(1))
        onlygood = totalLong + 2
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
    
        ReDim goodPacket(totalLong + 7)
        If dbact = True Then
            frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "DEBUG2: " & frmMain.showAsStr(goodPacket, True)
            dbact = True
        End If
        hbytes = UBound(packet) + 1
        RtlMoveMemory goodPacket(6), packet(0), (onlygood)
        goodPacket(0) = LowByteOfLong(UBound(goodPacket) - 1)
        goodPacket(1) = HighByteOfLong(UBound(goodPacket) - 1)
        pres = EncipherTibiaProtectedSP(goodPacket(0), packetKey(Index).Key(0), UBound(goodPacket), UBound(packetKey(Index).Key))
        thedamnCRC = GetTibiaCRC(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
        longToBytes fourBytesCRC, thedamnCRC
        'Debug.Print "t1:" & GoodHex(fourBytesCRC(0)) & " " & GoodHex(fourBytesCRC(1)) & " " & GoodHex(fourBytesCRC(2)) & " " & GoodHex(fourBytesCRC(3))
    
'        thedamnCRC2 = GetTibiaCRC2(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
'        longToBytes fourBytesCRC2, thedamnCRC2
'        If Not (((fourBytesCRC(0) = fourBytesCRC2(0)) And (fourBytesCRC(1) = fourBytesCRC2(1)) And (fourBytesCRC(2) = fourBytesCRC2(2)) And (fourBytesCRC(3) = fourBytesCRC2(3)))) Then
'          Debug.Print "no match!!!"
'          Debug.Print "res1:" & GoodHex(fourBytesCRC(0)) & " " & GoodHex(fourBytesCRC(1)) & " " & GoodHex(fourBytesCRC(2)) & " " & GoodHex(fourBytesCRC(3))
'          Debug.Print "res2:" & GoodHex(fourBytesCRC2(0)) & " " & GoodHex(fourBytesCRC2(1)) & " " & GoodHex(fourBytesCRC2(2)) & " " & GoodHex(fourBytesCRC2(3))
'        End If
        goodPacket(2) = fourBytesCRC(0)
        goodPacket(3) = fourBytesCRC(1)
        goodPacket(4) = fourBytesCRC(2)
        goodPacket(5) = fourBytesCRC(3)
        If dbact = True Then
            frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "DEBUG3: " & frmMain.showAsStr(goodPacket, True)
            dbact = True
        End If
        dbg1 = UBound(goodPacket) - 1
        dbg2 = GetTheLong(goodPacket(0), goodPacket(1))
        If dbg1 <> dbg2 Then
            frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unstable packets detected: " & CStr(dbg1) & " <> " & CStr(dbg2)
        End If
        'Debug.Print "2<< " & frmMain.showAsStr(goodPacket, True) ' DEBUGGGGGGGGGGGGGGGGGGGGGGGGGG
    End If
    
    If (pres < 0) Then
        GiveCrackdDllErrorMessage pres, goodPacket, packetKey(Index).Key, UBound(goodPacket), UBound(packetKey(Index).Key), 201
        Exit Sub
    End If
    lngwsck = sckClientGame(Index).State
    If lngwsck = sckConnected Then
        sckClientGame(Index).SendData goodPacket
    Else
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "GAMECLIENT #" & CStr(Index) & " (" & CharacterName(Index) & ") closed because winsock state was not connected (" & CStr(lngwsck) & ")"
        frmMain.DoCloseActions Index
        DoEvents
    End If
  Else
    lngwsck = sckClientGame(Index).State
    If lngwsck = sckConnected Then
        sckClientGame(Index).SendData packet
    Else
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "GAMECLIENT #" & CStr(Index) & " (" & CharacterName(Index) & ") closed because winsock state was not connected (" & CStr(lngwsck) & ")"
        frmMain.DoCloseActions Index
        DoEvents
    End If
  End If
  End If
End Sub

Public Sub UnifiedSendToClient(ByVal Index As Integer, ByRef packet() As Byte, Optional forceOldMode As Boolean = False, _
 Optional packetHaveStrangeBytes As Boolean = False)
  Dim extrab As Long
  Dim i As Long
  Dim rnumber As Byte
  Dim totalLong As Long
  Dim goodPacket() As Byte
  
  Dim hbytes As Long
  Dim pres As Long
  Dim thedamnCRC As Long
  Dim fourBytesCRC(3) As Byte
  If Connected(Index) = True Then
  If ((UseCrackd = True) And (forceOldMode = False)) Then
    If packetHaveStrangeBytes = False Then
        totalLong = GetTheLong(packet(0), packet(1))
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
        totalLong = totalLong + 2
        ReDim goodPacket(totalLong + 1)
        hbytes = UBound(packet) + 1
        RtlMoveMemory goodPacket(2), packet(0), (totalLong)
        goodPacket(0) = LowByteOfLong(totalLong)
        goodPacket(1) = HighByteOfLong(totalLong)
        pres = EncipherTibiaProtected(goodPacket(0), loginPacketKey(Index).Key(0), UBound(goodPacket), UBound(loginPacketKey(Index).Key))
        If (pres < 0) Then
            GiveCrackdDllErrorMessage pres, goodPacket, loginPacketKey(Index).Key, UBound(goodPacket), UBound(loginPacketKey(Index).Key), 202
            Exit Sub
        End If
    
        SckClient(Index).SendData goodPacket
    Else ' new since 8.3 , 4 CRC bytes
        totalLong = UBound(packet) + 1
        ReDim goodPacket(totalLong - 1)
        RtlMoveMemory goodPacket(0), packet(0), (totalLong)
        

        pres = EncipherTibiaProtectedSP(goodPacket(0), loginPacketKey(Index).Key(0), UBound(goodPacket), UBound(loginPacketKey(Index).Key))
        
        ' fix CRC
        
        thedamnCRC = GetTibiaCRC(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
        longToBytes fourBytesCRC, thedamnCRC
        goodPacket(2) = fourBytesCRC(0)
        goodPacket(3) = fourBytesCRC(1)
        goodPacket(4) = fourBytesCRC(2)
        goodPacket(5) = fourBytesCRC(3)
        
        
        If (pres < 0) Then
            GiveCrackdDllErrorMessage pres, goodPacket, loginPacketKey(Index).Key, UBound(goodPacket), UBound(loginPacketKey(Index).Key), 203
            Exit Sub
        End If
        SckClient(Index).SendData goodPacket
    End If
  Else
    SckClient(Index).SendData packet
  End If
  End If
End Sub

Public Sub UnifiedSendToServerGame(ByVal Index As Integer, ByRef packet() As Byte, logIt As Boolean)
  Dim extrab As Long
  Dim i As Long
  Dim rnumber As Byte
  Dim totalLong As Long
  Dim goodPacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim thedamnCRC As Long
  Dim fourBytesCRC(3) As Byte
  Dim onlygood As Long
  
  If sckServerGame(Index).State <> sckConnected Then
    If frmHardcoreCheats.chkAutorelog.value = 1 Then
      If ReconnectionStage(Index) = 0 Then
        pres = GiveGMmessage(Index, "The connection with the server was lost, doing reconnection now", "Warning")
        DoEvents
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "The connection with the server have been lost on client #" & CStr(Index)
        StartReconnection Index
      Else
        Exit Sub
      End If
    Else
      If (PlayTheDangerSound = False) Then
        ChangePlayTheDangerSound True
        pres = GiveGMmessage(Index, "The connection with the server was lost, use exiva close or exiva relog", "BlackdProxy")
        DoEvents
        frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "The connection with the server have been lost on client #" & CStr(Index)
      End If
    End If
    Exit Sub
  End If
  If GameConnected(Index) = True Then
  If (UseCrackd = True) Then
    If chkLogPackets.value = 1 Then
      If logIt = True Then
        LogLine "GAMECLIENT" & Index & ":"
        LogPacket packet
        txtPackets.Text = txtPackets.Text & vbCrLf & "GAMECLIENT" & Index & ">" & showAsStr2(packet, 0)
        txtPackets.SelStart = Len(txtPackets.Text)
      End If
    End If

    
    If TibiaVersionLong < 830 Then
        totalLong = GetTheLong(packet(0), packet(1))
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
        totalLong = totalLong + 2
        ReDim goodPacket(totalLong + 1)
        hbytes = UBound(packet) + 1
        RtlMoveMemory goodPacket(2), packet(0), (totalLong)
        goodPacket(0) = LowByteOfLong(totalLong)
        goodPacket(1) = HighByteOfLong(totalLong)
        pres = EncipherTibiaProtected(goodPacket(0), packetKey(Index).Key(0), UBound(goodPacket), UBound(packetKey(Index).Key))
    
    Else
        'Debug.Print "1>> " & frmMain.showAsStr(packet, True) ' DEBUGGGGGGGGGGGGGGGGGGGGGGGGGG
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
        pres = EncipherTibiaProtectedSP(goodPacket(0), packetKey(Index).Key(0), UBound(goodPacket), UBound(packetKey(Index).Key))
        thedamnCRC = GetTibiaCRC(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
        longToBytes fourBytesCRC, thedamnCRC
        goodPacket(2) = fourBytesCRC(0)
        goodPacket(3) = fourBytesCRC(1)
        goodPacket(4) = fourBytesCRC(2)
        goodPacket(5) = fourBytesCRC(3)
        'Debug.Print "2>> " & frmMain.showAsStr(goodPacket, True) ' DEBUGGGGGGGGGGGGGGGGGGGGGGGGGG
    End If
    
    
    If (pres < 0) Then
        GiveCrackdDllErrorMessage pres, goodPacket, packetKey(Index).Key, UBound(goodPacket), UBound(packetKey(Index).Key), 3
        Exit Sub
    End If
    If (sckServerGame(Index).State = sckConnected) Then
        sckServerGame(Index).SendData goodPacket
    Else
       If Index > 0 Then
        DoCloseActions Index
        DoEvents
       Exit Sub
       End If
    End If
  Else
    If chkLogPackets.value = 1 Then
      If logIt = True Then
        LogLine "GAMECLIENT" & Index & ":"
        LogPacket packet
        txtPackets.Text = txtPackets.Text & vbCrLf & "GAMECLIENT" & Index & ">" & showAsStr2(packet, 0)
        txtPackets.SelStart = Len(txtPackets.Text)
      End If
    End If
    sckServerGame(Index).SendData packet
  End If
  End If
End Sub

' Unused code
'Private Sub sckFasterLogin_Connect(Index As Integer)
'    fastestconnect = CLng(Index)
'    setsockopt sckFasterLogin(Index).SocketHandle, IPPROTO_TCP, TCP_NODELAY, 1, 4
'End Sub





Private Sub SckServer_Close(Index As Integer)
  ' server closes
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  sckServer(Index).Close
  DoEvents
  SckClient(Index).Close 'close his brother client ??
  
  If ((Connected(Index) = True) Or (DoingMainLoopLogin(Index) = True)) Then
    txtPackets.Text = txtPackets.Text & vbCrLf & "#server" & Index & " closed (by server) #"
    Connected(Index) = False
    DoingMainLoopLogin(Index) = False
  End If
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckServer_Close(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub

Private Sub SckServer_Connect(Index As Integer)
  ' server connects
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  setsockopt sckServer(Index).SocketHandle, IPPROTO_TCP, TCP_NODELAY, 1, 4
  If Index > 0 Then
    ConnectionSignal(Index) = True
  End If
  ReDim ConnectionBufferLogin(Index).packet(0)
  ConnectionBufferLogin(Index).numbytes = 0
  DoingMainLoopLogin(Index) = False
  txtPackets.Text = txtPackets.Text & vbCrLf & "#server" & Index & " connected#"
  Exit Sub
gotErr:
 frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckServer_Connect(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub



Private Function LearnFromServerLogin(ByRef packet() As Byte, ByVal Index As Integer, ByVal strIP As String, Optional bstart As Long = 2) As Long
    Dim c As Byte
    Dim res As Long
Dim tmpLong As Long

    If UBound(packet) < 2 Then
        LearnFromServerLogin = 0
        Exit Function
    End If
    c = packet(bstart)
    Select Case c
    Case &H28
       'Debug.Print "LOGIN TYPE " & GoodHex(c)
       If TibiaVersionLong >= 1091 Then
          res = PacketIPchange6(packet, Index, strIP, bstart)
      
            If res = 1 Then
          
              If CloseLoginServerAfterCharList = True Then
                 Debug.Print "Closing login connection"
                 If Index > 0 Then
                    sckServer(Index).Close
                 End If
              End If
              LearnFromServerLogin = 1
              Exit Function
            End If
       
       ElseIf TibiaVersionLong >= 1074 Then
         res = PacketIPchange5(packet, Index, strIP, bstart)
      
            If res <> 1 Then
              txtPackets.Text = txtPackets.Text & vbCrLf & "ERROR: FAILED TO MODIFY LOGIN PACKET!"
            Else
              If CloseLoginServerAfterCharList = True Then
                 If Index > 0 Then
                    sckServer(Index).Close
                 End If
              End If
              LearnFromServerLogin = 1
              Exit Function
            End If
       End If
    Case &H14
     ' Debug.Print "LOGIN TYPE " & GoodHex(c)
      If TibiaVersionLong >= 1091 Then
        res = PacketIPchange6(packet, Index, strIP, bstart)
        If res = 1 Then
              If CloseLoginServerAfterCharList = True Then
                 If Index > 0 Then
                    Debug.Print "Closing login connection"
                    sckServer(Index).Close
                 End If
              End If
              LearnFromServerLogin = 1
              Exit Function
        Else
          LearnFromServerLogin = 0
          Exit Function
        End If
      ElseIf TibiaVersionLong >= 1074 Then
        res = PacketIPchange5b(packet, Index, strIP, bstart)
      ElseIf TibiaVersionLong >= 1012 Then
        res = PacketIPchange4(packet, Index, strIP, bstart)
      ElseIf TibiaVersionLong >= 1011 Then
        res = PacketIPchange3(packet, Index, strIP, bstart)
      Else
        res = PacketIPchange2(packet, Index, strIP, bstart)
      End If
      If res <> 1 Then
         txtPackets.Text = txtPackets.Text & vbCrLf & "ERROR: FAILED TO MODIFY LOGIN PACKET!"
      Else
         If CloseLoginServerAfterCharList = True Then
          If Index > 0 Then
             sckServer(Index).Close
          End If
         End If
         If TibiaVersionLong >= 1012 Then
           LearnFromServerLogin = 1
        '   LearnFromServerLogin = 1 ' skip original packet (function PacketIPchange4 already sent modified packet)
           Exit Function
         End If
      End If
    Case Else
      Debug.Print "unknown server login packet (" & GoodHex(c) & ") : " & frmMain.showAsStr(packet, True);
      
    End Select
    LearnFromServerLogin = 0
End Function



Private Sub SckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  ' data arrives from game server
  Dim rawpacket() As Byte
  Dim newSizeBuffer As Long
  Dim i As Long
  Dim j As Long
  Dim startB As Long
  Dim endB As Long
  Dim iniB As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ' Get it
  sckServer(Index).GetData rawpacket, vbArray + vbByte
  

  If Index > 0 Then
  #If BufferDebug Then
    LogOnFile "bufferLogLogin.txt", "NEW RAWPACKET:"
    LogOnFile "bufferLogLogin.txt", showAsStr2(rawpacket, 0)
  #End If
  iniB = ConnectionBufferLogin(Index).numbytes ' save initial bytes of buffer
  ' enlarge buffer if needed
  If (UBound(rawpacket) + 1) > ((UBound(ConnectionBufferLogin(Index).packet) + 1) - ConnectionBufferLogin(Index).numbytes) Then
    newSizeBuffer = ConnectionBufferLogin(Index).numbytes + UBound(rawpacket)
    ReDim Preserve ConnectionBufferLogin(Index).packet(newSizeBuffer)
    
    #If BufferDebug Then
    LogOnFile "bufferLogLogin.txt", "BUFFER WAS RESIZED TO " & CStr(newSizeBuffer)
    #End If
  End If
  startB = iniB
  endB = startB + UBound(rawpacket)
  ConnectionBufferLogin(Index).numbytes = iniB + 1 + UBound(rawpacket)

  RtlMoveMemory ConnectionBufferLogin(Index).packet(startB), rawpacket(0), (OptCte4 * (endB - startB + 1))
  'j = 0
  'For i = startB To endB
  '  ConnectionBuffer(index).packet(i) = rawpacket(j)
  '  j = j + 1
  'Next i
  #If BufferDebug Then
  LogOnFile "bufferLogLogin.txt", "USEFULL BUFFER ARE FIRST " & CStr(ConnectionBufferLogin(Index).numbytes) & " BYTES OF THIS :"
  LogOnFile "bufferLogLogin.txt", showAsStr2(ConnectionBufferLogin(Index).packet, 0)
  #End If
  If DoingMainLoopLogin(Index) = False Then ' if not doing main loop right now, then start it
    DoMainLoopLogin Index
  End If
  End If
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# internal ID" & Index & " lost connection at SckServer_DataArrival #"
  'frmMain.DoCloseActions Index
  Connected(Index) = False
  DoEvents
End Sub

Public Sub DoMainLoopLogin(idConnection As Integer)
  ' Here the buffer of data we got from server is processed
  Dim startB As Long
  Dim lastB As Long
  Dim longPacket As Long
  Dim lPminusOne As Long
  Dim i As Long
  Dim packet() As Byte
  Dim amLeft As Long
  Dim tmpV As Long
  Dim lRes As Integer
  Dim timeOut As Long
  Dim withHeaderL As Long
  Dim withHeaderS As Long
  Dim nBytes As Long
  Dim rawpacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim strIP As String
  Dim specialMessage As Boolean
  Dim debugloginline As Long
  Dim extradebugl As String
  Dim thedamnCRC As Long
  Dim fourBytesCRC(3) As Byte
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  extradebugl = ""
  'INITIAL CHECKS (only once per call)
  debugloginline = 1
  DoingMainLoopLogin(idConnection) = True
  debugloginline = 2
  specialMessage = False
  debugloginline = 3
  startB = 0
  debugloginline = 4
  If ConnectionBufferLogin(idConnection).numbytes < 2 Then
    debugloginline = 5
    #If BufferDebug Then
    LogOnFile "bufferLogLogin.txt", "not even 2 bytes at start..."
    #End If
    debugloginline = 6
    DoingMainLoopLogin(idConnection) = False
    debugloginline = 7
    Exit Sub ' not even 2 bytes at start ...
  End If
  debugloginline = 8
  longPacket = GetTheLong(ConnectionBufferLogin(idConnection).packet(0), ConnectionBufferLogin(idConnection).packet(1))
  debugloginline = 9
  If longPacket > ((ConnectionBufferLogin(idConnection).numbytes) - 2) Then
    #If BufferDebug Then
    LogOnFile "bufferLogLogin.txt", "no complete packet at start..."
    #End If
    debugloginline = 10
    DoingMainLoopLogin(idConnection) = False
    Exit Sub ' no complete packet at start...
  End If
  debugloginline = 11
  startB = 2
  'extract 1 complete packet
  lPminusOne = longPacket - 1
  debugloginline = 12
  lastB = startB + lPminusOne
  debugloginline = 13
nextLoop:
debugloginline = 14
  withHeaderL = lPminusOne + 2
  debugloginline = 15
  withHeaderS = startB - 2
  debugloginline = 16
  nBytes = withHeaderL + 1
  debugloginline = 17
  
  ' decipher it
  If UseCrackd = True Then
  debugloginline = 18
    ReDim rawpacket(withHeaderL)
    debugloginline = 19
    RtlMoveMemory rawpacket(0), ConnectionBufferLogin(idConnection).packet(withHeaderS), (OptCte4 * nBytes)
    debugloginline = 20
    #If BufferDebug Then
    ' conexion parte 1
    LogOnFile "bufferLogLogin.txt", "EXTRACTING 1 COMPLETE PACKET:"
    debugloginline = 21
    LogOnFile "bufferLogLogin.txt", showAsStr2(rawpacket, 0)
    debugloginline = 22
    #End If
    specialMessage = False
    debugloginline = 23
    If TibiaVersionLong >= 830 Then
        'testing the function GetTibiaCRC
        'theDamnCRC = GetTibiaCRC(rawpacket(6), UBound(rawpacket) - 5) ' (number of bytes - 6)
        'longToBytes fourBytesCRC, theDamnCRC
        'debug.Print frmmain.showAsStr(rawpacket,True)
        
        pres = DecipherTibiaProtectedSP(rawpacket(0), loginPacketKey(idConnection).Key(0), UBound(rawpacket), UBound(loginPacketKey(idConnection).Key))
    Else
        pres = DecipherTibiaProtected(rawpacket(0), loginPacketKey(idConnection).Key(0), UBound(rawpacket), UBound(loginPacketKey(idConnection).Key))
    End If
    debugloginline = 24
    ' LETS REPAIRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR THIS!
    If (pres = -1) And (rawpacket(2) = &HA) And (TibiaVersionLong < 830) Then
        debugloginline = 25
        specialMessage = True
    Else
        debugloginline = 26
        If (pres < 0) Then
            debugloginline = 27
            GiveCrackdDllErrorMessage pres, rawpacket, loginPacketKey(idConnection).Key, UBound(rawpacket), UBound(loginPacketKey(idConnection).Key), 6
            Exit Sub
        End If
        debugloginline = 28
        If TibiaVersionLong < 830 Then
            hbytes = GetTheLong(rawpacket(2), rawpacket(3))
            debugloginline = 29
            ReDim packet(hbytes + 1)
            debugloginline = 30
            RtlMoveMemory packet(0), rawpacket(2), (hbytes + 2)
            debugloginline = 31
        Else
            hbytes = UBound(rawpacket) - 1
            debugloginline = 29
            ReDim packet(UBound(rawpacket))
            debugloginline = 30
            RtlMoveMemory packet(0), rawpacket(0), (hbytes + 2)
            debugloginline = 31
        End If
        #If BufferDebug Then
        ' parte 2
        LogOnFile "bufferLogLogin.txt", "DECIPHERED IT:"
        LogOnFile "bufferLogLogin.txt", showAsStr2(packet, 0)
        #End If
    End If

  Else
    debugloginline = 32
    ReDim packet(withHeaderL)
    debugloginline = 33
    RtlMoveMemory packet(0), ConnectionBufferLogin(idConnection).packet(withHeaderS), (OptCte4 * nBytes)
    
    #If BufferDebug Then
    LogOnFile "bufferLogLogin.txt", "EXTRACTING 1 COMPLETE PACKET:"
    LogOnFile "bufferLogLogin.txt", showAsStr2(packet, 0)
    #End If
  End If
  
    debugloginline = 34
      If chkLogPackets.value = 1 Then
        debugloginline = 35
        LogLine "SERVER" & idConnection & ":"
        debugloginline = 36
        LogPacket packet
        debugloginline = 37
        txtPackets.Text = txtPackets.Text & vbCrLf & "SERVER" & idConnection & "<" & showAsStr2(packet, 0)
        debugloginline = 38
        txtPackets.SelStart = Len(txtPackets.Text)
      End If
        debugloginline = 39
        strIP = SckClient(idConnection).LocalIP
        debugloginline = 40
        If TibiaVersionLong >= 830 Then
            ' parte 3
            lRes = LearnFromServerLogin(packet, idConnection, strIP, 8)
        Else
            lRes = LearnFromServerLogin(packet, idConnection, strIP)
        End If
        debugloginline = 41
        If lRes = 1 Then ' Hardcore cheats require skiping this packet
           GoTo nextP
        ElseIf lRes = 3 Then ' Hardcore cheats require losing connection
            debugloginline = 42
           sckServer(idConnection).Close
           debugloginline = 43
           SckClient(idConnection).Close
           debugloginline = 44
           Connected(idConnection) = False
           debugloginline = 45
           ConnectionBufferLogin(idConnection).numbytes = 0
           Exit Sub
        End If
    debugloginline = 46
      timeOut = GetTickCount() + 30000
      debugloginline = 47
      While ((Connected(idConnection) = True) And (SckClient(idConnection).State <> sckConnected))
      debugloginline = 48
        If GetTickCount() >= timeOut Then
            debugloginline = 49
          'frmMain.DoCloseActions idconnection
          'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TIMEOUT(at gameserver) for ID " & CStr(idconnection)
          Exit Sub
        End If
        debugloginline = 50
        If SckClient(idConnection).State = sckClosed Then
            debugloginline = 51
          Connected(idConnection) = False
        End If
        debugloginline = 52
        DoEvents 'wait
      Wend
      debugloginline = 53
      If Connected(idConnection) = True And SckClient(idConnection).State = sckConnected Then
        debugloginline = 54
        If TibiaVersionLong < 830 Then
            UnifiedSendToClient idConnection, packet, specialMessage, False ' fixed since 13.4
        Else
            UnifiedSendToClient idConnection, packet, specialMessage, True
        End If
debugloginline = 55
    If chkLogPackets.value = 1 Then
        debugloginline = 56
        LogLine "SERVER" & idConnection & ":"
        debugloginline = 57
        LogPacket packet
        debugloginline = 58
        txtPackets.Text = txtPackets.Text & vbCrLf & "SERVER" & idConnection & "<" & showAsStr2(packet, 0)
        debugloginline = 59
        txtPackets.SelStart = Len(txtPackets.Text)
        debugloginline = 60
    End If
    debugloginline = 61
    
        DoEvents
      End If
nextP:
    debugloginline = 62
     If Connected(idConnection) = False Then
        debugloginline = 63
       DoingMainLoopLogin(idConnection) = False
       Exit Sub
     End If
     debugloginline = 64
     ' move pointer
     startB = startB + longPacket
     ' if no complete packet left, move residue to start and end
     debugloginline = 65
     If startB = ConnectionBufferLogin(idConnection).numbytes Then
       ' buffer is now empty
        debugloginline = 66
       ConnectionBufferLogin(idConnection).numbytes = 0
       debugloginline = 67
       DoingMainLoopLogin(idConnection) = False
       debugloginline = 68
       Exit Sub
     End If
     debugloginline = 69
     If (startB + 1) = ConnectionBufferLogin(idConnection).numbytes Then
       ' a single byte left
       debugloginline = 70
       #If BufferDebug Then
       LogOnFile "bufferLogLogin.txt", "a single byte left at the end..."
       #End If
       debugloginline = 71
       ConnectionBufferLogin(idConnection).numbytes = 1
       debugloginline = 72
       ConnectionBufferLogin(idConnection).packet(0) = ConnectionBufferLogin(idConnection).packet(startB)
       debugloginline = 73
       DoingMainLoopLogin(idConnection) = False
       debugloginline = 74
       Exit Sub
     End If
     debugloginline = 75
     If (startB + 2) = ConnectionBufferLogin(idConnection).numbytes Then
       ' two bytes left
       debugloginline = 76
       #If BufferDebug Then
       LogOnFile "bufferLog.txt", "a pair of bytes left at the end..."
       #End If
       debugloginline = 77
       ConnectionBufferLogin(idConnection).numbytes = 2
       debugloginline = 78
       ConnectionBufferLogin(idConnection).packet(0) = ConnectionBufferLogin(idConnection).packet(startB)
       debugloginline = 79
       ConnectionBufferLogin(idConnection).packet(1) = ConnectionBufferLogin(idConnection).packet(startB + 1)
       debugloginline = 80
       DoingMainLoop(idConnection) = False
       debugloginline = 81
       Exit Sub
     End If
     debugloginline = 82
     longPacket = GetTheLong(ConnectionBufferLogin(idConnection).packet(startB), ConnectionBufferLogin(idConnection).packet(startB + 1))
     debugloginline = 83
     lPminusOne = longPacket - 1
     debugloginline = 84
     startB = startB + 2
     debugloginline = 85
     lastB = startB + lPminusOne
     debugloginline = 86
     If (startB + longPacket) > ConnectionBufferLogin(idConnection).numbytes Then
        debugloginline = 87
       ' not complete packets left - save rest in start of buffer
       startB = startB - 2
       debugloginline = 88
       tmpV = (ConnectionBufferLogin(idConnection).numbytes) - startB
       debugloginline = 89
       amLeft = tmpV - 1
       debugloginline = 90
       ConnectionBufferLogin(idConnection).numbytes = tmpV
       debugloginline = 91
       #If BufferDebug Then
       LogOnFile "bufferLogLogin.txt", CStr(tmpV) & " bytes left at the end..."
       #End If
       debugloginline = 92
       RtlMoveMemory ConnectionBufferLogin(idConnection).packet(0), ConnectionBufferLogin(idConnection).packet(startB), (OptCte4 * tmpV)
       debugloginline = 93
       DoingMainLoopLogin(idConnection) = False
       debugloginline = 94
       Exit Sub
     End If
     debugloginline = 95
     If Connected(idConnection) = False Then
        debugloginline = 96
       DoingMainLoopLogin(idConnection) = False
       debugloginline = 97
       Exit Sub
     End If
     debugloginline = 98
     GoTo nextLoop ' loop until buffer have no complete packets
errclose:
  DoingMainLoopLogin(idConnection) = False
  If debugloginline = 40 Then
  extradebugl = " >> WHILE USING THE FOLLOWING PARAMETERS " & showAsStr2(packet, 0) & ", " & CStr(idConnection) & ", " & strIP
  End If
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# internal ID " & idConnection & _
   " lost connection at DoMainLoopLogin, line " & CStr(debugloginline) & " with error #" & CStr(Err.Number) & ":" & Err.Description & extradebugl
  LogOnFile "errors.txt", "# internal ID " & idConnection & _
   " lost connection at DoMainLoopLogin, line " & CStr(debugloginline) & " with error #" & CStr(Err.Number) & ":" & Err.Description & extradebugl
  Connected(idConnection) = False
  DoingMainLoopLogin(idConnection) = False
  DoEvents
End Sub

Private Sub SckServerGame_Close(Index As Integer)
  ' game server closes
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  Dim rwait As Long
  If TibiaVersionLong >= 841 Then
  GameConnected(Index) = False
  End If

    If ReconnectionStage(Index) = 0 Then
        If frmHardcoreCheats.chkAutorelog.value = 1 Then
            If logoutAllowed(Index) < GetTickCount() Then ' not allowed by player
                sckServerGame(Index).Close
                StartReconnection Index
                Exit Sub
            End If
        End If

        sckServerGame(Index).Close
        If TibiaVersionLong >= 841 Then
            DoEvents
            rwait = randomNumberBetween(500, 700)
            wait (rwait)
        End If
        If frmRunemaker.chkCloseSound.value = 1 Then
            ChangePlayTheDangerSound True
        End If
        If logoutAllowed(Index) >= GetTickCount() Then ' allowed by player
            txtPackets.Text = txtPackets.Text & vbCrLf & "#gameserver" & Index & " closed (disconnected by user logout)#"
            DoCloseActions Index
            txtPackets.Text = txtPackets.Text & vbCrLf & "(disabling the alarm because it was a desired logout)"
            ChangePlayTheDangerSound False
        Else
            txtPackets.Text = txtPackets.Text & vbCrLf & "#gameserver" & Index & " closed (disconnected by server)#"
            DoCloseActions Index
        End If
'      wait (500) ' avoid fast change of gameserver
        sckClientGame(Index).Close 'close his brother client
        If TibiaVersionLong >= 841 Then
            DoEvents
        End If
    Else
        sckServerGame(Index).Close
    End If
    Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckServerGame_Close(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
  DoCloseActions Index
End Sub


Private Sub SckServerGame_Connect(Index As Integer)
  ' game server connects
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  setsockopt sckServerGame(Index).SocketHandle, IPPROTO_TCP, TCP_NODELAY, 1, 4
'  If TibiaVersionLong >= 841 Then
    'Debug.Print "servergame (" & Index & ") connected to " & sckServerGame(Index).RemoteHostIP & ":" & sckServerGame(Index).RemotePort
'  End If
  'Debug.Print sckServerGame(Index).LocalPort
  lastPing(Index) = GetTickCount()
  If ReconnectionStage(Index) = 0 Then
    txtPackets.Text = txtPackets.Text & vbCrLf & "#gameserver" & Index & " connected#"
  Else
    ReconnectionStage(Index) = 2
    sentFirstPacket(Index) = False
    sentWelcome(Index) = False
    GameConnected(Index) = True
    frmMain.sckServerGame(Index).SendData ReconnectionPacket(Index).packet
    DoEvents
  End If
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error during SckServerGame_Connect(" & Index & ") Number: " & Err.Number & " Description: " & Err.Description & " Source: " & Err.Source
End Sub



Private Sub SckServerGame_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  ' data arrives from game server
  Dim rawpacket() As Byte
  Dim newSizeBuffer As Long
  Dim i As Long
  Dim j As Long
  Dim startB As Long
  Dim endB As Long
  Dim iniB As Long
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  

  ' Get it
  sckServerGame(Index).GetData rawpacket, vbArray + vbByte
  If UBound(rawpacket) < 0 Then
    Debug.Print "Warning: Connection probably lost"
    Exit Sub
  End If
  
  'LogOnFile "weird.txt", showAsStr2(rawpacket, 0)
  'MsgBox "continue2"
  If IgnoreServer(Index) = True Then
    Exit Sub
  End If
  
  'Exit Sub ' Uncomment to debug login packets
  
  lastPing(Index) = GetTickCount()
  ' Store in buffer
  If Index > 0 Then
  #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(Index) & ") NEW RAWPACKET:"
    LogOnFile "bufferLog.txt", showAsStr2(rawpacket, 0)
  #End If
  'Debug.Print "0<< " & frmMain.showAsStr(rawpacket, True)
  iniB = ConnectionBuffer(Index).numbytes ' save initial bytes of buffer
  ' enlarge buffer if needed
  If (UBound(rawpacket) + 1) > ((UBound(ConnectionBuffer(Index).packet) + 1) - ConnectionBuffer(Index).numbytes) Then
    newSizeBuffer = ConnectionBuffer(Index).numbytes + UBound(rawpacket)
    ReDim Preserve ConnectionBuffer(Index).packet(newSizeBuffer)
    
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "BUFFER WAS RESIZED TO " & CStr(newSizeBuffer)
    #End If
  End If
  startB = iniB
  endB = startB + UBound(rawpacket)
  ConnectionBuffer(Index).numbytes = iniB + 1 + UBound(rawpacket)

  RtlMoveMemory ConnectionBuffer(Index).packet(startB), rawpacket(0), (OptCte4 * (endB - startB + 1))
  'j = 0
  'For i = startB To endB
  '  ConnectionBuffer(index).packet(i) = rawpacket(j)
  '  j = j + 1
  'Next i
  #If BufferDebug Then
  LogOnFile "bufferLog.txt", "(" & CStr(Index) & ") USEFULL BUFFER ARE FIRST " & CStr(ConnectionBuffer(Index).numbytes) & " BYTES OF THIS :"
  LogOnFile "bufferLog.txt", showAsStr2(ConnectionBuffer(Index).packet, 0)
  #End If
  If DoingMainLoop(Index) = False Then ' if not doing main loop right now, then start it
    DoMainLoop Index
  End If
  End If
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & Index & " lost connection at SckServerGame_DataArrival #"
  frmMain.DoCloseActions Index
  DoEvents
End Sub

Public Sub DoMainLoop(idConnection As Integer)
  ' Here the buffer of data we got from server is processed
  Dim startB As Long
  Dim lastB As Long
  Dim longPacket As Long
  Dim lPminusOne As Long
  Dim i As Long
  Dim packet() As Byte
  Dim amLeft As Long
  Dim tmpV As Long
  Dim lRes As Integer
  Dim timeOut As Long
  Dim withHeaderL As Long
  Dim withHeaderS As Long
  Dim nBytes As Long
  Dim rawpacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim ignoreConnected As Boolean
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  ignoreConnected = False
  'INITIAL CHECKS (only once per call)
  DoingMainLoop(idConnection) = True
  startB = 0
  If ConnectionBuffer(idConnection).numbytes < 2 Then
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") not even 2 bytes at start..."
    #End If
    DoingMainLoop(idConnection) = False
    Exit Sub ' not even 2 bytes at start ...
  End If
  longPacket = GetTheLong(ConnectionBuffer(idConnection).packet(0), ConnectionBuffer(idConnection).packet(1))
  If longPacket > ((ConnectionBuffer(idConnection).numbytes) - 2) Then
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") no complete packet at start..."
    #End If
    DoingMainLoop(idConnection) = False
    Exit Sub ' no complete packet at start...
  End If
  startB = 2
  'extract 1 complete packet
  lPminusOne = longPacket - 1
  lastB = startB + lPminusOne
nextLoop:
  withHeaderL = lPminusOne + 2
  withHeaderS = startB - 2
  nBytes = withHeaderL + 1
  
  ' decipher it
  If ((UseCrackd = True) And (NeedToIgnoreFirstGamePacket(idConnection) = False)) Then

        ReDim rawpacket(withHeaderL)
        RtlMoveMemory rawpacket(0), ConnectionBuffer(idConnection).packet(withHeaderS), (OptCte4 * nBytes)

    
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") EXTRACTING 1 COMPLETE PACKET:"
    LogOnFile "bufferLog.txt", showAsStr2(rawpacket, 0)
    #End If

     If TibiaVersionLong < 830 Then
        pres = DecipherTibiaProtected(rawpacket(0), packetKey(idConnection).Key(0), UBound(rawpacket), UBound(packetKey(idConnection).Key))
    Else
        pres = DecipherTibiaProtectedSP(rawpacket(0), packetKey(idConnection).Key(0), UBound(rawpacket), UBound(packetKey(idConnection).Key))
    End If
    If (pres < 0) Then
        GiveCrackdDllErrorMessage pres, rawpacket, packetKey(idConnection).Key, UBound(rawpacket), UBound(packetKey(idConnection).Key), 6
        Exit Sub
    End If
    If TibiaVersionLong < 830 Then
        hbytes = GetTheLong(rawpacket(2), rawpacket(3))
        ReDim packet(hbytes + 1)
        RtlMoveMemory packet(0), rawpacket(2), (hbytes + 2)
    Else
        hbytes = GetTheLong(rawpacket(6), rawpacket(7)) ' format: 2x SIZE, 4xCRC , 2xSUBSIZE, PACKET,TRASH SO bytes of (subsize+packet) BECOMES MULTIPLIER OF 8
        ReDim packet(hbytes + 1)
        RtlMoveMemory packet(0), rawpacket(6), (hbytes + 2)
    End If
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") DECIPHERED IT:"
    LogOnFile "bufferLog.txt", showAsStr2(packet, 0)
    #End If
    
  ElseIf (NeedToIgnoreFirstGamePacket(idConnection) = True) Then
    NeedToIgnoreFirstGamePacket(idConnection) = False
    If cteDebugConEvents = True Then
         LogConEvent "N = False"
    End If
   
    
    ReDim packet(withHeaderL)
    RtlMoveMemory packet(0), ConnectionBuffer(idConnection).packet(withHeaderS), (OptCte4 * nBytes)
    If cteDebugConEvents = True Then
         LogConEvent "Initial packet from gameserver: " & frmMain.showAsStr2(packet, 0)
    End If
    If TibiaVersionLong >= 1099 Then
        If (UBound(packet) > 8) Then
            If packet(8) = &H1F Then
              If cteDebugConEvents = True Then
                LogConEvent "OK: Good initial packet"
              End If
            Else
              If cteDebugConEvents = True Then
                LogConEvent "WARNING: Unexpected initial packet format. Trying experimental solution"
                LogConEvent "N = True"
              End If
              NeedToIgnoreFirstGamePacket(idConnection) = True
            End If
        Else
           If cteDebugConEvents = True Then
              LogConEvent "WARNING: Too short packet!! Ubound = " & CStr(UBound(packet))
              LogConEvent "N = True"
           End If
           NeedToIgnoreFirstGamePacket(idConnection) = True
        End If
    End If
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") EXTRACTING 1 COMPLETE PACKET (FIRST):"
    LogOnFile "bufferLog.txt", showAsStr2(packet, 0)
    #End If
      If chkLogPackets.value = 1 Then
        LogLine "GAMESERVER" & idConnection & ":"
        LogPacket packet
        txtPackets.Text = txtPackets.Text & vbCrLf & "GAMESERVER" & idConnection & "<(1st packet)" & showAsStr2(packet, 0)
        txtPackets.SelStart = Len(txtPackets.Text)
      End If
      UnifiedSendToClientGame idConnection, packet, True
        ignoreConnected = True
      GoTo nextP
  Else
    ReDim packet(withHeaderL)
    RtlMoveMemory packet(0), ConnectionBuffer(idConnection).packet(withHeaderS), (OptCte4 * nBytes)
    
    #If BufferDebug Then
    LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") EXTRACTING 1 COMPLETE PACKET:"
    LogOnFile "bufferLog.txt", showAsStr2(packet, 0)
    #End If
  End If
  

      If chkLogPackets.value = 1 Then
        LogLine "GAMESERVER" & idConnection & ":"
        LogPacket packet
        txtPackets.Text = txtPackets.Text & vbCrLf & "GAMESERVER" & idConnection & "<" & showAsStr2(packet, 0)
        txtPackets.SelStart = Len(txtPackets.Text)
      End If
      If frmHardcoreCheats.chkApplyCheats.value = 1 Then
        lRes = LearnFromServer(packet, idConnection)
        If lRes = 1 Then ' Hardcore cheats require skiping this packet
           GoTo nextP
        ElseIf lRes = 3 Then ' Hardcore cheats require losing connection
           sckServerGame(idConnection).Close
           sckClientGame(idConnection).Close
           MustCheckFirstClientPacket(idConnection) = True
           sentFirstPacket(idConnection) = False
           IDstring(idConnection) = ""
           CharacterName(idConnection) = ""
           GameConnected(idConnection) = False
           ConnectionBuffer(idConnection).numbytes = 0
           Exit Sub
        End If
      End If
      timeOut = GetTickCount() + 30000
      While ((GameConnected(idConnection) = True) And (sckClientGame(idConnection).State <> sckConnected))
        If GetTickCount() >= timeOut Then
          'frmMain.DoCloseActions idconnection
          'frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "TIMEOUT(at gameserver) for ID " & CStr(idconnection)
          Exit Sub
        End If
        If sckClientGame(idConnection).State = sckClosed Then
          GameConnected(idConnection) = False
        End If
        DoEvents 'wait
      Wend
      If GameConnected(idConnection) = True And sckClientGame(idConnection).State = sckConnected Then
        UnifiedSendToClientGame idConnection, packet
        If lRes = 2 Then
          sentFirstPacket(idConnection) = True
        End If
        DoEvents
      End If
nextP:
     If ignoreConnected = False Then
     If GameConnected(idConnection) = False Then
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     End If
     ' move pointer
     startB = startB + longPacket
     ' if no complete packet left, move residue to start and end
     If startB = ConnectionBuffer(idConnection).numbytes Then
       ' buffer is now empty
       ConnectionBuffer(idConnection).numbytes = 0
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     If (startB + 1) = ConnectionBuffer(idConnection).numbytes Then
       ' a single byte left
       #If BufferDebug Then
       LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") a single byte left at the end..."
       #End If
       ConnectionBuffer(idConnection).numbytes = 1
       ConnectionBuffer(idConnection).packet(0) = ConnectionBuffer(idConnection).packet(startB)
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     If (startB + 2) = ConnectionBuffer(idConnection).numbytes Then
       ' two bytes left
       #If BufferDebug Then
       LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") a pair of bytes left at the end..."
       #End If
       ConnectionBuffer(idConnection).numbytes = 2
       ConnectionBuffer(idConnection).packet(0) = ConnectionBuffer(idConnection).packet(startB)
       ConnectionBuffer(idConnection).packet(1) = ConnectionBuffer(idConnection).packet(startB + 1)
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     longPacket = GetTheLong(ConnectionBuffer(idConnection).packet(startB), ConnectionBuffer(idConnection).packet(startB + 1))
     lPminusOne = longPacket - 1
     startB = startB + 2
     lastB = startB + lPminusOne
     If (startB + longPacket) > ConnectionBuffer(idConnection).numbytes Then
       ' not complete packets left - save rest in start of buffer
       startB = startB - 2
       tmpV = (ConnectionBuffer(idConnection).numbytes) - startB
       amLeft = tmpV - 1
       ConnectionBuffer(idConnection).numbytes = tmpV
       #If BufferDebug Then
       LogOnFile "bufferLog.txt", "(" & CStr(idConnection) & ") " & CStr(tmpV) & " bytes left at the end..."
       #End If
       
       RtlMoveMemory ConnectionBuffer(idConnection).packet(0), ConnectionBuffer(idConnection).packet(startB), (OptCte4 * tmpV)
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     If ignoreConnected = False Then
     If GameConnected(idConnection) = False Then
       DoingMainLoop(idConnection) = False
       Exit Sub
     End If
     End If
     GoTo nextLoop ' loop until buffer have no complete packets
errclose:
  DoingMainLoop(idConnection) = False
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "# ID" & idConnection & " lost connection at DoMainLoop # :" & CStr(Err.Number) & ":" & Err.Description
  LogOnFile "errors.txt", "# ID" & idConnection & " lost connection at DoMainLoop # :" & CStr(Err.Number) & ":" & Err.Description
  frmMain.DoCloseActions idConnection
  DoEvents
End Sub



Private Sub Timer1_Timer()
  #If FinalMode Then
  On Error GoTo endST
  #End If
  Dim numWindows As Long
  Dim numInt As Integer
  Dim pok As Boolean
  Dim i As Long
  Dim gtc As Long
'        If VarProtection1 <> 1 Then
'            End
'        End If
'        If VarProtection2 <> 2 Then
'            End
'        End If
'        If VarProtection3 <> 3 Then
'            End
'        End If
'        If VarProtection4 <> 4 Then
'            End
'        End If
'        If VarProtection5 <> 5 Then
'            End
'        End If
'        If VarProtection6 <> 6 Then
'            End
'        End If
'        If VarProtection7 <> 7 Then
'            End
'        End If

  If (TibiaVersionLong >= 1100) Then
    Exit Sub
  Else
    gtc = GetTickCount()
    numWindows = CountTibiaWindows()
    If numWindows > LastNumTibiaClients Then
      If frmMain.chckMemoryIP.value = 1 Then
        ' modify IPs in all tibia windows so they connect to this program at localhost
        ' and port
        ModifyTibiaIPs
        If frmMain.TrueServer3.value = True Then
          ModifyTibiaRSAs
        End If
      End If
      ' modify CPU priority of tibia
      pok = UpdateTibiaPriority()
    End If
    LastNumTibiaClients = numWindows
  End If
endST:
End Sub



Private Sub TimerTibiaQ_Timer()
  If (TibiaVersionLong >= 1100) Then
    RedirectAllServersHere
  End If
End Sub

Private Sub timeToSpam_Timer()
  Dim i As Integer
  Dim aRes As Long
  Dim gtc As Long
  Dim conds As Boolean
  #If FinalMode Then
    On Error GoTo gotErr
  #End If
  i = 0
  gtc = GetTickCount()
  For i = 1 To MAXCLIENTS
    If (ReconnectionStage(i) > 0) Then
        If (gtc >= nextReconnectionRetry(i)) Then
            If (ReconnectionStage(i) = 3) Or (ReconnectionStage(i) = 10) Then
                ReconnectionStage(i) = 0
                reconnectionRetryCount(i) = 0
            Else
                conds = (frmMain.sckClientGame(i).State = sckClosed)
                ReconnectionStage(i) = 0
                If (conds = True) Then
                    reconnectionRetryCount(i) = 0
                Else
                    StartReconnection i
                    nextReconnectionRetry(i) = gtc + RETRYDELAY
                End If
            End If
        End If
    Else
        reconnectionRetryCount(i) = 0
    End If
    If (GameConnected(i) = True) And (GotPacketWarning(i) = False) And (posSpamActivated(i) = True) Then
      aRes = SendChannelMessage(i, "@" & myX(i) & "," & myY(i) & "," & myZ(i), _
       posSpamChannelB1(i), posSpamChannelB2(i))
      DoEvents
    End If
  Next i
  Exit Sub
gotErr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error while trying to send position client " & CStr(i) & " : " & Err.Description
End Sub

Private Sub TrueServer1_Click()
  ' click in option True server
  If TrueServer1.value = True Then
    lblEnterOtherComputerIP.enabled = False
    lblLoginPort.enabled = False
    lblGamePort.enabled = False
    ForwardGameTo.enabled = False
    txtServerLoginP.enabled = False
    txtServerGameP.enabled = False
    lblWarning.Visible = False
    lblGamePort.Visible = False
    lblGamePort.enabled = False
    txtServerGameP.Visible = False
    txtServerGameP.enabled = False
  End If
  If WARNING_USING_OTSERVER_RSA = True Then
    closeAllTibiaClientsExcept -1
    frmMenu.Caption = "Tibia clients need reload"
  End If
End Sub

Private Sub TrueServer2_Click()
  ' click in option other proxy
  If TrueServer2.value = True Then
    lblEnterOtherComputerIP.enabled = True
    lblLoginPort.enabled = True
    lblGamePort.Visible = True
    lblGamePort.enabled = True
    ForwardGameTo.enabled = True
    txtServerLoginP.Visible = True
    txtServerLoginP.enabled = True
    txtServerGameP.Visible = True
    txtServerGameP.enabled = True
    lblWarning.Visible = True
    lblWarning.enabled = True
    lblEnterOtherComputerIP.Caption = "Enter other proxy IP ..."
  End If
End Sub

Private Sub TrueServer3_Click()
  ' click in option OT server
  If TrueServer3.value = True Then
    lblEnterOtherComputerIP.enabled = True
    lblLoginPort.enabled = True
    lblGamePort.Visible = False
    lblGamePort.enabled = False
    ForwardGameTo.enabled = True
    txtServerLoginP.enabled = True
    txtServerGameP.Visible = False
    txtServerGameP.enabled = False
    lblWarning.Visible = False
    lblWarning.enabled = False
    lblEnterOtherComputerIP.Caption = "Enter OT server IP ..."
  End If
End Sub

Private Sub txtClientGameP_Validate(Cancel As Boolean)
  ' change in game port
  Dim newP As Long
  newP = CLng(txtClientGameP.Text)
  If TibiaVersionLong >= 841 Then
    newP = 0
  End If
  If newP >= 0 Then
    sckClientGame(0).Close
    sckClientGame(0).LocalPort = newP
    sckClientGame(0).Listen
    If chckMemoryIP.value = 1 Then
      LastNumTibiaClients = 0 ' this will force change IPs now
    End If
  End If

End Sub

Private Sub txtClientLoginP_Validate(Cancel As Boolean)
  On Error GoTo gotErr
  ' change in login port
  Dim newP As Long
  Dim failedline As String
  failedline = ""
  'failedline = failedline & vbCrLf & "newP = CLng(""" & txtClientLoginP.Text & """)"
  newP = CLng(txtClientLoginP.Text)
  If TibiaVersionLong >= 841 Then
    newP = 0
  End If

  'failedline = failedline & vbCrLf & "If newP > 0 Then"
  If newP >= 0 Then
    failedline = "' Building the bind at " & SckClient(0).LocalIP & " ..."
    failedline = failedline & vbCrLf & "SckClient(0).Close"
    SckClient(0).Close
    DoEvents
    failedline = failedline & vbCrLf & "SckClient(0).LocalPort = " & CStr(newP)
    SckClient(0).LocalPort = newP
    failedline = failedline & vbCrLf & "SckClient(0).Listen"
    SckClient(0).Listen
    failedline = failedline & vbCrLf & "Connected"
  End If
  Exit Sub
gotErr:
  MsgBox "Sorry, Blackd Proxy was not able to initialize..." & vbCrLf & "Possible reasons:" & vbCrLf & _
  " - Blackd Proxy already open" & vbCrLf & _
  " - Bugged Tibia client blocking connections (try closing all Tibia clients first)" & vbCrLf & _
  " - Firewall blocking binds to port " & CStr(newP) & vbCrLf & _
  " Details:" & vbCrLf & _
  " - Error number " & Err.Number & vbCrLf & _
  " - Error description: " & Err.Description & vbCrLf & _
  " - Location: tClientLoginP_Validate" & vbCrLf & _
  " - Trace: " & vbCrLf & failedline, vbOKOnly + vbCritical, "Critical error"
  End
End Sub

Private Sub txtPackets_Change()
  ' change in txtPackets
  While Len(txtPackets.Text) > CLng(txtMaxChar.Text)
    If frmMain.LogFull1.value = True Then
      txtPackets.Text = ""
    ElseIf frmMain.LogFull2.value = True Then
      DeleteFirstLine
    Else
      LogOnFile txtLogFile.Text, txtPackets.Text
      txtPackets.Text = ""
    End If
    DoEvents
  Wend
  txtPackets.SelStart = Len(txtPackets.Text)
End Sub


            
Public Function showAsStr2(ByRef packet() As Byte, hexad As Byte, Optional limitUbound As Long = 0) As String
  ' show a packet as string
  ' hexad:
  ' 0 -> hex with header
  ' 1 -> ascii with header
  ' 2 -> hex without header
  ' limitUbound: return result as if packet only had that ubound
  Dim i As Long
  Dim itemsNumber As Long
  Dim strShow As String
  itemsNumber = UBound(packet)
  If limitUbound > 0 Then
    If limitUbound < itemsNumber Then
        itemsNumber = limitUbound
    End If
  End If
  
  ' depending hexad parameter, show it as hex or as ascii
  If hexad = 0 Then
     strShow = "( hex ) "
  ElseIf hexad = 1 Then
     strShow = "( ascii ) "
  Else
     strShow = ""
  End If
  For i = 0 To itemsNumber
   If hexad = 1 Then
     strShow = strShow & Chr(packet(i))
   Else
     strShow = strShow & GoodHex(packet(i)) & " "
   End If
  Next i
  showAsStr2 = strShow
End Function

Private Sub LogPacket(ByRef packet() As Byte)
  ' logs a packet
  Dim i As Long
  Dim j As Long
  Dim co As Long
  Dim currentLine As String
  Dim useRow As Long
  Dim byteStart As Long
  Dim byteEnd As Long
  Dim bytesLeft As Long
  Dim convHex As String
  Dim MaxLogLines As Long
  MaxLogLines = CLng(txtMaxLines.Text)
  byteStart = 0
  bytesLeft = UBound(packet) + 1
  Do
    If gridLog.Rows = MaxLogLines Then
      If frmMain.LogFull1.value = True Or frmMain.LogFull3.value = True Then
        InitGridLog
        gridLog.Rows = 2
        useRow = 1
      Else
        For i = 0 To MaxLogLines - 2
          For j = 0 To 20
            gridLog.TextMatrix(i, j) = gridLog.TextMatrix(i + 1, j)
          Next j
        Next i
        useRow = MaxLogLines - 1
      End If
    Else
      gridLog.Rows = gridLog.Rows + 1
      useRow = gridLog.Rows - 1
    End If
   
    If bytesLeft < 11 Then
      byteEnd = byteStart + bytesLeft - 1
      bytesLeft = 0
    Else
      byteEnd = byteStart + 9
      bytesLeft = bytesLeft - 10
    End If
    co = 0
    For i = byteStart To byteEnd
      convHex = GoodHex(packet(i))
      gridLog.TextMatrix(useRow, co) = convHex
      gridLog.TextMatrix(useRow, co + 11) = Chr(packet(i))
      co = co + 1
    Next i
    For i = byteEnd + 1 To byteStart + 9
      gridLog.TextMatrix(useRow, co) = ""
      gridLog.TextMatrix(useRow, co + 11) = ""
      co = co + 1
    Next i
    gridLog.TextMatrix(useRow, 10) = "#"
    byteStart = byteStart + 10
  Loop Until bytesLeft = 0
  
  If gridLog.Rows > 10 Then 'scrolls down
    gridLog.TopRow = gridLog.Rows - 10
  End If
End Sub
Private Sub LogLine(strLine As String)
  ' log a string line
  Dim i As Integer
  Dim j As Integer
  Dim currentLine As String
  Dim useRow As Integer
  Dim MaxLogLines As Long
  MaxLogLines = CLng(txtMaxLines.Text)
  Do
    If gridLog.Rows = MaxLogLines Then
       If frmMain.LogFull1.value = True Or frmMain.LogFull3.value = True Then
        InitGridLog
        gridLog.Rows = 2
        useRow = 1
      Else
        For i = 0 To MaxLogLines - 2
          For j = 0 To 20
            gridLog.TextMatrix(i, j) = gridLog.TextMatrix(i + 1, j)
          Next j
        Next i
        useRow = MaxLogLines - 1
      End If
    Else
      gridLog.Rows = gridLog.Rows + 1
      useRow = gridLog.Rows - 1
    End If
   
    If Len(strLine) < 22 Then
      currentLine = strLine
      strLine = ""
    Else
      currentLine = Left(strLine, 21)
      strLine = Right(strLine, Len(strLine) - 21)
    End If
    For i = 0 To Len(currentLine) - 1
      gridLog.TextMatrix(useRow, i) = Mid(currentLine, i + 1, 1)
    Next i
    For i = Len(currentLine) To 20
      gridLog.TextMatrix(useRow, i) = ""
    Next i
  Loop Until Len(strLine) = 0
  If gridLog.Rows > 10 Then 'scrolls down
    gridLog.TopRow = gridLog.Rows - 10
  End If
End Sub

Private Sub DeleteFirstLine()
  ' deletes first line - seems to be slow
  Dim endFirstLine As Long
  endFirstLine = InStr(1, txtPackets.Text, vbCrLf, vbTextCompare)
  If endFirstLine = -1 Then
    txtPackets.Text = "Error"
  Else
    txtPackets.SelStart = 0
    txtPackets.SelLength = endFirstLine
    txtPackets.SelText = "" 'delete first line
  End If
End Sub



Public Sub txtTibiaPath_Validate(Cancel As Boolean)

  Dim res As String
  
  res = ValidateTibiaPath(txtTibiaPath.Text)
  TibiaPath = res
  txtTibiaPath.Text = res
  

End Sub


Public Function showAsStr(ByRef packet() As Byte, hexad As Boolean) As String
  'legacy function
  If hexad = True Then
    showAsStr = showAsStr2(packet, 0)
  Else
    showAsStr = showAsStr2(packet, 1)
  End If
End Function

Public Function showAsStr3(ByRef packet() As Byte, ByVal hexad As Boolean, ByVal first As Long, ByVal last As Long) As String
  Dim i As Long
  Dim strShow As String
  Dim itemsNumber As Long
  
  itemsNumber = UBound(packet)
  
  ' depending hexad parameter, show it as hex or as ascii
  If hexad = True Then
     strShow = "( hex ) "
  ElseIf hexad = False Then
     strShow = "( ascii ) "
  Else
     strShow = ""
  End If
  If last > itemsNumber Then
     last = itemsNumber
  End If
  For i = first To last
   If hexad = False Then
     strShow = strShow & Chr(packet(i))
   Else
     strShow = strShow & GoodHex(packet(i)) & " "
   End If
  Next i
  showAsStr3 = strShow
End Function

Private Function GetWITHTIBIADATtrivial() As String
    Dim res As String
    If ((Right$(TibiaExePath, 1) = "\") Or (Right$(TibiaExePath, 1) = "/")) Then
       res = TibiaExePath & "Tibia.dat"
    Else
       res = TibiaExePath & "\Tibia.dat"
    End If
    If (TibiaVersionLong >= 1100) Then
       res = res & "packages\Tibia\assets\"
    Else
       res = res & "Tibia.dat"
    End If
    GetWITHTIBIADATtrivial = res
End Function
Private Function GetWITHTIBIADAT() As String
    'On Error GoTo goterr
    Dim fso As New Scripting.FileSystemObject
    Dim fol As Scripting.Folder
    Dim fil As Scripting.File
    Dim thename As String
    Dim usethisfolder As String
    Dim res As String
    Dim foundit As Boolean
    Dim gotTibiaDat As Boolean
    Dim lastDatFound As String
    foundit = False
    If ((Right$(TibiaExePath, 1) = "\") Or (Right$(TibiaExePath, 1) = "/")) Then
       usethisfolder = TibiaExePath
    Else
       usethisfolder = TibiaExePath & "\"
    End If
    If (TibiaVersionLong >= 1100) Then
       ' For ots still using old tibia.dat you only need to place your old tibia.dat in the main Tibia folder.
       ' Then Blackd Proxy will read it instead, using old parser.
       If fso.FileExists(usethisfolder & "tibia.dat") = True Then
          GetWITHTIBIADAT = usethisfolder & "tibia.dat"
          Exit Function
       End If
       usethisfolder = usethisfolder & "packages\Tibia\assets\"
    End If
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(usethisfolder) = False Then
        MsgBox "The Tibia folder you selected does not exist:" & vbCrLf & _
        usethisfolder & vbCrLf & _
        "Reload Blackd Proxy and select a correct folder", vbCritical + vbOKOnly, "Blackd Proxy - Config Error"
        SaveConfigWizard True
        End
    End If
    Set fol = fso.GetFolder(usethisfolder)
    gotTibiaDat = False
    For Each fil In fol.Files
        thename = fil.name
        If Len(thename) > 4 Then
            If LCase(Right$(thename, 4)) = ".dat" Then ' allows custom dats of some ot servers
                If Not (LCase(thename) = "unins000.dat") Then ' don't take this strange .dat
                    foundit = True
                    lastDatFound = thename
                    If LCase(thename) = "tibia.dat" Then
                        Exit For
                    End If
                    If (TibiaVersionLong >= 1100) Then
                        GetWITHTIBIADAT = usethisfolder
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    If foundit = False Then
    thename = "Tibia.dat"
    End If
    Set fil = Nothing
    Set fol = Nothing
    Set fso = Nothing
    res = usethisfolder & lastDatFound
    GetWITHTIBIADAT = res
    Exit Function
gotErr:
    GetWITHTIBIADAT = GetWITHTIBIADATtrivial()
End Function




Public Function HTTPGet(uri As String) As String
Dim parsedURL As url
Dim strHeaders As String
Dim strMethod As String
Dim strHTTP As String
Dim strPostData As String
'currently hardcoded to only 1 (0), i guess theoretically we could support several requests at once...
'.........but i don't wanna figure that out now x.x
ReDim Preserve HTTPGetConnected(0 To 1)
ReDim Preserve HTTPGetResponseBuffer(0 To 1)

strHeaders = "Connection: close" & vbCrLf

strMethod = "GET"
parsedURL = ExtractUrl(uri)
If parsedURL.Host = vbNullString Then
    Err.Raise vbCritical, uri, "Invalid Host"
End If
    'hmm, support multiple connections?
    SckHttpGet(0).Protocol = sckTCPProtocol
    SckHttpGet(0).RemoteHost = parsedURL.Host
    
    parsedURL.Scheme = "http" 'bleh. supporting anything else would be a lot of work.
    If parsedURL.Scheme = "http" Then
        If parsedURL.port > 0 Then
            SckHttpGet(0).RemotePort = parsedURL.port
        Else
            SckHttpGet(0).RemotePort = 80
        End If
    ElseIf parsedURL.Scheme = vbNullString Then
        SckHttpGet(0).RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
    strHTTP = strMethod & " " & parsedURL.uri & "?" & parsedURL.Query & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & parsedURL.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData
    HTTPGetConnected(0) = False
    SckHttpGet(0).Connect
    ' wait for a connection
    While Not HTTPGetConnected(0)
        DoEvents
    Wend
    'connected
    SckHttpGet(0).SendData strHTTP
    ' wait for response
    While HTTPGetConnected(0)
        DoEvents
    Wend
    'have our response, connection is closed.
    
    HTTPGet = HTTPGetResponseBuffer(0)
    HTTPGet = Mid$(HTTPGet, InStr(HTTPGet, vbCrLf & vbCrLf) + 4) 'trim out HTTP headers.
    HTTPGetResponseBuffer(0) = vbNullString 'clear memory
End Function




Private Sub SckHttpGet_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'??? now what?
Err.Raise Number, "SckHttpGet", Description & " source: " & Source
HTTPGetConnected(Index) = Not HTTPGetConnected(Index) 'flip it.
End Sub

Private Sub SckHttpGet_Connect(Index As Integer)
    HTTPGetConnected(Index) = True
End Sub

Private Sub SckHttpGet_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strResponse As String
    SckHttpGet(Index).GetData strResponse, vbString, bytesTotal
    
'    strResponse = FormatLineEndings(strResponse)
    
    ' we append this to the string becuase data arrives
    ' in multiple packets
    HTTPGetResponseBuffer(Index) = HTTPGetResponseBuffer(Index) & strResponse
End Sub
Private Sub SckHttpGet_Close(Index As Integer)
    HTTPGetConnected(Index) = False
    SckHttpGet(Index).Close 'DO NOT REMOVE THIS (seemingly useless) LINE. SHIT WILL BREAK AFTER 2 REQUESTS WITHOUT THIS LINE!!
End Sub

