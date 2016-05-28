VERSION 5.00
Begin VB.Form frmOld 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Proxy"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   Icon            =   "frmOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOldMenu 
      BackColor       =   &H00000000&
      Caption         =   "Set this Menu default. Next time Blackd is open this will be the main menu."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   5775
   End
   Begin VB.CommandButton cmdVIPsupport 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to VIP support page"
      Height          =   315
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdHardcoreCheats 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      Picture         =   "frmOld.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdRunemaker 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1560
      Picture         =   "frmOld.frx":17BD
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCavebot 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      Picture         =   "frmOld.frx":2AE8
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheats 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      Picture         =   "frmOld.frx":3C37
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogs 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1560
      Picture         =   "frmOld.frx":4D44
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutorial 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      Picture         =   "frmOld.frx":5A85
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdStopAlarm 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      Picture         =   "frmOld.frx":6995
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdHotkeys 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   4440
      Picture         =   "frmOld.frx":79CD
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdvanced 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1560
      Picture         =   "frmOld.frx":8A07
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdEvents 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5880
      Picture         =   "frmOld.frx":9A5F
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdWarbot 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5880
      Picture         =   "frmOld.frx":A8B8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTrainer 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   4440
      Picture         =   "frmOld.frx":BE59
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdMagebomb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   975
      Left            =   7320
      Picture         =   "frmOld.frx":CD37
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnknownFeature 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   7320
      Picture         =   "frmOld.frx":DC98
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdHPmana 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      Picture         =   "frmOld.frx":F012
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAd 
      BackColor       =   &H00C0C000&
      Height          =   975
      Left            =   4440
      Picture         =   "frmOld.frx":1007F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdStealth 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1560
      Picture         =   "frmOld.frx":10E7C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLaunchTibia 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5880
      Picture         =   "frmOld.frx":124AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdLaunchTibiaMC 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   7320
      Picture         =   "frmOld.frx":15107
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdNews 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      Picture         =   "frmOld.frx":160EE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdBroadcast 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      Picture         =   "frmOld.frx":17DE3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "If you purchased us any gold in the last month, we give you VIP support"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   6600
      TabIndex        =   22
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Official sites:"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblMainSite 
      BackColor       =   &H00000000&
      Caption         =   "www.blackdtools.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblUpdates 
      BackColor       =   &H00000000&
      Caption         =   "[updates]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblForum 
      BackColor       =   &H00000000&
      Caption         =   "[forum]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblAltSite 
      BackColor       =   &H00000000&
      Caption         =   "www.blackdtools.es"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   3840
      Width           =   2175
   End
End
Attribute VB_Name = "frmOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
#Const BlockCavebot = 0
#Const BlockTools = 0
#Const BlockRunemaker = 0
#Const BlockAllCheats = 0
#Const KeepRemote = 0
#Const DoSave = 1
Option Explicit

Private Sub chkOldMenu_Click()
  'custom ng
  ' write ini file
  Dim i As Integer
  Dim strInfo As String
  Dim here As String
  Dim idLoginSP As Long

  
  If configPath = "" Then
    here = myMainConfigINIPath()
  Else
    here = App.path & "\" & configPath & "\config.ini"
  End If
  
  If chkOldMenu.Value = 1 Then
  strInfo = CStr(frmOld.chkOldMenu.Value)
  i = setBlackdINI("OldMenu", "chkOldMenu", strInfo, here)
  Else
  strInfo = CStr(frmOld.chkOldMenu.Value)
  i = setBlackdINI("OldMenu", "chkOldMenu", strInfo, here)
  End If

End Sub

Private Sub cmdAd_Click()
Dim a

  a = ShellExecute(Me.hwnd, "Open", "https://blackdtools.com/worldtrade.php", &O0, &O0, SW_NORMAL)

End Sub

Private Sub cmdAdvanced_Click()
  ' show Advanced form
  frmAdvanced.WindowState = vbNormal
  frmAdvanced.Show
  frmAdvanced.SetFocus
  SetWindowPos frmAdvanced.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdBroadcast_Click()
  frmBroadcast.WindowState = vbNormal
  frmBroadcast.Show
  frmBroadcast.SetFocus
  SetWindowPos frmBroadcast.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

'#Const FinalMode =0
'#Const BlockCavebot = 1
'#Const BlockTools = 1
'#Const BlockRunemaker = 1
'#Const BlockAllCheats = 1
'#Const KeepRemote = 1
'#Const DoSave = 0
Private Sub cmdCavebot_Click()
  ' show cavebot form
  frmCavebot.WindowState = vbNormal
  frmCavebot.Show
  frmCavebot.SetFocus
  SetWindowPos frmCavebot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdCheats_Click()
  ' show Tools form
  frmCheats.WindowState = vbNormal
  frmCheats.Show
  frmCheats.SetFocus
  SetWindowPos frmCheats.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdEvents_Click()
  frmEvents.WindowState = vbNormal
  frmEvents.Show
  frmEvents.SetFocus
  SetWindowPos frmEvents.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdHardcoreCheats_Click()
  ' show Cheats form
  frmHardcoreCheats.WindowState = vbNormal
  frmHardcoreCheats.Show
  frmHardcoreCheats.SetFocus
  SetWindowPos frmHardcoreCheats.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdHotkeys_Click()
  ' show hotkeys
  frmHotkeys.WindowState = vbNormal
  frmHotkeys.Show
  frmHotkeys.SetFocus
  SetWindowPos frmHotkeys.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdHPmana_Click()
  ' show Advanced form
  frmHPmana.WindowState = vbNormal
  frmHPmana.Show
  frmHPmana.SetFocus
  SetWindowPos frmHPmana.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub




Private Sub cmdLaunchTibia_Click()
  ' open Shoot Fruits in default web navigator
  Dim X
  X = ShellExecute(Me.hwnd, "Open", "http://shootfruits.com", &O0, &O0, SW_NORMAL)
  
'    Dim res As String
'    Dim tpath As String
'    tpath = TibiaExePath
'    If tpath = "" Then
'        Label3.Caption = "FILESYSTEM ERROR"
'        Exit Sub
'    End If
'    res = LaunchTibia(tpath, False)
'    If res <> "" Then
'        Label3.Caption = "TIBIA NOT FOUND"
'        Exit Sub
'    End If
End Sub

Private Sub cmdLaunchTibiaMC_Click()
    Dim res As String
    Dim tpath As String
    tpath = TibiaExePath
    If tpath = "" Then
        Label3.Caption = "FILESYSTEM ERROR"
        Exit Sub
    End If
    res = LaunchTibia(tpath, True)
    If res <> "" Then
        Label3.Caption = "TIBIA NOT FOUND"
        Exit Sub
    End If
End Sub


Private Sub cmdLoad_Click()
Dim aRes As Long
Dim louade As String
Dim idConnection As Integer
Dim i As Integer
    For i = 1 To MAXCLIENTS
        idConnection = i
        louade = "exiva load"
        aRes = ExecuteInTibia(louade, idConnection, True)
        aRes = ExecuteInTibia(louade, idConnection, True)
        DoEvents
        Next i
End Sub

Private Sub cmdLogs_Click()
  ' show main form
  frmMain.WindowState = vbNormal
  frmMain.Show
  frmMain.SetFocus
  SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdMagebomb_Click()
  frmMagebomb.WindowState = vbNormal
  frmMagebomb.Show
  frmMagebomb.SetFocus
  SetWindowPos frmMagebomb.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

'Private Sub cmdMagebot_Click()
'    Dim res As String
'    Dim tpath As String
'    Dim tfile As String
'    tpath = MagebotPath
'    tfile = MagebotExe
'    If tpath = "" Then
'        Label3.Caption = "FILESYSTEM ERROR"
'        Exit Sub
'    End If
'    res = LaunchFileNormalWay(tpath, tfile)
'    If res <> "" Then
'        Label3.Caption = "TIBIA NOT FOUND"
'        Exit Sub
'    End If
'End Sub

Private Sub cmdNews_Click()
  frmNews.WindowState = vbNormal
  frmNews.Show
  frmNews.SetFocus
  SetWindowPos frmNews.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdRunemaker_Click()
  ' show Runemaker form
  frmRunemaker.WindowState = vbNormal
  frmRunemaker.Show
  frmRunemaker.SetFocus
  SetWindowPos frmRunemaker.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdSave_Click()
Dim aRes As Long
Dim salve As String
Dim idConnection As Integer
Dim i As Integer
'frmHardcoreCheats.UpdateValues

    For i = 1 To MAXCLIENTS
        idConnection = i
        salve = "exiva save"
        aRes = ExecuteInTibia(salve, idConnection, True)
        DoEvents
        Next i
End Sub

Private Sub cmdStealth_Click()
  frmStealth.WindowState = vbNormal
  frmStealth.Show
  frmStealth.SetFocus
  SetWindowPos frmStealth.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdStopAlarm_Click()
  ' stop alarms
  Dim mcid As Integer
  For mcid = 1 To MAXCLIENTS
    DangerPK(mcid) = False
    DangerGM(mcid) = False
    DangerPlayer(mcid) = False
    LogoutTimeGM(mcid) = 0
    moveRetry(mcid) = 0
    RemoveSpamOrder mcid, 1
    UHRetryCount(mcid) = 0
    logoutAllowed(mcid) = 0
  Next mcid
  ChangePlayTheDangerSound False
End Sub

Private Sub cmdStopAlarmNG_Click()
  ' stop alarms
  Dim mcid As Integer
  For mcid = 1 To MAXCLIENTS
    DangerPK(mcid) = False
    DangerGM(mcid) = False
    DangerPlayer(mcid) = False
    LogoutTimeGM(mcid) = 0
    moveRetry(mcid) = 0
    RemoveSpamOrder mcid, 1
    UHRetryCount(mcid) = 0
    logoutAllowed(mcid) = 0
  Next mcid
  ChangePlayTheDangerSound False
End Sub

Private Sub cmdTrainer_Click()
  frmTrainer.WindowState = vbNormal
  frmTrainer.Show
  frmTrainer.SetFocus
  SetWindowPos frmTrainer.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdTutorial_Click()
  ' open tutorial in default web navigator
  Dim X
  X = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/showthread.php?t=221", &O0, &O0, SW_NORMAL)
End Sub




Private Sub cmdUnknownFeature_Click()
  frmCondEvents.WindowState = vbNormal
  frmCondEvents.Show
  frmCondEvents.SetFocus
  SetWindowPos frmCondEvents.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


Private Sub cmdVIPsupport_Click()
 ' open tutorial in default web navigator
  Dim X
  X = ShellExecute(Me.hwnd, "Open", "https://blackdtools.com/vip.php", &O0, &O0, SW_NORMAL)

End Sub

Private Sub cmdWarbot_Click()
  frmWarbot.WindowState = vbNormal
  frmWarbot.Show
  frmWarbot.SetFocus
  SetWindowPos frmWarbot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub


'Private Sub Command1_Click()
'  Dim tibiaclient As Long
'  Dim res As Long
'  tibiaclient = 0
'  Do
'        tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
'        If tibiaclient = 0 Then
'            Exit Do
'        Else
'            res = ReadCurrentAddress(tibiaclient, adrSelectedCharIndex, -1, True)
'            MsgBox ("SEL CHAR=" & CStr(res))
'        End If
'  Loop
'End Sub


Public Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        frmOld.Hide
    End If
End Sub

Private Sub lblAltSite_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim a
If Button = 1 Then
 a = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.es/index.php", &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lblForum_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim a
If Button = 1 Then
  a = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/forum/index.php", &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lblMainSite_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim a
If Button = 1 Then
 a = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/index.php", &O0, &O0, SW_NORMAL)
End If
End Sub



Private Sub lblUpdates_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim a
If Button = 1 Then
  a = ShellExecute(Me.hwnd, "Open", "http://www.blackdtools.com/freedownloads.php", &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub mPopExit_Click()
  ' exit by tray menu
  Dim btemp As Integer
  btemp = 0
  If confirmedExit = False Then
    frmConfirm.Show
    frmConfirm.SetFocus
    Exit Sub
  End If
  'called when user clicks the popup menu Exit command
End Sub

Private Sub mPopRestore_Click()
  'called when the user clicks the popup menu Restore command
  Dim result As Long
  Me.WindowState = vbNormal
  result = SetForegroundWindow(Me.hwnd)
  Me.Show
End Sub

Private Sub mPopShowTibia_Click()
  SetTibiaClientsVisible True
End Sub

Private Sub mPopHideTibia_Click()
  SetTibiaClientsVisible False
End Sub

Public Sub ApplyLimits()
' If compiling a limited version, then disable and hide some options
Dim save1 As Long
#If BlockRunemaker Then
With frmRunemaker
.UseRightHand.enabled = False
.UseLeftHand.enabled = False
.chkActivate.enabled = False
.chkFood.enabled = False
.chkManaFluid.enabled = False
.chkLogoutDangerAny.enabled = False
.chkLogoutDangerCurrent.enabled = False
.chkLogoutOutRunes.enabled = False
.chkWaste.enabled = False
.chkmsgSound.enabled = False
.chkmsgSound2.enabled = False
.txtAction1.enabled = False
.txtManaAction1.enabled = False
.txtAction2.enabled = False
.txtManaAction2.enabled = False
.txtSoulAction2.enabled = False
.lstFriends.enabled = False
.cmdLoad.enabled = False
.cmdSave.enabled = False
.txtFile.enabled = False
.txtAddFriend.enabled = False
.cmdAddFriend.enabled = False
.cmdRemoveFriend.enabled = False
.ChkDangerSound.Value = 0
.ChkDangerSound.enabled = False
.chkCloseSound.Value = 0
.chkOnDangerSS.Value = 0
.chkCloseSound.enabled = False
.cmdStopAlarm.enabled = False
.cmdApply.enabled = False
.cmdDebug.enabled = False
End With
frmMenu.cmdRunemaker.enabled = False
#End If
#If BlockRunemaker Then
With frmCavebot
.chkEnabled.enabled = False
.chkChangePkHeal.Value = 0
.chkChangePkHeal.enabled = False
End With
frmMenu.cmdCavebot.enabled = False
#End If
#If BlockTools Then
frmCheats.chkInspectTileID.Value = 0
frmCheats.chkInspectTileID.enabled = False
#End If
save1 = frmHardcoreCheats.chkAcceptSDorder.Value
#If BlockAllCheats Then
With frmHardcoreCheats
.txtRemoteLeader.Text = LimitedLeader
.chkLogoutIfDanger.Value = 0
.chkLogoutIfDanger.enabled = False
.chkReveal.Value = 0
.chkReveal.enabled = False
.chkLight.Value = 0
.chkLight.enabled = False
.chkAutoHeal.Value = 0
.chkAutoHeal.enabled = False
.chkAutoVita.Value = 0
.chkAutoVita.enabled = False
.chkAcceptSDorder.Value = 0
.chkAcceptSDorder.enabled = False
.chkColorEffects.Value = 0
.chkColorEffects.enabled = False
.cmdOpenTrueRadar.enabled = False
.cmdUpdateMap.enabled = False
.cmdOpenBackpacks.enabled = False
.chkLogoutIfDanger.Visible = False
.chkReveal.Visible = False
.chkLight.Visible = False
.chkAutoHeal.Visible = False
.chkAutoVita.Visible = False
.chkAcceptSDorder.Visible = False
.chkColorEffects.Visible = False
.cmdOpenTrueRadar.Visible = False
.cmdUpdateMap.Visible = False
.cmdOpenBackpacks.Visible = False
.chkApplyCheats.Visible = False
.cmdReset.Visible = False
.Line3.Visible = False
.scrollLight.Visible = False
.lblLightValue.Visible = False
.scrollHP.Visible = False
.lblHPvalue.Visible = False
.scrollHP2.Visible = False
.lblHPvalue2.Visible = False
.txtOrder.Visible = False
.lblOrder2.Visible = False
.lblRead.Visible = False
.cmbOrderType.Visible = False
.lblOn.Visible = False
.lblLeader.Visible = False
.txtRemoteLeader.Visible = False
.txtCommands.Visible = False
.chkColorEffects.Visible = False
.cmdOpenTrueRadar.Visible = False
.cmdUpdateMap.Visible = False
.chkLockOnMyFloor.Visible = False
.chkOnTop.Visible = False
.cmdOpenBackpacks.Visible = False
.lblChar.Visible = False
.cmbCharacter.Visible = False
.lblYourPos.Visible = False
.lblPosition.Visible = False
.chkManualUpdate.Visible = False
.chkUpdateMs.Visible = False
.chkAutoUpdateMap.Visible = False
.Label1.Visible = False
.lblArraySelected.Visible = False
.cmdMs.Visible = False
.cmdChange.Visible = False
.lblAdvanced.Visible = False
.pushID.Visible = False
.ActionInspect.Visible = False
.ActionMove.Visible = False
.ActionNothing.Visible = False
.ActionPath.Visible = False
.Frame1.Visible = False
.chkRuneAlarm.Value = 0
.chkRuneAlarm.enabled = False
.chkRuneAlarm.Visible = False
.txtAlarmUHs.Text = -1
.txtAlarmUHs.enabled = False
.txtAlarmUHs.Visible = False
End With
frmMenu.cmdHardcoreCheats.enabled = False
#End If
#If KeepRemote Then
With frmHardcoreCheats
.Caption = "Cheats (limited to accept remote orders)"
.lblLeader.Caption = "Only accept order from this leader (locked in this version) :"
.chkAcceptSDorder.Value = save1
.chkAcceptSDorder.enabled = True
.txtRemoteLeader.enabled = False
.chkAcceptSDorder.Visible = True
.txtOrder.Visible = True
.lblOrder2.Visible = True
.lblRead.Visible = True
.cmbOrderType.Visible = True
.lblOn.Visible = True
.lblLeader.Visible = True
.txtRemoteLeader.Visible = True
.chkAcceptSDorder.Top = 100
.txtOrder.Top = 100
.lblOrder2.Top = 100
.lblRead.Top = 340
.cmbOrderType.Top = 340
.lblOn.Top = 340
.lblLeader.Top = 700
.txtRemoteLeader.Top = 680
.txtRemoteLeader.Left = 4500
.txtRemoteLeader.Width = 1000
.Height = 1500
End With
frmMenu.cmdHardcoreCheats.enabled = True
#End If
End Sub



