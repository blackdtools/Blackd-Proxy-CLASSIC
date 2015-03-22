VERSION 5.00
Begin VB.Form frmHardcoreCheats 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheats ..."
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmHardcoreCheats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkProtectedShots 
      BackColor       =   &H00000000&
      Caption         =   "Avoid shoting damage runes if your %hp < AutoRuneHeal %hp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   73
      Top             =   5160
      Width           =   5175
   End
   Begin VB.CheckBox chkGmMessagesPauseAll 
      BackColor       =   &H00000000&
      Caption         =   "Gm messages trigger special events and pauses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   72
      Top             =   4920
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.Frame frmNewCheats 
      BackColor       =   &H0000C000&
      Caption         =   "Options for newest Tibia version"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   5760
      TabIndex        =   68
      Top             =   2520
      Width           =   5055
      Begin VB.OptionButton chkClassic 
         BackColor       =   &H0000C000&
         Caption         =   "Classic mode. 0 % waste."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   4815
      End
      Begin VB.OptionButton chkEnhancedCheats 
         BackColor       =   &H0000C000&
         Caption         =   "No need to open bps, exact cast. Little chance of waste."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Value           =   -1  'True
         Width           =   4815
      End
      Begin VB.OptionButton chkTotalWaste 
         BackColor       =   &H0000C000&
         Caption         =   "War mode : Use the fastest method,. Big chance of waste."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   4815
      End
   End
   Begin VB.TextBox tibiaTittleFormat 
      Height          =   285
      Left            =   7440
      TabIndex        =   67
      Text            =   "$charactername$ - $expleft$ exp to lv $nextlevel$ - $exph$ exp/h"
      Top             =   6430
      Width           =   3255
   End
   Begin VB.ComboBox cmbWhere 
      Height          =   315
      Left            =   9000
      TabIndex        =   63
      Text            =   "19 : white center"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtExivaExpFormat 
      Height          =   285
      Left            =   7440
      TabIndex        =   61
      Text            =   $"frmHardcoreCheats.frx":0442
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox txtRelogBackpacks 
      Height          =   285
      Left            =   9960
      TabIndex        =   59
      Text            =   "4"
      Top             =   4680
      Width           =   375
   End
   Begin VB.CheckBox chkAutorelog 
      BackColor       =   &H00000000&
      Caption         =   "Auto relog in kick or serversave and reopen"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   58
      Top             =   4680
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5400
      Width           =   375
   End
   Begin VB.TextBox txtBlueauraDelay 
      Height          =   285
      Left            =   9600
      TabIndex        =   56
      Text            =   "300"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtExuraVitaMana 
      Height          =   285
      Left            =   10200
      TabIndex        =   54
      Text            =   "160"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtExuraVita 
      Height          =   285
      Left            =   8760
      TabIndex        =   51
      Text            =   "exura vita"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox chkAutoGratz 
      BackColor       =   &H00000000&
      Caption         =   "Auto gratz at level advances"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   50
      Top             =   4440
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkCaptionExp 
      BackColor       =   &H00000000&
      Caption         =   "Show exp in Tibia window title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   49
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton cmdBigMap 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Big map"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtAlarmUHs 
      Height          =   285
      Left            =   5040
      TabIndex        =   47
      Text            =   "5"
      Top             =   720
      Width           =   495
   End
   Begin VB.CheckBox chkRuneAlarm 
      BackColor       =   &H00000000&
      Caption         =   "When autohealing, alarm when UHS <"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   46
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtRemoteLeader 
      Height          =   285
      Left            =   4680
      TabIndex        =   45
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCommands 
      Height          =   2295
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "frmHardcoreCheats.frx":0521
      Top             =   120
      Width           =   5175
   End
   Begin VB.CheckBox chkColorEffects 
      BackColor       =   &H00000000&
      Caption         =   "Show colour effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   43
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox pushID 
      Height          =   375
      Left            =   4920
      TabIndex        =   41
      Text            =   "9"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Timer timerSpam 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   2760
   End
   Begin VB.HScrollBar scrollHP2 
      Height          =   255
      Left            =   2880
      Max             =   100
      TabIndex        =   38
      Top             =   1560
      Value           =   70
      Width           =   1935
   End
   Begin VB.CheckBox chkAutoVita 
      BackColor       =   &H00000000&
      Caption         =   "*AutoExuraVita if hp drop under"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ComboBox cmbOrderType 
      Height          =   315
      Left            =   2040
      TabIndex        =   35
      Text            =   "type 0 : SD (XYZ)"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   3120
      TabIndex        =   32
      Text            =   "firenow"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkAcceptSDorder 
      BackColor       =   &H00000000&
      Caption         =   "Accept order if you get in a channel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reactivate (will CLOSE proxy)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   2655
   End
   Begin VB.HScrollBar scrollHP 
      Height          =   255
      Left            =   2880
      Max             =   100
      TabIndex        =   27
      Top             =   1320
      Value           =   60
      Width           =   1935
   End
   Begin VB.HScrollBar scrollLight 
      Height          =   255
      Left            =   2880
      Max             =   15
      TabIndex        =   4
      Top             =   1080
      Value           =   15
      Width           =   1935
   End
   Begin VB.CheckBox chkAutoHeal 
      BackColor       =   &H00000000&
      Caption         =   "AutoRuneHeal if hp drop under"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton cmdOpenBackpacks 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Backpacks"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Map Click action"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3600
      TabIndex        =   22
      Top             =   3000
      Width           =   1815
      Begin VB.OptionButton ActionPath 
         BackColor       =   &H00000000&
         Caption         =   "Move there"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton ActionNothing 
         BackColor       =   &H00000000&
         Caption         =   "Do nothing"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton ActionMove 
         BackColor       =   &H00000000&
         Caption         =   "Summon to bag"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton ActionInspect 
         BackColor       =   &H00000000&
         Caption         =   "Game Inspect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Text            =   "-"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H00000000&
      Caption         =   "ontop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkLockOnMyFloor 
      BackColor       =   &H00000000&
      Caption         =   "lock"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdateMap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- update now !"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.OptionButton chkAutoUpdateMap 
      BackColor       =   &H00000000&
      Caption         =   "Full auto update (slow ! )"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   5280
      Width           =   2295
   End
   Begin VB.OptionButton chkUpdateMs 
      BackColor       =   &H00000000&
      Caption         =   "Update each x mseconds:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   5040
      Width           =   2295
   End
   Begin VB.OptionButton chkManualUpdate 
      BackColor       =   &H00000000&
      Caption         =   "No auto update"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   4800
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Timer timerAutoUpdater 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   5400
   End
   Begin VB.TextBox cmdMs 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Text            =   "1000"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   375
   End
   Begin VB.Timer timerLight 
      Interval        =   1000
      Left            =   3120
      Top             =   2760
   End
   Begin VB.CommandButton cmdOpenTrueRadar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show True Map"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkReveal 
      BackColor       =   &H00000000&
      Caption         =   "Reveal all invisible creatures"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox chkLight 
      BackColor       =   &H00000000&
      Caption         =   "Change light to this intensity :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkLogoutIfDanger 
      BackColor       =   &H00000000&
      Caption         =   "Logout! if danger on screen at start"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CheckBox chkApplyCheats 
      BackColor       =   &H00000000&
      Caption         =   "Activate hardcore cheats  "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "tibia window tittle:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   66
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "exiva exp message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   65
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Display exiva exp as:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   64
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblRedefineExp 
      BackColor       =   &H00000000&
      Caption         =   "Redefine exp info:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   62
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblBackpacks 
      BackColor       =   &H00000000&
      Caption         =   "bps"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   60
      Top             =   4740
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Blue aura delay between casts ( in mseconds):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   55
      Top             =   5475
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "mana:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10320
      TabIndex        =   53
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblRedefine 
      BackColor       =   &H00000000&
      Caption         =   "*Redefine ExuraVita:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   52
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5640
      X2              =   5640
      Y1              =   2520
      Y2              =   6720
   End
   Begin VB.Label lblLeader 
      BackColor       =   &H00000000&
      Caption         =   "Only accept order from this leader (leave blank for no leader) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label lblAdvanced 
      BackColor       =   &H00000000&
      Caption         =   "internal p delay"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   40
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblHPvalue2 
      BackColor       =   &H00000000&
      Caption         =   "70 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   39
      Top             =   1560
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   5520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblOn 
      BackColor       =   &H00000000&
      Caption         =   "on targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   36
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblRead 
      BackColor       =   &H00000000&
      Caption         =   "read order as:  cast"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOrder2 
      BackColor       =   &H00000000&
      Caption         =   ":targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblHPvalue 
      BackColor       =   &H00000000&
      Caption         =   "60 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00000000&
      Caption         =   "Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblYourPos 
      BackColor       =   &H00000000&
      Caption         =   "Your position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Tile stack for last selected position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblArraySelected 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Label lblLightValue 
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmHardcoreCheats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Public Sub chkApplyCheats_Click()
  Dim i As Integer
  If chkApplyCheats.Value = 0 Then
    chkRuneAlarm.Value = 0
    chkRuneAlarm.enabled = False
    chkLogoutIfDanger.Value = 0
    chkReveal.Value = 0
    chkLight.Value = 0
    frmTrueMap.Hide
    frmBackpacks.Hide
    scrollLight.enabled = False
    chkLogoutIfDanger.enabled = False
    chkReveal.enabled = False
    chkLight.enabled = False
    cmdOpenTrueRadar.enabled = False
    lblLightValue.enabled = False
    chkApplyCheats.enabled = False
    chkManualUpdate.Value = True
    chkUpdateMs.Value = False
    chkAutoUpdateMap.Value = False
    chkManualUpdate.enabled = False
    chkUpdateMs.enabled = False
    chkAutoUpdateMap.enabled = False
    timerAutoUpdater.enabled = False
    cmdMs.enabled = False
    cmdChange.enabled = False
    cmdUpdateMap.enabled = False
    Frame1.enabled = False
    chkLockOnMyFloor.enabled = False
    chkOnTop.enabled = False
    cmbCharacter.enabled = False
    ActionInspect.enabled = False
    ActionMove.enabled = False
    ActionNothing.enabled = False
    ActionPath.enabled = False
    cmdOpenBackpacks.enabled = False
    scrollHP.enabled = False
    chkAutoHeal.enabled = False
    chkAutoHeal.Value = 0
    cmdReset.enabled = True
    lblHPvalue.enabled = False
    txtCommands.enabled = False
    txtOrder.enabled = False
    lblOrder2.enabled = False
    chkAcceptSDorder.Value = 0
    chkAcceptSDorder.enabled = False
    cmbOrderType.enabled = False
    lblRead.enabled = False
    lblOn.enabled = False
    chkAutoVita.enabled = False
    chkAutoVita.Value = 0
    scrollHP2.enabled = False
    lblHPvalue2.enabled = False
    chkColorEffects.Value = 0
    chkColorEffects.enabled = False
    lblLeader.enabled = False
    txtRemoteLeader.enabled = False
    lblAdvanced.enabled = False
    pushID.enabled = False
    cmdBigMap.enabled = False
    For i = 1 To MAXCLIENTS
      GotPacketWarning(i) = True
    Next i
  End If
End Sub

Private Sub chkAutoHeal_Click()
  Dim i As Integer
  If chkAutoHeal.Value = 0 Then
    For i = 1 To MAXCLIENTS
      RemoveSpamOrder i, 1 'remove  auto UH
    Next i
  End If
End Sub

Private Sub chkAutoUpdateMap_Click()
  timerAutoUpdater.enabled = False
End Sub




Private Sub chkLockOnMyFloor_Click()
  If chkLockOnMyFloor.Value = 1 And cmbCharacter.ListIndex = mapIDselected And mapIDselected > 0 Then
    If mapFloorSelected <> myZ(mapIDselected) Then
      mapFloorSelected = myZ(mapIDselected)
      frmTrueMap.DrawFloor
    End If
  End If
End Sub

Private Sub chkManualUpdate_Click()
  timerAutoUpdater.enabled = False
End Sub

Public Sub chkOnTop_Click()
  If chkOnTop.Value = 1 Then
    ToggleTopmost frmTrueMap.hwnd, True
    ToggleTopmost frmMapReader.hwnd, True
    MapWantedOnTop = True
  Else
    ToggleTopmost frmTrueMap.hwnd, False
    ToggleTopmost frmMapReader.hwnd, False
    MapWantedOnTop = False
  End If
End Sub







Private Sub chkUpdateMs_Click()
  cmdMs_Change
  timerAutoUpdater.enabled = True
End Sub


Private Sub cmbCharacter_Click()
  mapIDselected = cmbCharacter.ListIndex
  If mapIDselected > 0 Then
    If TrialVersion = True Then
      If GameConnected(mapIDselected) = True And sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
          mapFloorSelected = myZ(mapIDselected)
          lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
          frmTrueMap.SetButtonColours
          frmTrueMap.DrawFloor
      End If
    Else
      If GameConnected(mapIDselected) = True Then
        mapFloorSelected = myZ(mapIDselected)
        lblPosition = "x=" & myX(mapIDselected) & ", y=" & myY(mapIDselected) & ", z=" & myZ(mapIDselected)
        frmTrueMap.SetButtonColours
        frmTrueMap.DrawFloor
      End If
    End If
  End If
End Sub









Private Sub cmdBigMap_Click()
  frmMapReader.WindowState = vbNormal
  frmMapReader.Show
  DoEvents
  frmMapReader.ShowCenter
  DoEvents
  frmMapReader.timerBigMapUpdate.enabled = True
End Sub

Public Sub cmdMs_Change()
 Dim lngValue
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lngValue = CLng(cmdMs.Text)
  If lngValue >= 10 And lngValue <= 500000 Then
    timerAutoUpdater.Interval = lngValue
  Else
    cmdMs.Text = "1000"
    timerAutoUpdater.Interval = 1000
  End If
  Exit Sub
gotError:
  cmdMs.Text = "1000"
  timerAutoUpdater.Interval = 1000
End Sub


Private Sub cmdOpenBackpacks_Click()
  frmBackpacks.Show
End Sub

Private Sub cmdOpenTrueRadar_Click()
  frmTrueMap.WindowState = vbNormal
  frmTrueMap.Show
End Sub

Private Sub cmdReset_Click()
  chkApplyCheats.Value = 1
  chkReveal.Value = 1
  chkLight.Value = 1
  chkAutoHeal.Value = 1
  frmMenu.Form_Unload False
End Sub

Private Sub cmdUpdateMap_Click()
  If TrialVersion = True Then
    If sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
      frmTrueMap.DrawFloor
    End If
  Else
    frmTrueMap.DrawFloor
  End If
End Sub



Private Sub Form_Load()
  cmbOrderType.Clear
  cmbOrderType.AddItem "type 0 : SD (XYZ)", 0
  cmbOrderType.AddItem "type 1 : HMM (XYZ)", 1
  cmbOrderType.AddItem "type 2 : Explosion (XYZ)", 2
  cmbOrderType.AddItem "type 3 : IH (XYZ)", 3
  cmbOrderType.AddItem "type 4 : UH (XYZ)", 4
  cmbOrderType.AddItem "type 5 : SD (battlelist)", 5
  cmbOrderType.AddItem "type 6 : HMM (battlelist)", 6
  cmbOrderType.AddItem "type 7 : Explosion (battlelist)", 7
  cmbOrderType.AddItem "type 8 : IH (battlelist)", 8
  cmbOrderType.AddItem "type 9 : UH (battlelist)", 9
  cmbOrderType.AddItem "type A : Say (text)", 10
  cmbOrderType.AddItem "type B : fireball (battlelist)", 11
  cmbOrderType.AddItem "type C : stalagmite (battlelist)", 12
  cmbOrderType.AddItem "type D : icicle (battlelist)", 13
  cmbOrderType.Text = "type 5 : SD (battlelist)"
  cmbWhere.Clear
  cmbWhere.AddItem "01 : yellow default"
  cmbWhere.AddItem "02 : yellow default"
  cmbWhere.AddItem "03 : yellow default"
  cmbWhere.AddItem "04 : blue default"
  cmbWhere.AddItem "05 : blue default"
  cmbWhere.AddItem "06 : invisible?"
  cmbWhere.AddItem "07 : invisible?"
  cmbWhere.AddItem "08 : invisible?"
  cmbWhere.AddItem "09 : red default"
  cmbWhere.AddItem "10 : invisible?"
  cmbWhere.AddItem "11 : red default"
  cmbWhere.AddItem "12 : red default"
  cmbWhere.AddItem "13 : red default"
  cmbWhere.AddItem "14 : invisible?"
  cmbWhere.AddItem "15 : red default"
  cmbWhere.AddItem "16 : orange default"
  cmbWhere.AddItem "17 : orange default"
  cmbWhere.AddItem "18 : red center"
  cmbWhere.AddItem "19 : white center"
  cmbWhere.AddItem "20 : 1 line log"
  cmbWhere.AddItem "21 : white system"
  cmbWhere.AddItem "22 : green middle"
  cmbWhere.AddItem "23 : white system"
  cmbWhere.AddItem "24 : purple default"
  cmbWhere.Text = ExivaExpPlace
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub









Private Sub pushID_Change()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  PUSHDELAYTIMES = CLng(pushID.Text)
  Exit Sub
goterr:
  PUSHDELAYTIMES = 9
  pushID.Text = 9
End Sub

Public Sub scrollHP_Change()
  ChangeGLOBAL_RUNEHEAL_HP scrollHP.Value
End Sub

Public Sub scrollHP2_Change()
  lblHPvalue2.Caption = CStr(scrollHP2.Value) & " %"
End Sub

Public Sub scrollLight_Change()
 lblLightValue.Caption = CStr(Round((scrollLight.Value / 15) * 100)) & " %"
  LightIntesityHex = GoodHex(CByte(scrollLight.Value))
End Sub

Private Sub timerAutoUpdater_Timer()
 If mapIDselected > 0 Then
    If TrialVersion = True Then
      If sentWelcome(mapIDselected) = True And GotPacketWarning(mapIDselected) = False Then
        If chkLockOnMyFloor.Value = 1 Then
          mapFloorSelected = myZ(mapIDselected)
        End If
        frmTrueMap.SetButtonColours
        frmTrueMap.DrawFloor
      End If
    Else
      If chkLockOnMyFloor.Value = 1 Then
        mapFloorSelected = myZ(mapIDselected)
      End If
      frmTrueMap.SetButtonColours
      frmTrueMap.DrawFloor
    End If
  End If
End Sub

Private Sub timerLight_Timer()
  Dim i As Integer
  'Dim cPacket() As Byte
  Dim errorD As Integer
  Dim inRes As Integer
  Dim aRes As Long
  Dim playerS As String
  'Exit Sub '
  #If FinalMode Then
  On Error GoTo endT
  #End If
  If (TrialVersion = False) And (trialSafety4 <> 4) Then
    End
  End If
  If chkApplyCheats.Value = 1 Then
  
  If (Me.chkCaptionExp.Value = 1) Then
    UpdateTibiaTitles
  End If
  
  For i = 1 To HighestConnectionID
      errorD = i
  If (GameConnected(i) = True) And (sentWelcome(i) = True) And (GotPacketWarning(i) = False) Then
    If ReconnectionStage(i) = 3 Then
      If frmBackpacks.totalbpsOpen(i) < CLng(txtRelogBackpacks.Text) Then
        aRes = openBP(i)
        If aRes = -1 Then
          
          ReconnectionStage(i) = 0 'forced
          If frmBackpacks.totalbpsOpen(i) = 0 Then
             If TibiaVersionLong < 790 Then
               ReconnectionStage(i) = 10
               frmMain.DoCloseActions i
               frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "WARNING: Character " & CharacterName(i) & " had no backpack after relog so it was left closed"
               ReconnectionStage(i) = 0
               logoutAllowed(i) = 0
             Else
               aRes = SendLogSystemMessageToClient(i, "It was not possible to open the desired number of backpacks.")
               DoEvents
             End If
          Else
            aRes = SendLogSystemMessageToClient(i, "It was not possible to open the desired number of backpacks.")
            DoEvents
          End If
        Else
          DoEvents
        End If
      Else
        ReconnectionStage(i) = 0
        logoutAllowed(i) = 0
        aRes = SendLogSystemMessageToClient(i, "Successfully opened " & CStr(txtRelogBackpacks.Text) & " containers.")
        DoEvents
      End If
    End If
      ' ALIVE? (45seconds without packet is not good)
    If (lastPing(i) < (GetTickCount() - MaxTimeWithoutServerPackets)) Then
      If frmHardcoreCheats.chkAutorelog.Value = 1 Then
        aRes = GiveGMmessage(i, "ISP - server down detected (too much time without receiving anything from server)", "Blackdproxy")
        DoEvents
        lastPing(i) = GetTickCount() + 3600000
        StartReconnection i
      Else
        aRes = GiveGMmessage(i, "ISP - server down detected (too much time without receiving anything from server)", "Blackdproxy")
        DoEvents
        lastPing(i) = GetTickCount()
        If frmRunemaker.chkCloseSound.Value = 1 Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "(Giving alarm because server or ISP went probably down)"
          ChangePlayTheDangerSound True
        End If
      End If
    End If
    If CheatsPaused(i) = False Then
      If (RuneMakerOptions(i).msgSound2 = True) Then
        playerS = PlayerOnScreen(i)
        If playerS <> "" Then
        'If DangerPlayer(i) = True Then
        '  playerS = DangerPlayerName(i)
          PlayMsgSound2 = True
          If publicDebugMode = True Then
            aRes = SendLogSystemMessageToClient(i, "[Debug] Giving alarm because you have on screen: " & playerS)
            DoEvents
          End If
        End If
      End If

      If DangerGM(i) = True Then
        If (GetTickCount() > LogoutTimeGM(i)) And (LogoutTimeGM(i) <> 0) Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(i) & " closed after random time,because : " & GMname(i)
          aRes = GiveServerError("Closed after random time, because : " & GMname(i), i)
          DoEvents
          ReconnectionStage(i) = 10
          frmMain.DoCloseActions i
          ReconnectionStage(i) = 0
          DoEvents
          Exit Sub
        End If
       ' If frmRunemaker.ChkDangerSound.Value = 1 Then
          ChangePlayTheDangerSound True
        'End If
      End If
      If DangerPK(i) = True Then
          If PlayTheDangerSound = False Then
            If frmCavebot.chkChangePkHeal.Value = 1 Then
              ChangeGLOBAL_RUNEHEAL_HP frmCavebot.scrollPkHeal.Value
              aRes = GiveGMmessage(i, "WARNING : YOU ARE UNDER PK ATTACK ! (" & DangerPKname(i) & ") Auto heal have been auto increased to " & frmCavebot.lblPKhealValue.Caption, "BlackdProxy")
              DoEvents
            Else
              aRes = GiveGMmessage(i, "WARNING : YOU ARE UNDER PK ATTACK ! (" & DangerPKname(i) & ")", "BlackdProxy")
              DoEvents
            End If
            aRes = SendLogSystemMessageToClient(i, "BlackdProxy: To deactivate alarm do Exiva cancel")
            DoEvents
          End If
          ChangePlayTheDangerSound True
      End If
    End If
      If (chkLight.Value = 1) Then
        If IDstring(i) <> "" Then
          If frmMain.sckClientGame(i).State = sckConnected Then
            If PlayTheDangerSound = True Then
              nextLight(i) = "FD"
              enLight i
              nextLight(i) = "D7"
            ElseIf nextLight(i) <> "D7" Then
              nextLight(i) = "D7"
            Else
              enLight i
            End If
          Else
            frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Client #" & CStr(i) & "# lost connection during timerLight_Timer"
            frmMain.DoCloseActions i
            DoEvents
          End If
        End If
      End If
  End If
  Next i
  

  
  If PlayMsgSound = True Then
    PlayMsgSound = False
    If frmRunemaker.ChkDangerSound.Value = 1 Then
        DirectX_PlaySound 3 ' play ding.wav
    End If
    If ((frmRunemaker.chkOnDangerSS.Value = 1) And (frmRunemaker.timerSS.enabled = False)) Then
        frmRunemaker.timerSS.enabled = True
    End If
  End If
  If PlayMsgSound2 = True Then
    PlayMsgSound2 = False
    If frmRunemaker.ChkDangerSound.Value = 1 Then
        DirectX_PlaySound 1 ' play player.wav
    End If
    If ((frmRunemaker.chkOnDangerSS.Value = 1) And (frmRunemaker.timerSS.enabled = False)) Then
        frmRunemaker.timerSS.enabled = True
    End If
  End If
  If (PlayTheDangerSound = True) Then ' And (frmRunemaker.ChkDangerSound.Value = 1) Then
    If frmRunemaker.ChkDangerSound.Value = 1 Then
        DirectX_PlaySound 2 ' play danger.wav
    End If
  End If
      
  End If
  Exit Sub
endT:
  On Error GoTo severeE:
  If PlayMsgSound = True Then
    PlayMsgSound = False
    DirectX_PlaySound 3 ' play ding.wav
  End If
  If PlayMsgSound2 = True Then
    PlayMsgSound2 = False
    DirectX_PlaySound 1 ' play player.wav
  End If
  If (PlayTheDangerSound = True) And (frmRunemaker.ChkDangerSound.Value = 1) Then
    DirectX_PlaySound 2 ' play danger.wav
  End If
severeE:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Got unexpected error at timerLight_Timer - Client ID terminated!"
  frmMain.DoCloseActions i
  DoEvents
End Sub

Private Sub timerSpam_Timer()
  Dim i As Integer
  Dim order As Integer
  Dim cid As Integer
  Dim resA As Long
  Dim gtc As Long
  Dim posSetting As Long
  Dim act As String
  #If FinalMode Then
  On Error GoTo errIgnore
  #End If
  For i = 1 To MAXCLIENTS
    If (GameConnected(i) = True) Then
      If ((SpamAutoFastHeal(i) = True) And _
       (CheatsPaused(i) = False)) Then
        gtc = GetTickCount()
        If (nextFastHeal(i) <= gtc) Then
          resA = UseFastUH(i)
          If resA = 0 Then
            cancelAllMove(i) = GetTickCount() + 500
            If frmHardcoreCheats.chkColorEffects.Value = 1 Then
              If Not (nextLight(i) = "04") Then
                nextLight(i) = "04"
               enLight i
             End If
            End If
          End If
          nextFastHeal(i) = gtc + BlueAuraDelay
        End If
       ElseIf (SpamAutoHeal(i) = True) And _
      ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
        UHRetryCount(i) = UHRetryCount(i) + 1
          If (UHRetryCount(i) < 50) Then
            If (TibiaVersionLong < 780) Then
                cancelAllMove(i) = GetTickCount() + 500
            End If
          ElseIf (PlayTheDangerSound = False) Then
            'give msg !
            If chkClassic.Value = False Then
                resA = GiveGMmessage(i, "Autohealer is unable to heal you. Maybe no UHs left!", "Warning")
                ChangePlayTheDangerSound True
                DoEvents
            End If
          End If
          'heal
          resA = UseUH(i)
          If resA = 0 Then
            If frmHardcoreCheats.chkColorEffects.Value = 1 Then
              If Not (nextLight(i) = "04") Then
                nextLight(i) = "04"
                enLight i
              End If
            End If
          End If
      ElseIf ((SpamAutoPush(i) = True) And _
       (CheatsPaused(i) = False)) Then
        If pushDelay(i) = 0 Then
          resA = DoPush(i)
          pushDelay(i) = PUSHDELAYTIMES
          DoEvents
        Else
          pushDelay(i) = pushDelay(i) - 1
        End If
      ElseIf ((SpamAutoMana(i) = True) And _
       (CheatsPaused(i) = False)) Then
       ' Note: it can be a problem when there are no manas left!
        resA = UseFluid(i, byteMana)
      End If
    End If
  Next i
  Exit Sub
errIgnore:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Unexpected error at timerSpam_Timer()"
End Sub

Private Function GetOneTitle(idConnection As Integer) As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  'UpdateExpVars idConnection
  var_lf(idConnection) = ". "
  GetOneTitle = parseVars(idConnection, tibiaTittleFormat.Text)
  Exit Function
goterr:
  GetOneTitle = "Tibia"
End Function

Private Sub UpdateTibiaTitles()
  Dim i As Integer
  Dim Message As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If frmStealth.chkStealthExp = 1 Then
    If stealthIDselected <> 0 Then
        If GameConnected(stealthIDselected) = True Then
            Message = GetOneTitle(stealthIDselected)
            frmStealth.Caption = Message
        End If
    End If
  Else
  GetProcessAllProcessIDs
  For i = 1 To MAXCLIENTS
    If (GameConnected(i) = True) Then
      Message = GetOneTitle(i)
      If ProcessID(i) > 0 Then
        SetWindowText ProcessID(i), Message
      End If
    End If
  Next i
  End If
goterr:
  ' just end...
End Sub

Private Sub txtBlueauraDelay_Change()
 Dim lngValue
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lngValue = CLng(txtBlueauraDelay.Text)
  If lngValue >= 10 And lngValue <= 500000 Then
    BlueAuraDelay = lngValue
  Else
    txtBlueauraDelay.Text = "300"
    BlueAuraDelay = 300
  End If
  Exit Sub
gotError:
  txtBlueauraDelay.Text = "300"
  BlueAuraDelay = 300
End Sub



