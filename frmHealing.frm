VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmHealing 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Healing"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMana 
      Height          =   285
      Left            =   1140
      TabIndex        =   102
      Text            =   "SELF MANA"
      Top             =   2100
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtPot 
      Height          =   285
      Left            =   1140
      TabIndex        =   101
      Text            =   "SELF UHEAL"
      Top             =   1740
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer timerHeal 
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.Frame frmNewCheats 
      Caption         =   "Heal Method"
      Height          =   975
      Left            =   120
      TabIndex        =   97
      Top             =   2640
      Width           =   5655
      Begin VB.OptionButton chkTotalWaste 
         Caption         =   "Hotkeys. leave this default for newer versions"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   180
         TabIndex        =   100
         Top             =   600
         Value           =   -1  'True
         Width           =   3555
      End
      Begin VB.OptionButton chkEnhancedCheats 
         BackColor       =   &H80000018&
         Caption         =   "No need to open bps, exact cast. Little chance of waste."
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   3660
         TabIndex        =   99
         Top             =   240
         Width           =   15
      End
      Begin VB.OptionButton chkClassic 
         Caption         =   "Classic mode. for Old Tibia Clients"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   180
         TabIndex        =   98
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   1200
      TabIndex        =   85
      Text            =   "-"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtSpelllo 
      Height          =   285
      Left            =   1140
      TabIndex        =   84
      Text            =   "exura vita"
      Top             =   1020
      Width           =   1275
   End
   Begin VB.TextBox txtSpellhi 
      Height          =   285
      Left            =   1140
      TabIndex        =   83
      Text            =   "exura gran"
      Top             =   660
      Width           =   1275
   End
   Begin VB.TextBox txtManahi 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      TabIndex        =   82
      Text            =   "70"
      Top             =   660
      Width           =   615
   End
   Begin VB.TextBox txtManalo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      TabIndex        =   81
      Text            =   "160"
      Top             =   1020
      Width           =   615
   End
   Begin VB.TextBox txtHealpot 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   80
      Text            =   "0"
      Top             =   1740
      Width           =   735
   End
   Begin VB.TextBox txtManapot 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   79
      Text            =   "0"
      Top             =   2100
      Width           =   735
   End
   Begin VB.TextBox txtHealthlo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   78
      Text            =   "0"
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtHealthhi 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   77
      Text            =   "0"
      Top             =   660
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4320
      TabIndex        =   76
      Text            =   "-"
      Top             =   1740
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   75
      Text            =   "-"
      Top             =   2100
      Width           =   1215
   End
   Begin VB.CheckBox chkApplyCheats 
      BackColor       =   &H80000018&
      Caption         =   "Activate hardcore cheats  "
      Height          =   375
      Left            =   5640
      TabIndex        =   51
      Top             =   7560
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox chkLogoutIfDanger 
      BackColor       =   &H00000000&
      Caption         =   "Logout! if danger on screen at start"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   50
      Top             =   9120
      Width           =   3135
   End
   Begin VB.CheckBox chkLight 
      BackColor       =   &H00000000&
      Caption         =   "Change light to this intensity :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   49
      Top             =   9360
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdOpenTrueRadar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show True Map"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   10800
      Width           =   1335
   End
   Begin VB.Timer timerLight 
      Interval        =   1000
      Left            =   5880
      Top             =   6960
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   12720
      Width           =   375
   End
   Begin VB.TextBox cmdMs 
      Height          =   285
      Left            =   4560
      TabIndex        =   46
      Text            =   "1000"
      Top             =   12720
      Width           =   735
   End
   Begin VB.Timer timerAutoUpdater 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   13080
   End
   Begin VB.OptionButton chkManualUpdate 
      BackColor       =   &H00000000&
      Caption         =   "No auto update"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   45
      Top             =   12480
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton chkUpdateMs 
      BackColor       =   &H00000000&
      Caption         =   "Update each x mseconds:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   44
      Top             =   12720
      Width           =   2295
   End
   Begin VB.OptionButton chkAutoUpdateMap 
      BackColor       =   &H00000000&
      Caption         =   "Full auto update (slow ! )"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   43
      Top             =   12960
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdateMap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- update now !"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   10800
      Width           =   1575
   End
   Begin VB.CheckBox chkLockOnMyFloor 
      BackColor       =   &H00000000&
      Caption         =   "lock"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   41
      Top             =   11280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H00000000&
      Caption         =   "ontop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   40
      Top             =   11280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Map Click action"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   5280
      TabIndex        =   35
      Top             =   10680
      Width           =   1815
      Begin VB.OptionButton ActionInspect 
         BackColor       =   &H00000000&
         Caption         =   "Game Inspect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton ActionMove 
         BackColor       =   &H00000000&
         Caption         =   "Summon to bag"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton ActionNothing 
         BackColor       =   &H00000000&
         Caption         =   "Do nothing"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton ActionPath 
         BackColor       =   &H00000000&
         Caption         =   "Move there"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOpenBackpacks 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Backpacks"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10440
      Width           =   1575
   End
   Begin VB.HScrollBar scrollLight 
      Height          =   255
      Left            =   4200
      Max             =   15
      TabIndex        =   33
      Top             =   8880
      Value           =   15
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reactivate (will CLOSE proxy)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CheckBox chkAcceptSDorder 
      BackColor       =   &H00000000&
      Caption         =   "Accept order if you get in a channel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   9720
      Width           =   3495
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   4800
      TabIndex        =   30
      Text            =   "firenow"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.ComboBox cmbOrderType 
      Height          =   315
      Left            =   3720
      TabIndex        =   29
      Text            =   "type 0 : SD (XYZ)"
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Timer timerSpam 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   7080
   End
   Begin VB.TextBox pushID 
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Text            =   "9"
      Top             =   12720
      Width           =   495
   End
   Begin VB.CheckBox chkColorEffects 
      BackColor       =   &H00000000&
      Caption         =   "Show colour effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   27
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtRemoteLeader 
      Height          =   285
      Left            =   6480
      TabIndex        =   26
      Top             =   9240
      Width           =   735
   End
   Begin VB.CheckBox chkRuneAlarm 
      BackColor       =   &H00000000&
      Caption         =   "When autohealing, alarm when UHS <"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   8280
      Width           =   1815
   End
   Begin VB.TextBox txtAlarmUHs 
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Text            =   "5"
      Top             =   8400
      Width           =   495
   End
   Begin VB.CommandButton cmdBigMap 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Big map"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10440
      Width           =   1335
   End
   Begin VB.CheckBox chkCaptionExp 
      BackColor       =   &H00000000&
      Caption         =   "Show exp in Tibia window title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   7800
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox chkAutoGratz 
      BackColor       =   &H00000000&
      Caption         =   "Auto gratz at level advances"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox txtBlueauraDelay 
      Height          =   285
      Left            =   9480
      TabIndex        =   20
      Text            =   "300"
      Top             =   8400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ok"
      Height          =   300
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9360
      Width           =   375
   End
   Begin VB.CheckBox chkAutorelog 
      BackColor       =   &H00000000&
      Caption         =   "Auto relog in kick or serversave and reopen"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   8400
      Width           =   4215
   End
   Begin VB.TextBox txtRelogBackpacks 
      Height          =   285
      Left            =   7680
      TabIndex        =   17
      Text            =   "4"
      Top             =   8640
      Width           =   375
   End
   Begin VB.TextBox txtExivaExpFormat 
      Height          =   285
      Left            =   8520
      TabIndex        =   16
      Text            =   $"frmHealing.frx":0000
      Top             =   9360
      Width           =   3255
   End
   Begin VB.ComboBox cmbWhere 
      Height          =   315
      Left            =   10200
      TabIndex        =   15
      Text            =   "19 : white center"
      Top             =   10320
      Width           =   1695
   End
   Begin VB.TextBox tibiaTittleFormat 
      Height          =   285
      Left            =   7020
      TabIndex        =   14
      Text            =   "$charactername$ - $expleft$ exp to lv $nextlevel$ - $exph$ exp/h"
      Top             =   10155
      Width           =   3255
   End
   Begin VB.CheckBox chkGmMessagesPauseAll 
      BackColor       =   &H00000000&
      Caption         =   "Gm messages trigger special events and pauses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   13
      Top             =   7800
      Width           =   4215
   End
   Begin VB.CheckBox chkProtectedShots 
      BackColor       =   &H00000000&
      Caption         =   "Avoid shoting damage runes if your %hp < AutoRuneHeal %hp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   9120
      Width           =   5175
   End
   Begin VB.TextBox txtExuraVitaMana3 
      Height          =   285
      Left            =   9600
      TabIndex        =   11
      Text            =   "0"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar scrollHP 
      Height          =   255
      Left            =   10320
      Max             =   100
      TabIndex        =   10
      Top             =   9840
      Value           =   60
      Width           =   1455
   End
   Begin VB.HScrollBar scrollHP2 
      Height          =   255
      Left            =   10320
      Max             =   100
      TabIndex        =   9
      Top             =   6960
      Value           =   70
      Width           =   1455
   End
   Begin VB.HScrollBar scrollHP22 
      Height          =   255
      Left            =   10320
      Max             =   100
      TabIndex        =   8
      Top             =   6600
      Value           =   70
      Width           =   1455
   End
   Begin VB.HScrollBar scrollHP3 
      Height          =   255
      Left            =   10320
      Max             =   100
      TabIndex        =   7
      Top             =   7680
      Value           =   70
      Width           =   1455
   End
   Begin VB.HScrollBar scrollHP4 
      Height          =   255
      Left            =   10320
      Max             =   100
      TabIndex        =   6
      Top             =   8040
      Value           =   40
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9960
      TabIndex        =   5
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9960
      TabIndex        =   4
      Text            =   "0"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9960
      TabIndex        =   3
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9960
      TabIndex        =   2
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   9960
      TabIndex        =   1
      Text            =   "0"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "arme3"
      Height          =   195
      Left            =   9180
      TabIndex        =   0
      Top             =   6480
      Width           =   735
   End
   Begin JwldButn2b.JeweledButton cmdApply 
      Height          =   255
      Left            =   4560
      TabIndex        =   74
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "apply"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Hover=   16761024
      BorderColor_Inner=   16777215
   End
   Begin VB.Label Label8 
      Caption         =   "FOR THE CHANGES TAKE EFFECT, HIT THE APPLY BUTTON ON THE BOT"
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   3630
      Width           =   6015
   End
   Begin VB.Label lblChar 
      Caption         =   "Auto-Heal :"
      Height          =   255
      Left            =   240
      TabIndex        =   96
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblYourPos 
      Caption         =   "Mana :"
      Height          =   255
      Left            =   4320
      TabIndex        =   95
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblYourPos2 
      Caption         =   "Mana :"
      Height          =   255
      Left            =   4320
      TabIndex        =   94
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   2460
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000010&
      X1              =   5760
      X2              =   120
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000015&
      X1              =   5760
      X2              =   120
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000010&
      X1              =   5760
      X2              =   5760
      Y1              =   540
      Y2              =   2460
   End
   Begin VB.Label Label2 
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   93
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   92
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Mana :"
      Height          =   255
      Left            =   2760
      TabIndex        =   91
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Health :"
      Height          =   255
      Left            =   2700
      TabIndex        =   90
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Spell Hi :"
      Height          =   255
      Left            =   300
      TabIndex        =   89
      Top             =   705
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Spell Lo :"
      Height          =   255
      Left            =   300
      TabIndex        =   88
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Heal Pot"
      Height          =   255
      Left            =   300
      TabIndex        =   87
      Top             =   1785
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Mana Pot"
      Height          =   255
      Left            =   300
      TabIndex        =   86
      Top             =   2145
      Width           =   735
   End
   Begin VB.Label lblLightValue 
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   73
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblArraySelected 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   72
      Top             =   13560
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Tile stack for last selected position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   71
      Top             =   13320
      Width           =   5175
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00000000&
      Caption         =   "Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   70
      Top             =   12120
      Width           =   2055
   End
   Begin VB.Label lblOrder2 
      BackColor       =   &H00000000&
      Caption         =   ":targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   69
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblRead 
      BackColor       =   &H00000000&
      Caption         =   "read order as:  cast"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   68
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Label lblOn 
      BackColor       =   &H00000000&
      Caption         =   "on targetname"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   67
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   7200
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label lblAdvanced 
      BackColor       =   &H00000000&
      Caption         =   "internal p delay"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   66
      Top             =   12720
      Width           =   855
   End
   Begin VB.Label lblLeader 
      BackColor       =   &H00000000&
      Caption         =   "Only accept order from this leader (leave blank for no leader) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   65
      Top             =   10080
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10860
      X2              =   9840
      Y1              =   9180
      Y2              =   7800
   End
   Begin VB.Label lblRedefine 
      BackColor       =   &H00000000&
      Caption         =   "*Redefine ExuraVita:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   64
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Blue aura delay between casts ( in mseconds):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   63
      Top             =   9480
      Width           =   3615
   End
   Begin VB.Label lblBackpacks 
      BackColor       =   &H00000000&
      Caption         =   "bps"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   62
      Top             =   8700
      Width           =   495
   End
   Begin VB.Label lblRedefineExp 
      BackColor       =   &H00000000&
      Caption         =   "Redefine exp info:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   61
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Display exiva exp as:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7020
      TabIndex        =   60
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "exiva exp message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   59
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "tibia window tittle:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   58
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label lblYourPos3 
      BackColor       =   &H80000018&
      Caption         =   "Mana :"
      Height          =   255
      Left            =   9000
      TabIndex        =   57
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblHPvalue3 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   11880
      TabIndex        =   56
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label lblHPvalue4 
      BackColor       =   &H80000018&
      Caption         =   "40 %"
      Height          =   255
      Left            =   11880
      TabIndex        =   55
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label lblHPvalue 
      BackColor       =   &H80000018&
      Caption         =   "60 %"
      Height          =   255
      Left            =   11880
      TabIndex        =   54
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label lblHPvalue22 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   11640
      TabIndex        =   53
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label lblHPvalue2 
      BackColor       =   &H80000018&
      Caption         =   "70 %"
      Height          =   255
      Left            =   11880
      TabIndex        =   52
      Top             =   6960
      Width           =   495
   End
End
Attribute VB_Name = "frmHealing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub chkClassic_Click()

frmHardcoreCheats.chkClassic.value = True

End Sub

Private Sub chkTotalWaste_Click()

frmHardcoreCheats.chkTotalWaste.value = True

End Sub

Private Sub cmbCharacter_Click()

 healingIDselected = cmbCharacter.ListIndex
  If healingIDselected > 0 Then
      UpdateValues
  End If

End Sub

Private Sub cmdApply_Click()

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtHealthhi = txtHealthhi.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtHealthlo = txtHealthlo.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtManahi = txtManahi.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtManalo = txtManalo.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtHealpot = txtHealpot.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtManapot = txtManapot.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).Combo1 = Combo1.Text
End If

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).Combo2 = Combo2.Text
End If

UpdateValues

End Sub

Public Sub UpdateValues()
Dim i As Integer
Dim idConnection As Integer

If healingIDselected <= 0 Then
frmHealing.txtSpelllo.Text = healingCheatsOptions_txtSpelllo_default
frmHealing.txtPot.Text = healingCheatsOptions_txtPot_default
frmHealing.txtMana.Text = healingCheatsOptions_txtMana_default
frmHealing.txtHealthhi.Text = healingCheatsOptions_txtHealthhi_default
frmHealing.txtHealthlo.Text = healingCheatsOptions_txtHealthlo_default
frmHealing.txtHealpot.Text = healingCheatsOptions_txtHealpot_default
frmHealing.txtManapot.Text = healingCheatsOptions_txtManapot_default
frmHealing.txtManahi.Text = healingCheatsOptions_txtManahi_default
frmHealing.txtManalo.Text = healingCheatsOptions_txtManalo_default
frmHealing.Combo1.Text = healingCheatsOptions_Combo1_default
frmHealing.Combo2.Text = healingCheatsOptions_Combo2_default
 Else
frmHealing.txtSpellhi.Text = healingCheatsOptions(healingIDselected).txtSpellhi
frmHealing.txtSpelllo.Text = healingCheatsOptions(healingIDselected).txtSpelllo
frmHealing.txtPot.Text = healingCheatsOptions(healingIDselected).txtPot
frmHealing.txtMana.Text = healingCheatsOptions(healingIDselected).txtMana
frmHealing.txtHealthhi.Text = healingCheatsOptions(healingIDselected).txtHealthhi
frmHealing.txtHealthlo.Text = healingCheatsOptions(healingIDselected).txtHealthlo
frmHealing.txtHealpot.Text = healingCheatsOptions(healingIDselected).txtHealpot
frmHealing.txtManapot.Text = healingCheatsOptions(healingIDselected).txtManapot
frmHealing.txtManahi.Text = healingCheatsOptions(healingIDselected).txtManahi
frmHealing.txtManalo.Text = healingCheatsOptions(healingIDselected).txtManalo
frmHealing.Combo1.Text = healingCheatsOptions(healingIDselected).Combo1
frmHealing.Combo2.Text = healingCheatsOptions(healingIDselected).Combo2
End If

End Sub

Private Sub Combo1_Click()
Dim Index As Integer
Dim idConnection As Integer

If healingIDselected > 0 Then

    If Combo1.ListIndex = 0 Then
        healingCheatsOptions(healingIDselected).Combo1 = "Health Potion"
    ElseIf Combo1.ListIndex = 1 Then
        healingCheatsOptions(healingIDselected).Combo1 = "Strong Health"
    ElseIf Combo1.ListIndex = 2 Then
        healingCheatsOptions(healingIDselected).Combo1 = "Great Health"
    ElseIf Combo1.ListIndex = 3 Then
        healingCheatsOptions(healingIDselected).Combo1 = "Ultimate Health"
    ElseIf Combo1.ListIndex = 4 Then
        healingCheatsOptions(healingIDselected).Combo1 = "Spirit"
    ElseIf Combo1.ListIndex = 5 Then
        healingCheatsOptions(healingIDselected).Combo1 = "UH"
    End If
    
End If

End Sub

Private Sub Combo2_Click()
Dim Index As Integer

If healingIDselected > 0 Then

    If Combo2.ListIndex = 0 Then
        healingCheatsOptions(healingIDselected).Combo2 = "Mana Fluid"
    ElseIf Combo2.ListIndex = 1 Then
        healingCheatsOptions(healingIDselected).Combo2 = "Mana Potion"
    ElseIf Combo2.ListIndex = 2 Then
        healingCheatsOptions(healingIDselected).Combo2 = "Strong Mana"
    ElseIf Combo2.ListIndex = 3 Then
        healingCheatsOptions(healingIDselected).Combo2 = "Great Mana"
    ElseIf Combo2.ListIndex = 4 Then
        healingCheatsOptions(healingIDselected).Combo2 = "Spirit"
    End If
    
End If

End Sub

Private Sub Form_Load()
LoadHealingChars
  Combo1.AddItem "Health Potion", 0
  Combo1.AddItem "Strong Health", 1
  Combo1.AddItem "Great Health", 2
  Combo1.AddItem "Ultimate Health", 3
  Combo1.AddItem "Spirit", 4
  Combo1.AddItem "UH", 5
  Combo2.AddItem "Mana Fluid", 0
  Combo2.AddItem "Mana Potion", 1
  Combo2.AddItem "Strong Mana", 2
  Combo2.AddItem "Great Mana", 3
  Combo2.AddItem "Spirit", 4
  
End Sub

Public Sub LoadHealingChars()
  Dim i As Long
  Dim firstC As Long

  firstC = 0
  cmbCharacter.Clear
  cmbCharacter.AddItem "-", 0
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True Then
      If firstC = 0 Then
        firstC = i
      End If
      cmbCharacter.AddItem CharacterName(i), i
    Else
      cmbCharacter.AddItem "-" & CStr(i) & "- NOT CONNECTED", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  healingIDselected = firstC
  UpdateValues
  
End Sub

Private Sub timerHeal_Timer()
Dim idConnection As Integer
Dim learnResult As TypeLearnResult
Dim aRes As Long
  
For idConnection = 1 To MAXCLIENTS
If GameConnected(idConnection) = True Then


'spell lo
   If (myHP(idConnection) < CLng(healingCheatsOptions(idConnection).txtHealthlo)) And (sentFirstPacket(idConnection) = True) And (myMana(idConnection) >= CLng(healingCheatsOptions(idConnection).txtManalo)) Then
                healingCheatsOptions(idConnection).exaust = True
                aRes = ExecuteInTibia(healingCheatsOptions(idConnection).txtSpelllo, idConnection, True)
                DoEvents
    Else
                healingCheatsOptions(idConnection).exaust = False
   End If
   
'spell hi
   If (myHP(idConnection) < CLng(healingCheatsOptions(idConnection).txtHealthhi)) And (sentFirstPacket(idConnection) = True) And (myMana(idConnection) >= CLng(healingCheatsOptions(idConnection).txtManahi)) Then
        If healingCheatsOptions(idConnection).exaust = False Then
                aRes = ExecuteInTibia(healingCheatsOptions(idConnection).txtSpellhi, idConnection, True)
                DoEvents
        End If
   End If
   healingCheatsOptions(idConnection).exaust = False

'mana heal
   If (sentFirstPacket(idConnection) = True) And (myMana(idConnection) < CLng(healingCheatsOptions(idConnection).txtManapot)) Then
            If healingCheatsOptions(idConnection).Combo2 = "Mana Fluid" Then
                aRes = ExecuteInTibia("exiva mana_fluid", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo2 = "Mana Potion" Then
                aRes = ExecuteInTibia("exiva mana_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo2 = "Strong Mana" Then
                aRes = ExecuteInTibia("exiva strong_mana_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo2 = "Great Mana" Then
                aRes = ExecuteInTibia("exiva great_mana_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo2 = "Spirit" Then
                aRes = ExecuteInTibia("exiva great_spirit_potion", idConnection, True)
                DoEvents
            End If
   End If

'pot heal
   If (sentFirstPacket(idConnection) = True) And (myHP(idConnection) < CLng(healingCheatsOptions(idConnection).txtHealpot)) Then
            If healingCheatsOptions(idConnection).Combo1 = "Health Potion" Then
                aRes = ExecuteInTibia("exiva health_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo1 = "Strong Health" Then
                aRes = ExecuteInTibia("exiva strong_health_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo1 = "Great Health" Then
                aRes = ExecuteInTibia("exiva great_health_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo1 = "Ultimate Health" Then
                aRes = ExecuteInTibia("exiva ultimate_health_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo1 = "Spirit" Then
                aRes = ExecuteInTibia("exiva great_spirit_potion", idConnection, True)
                DoEvents
            ElseIf healingCheatsOptions(idConnection).Combo1 = "UH" Then
                aRes = ExecuteInTibia("exiva uh", idConnection, True)
                DoEvents
            End If
   End If
   
   
End If
Next idConnection

End Sub



Private Sub txtHealpot_change()

If IsNumeric(txtHealpot) = True Then
    ' ok
Else
    txtHealpot.Text = "0"
End If

End Sub

Private Sub txtHealthhi_change()

If IsNumeric(txtHealthhi) = True Then
    ' ok
Else
    txtHealthhi.Text = "0"
End If

End Sub

Private Sub txtHealthlo_change()

If IsNumeric(txtHealthlo) = True Then
    ' ok
Else
    txtHealthlo.Text = "0"
End If

End Sub


Private Sub txtMana_Validate(Cancel As Boolean)

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtMana = txtMana.Text
End If

End Sub

Private Sub txtManahi_change()

If IsNumeric(txtManahi) = True Then
    ' ok
Else
    txtManahi.Text = "0"
End If

End Sub

Private Sub txtManalo_change()

If IsNumeric(txtManalo) = True Then
    ' ok
Else
    txtManalo.Text = "0"
End If

End Sub

Private Sub txtManapot_change()

If IsNumeric(txtManapot) = True Then
    ' ok
Else
    txtManapot.Text = "0"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub


Private Sub txtPot_Validate(Cancel As Boolean)

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtPot = txtPot.Text
End If

End Sub

Private Sub txtSpellhi_Validate(Cancel As Boolean)

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtSpellhi = txtSpellhi.Text
End If

End Sub

Private Sub txtSpelllo_Validate(Cancel As Boolean)

If healingIDselected > 0 Then
  healingCheatsOptions(healingIDselected).txtSpelllo = txtSpelllo.Text
End If

End Sub
