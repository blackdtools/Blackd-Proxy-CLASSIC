VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blackd Proxy"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVIPsupport 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to VIP support page"
      Height          =   315
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdBroadcast 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2880
      Picture         =   "frmMenu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdNews 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   0
      Picture         =   "frmMenu.frx":2122
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdLaunchTibiaMC 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   7200
      Picture         =   "frmMenu.frx":3E17
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdLaunchTibia 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5760
      Picture         =   "frmMenu.frx":4DFE
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdStealth 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1440
      Picture         =   "frmMenu.frx":7A59
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAd 
      BackColor       =   &H00C0C000&
      Height          =   975
      Left            =   4320
      Picture         =   "frmMenu.frx":9089
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdHPmana 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2880
      Picture         =   "frmMenu.frx":9E86
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnknownFeature 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   7200
      Picture         =   "frmMenu.frx":AEF3
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdMagebomb 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   7200
      Picture         =   "frmMenu.frx":C26D
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdTrainer 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   4320
      Picture         =   "frmMenu.frx":D1CE
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdWarbot 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5760
      Picture         =   "frmMenu.frx":E0AC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdEvents 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   5760
      Picture         =   "frmMenu.frx":F64D
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdvanced 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1440
      Picture         =   "frmMenu.frx":104A6
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdHotkeys 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   4320
      Picture         =   "frmMenu.frx":114FE
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdStopAlarm 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2880
      Picture         =   "frmMenu.frx":12538
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutorial 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   0
      Picture         =   "frmMenu.frx":13570
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogs 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1440
      Picture         =   "frmMenu.frx":14480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheats 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   0
      Picture         =   "frmMenu.frx":151C1
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCavebot 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2880
      Picture         =   "frmMenu.frx":162CE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdRunemaker 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1440
      Picture         =   "frmMenu.frx":1741D
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdHardcoreCheats 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   0
      Picture         =   "frmMenu.frx":18748
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin JwldButn2b.JeweledButton cmdHealingNG 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Healing"
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
   Begin JwldButn2b.JeweledButton cmdExtrasNG 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Extras"
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
   Begin JwldButn2b.JeweledButton cmdCavebotNG 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Cavebot"
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
   Begin JwldButn2b.JeweledButton cmdAimbot 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Aimbot"
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
   Begin JwldButn2b.JeweledButton cmdWarbotNG 
      Height          =   300
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Friend List"
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
   Begin JwldButn2b.JeweledButton cmdStopAlarmNG 
      Height          =   300
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Stop Alarm"
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
   Begin JwldButn2b.JeweledButton cmdHotkeysNG 
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Hotkeys"
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
   Begin JwldButn2b.JeweledButton cmdTrainerNG 
      Height          =   300
      Left            =   2160
      TabIndex        =   8
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Trainer"
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
   Begin JwldButn2b.JeweledButton cmdUnknownFeatureNG 
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Conditions"
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
   Begin JwldButn2b.JeweledButton cmdBoradcastNG 
      Height          =   300
      Left            =   1080
      TabIndex        =   10
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Broadcast"
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
   Begin JwldButn2b.JeweledButton cmdLogsNG 
      Height          =   300
      Left            =   2160
      TabIndex        =   11
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Proxy"
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
   Begin JwldButn2b.JeweledButton cmdChangeStyle 
      Height          =   300
      Left            =   4320
      TabIndex        =   12
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Old Menu"
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
      BackColor       =   16777088
      BorderColor_Hover=   16761024
      BorderColor_Inner=   16777215
   End
   Begin JwldButn2b.JeweledButton cmdLaunchTibiaMCNG 
      Height          =   300
      Left            =   3240
      TabIndex        =   13
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Tibia MC"
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
   Begin JwldButn2b.JeweledButton cmdSave 
      Height          =   300
      Left            =   4320
      TabIndex        =   14
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Save"
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
   Begin JwldButn2b.JeweledButton cmdAdvancedNG 
      Height          =   300
      Left            =   5400
      TabIndex        =   15
      Top             =   330
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   529
      Caption         =   "Config"
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
   Begin JwldButn2b.JeweledButton cmdLoad 
      Height          =   300
      Left            =   4320
      TabIndex        =   16
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "Load"
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
   Begin JwldButn2b.JeweledButton cmdNewsNG 
      Height          =   300
      Left            =   5400
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   529
      Caption         =   "News"
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
   Begin JwldButn2b.JeweledButton ccmdPersistent 
      Height          =   300
      Left            =   5400
      TabIndex        =   45
      Top             =   0
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   529
      Caption         =   "Persist"
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
      Left            =   6480
      TabIndex        =   44
      Top             =   4320
      Width           =   2055
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
      Left            =   4320
      TabIndex        =   43
      Top             =   5160
      Width           =   2175
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
      Left            =   5400
      TabIndex        =   42
      Top             =   5400
      Width           =   1095
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
      Left            =   4320
      TabIndex        =   41
      Top             =   5400
      Width           =   975
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
      Left            =   4320
      TabIndex        =   40
      Top             =   4920
      Width           =   2295
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
      Left            =   4320
      TabIndex        =   39
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Close"
      End
      Begin VB.Menu mPopShowTibia 
         Caption         =   "&Show Tibia"
      End
      Begin VB.Menu mPopHideTibia 
         Caption         =   "&Hide Tibia"
      End
   End
End
Attribute VB_Name = "frmMenu"
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

Private Sub ccmdPersistent_Click()
  ' show persistent
  frmPersistent.WindowState = vbNormal
  frmPersistent.Show
  frmPersistent.SetFocus
  SetWindowPos frmPersistent.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
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
End Sub

Private Sub cmdAdvancedNG_Click()
  ' show Advanced form
  frmAdvanced.WindowState = vbNormal
  frmAdvanced.Show
  frmAdvanced.SetFocus
  SetWindowPos frmAdvanced.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdAimbot_Click()
  frmAimbot.WindowState = vbNormal
  frmAimbot.Show
  frmAimbot.SetFocus
  SetWindowPos frmAimbot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdBoradcastNG_Click()
  frmBroadcast.WindowState = vbNormal
  frmBroadcast.Show
  frmBroadcast.SetFocus
  SetWindowPos frmBroadcast.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdBroadcast_Click()
  frmBroadcast.WindowState = vbNormal
  frmBroadcast.Show
  frmBroadcast.SetFocus
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
End Sub

Private Sub cmdCavebotNG_Click()
  ' show cavebot form
  frmCavebot.WindowState = vbNormal
  frmCavebot.Show
  frmCavebot.SetFocus
  SetWindowPos frmCavebot.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdChangeStyle_Click()
  ' show old skin
  frmOld.WindowState = vbNormal
  frmOld.Show
  frmOld.SetFocus
  SetWindowPos frmOld.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
  Me.Hide
End Sub

Private Sub cmdCheats_Click()
  ' show Tools form
  frmCheats.WindowState = vbNormal
  frmCheats.Show
  frmCheats.SetFocus
End Sub

Private Sub cmdEvents_Click()
  frmEvents.WindowState = vbNormal
  frmEvents.Show
  frmEvents.SetFocus
End Sub

Private Sub cmdExtrasNG_Click()
  ' show extrasNG
  frmExtras.WindowState = vbNormal
  frmExtras.Show
  frmExtras.SetFocus
  SetWindowPos frmExtras.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdHardcoreCheats_Click()
  ' show Cheats form
  frmHardcoreCheats.WindowState = vbNormal
  frmHardcoreCheats.Show
  frmHardcoreCheats.SetFocus
End Sub

Private Sub cmdHealingNG_Click()
  ' show healingNG
  frmHealing.WindowState = vbNormal
  frmHealing.Show
  frmHealing.SetFocus
  SetWindowPos frmHealing.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub cmdHotkeys_Click()
  ' show hotkeys
  frmHotkeys.WindowState = vbNormal
  frmHotkeys.Show
  frmHotkeys.SetFocus
End Sub

Private Sub cmdHotkeysNG_Click()
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

Private Sub cmdLaunchTibiaMCNG_Click()
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
End Sub

Private Sub cmdLogsNG_Click()
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
End Sub

Private Sub cmdNewsNG_Click()
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
End Sub

Private Sub cmdSave_Click()
Dim aRes As Long
Dim idConnection As Integer
Dim i As Integer

    For i = 1 To MAXCLIENTS
        idConnection = i
        aRes = ExecuteInTibia("exiva save", idConnection, True)
        DoEvents
    Next i
    
End Sub

Private Sub cmdStealth_Click()
  frmStealth.WindowState = vbNormal
  frmStealth.Show
  frmStealth.SetFocus
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
    'custom ng
    'DangerPK(idConnection) = False
    'extrasOptions(extrasIDselected).chkPM = False
    frmRunemaker.chkCloseSound.Value = 0
    frmHardcoreCheats.chkRuneAlarm.Value = 0
    frmRunemaker.ChkDangerSound.Value = 0
    frmEvents.chkReconnectionAlarm.Value = 0
    frmCavebot.chkChangePkHeal.Value = 0
    
  Next mcid
  ChangePlayTheDangerSound False
End Sub

Private Sub cmdTrainer_Click()
  frmTrainer.WindowState = vbNormal
  frmTrainer.Show
  frmTrainer.SetFocus
End Sub

Private Sub cmdTrainerNG_Click()
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
End Sub

Private Sub cmdUnknownFeatureNG_Click()
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
End Sub





Private Sub cmdWarbotNG_Click()
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

Private Sub Form_Load()
  Dim pok As Boolean
  If thisShouldNotBeLoading = 0 Then
    Unload Me
    Exit Sub
  End If
  
  If TibiaVersionLong >= 841 Then
    frmMenu.cmdMagebomb.enabled = False
    frmOld.cmdMagebomb.enabled = False
  End If
  
  'If IamAdmin = True Then
  '  lblAdminInfo.Caption = "Running as admin: " & App.EXEName & ".exe"
  'Else
  '  lblAdminInfo.Caption = "Running as user: " & App.EXEName & ".exe"
  'End If
  Me.Caption = frmMain.Caption
  frmMain.Caption = "Proxy (connection and logs)"

  CornerMessage = "If you purchased us any gold in the last month, we give you VIP support"
 
  Label3.Caption = CornerMessage
  Label3.ForeColor = CornerColor
  ApplyLimits
  Me.Show
  Me.Refresh
  With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Blackd Proxy" & vbNullChar
  End With
  Shell_NotifyIcon NIM_ADD, nid
  DoEvents
  If FirstExecute = True Then
    cmdTutorial_Click
  End If
  pok = True
  If MyPriorityID <> 2 Then
    pok = UpdateMyPriority()
  End If
  If (TibiaPriorityID <> 2) And (pok = True) Then
    pok = UpdateTibiaPriority()
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
         Single, y As Single)
  'this procedure receives the callbacks from the System Tray icon.
  Dim result As Long
  Dim msg As Long
  'the value of X will vary depending upon the scalemode setting
  If Me.ScaleMode = vbPixels Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If
  
  Select Case msg
  Case WM_LBUTTONUP        '514 restore form window
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Hide
    Me.Show
  Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Show
  Case WM_RBUTTONUP        '517 display popup menu
    result = SetForegroundWindow(Me.hwnd)
    Me.PopupMenu Me.mPopupSys
  End Select
End Sub

Private Sub Form_Resize()
  ' this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then
    Me.Hide
  End If
End Sub

Private Sub JeweledButton16_Click()

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
    'custom ng
    SetWindowPos frmConfirm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
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

