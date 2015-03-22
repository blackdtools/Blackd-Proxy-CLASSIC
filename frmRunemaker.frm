VERSION 5.00
Begin VB.Form frmRunemaker 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runemaker"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   Icon            =   "frmRunemaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNoHands 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Tibia 8.72+"
      ForeColor       =   &H00FFFFFF&
      Height          =   535
      Left            =   120
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton NoHands 
         BackColor       =   &H00000000&
         Caption         =   "Don't move runes to hands"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.TextBox txrRunemakerChaos2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   44
      Text            =   "10000"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveRunemakerChaos 
      Caption         =   "Change"
      Height          =   285
      Left            =   2880
      TabIndex        =   41
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txrRunemakerChaos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2280
      TabIndex        =   40
      Text            =   "600"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Timer timerSS 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4800
      Top             =   4920
   End
   Begin VB.CheckBox chkOnDangerSS 
      BackColor       =   &H00000000&
      Caption         =   "On danger screenshot"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   39
      Top             =   3210
      Width           =   2055
   End
   Begin VB.TextBox txtLowMana 
      Height          =   285
      Left            =   3240
      TabIndex        =   38
      Text            =   "100"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkManaFluid 
      BackColor       =   &H00000000&
      Caption         =   "Drink a mana fluid if mana is less than "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CheckBox chkmsgSound2 
      BackColor       =   &H00000000&
      Caption         =   "Play player.wav when an unfriendly player-creature pop on screen"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   36
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CheckBox chkmsgSound 
      BackColor       =   &H00000000&
      Caption         =   "Play ding.wav at messages"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdStopAlarm 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stop alarms !"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox chkCloseSound 
      BackColor       =   &H00000000&
      Caption         =   "Close = danger too"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   5160
      TabIndex        =   33
      Text            =   "fri.txt"
      ToolTipText     =   "Load - save file name"
      Top             =   1820
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000C000&
      Caption         =   "Load"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Loads from given file"
      Top             =   1820
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Saves to given file"
      Top             =   1820
      Width           =   615
   End
   Begin VB.OptionButton UseLeftHand 
      BackColor       =   &H00000000&
      Caption         =   "Always make runes in LEFT hand"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Width           =   3015
   End
   Begin VB.OptionButton UseRightHand 
      BackColor       =   &H00000000&
      Caption         =   "Always make runes in RIGHT hand"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton cmdDebug 
      BackColor       =   &H0080FF80&
      Caption         =   "DEBUG"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1800
      Width           =   855
   End
   Begin VB.Timer TimerMaker 
      Interval        =   400
      Left            =   3960
      Top             =   5280
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apply changes"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemoveFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remove sel"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtAddFriend 
      Height          =   285
      Left            =   3960
      TabIndex        =   19
      Text            =   "friendNameHere"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ListBox lstFriends 
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox chkLogoutOutRunes 
      BackColor       =   &H00000000&
      Caption         =   "Auto logout if out of runes or soulpoints"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CheckBox chkLogoutDangerCurrent 
      BackColor       =   &H00000000&
      Caption         =   "Auto logout if danger in current floor"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CheckBox chkLogoutDangerAny 
      BackColor       =   &H00000000&
      Caption         =   "Auto logout if danger in any floor"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CheckBox chkFood 
      BackColor       =   &H00000000&
      Caption         =   "Auto eat food for this character"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtSoulAction2 
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Text            =   "3"
      Top             =   6090
      Width           =   495
   End
   Begin VB.TextBox txtManaAction2 
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Text            =   "400"
      Top             =   6090
      Width           =   735
   End
   Begin VB.TextBox txtAction2 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "adura vita"
      Top             =   6090
      Width           =   1695
   End
   Begin VB.TextBox txtManaAction1 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "25"
      Top             =   5250
      Width           =   735
   End
   Begin VB.TextBox txtAction1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "exura"
      Top             =   5250
      Width           =   1695
   End
   Begin VB.CheckBox chkActivate 
      BackColor       =   &H00000000&
      Caption         =   "Activate runemaker for this character"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "-"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chkWaste 
      BackColor       =   &H00000000&
      Caption         =   "If out of runes or out of soulpoints then waste mana with this spell:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   4680
      Width           =   5175
   End
   Begin VB.CheckBox ChkDangerSound 
      BackColor       =   &H00000000&
      Caption         =   "On danger play sound"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   2920
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "<- NOT RECOMMENDED"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   45
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   " After enough mana wait up to... (ms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Chaos (ms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   42
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Important: you should have free space in a container called <something>BACKPACK"
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
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   3735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   3720
      X2              =   0
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   6000
      X2              =   3720
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   3720
      X2              =   3720
      Y1              =   1680
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   3960
      X2              =   5760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblGlobal 
      BackColor       =   &H00000000&
      Caption         =   "GLOBAL (applies for all)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Add exception:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblFriends 
      BackColor       =   &H00000000&
      Caption         =   "Don't consider following names as danger:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "soulpoint cost :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblMana2 
      BackColor       =   &H00000000&
      Caption         =   "mana cost :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblAction2 
      BackColor       =   &H00000000&
      Caption         =   "Make this rune if enough mana / soulpoints :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label lblMana1 
      BackColor       =   &H00000000&
      Caption         =   "mana cost :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmRunemaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Public Sub UpdateValues()
  If runemakerIDselected = 0 Then
    If RuneMakerOptions_activated_default = True Then
      chkActivate.Value = 1
    Else
      chkActivate.Value = 0
    End If
    If RuneMakerOptions_autoEat_default = True Then
      chkFood.Value = 1
    Else
      chkFood.Value = 0
    End If
    If RuneMakerOptions_ManaFluid_default = True Then
      chkManaFluid.Value = 1
    Else
      chkManaFluid.Value = 0
    End If
    If RuneMakerOptions_autoLogoutAnyFloor_default = True Then
      chkLogoutDangerAny.Value = 1
    Else
      chkLogoutDangerAny.Value = 0
    End If
    If RuneMakerOptions_autoLogoutCurrentFloor_default = True Then
      chkLogoutDangerCurrent.Value = 1
    Else
      chkLogoutDangerCurrent.Value = 0
    End If
    If RuneMakerOptions_autoLogoutOutOfRunes_default = True Then
      chkLogoutOutRunes.Value = 1
    Else
      chkLogoutOutRunes.Value = 0
    End If
    If RuneMakerOptions_autoWaste_default = True Then
      chkWaste.Value = 1
    Else
      chkWaste.Value = 0
    End If
    If RuneMakerOptions_msgSound_default = True Then
      chkmsgSound.Value = 1
    Else
      chkmsgSound.Value = 0
    End If
    If RuneMakerOptions_msgSound2_default = True Then
      chkmsgSound2.Value = 1
    Else
      chkmsgSound2.Value = 0
    End If
    txtAction1.Text = RuneMakerOptions_firstActionText_default
    txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
    txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
    txtAction2.Text = RuneMakerOptions_secondActionText_default
    txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
    txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
  Else
    If RuneMakerOptions(runemakerIDselected).activated = True Then
      chkActivate.Value = 1
    Else
      chkActivate.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoEat = True Then
      chkFood.Value = 1
    Else
      chkFood.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).ManaFluid = True Then
      chkManaFluid.Value = 1
    Else
      chkManaFluid.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = True Then
      chkLogoutDangerAny.Value = 1
    Else
      chkLogoutDangerAny.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = True Then
      chkLogoutDangerCurrent.Value = 1
    Else
      chkLogoutDangerCurrent.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = True Then
      chkLogoutOutRunes.Value = 1
    Else
      chkLogoutOutRunes.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).autoWaste = True Then
      chkWaste.Value = 1
    Else
      chkWaste.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).msgSound = True Then
      chkmsgSound.Value = 1
    Else
      chkmsgSound.Value = 0
    End If
    If RuneMakerOptions(runemakerIDselected).msgSound2 = True Then
      chkmsgSound2.Value = 1
    Else
      chkmsgSound2.Value = 0
    End If
    txtAction1.Text = RuneMakerOptions(runemakerIDselected).firstActionText
    txtManaAction1.Text = CStr(RuneMakerOptions(runemakerIDselected).firstActionMana)
    txtLowMana.Text = CStr(RuneMakerOptions(runemakerIDselected).LowMana)
    txtAction2.Text = RuneMakerOptions(runemakerIDselected).secondActionText
    txtManaAction2.Text = CStr(RuneMakerOptions(runemakerIDselected).secondActionMana)
    txtSoulAction2.Text = CStr(RuneMakerOptions(runemakerIDselected).secondActionSoulpoints)
  End If
End Sub

Public Sub SetChk(typeChk As String, v As Integer)
  Select Case typeChk
  Case "chkActivate"
    lock_chkActivate = True
    chkActivate.Value = v
    lock_chkActivate = False
    
  Case "chkFood"
    lock_chkFood = True
    chkFood.Value = v
    lock_chkFood = False
    
  Case "chkManaFluid"
    lock_chkManaFluid = True
    chkManaFluid.Value = v
    lock_chkManaFluid = False
    
  Case "chkLogoutDangerAny"
    lock_chkLogoutDangerAny = True
    chkLogoutDangerAny.Value = v
    lock_chkLogoutDangerAny = False
    
  Case "chkLogoutDangerCurrent"
    lock_chkLogoutDangerCurrent = True
    chkLogoutDangerCurrent.Value = v
    lock_chkLogoutDangerCurrent = False
    
  Case "chkLogoutOutRunes"
    lock_chkLogoutOutRunes = True
    chkLogoutOutRunes.Value = v
    lock_chkLogoutOutRunes = False
    
  Case "chkWaste"
    lock_chkWaste = True
    chkWaste.Value = v
    lock_chkWaste = False
    
  Case "chkmsgSound"
    lock_chkmsgSound = True
    chkmsgSound.Value = v
    lock_chkmsgSound = False
  
  Case "chkmsgSound2"
    lock_chkmsgSound2 = True
    chkmsgSound2.Value = v
    lock_chkmsgSound2 = False
  End Select
End Sub

Public Sub DisableAll(id As Integer)
  If id = CInt(runemakerIDselected) Then
    SetChk "chkActivate", 0
    SetChk "chkFood", 0
    SetChk "chkManaFluid", 0
    SetChk "chkLogoutDangerAny", 0
    SetChk "chkLogoutDangerCurrent", 0
    SetChk "chkLogoutOutRunes", 0
    SetChk "chkWaste", 0
    SetChk "chkmsgSound", 0
    SetChk "chkmsgSound2", 0
  End If
  RuneMakerOptions(id).activated = False
  RuneMakerOptions(id).autoEat = False
  RuneMakerOptions(id).ManaFluid = False
  RuneMakerOptions(id).autoLogoutAnyFloor = False
  RuneMakerOptions(id).autoLogoutCurrentFloor = False
  RuneMakerOptions(id).autoLogoutOutOfRunes = False
  RuneMakerOptions(id).autoWaste = False
  RuneMakerOptions(id).msgSound = False
  RuneMakerOptions(id).msgSound2 = False
End Sub
Private Sub chkActivate_Click()
Dim tileID As Long
Dim aRes As Long
#If FinalMode Then
On Error GoTo goterr
#End If
If lock_chkActivate = False Then
If runemakerIDselected > 0 Then
  If chkActivate.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).activated = True
    If TibiaVersionLong >= 872 Then
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t3
      End If
      aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started.")
      DoEvents
    Else
    If UseRightHand.Value = True Then
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on right hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2))
        DoEvents
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_RIGHTHAND).t3
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on right hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2) & " (with amount byte " & GoodHex(savedItem(runemakerIDselected).t3) & " )")
        DoEvents
      End If
    Else
      savedItem(runemakerIDselected).t1 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t1
      savedItem(runemakerIDselected).t2 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t2
      tileID = GetTheLong(savedItem(runemakerIDselected).t1, savedItem(runemakerIDselected).t2)
      If DatTiles(tileID).stackable = False Then
        savedItem(runemakerIDselected).t3 = 0
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on left hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2))
        DoEvents
      Else
        savedItem(runemakerIDselected).t3 = mySlot(runemakerIDselected, SLOT_LEFTHAND).t3
        aRes = SendLogSystemMessageToClient(CInt(runemakerIDselected), "Runemaker started. Saved your current item on left hand : " & GoodHex(savedItem(runemakerIDselected).t1) & " " & GoodHex(savedItem(runemakerIDselected).t2) & " (with amount byte  " & GoodHex(savedItem(runemakerIDselected).t3) & " )")
        DoEvents
      End If
    End If
    End If
  Else
    RuneMakerOptions(runemakerIDselected).activated = False
  End If
End If
End If
Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Warning: connection fail during the runemaker activation - ignoring"
End Sub





Private Sub chkFood_Click()
If lock_chkFood = False Then
If runemakerIDselected > 0 Then
  If chkFood.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoEat = True
  Else
    RuneMakerOptions(runemakerIDselected).autoEat = False
  End If
End If
End If
End Sub

Private Sub chkLogoutDangerAny_Click()
If lock_chkLogoutDangerAny = False Then
If runemakerIDselected > 0 Then
  If chkLogoutDangerAny.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = True
    SetChk "chkLogoutDangerCurrent", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = False
  End If
End If
End If
End Sub

Private Sub chkLogoutDangerCurrent_Click()
If lock_chkLogoutDangerCurrent = False Then
If runemakerIDselected > 0 Then
  If chkLogoutDangerCurrent.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = True
    SetChk "chkLogoutDangerAny", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = False
  End If
End If
End If
End Sub

Private Sub chkLogoutOutRunes_Click()
If lock_chkLogoutOutRunes = False Then
If runemakerIDselected > 0 Then
  If chkLogoutOutRunes.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = True
    SetChk "chkWaste", 0
    RuneMakerOptions(runemakerIDselected).autoWaste = False
  Else
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = False
  End If
End If
End If
End Sub

Private Sub chkManaFluid_Click()
If lock_chkManaFluid = False Then
If runemakerIDselected > 0 Then
  If chkManaFluid.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).ManaFluid = True
  Else
    RuneMakerOptions(runemakerIDselected).ManaFluid = False
    RemoveSpamOrder CInt(runemakerIDselected), 4 'remove auto mana
  End If
End If
End If
End Sub

Private Sub chkmsgSound2_Click()
If lock_chkmsgSound2 = False Then
If runemakerIDselected > 0 Then
  If chkmsgSound2.Value = 1 Then
    DangerPlayer(runemakerIDselected) = False
    RuneMakerOptions(runemakerIDselected).msgSound2 = True
  Else
    DangerPlayer(runemakerIDselected) = False
    RuneMakerOptions(runemakerIDselected).msgSound2 = False
  End If
End If
End If
End Sub

Private Sub chkWaste_Click()
If lock_chkWaste = False Then
If runemakerIDselected > 0 Then
  If chkWaste.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).autoWaste = True
    SetChk "chkLogoutOutRunes", 0
    RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = False
  Else
    RuneMakerOptions(runemakerIDselected).autoWaste = False
  End If
End If
End If
End Sub


Private Sub chkmsgSound_Click()
If lock_chkmsgSound = False Then
If runemakerIDselected > 0 Then
  If chkmsgSound.Value = 1 Then
    RuneMakerOptions(runemakerIDselected).msgSound = True
  Else
    RuneMakerOptions(runemakerIDselected).msgSound = False
  End If
End If
End If
End Sub



Private Sub cmbCharacter_Click()
 runemakerIDselected = cmbCharacter.ListIndex
  If runemakerIDselected > 0 Then
      UpdateValues
  End If
End Sub
Public Function IsFriend(strName As String) As Boolean
  'strname should come in lcase
  Dim i As Long
  Dim totI As Long
  Dim foundI As Long
  totI = lstFriends.ListCount - 1
  foundI = -1
  For i = 0 To totI
    If lstFriends.List(i) = strName Then
      foundI = i
    End If
  Next i
  If foundI = -1 Then
    IsFriend = False
  Else
    IsFriend = True
  End If
End Function

Private Sub cmdAddFriend_Click()
  If IsFriend(LCase(txtAddFriend.Text)) = False Then
    lstFriends.AddItem LCase(txtAddFriend.Text)
  End If
End Sub

Private Sub cmdApply_Click()

    UpdateValues
 
End Sub

Private Sub cmdDebug_Click()
  Dim aRes As Long
  Dim i As Long
  If runemakerIDselected > 0 Then
    publicDebugMode = Not publicDebugMode
    If publicDebugMode = True Then
      aRes = GiveGMmessage(CInt(runemakerIDselected), "DEBUG MODE ENABLED", "Blackd")
      For i = 1 To MAXCLIENTS
        If GameConnected(i) = True Then
          frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Process ID of client #" & CStr(i) & " (" & CharacterName(i) & ") =" & CStr(ProcessID(i))
        End If
      Next i
      DoEvents
    Else
      aRes = GiveGMmessage(CInt(runemakerIDselected), "DEBUG MODE DISABLED", "Blackd")
      DoEvents
    End If
  End If
End Sub

Private Sub cmdLoad_Click()
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim filename As String
  Set fso = New scripting.FileSystemObject
    lstFriends.Clear
    filename = App.path & "\" & txtFile.Text
    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        If strLine <> "" Then
        If IsFriend(LCase(strLine)) = False Then
          lstFriends.AddItem LCase(strLine)
        End If
        End If
      Wend
      Close #fn
    End If
End Sub

Private Sub cmdRemoveFriend_Click()
  If lstFriends.ListIndex > -1 Then
    lstFriends.RemoveItem (lstFriends.ListIndex)
  End If
End Sub

Private Sub cmdSave_Click()
  Dim fn As Integer
  Dim limI As Long
  Dim i As Long
    limI = lstFriends.ListCount - 1
    fn = FreeFile
    Open App.path & "\" & txtFile.Text For Output As #fn
    For i = 0 To limI
      Print #fn, lstFriends.List(i)
    Next i
    Close #fn
End Sub

Private Sub cmdSaveRunemakerChaos_Click()
    On Error GoTo goterr
    Dim lngCast As Long
    Dim lngCast2 As Long
    lngCast = CLng(frmRunemaker.txrRunemakerChaos.Text)
    lngCast2 = CLng(frmRunemaker.txrRunemakerChaos2.Text)
    If (lngCast >= 20) And (lngCast2 >= 0) Then
        RunemakerChaos = lngCast
        RunemakerChaos2 = lngCast2
        Me.txrRunemakerChaos.Text = CStr(RunemakerChaos)
        Me.txrRunemakerChaos2.Text = CStr(RunemakerChaos2)
        frmRunemaker.Caption = "Runemaker - chaos updated"
    Else
        GoTo goterr
    End If
    Exit Sub
goterr:
    frmRunemaker.Caption = "Runemaker - invalid chaos values"
End Sub

Private Sub cmdStopAlarm_Click()
  Dim mcid As Integer
  For mcid = 1 To MAXCLIENTS
    DangerPK(mcid) = False
    DangerGM(mcid) = False
    LogoutTimeGM(mcid) = 0
    moveRetry(mcid) = 0
    RemoveSpamOrder mcid, 1
    UHRetryCount(mcid) = 0
  Next mcid
  ChangePlayTheDangerSound False
End Sub



Private Sub Form_Load()
LoadRuneChars
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub
Public Sub LoadRuneChars()
  Dim i As Long
  Dim firstC As Long
  If TibiaVersionLong >= 872 Then
    UseRightHand.Visible = False
    UseLeftHand.Visible = False
    fraNoHands.Visible = True
    Label2.Caption = "IMPORTANT: Blackd Proxy will only count runes displayed in opened backpacks!"
  End If
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
  runemakerIDselected = firstC
  UpdateValues
End Sub

Private Sub TimerMaker_Timer()
  Dim resS As TypeSearchItemResult2
  Dim idConnection As Integer
  Dim aRes As Long
  Dim sCheat As String
  Dim cPacket() As Byte
  Dim inRes As Integer
  Dim cond1 As Boolean
  Dim cond2 As Boolean
  Dim cond3 As Boolean
  Dim cond4 As Boolean
  Dim cond5 As Boolean
  Dim tmpcond As Boolean
  Dim playerS As String
  Dim gtc As Long
  Dim eatDone() As Boolean
  Dim i As Integer
  #If FinalMode Then
  On Error GoTo errclose
  #End If
  'add chaos to the timer
  TimerMaker.Interval = 400 + randomNumberBetween(0, RunemakerChaos)
      
  ' EAT FOOD EACH 60 TURNS (1 turn = 400ms)
  ' when eating food > do nothing else in this mc
  ReDim eatDone(1 To MAXCLIENTS)
  For i = 1 To MAXCLIENTS
    eatDone(i) = False
  Next i
  For idConnection = 1 To MAXCLIENTS
    If (runeTurn(idConnection) > 59) Then
        If ((GameConnected(idConnection) = True) And (sentWelcome(idConnection) = True) And (GotPacketWarning(idConnection) = False)) Then
          If makingRune(idConnection) = False Then
            If CheatsPaused(idConnection) = False Then
              If RuneMakerOptions(idConnection).autoEat = True Then
               ' We are allowed to eat.
               ' Lets search food...
                resS = SearchFood(idConnection)
                If (resS.foundcount > 0) Then
                  ' Food found, eat it now
                  aRes = EatFood(idConnection, resS.b1, resS.b2, resS.bpID, resS.slotID)
                  DoEvents
                End If
              End If
            End If
          End If
        End If
        runeTurn(idConnection) = randomNumberBetween(0, 29)
        eatDone(idConnection) = True
    Else
      runeTurn(idConnection) = runeTurn(idConnection) + 1
    End If
  Next idConnection
  
  gtc = GetTickCount()
  For idConnection = 1 To MAXCLIENTS
    If (eatDone(idConnection) = False) Then
    If (CheatsPaused(idConnection) = False) And (GameConnected(idConnection) = True) And _
       (sentWelcome(idConnection) = True) And _
     (GotPacketWarning(idConnection) = False) And (DangerGM(idConnection) = False) And _
     (gtc > lootTimeExpire(idConnection)) Then
    If TibiaVersionLong >= 872 Then ' do not move runes to hand!
      If RuneMakerOptions(idConnection).activated = True Then
        resS = SearchItem(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank))
        cond1 = ((mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2)) Or _
          ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2))
        If (resS.foundcount = 0) Then
          tmpcond = True
        Else
          tmpcond = False
        End If
        cond2 = (tmpcond = True) And (cond1 = False)
        cond3 = mySoulpoints(idConnection) < RuneMakerOptions(idConnection).secondActionSoulpoints
        If (cond2 Or cond3) Then
            ' can't make rune
            makingRune(idConnection) = False 'not making rune mode
            runemakerMana1(idConnection) = -1
            If RuneMakerOptions(idConnection).autoLogoutOutOfRunes = True Then
              If ReconnectionStage(idConnection) = 0 Then
                frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " did runemaker logout - logged out because : out of runes or soulpoints"
                sCheat = "14"
                SafeCastCheatString "TimerMaker1", idConnection, sCheat
                aRes = GiveServerError("Runemaker logout - logged out because : out of runes or soulpoints", idConnection)
                DoEvents
                frmMain.DoCloseActions idConnection
                DoEvents
              End If
            ' waste mana option
            ElseIf (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              End If
            End If
        Else
            If runemakerMana1(idConnection) = -1 Then
              runemakerMana1(idConnection) = gtc + randomNumberBetween(0, RunemakerChaos2)
            End If
            If gtc >= runemakerMana1(idConnection) Then
                If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
                    If ((runeTurn(idConnection) Mod 5) = 0) Then
                  ' make the rune now!
                  aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                  makingRune(idConnection) = False 'all is ok again, not making rune mode
                  runemakerMana1(idConnection) = -1
                    End If
                End If
            End If
        End If
      End If
    Else ' old mode
      If RuneMakerOptions(idConnection).activated = True Then
          resS = SearchItem(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank))
          cond1 = ((mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2)) Or _
            ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2))
          If (resS.foundcount = 0) Then
            tmpcond = True
          Else
            tmpcond = False
          End If
          cond2 = (tmpcond = True) And (cond1 = False)
          cond3 = mySoulpoints(idConnection) < RuneMakerOptions(idConnection).secondActionSoulpoints
          
          cond4 = (UseRightHand.Value = True) And _
             (Not ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = savedItem(idConnection).t1) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = savedItem(idConnection).t2)))
            
          cond5 = (UseLeftHand.Value = True) And _
             (Not ((mySlot(idConnection, SLOT_LEFTHAND).t1 = savedItem(idConnection).t1) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = savedItem(idConnection).t2)))
                
          If (cond2 Or cond3) And (Not (cond4 Or cond5)) Then
            ' out of runes or soulpoints
            ' logout option
            makingRune(idConnection) = False
            If RuneMakerOptions(idConnection).autoLogoutOutOfRunes = True Then
              If ReconnectionStage(idConnection) = 0 Then
                frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & CharacterName(idConnection) & " did runemaker logout - logged out because : out of runes or soulpoints"
                sCheat = "14"
                SafeCastCheatString "TimerMaker2", idConnection, sCheat

                aRes = GiveServerError("Runemaker logout - logged out because : out of runes or soulpoints", idConnection)
                DoEvents
                frmMain.DoCloseActions idConnection
                DoEvents
              End If
            ' waste mana option
            ElseIf (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              End If
            End If
          Else
          
            If runemakerMana1(idConnection) = -1 Then
              runemakerMana1(idConnection) = gtc + randomNumberBetween(0, RunemakerChaos2)
            End If
            If gtc >= runemakerMana1(idConnection) Then
          
            If UseRightHand.Value = True Then
             ' right hand
            If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
              If mySlot(idConnection, SLOT_RIGHTHAND).t1 = blank1 And mySlot(idConnection, SLOT_RIGHTHAND).t2 = blank2 Then
                If ((runeTurn(idConnection) Mod 5) = 0) Then
                ' make the rune now!
               ' MODIFIFIED IN 9.35 to allow exiva testsound
                aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                End If
              ElseIf ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = &H0) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = &H0)) Then
                'move blank rune to right hand
                initialRuneBackpack(idConnection) = resS.bpID
                aRes = MoveItemToRightHand(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank), 0, resS.bpID, resS.slotID, False)
              Else
                makingRune(idConnection) = True
  
                aRes = SaveHand(idConnection, True, CByte(SLOT_RIGHTHAND), initialRuneBackpack(idConnection))
                If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
                End If
                DoEvents
                
              End If
            ElseIf ((mySlot(idConnection, SLOT_RIGHTHAND).t1 = 0) And (mySlot(idConnection, SLOT_RIGHTHAND).t2 = 0)) Then
                aRes = MoveItemToRightHand(idConnection, savedItem(idConnection).t1, savedItem(idConnection).t2, savedItem(idConnection).t3, 0, 0, True)
                
            ElseIf (Not (mySlot(idConnection, SLOT_RIGHTHAND).t1 = savedItem(idConnection).t1 And mySlot(idConnection, SLOT_RIGHTHAND).t2 = savedItem(idConnection).t2)) Then
              ' put made rune in backpack
              aRes = SaveHand(idConnection, False, CByte(SLOT_RIGHTHAND), initialRuneBackpack(idConnection))
              If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
              End If
              DoEvents
            Else
              makingRune(idConnection) = False 'all is ok again, not making rune mode
              runemakerMana1(idConnection) = -1
            End If
            
            
            Else
              'left hand
            If myMana(idConnection) >= RuneMakerOptions(idConnection).secondActionMana Then
              If mySlot(idConnection, SLOT_LEFTHAND).t1 = blank1 And mySlot(idConnection, SLOT_LEFTHAND).t2 = blank2 Then
                If ((runeTurn(idConnection) Mod 5) = 0) Then
                ' make the rune now!
               ' MODIFIFIED IN 9.35 to allow exiva testsound
                aRes = ExecuteInTibia(RuneMakerOptions(idConnection).secondActionText, idConnection, True)
                End If
              ElseIf ((mySlot(idConnection, SLOT_LEFTHAND).t1 = &H0) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = &H0)) Then
               'move blank rune to left hand
                initialRuneBackpack(idConnection) = resS.bpID
                aRes = MoveItemToLeftHand(idConnection, LowByteOfLong(tileID_Blank), HighByteOfLong(tileID_Blank), 0, resS.bpID, resS.slotID, False)
                
              Else
               makingRune(idConnection) = True

                aRes = SaveHand(idConnection, True, CByte(SLOT_LEFTHAND), initialRuneBackpack(idConnection))
                If (aRes = -1) Then
                  makingRune(idConnection) = False
                  runemakerMana1(idConnection) = -1
                End If
                DoEvents
              End If
            ElseIf ((mySlot(idConnection, SLOT_LEFTHAND).t1 = 0) And (mySlot(idConnection, SLOT_LEFTHAND).t2 = 0)) Then
                aRes = MoveItemToLeftHand(idConnection, savedItem(idConnection).t1, savedItem(idConnection).t2, savedItem(idConnection).t3, 0, 0, True)
                
            ElseIf (Not (mySlot(idConnection, SLOT_LEFTHAND).t1 = savedItem(idConnection).t1 And mySlot(idConnection, SLOT_LEFTHAND).t2 = savedItem(idConnection).t2)) Then
              ' put made rune in backpack
              aRes = SaveHand(idConnection, False, CByte(SLOT_LEFTHAND), initialRuneBackpack(idConnection))
              If (aRes = -1) Then
                makingRune(idConnection) = False
                runemakerMana1(idConnection) = -1
              End If
              DoEvents
            Else
              makingRune(idConnection) = False 'all is ok again, not making rune mode
              runemakerMana1(idConnection) = -1
            End If
            
            End If
            
            End If
          End If
      Else
            If (RuneMakerOptions(idConnection).autoWaste = True) And (myMana(idConnection) >= RuneMakerOptions(idConnection).firstActionMana) Then
              If ((runeTurn(idConnection) Mod 10) = 0) Then
              ' MODIFIFIED IN 9.35 to allow exiva testsound
              aRes = ExecuteInTibia(RuneMakerOptions(idConnection).firstActionText, idConnection, True)
              End If
            End If
      End If
      
    End If ' tibia version
    End If
    End If ' eat done=false
  Next idConnection
  Exit Sub
errclose:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Warning: connection fail during the runemaker function - ignoring"
End Sub

Private Sub timerSS_Timer()
    timerSS.enabled = False
    GetScreenshot frmScreenshot, getScreenshotname()
    DoEvents
End Sub



Private Sub txtAction1_Validate(Cancel As Boolean)
If runemakerIDselected > 0 Then
  RuneMakerOptions(runemakerIDselected).firstActionText = txtAction1.Text
End If
End Sub
Private Sub txtAction2_Validate(Cancel As Boolean)
If runemakerIDselected > 0 Then
  RuneMakerOptions(runemakerIDselected).secondActionText = txtAction2.Text
End If
End Sub



Private Sub txtLowMana_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtLowMana.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).LowMana = lonN
  Else
    txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
    RuneMakerOptions(runemakerIDselected).LowMana = RuneMakerOptions_LowMana_default
  End If
  End If
  Exit Sub
gotError:
  txtLowMana.Text = CStr(RuneMakerOptions_LowMana_default)
  RuneMakerOptions(runemakerIDselected).LowMana = RuneMakerOptions_LowMana_default
End Sub

Private Sub txtManaAction1_Validate(Cancel As Boolean)
  Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtManaAction1.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).firstActionMana = lonN
  Else
    txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
    RuneMakerOptions(runemakerIDselected).firstActionMana = RuneMakerOptions_firstActionMana_default
  End If
  End If
  Exit Sub
gotError:
  txtManaAction1.Text = CStr(RuneMakerOptions_firstActionMana_default)
  RuneMakerOptions(runemakerIDselected).firstActionMana = RuneMakerOptions_firstActionMana_default
End Sub
Private Sub txtManaAction2_Validate(Cancel As Boolean)
 Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  If runemakerIDselected > 0 Then
  lonN = CLng(txtManaAction2.Text)
  If lonN > 0 Then
    RuneMakerOptions(runemakerIDselected).secondActionMana = lonN
  Else
    txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
    RuneMakerOptions(runemakerIDselected).secondActionMana = RuneMakerOptions_secondActionMana_default
  End If
  End If
  Exit Sub
gotError:
  txtManaAction2.Text = CStr(RuneMakerOptions_secondActionMana_default)
  RuneMakerOptions(runemakerIDselected).secondActionMana = RuneMakerOptions_secondActionMana_default
End Sub
Private Sub txtSoulAction2_Validate(Cancel As Boolean)
Dim lonN As Long
  #If FinalMode Then
  On Error GoTo gotError
  #End If
  lonN = CLng(txtSoulAction2.Text)
  If runemakerIDselected > 0 Then
  If lonN >= 0 Then
    RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = lonN
  Else
    txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
    RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default
  End If
  End If
  Exit Sub
gotError:
  txtSoulAction2.Text = CStr(RuneMakerOptions_secondActionSoulpoints_default)
  RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = RuneMakerOptions_secondActionSoulpoints_default
End Sub

Private Sub UseLeftHand_Click()
  Dim aRes As Long
  Dim i As Integer
    #If FinalMode Then
  On Error GoTo gotError
  #End If
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True And GotPacketWarning(i) = False And RuneMakerOptions(i).activated = True Then
      DisableAll i
      aRes = GiveGMmessage(i, "Runemaker have been disabled because the change of hand option. You should reactivate it now.", "Blackd")
      DoEvents
    End If
  Next i
  Exit Sub
gotError:
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error at UseLeftHand_Click"
End Sub

Private Sub UseRightHand_Click()
  Dim aRes As Long
  Dim i As Integer
    #If FinalMode Then
  On Error GoTo gotError
  #End If
  For i = 1 To MAXCLIENTS
    If GameConnected(i) = True And GotPacketWarning(i) = False And RuneMakerOptions(i).activated = True Then
      DisableAll i
      aRes = GiveGMmessage(i, "Runemaker have been disabled because the change of hand option. You should reactivate it now.", "Blackd")
      DoEvents
    End If
  Next i
  Exit Sub
gotError:
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Error at UseRightHand_Click"
End Sub
