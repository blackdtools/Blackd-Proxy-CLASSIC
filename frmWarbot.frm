VERSION 5.00
Begin VB.Form frmWarbot 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warbot"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmWarbot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAutohealDelay2 
      Height          =   285
      Left            =   8760
      TabIndex        =   44
      Text            =   "700"
      Top             =   4800
      Width           =   615
   End
   Begin VB.OptionButton AutoHealOption3 
      BackColor       =   &H00000000&
      Caption         =   "Heal with exura sio ""friendName"""
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   43
      Top             =   6000
      Width           =   4215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear"
      Height          =   255
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   7800
      TabIndex        =   41
      Text            =   "wargroups\autoheal.txt"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveAutoheal 
      BackColor       =   &H0000C000&
      Caption         =   "SAVE"
      Height          =   255
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdReloadAutoHealFriends 
      BackColor       =   &H0000C000&
      Caption         =   "LOAD"
      Height          =   255
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtAddFriend 
      Height          =   285
      Left            =   5280
      TabIndex        =   36
      Text            =   "friendNameHere"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add"
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdRemoveFriend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remove sel"
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Timer timerFriendHealer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9120
      Top             =   120
   End
   Begin VB.TextBox txtAutohealDelay 
      Height          =   285
      Left            =   7680
      TabIndex        =   32
      Text            =   "300"
      Top             =   4800
      Width           =   615
   End
   Begin VB.OptionButton AutoHealOption2 
      BackColor       =   &H00000000&
      Caption         =   "Try to heal friends with any kind of UH "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   5760
      Width           =   4215
   End
   Begin VB.OptionButton AutoHealOption1 
      BackColor       =   &H00000000&
      Caption         =   "Only heal with the UHs found in opened backpacks"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Top             =   5520
      Value           =   -1  'True
      Width           =   4215
   End
   Begin VB.HScrollBar scrollSafeToHealHP 
      Height          =   255
      Left            =   6840
      Max             =   100
      TabIndex        =   25
      Top             =   4440
      Value           =   80
      Width           =   1935
   End
   Begin VB.HScrollBar scrollFriendsHP 
      Height          =   255
      Left            =   6840
      Max             =   100
      TabIndex        =   23
      Top             =   4080
      Value           =   60
      Width           =   1935
   End
   Begin VB.CheckBox chkAutoHealFriendEnabled 
      BackColor       =   &H00000000&
      Caption         =   "Enable the autofriend healer (this mode affects ALL mcs)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   6360
      Width           =   4455
   End
   Begin VB.ListBox lstAutoheal 
      Height          =   1620
      Left            =   5280
      TabIndex        =   19
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdReloadWaroutfits 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reload list"
      Height          =   300
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangeOutfit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Change outfit"
      Height          =   300
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   600
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtOutfit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdMore 
      BackColor       =   &H00C0FFC0&
      Caption         =   "+"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdLess 
      BackColor       =   &H00C0FFC0&
      Caption         =   "-"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.ListBox lstGroups 
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.ListBox lstAllNames 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   3120
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DEACTIVATED - press to run changer"
      Height          =   420
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "to"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   45
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "File name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   40
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Add friend to the autoheal list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   37
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Delay between heal scans: from"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   33
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Healer mode:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Unless I am under:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   28
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Heal friends under:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label label_scrollSafeToHealHP 
      BackColor       =   &H00000000&
      Caption         =   "80 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   26
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label label_scrollFriendsHP 
      BackColor       =   &H00000000&
      Caption         =   "60 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   24
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "List of friends that should get autohealed:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "AUTOFRIEND HEALER:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5160
      X2              =   5160
      Y1              =   120
      Y2              =   6600
   End
   Begin VB.Label lblGlobalEvents 
      BackColor       =   &H00000000&
      Caption         =   "WAR OUTFIT CHANGER:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblStep4 
      BackColor       =   &H00000000&
      Caption         =   "Redefine the outfit for selected group:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Outfit ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Colors (only used in players outfit) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Current list of group files (*.txt) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Following changes will be made:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   2775
   End
End
Attribute VB_Name = "frmWarbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit



Public Function StrToLong(str As String) As Long
  'Function to convert hex string to long
  Dim res As Long
  res = CLng(str)
  StrToLong = res
End Function





Public Sub RenameIt(tibiaclient As Long, newName As String)
  Dim myID As Long
  Dim tmpID As Long
  Dim bPos As Long
  Dim lastPos As Long
  Dim b As Byte
  Dim strB As String
  Dim i As Long
  Dim buffName As String
        myID = Memory_ReadLong(adrNum, tibiaclient)
        
        lastPos = -1
        For bPos = 0 To LAST_BATTLELISTPOS
          tmpID = Memory_ReadLong(adrNChar + (bPos * CharDist), tibiaclient)
          If tmpID = myID Then
            lastPos = bPos

            Exit For
          End If
        Next bPos
        If lastPos = -1 Then
 
        Else
          If Len(newName) <= MAX_NAME_LENGHT Then
          buffName = newName
          i = 0
          While Len(buffName) > 0
            strB = Left(buffName, 1)
            b = CByte(Asc(strB))
            buffName = Right(buffName, Len(buffName) - 1)
            Memory_WriteByte (adrNChar + (lastPos * CharDist) + NameDist + i), b, tibiaclient
            i = i + 1
          Wend
          b = &H0
          Memory_WriteByte (adrNChar + (lastPos * CharDist) + NameDist + i), b, tibiaclient
          End If
        End If
End Sub




Private Sub AutoHealOption1_Click()
  If AutoHealOption1.Value = True Then
    GLOBAL_AUTOFRIENDHEAL_MODE = 1
  End If
End Sub

Private Sub AutoHealOption2_Click()
  If AutoHealOption2.Value = True Then
    GLOBAL_AUTOFRIENDHEAL_MODE = 2
  End If
End Sub

Private Sub AutoHealOption3_Click()
  If AutoHealOption3.Value = True Then
    GLOBAL_AUTOFRIENDHEAL_MODE = 3
  End If
End Sub

Public Sub chkAutoHealFriendEnabled_Click()
If chkAutoHealFriendEnabled.Value = 1 Then
  frmWarbot.timerFriendHealer.enabled = True
Else
  frmWarbot.timerFriendHealer.enabled = False
End If
End Sub



Public Sub cmdChangeOutfit_Click()
  Dim b0 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim b5 As Byte
  Dim filename As String
  b0 = CByte(CLng(txtOutfit(0).Text))
  If ((b0 = 0) And (TibiaVersionLong > 760)) Then
    b0 = firstValidOutfit
  End If
  b1 = CByte(CLng(txtOutfit(1).Text))
  b2 = CByte(CLng(txtOutfit(2).Text))
  b3 = CByte(CLng(txtOutfit(3).Text))
  b4 = CByte(CLng(txtOutfit(4).Text))
  b5 = CByte(CLng(txtOutfit(5).Text))
  If (lstGroups.ListIndex >= 0) Then
    filename = lstGroups.List(lstGroups.ListIndex)
    filename = Left$(filename, Len(filename) - 3) & "out"
    SaveOutfit filename, b0, b1, b2, b3, b4, b5
    ReLoadAllCharOutfits
  End If
End Sub







Private Sub cmdClear_Click()
  lstAutoheal.Clear
End Sub

Private Sub cmdLess_Click()
  Dim curr As Long
  Dim bvalid As Boolean
  curr = CLng(txtOutfit(0).Text)
recheck:
  If TibiaVersionLong >= 773 Then
  curr = curr - 1
  If curr < firstValidOutfit Then
    curr = lastValidOutfit
  End If
  bvalid = True
  Select Case curr
  Case 135
    bvalid = False
  End Select
  If bvalid = False Then
    GoTo recheck
  End If
  txtOutfit(0).Text = CStr(curr)
  lastValid(0) = CStr(curr)
  cmdChangeOutfit_Click
  Else
  curr = curr - 1
  If curr < 0 Then
    curr = 142
  End If
  bvalid = True
  Select Case curr
  Case 1, 10, 11, 12, 20, 46, 47, 72, 77, 93, 96, 97, 98, 135
    bvalid = False
  End Select
  If bvalid = False Then
    GoTo recheck
  End If
  txtOutfit(0).Text = CStr(curr)
  lastValid(0) = CStr(curr)
  cmdChangeOutfit_Click
  End If
End Sub



Private Sub cmdMore_Click()
  Dim curr As Long
  Dim bvalid As Boolean
  curr = CLng(txtOutfit(0).Text)
recheck:
  If TibiaVersionLong >= 773 Then
  curr = curr + 1
  If curr > lastValidOutfit Then
    curr = firstValidOutfit
  End If
  bvalid = True
  Select Case curr
  Case 135
    bvalid = False
  End Select
  If bvalid = False Then
    GoTo recheck
  End If
  txtOutfit(0).Text = CStr(curr)
  lastValid(0) = CStr(curr)
  cmdChangeOutfit_Click
  Else
  curr = curr + 1
  If curr > 142 Then
    curr = 0
  End If
  bvalid = True
  Select Case curr
  Case 1, 10, 11, 12, 20, 46, 47, 72, 77, 93, 96, 97, 98, 135
    bvalid = False
  End Select
  If bvalid = False Then
    GoTo recheck
  End If
  txtOutfit(0).Text = CStr(curr)
  lastValid(0) = CStr(curr)
  cmdChangeOutfit_Click
  End If
End Sub







Private Sub cmdReloadAutoHealFriends_Click()
  ReloadAutohealFile
End Sub

Private Sub cmdReloadWaroutfits_Click()
  allowRename = False
  lastValid(0) = CStr(firstValidOutfit)
  lastValid(1) = 0
  lastValid(2) = 0
  lastValid(3) = 0
  lastValid(4) = 0
  lastValid(5) = 0
  txtOutfit(0) = lastValid(0)
  txtOutfit(1) = lastValid(1)
  txtOutfit(2) = lastValid(2)
  txtOutfit(3) = lastValid(3)
  txtOutfit(4) = lastValid(4)
  txtOutfit(5) = lastValid(5)
  
 OutfitOfName(0).RemoveAll
 OutfitOfName(1).RemoveAll
 OutfitOfName(2).RemoveAll
 OutfitOfName(3).RemoveAll
 OutfitOfName(4).RemoveAll
 OutfitOfName(5).RemoveAll
 OutfitOfChar(0).RemoveAll
 OutfitOfChar(1).RemoveAll
 OutfitOfChar(2).RemoveAll
 OutfitOfChar(3).RemoveAll
  OutfitOfChar(4).RemoveAll
  OutfitOfChar(5).RemoveAll
  lstAllNames.Clear
  lstGroups.Clear
  LoadWarbotFiles
  lastClient = -1
  Me.cmdStart.Caption = "DEACTIVATED - press to run changer"
  Me.Timer1.enabled = False
End Sub



Private Sub cmdSaveAutoheal_Click()
  SaveAutoHealList
End Sub

Private Sub cmdStart_Click()
  If Timer1.enabled = False Then
    cmdStart.Caption = "ACTIVATED - press again and relog to restore outfits"
    Timer1.enabled = True
  Else
    cmdStart.Caption = "DEACTIVATED - press to run changer"
    Timer1.enabled = False
  End If
End Sub








Private Sub Form_Load()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  
  allowRename = False
  If TibiaVersionLong > 760 Then
    firstValidOutfit = &H2
  Else
    firstValidOutfit = &H0
  End If
  lastValid(0) = CStr(firstValidOutfit)
  lastValid(1) = firstValidOutfit
  lastValid(2) = 0
  lastValid(3) = 0
  lastValid(4) = 0
  lastValid(5) = 0
  txtOutfit(0) = lastValid(0)
  txtOutfit(1) = lastValid(1)
  txtOutfit(2) = lastValid(2)
  txtOutfit(3) = lastValid(3)
  txtOutfit(4) = lastValid(4)
  txtOutfit(5) = lastValid(5)
  
  Set OutfitOfName(0) = New scripting.Dictionary
  Set OutfitOfName(1) = New scripting.Dictionary
  Set OutfitOfName(2) = New scripting.Dictionary
  Set OutfitOfName(3) = New scripting.Dictionary
  Set OutfitOfName(4) = New scripting.Dictionary
  Set OutfitOfName(5) = New scripting.Dictionary
  Set OutfitOfChar(0) = New scripting.Dictionary
  Set OutfitOfChar(1) = New scripting.Dictionary
  Set OutfitOfChar(2) = New scripting.Dictionary
  Set OutfitOfChar(3) = New scripting.Dictionary
  Set OutfitOfChar(4) = New scripting.Dictionary
  Set OutfitOfChar(5) = New scripting.Dictionary
  LoadWarbotFiles
  lastClient = -1
  Me.cmdStart.Caption = "DEACTIVATED - press to run changer"
  Me.Timer1.enabled = False
  Exit Sub
goterr:
  gotDictErr = 1
End Sub








Private Sub lstGroups_Click()
  Dim filename As String
  If lstGroups.ListIndex >= 0 Then
    filename = lstGroups.List(lstGroups.ListIndex)
    LoadGroupOutfit filename
  End If
End Sub






Private Sub scrollFriendsHP_Change()
  ChangeGLOBAL_FRIENDSLOWLIMIT_HP scrollFriendsHP.Value
End Sub

Private Sub scrollSafeToHealHP_Change()
  ChangeGLOBAL_MYSAFELIMIT_HP scrollSafeToHealHP.Value
End Sub




Private Sub Timer1_Timer()
  ProcessAllBattleLists
End Sub

Private Sub timerFriendHealer_Timer()
    On Error GoTo goterr
    Dim v1 As Long
    Dim v2 As Long
    v1 = 300
    v2 = 700
    If IsNumeric(txtAutohealDelay.Text) Then
        v1 = CLng(txtAutohealDelay.Text)
    End If
    If IsNumeric(txtAutohealDelay2.Text) Then
        v2 = CLng(txtAutohealDelay2.Text)
    End If
    timerFriendHealer.Interval = randomNumberBetween(v1, v2)
  ProcessAllFriendHeals
  Exit Sub
goterr:
  txtAutohealDelay.Text = "300"
  txtAutohealDelay2.Text = "700"
End Sub





Private Sub txtOutfit_Validate(Index As Integer, Cancel As Boolean)
Dim byteValue As Byte
Dim longValue As Long
On Error GoTo goterr
  If TibiaVersionLong >= 773 Then

  longValue = CLng(txtOutfit(Index).Text)
  If Index = 0 Then
    If longValue > lastValidOutfit Then
      txtOutfit(Index).Text = firstValidOutfit
      Cancel = True
      Exit Sub
    End If
    If longValue < firstValidOutfit Then
      txtOutfit(Index).Text = lastValidOutfit
      Cancel = True
      Exit Sub
    End If
    Select Case longValue
    Case 135
      txtOutfit(Index).Text = lastValid(Index)
      Cancel = True
      Exit Sub
    End Select
  End If
  If longValue >= 0 And longValue <= 255 Then
    byteValue = CByte(longValue)
    lastValid(Index) = CStr(longValue)
  Else
    txtOutfit(Index).Text = lastValid(Index)
    Cancel = True
  End If
  Exit Sub
  Else
  longValue = CLng(txtOutfit(Index).Text)
  If Index = 0 Then
    If longValue < firstValidOutfit Then
      txtOutfit(Index).Text = lastValidOutfit
      Cancel = True
      Exit Sub
    End If
    If longValue > lastValidOutfit Then
      txtOutfit(Index).Text = firstValidOutfit
      Cancel = True
      Exit Sub
    End If
    Select Case longValue
    Case 1, 10, 11, 12, 20, 46, 47, 72, 77, 93, 96, 97, 98, 135
      txtOutfit(Index).Text = lastValid(Index)
      Cancel = True
      Exit Sub
    End Select
  End If
  If longValue >= 0 And longValue <= 255 Then
    byteValue = CByte(longValue)
    lastValid(Index) = CStr(longValue)
  Else
    txtOutfit(Index).Text = lastValid(Index)
    Cancel = True
  End If
  Exit Sub
  
  
  End If
goterr:
  txtOutfit(Index).Text = lastValid(Index)
  Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub ReloadAutohealFile()
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim filename As String
  Set fso = New scripting.FileSystemObject
    lstAutoheal.Clear
    filename = App.path & "\" & txtFileName.Text
    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strLine
        If strLine <> "" Then
        If IsAutoHealFriend(LCase(strLine)) = False Then
          lstAutoheal.AddItem LCase(strLine)
        End If
        End If
      Wend
      Close #fn
    End If
  Exit Sub
goterr:
  lstAutoheal.Clear
End Sub






Public Sub SaveAutoHealList()
  Dim i As Long
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
  Dim fn As Integer
  Dim limI As Long
  limI = lstAutoheal.ListCount - 1
  fn = FreeFile
  Open App.path & "\" & txtFileName.Text For Output As #fn
    For i = 0 To limI
      Print #fn, lstAutoheal.List(i)
    Next i
  Close #fn
  Exit Sub
goterr:
  i = -1
End Sub

Public Function IsAutoHealFriend(strName As String) As Boolean
  'strname should come in lcase
  Dim i As Long
  Dim totI As Long
  Dim foundI As Long
  totI = lstAutoheal.ListCount - 1
  foundI = -1
  For i = 0 To totI
    If lstAutoheal.List(i) = strName Then
      foundI = i
    End If
  Next i
  If foundI = -1 Then
    IsAutoHealFriend = False
  Else
    IsAutoHealFriend = True
  End If
End Function

Private Sub cmdAddFriend_Click()
  If IsAutoHealFriend(LCase(txtAddFriend.Text)) = False Then
    lstAutoheal.AddItem LCase(txtAddFriend.Text)
  End If
End Sub

Private Sub cmdRemoveFriend_Click()
  If lstAutoheal.ListIndex > -1 Then
    lstAutoheal.RemoveItem (lstAutoheal.ListIndex)
  End If
End Sub
