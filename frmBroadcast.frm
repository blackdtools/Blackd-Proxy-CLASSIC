VERSION 5.00
Begin VB.Form frmBroadcast 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Broadcast by private messages"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7785
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBroadcast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMC 
      BackColor       =   &H00000000&
      Caption         =   "(optional) Multiple sources - multiclient broadcast by turns in same server. This allow smaller delay"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   6720
      Width           =   7575
   End
   Begin VB.Timer timerBroadcast 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6840
      Top             =   5640
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   3360
      TabIndex        =   22
      Text            =   "-"
      ToolTipText     =   "Select one of your connected characters"
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STOP"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtBroadcastDelay2 
      Height          =   285
      Left            =   5760
      MaxLength       =   7
      TabIndex        =   12
      Text            =   "30000"
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtBroadcastDelay1 
      Height          =   285
      Left            =   4080
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "20000"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PLAY"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtBroadcastThis 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   7575
   End
   Begin VB.ListBox lstList 
      Height          =   2010
      Left            =   5040
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "> read >"
      Height          =   1455
      Left            =   4200
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtRaw 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.tibia.com/community/?subtopic=worlds&world=Aldora"
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00000000&
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
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   7455
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "0%"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Step 6. Select who will send the broadcast:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label lblCurrentBroadcast 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-nobody"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Currently broadcasting to:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Example 2: hello, I don't care about your name. This is a test"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   6975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Example 1: hello my friend $broadcast$ , I am broadcasting this by private messages"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   6975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Step 7: Press play to start. Press stop if you want to stop it"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   6375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "ms"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "ms and"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Step 5: Set delay between messages. Delay between..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   $"frmBroadcast.frx":0442
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Step 3: press the button below:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Step 2: copy/paste the list of online players here. Don't worry about level and class. Paste all."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   "Step 1: (for real servers) copy/paste this in your browser. Change server at the end:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmBroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private currentBroadcastID As Long

Public Function getCurrentBID() As Long
    getCurrentBID = currentBroadcastID
End Function
Public Function nextValidBID() As Long
    Dim nothingElse As Long
    Dim testIt As Long
    Dim i As Long
    nothingElse = broadcastIDselected
    testIt = nothingElse + 1
    For i = testIt To MAXCLIENTS
        If GameConnected(i) = True Then
            nextValidBID = i
            Exit Function
        End If
    Next i
    For i = 1 To (nothingElse - 1)
        If GameConnected(i) = True Then
            nextValidBID = i
            Exit Function
        End If
    Next i
    nextValidBID = nothingElse
End Function

Private Function filterIt(ByVal strName As String) As String
    Dim lenOfStrname As String
    Dim i As Long
    Dim p As String
    Dim filteredN As String
    filteredN = ""
    lenOfStrname = Len(strName)
    i = 1
    Do
        p = Mid$(strName, i, 1)
        i = i + 1
        If IsNumeric(p) = False Then
            filteredN = filteredN & p
        End If
    Loop Until (IsNumeric(p) = True) Or (i > lenOfStrname)

    filteredN = Trim$(filteredN)
    filteredN = Replace(filteredN, Chr(9), "") ' for firefox explorer
    filterIt = filteredN
End Function
Private Function ReadRaw()
    Dim pos1 As Long
    Dim pos2 As Long
    Dim part As String
    Dim nothingElse As Boolean
    Dim posI As Long
    Dim lenText As Long
    Dim strRaw As String
    Dim a As String
    Dim strName As String
    nothingElse = False
    strRaw = txtRaw.Text & vbCrLf
    lenText = Len(strRaw)
    lstList.Clear
    pos1 = 1
    Do
      pos2 = InStr(pos1, strRaw, vbCr)
      If pos2 > 0 Then
        part = Mid$(strRaw, pos1, pos2 - pos1)
        pos1 = pos2 + 2
        strName = filterIt(part)
        If strName <> "" Then
          lstList.AddItem strName
        End If
      End If
    Loop Until pos2 <= 0
End Function



Private Sub cmbCharacter_Click()
  broadcastIDselected = cmbCharacter.ListIndex
End Sub

Private Sub UpdateTimer()
    Dim lng1 As Long
    Dim lng2 As Long
    If (IsNumeric(txtBroadcastDelay1.Text) = True) And _
     (IsNumeric(txtBroadcastDelay2.Text)) Then
        lng1 = CLng(txtBroadcastDelay1.Text)
        lng2 = CLng(txtBroadcastDelay2.Text)
    Else
        txtBroadcastDelay1.Text = "20000"
        txtBroadcastDelay2.Text = "30000"
        lng1 = 300
        lng2 = 2000
    End If
    If (lng1 < 1) Or (lng2 < 1) Or (lng1 > lng2) Then
        txtBroadcastDelay1.Text = "20000"
        txtBroadcastDelay2.Text = "30000"
        lng1 = 300
        lng2 = 2000
    End If
    timerBroadcast.Interval = randomNumberBetween(lng1, lng2)
End Sub
Private Sub cmdPlay_Click()
    If lstList.ListCount = 0 Then
        lblInfo.Caption = "Destination list is empty. Nothing to broadcast"
    ElseIf txtBroadcastThis.Text = "" Then
        lblInfo.Caption = "First type a text for the broadcast message!"
    ElseIf broadcastIDselected = 0 Then
        lblInfo.Caption = "First select who will send the broadcast messages!"
    ElseIf GameConnected(broadcastIDselected) = False Then
        lblInfo.Caption = "Broadcast cancelled. Selected tibia client is offline"
    Else
        currentBroadcastID = broadcastIDselected - 1
        timerBroadcast.Interval = 300
        timerBroadcast.enabled = True
        lblInfo.Caption = "Broadcast started. Number of players: " & CStr(lstList.ListCount)
    End If
End Sub

Private Sub cmdRead_Click()
ReadRaw
End Sub

Private Sub cmdStop_Click()
    currentBroadcastIndex = -1
    Me.timerBroadcast.enabled = False
    Me.lblPer = "0%"
    Me.lblCurrentBroadcast = "-nobody"
    lblInfo.Caption = "Stopped by user request"
End Sub

Private Sub Form_Load()

  LoadBroadcastChars
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub LoadBroadcastChars()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
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
  broadcastIDselected = firstC
  Exit Sub
goterr:
 LogOnFile "errors.txt", "Error at LoadBroadcastChars(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub


Private Sub timerBroadcast_Timer()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
    Dim lngv As Long
    Dim dblPer As Double
    Dim lstC As Long
    Dim strD As String
    Dim privm As String
    Dim aRes As Long
    Dim currentBID As Long
    If broadcastIDselected = 0 Then
        timerBroadcast.enabled = False
        Exit Sub
    End If
    If chkMC.Value = 1 Then
        currentBID = nextValidBID()
        broadcastIDselected = currentBID
        Me.cmbCharacter.ListIndex = broadcastIDselected
    Else
        currentBID = broadcastIDselected
    End If
    If GameConnected(currentBID) = False Then
        timerBroadcast.enabled = False
        Exit Sub
    End If
    If CheatsPaused(currentBID) = True Then
        Exit Sub
    End If
    lngv = currentBroadcastIndex
    lngv = lngv + 1
    lstC = Me.lstList.ListCount
    If lngv >= lstC Then
        timerBroadcast.enabled = False
        currentBroadcastIndex = -1
        Me.lblCurrentBroadcast = "-nobody"
        Me.lblPer = "0%"
        lblInfo.Caption = "Broadcast finished. Reached end of list"
    Else
        dblPer = Round((CDbl(lngv + 1) / CDbl(lstC)) * 100, 2)
        Me.lblPer = CStr(dblPer) & "%"
        currentBroadcastIndex = lngv
        strD = Me.lstList.List(lngv)
        Me.lblCurrentBroadcast = strD
        privm = "*" & strD & "* " & parseVars(CInt(currentBID), Trim$(txtBroadcastThis.Text))
        aRes = ExecuteInTibia(privm, CInt(currentBID), True)
        UpdateTimer
        DoEvents
    End If
    Exit Sub
goterr:
        timerBroadcast.enabled = False
        currentBroadcastIndex = -1
        Me.lblCurrentBroadcast = "-nobody"
        Me.lblPer = "0%"
        lblInfo.Caption = "Broadcast finished. Unexpected error " & CStr(Err.Number)
End Sub
