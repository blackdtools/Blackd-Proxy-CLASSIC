VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEvents 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridEvents 
      Height          =   1815
      Left            =   120
      TabIndex        =   53
      Top             =   4080
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.TextBox txtTelephoneNumber 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "+34012345678"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CheckBox chkReconnectionAlarm 
      BackColor       =   &H00000000&
      Caption         =   "Activate big alarm on autoreconection"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Top             =   840
      Width           =   3255
   End
   Begin VB.Timer timerScheduledActions 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9840
      Top             =   2280
   End
   Begin VB.TextBox txtDelay 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   4320
      TabIndex        =   50
      Text            =   "0"
      Top             =   8040
      Width           =   735
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H000000FF&
      Caption         =   "Keep this event working even if cheats get paused"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   5400
      TabIndex        =   49
      Top             =   8040
      Width           =   4215
   End
   Begin VB.TextBox txtSkypeHelp 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmEvents.frx":0442
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox txtTheVar 
      Height          =   285
      Left            =   6840
      TabIndex        =   47
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtTrans 
      Height          =   855
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   45
      Text            =   "frmEvents.frx":05E6
      Top             =   2640
      Width           =   4575
   End
   Begin VB.CommandButton cmdUpdateVar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UPDATE"
      Height          =   315
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Loads from given file"
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox lstVar 
      Height          =   840
      Left            =   6120
      TabIndex        =   43
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdDelAll 
      BackColor       =   &H0000C000&
      Caption         =   "Delete all"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Delete all"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdReloadFiles 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- Reload"
      Height          =   315
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Reload"
      Top             =   6000
      Width           =   855
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Monster2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   6000
      TabIndex        =   38
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Monster1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   4800
      TabIndex        =   37
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "RAID MSG"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   3360
      TabIndex        =   36
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Unknown0E"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   35
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Unknown0D"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Tutor to channel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   8280
      TabIndex        =   33
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "rare GM msg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   6840
      TabIndex        =   32
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "GM Priv msg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   5160
      TabIndex        =   31
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "GM to channel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   30
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Unknown08"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   29
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Counsellor Priv msg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Counsellor to channel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   27
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Channel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   26
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Priv msg"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   25
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Yell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   24
      Top             =   6960
      Width           =   855
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Whisper"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   23
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "Say"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   22
      Top             =   6960
      Width           =   975
   End
   Begin VB.CheckBox chkParameter 
      BackColor       =   &H00000000&
      Caption         =   "SYSTEM"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   21
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtAction 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   8400
      Width           =   7215
   End
   Begin VB.TextBox txtTrigger 
      Height          =   285
      Left            =   6720
      TabIndex        =   17
      Top             =   6600
      Width           =   3975
   End
   Begin VB.ComboBox cmbEventType 
      Height          =   315
      Left            =   3000
      TabIndex        =   13
      Text            =   "0 - MSG THAT CONTAINS..."
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddEvent 
      BackColor       =   &H0000C000&
      Caption         =   "Add event"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "ADD"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeleteSel 
      BackColor       =   &H0000C000&
      Caption         =   "Delete selected events"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Delete selected events"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoadEv 
      BackColor       =   &H0000C000&
      Caption         =   "Load events"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Loads from given file"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveEv 
      BackColor       =   &H0000C000&
      Caption         =   "Save events"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Saves to given file"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Text            =   "ev_example.txt"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Text            =   "-"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CheckBox chkTelephoneAlarm 
      BackColor       =   &H00000000&
      Caption         =   "Call to this telephone (using Skype) when big alarm start to sound:"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Delay action this time (ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   51
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Test:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   48
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Value:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   46
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "LIST OF SOME AVAILABLE VARIABLES: Select one and click update button to test current value in the selected char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6120
      TabIndex        =   42
      Top             =   0
      Width           =   4575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6000
      X2              =   6000
      Y1              =   960
      Y2              =   3480
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
      Height          =   495
      Left            =   6000
      TabIndex        =   41
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label lblFlags 
      BackColor       =   &H00000000&
      Caption         =   "Activated from:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblAction 
      BackColor       =   &H00000000&
      Caption         =   "Action :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Trigger text :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "NEW CUSTOM EVENT:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   "Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   6600
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   10680
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblFile 
      BackColor       =   &H00000000&
      Caption         =   "File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "CUSTOM EVENTS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblGlobalEvents 
      BackColor       =   &H00000000&
      Caption         =   "GLOBAL EVENTS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Public Sub UpdateValues()
  ' display event list
  ' CustomEvents(eventsIDselected)
  Dim numofEv As Long
  Dim i As Long
  If eventsIDselected <= 0 Then
    gridEvents.Rows = 1
  Else
    numofEv = CustomEvents(eventsIDselected).Number
    gridEvents.Rows = numofEv + 1
    For i = 1 To numofEv
      With gridEvents
      .TextMatrix(i, 0) = CStr(CustomEvents(eventsIDselected).ev(i).id)
      .TextMatrix(i, 1) = CustomEvents(eventsIDselected).ev(i).flags
      .TextMatrix(i, 2) = CustomEvents(eventsIDselected).ev(i).trigger
      .TextMatrix(i, 3) = CustomEvents(eventsIDselected).ev(i).action
      .Row = i
      .Col = 0
      .CellAlignment = flexAlignCenterCenter
      .Col = 1
      .CellAlignment = flexAlignLeftCenter
      .Col = 2
      .CellAlignment = flexAlignLeftCenter
      .Col = 3
      .CellAlignment = flexAlignLeftCenter
      End With
    Next i
  End If
End Sub

Public Sub LoadEventChars()
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
      cmbCharacter.AddItem "-", i
    End If
  Next i
  cmbCharacter.ListIndex = firstC
  cmbCharacter.Text = cmbCharacter.List(firstC)
  eventsIDselected = firstC
  UpdateValues
End Sub


Private Sub chkDisableEventsAtPause_Click()
'...chkDisableEventsAtPause
End Sub

Private Sub cmbCharacter_Click()
  eventsIDselected = cmbCharacter.ListIndex
  UpdateValues
End Sub
Public Sub DeleteAllEvents(idEv As Long)
  If (idEv <= 0) Then
    Exit Sub
  Else
    CustomEvents(idEv).Number = 0
  End If
End Sub



Public Function AddEvent(idConnection As Integer, id As Integer, fl As String, _
 tr As String, ac As String) As Long
  Dim curr As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (idConnection <= 0) Then
    AddEvent = -1
    Exit Function
  End If
  If (CustomEvents(idConnection).Number = MAXEVENTS) Then
    AddEvent = -1
    Exit Function
  End If
  curr = (CustomEvents(idConnection).Number) + 1
  CustomEvents(idConnection).Number = curr
  CustomEvents(idConnection).ev(curr).id = id
  CustomEvents(idConnection).ev(curr).flags = fl
  CustomEvents(idConnection).ev(curr).trigger = tr
  CustomEvents(idConnection).ev(curr).action = ac
  AddEvent = 0
  Exit Function
goterr:
  AddEvent = -1
End Function


Private Sub cmdAddEvent_Click()
  Dim aRes As Long
  Dim theFlags As String
  Dim i As Long
  Dim id As Integer
  If eventsIDselected > 0 Then
    theFlags = ""
    For i = 0 To 18
      theFlags = theFlags & CStr(CLng(chkParameter(i).Value))
    Next i
    If cmbEventType.ListIndex < 0 Then
      id = 0
    Else
      id = CInt(cmbEventType.ListIndex)
    End If
    theFlags = theFlags & ":" & CStr(txtDelay.Text)
    If Left$(theFlags, 18) = "000000000000000000" Then
      lblInfo.Caption = "ERROR: No sources selected!"
    ElseIf txtAction.Text = "" Then
      lblInfo.Caption = "ERROR: No action selected!"
    Else
    aRes = AddEvent(CInt(eventsIDselected), id, theFlags, _
     txtTrigger.Text, txtAction.Text)
    UpdateValues
    lblInfo.Caption = "New event added OK"
    End If
  Else
    lblInfo.Caption = "ERROR: no char selected!"
  End If
End Sub

Private Sub cmdDelAll_Click()
  DeleteAllEvents eventsIDselected
  UpdateValues
End Sub

Private Sub cmdDeleteSel_Click()
  Dim vrow As Long
  Dim vrowsel As Long
  Dim firstrow As Long
  Dim lastrow As Long
  Dim firstI As Long
  Dim lasti As Long
  Dim i As Long
  Dim difR As Long
  Dim numofEv As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (eventsIDselected > 0) Then
  vrow = gridEvents.Row
  vrowsel = gridEvents.RowSel
  If vrow > vrowsel Then
    firstrow = vrowsel
    lastrow = vrow
  Else
    firstrow = vrow
    lastrow = vrowsel
  End If
  numofEv = CustomEvents(eventsIDselected).Number
  If lastrow > numofEv Then
    lastrow = numofEv
  End If
  If (firstrow > lastrow) Or (numofEv = 0) Then
   'lblDebug.Caption = "Invalid selection"
  Else
  ' lblDebug.Caption = "First = " & firstRow & " ; Last = " & lastRow
   firstI = firstrow
   lasti = lastrow
   difR = lasti - firstI + 1
   For i = firstI To numofEv
     If i + difR <= MAXEVENTS Then
       CustomEvents(eventsIDselected).ev(i).id = CustomEvents(eventsIDselected).ev(i + difR).id
       CustomEvents(eventsIDselected).ev(i).flags = CustomEvents(eventsIDselected).ev(i + difR).flags
       CustomEvents(eventsIDselected).ev(i).trigger = CustomEvents(eventsIDselected).ev(i + difR).trigger
       CustomEvents(eventsIDselected).ev(i).action = CustomEvents(eventsIDselected).ev(i + difR).action
      End If
    Next i
    CustomEvents(eventsIDselected).Number = numofEv - difR
  End If
  UpdateValues
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Function cmdDeleteSel_Click() failed : " & Err.Description
  LogOnFile "errors.txt", "Function cmdDeleteSel_Click() failed : " & Err.Description
End Sub

Private Sub cmdLoadEv_Click()
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine(1 To 4) As String
  Dim filename As String
  Dim i As Long
  Dim p As Long
  Dim completed As Boolean
  Dim aRes As Long
  Dim thelo As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Set fso = New scripting.FileSystemObject
  If eventsIDselected > 0 Then
    lblInfo.Caption = "Load OK"
    DeleteAllEvents eventsIDselected
    filename = App.path & "\events\" & txtFile.Text
    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      i = 0
      While Not EOF(fn)
        completed = True
        For p = 1 To 4
        If Not EOF(fn) Then
          Line Input #fn, strLine(p)
        Else
          strLine(p) = ""
          completed = False
        End If
        strLine(p) = LTrim$(strLine(p))
        Next p
        If completed = True Then
          thelo = Len(strLine(2))
          If thelo = 18 Then
            strLine(2) = strLine(2) & "0"
          End If
          If thelo < 20 Then
            strLine(2) = strLine(2) & ":0"
          End If
          aRes = AddEvent(CInt(eventsIDselected), CInt(strLine(1)), strLine(2), strLine(3), strLine(4))
        End If
      Wend
      Close #fn
    Else
       lblInfo.Caption = "FAILED TO LOAD"
    End If
  End If
  UpdateValues
  Exit Sub
goterr:
  lblInfo.Caption = "Load ERROR (" & Err.Number & "):" & Err.Description
End Sub

Private Sub cmdReloadFiles_Click()
  ReloadFiles
End Sub



Private Sub cmdSaveEv_Click()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim fn As Integer
  Dim i As Long
  If eventsIDselected > 0 Then
    fn = FreeFile
    Open App.path & "\events\" & txtFile.Text For Output As #fn
    For i = 1 To CustomEvents(eventsIDselected).Number
      Print #fn, CStr(CustomEvents(eventsIDselected).ev(i).id)
      Print #fn, CustomEvents(eventsIDselected).ev(i).flags
      Print #fn, CustomEvents(eventsIDselected).ev(i).trigger
      Print #fn, CustomEvents(eventsIDselected).ev(i).action
    Next i
    Close #fn
    lblInfo.Caption = "Save OK"
  End If
  Exit Sub
goterr:
  lblInfo.Caption = "Save ERROR (" & Err.Number & "):" & Err.Description
End Sub

Private Sub cmdUpdateVar_Click()
  If (eventsIDselected > 0) Then
    If GameConnected(eventsIDselected) = True Then
      var_lf(eventsIDselected) = vbLf
      txtTrans.Text = parseVars(CInt(eventsIDselected), txtTheVar.Text)
    Else
      txtTrans.Text = "Error: selected char is disconnected"
    End If
  Else
    txtTrans.Text = "Error: select a char first"
  End If
End Sub



Private Sub Form_Load()
cmbEventType.Clear
cmbEventType.AddItem "0 - MSG THAT CONTAINS..."
cmbEventType.AddItem "1 - EXACT MESSAGE"
cmbEventType.AddItem "2 - Like Regex.."
cmbEventType.Text = "0 - MSG THAT CONTAINS..."
ReloadFiles
  With lstVar
    .Clear
    .AddItem "$$"
    .AddItem "$expleft$"
    .AddItem "$nextlevel$"
    .AddItem "$exph$"
    .AddItem "$timeleft$"
    .AddItem "$played$"
    .AddItem "$expgained$"
    .AddItem "$charactername$"
    .AddItem "$lastsender$"
    .AddItem "$lastmsg$"
    .AddItem "$lf$"
    .AddItem "$myhp$"
    .AddItem "$myhppercent$"
    .AddItem "$mymanapercent$"
    .AddItem "$mymana$"
    .AddItem "$mylevel$"
    .AddItem "$mysoulpoints$"
    .AddItem "$myexp$"
    .AddItem "$lastpkname$"
    .AddItem "$lastgmname$"
    .AddItem "$date$"
    .AddItem "$time$"
    .AddItem "$shorttime$"
    .AddItem "$mycap$"
    .AddItem "$mystamina$"
    .AddItem "$randomlineof:hi.txt$"
    .AddItem "$nlineoflabel:labelname$"
    .AddItem "$hex-equiped-item:01$"
    .AddItem "$hex-equiped-ammount:01$"
    .AddItem "$num-equiped-ammount:01$"
    .AddItem "$hex-equiped-ammount-special:01$"
    .AddItem "$urlencode:are you there?$"
    .AddItem "$hex-tibiastr:hello$"
    .AddItem "$hex-lastattackedid$"
    .AddItem "$nameofhex-id:AB CD EF 00$"
    .AddItem "$hex-idofname:name$"
    .AddItem "$httpget:http://example.org/what.php?foo=bar&baz=blabla$"
    .AddItem "$numbertohex1:171$"
    .AddItem "$numbertohex2:52651$"
    .AddItem "$hex1tonumber:AB$"
    .AddItem "$hex2tonumber:AB CD$"
    .AddItem "$numericalexp:2+3$"
    .AddItem "$numericalexp:7-4$"
    .AddItem "$numericalexp:3*5$"
    .AddItem "$numericalexp:8/2$"
    .AddItem "$myx$"
    .AddItem "$myy$"
    .AddItem "$myz$"
    .AddItem "$numericalexp:{$mylevel$}+1$"
    .AddItem "$comboorder$"
    .AddItem "$comboleader$"
    .AddItem "$lastusedchannelid$"
    .AddItem "$lastrecchannelid$"
    .AddItem "$pksonrelativefloor:0$"
    .AddItem "$gmsonrelativefloor:0$"
    .AddItem "$pksandgmsonrelativefloor:0$"
    .AddItem "$statusbit:8$"
    .AddItem "$_customvariable$"
    .AddItem "$__customglobalvariable$"
    .AddItem "$countitems:D7 0B$"
    .AddItem "$countitems:D7 0B 64$"
    .AddItem "$hpofhex-id:AB CD EF 00$"
    .AddItem "$dirofhex-id:AB CD EF 00$"
    .AddItem "$lasthpchange$"
    .AddItem "$cavebottimewithsametarget$"
    .AddItem "$lastattackedid$"
    .AddItem "$bestenemy$"
    .AddItem "$bestenemyhp$"
    .AddItem "$bestenemyid$"
    .AddItem "$hex-bestenemyid$"
    .AddItem "$randomnumber:1>6$"
    .AddItem "$broadcast$"
    .AddItem "$lastcheckresult$"
    .AddItem "$lastchecktileid$"
    .AddItem "$useitemwithamount:XX XX,A$"
    .AddItem "$hex-currenttargetid$"
    .AddItem "$check:{$a$},#conditionOperator#,{$b$}$"
    .AddItem "$istrue:{$var1$},{$var2$}, ... , {$varN$}$"
  End With
  With gridEvents
  .ColWidth(0) = 800
  .ColWidth(1) = 2500
  .ColWidth(2) = 2000
  .ColWidth(3) = 3700
  .TextMatrix(0, 0) = "Type"
  .TextMatrix(0, 1) = "Flags"
  .TextMatrix(0, 2) = "Trigger"
  .TextMatrix(0, 3) = "Action (Command / message to say)"
  .Row = 0
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignLeftCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  .Col = 3
  .CellAlignment = flexAlignLeftCenter
  End With
End Sub

Public Sub ReloadFiles()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim strPath As String
  Dim fs As scripting.FileSystemObject
  Dim f As scripting.Folder
  Dim f1 As scripting.File
  Set fs = New scripting.FileSystemObject
  strPath = App.path & "\events"
  Set f = fs.GetFolder(strPath)
  txtFile.Clear
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "txt" Then
        txtFile.AddItem f1.name
    End If
  Next
  txtFile.Text = "ev_example.txt"
  Exit Sub
goterr:
  LogOnFile "errors.txt", "ERROR WITH FILESYSTEM OBJECT at ReloadFiles (" & Err.Number & ") : " & Err.Description & " (path : " & strPath & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Private Sub lstVar_Click()
  txtTheVar = lstVar.Text
End Sub

Private Sub timerScheduledActions_Timer()
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim i As Long
  Dim ci As Long
  Dim iRes As Integer
  Dim gtc As Long
  gtc = GetTickCount()
  For i = 1 To MAXSCHEDULED
    If scheduledActions(i).pending = True Then
      If gtc > scheduledActions(i).tickc Then
        ci = scheduledActions(i).clientID
        If ((GameConnected(ci) = True) And (sentWelcome(ci) = True) And (GotPacketWarning(ci) = False)) Then
          iRes = ExecuteInTibia(scheduledActions(i).action, scheduledActions(i).clientID, False)
        End If
        scheduledActions(i).pending = False
      End If
    End If
  Next i
  Exit Sub
goterr:
  iRes = -1
End Sub


