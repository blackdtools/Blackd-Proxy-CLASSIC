VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCondEvents 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conditional events"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCondEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridEvents 
      Height          =   1695
      Left            =   0
      TabIndex        =   32
      Top             =   2160
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.TextBox txtLock 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   9720
      TabIndex        =   30
      Text            =   "0"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox chkKeep 
      BackColor       =   &H000000FF&
      Caption         =   "Keep this event working even if cheats get paused"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Timer timerCheck 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   6840
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H0000C000&
      Caption         =   "Modify selected condition"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "MODIFY"
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "GLOBAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   23
      Top             =   1080
      Width           =   6135
      Begin VB.TextBox txtMs2 
         Height          =   285
         Left            =   4080
         TabIndex        =   33
         Text            =   "700"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtMs 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Text            =   "300"
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   3960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "ms"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "* TIMER TICK: All conditions each ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox txtDelay 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtThing1 
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Text            =   "$mymana$"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   4800
      Width           =   3375
   End
   Begin VB.ComboBox cmbOperator 
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Text            =   "#number<=#"
      ToolTipText     =   "Operator"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtThing2 
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Text            =   "100"
      ToolTipText     =   "number, text or $var$ <- read list in events module"
      Top             =   4800
      Width           =   4335
   End
   Begin VB.CommandButton cmdAddEvent 
      BackColor       =   &H0000C000&
      Caption         =   "Add as new conditional event"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "ADD"
      Top             =   6840
      Width           =   2895
   End
   Begin VB.TextBox txtAction 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   6360
      Width           =   9375
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Text            =   "c_example.txt"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveEv 
      BackColor       =   &H0000C000&
      Caption         =   "Save conds"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Saves to given file"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadEv 
      BackColor       =   &H0000C000&
      Caption         =   "Load conds"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Loads from given file"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteSel 
      BackColor       =   &H0000C000&
      Caption         =   "Delete selected conds"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Delete selected conditional events"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdReloadFiles 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<- Reload"
      Height          =   315
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Reload the file list"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdDelAll 
      BackColor       =   &H0000C000&
      Caption         =   "Delete all"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete all the list"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "-"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Do not execute the action again if that action was done ""few time ago"" because the same condition . Define ""few time ago"" (ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   9615
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
      Left            =   120
      TabIndex        =   29
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   10680
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Condition :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "List of conditional events - You can click on a conditional event to display it / modify it :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Delay action this time (ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "thing1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "thing2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "operator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblAction 
      BackColor       =   &H00000000&
      Caption         =   "Action :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lblFile 
      BackColor       =   &H00000000&
      Caption         =   "File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
      Caption         =   $"frmCondEvents.frx":0442
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Char:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmCondEvents"
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
  If condEventsIDselected <= 0 Then
    gridEvents.Rows = 1
  Else
    numofEv = CustomCondEvents(condEventsIDselected).Number
    gridEvents.Rows = numofEv + 1
    For i = 1 To numofEv
      With gridEvents
      .TextMatrix(i, 0) = CustomCondEvents(condEventsIDselected).ev(i).thing1
      .TextMatrix(i, 1) = CustomCondEvents(condEventsIDselected).ev(i).operator
      .TextMatrix(i, 2) = CustomCondEvents(condEventsIDselected).ev(i).thing2
      .TextMatrix(i, 3) = CustomCondEvents(condEventsIDselected).ev(i).delay
      .TextMatrix(i, 4) = CustomCondEvents(condEventsIDselected).ev(i).lock
      .TextMatrix(i, 5) = CustomCondEvents(condEventsIDselected).ev(i).keep
      .TextMatrix(i, 6) = CustomCondEvents(condEventsIDselected).ev(i).action
      .Row = i
      .Col = 0
      .CellAlignment = flexAlignCenterCenter
      .Col = 1
      .CellAlignment = flexAlignCenterCenter
      .Col = 2
      .CellAlignment = flexAlignCenterCenter
      .Col = 3
      .CellAlignment = flexAlignCenterCenter
      .Col = 4
      .CellAlignment = flexAlignCenterCenter
      .Col = 5
      .CellAlignment = flexAlignCenterCenter
      .Col = 6
      .CellAlignment = flexAlignLeftCenter
      End With
    Next i
  End If
End Sub

Public Sub LoadCondEventChars()
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
  condEventsIDselected = firstC
  UpdateValues
End Sub


Private Sub cmbCharacter_Click()
  condEventsIDselected = cmbCharacter.ListIndex
  UpdateValues
End Sub

Private Sub cmdAddEvent_Click()
  Dim aRes As Long
  aRes = AddCondEvent(CInt(condEventsIDselected), txtThing1.Text, cmbOperator.Text, txtThing2.Text, txtDelay.Text, txtLock.Text, CStr(chkKeep.Value), txtAction.Text)
  If aRes = 0 Then
    UpdateValues
    lblInfo.Caption = "New cond added OK"
  ElseIf aRes = -2 Then
    lblInfo.Caption = "ERROR: no char selected!"
  Else
    lblInfo.Caption = "ERROR: can't add more conds!"
  End If
End Sub

Private Sub cmdDelAll_Click()
  DeleteAllCondEvents condEventsIDselected
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
  If (condEventsIDselected > 0) Then
  vrow = gridEvents.Row
  vrowsel = gridEvents.RowSel
  If vrow > vrowsel Then
    firstrow = vrowsel
    lastrow = vrow
  Else
    firstrow = vrow
    lastrow = vrowsel
  End If
  numofEv = CustomCondEvents(condEventsIDselected).Number
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
     If i + difR <= MAXCONDS Then
       CustomCondEvents(condEventsIDselected).ev(i).action = CustomCondEvents(condEventsIDselected).ev(i + difR).action
       CustomCondEvents(condEventsIDselected).ev(i).delay = CustomCondEvents(condEventsIDselected).ev(i + difR).delay
       CustomCondEvents(condEventsIDselected).ev(i).keep = CustomCondEvents(condEventsIDselected).ev(i + difR).keep
       CustomCondEvents(condEventsIDselected).ev(i).lock = CustomCondEvents(condEventsIDselected).ev(i + difR).lock
       CustomCondEvents(condEventsIDselected).ev(i).operator = CustomCondEvents(condEventsIDselected).ev(i + difR).operator
       CustomCondEvents(condEventsIDselected).ev(i).thing1 = CustomCondEvents(condEventsIDselected).ev(i + difR).thing1
       CustomCondEvents(condEventsIDselected).ev(i).thing2 = CustomCondEvents(condEventsIDselected).ev(i + difR).thing2
       CustomCondEvents(condEventsIDselected).ev(i).nextunlock = CustomCondEvents(condEventsIDselected).ev(i + difR).nextunlock
      End If
    Next i
    CustomCondEvents(condEventsIDselected).Number = numofEv - difR
  End If
  UpdateValues
  End If
  Exit Sub
goterr:
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Function cmdDeleteSel_Click() - cond - failed : " & Err.Description
  LogOnFile "errors.txt", "Function cmdDeleteSel_Click() - cond - failed : " & Err.Description
End Sub

Private Sub cmdLoadEv_Click()
  Dim fso As scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine(1 To 7) As String
  Dim filename As String
  Dim p As Long
  Dim seguir As Boolean
  Dim completed As Boolean
  Dim aRes As Long
  Dim thelo As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Set fso = New scripting.FileSystemObject
  If condEventsIDselected > 0 Then
    lblInfo.Caption = "Load OK"
    DeleteAllCondEvents condEventsIDselected
    filename = App.path & "\conds\" & txtFile.Text
    If fso.FileExists(filename) = True Then
      fn = FreeFile
      Open filename For Input As #fn
      While Not EOF(fn)
        completed = True
        For p = 1 To 7
        If Not EOF(fn) Then
          Line Input #fn, strLine(p)
          strLine(p) = Trim$(strLine(p))
        Else
          strLine(p) = ""
        End If
        Next p
        aRes = AddCondEvent(CInt(condEventsIDselected), strLine(1), strLine(2), strLine(3), strLine(4), strLine(5), strLine(6), strLine(7))
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

Private Sub cmdModify_Click()
  Dim vrow As Long
  If condEventsIDselected > 0 Then
    vrow = gridEvents.Row
    If vrow > 0 Then
      CustomCondEvents(condEventsIDselected).ev(vrow).thing1 = txtThing1.Text
      CustomCondEvents(condEventsIDselected).ev(vrow).operator = cmbOperator.Text
      CustomCondEvents(condEventsIDselected).ev(vrow).thing2 = txtThing2.Text
      If chkKeep.Value = 1 Then
        CustomCondEvents(condEventsIDselected).ev(vrow).keep = "1"
      Else
        CustomCondEvents(condEventsIDselected).ev(vrow).keep = "0"
      End If
      CustomCondEvents(condEventsIDselected).ev(vrow).action = txtAction.Text
      CustomCondEvents(condEventsIDselected).ev(vrow).delay = txtDelay.Text
      CustomCondEvents(condEventsIDselected).ev(vrow).lock = txtLock.Text
      UpdateValues
      gridEvents.Row = vrow
      gridEvents.RowSel = vrow
      gridEvents.Col = 0
      gridEvents.ColSel = 6
      lblInfo.Caption = "Modified OK"
    Else
      lblInfo.Caption = "Nothing selected"
    End If
  Else
    lblInfo.Caption = "Select a character first"
  End If
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
  If condEventsIDselected > 0 Then
    fn = FreeFile
    Open App.path & "\conds\" & txtFile.Text For Output As #fn
    For i = 1 To CustomCondEvents(condEventsIDselected).Number
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).thing1
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).operator
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).thing2
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).delay
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).lock
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).keep
      Print #fn, CustomCondEvents(condEventsIDselected).ev(i).action
    Next i
    Close #fn
    lblInfo.Caption = "Save OK"
  End If
  Exit Sub
goterr:
  lblInfo.Caption = "Save ERROR (" & Err.Number & "):" & Err.Description
End Sub



Private Sub Form_Load()
 On Error GoTo goterr
 With cmbOperator
 .Clear
 .AddItem "#number=#"
 .AddItem "#number<=#"
 .AddItem "#number>=#"
 .AddItem "#number<>#"
 .AddItem "#number<#"
 .AddItem "#number>#"
 .AddItem "#string=#"
 .AddItem "#string<>#"
 .Text = "#number<=#"
 End With
 
  With gridEvents
  .ColWidth(0) = 2000
  .ColWidth(1) = 1000
  .ColWidth(2) = 2000
  .ColWidth(3) = 600
  .ColWidth(4) = 600
  .ColWidth(5) = 600
  .ColWidth(6) = 4400
  .TextMatrix(0, 0) = "Thing 1"
  .TextMatrix(0, 1) = "Operator"
  .TextMatrix(0, 2) = "Thing 2"
  .TextMatrix(0, 3) = "Delay"
  .TextMatrix(0, 4) = "Lock"
  .TextMatrix(0, 5) = "Keep"
  .TextMatrix(0, 6) = "Action (Command / message to say)"
  .Row = 0
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignCenterCenter
  .Col = 3
  .CellAlignment = flexAlignCenterCenter
  .Col = 4
  .CellAlignment = flexAlignCenterCenter
  .Col = 5
  .CellAlignment = flexAlignCenterCenter
  .Col = 6
  .CellAlignment = flexAlignCenterCenter
  End With
  ReloadFiles
 Exit Sub
goterr:
  LogOnFile "errors.txt", "Could not load cavebot module. Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub

Public Sub ProcessClientConditions(idConnection As Integer, condid As Long)
  Dim a As Integer
  Dim part1 As String
  Dim part2 As String
  Dim res As Boolean
  Dim op As String
  Dim mustdelay As Long
  Dim intRes As Integer
  Dim executeThis As String
  Dim gtc As Long
  Dim addlock As Long
  On Error GoTo ignoreit
  a = 0
  part1 = parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).thing1)
  part2 = parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).thing2)
  op = parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).operator)
  Select Case op
    Case "#number=#"
      If safeDouble(part1) = safeDouble(part2) Then
        res = True
      End If
    Case "#number<=#"
      If safeDouble(part1) <= safeDouble(part2) Then
        res = True
      End If
    Case "#number>=#"
      If safeDouble(part1) >= safeDouble(part2) Then
        res = True
      End If
    Case "#number<>#"
      If safeDouble(part1) <> safeDouble(part2) Then
        res = True
      End If
    Case "#number<#"
      If safeDouble(part1) < safeDouble(part2) Then
        res = True
      End If
    Case "#number>#"
      If safeDouble(part1) > safeDouble(part2) Then
        res = True
      End If
    Case "#string=#"
      If part1 = part2 Then
        res = True
      End If
    Case "#string<>#"
      If part1 <> part2 Then
        res = True
      End If
    Case Else
      res = False
  End Select
  If (res = True) Then
    gtc = GetTickCount()
    If (gtc >= CustomCondEvents(idConnection).ev(condid).nextunlock) Then
      mustdelay = safeLong(parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).delay))
      addlock = safeLong(parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).lock))
      If addlock = 0 Then
        CustomCondEvents(idConnection).ev(condid).nextunlock = 0
      Else
        CustomCondEvents(idConnection).ev(condid).nextunlock = gtc + addlock
      End If
      executeThis = parseVars(idConnection, CustomCondEvents(idConnection).ev(condid).action)
      If mustdelay = 0 Then
        intRes = ExecuteInTibia(executeThis, idConnection, True)
      Else
        mustdelay = mustdelay + GetTickCount()
        AddSchedule idConnection, executeThis, mustdelay
      End If
    End If
  End If
  Exit Sub
ignoreit:
  a = -1
End Sub



Private Sub gridEvents_Click()
  Dim vrow As Long
  Dim vrowsel As Long
  Dim firstrow As Long
  Dim lastrow As Long
  If condEventsIDselected > 0 Then
    vrow = gridEvents.Row
    If vrow > 0 Then
      txtThing1.Text = CustomCondEvents(condEventsIDselected).ev(vrow).thing1
      cmbOperator.Text = CustomCondEvents(condEventsIDselected).ev(vrow).operator
      txtThing2.Text = CustomCondEvents(condEventsIDselected).ev(vrow).thing2
      If CustomCondEvents(condEventsIDselected).ev(vrow).keep = "1" Then
        chkKeep.Value = 1
      Else
        chkKeep.Value = 0
      End If
      txtAction.Text = CustomCondEvents(condEventsIDselected).ev(vrow).action
      txtDelay.Text = CustomCondEvents(condEventsIDselected).ev(vrow).delay
      txtLock.Text = CustomCondEvents(condEventsIDselected).ev(vrow).lock
    End If
  End If
End Sub

Private Sub timerCheck_Timer()
  Dim i As Integer
  Dim ev As Long
  timerCheck.Interval = randomNumberBetween(TimerConditionTick, TimerConditionTick2)
  
  For i = 1 To MAXCLIENTS
  
    If ((GameConnected(i) = True) And (sentWelcome(i) = True) And (GotPacketWarning(i) = False)) Then
 
    For ev = 1 To CustomCondEvents(i).Number
      If ((CheatsPaused(i) = False) Or (CustomCondEvents(i).ev(ev).keep = "1")) Then
        ProcessClientConditions i, ev
      End If
    Next ev
    End If
  Next i
End Sub

Private Sub txtMs_Change()
  On Error GoTo goterr
  Dim lngTick As Long
  lngTick = CLng(txtMs.Text)
  If lngTick < 20 Then
    lngTick = 20
  End If
  TimerConditionTick = lngTick
  timerCheck.Interval = lngTick
  Exit Sub
goterr:
  txtMs.Text = "300"
  TimerConditionTick = 300
  timerCheck.Interval = 300
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
  strPath = App.path & "\conds"
  Set f = fs.GetFolder(strPath)
  txtFile.Clear
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "txt" Then
        txtFile.AddItem f1.name
    End If
  Next
  txtFile.Text = "c_example.txt"
  Exit Sub
goterr:
  LogOnFile "errors.txt", "ERROR WITH FILESYSTEM OBJECT at ReloadFiles (" & Err.Number & ") : " & Err.Description & " (at cond events - path : " & strPath & ")"
End Sub

Public Sub DeleteAllCondEvents(idEv As Long)
  If (idEv <= 0) Then
    Exit Sub
  Else
    CustomCondEvents(idEv).Number = 0
  End If
End Sub

Public Function AddCondEvent(idConnection As Integer, t1 As String, op As String, _
 t2 As String, de As String, lo As String, ke As String, ac As String) As Long
  Dim curr As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If (idConnection <= 0) Then
    AddCondEvent = -2
    Exit Function
  End If
  If (CustomCondEvents(idConnection).Number = MAXCONDS) Then
    AddCondEvent = -1
    Exit Function
  End If
  curr = (CustomCondEvents(idConnection).Number) + 1
  CustomCondEvents(idConnection).Number = curr
  CustomCondEvents(idConnection).ev(curr).thing1 = t1
  CustomCondEvents(idConnection).ev(curr).operator = op
  CustomCondEvents(idConnection).ev(curr).thing2 = t2
  CustomCondEvents(idConnection).ev(curr).delay = de
  CustomCondEvents(idConnection).ev(curr).lock = lo
  CustomCondEvents(idConnection).ev(curr).keep = ke
  CustomCondEvents(idConnection).ev(curr).action = ac
  CustomCondEvents(idConnection).ev(curr).nextunlock = 0
  AddCondEvent = 0
  Exit Function
goterr:
  AddCondEvent = -1
End Function

Private Sub txtMs2_Change()
  On Error GoTo goterr
  Dim lngTick2 As Long
  lngTick2 = CLng(txtMs2.Text)
  If lngTick2 < (TimerConditionTick + 300) Then
    lngTick2 = TimerConditionTick + 300
  End If
  TimerConditionTick2 = lngTick2
  Exit Sub
goterr:
  txtMs2.Text = CStr(TimerConditionTick + 300)
  TimerConditionTick2 = TimerConditionTick + 300
End Sub
