VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHotkeys 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotkeys"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmHotkeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridHotkeys 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Text            =   "500"
      Top             =   4400
      Width           =   615
   End
   Begin VB.CheckBox chkRepeat 
      BackColor       =   &H00000000&
      Caption         =   "Repeat action while key pressed, each"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CheckBox chkHotkeysActivated 
      BackColor       =   &H00000000&
      Caption         =   "Activate hotkeys"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Timer timerHotkeys 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   3600
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save changes"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel changes"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteSel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete selected"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete all"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "ms"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   4455
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "OPTIONS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblDebug 
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Current hotkeys:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private reenableHotkeyTime As Long
Private Sub cmdCancel_Click()
  Dim res As Integer
  ReleasePressKey
  res = LoadHotkeys()
  If res = -1 Then
    lblDebug.Caption = "Could not load hotkeys.ini"
    lblDebug.ForeColor = &HFF&
  Else
    lblDebug.Caption = "Loaded hotkeys.ini successfully"
    lblDebug.ForeColor = &HFF00&
  End If
End Sub

Private Sub cmdClear_Click()
  ReleasePressKey
  NumberOfHotkeys = 0
  With gridHotkeys
    .Rows = 2
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = ""
    .TextMatrix(1, 2) = ""
  End With
  ReDim Hotkeys(0)
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
  ReleasePressKey
  vrow = gridHotkeys.Row
  vrowsel = gridHotkeys.RowSel
  If vrow > vrowsel Then
    firstrow = vrowsel
    lastrow = vrow
  Else
    firstrow = vrow
    lastrow = vrowsel
  End If
  If lastrow > NumberOfHotkeys Then
    lastrow = NumberOfHotkeys
  End If
  If (firstrow > lastrow) Or (NumberOfHotkeys = 0) Then
   'lblDebug.Caption = "Invalid selection"
  Else
  ' lblDebug.Caption = "First = " & firstRow & " ; Last = " & lastRow
   firstI = firstrow - 1
   difR = lastrow - firstrow + 1
   lasti = NumberOfHotkeys - difR - 1
    For i = firstI To lasti
      If i + difR <= (NumberOfHotkeys - 1) Then
        Hotkeys(i).key1 = Hotkeys(i + difR).key1
        Hotkeys(i).key2 = Hotkeys(i + difR).key2
        Hotkeys(i).command = Hotkeys(i + difR).command
      End If
    Next i
    NumberOfHotkeys = NumberOfHotkeys - difR
    If NumberOfHotkeys = 0 Then
      With gridHotkeys
       .Rows = 2
       .TextMatrix(1, 0) = ""
       .TextMatrix(1, 1) = ""
       .TextMatrix(1, 2) = ""
      End With
      ReDim Hotkeys(0)
    Else
      ReDim Preserve Hotkeys(NumberOfHotkeys - 1)
      With gridHotkeys
       .Rows = NumberOfHotkeys + 2
       .TextMatrix(NumberOfHotkeys + 1, 0) = ""
       .TextMatrix(NumberOfHotkeys + 1, 1) = ""
       .TextMatrix(NumberOfHotkeys + 1, 2) = ""
      End With
      For i = firstrow To NumberOfHotkeys
        gridHotkeys.TextMatrix(i, 0) = TranslateHotkeyID(Hotkeys(i - 1).key1)
        gridHotkeys.TextMatrix(i, 1) = TranslateHotkeyID(Hotkeys(i - 1).key2)
        gridHotkeys.TextMatrix(i, 2) = Hotkeys(i - 1).command
      Next i
    End If
  End If
End Sub
Public Function SaveHotkeys() As Integer
  Dim fn As Integer
  Dim strLine As String
  Dim res As Integer
  Dim i As Integer
  #If FinalMode Then
  On Error GoTo justend
  #End If
  res = -1
  fn = FreeFile
  Open App.path & "\" & "hotkeys.ini" For Output As #fn
    Print #fn, CStr(NumberOfHotkeys)
    For i = 1 To NumberOfHotkeys
      strLine = "#" & HotkeyIDFixedLen(Hotkeys(i - 1).key1) & " + #" & _
       HotkeyIDFixedLen(Hotkeys(i - 1).key2) & _
       " : " & Hotkeys(i - 1).command
      Print #fn, strLine
    Next i
  Close #fn
  res = 0
justend:
  SaveHotkeys = res
End Function

Public Sub SafeJustEndofLoadHotkeys()
  On Error GoTo justend
  NumberOfHotkeys = 0
  ReDim Hotkeys(0)
  With gridHotkeys
  .Rows = 2
  .Row = 1
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  .TextMatrix(1, 0) = ""
  .TextMatrix(1, 1) = ""
  .TextMatrix(1, 2) = ""
  End With
  Exit Sub
justend:
 LogOnFile "errors.txt", "Error caught at SafeJustEndofLoadHotkeys(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub
Public Function LoadHotkeys() As Integer
  Dim fn As Integer
  Dim strLine As String
  Dim i As Integer
  Dim key1s As String
  Dim key2s As String
  Dim comms As String
  On Error GoTo justend
  fn = FreeFile
  Open App.path & "\hotkeys.ini" For Input As #fn
    Line Input #fn, strLine
    NumberOfHotkeys = CLng(strLine)
    If NumberOfHotkeys = 0 Then
      ReDim Hotkeys(0)
    Else
      ReDim Hotkeys(NumberOfHotkeys - 1)
    End If
    For i = 0 To (NumberOfHotkeys - 1)
      Line Input #fn, strLine
      key1s = Mid(strLine, 2, 3)
      key2s = Mid(strLine, 9, 3)
      comms = Right(strLine, Len(strLine) - 14)
      Hotkeys(i).key1 = CByte(CLng(key1s))
      Hotkeys(i).key2 = CByte(CLng(key2s))
      Hotkeys(i).command = comms
    Next i
  Close #fn
  gridHotkeys.Rows = NumberOfHotkeys + 2
  For i = 1 To NumberOfHotkeys
    With gridHotkeys
    .TextMatrix(i, 0) = TranslateHotkeyID(Hotkeys(i - 1).key1)
    .TextMatrix(i, 1) = TranslateHotkeyID(Hotkeys(i - 1).key2)
    .TextMatrix(i, 2) = Hotkeys(i - 1).command
    .Row = i
    .Col = 0
    .CellAlignment = flexAlignCenterCenter
    .Col = 1
    .CellAlignment = flexAlignCenterCenter
    .Col = 2
    .CellAlignment = flexAlignLeftCenter
    End With
  Next i
  With gridHotkeys
  .Row = NumberOfHotkeys + 1
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  .TextMatrix(i, 0) = ""
  .TextMatrix(i, 1) = ""
  .TextMatrix(i, 2) = ""
  LoadHotkeys = 0
  End With
  Exit Function
justend:
  SafeJustEndofLoadHotkeys
  LoadHotkeys = -1
End Function

Private Sub cmdSave_Click()
  Dim res As Integer
  ReleasePressKey
  res = SaveHotkeys()
  If res = -1 Then
    lblDebug.Caption = "Could not save hotkeys.ini"
    lblDebug.ForeColor = &HFF&
  Else
    lblDebug.Caption = "Saved hotkeys.ini successfully"
    lblDebug.ForeColor = &HFF00&
  End If
End Sub

Private Sub Form_Load()
  Dim res As Integer
  Dim sRes As String
  #If FinalMode Then
  On Error GoTo justend
  #End If
  reenableHotkeyTime = 0
  With gridHotkeys
  .ColWidth(0) = 1600
  .ColWidth(1) = 1600
  .ColWidth(2) = 3200
  .TextMatrix(0, 0) = "key1"
  .TextMatrix(0, 1) = "key2 (opt)"
  .TextMatrix(0, 2) = "Command / message to say"
  .TextMatrix(1, 0) = ""
  .TextMatrix(1, 1) = ""
  .TextMatrix(1, 2) = ""
  .Row = 0
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  .Row = 1
  .Col = 0
  .CellAlignment = flexAlignCenterCenter
  .Col = 1
  .CellAlignment = flexAlignCenterCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  End With
  res = LoadHotkeys()
  If res = -1 Then
    lblDebug.Caption = "Could not load hotkeys.ini"
    lblDebug.ForeColor = &HFF&
  Else
    lblDebug.Caption = "Loaded hotkeys.ini successfully"
    lblDebug.ForeColor = &HFF00&
  End If
  espectingHotkey = False
  lastHotkeyCol = 0
  lastHotkeyRow = 0
  sRes = InitDI()
  If sRes = "" Then
    timerHotkeys.enabled = True
  Else
    LogOnFile "errors.txt", "Could not load hotkey module." & debugdxError & vbCrLf & " This error happened exactly while trying to execute this line:" & vbCrLf & sRes
  End If
  Exit Sub
justend:
  LogOnFile "errors.txt", "Could not load hotkey module. Err number: " & CStr(Err.Number) & " ; Err description: " & Err.Description
End Sub

Private Sub Form_Resize()
  ReleasePressKey
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ReleasePressKey
  Me.Hide
  Cancel = BlockUnload
End Sub
Private Sub ReleasePressKey()
  If espectingHotkey = True Then
    espectingHotkey = False
    If lastHotkeyCol = 0 Then
      gridHotkeys.TextMatrix(lastHotkeyRow, lastHotkeyCol) = TranslateHotkeyID(Hotkeys(lastHotkeyRow - 1).key1)
    Else
      gridHotkeys.TextMatrix(lastHotkeyRow, lastHotkeyCol) = TranslateHotkeyID(Hotkeys(lastHotkeyRow - 1).key2)
    End If
  End If
End Sub
Private Sub gridHotkeys_Click()
  Dim sCol As Long
  Dim sRow As Long
  Dim thereIsSelection As Boolean
  sCol = gridHotkeys.Col
  sRow = gridHotkeys.Row
  If sRow = (NumberOfHotkeys + 1) Then
    ReDim Preserve Hotkeys(NumberOfHotkeys)
    Hotkeys(NumberOfHotkeys).key1 = 0
    Hotkeys(NumberOfHotkeys).key2 = 0
    Hotkeys(NumberOfHotkeys).command = ""
    NumberOfHotkeys = NumberOfHotkeys + 1
    gridHotkeys.Rows = gridHotkeys.Rows + 1
    gridHotkeys.TextMatrix(NumberOfHotkeys, 0) = TranslateHotkeyID(Hotkeys(NumberOfHotkeys - 1).key1)
    gridHotkeys.TextMatrix(NumberOfHotkeys, 1) = TranslateHotkeyID(Hotkeys(NumberOfHotkeys - 1).key2)
    gridHotkeys.TextMatrix(NumberOfHotkeys, 2) = ""
    gridHotkeys.Row = gridHotkeys.Rows - 1
    gridHotkeys.Col = 0
    gridHotkeys.CellAlignment = flexAlignCenterCenter
    gridHotkeys.Col = 1
    gridHotkeys.CellAlignment = flexAlignCenterCenter
    gridHotkeys.Col = 2
    gridHotkeys.CellAlignment = flexAlignLeftCenter
    gridHotkeys.Col = sCol
    gridHotkeys.Row = sRow
  End If
  ReleasePressKey
  If (gridHotkeys.RowSel = sRow) And (gridHotkeys.ColSel = sCol) Then
    thereIsSelection = False
  Else
    thereIsSelection = True
  End If
  If (sRow <> 0) And (thereIsSelection = False) Then
    lastHotkeyCol = sCol
    lastHotkeyRow = sRow
    If sCol < 2 Then
      gridHotkeys.TextMatrix(sRow, sCol) = "<PRESS KEY>"
      espectingHotkey = True
    End If
  Else
    lastHotkeyCol = 0
    lastHotkeyRow = 0
  End If
End Sub


Private Sub gridHotkeys_KeyPress(KeyAscii As Integer)
  Dim sCol As Long
  Dim sRow As Long
  Dim thereIsSelection As Boolean
  Dim lenS As Long
  sCol = gridHotkeys.Col
  sRow = gridHotkeys.Row
  If (gridHotkeys.RowSel = sRow) And (gridHotkeys.ColSel = sCol) Then
    thereIsSelection = False
  Else
    thereIsSelection = True
  End If
  If (sCol = 2) And (thereIsSelection = False) Then
  If sRow = (NumberOfHotkeys + 1) Then
    ReDim Preserve Hotkeys(NumberOfHotkeys)
    Hotkeys(NumberOfHotkeys).key1 = 0
    Hotkeys(NumberOfHotkeys).key2 = 0
    Hotkeys(NumberOfHotkeys).command = ""
    NumberOfHotkeys = NumberOfHotkeys + 1
    gridHotkeys.Rows = gridHotkeys.Rows + 1
    gridHotkeys.TextMatrix(NumberOfHotkeys, 0) = TranslateHotkeyID(Hotkeys(NumberOfHotkeys - 1).key1)
    gridHotkeys.TextMatrix(NumberOfHotkeys, 1) = TranslateHotkeyID(Hotkeys(NumberOfHotkeys - 1).key2)
    gridHotkeys.TextMatrix(NumberOfHotkeys, 2) = ""
    gridHotkeys.Row = gridHotkeys.Rows - 1
    gridHotkeys.Col = 0
    gridHotkeys.CellAlignment = flexAlignCenterCenter
    gridHotkeys.Col = 1
    gridHotkeys.CellAlignment = flexAlignCenterCenter
    gridHotkeys.Col = 2
    gridHotkeys.CellAlignment = flexAlignLeftCenter
    gridHotkeys.Col = sCol
    gridHotkeys.Row = sRow
  End If
  If KeyAscii = 8 Then
    lenS = Len(gridHotkeys.TextMatrix(sRow, sCol))
    If lenS > 0 Then
      gridHotkeys.TextMatrix(sRow, sCol) = Left(gridHotkeys.TextMatrix(sRow, sCol), lenS - 1)
    End If
  Else
   gridHotkeys.TextMatrix(sRow, sCol) = gridHotkeys.TextMatrix(sRow, sCol) & Chr(KeyAscii)
  End If
  Hotkeys(sRow - 1).command = gridHotkeys.TextMatrix(sRow, sCol)
  End If
End Sub

Private Sub timerHotkeys_Timer()
  #If FinalMode Then
  On Error GoTo justend
  #End If
  If HotkeysAreUsable = False Then
    Exit Sub
  End If
  If SoundIsUsable = True Then
  Dim i As Integer
  Dim limhot As Integer
  Dim aRes As Integer
  Dim activated As Boolean
  Dim gt As Long
  DIV.GetDeviceStateKeyboard KeyB
  If (espectingHotkey = True) Then ' defining hotkeys
    aRes = 0
    For i = 1 To 255
      If KeyB.key(i) > &H0 Then
        aRes = i
      End If
    Next i
    If aRes > 0 Then
      If lastHotkeyCol = 0 Then
        Hotkeys(lastHotkeyRow - 1).key1 = aRes
        gridHotkeys.TextMatrix(lastHotkeyRow, lastHotkeyCol) = TranslateHotkeyID(Hotkeys(lastHotkeyRow - 1).key1)
      Else
        Hotkeys(lastHotkeyRow - 1).key2 = aRes
        gridHotkeys.TextMatrix(lastHotkeyRow, lastHotkeyCol) = TranslateHotkeyID(Hotkeys(lastHotkeyRow - 1).key2)
      End If
      espectingHotkey = False
    End If
  Else ' playing
    If chkHotkeysActivated.Value = 0 Then
      Exit Sub
    End If
    gt = GetTickCount()
    limhot = NumberOfHotkeys - 1
    If (chkRepeat.Value = 1) Then
      If (gt > reenableHotkeyTime) Then 'reactivate all hotbuttons
        For i = 0 To limhot
          Hotkeys(i).usable = True
        Next i
      End If
    End If
    For i = 0 To limhot
      activated = False
      If (KeyB.key(Hotkeys(i).key1) > 0) Or (Hotkeys(i).key1 = 0) Then
        If (KeyB.key(Hotkeys(i).key2) > 0) Or (Hotkeys(i).key2 = 0) Then
          If Not ((Hotkeys(i).key1 = 0) And (Hotkeys(i).key2 = 0)) Then
            activated = True
            If Hotkeys(i).usable = True Then
              aRes = ExecuteInFocusedTibia(Hotkeys(i).command)
              Hotkeys(i).usable = False
              reenableHotkeyTime = gt + CLng(Me.txtDelay.Text)
            End If
          End If
        End If
      End If
      If (activated = False) And (Hotkeys(i).usable = False) Then
        Hotkeys(i).usable = True
      End If
    Next i
  End If
  End If
Exit Sub
justend:
 LogOnFile "errors.txt", "Error caught at timerHotkeys_Timer(). Err number " & CStr(Err.Number) & " ; Err description " & Err.Description
End Sub
