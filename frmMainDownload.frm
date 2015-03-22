VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMainDownload 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Proxy Update Manager 1.0"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   Icon            =   "frmMainDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timerStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8040
      Top             =   360
   End
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   4215
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7435
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMainDownload.frx":0442
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4200
      Picture         =   "frmMainDownload.frx":04C4
      ScaleHeight     =   735
      ScaleWidth      =   3615
      TabIndex        =   7
      Top             =   0
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancelDownload 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel download"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin ComctlLib.ProgressBar pgBar 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   5400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdButton1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Run Blackd Proxy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin ComctlLib.ProgressBar pgTotal 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5760
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblProgress 
      BackColor       =   &H00000000&
      Caption         =   "Please wait while Blackd Proxy gets updated..."
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
      TabIndex        =   9
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Label lblOfficial 
      BackColor       =   &H00000000&
      Caption         =   "Blackd Proxy official site: www.blackdtools.com"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Total progress:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Current file:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmMainDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As Scripting.FileSystemObject
Dim strLogCambios As String
Dim DictLogPrevio As Dictionary
Dim DictLogPrevio2 As Dictionary
Dim lngCurrentVersion As Long
Dim lngWantedVersion As Long

Public Sub OverwriteOnFileSimple(file_name As String, strText As String)
  Dim fn As Integer
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
    writeThis = strText
  Open App.Path & "\" & file_name For Output As #fn
    Print #fn, writeThis
  Close #fn

  Exit Sub
ignoreit:
  a = -1
End Sub


Public Function ReadLongFromFile(file_name As String) As Long
    On Error GoTo ignoreit
    Dim fn As Integer
    Dim strValue As String
    Dim a As Long
    Dim fs As Scripting.FileSystemObject
    a = 0
    Set fs = New Scripting.FileSystemObject
    If fs.FileExists(App.Path & "\" & file_name) = False Then
          ReadLongFromFile = 0
          Exit Function
    End If
    Set fs = Nothing
    fn = FreeFile
    Open App.Path & "\" & file_name For Input As #fn
      Line Input #fn, strValue
    Close #fn
    ReadLongFromFile = CLng(strValue)
    Exit Function
ignoreit:
 ReadLongFromFile = 0
End Function


Public Sub AddLog(strMsg As String)
    If txtLog.Text = "" Then
         txtLog.Text = strMsg
    Else
         txtLog.Text = txtLog.Text & vbCrLf & strMsg
    End If
    txtLog.SelStart = Len(txtLog.Text)
    DoEvents
End Sub

Public Sub pintar(dblValor As Double)
    Dim percentTotalSoFar As Double
    Dim realDblValor As Double
    Dim tmp1 As Double
    Dim tmp2 As Double
    Dim tmp3 As Double
    If dblTotalToDownload > 0 Then
        pgBar.Value = CInt(Round(dblValor, 0))
        tmp1 = dblValor / 100
        realDblValor = dblProgressPortion * tmp1
        tmp2 = dblProgressSoFar + realDblValor
        tmp3 = dblTotalToDownload
        percentTotalSoFar = (tmp2 / tmp3) * 100
        If percentTotalSoFar > 100 Then
            percentTotalSoFar = 100
        End If
        lblProgress.Caption = "Downloaded " & CStr(Round(percentTotalSoFar, 2)) & "% of " & dblTotalToDownload & " bytes"
        pgTotal.Value = CInt(Round(percentTotalSoFar, 0))
    End If
    DoEvents
End Sub



Public Sub AddFileVersion(ByVal strPath As String, ByVal lngValue As Long)
  ' add item to dictionary
  Dim res As Boolean
  DictLogPrevio.Item(strPath) = lngValue
  Exit Sub
End Sub

Public Function GetFileVersion(ByVal strPath As String) As Long
  On Error GoTo goterr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  If DictLogPrevio.Exists(strPath) = True Then
    GetFileVersion = DictLogPrevio.Item(strPath)
  Else
    GetFileVersion = -1
  End If
  Exit Function
goterr:
  GetFileVersion = -1
End Function


Public Sub AddFileSize(ByVal strPath As String, ByVal lngValue As Long)
  ' add item to dictionary
  Dim res As Boolean
  DictLogPrevio2.Item(strPath) = lngValue
  Exit Sub
End Sub

Public Function GetFileSize(ByVal strPath As String) As Long
  On Error GoTo goterr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  If DictLogPrevio2.Exists(strPath) = True Then
    GetFileSize = DictLogPrevio2.Item(strPath)
  Else
    GetFileSize = -1
  End If
  Exit Function
goterr:
  GetFileSize = -1
End Function

Private Function ReadFileHere(strFile As String) As String
    On Error GoTo goterr
    Dim fs As Scripting.FileSystemObject
    Dim fn As Integer
    Dim strValue As String
    Dim strReaded As String
    strReaded = ""
    Set fs = New Scripting.FileSystemObject
    If fs.FileExists(strFile) = False Then
          ReadFileHere = ""
          Exit Function
    End If
    Set fs = Nothing
    fn = FreeFile
    Open strFile For Input As #fn
      While Not EOF(fn)
        Line Input #fn, strValue
        If strReaded = "" Then
          strReaded = strValue
        Else
          strReaded = strReaded & vbCrLf & strValue
        End If
      Wend
    Close #fn
    ReadFileHere = strReaded
    Exit Function
goterr:
    AddLog "Error " & CStr(Err.Number) & " : " & Err.Description
    ReadFileHere = ""
End Function
Private Sub ReadUpdatesFromWeb()
    On Error GoTo goterr
    Dim strReceived As String
    Dim lngNextFilePos As Long
    Dim lngNextPart As Long
    Dim strParte(2) As String
    Dim i As Long
    Dim strLonTot As Long
    Dim blnRes As Boolean
    Dim strAqui As String
    Dim strRes As String
    dblTotalToDownload = 0
    InitFilesToDownload
    DictLogPrevio.RemoveAll
    DictLogPrevio2.RemoveAll
     
    blnRes = conectar("Updater FTP", "ftp.blackdtools.com", "autoupdate@blackdtools.com", "p010101010101", "")
    If estaConectado() = False Then
          AddLog "Unable to connect to blackdtools FTP server."
          cmdButton1.Caption = "Retry"
          cmdButton1.Enabled = True
          Exit Sub
    End If
    strRes = bajar("/", "index.php", frmMainDownload)
    If strRes <> "" Then
        AddLog "Failed to download update list: " & strRes
        ScanDone = False
        Exit Sub
    End If
    DoEvents
    strAqui = App.Path
    If Right$(strAqui, 1) = "/" Or Right$(strAqui, 1) = "\" Then
       strAqui = Left$(strAqui, Len(strAqui) - 1)
    End If
    strReceived = ReadFileHere(strAqui & "\index.php")
    If strReceived = "" Then
        AddLog "Failed to read update list from hard disk."
        ScanDone = False
        Exit Sub
    End If
    lngNextFilePos = 1
    lngNextFilePos = InStr(lngNextFilePos, strReceived, "<td>")
    strLonTot = Len(strReceived)
    While lngNextFilePos <> 0
      strParte(0) = ""
      strParte(1) = ""
      strParte(2) = ""
      For i = 0 To 2
        If lngNextFilePos <> 0 Then
          lngNextPart = InStr(lngNextFilePos, strReceived, "</td>")
          If lngNextPart = 0 Then
              lngNextFilePos = 0
          End If
        End If
        If lngNextFilePos <> 0 Then
          strParte(i) = Mid$(strReceived, lngNextFilePos + 4, lngNextPart - lngNextFilePos - 4)
        End If
        If lngNextFilePos <> 0 Then
          lngNextFilePos = InStr(lngNextPart, strReceived, "<td>")
        End If
      Next i
      If lngNextPart <> 0 Then
        If strParte(0) = "LAST_VERSION" Then
            lngWantedVersion = CLng(strParte(1))
        Else
            SimpleAddFile strParte(0), CLng(strParte(1)), CLng(strParte(2))
        End If
      End If
    Wend
    ScanDone = True
    Exit Sub
goterr:
    AddLog "Failed to download update list." & vbCrLf & "Error code = " & CStr(Err.Number) & vbCrLf & _
    "Error description = " & Err.Description
    ScanDone = False
End Sub


Private Sub InitFilesToDownload()
  NumberOfFilesToDownload = 0
  txtLog.Text = "Checking for updated files, please wait..."
End Sub

Private Sub SimpleAddFile(strFileName As String, _
      lngVersion As Long, strSize As Long)
      On Error GoTo goterr
      Dim lngCurRow As Long
      Dim i As Long
      Dim blnCheck As Boolean
      Dim strPath As String
      Dim strAll As String
      Dim fs As Scripting.FileSystemObject
      Dim fil As Scripting.File
      Set fs = New Scripting.FileSystemObject
      strPath = App.Path
      If ((Right$(strPath, 1) = "\") Or (Right$(strPath, 1) = "/")) Then
        strPath = Left$(strPath, Len(strPath) - 1)
      End If
    
      FilesToDownload(NumberOfFilesToDownload, 0) = strFileName
      FilesToDownload(NumberOfFilesToDownload, 1) = CStr(lngVersion)
      FilesToDownload(NumberOfFilesToDownload, 2) = CStr(strSize)
      strAll = strPath & "\" & strFileName
      If ((Right$(strAll, 1) = "\") Or (Right$(strAll, 1) = "/")) Then
            blnCheck = (fs.FolderExists(strAll) = False)
            If blnCheck = True Then
                FilesToDownload(NumberOfFilesToDownload, 3) = cte_MakeFolderRequired
            Else
                FilesToDownload(NumberOfFilesToDownload, 3) = cte_NoUpdateRequired
                dblTotalToDownload = dblTotalToDownload + CDbl(strSize)
            End If
      Else
            If fs.FileExists(strAll) = False Then
                FilesToDownload(NumberOfFilesToDownload, 3) = cte_DownloadRequired
                dblTotalToDownload = dblTotalToDownload + CDbl(strSize)
            Else
                If (lngVersion > CLng(lngCurrentVersion)) Then
                    FilesToDownload(NumberOfFilesToDownload, 3) = cte_DownloadRequired
                    dblTotalToDownload = dblTotalToDownload + CDbl(strSize)
                Else
                    Set fil = fs.GetFile(strAll)
                    If fil.Size <> CLng(strSize) Then
                        FilesToDownload(NumberOfFilesToDownload, 3) = cte_DownloadRequired
                        dblTotalToDownload = dblTotalToDownload + CDbl(strSize)
                    Else
                        FilesToDownload(NumberOfFilesToDownload, 3) = cte_NoUpdateRequired
                    End If
                End If
            End If
      End If
      Set fs = Nothing
      NumberOfFilesToDownload = NumberOfFilesToDownload + 1
      Exit Sub
goterr:
      MsgBox "Got error " & CStr(Err.Number) & " at SimpleAddFile : " & Err.Description, vbOKOnly + vbCritical, "Error"
End Sub






Private Sub canceldownload()
    finbajada = 2
    DoEvents
End Sub

Private Sub cmdButton1_Click()
    If cmdButton1.Caption = "Run Blackd Proxy" Then
       Dim a As Long
       a = ShellExecute(Me.hwnd, "Open", App.Path & "\blackdproxy.exe", &O0, &O0, SW_NORMAL)
       End
    Else
    finbajada = 0
    DoUpdateFromBlackdtools
    End If
End Sub

Private Sub cmdCancelDownload_Click()
    AddLog "Trying to cancel download, please wait..."
    canceldownload
End Sub




Private Sub DoUpdateFromBlackdtools()
  On Error GoTo goterr
  cmdButton1.Enabled = False
  cmdCancelDownload.Enabled = True
  Dim strGuardado As String
  Dim blnRes As Boolean
  Dim strRes As String
  Dim lngLim As Long
  Dim strRuta As String
  Dim strAct As String
  Dim fs As Scripting.FileSystemObject
  Dim fol As Scripting.Folder
  Dim strRealDest As String
  Dim sDest As String
  Dim blnStarted As Boolean
  Dim i As Long
  blnStarted = False
  ReadUpdatesFromWeb
  If ScanDone = False Then
       'AddLog "Unable to read file list"
       cmdButton1.Caption = "Retry"
       cmdButton1.Enabled = True
       Exit Sub
  End If
  lngLim = NumberOfFilesToDownload
  If dblTotalToDownload = 0 Then
      AddLog "All already updated."
  Else
      dblProgressSoFar = 0
      pgTotal.Value = 0
      DoEvents
      i = 0
      While i <= lngLim
          If finbajada = 2 Then
                AddLog "Download cancelled by user."
                cmdButton1.Caption = "Retry"
                cmdButton1.Enabled = True
                Exit Sub
          End If
          strAct = FilesToDownload(i, 3)
          Select Case strAct
          Case cte_MakeFolderRequired
                strRealDest = App.Path
                If Right$(strRealDest, 1) = "\" Or Right$(strRealDest, 1) = "/" Then
                      strRealDest = Left$(strRealDest, Len(strRealDest) - 1)
                End If
                sDest = strRealDest & "\" & filtroBD(FilesToDownload(i, 0))
                AddLog "Making folder " & FilesToDownload(i, 0) & " ..."
                Set fs = New Scripting.FileSystemObject
                Set fol = fs.CreateFolder(sDest)
                Set fol = Nothing
                Set fs = Nothing
                i = i + 1
          Case cte_DownloadRequired
                If blnStarted = False Then
                    AddLog "Now updating to Blackd Proxy " & Round(CDbl(lngWantedVersion) / 1000, 3)
                    blnStarted = True
                End If
                dblProgressPortion = CDbl(FilesToDownload(i, 2))
                pintar 0
                
                                          strRealDest = App.Path
                          If Right$(strRealDest, 1) = "\" Or Right$(strRealDest, 1) = "/" Then
                                 strRealDest = Left$(strRealDest, Len(strRealDest) - 1)
                          End If
                          sDest = strRealDest & "\" & filtroBD(FilesToDownload(i, 0))
                             Set fs = New Scripting.FileSystemObject
                          If fs.FileExists(sDest) = True Then
                                fs.DeleteFile sDest, True
                          End If
                          Set fs = Nothing
                blnRes = conectar("Updater FTP", "ftp.blackdtools.com", "autoupdate@blackdtools.com", "p010101010101", "")
                If estaConectado() = False Then
                      AddLog "Unable to connect to blackdtools FTP server."
                      cmdButton1.Caption = "Retry"
                      cmdButton1.Enabled = True
                      Exit Sub
                End If
                If finbajada = 2 Then
                    AddLog "Download cancelled by user."
                    cmdButton1.Caption = "Retry"
                    cmdButton1.Enabled = True
                    Exit Sub
                Else
                    finbajada = 0
                End If
                strRuta = CStr(lngWantedVersion) & "/"
                AddLog "Downloading " & FilesToDownload(i, 0) & " ..."
                strRes = bajar(strRuta, FilesToDownload(i, 0), frmMainDownload)
                If strRes <> "" Then
                      AddLog "Download failed: " & FilesToDownload(i, 0) & " : " & strRes
                Else
                      Do
                          If (finbajada > 0) Then
                              Exit Do
                          End If
                          DoEvents
                      Loop
                      Select Case finbajada
                      Case 1
                          dblProgressSoFar = dblProgressSoFar + FilesToDownload(i, 2)
                          pintar 0
                          i = i + 1
                      Case 2
                          blnRes = desconectar()
                          strRealDest = App.Path
                          If Right$(strRealDest, 1) = "\" Or Right$(strRealDest, 1) = "/" Then
                                 strRealDest = Left$(strRealDest, Len(strRealDest) - 1)
                          End If
                          sDest = strRealDest & "\" & filtroBD(FilesToDownload(i, 0))
                          Set fs = New Scripting.FileSystemObject
                          If fs.FileExists(sDest) = True Then
                                fs.DeleteFile sDest, True
                                AddLog FilesToDownload(i, 0) & " deleted."
                          End If
                          Set fs = Nothing
                          AddLog "Download cancelled by user."
                          cmdButton1.Caption = "Retry"
                          cmdButton1.Enabled = True
                          Exit Sub
                      Case 3
                          AddLog "Download failed: " & FilesToDownload(i, 0)
                      End Select
                End If
                DoEvents
                blnRes = desconectar()
                DoEvents
          Case Else
                DoEvents
          
                i = i + 1
                
          End Select
      Wend
  End If
  AddLog "Download completed."
  blnRes = desconectar()
  pgBar.Value = 100
  pgTotal.Value = 100
  lblProgress.Caption = "Blackd Proxy is now updated!"
  OverwriteOnFileSimple "version.cfg", CStr(lngWantedVersion)
  cmdButton1.Caption = "Run Blackd Proxy"
  cmdButton1.Enabled = True
  cmdCancelDownload.Enabled = False
  Exit Sub
goterr:
  MsgBox "Got error " & CStr(Err.Number) & " at DoUpdateFromBlackdtools : " & Err.Description, vbOKOnly + vbCritical, "Error"
End Sub






Private Sub Form_Load()
    Set DictLogPrevio = New Dictionary
    Set DictLogPrevio2 = New Dictionary
    lngCurrentVersion = ReadLongFromFile("version.cfg")
    lngWantedVersion = 9380
    ScanDone = False
    timerStart.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo justend
    desconectar
justend:
    End
End Sub



Private Sub timerStart_Timer()
  On Error GoTo goterr
    timerStart.Enabled = False
    WebBrowser1.Navigate "http://www.blackdtools.com/lastchanges.php"
    DoEvents
    WebBrowser1.Visible = True
    DoUpdateFromBlackdtools
  Exit Sub
goterr:
  MsgBox "Got error " & CStr(Err.Number) & " at timerStart_Timer : " & Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

