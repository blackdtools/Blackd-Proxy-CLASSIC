VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHPmana 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HP and Mana"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7770
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmHPmana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLimitRandomizator 
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
      Left            =   3960
      TabIndex        =   51
      Text            =   "10"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdSaveLimitRandomizator 
      Caption         =   "Change ( current = 10 % )"
      Height          =   285
      Left            =   4800
      TabIndex        =   50
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdSaveRecast2 
      Caption         =   "Change ( current = 700 ms )"
      Height          =   285
      Left            =   4800
      TabIndex        =   47
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtRecast2 
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
      Left            =   3960
      TabIndex        =   46
      Text            =   "700"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdDeleteSel 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE SELECTED SETTING"
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeleteAll 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE ALL SETTINGS"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Timer tmrHPmana 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   480
   End
   Begin VB.TextBox txtRecast 
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
      Left            =   3960
      TabIndex        =   37
      Text            =   "300"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdSaveRecast 
      Caption         =   "Change ( current = 300 ms )"
      Height          =   285
      Left            =   4800
      TabIndex        =   38
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080FF80&
      Caption         =   "ADD AS NEW SETTING"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdSaveOnMemory 
      BackColor       =   &H0080FFFF&
      Caption         =   "OVERWRITE"
      Height          =   255
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Frame fraHP 
      BackColor       =   &H00000000&
      Caption         =   "HP SETTINGS"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   3735
      Begin VB.OptionButton HPopt11 
         BackColor       =   &H00000000&
         Caption         =   "Heal with small health potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   3840
         Width           =   2535
      End
      Begin VB.OptionButton HPopt10 
         BackColor       =   &H00000000&
         Caption         =   "Heal with great spirit potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   3600
         Width           =   2535
      End
      Begin VB.OptionButton HPopt9 
         BackColor       =   &H00000000&
         Caption         =   "Heal with ultimate health potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   3360
         Width           =   2535
      End
      Begin VB.OptionButton HPopt7 
         BackColor       =   &H00000000&
         Caption         =   "Heal with great health potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox hpTEXT 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Text            =   "exura vita"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.OptionButton HPopt8 
         BackColor       =   &H00000000&
         Caption         =   "Cast spell or use command ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2760
         Width           =   2535
      End
      Begin VB.OptionButton HPopt6 
         BackColor       =   &H00000000&
         Caption         =   "Heal with strong health potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2280
         Width           =   3015
      End
      Begin VB.OptionButton HPopt5 
         BackColor       =   &H00000000&
         Caption         =   "Heal with health potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2040
         Width           =   3015
      End
      Begin VB.OptionButton HPopt4 
         BackColor       =   &H00000000&
         Caption         =   "Heal with life fluid"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton HPopt3 
         BackColor       =   &H00000000&
         Caption         =   "Heal with UH rune"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton HPopt2 
         BackColor       =   &H00000000&
         Caption         =   "Heal with IH rune"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   3015
      End
      Begin VB.OptionButton HPopt1 
         BackColor       =   &H00000000&
         Caption         =   "Do nothing"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar scrollHP 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   10
         Top             =   600
         Value           =   63
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "When the %HP is less than ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblHPvalue 
         BackColor       =   &H00000000&
         Caption         =   "63 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraMANA 
      BackColor       =   &H00000000&
      Caption         =   "MANA SETTINGS"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   3960
      TabIndex        =   7
      Top             =   4560
      Width           =   3735
      Begin VB.OptionButton MANAopt7 
         BackColor       =   &H00000000&
         Caption         =   "Recharge with great spirit potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   2880
         Width           =   3015
      End
      Begin VB.OptionButton MANAopt2 
         BackColor       =   &H00000000&
         Caption         =   "Recharge with mana fluid"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox manaTEXT 
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Text            =   "exiva close"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.OptionButton MANAopt6 
         BackColor       =   &H00000000&
         Caption         =   "Cast spell or use command ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2280
         Width           =   2535
      End
      Begin VB.OptionButton MANAopt5 
         BackColor       =   &H00000000&
         Caption         =   "Recharge with great mana potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   2040
         Width           =   3015
      End
      Begin VB.OptionButton MANAopt4 
         BackColor       =   &H00000000&
         Caption         =   "Recharge with strong mana potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton MANAopt3 
         BackColor       =   &H00000000&
         Caption         =   "Recharge with mana potion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1560
         Width           =   3015
      End
      Begin VB.OptionButton MANAopt1 
         BackColor       =   &H00000000&
         Caption         =   "Do nothing"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.HScrollBar scrollMANA 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   13
         Top             =   600
         Value           =   50
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "When the %MANA is less than ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblMANAvalue 
         BackColor       =   &H00000000&
         Caption         =   "50 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbCharacter 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Text            =   "-"
      Top             =   4080
      Width           =   4095
   End
   Begin VB.CommandButton cmdLoadFromHD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RELOAD FROM HD"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveOnHD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SAVE ON HD"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid gridHPmana 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      BackColorBkg    =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7680
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   49
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "RECHARGE LIMITS RANDOMIZED BY... :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   48
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "To ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   45
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "From ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   44
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "IMPORTANT: when doing multiheal ... SET SMALLER HEAL PERCENTS FIRST!"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Label lblRecast 
      BackColor       =   &H00000000&
      Caption         =   "GLOBAL RECAST TIMER (in ms) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblMsg 
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
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   3960
      TabIndex        =   32
      Top             =   8280
      Width           =   3735
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Character name (you are free to type any) :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7680
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Current settings loaded in memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"frmHPmana.frx":0442
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmHPmana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Private Sub cmdAdd_Click()
    Dim strSel As String
    Dim strMANAact As String
    Dim strHPact As String
    Dim bMANA As Byte
    Dim bHP As Byte

    
    strSel = LCase(cmbCharacter.Text)
    If strSel = "" Or strSel = "-" Then
        ' do nothing
        lblMsg.Caption = "Select/type a character first!"
        lblMsg.ForeColor = &HFFFF&
    Else
        bHP = CByte(scrollHP.Value)
        bMANA = CByte(scrollMANA.Value)
        If HPopt1.Value = True Then
            strHPact = "no action"
        ElseIf HPopt2.Value = True Then
            strHPact = "IH"
        ElseIf HPopt3.Value = True Then
            strHPact = "UH"
        ElseIf HPopt4.Value = True Then
            strHPact = "life fluid"
        ElseIf HPopt5.Value = True Then
            strHPact = "health potion"
        ElseIf HPopt6.Value = True Then
            strHPact = "strong health potion"
        ElseIf HPopt7.Value = True Then
            strHPact = "great health potion"
        ElseIf HPopt9.Value = True Then
            strHPact = "ultimate health potion"
        ElseIf HPopt10.Value = True Then
            strHPact = "great spirit potion"
        ElseIf HPopt11.Value = True Then
            strHPact = "small health potion"
        ElseIf HPopt8.Value = True Then
            strHPact = hpTEXT.Text
        End If
        
        If MANAopt1.Value = True Then
            strMANAact = "no action"
        ElseIf MANAopt2.Value = True Then
            strMANAact = "mana fluid"
        ElseIf MANAopt3.Value = True Then
            strMANAact = "mana potion"
        ElseIf MANAopt4.Value = True Then
            strMANAact = "strong mana potion"
        ElseIf MANAopt5.Value = True Then
            strMANAact = "great mana potion"
        ElseIf MANAopt7.Value = True Then
            strMANAact = "great spirit potion"
        ElseIf MANAopt6.Value = True Then
            strMANAact = manaTEXT.Text
        End If
        
        AddHPmanaSetting strSel, bHP, strHPact, bMANA, strMANAact
        DisplayLoadedHPmanaConfig
        lblMsg.Caption = "Character <" & strSel & "> saved sucesfully!"
        lblMsg.ForeColor = &HFF00&
    End If
End Sub

Private Sub cmdDeleteAll_Click()
    DeleteAllSettings
    DisplayLoadedHPmanaConfig
End Sub

Private Sub cmdDeleteSel_Click()
    Dim lineSel As Long
    lineSel = Me.gridHPmana.RowSel
    If lineSel < 1 Then
        lblMsg.Caption = "First select the setting that you want to delete!"
        lblMsg.ForeColor = &HFFFF&
        Exit Sub
    End If
    deleteHPmanaSetting lineSel
    DisplayLoadedHPmanaConfig
End Sub

Private Sub cmdLoadFromHD_Click()
    Dim strRes As String
    strRes = LoadHPmanaConfig()
    If strRes = "" Then
         lblMsg.Caption = "Loaded sucesfully!"
         lblMsg.ForeColor = &HFF00&
    Else
         lblMsg.Caption = strRes
         lblMsg.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdSaveLimitRandomizator_Click()
    On Error GoTo goterr
    Dim lngCast As Long
    lngCast = CLng(Me.txtLimitRandomizator.Text)
    If ((lngCast >= 0) And (lngCast <= 99)) Then
        LimitRandomizator = lngCast
        Me.txtLimitRandomizator.Text = CStr(LimitRandomizator)
        cmdSaveLimitRandomizator.Caption = "CHANGE ( current = " & CStr(LimitRandomizator) & " % )"
    Else
        GoTo goterr
    End If
    lblMsg.Caption = "Changed limit randomizator succesfully"
    lblMsg.ForeColor = &HFF00&
    Exit Sub
goterr:
    lblMsg.Caption = "Invalid setting"
    lblMsg.ForeColor = &HFF&
End Sub

Private Sub cmdSaveOnHD_Click()
    Dim strRes As String
    strRes = SaveHPmanaConfig()
    If strRes = "" Then
         lblMsg.Caption = "Saved sucesfully!"
         lblMsg.ForeColor = &HFF00&
    Else
         lblMsg.Caption = strRes
         lblMsg.ForeColor = &HFF&
    End If
End Sub

Private Sub cmdSaveOnMemory_Click()
    Dim strSel As String
    Dim strMANAact As String
    Dim strHPact As String
    Dim bMANA As Byte
    Dim bHP As Byte
    Dim lineSel As Long
    lineSel = Me.gridHPmana.RowSel
    If lineSel < 1 Then
        lblMsg.Caption = "First select the setting that you want to overwrite!"
        lblMsg.ForeColor = &HFFFF&
        Exit Sub
    End If
    
    strSel = LCase(cmbCharacter.Text)
    If strSel = "" Or strSel = "-" Then
        ' do nothing
        lblMsg.Caption = "Select/type a character first!"
        lblMsg.ForeColor = &HFFFF&
    Else
        bHP = CByte(scrollHP.Value)
        bMANA = CByte(scrollMANA.Value)
        If HPopt1.Value = True Then
            strHPact = "no action"
        ElseIf HPopt2.Value = True Then
            strHPact = "IH"
        ElseIf HPopt3.Value = True Then
            strHPact = "UH"
        ElseIf HPopt4.Value = True Then
            strHPact = "life fluid"
        ElseIf HPopt5.Value = True Then
            strHPact = "health potion"
        ElseIf HPopt6.Value = True Then
            strHPact = "strong health potion"
        ElseIf HPopt7.Value = True Then
            strHPact = "great health potion"
        ElseIf HPopt9.Value = True Then
            strHPact = "ultimate health potion"
        ElseIf HPopt10.Value = True Then
            strHPact = "great spirit potion"
        ElseIf HPopt11.Value = True Then
            strHPact = "small health potion"
        ElseIf HPopt8.Value = True Then
            strHPact = hpTEXT.Text
        End If
        
        If MANAopt1.Value = True Then
            strMANAact = "no action"
        ElseIf MANAopt2.Value = True Then
            strMANAact = "mana fluid"
        ElseIf MANAopt3.Value = True Then
            strMANAact = "mana potion"
        ElseIf MANAopt4.Value = True Then
            strMANAact = "strong mana potion"
        ElseIf MANAopt5.Value = True Then
            strMANAact = "great mana potion"
        ElseIf MANAopt7.Value = True Then
            strMANAact = "great spirit potion"
        ElseIf MANAopt6.Value = True Then
            strMANAact = manaTEXT.Text
        End If
        
        UpdateHPmanaSetting strSel, bHP, strHPact, bMANA, strMANAact, lineSel
        DisplayLoadedHPmanaConfig
        lblMsg.Caption = "Character <" & strSel & "> saved sucesfully!"
        lblMsg.ForeColor = &HFF00&
    End If
End Sub

Private Sub cmdSaveRecast_Click()
    On Error GoTo goterr
    Dim lngCast As Long
    lngCast = CLng(Me.txtRecast.Text)
    If ((lngCast >= 20) And (lngCast <= HPmanaRECAST2)) Then
        HPmanaRECAST = lngCast
        Me.txtRecast.Text = CStr(HPmanaRECAST)
        cmdSaveRecast.Caption = "CHANGE ( current = " & CStr(HPmanaRECAST) & " ms )"
    Else
        GoTo goterr
    End If
    lblMsg.Caption = "Changed recast succesfully"
    lblMsg.ForeColor = &HFF00&
    Exit Sub
goterr:
    lblMsg.Caption = "Invalid setting"
    lblMsg.ForeColor = &HFF&
End Sub



Private Sub cmdSaveRecast2_Click()
    On Error GoTo goterr
    Dim lngCast As Long
    lngCast = CLng(Me.txtRecast2.Text)
    If ((lngCast >= 20) And (lngCast >= HPmanaRECAST)) Then
        HPmanaRECAST2 = lngCast
        Me.txtRecast2.Text = CStr(HPmanaRECAST2)
        cmdSaveRecast2.Caption = "CHANGE ( current = " & CStr(HPmanaRECAST2) & " ms )"
    Else
        GoTo goterr
    End If
    lblMsg.Caption = "Changed recast succesfully"
    lblMsg.ForeColor = &HFF00&
    Exit Sub
goterr:
    lblMsg.Caption = "Invalid setting"
    lblMsg.ForeColor = &HFF&
End Sub


Private Sub Form_Load()
    Dim strRes As String
  With gridHPmana
  .ColWidth(0) = 200
  .ColWidth(1) = 1800
  .ColWidth(2) = 760
  .ColWidth(3) = 1800
  .ColWidth(4) = 760
  .ColWidth(5) = 1900
  .TextMatrix(0, 0) = ">"
  .TextMatrix(0, 1) = "Char"
  .TextMatrix(0, 2) = "%HP"
  .TextMatrix(0, 3) = "Action"
  .TextMatrix(0, 4) = "%MANA"
  .TextMatrix(0, 5) = "Action"
  .Row = 0
  .Col = 0
  .CellAlignment = flexAlignLeftCenter
  .Col = 1
  .CellAlignment = flexAlignLeftCenter
  .Col = 2
  .CellAlignment = flexAlignLeftCenter
  .Col = 3
  .CellAlignment = flexAlignLeftCenter
  .Col = 4
  .CellAlignment = flexAlignLeftCenter
  .Col = 5
  .CellAlignment = flexAlignLeftCenter
  End With
  strRes = LoadHPmanaConfig()
  If strRes = "" Then
    lblMsg.Caption = "Settings recovered from HD succesfully"
    lblMsg.ForeColor = &HFF00&
  Else
    lblMsg.Caption = strRes
    lblMsg.ForeColor = &HFF&
  End If
  Me.txtRecast.Text = CStr(HPmanaRECAST)
  Me.txtRecast2.Text = CStr(HPmanaRECAST2)
  Me.txtLimitRandomizator.Text = CStr(LimitRandomizator)
  cmdSaveRecast.Caption = "CHANGE ( current = " & CStr(HPmanaRECAST) & " ms )"
  cmdSaveRecast2.Caption = "CHANGE ( current = " & CStr(HPmanaRECAST2) & " ms )"
  cmdSaveLimitRandomizator.Caption = "CHANGE ( current =" & CStr(LimitRandomizator) & " % )"
  tmrHPmana.Interval = 100
  tmrHPmana.enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub



Private Sub scrollHP_Change()
    lblHPvalue.Caption = scrollHP.Value & " %"
End Sub

Private Sub scrollMANA_Change()
    lblMANAvalue.Caption = scrollMANA.Value & " %"
End Sub

Public Sub LoadHPmanaChars()
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
End Sub

Private Sub OLDTIMERMETHODHERE()

    Dim i As Integer
    Dim resA As Long
    Dim act As String
    Dim ult As Long
    Dim idset As Long
    Dim gtc As Long
    Dim hpPercent As Long
    Dim manaPercent As Long
    Dim allowed As Boolean
    Dim usedHeal As Boolean
    Dim idConnection As Integer
    gtc = GetTickCount()
    
    ult = UBound(HPmanaConfig)
    For i = 1 To MAXCLIENTS
        idConnection = i
        'myMana(idConnection) = 0 ' DEBUGGG! delete later
        allowed = True
        If (GameConnected(i) = True) And (sentWelcome(i) = True) And (GotPacketWarning(i) = False) Then
            If gtc >= LastHealTime(i) Then
                hpPercent = 100 * ((myHP(i) / myMaxHP(i)))
                If myMaxMana(i) = 0 Then
                    manaPercent = 0
                Else
                    manaPercent = 100 * ((myMana(i) / myMaxMana(i)))
                End If
                usedHeal = False
                If NoHealingNextTurn(idConnection) = True Then
                    NoHealingNextTurn(idConnection) = False
                Else
                For idset = 1 To ult
                    If LCase(HPmanaConfig(idset).charName) = LCase(CharacterName(i)) Then
                        If ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
                            If (hpPercent < HPmanaConfig(idset).hpVal) Then
                                act = LCase(HPmanaConfig(idset).hpACTION)
                                resA = 0
                                Select Case act
                                Case "uh"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseUH(i)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)
       
                                Case "ih"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseIH(i)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "life fluid"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseFluid(i, byteLife)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "strong health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_strong_health_potion)
                                    usedHeal = True
                                    
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)
    
                                Case "great health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_great_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "small health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_small_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "ultimate health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_ultimate_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "great spirit potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_great_spirit_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "no action"
                                    ' nothing
                                Case Else
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    If allowed = True Then
                                        HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                        allowed = False
                                        resA = ExecuteInTibia(HPmanaConfig(idset).hpACTION, i, True)
                                    End If
                                    resA = 0
                                    usedHeal = True
                                End Select
                                If (resA = -1) Then
                                    If (PlayTheDangerSound = False) Then
                                        'give msg !
                                        resA = GiveGMmessage(i, "Unable to recharge HP with the selected method!", "Warning")
                                        ChangePlayTheDangerSound True
                                        DoEvents
                                    End If
                                End If
                            End If
                        End If ' // if cheats paused
                    End If
                   
                Next idset
                End If

 
                ' mana
                For idset = 1 To ult
                    If LCase(HPmanaConfig(idset).charName) = LCase(CharacterName(i)) Then
                        If ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
                            If (manaPercent < HPmanaConfig(idset).manaVal) Then
                                act = LCase(HPmanaConfig(idset).manaACTION)
                                resA = 0
                                Select Case act
                                Case "mana fluid"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UseFluid(i, byteMana)
                                    End If
                                Case "mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_mana_potion)
                                    End If
                                Case "strong mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_strong_mana_potion)
                                    End If
                                Case "great mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_great_mana_potion)
                                    End If
                                Case "great spirit potion"
                                    If (usedHeal = True) Then
                                    
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_great_spirit_potion)
                                    End If
                                Case "no action"
                                    ' nothing
                                Case Else
                                
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        If allowed = True Then
                                            HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                            allowed = False
                                            resA = ExecuteInTibia(HPmanaConfig(idset).manaACTION, i, True)
                                        
                                        End If
                                        resA = 0
                                    End If
                                End Select
                                If (resA = -1) Then
                                    If (PlayTheDangerSound = False) Then
                                        'give msg !
                                        resA = GiveGMmessage(i, "Unable to recharge MANA with the selected method!", "Warning")
                                        ChangePlayTheDangerSound True
                                        DoEvents
                                    End If
                                End If
                            End If
                        End If ' // if cheats paused
                    End If ' // if char name
                Next idset
            End If
        End If
    Next i

End Sub


Private Sub tmrHPmana_Timer()

    Dim i As Integer
    Dim resA As Long
    Dim act As String
    Dim ult As Long
    Dim idset As Long
    Dim gtc As Long
    Dim hpPercent As Long
    Dim manaPercent As Long
    Dim allowed As Boolean
    Dim usedHeal As Boolean
    Dim idConnection As Integer
    Dim iSmart As Long
    Dim lowestSmart As Long
    Dim winnerSmart As Long
    gtc = GetTickCount()
    
    ult = UBound(HPmanaConfig)
    For i = 1 To MAXCLIENTS
        idConnection = i
        'myMana(idConnection) = 0 ' DEBUGGG! delete later
        allowed = True
        If (GameConnected(i) = True) And (sentWelcome(i) = True) And (GotPacketWarning(i) = False) Then
            If gtc >= LastHealTime(i) Then
                hpPercent = 100 * ((myHP(i) / myMaxHP(i)))
                If myMaxMana(i) = 0 Then
                    manaPercent = 0
                Else
                    manaPercent = 100 * ((myMana(i) / myMaxMana(i)))
                End If
                usedHeal = False
                If NoHealingNextTurn(idConnection) = True Then
                    NoHealingNextTurn(idConnection) = False
                Else
                lowestSmart = 200
                winnerSmart = 0
                For iSmart = 1 To ult
                    If LCase(HPmanaConfig(iSmart).charName) = LCase(CharacterName(i)) Then
                        If ((HPmanaConfig(iSmart).hpVal <= lowestSmart) And ((hpPercent <= HPmanaConfig(iSmart).hpVal))) Then
                            winnerSmart = iSmart
                            lowestSmart = HPmanaConfig(iSmart).hpVal
                        End If
                    End If
                Next iSmart
                idset = winnerSmart
                If idset > 0 Then
                    If LCase(HPmanaConfig(idset).charName) = LCase(CharacterName(i)) Then
                        If ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
                            If (hpPercent < HPmanaConfig(idset).hpVal) Then
                                act = LCase(HPmanaConfig(idset).hpACTION)
                                resA = 0
                                Select Case act
                                Case "uh"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseUH(i)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)
       
                                Case "ih"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseIH(i)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "life fluid"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UseFluid(i, byteLife)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "strong health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_strong_health_potion)
                                    usedHeal = True
                                    
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)
    
                                Case "great health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_great_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "small health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_small_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "ultimate health potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_ultimate_health_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "great spirit potion"
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    resA = UsePotion(i, tileID_great_spirit_potion)
                                    usedHeal = True
                                    HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                Case "no action"
                                    ' nothing
                                Case Else
                                    LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                    If allowed = True Then
                                        HPmanaConfig(idset).hpVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseHpVal)

                                        allowed = False
                                        resA = ExecuteInTibia(HPmanaConfig(idset).hpACTION, i, True)
                                    End If
                                    resA = 0
                                    usedHeal = True
                                End Select
                                If (resA = -1) Then
                                    If (PlayTheDangerSound = False) Then
                                        'give msg !
                                        resA = GiveGMmessage(i, "Unable to recharge HP with the selected method!", "Warning")
                                        ChangePlayTheDangerSound True
                                        DoEvents
                                    End If
                                End If
                            End If
                        End If ' // if cheats paused
                    End If
                   
                  End If ' idset >0
                End If

 
                ' mana
                lowestSmart = 200
                winnerSmart = 0
                For iSmart = 1 To ult
                    If LCase(HPmanaConfig(iSmart).charName) = LCase(CharacterName(i)) Then
                        If ((HPmanaConfig(iSmart).manaVal <= lowestSmart) And ((manaPercent <= HPmanaConfig(iSmart).manaVal))) Then
                            winnerSmart = iSmart
                            lowestSmart = HPmanaConfig(iSmart).manaVal
                        End If
                    End If
                Next iSmart
                idset = winnerSmart
                If idset > 0 Then
                    If LCase(HPmanaConfig(idset).charName) = LCase(CharacterName(i)) Then
                        If ((CheatsPaused(i) = False) Or (AllowUHpaused(i) = True)) Then
                            If (manaPercent < HPmanaConfig(idset).manaVal) Then
                                act = LCase(HPmanaConfig(idset).manaACTION)
                                resA = 0
                                Select Case act
                                Case "mana fluid"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UseFluid(i, byteMana)
                                    End If
                                Case "mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_mana_potion)
                                    End If
                                Case "strong mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_strong_mana_potion)
                                    End If
                                Case "great mana potion"
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_great_mana_potion)
                                    End If
                                Case "great spirit potion"
                                    If (usedHeal = True) Then
                                    
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        resA = UsePotion(i, tileID_great_spirit_potion)
                                    End If
                                Case "no action"
                                    ' nothing
                                Case Else
                                
                                    If (usedHeal = True) Then
                                        NoHealingNextTurn(idConnection) = True
                                    Else
                                        LastHealTime(i) = gtc + randomNumberBetween(HPmanaRECAST, HPmanaRECAST2)
                                        If allowed = True Then
                                            HPmanaConfig(idset).manaVal = ChaotizeRechargeLevel(HPmanaConfig(idset).baseManaVal)

                                            allowed = False
                                            resA = ExecuteInTibia(HPmanaConfig(idset).manaACTION, i, True)
                                        
                                        End If
                                        resA = 0
                                    End If
                                End Select
                                If (resA = -1) Then
                                    If (PlayTheDangerSound = False) Then
                                        'give msg !
                                        resA = GiveGMmessage(i, "Unable to recharge MANA with the selected method!", "Warning")
                                        ChangePlayTheDangerSound True
                                        DoEvents
                                    End If
                                End If
                            End If
                        End If ' // if cheats paused
                    End If ' // if char name
               End If ' idset >0
            End If
        End If
    Next i

End Sub
