VERSION 5.00
Begin VB.Form frmScreenshot 
   BackColor       =   &H00000000&
   Caption         =   "Taking screenshot ..."
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4185
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmScreenshot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picScreen 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   2880
         ScaleHeight     =   15
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   960
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmScreenshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub Form_Load()
  Me.WindowState = 1
  Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub
