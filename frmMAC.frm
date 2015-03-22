VERSION 5.00
Begin VB.Form frmMAC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd MAC"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   Icon            =   "frmMAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Your MAC address is:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32

Private Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte 'Reserved, must be 0
   ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Private Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
   (pncb As NET_CONTROL_BLOCK) As Byte
     
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, ByVal _
   hpvSource As Long, ByVal _
   cbCopy As Long)
     
Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
  (ByVal hHeap As Long, _
   ByVal dwFlags As Long, _
   lpMem As Any) As Long



Private Sub Command1_Click()
  Unload Me
End Sub


Private Function GetMACAddress(sDelimiter As String) As String

  'retrieve the MAC Address for the network controller
  'installed, returning a formatted string
   
   Dim tmp As String
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT
   Dim cnt As Long

  'The IBM NetBIOS 3.0 specifications defines four basic
  'NetBIOS environments under the NCBRESET command. Win32
  'follows the OS/2 Dynamic Link Routine (DLR) environment.
  'This means that the first NCB issued by an application
  'must be a NCBRESET, with the exception of NCBENUM.
  'The Windows NT implementation differs from the IBM
  'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
   
  'To get the Media Access Control (MAC) address for an
  'ethernet adapter programmatically, use the Netbios()
  'NCBASTAT command and provide a "*" as the name in the
  'NCB.ncb_CallName field (in a 16-chr string).
   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT
   
  'For machines with multiple network adapters you need to
  'enumerate the LANA numbers and perform the NCBASTAT
  'command on each. Even when you have a single network
  'adapter, it is a good idea to enumerate valid LANA numbers
  'first and perform the NCBASTAT on one of the valid LANA
  'numbers. It is considered bad programming to hardcode the
  'LANA number to 0 (see the comments section below).
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   
   pASTAT = HeapAlloc(GetProcessHeap(), _
                      HEAP_GENERATE_EXCEPTIONS Or _
                      HEAP_ZERO_MEMORY, _
                      NCB.ncb_length)
            
   If pASTAT <> 0 Then
   
      NCB.ncb_buffer = pASTAT
      Call Netbios(NCB)
      
      CopyMemory AST, NCB.ncb_buffer, Len(AST)
              
     'convert the byte array to a string
      GetMACAddress = MakeMacAddress(AST.adapt.adapter_address(), sDelimiter)
                          
      HeapFree GetProcessHeap(), 0, pASTAT

    Else
      GetMACAddress = "NO VALID MAC"
      Exit Function
   End If
   
End Function


Private Function MakeMacAddress(b() As Byte, sDelim As String) As String

   Dim cnt As Long
   Dim buff As String
   
   On Local Error GoTo MakeMac_error
 
  'so far, MAC addresses are
  'exactly 6 segments in size (0-5)
   If UBound(b) = 5 Then
   
     'concatenate the first five values
     'together and separate with the
     'delimiter char
      For cnt = 0 To 4
         buff = buff & Right$("00" & Hex(b(cnt)), 2) & sDelim
      Next
      
     'and append the last value
      buff = buff & Right$("00" & Hex(b(5)), 2)
         
   End If  'UBound(b)
   
   MakeMacAddress = buff
   
MakeMac_exit:
   Exit Function
   
MakeMac_error:
   MakeMacAddress = "(error building MAC address)"
   Resume MakeMac_exit
   
End Function



Private Sub Form_Load()
   Text1.Text = GetMACAddress(":")
   Text1.Locked = True
End Sub
