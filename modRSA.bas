Attribute VB_Name = "modRSA"
#Const FinalMode = 1
Option Explicit

Public adrRSA As Long
'Public Const RLserverRSAkey As String = "124710459426827943004376449897985582167801707960697037164044904862948569380850421396904597686953877022394604239428185498284169068581802277612081027966724336319448537811441719076484340922854929273517308661370727105382899118999403808045846444647284499123164879035103627004668521005328367415259939915284902061793"
Public Const RLserverRSAkey1075 As String = "132127743205872284062295099082293384952776326496165507967876361843343953435544496682053323833394351797728954155097012103928360786959821132214473291575712138800495033169914814069637740318278150290733684032524174782740134357629699062987023311132821016569775488792221429527047321331896351555606801473202394175817"

Public Const OTserverRSAkey As String = "109120132967399429278860960508995541528237502902798129123468757937266291492576446330739696001110603907230888610072655818825358503429057592827629436413108566029093628212635953836686562675849720620786279431090218017681061521755056710823876476444260558147179707119674283982419152118103759076030616683978566631413"

Public WARNING_USING_OTSERVER_RSA As Boolean

Public Sub AutoUpdateRSA(ByVal pid As Long)
  On Error GoTo goterr
  Dim pg As Integer
  Dim i As Long
  Dim b As Byte
  Dim sb As String
  Dim s As String
  Dim si As Integer
 ' Dim sc As String
  Dim maxsi As Integer
  Dim backupi As Long
  Dim reskey As String
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "Trying to autoupdate adrRSA..."
  reskey = ""
  pg = 0
  maxsi = 1
  si = 1
  'sc = Mid$(RLserverRSAkey, si, 1)
  sb = ""
  i = &H500000
  Do
     b = Memory_ReadByte(i, pid)
     sb = Chr(b)
     If (IsNumeric(sb)) Then
    ' If (sb = sc) Then
    reskey = reskey & sb
       si = si + 1
       If (si = 2) Then
         backupi = i
       End If
       If (si > maxsi) Then
         maxsi = si
         'Debug.Print ("New record (" & maxsi - 1 & ") at &H" & Hex(backupi)) & " : " & reskey
         If maxsi - 1 = 309 Then
           adrRSA = backupi
           frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "SUCCESS!! Found RSA key at &H" & (Hex(adrRSA)) & " : " & reskey
           Exit Sub
         End If
       End If
     Else
       reskey = ""
       If (si > 1) Then
         i = backupi
         si = 1
       End If
     End If
    ' sc = Mid$(RLserverRSAkey, si, 1)
     pg = pg + 1
     If (pg >= 10000) Then
       frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & Hex(i) & " Searching RSA key for this tibia client..."
       pg = 0
     End If
     i = i + 1
     DoEvents
  Loop Until i >= &HA00000
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... MEMORY SCAN COMPLETED WITHOUT RESULTS"
   Exit Sub
   
goterr:
  adrRSA = 0
   frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & "FAIL ... Error at AutoUpdateRSA (" & CStr(Err.Number) & ") : " & Err.Description
  Exit Sub
End Sub
Public Sub TryToUpdateRSA(ByVal pid As Long, ByVal strKey As String, Optional fixRSA As Boolean = False)
' at this moment fixRSA =true only will works at Windows XP (or in any Window if ASLR is disabled)
    Dim i As Long
    Dim writeChr As String
    Dim currByteAdr As Long
    Dim byteChr As Byte
    Dim byteChrR As Byte
    Dim RSA_bytes(308) As Byte
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    Dim res As Long
    Dim realAddress As Long
    'fixRSA = True
    
    If fixRSA = True Then
      If adrRSA = 0 Then
        AutoUpdateRSA (pid)
      End If
    End If

    If adrRSA = 0 Then
        Exit Sub
    End If
    If pid = -1 Then
        Exit Sub
    End If
   
    realAddress = Memory_BlackdAddressToFinalAdddress(adrRSA, pid)
    If (realAddress = 0) Then
      Exit Sub
    End If
   
    For i = 0 To 308
      writeChr = Mid$(strKey, i + 1, 1)
      byteChr = ConvStrToByte(writeChr)
      RSA_bytes(i) = byteChr
    Next i
   
    res = BlackdForceWrite(realAddress, RSA_bytes(0), 309, pid)
    Debug.Print "RSA key changed"

    WARNING_USING_OTSERVER_RSA = True
End Sub

Public Sub ModifyTibiaRSAs()
  Dim tibiaclient As Long
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
        Exit Do
    Else
        TryToUpdateRSA tibiaclient, OTserverRSAkey
    End If
  Loop
End Sub
