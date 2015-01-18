Attribute VB_Name = "modRSA"
#Const FinalMode = 1
Option Explicit

Public adrRSA As Long
' Public Const RLserverRSAkey As String = "124710459426827943004376449897985582167801707960697037164044904862948569380850421396904597686953877022394604239428185498284169068581802277612081027966724336319448537811441719076484340922854929273517308661370727105382899118999403808045846444647284499123164879035103627004668521005328367415259939915284902061793"

Public Const OTserverRSAkey As String = "109120132967399429278860960508995541528237502902798129123468757937266291492576446330739696001110603907230888610072655818825358503429057592827629436413108566029093628212635953836686562675849720620786279431090218017681061521755056710823876476444260558147179707119674283982419152118103759076030616683978566631413"

Public WARNING_USING_OTSERVER_RSA As Boolean

Public Sub TryToUpdateRSA(ByVal pid As Long, ByVal strKey As String)
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
    ' needs to be fixed

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
