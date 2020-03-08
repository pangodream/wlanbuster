Attribute VB_Name = "Prot_Composer"
Public Enum Eth_Protocol
     IP_Prot = &H800
     ARP_Prot = &H806
End Enum
Public Enum IP_Protocol
     ICMP_Prot = &H1
     IGMP_Prot = &H2
     IP_over_IP_Prot = &H4
     TCP_Prot = &H6
     UDP_Prot = &H11
End Enum
Public Enum ARP_Operation
     Request = &H1
     Reply = &H2
End Enum
Public Enum PCK_TYPE
    Management = 0
    Control = 1
    data = 2
End Enum
Public Enum PCK_SUBTYPE
    mgmt_AssocRequest = 0
    mgmt_AssocResponse = 1
    mgmt_ReAssocRequest = 2
    mgmt_ReAssocResponse = 3
    mgmt_ProbeRequest = 4
    mgmt_ProbeResponse = 5
    mgmt_Beacon = 8
    mgmt_ATIM = 9
    mgmt_Disassociation = 10
    mgmt_Authentication = 11
    mgmt_Deauthentication = 12
    ctrl_RTS = 11
    ctrl_CTS = 12
    ctrl_ACK = 13
    ctrl_CFEnd = 14
    ctrl_DFEnd_CFACK = 15
    data_NullData = 4
    data_Data = 0
End Enum

Public Type EthernetII_STR
     a00_dMAC As String
     a01_sMAC As String
     a02_Protocol As String
     data As String
End Type
Public EthernetTest As EthernetII_STR

Public Type ARP_STR
     a00_HW_Type As String
     a01_Prot_Type As String
     a02_HW_Size As String
     a03_Prot_Size As String
     a04_Operation As String
     a05_SourceMAC As String
     a06_SourceIP As String
     a07_DestinMAC As String
     a08_DestinIP As String
End Type
Public ARPTest As ARP_STR

Public Type IP_STR
     a00_Version_4 As String
     a01_Diff_Services As String
     a02_Length As String
     a03_Identification As String
     a04_Flags As String
     a05_FragmentOffset As String
     a06_TTL As String
     a07_Protocol As String
     a08_CheckSum As String
     a09_SourceIP As String
     a10_DestinIP As String
     data As String
End Type
Public IPTest As IP_STR

Public Type UDP_STR
     x00_SourceIP As String
     x01_DestinIP As String
     a00_SourcePort As String
     a01_DestinPort As String
     a02_Length As String
     a08_CheckSum As String
     data As String
End Type
Public UDPTest As UDP_STR

Public Type TCPFlags_STR
     f0_Fin As Boolean
     f1_SYN As Boolean
     f2_Reset As Boolean
     f3_Push As Boolean
     f4_ACK As Boolean
     f5_Urgent As Boolean
     f6_Echo As Boolean
     f7_CWR As Boolean
End Type
Public Type TCP_STR
     x00_SourceIP As String
     x01_DestinIP As String
     a00_SourcePort As String
     a01_DestinPort As String
     a02_RelSeqNum As String
     a03_RelAckNum As String
     a04_Length As String
     a05_Flags As TCPFlags_STR
     a05_FlagsByte As String
     a06_WindowsSize As String
     a07_CheckSum As String
     o01_Options As String
     data As String
End Type
Public TCPTest As TCP_STR

Public Type LLC_STR
     a01_DSAP As String
     a02_SSAP As String
     a03_Command As String
     a04_Protocol As String
End Type
Public LLCTest As LLC_STR


Public Type TPE_FrameControl
    Version As Byte '(2 bits)
    Type As Byte    '(2 bits)
    Subtype As Byte '(4 bits)
    ToDS As Boolean
    FromDS As Boolean
    MoreFragment As Boolean
    Retry As Boolean
    Power As Boolean
    More As Boolean
    WEP As Boolean
    Order As Boolean
End Type

Public Type WF80211_STR
    a01_FrameControl As TPE_FrameControl
    a01_FrameControlWord As String
    a02_Duration As String
    a03_BSSID As String
    a04_SourceMAC As String
    a05_DestinationMAC As String
    a06_FragmentNumber As String
    a07_SequenceNumber As String
End Type
Public WF80211TEst As WF80211_STR

Public Function CompileWF(ByRef pckWF As WF80211_STR) As String
    'Procesa un paquete 802.11 y devuelve cadena hexa
    pckWF.a01_FrameControlWord = CompileFrameControl(pckWF.a01_FrameControl)
    If pckWF.a02_Duration = "" Then pckWF.a02_Duration = "3000"
    If pckWF.a03_BSSID = "" Then MsgBox "CompileWF: Falta BSSID"
    If pckWF.a04_SourceMAC = "" Then MsgBox "CompileWF: Falta Source MAC"
    If pckWF.a05_DestinationMAC = "" Then MsgBox "CompileWF: Falta Source MAC"
''''    If pckWF.a06_FragmentNumber = "" Then pckWF.a06_FragmentNumber = "0000"
    If pckWF.a07_SequenceNumber = "" Then pckWF.a07_SequenceNumber = "F031" 'Cualquiera???
    
    CompileWF = CompileWF & pckWF.a01_FrameControlWord
    CompileWF = CompileWF & pckWF.a02_Duration
    CompileWF = CompileWF & pckWF.a03_BSSID
    CompileWF = CompileWF & pckWF.a04_SourceMAC
    CompileWF = CompileWF & pckWF.a05_DestinationMAC
    CompileWF = CompileWF & pckWF.a06_FragmentNumber
    CompileWF = CompileWF & pckWF.a07_SequenceNumber
End Function


Public Function CompileLLC(ByRef pckLLC As LLC_STR) As String
     'Procesa un paquete LLC y devuelve cadena hexa
     If pckLLC.a01_DSAP = "" Then pckLLC.a01_DSAP = "AA"
     If pckLLC.a02_SSAP = "" Then pckLLC.a02_SSAP = "AA"
     If pckLLC.a03_Command = "" Then pckLLC.a03_Command = "03"
     CompileLLC = CompileLLC & pckLLC.a01_DSAP
     CompileLLC = CompileLLC & pckLLC.a02_SSAP
     CompileLLC = CompileLLC & pckLLC.a03_Command
     CompileLLC = CompileLLC & "000000"
     CompileLLC = CompileLLC & pckLLC.a04_Protocol
End Function

Public Function CompileEthernetII(ByRef pckETH As EthernetII_STR) As String
     'Procesa un paquete Ethernet y devuelve una cadena hexa
     CompileEthernetII = CompileEthernetII & pckETH.a00_dMAC
     CompileEthernetII = CompileEthernetII & pckETH.a01_sMAC
     CompileEthernetII = CompileEthernetII & pckETH.a02_Protocol
     CompileEthernetII = CompileEthernetII & pckETH.data
End Function

Public Function CompileARP(ByRef pckARP As ARP_STR) As String
     'Procesa un paquete arp y devuelve una cadena hexa
     pckARP.a00_HW_Type = "0001" 'Ethernet
     pckARP.a01_Prot_Type = "0800" 'IP
     pckARP.a02_HW_Size = "06"
     pckARP.a03_Prot_Size = "04"
     If pckARP.a04_Operation = "" Then MsgBox "CompileARP: Falta Operation"
     If pckARP.a05_SourceMAC = "" Then MsgBox "CompileARP: Falta Source MAC"
     If pckARP.a06_SourceIP = "" Then MsgBox "CompileARP: Falta Source IP"
     If pckARP.a07_DestinMAC = "" Then MsgBox "CompileARP: Falta Destination MAC"
     If pckARP.a08_DestinIP = "" Then MsgBox "CompileARP: Falta Destination IP"
    
     CompileARP = CompileARP & pckARP.a00_HW_Type
     CompileARP = CompileARP & pckARP.a01_Prot_Type
     CompileARP = CompileARP & pckARP.a02_HW_Size
     CompileARP = CompileARP & pckARP.a03_Prot_Size
     CompileARP = CompileARP & pckARP.a04_Operation
     CompileARP = CompileARP & pckARP.a05_SourceMAC
     CompileARP = CompileARP & pckARP.a06_SourceIP
     CompileARP = CompileARP & pckARP.a07_DestinMAC
     CompileARP = CompileARP & pckARP.a08_DestinIP
End Function
Public Function CompileIP(ByRef pckIP As IP_STR) As String
     'Procesa un paquete IP y devuelve una cadena hexa
     pckIP.a00_Version_4 = "45"
     pckIP.a01_Diff_Services = "00"
     pckIP.a02_Length = WHexa(Len(pckIP.data) / 2 + 20)
     If pckIP.a03_Identification = "" Then pckIP.a03_Identification = "0102"
     If pckIP.a04_Flags = "" Then pckIP.a04_Flags = "40" 'Don't fragment
     pckIP.a05_FragmentOffset = "00"
     pckIP.a06_TTL = "80"
     If pckIP.a07_Protocol = "" Then MsgBox "CompileIP: Falta Protocol"
     pckIP.a08_CheckSum = IPCheckSum(pckIP)
     If pckIP.a09_SourceIP = "" Then MsgBox "CompileIP: Falta Source IP"
     If pckIP.a10_DestinIP = "" Then MsgBox "CompileIP: Falta Destination IP"
     'pckIP.Data
    
     CompileIP = CompileIP & pckIP.a00_Version_4
     CompileIP = CompileIP & pckIP.a01_Diff_Services
     CompileIP = CompileIP & pckIP.a02_Length
     CompileIP = CompileIP & pckIP.a03_Identification
     CompileIP = CompileIP & pckIP.a04_Flags
     CompileIP = CompileIP & pckIP.a05_FragmentOffset
     CompileIP = CompileIP & pckIP.a06_TTL
     CompileIP = CompileIP & pckIP.a07_Protocol
     CompileIP = CompileIP & pckIP.a08_CheckSum
     CompileIP = CompileIP & pckIP.a09_SourceIP
     CompileIP = CompileIP & pckIP.a10_DestinIP
     CompileIP = CompileIP & pckIP.data
End Function
Public Function CompileTCP(ByRef pckTCP As TCP_STR) As String
     'Procesa un paquete TCP y devuelve una cadena hexa
     Dim tf As Integer
     If pckTCP.a00_SourcePort = "" Then MsgBox "CompileTCP: Falta Source Port"
     If pckTCP.a01_DestinPort = "" Then MsgBox "CompileTCP: Falta Destination Port"
     If pckTCP.a02_RelSeqNum = "" Then pckTCP.a02_RelSeqNum = "00000000"
     If pckTCP.a03_RelAckNum = "" Then pckTCP.a03_RelAckNum = "00000000"
     'Calculo de longitud:
        ' En el nibble de la izquierda ponemos el número de bloques de 4 bytes (32 bits) de la longitud de la cabecera
        ' Así, si mide 20 son 5 paquetes de 4 bytes, es decir 50
        '      si mide 28 son 7 paquetes de 4 bytes, es decir 70
     If pckTCP.a05_Flags.f1_SYN = True Then
        pckTCP.a04_Length = "70"
     Else
        pckTCP.a04_Length = "50"
     End If
    
     'Cálculo de flags
     tf = tf + IIf(pckTCP.a05_Flags.f7_CWR, 128, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f6_Echo, 64, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f5_Urgent, 32, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f4_ACK, 16, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f3_Push, 8, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f2_Reset, 4, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f1_SYN, 2, 0)
     tf = tf + IIf(pckTCP.a05_Flags.f0_Fin, 1, 0)
     pckTCP.a05_FlagsByte = BHexa(tf)
    
     If pckTCP.a06_WindowsSize = "" Then pckTCP.a06_WindowsSize = "44E8"
     'pckTCP.data
     If pckTCP.a05_Flags.f1_SYN Then
        If pckTCP.o01_Options = "" Then pckTCP.o01_Options = "020404EC01010402"
     Else
        pckTCP.o01_Options = ""
     End If
     If pckTCP.x00_SourceIP = "" Then MsgBox "CompileTCP: Falta Source IP"
     If pckTCP.x01_DestinIP = "" Then MsgBox "CompileTCP: Falta Destination IP"
     pckTCP.a07_CheckSum = TCPCheckSum(pckTCP)
    
     CompileTCP = CompileTCP & pckTCP.a00_SourcePort
     CompileTCP = CompileTCP & pckTCP.a01_DestinPort
     CompileTCP = CompileTCP & pckTCP.a02_RelSeqNum
     CompileTCP = CompileTCP & pckTCP.a03_RelAckNum
     CompileTCP = CompileTCP & pckTCP.a04_Length
     CompileTCP = CompileTCP & pckTCP.a05_FlagsByte
     CompileTCP = CompileTCP & pckTCP.a06_WindowsSize
     CompileTCP = CompileTCP & pckTCP.a07_CheckSum & "0000"
     CompileTCP = CompileTCP & pckTCP.o01_Options
     CompileTCP = CompileTCP & pckTCP.data
End Function

Public Function CompileUDP(ByRef pckUDP As UDP_STR) As String
     'Procesa un paquete UDP y devuelve una cadena hexa
     If pckUDP.a00_SourcePort = "" Then MsgBox "CompileUDP: Falta Source Port"
     If pckUDP.a01_DestinPort = "" Then MsgBox "CompileUDP: Falta Destination Port"
     pckUDP.a02_Length = WHexa(Len(pckUDP.data) / 2 + 8)
     pckUDP.a08_CheckSum = UDPCheckSum(pckUDP)
     'pckUDP.Data
     If pckUDP.x00_SourceIP = "" Then MsgBox "CompileUDP: Falta Source IP"
     If pckUDP.x01_DestinIP = "" Then MsgBox "CompileUDP: Falta Destination IP"
    
     CompileUDP = CompileUDP & pckUDP.a00_SourcePort
     CompileUDP = CompileUDP & pckUDP.a01_DestinPort
     CompileUDP = CompileUDP & pckUDP.a02_Length
     CompileUDP = CompileUDP & pckUDP.a08_CheckSum
     CompileUDP = CompileUDP & pckUDP.data
End Function


Public Function IPCheckSum(ByRef pckIP As IP_STR) As String
   Dim dl As Long
   Dim suma As Long
   Dim crr As Long
   Dim b1 As Long
   Dim b2 As Long
   Dim tr As String
   Dim data As String
   dl = Len(pckIP.data) / 2
   If dl Mod 2 = 1 Then tr = "00"
  
   data = pckIP.a00_Version_4 & pckIP.a01_Diff_Services & pckIP.a02_Length & pckIP.a03_Identification
   data = data & pckIP.a04_Flags & pckIP.a05_FragmentOffset & pckIP.a06_TTL & pckIP.a07_Protocol & "0000"
   data = data & pckIP.a09_SourceIP & pckIP.a10_DestinIP
  
   data = HexaASCII(data)
   For n = 1 To Len(data) - 1 Step 2
       b1 = Asc(Mid$(data, n, 1))
       b1 = b1 * 256
       b2 = Asc(Mid$(data, n + 1, 1))
       suma = suma + b1 + b2
   Next n
   crr = suma \ 65536
   suma = (suma Mod 65536) + crr
   IPCheckSum = Right$(Hex((Not suma)), 4)
End Function
Public Function TCPCheckSum(ByRef pckTCP As TCP_STR) As String
   Dim dl As Long
   Dim suma As Long
   Dim crr As Long
   Dim b1 As Long
   Dim b2 As Long
   Dim tr As String
   Dim data As String
   dl = Len(pckTCP.data) / 2
   If dl Mod 2 = 1 Then tr = "00"
   data = pckTCP.x00_SourceIP & pckTCP.x01_DestinIP & "0006" & "00" & BHexa(Val(Left$(pckTCP.a04_Length, 1)) * 4) 'Pseudoheader
   data = data & pckTCP.a00_SourcePort & pckTCP.a01_DestinPort & pckTCP.a02_RelSeqNum & pckTCP.a03_RelAckNum & pckTCP.a04_Length & pckTCP.a05_FlagsByte & pckTCP.a06_WindowsSize & "0000" & pckTCP.o01_Options 'TCP Header
   data = data & pckTCP.data & tr
   data = HexaASCII(data)
   For n = 1 To Len(data) - 1 Step 2
       b1 = Asc(Mid$(data, n, 1))
       b1 = b1 * 256
       b2 = Asc(Mid$(data, n + 1, 1))
       suma = suma + b1 + b2
   Next n
   crr = suma \ 65536
   suma = (suma Mod 65536) + crr
   TCPCheckSum = Right$(Hex((Not suma)), 4)
End Function


Public Function UDPCheckSum(ByRef pckUDP As UDP_STR) As String
   Dim dl As Long
   Dim suma As Long
   Dim crr As Long
   Dim b1 As Long
   Dim b2 As Long
   Dim tr As String
   Dim data As String
   dl = Len(pckUDP.data) / 2
   If dl Mod 2 = 1 Then tr = "00"
   data = pckUDP.x00_SourceIP & pckUDP.x01_DestinIP & "0011" & pckUDP.a02_Length & pckUDP.a00_SourcePort & pckUDP.a01_DestinPort & pckUDP.a02_Length & "0000" & pckUDP.data & tr
   data = HexaASCII(data)
   For n = 1 To Len(data) - 1 Step 2
       b1 = Asc(Mid$(data, n, 1))
       b1 = b1 * 256
       b2 = Asc(Mid$(data, n + 1, 1))
       suma = suma + b1 + b2
   Next n
   crr = suma \ 65536
   suma = (suma Mod 65536) + crr
   UDPCheckSum = Right$(Hex((Not suma)), 4)
End Function


Public Function CompileFrameControl(ByRef FC As TPE_FrameControl) As String
    Dim b1 As Byte
    Dim b2 As Byte
    b1 = FC.Subtype * 16
    b1 = b1 + FC.Type * 4
    b1 = b1 + FC.Version
    b2 = b2 + IIf(FC.ToDS, 1, 0)
    b2 = b2 + IIf(FC.FromDS, 2, 0)
    b2 = b2 + IIf(FC.MoreFragment, 4, 0)
    b2 = b2 + IIf(FC.Retry, 8, 0)
    b2 = b2 + IIf(FC.Power, 16, 0)
    b2 = b2 + IIf(FC.More, 32, 0)
    b2 = b2 + IIf(FC.WEP, 64, 0)
    b2 = b2 + IIf(FC.Order, 128, 0)
    CompileFrameControl = DecHex(b1) & DecHex(b2)
End Function
Public Sub SetFrameControl(ByVal data As String, ByRef FrameControl As TPE_FrameControl)
    Dim b1 As Byte
    Dim b2 As Byte
    b1 = Asc(Mid$(data, 1, 1))
    b2 = Asc(Mid$(data, 2, 1))
    
    FrameControl.Version = b1 And 3
    FrameControl.Type = (b1 And 12) / 4
    FrameControl.Subtype = (b1 And 240) / 16
    
    FrameControl.ToDS = b2 And 1
    FrameControl.FromDS = b2 And 2
    FrameControl.MoreFragment = b2 And 4
    FrameControl.Retry = b2 And 8
    FrameControl.Power = b2 And 16
    FrameControl.More = b2 And 32
    FrameControl.WEP = b2 And 64
    FrameControl.Order = b2 And 128
    
End Sub
