Attribute VB_Name = "Tests"






Public Sub TestUDP()
    Dim d As String
    Dim pUDP As UDP_STR
    Dim pIP As IP_STR
    Dim pEth As EthernetII_STR
   
    pUDP.a00_SourcePort = "044f"
    pUDP.a01_DestinPort = "b1fe"
    pUDP.x00_SourceIP = "ac16029f"
    pUDP.x01_DestinIP = "ea82deb6"
    pUDP.data = "3032323731aced0005774f01000100000001ea82deb60000b1fe0000000001000100000001ac16029f0000044e000000113137322e32322e322e3135393a313039390000000000000002000450494e4700010000000d0000000170771a0003554450000100000014000d50726f63657373536572766572"
   
    pIP.data = CompileUDP(pUDP)
    pIP.a09_SourceIP = pUDP.x00_SourceIP
    pIP.a10_DestinIP = pUDP.x01_DestinIP
    pIP.a07_Protocol = BHexa(IP_Protocol.UDP_Prot)
    pIP.a03_Identification = "54f1"
   
    pEth.data = CompileIP(pIP)
    pEth.a00_dMAC = "01005e02deb6"
    pEth.a01_sMAC = "0019bbcef51d"
    pEth.a02_Protocol = WHexa(Eth_Protocol.IP_Prot)
   
    d = CompileEthernetII(pEth)
End Sub
