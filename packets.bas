Attribute VB_Name = "Packets_Composer"
Public Function ARPRequest(ByVal bssid As String, ByVal smac As String, ByVal dmac As String, ByVal sIP As String, ByVal dIP As String, ByVal clave As String) As String
    Dim wf As WF80211_STR
    Dim llc As LLC_STR
    Dim arp As ARP_STR
    Dim s As String
    Dim w As String
    
    bssid = Replace(bssid, ":", "")
    smac = Replace(smac, ":", "")
    dmac = Replace(dmac, ":", "")
    If Len(sIP) > 8 Then
       sIP = IPHex(sIP)
    End If
    If Len(dIP) > 8 Then
       dIP = IPHex(dIP)
    End If
    
    arp.a04_Operation = "0001"
    arp.a05_SourceMAC = smac
    arp.a06_SourceIP = sIP
    arp.a07_DestinMAC = "000000000000"
    arp.a08_DestinIP = dIP
    s = CompileARP(arp)
    
    llc.a04_Protocol = "0806"
    s = CompileLLC(llc) & s
    
    wf.a01_FrameControl.ToDS = True
    wf.a01_FrameControl.WEP = True
    wf.a01_FrameControl.Type = PCK_TYPE.data
    wf.a01_FrameControl.Subtype = PCK_SUBTYPE.data_Data
    
    wf.a03_BSSID = bssid
    wf.a04_SourceMAC = smac
    wf.a05_DestinationMAC = dmac
    w = CompileWF(wf)
    ARPRequest = HexaASCII(w) & CreaPaqueteWEP(HexaASCII(s), clave)
End Function

