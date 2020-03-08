VERSION 5.00
Begin VB.Form frm_lector 
   Caption         =   "Lector de ficheros .cap"
   ClientHeight    =   7050
   ClientLeft      =   2100
   ClientTop       =   3150
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   15030
   Begin VB.CommandButton pb_buscar 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox t_busq 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "Buscar..."
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox t_datos 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   7935
   End
   Begin VB.TextBox t_wep 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Text            =   "X0001389FA33B"
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   5910
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Clave WEP:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_lector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As Collection
Private Sub Form_Load()
    Dim a As String
    a = Dir$(App.Path & "\sesiones\inspector\WLAN*", vbDirectory)
    Do While a > ""
       List1.AddItem a
       a = Dir()
    Loop
End Sub

Private Sub List1_dblClick()
   List2.Clear
   ProcesaCAP App.Path & "\sesiones\inspector\" & List1.Text & "\WTC_" & List1.Text & ".cap"
End Sub
Private Sub ProcesaCAP(ByVal cap As String)
   Dim f As Integer
   Dim o As Integer
   Dim r As String
   Dim l As Long
   Dim sal As String
   Dim os As String
   Screen.MousePointer = 11
   Set P = New Collection
   f = FreeFile
   Open cap For Binary Access Read As #f
   o = FreeFile
   sal = Replace(cap, ".cap", ".txt")
   Open sal For Binary Access Write As #o
   r = Space$(24)
   Get #f, , r
   Do
      r = Space$(16)
      Get #f, , r
      l = Asc(Mid$(r, 9, 1)) + Asc(Mid$(r, 10, 1)) * 256
      If l = 0 Then Exit Do
      r = Space$(l)
      Get #f, , r
      os = Desencripta(r, t_wep.Text)
      P.Add "Sin datos"
      List2.AddItem ProcesaPaquete(os)
      Put #o, , os
   Loop
   Close
   Screen.MousePointer = 0
End Sub
Private Function ProcesaPaquete(ByVal PLD As String) As String
   Dim pwr As Integer
   Dim frmctr As TPE_FrameControl
   Dim pckdMAC As String
   Dim pcksMAC As String
   Dim pckBSSID As String
   Dim sIP As String
   Dim dIP As String
   Dim sPort As Long
   Dim dPort As Long
   Dim pl1 As Long
   Dim pl2 As Long
   Dim RutaPack As String
   Static Ultimo As String
   Dim Actual As String
   Dim DHCPData As String
   Dim iDHCPData As Integer
   Dim jDHCPData As Integer
   Dim lDHCPData As Integer
   P.Remove (P.Count)
   P.Add "Sin datos"
   Select Case ASCIIHexa(Mid$(PLD, 7, 2)) 'Protocolo 1
            Case "0800" 'IP
                 sIP = ip(Mid$(PLD, 21, 4))
                 dIP = ip(Mid$(PLD, 25, 4))
                 pl1 = Asc(Mid$(PLD, 29, 1))
                 pl2 = Asc(Mid$(PLD, 31, 1))
                 sPort = (pl1 * 256) + Asc(Mid$(PLD, 30, 1))
                 dPort = (pl2 * 256) + Asc(Mid$(PLD, 32, 1))
                 RutaPack = sIP & ":" & sPort & "-->" & dIP & ":" & dPort
                 Select Case Asc(Mid$(PLD, 18)) 'Protocolo 2
                        Case 1 'ICMP
                             ProcesaPaquete = "ICMP"
                        Case 6 'TCP
                            
                              
'''
                             
                             'MSN Messenger 1863
                             If sPort = 1863 Or dPort = 1863 Then
                                ProcesaPaquete = "MSN Messenger"
                                
                                

                                  ' Messenger Mid$(PLD, 1)

                             End If
                             'HTTP 80
                             If dPort = 80 Then
                                ProcesaPaquete = "HTTP"

                                P.Remove (P.Count)
                                'If InStr(PLD, "mandib") > 0 Then
                                '   kk = True
                                'End If
                                P.Add HTTP(PLD)
                                'Cookie dIP, PLD
                             End If
                             
                             'HTTPS 443
                             If dPort = 443 Then
                                ProcesaPaquete = "HTTPS"

                                'Cookie dIP, PLD
                             End If
                             
                             'FTP 21
                             If dPort = 21 Then
                                 ProcesaPaquete = "FTP"

                             End If
                             
                             '110 POP
                             If dPort = 110 Then
                                ProcesaPaquete = "POP3"

                             End If
                             
                             '5222 GTALK
                             If dPort = 5222 Then
                                ProcesaPaquete = "GTALK"

                             End If
                             
                             'Debug.Print dPort
                        Case 17 'UDP
                                ProcesaPaquete = "UDP"

                             
                        Case Else
                                ProcesaPaquete = "Otro " & "[" & Asc(Mid$(PLD, 18)) & "] "
                 End Select
    End Select
    ProcesaPaquete = ProcesaPaquete & " - " & RutaPack
End Function
Private Function HTTP(ByVal P As String) As String
    Dim i As Integer
    Dim T As Integer
    Dim dIP As String
    Dim l As String
    Dim Pet As String
    Dim li() As String
    Dim URL As String
    'p = Replace(p, Chr$(0), "·")
    dIP = ip(Mid$(P, 25, 4))
    
    If Mid$(P, 49, 4) = "POST" Then
       Pet = "P"
    ElseIf Mid$(P, 49, 3) = "GET" Then
       Pet = "G"
    End If
    'If Pet = "" Then Exit Function
    P = Mid$(P, 49)
    HTTP = P
    li = Split(P, vbCrLf)
    If Pet = "P" Then
       URL = Mid$(li(0), 6)
    Else
       URL = Mid$(li(0), 5)
    End If
    T = InStr(URL, " ")
    If T > 0 Then
       URL = Left$(URL, T - 1)
    End If
    'URL = "[" & Pet & "] " & "http://" & ResolveIP(dIP) & "/" & Replace(Mid$(li(1), 7) & URL, "//", "/")
    'lst_urls.AddItem URL
    'WHTTP URL
    'For i = 0 To UBound(li)
    '    If Left$(li(i), 8) = "Cookie: " Then
    '       WCookie Mid$(li(1), 7), Mid$(li(i), 9)
    '       Exit For
    '    End If
    'Next i
    'If Len(li(UBound(li))) > 4 And Pet = "P" Then
    '   lst_urls.AddItem "Datos: " & Left$(li(UBound(li)), Len(li(UBound(li))) - 4)
    '   WHTTP "Datos: " & Left$(li(UBound(li)), Len(li(UBound(li))) - 4)
    'End If
    
    'If lst_urls.ListCount > 13 Then
    '   lst_urls.TopIndex = lst_urls.ListCount - 13
    'End If
End Function

Private Sub List2_Click()
    t_datos.Text = P.Item(List2.ListIndex + 1)
End Sub

Private Sub pb_buscar_Click()
    Dim n As Long
    Dim i As Integer
    On Error Resume Next
    Screen.MousePointer = 11
    For n = List2.ListIndex + 1 To List2.ListCount - 1
        i = InStr(P(n + 1), t_busq.Text)
        If i > 0 Then
           List2.ListIndex = n
           DoEvents
           t_datos.SelStart = i - 1
           t_datos.SelLength = Len(t_busq.Text)
           t_datos.SetFocus
           Exit For
        End If
    Next n
    Screen.MousePointer = 0
End Sub
