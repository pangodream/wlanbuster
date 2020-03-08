VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_envio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envío de Claves WLAN"
   ClientHeight    =   5055
   ClientLeft      =   4995
   ClientTop       =   2625
   ClientWidth     =   6060
   Icon            =   "frm_envio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lst_claves 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   2790
      Left            =   2640
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5880
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton pb_enviar 
      Caption         =   "Enviar y Recibir WEP"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   4680
      Width           =   2535
   End
   Begin VB.ListBox lst_bssids 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "De antemano te damos las gracias."
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
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frm_envio.frx":0442
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
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frm_envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim n As Integer
    Dim k As String
    CargaListaBSSIDs lst_bssids
    For n = 0 To lst_bssids.ListCount - 1
        k = LeeWEP(lst_bssids.List(n))
        If k > "" Then
           lst_claves.List(n) = "WLAN_" & Right$(k, 2) & " --> " & k
        End If
    Next n
    If lst_bssids.ListCount = 0 Then
       MsgBox "Aun no se ha encontrado ninguna clave WEP."
       Unload Me
    Else
       ActivaEnvio
    End If

End Sub
Private Sub ActivaEnvio()
    Dim n As Integer
    Dim activa As Boolean
    For n = 0 To lst_bssids.ListCount - 1
        If lst_claves.List(n) = "" Then
           activa = True
           Exit For
        End If
    Next n
    pb_enviar.Enabled = activa
End Sub


Private Sub lst_bssids_DblClick()
   Dim hash As String
   hash = GetSetting("WLAN_Buster", "BSSID", lst_bssids.Text)
   Clipboard.Clear
   Clipboard.SetText hash
   MsgBox "Hash de clave copiado al portapapeles. Puedes desbloquear la clave en http://www.aiasec.net/www/ubwep.php"
   
End Sub

Private Sub pb_enviar_Click()
    Dim n As Integer
    Dim a As String
    Dim k As String
    Dim m As Integer
    Dim o As Integer
    'Err = 0
    'On Error Resume Next
    For n = 0 To lst_bssids.ListCount - 1
        If lst_claves.List(n) = "" Then
           a = Inet1.OpenURL("http://" & WebServer & "/www/keylistener.php?hash=" & LeeClave(lst_bssids.List(n)) & "&nick=Mandatory", 0)
              m = InStr(a, "[key]")
              If m > 0 Then
                 m = m + 5
                 o = InStr(a, "[/key]")
                 If o > 0 Then
                    k = Mid$(a, m, o - m)
                    lst_claves.List(n) = "WLAN_" & Right$(k, 2) & " --> " & k
                    EscribeWEP lst_bssids.List(n), k
                 End If
              Else
                 MsgBox "Se ha producido un error enviando el hash al servidor o el hash no es correcto."
              End If
        End If
    Next n
End Sub
