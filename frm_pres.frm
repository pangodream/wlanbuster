VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_pres 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WLAN Buster "
   ClientHeight    =   6120
   ClientLeft      =   4710
   ClientTop       =   2595
   ClientWidth     =   7305
   Icon            =   "frm_pres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_pres.frx":08CA
   ScaleHeight     =   6120
   ScaleWidth      =   7305
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7440
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton pb_salir 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton pb_disector 
      Caption         =   "Visor Capturas"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton pb_scanner 
      Caption         =   "Scanner"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton pb_iniciar 
      Caption         =   "Capturas"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin MSWinsockLib.Winsock ws01 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label l_version 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label l_aviso 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   1935
      Left            =   2160
      TabIndex        =   1
      Top             =   3960
      Width           =   4935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm_pres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Activado As Boolean
Private Sub Form_Activate()
   'frm_envio.Show
   If Not Activado Then
      Activar
      On Error Resume Next
      Me.SetFocus
      pb_disector.Enabled = True
      pb_salir.Enabled = True
      Activado = True
   End If
End Sub

Private Sub Activar()
   On Error Resume Next
   Dim ms As String
   Shell App.Path & "\airserv-ng.exe -d " & """" & "commview.dll|debug" & """" & " -p 666 y", vbNormalFocus
   If EsperaAirserv() = False Then
      ms = "No consigo arrancar el Airserv-NG automáticamente!!!" & vbCrLf
      ms = ms & "Esto puede deberse a que no se ha instalado correctamente la aplicación, o que se cierra automáticamente por un fallo de configuración o que no tienes la tarjeta apropiada." & vbCrLf
      ms = ms & "Para solucionarlo puedes intentar arrancarlo manualmente tecleando lo siguiente en una consola desde el directorio de WLAN Buster:" & vbCrLf
      ms = ms & "airserv-ng.exe -d " & """" & "commview.dll|debug" & """" & " -p 666 y" & vbCrLf
      ms = ms & "Si ves que te pregunta algo, contéstale --> y <-- (sin las flechas) y pulsa ENTER" & vbCrLf
      ms = ms & "La ventana de MSDOS debe permanecer abierta si todo va bien." & vbCrLf
      ms = ms & "Minimízala (NO LA CIERRES) y vuelve a arrancar WLAN Buster." & vbCrLf
      MsgBox ms, vbCritical + vbOKOnly, "IMPOSIBLE ARRANCAR AIRSERV-NG"
      Exit Sub
   End If
   AppActivate App.Path & "\airserv-ng.exe"
   Espera 750
   SendKeys "y"
   Espera 250
   SendKeys vbCr
   If Not AirservFunciona() Then
      ms = "Parece que he conseguido arrancar el Airserv-NG, pero no parece funcionar correctamente." & vbCrLf
      ms = ms & "Si está arrancado se debería ver una pantalla de MS-DOS con un título como " & App.Path & "\airserv-ng.exe" & vbCrLf
      ms = ms & "Si no está la ventana algo ha ido mal. Puede ser que el Airserv se haya quedado preguntando algo, en ese caso pulsa en su ventana --> y <-- (sin las flechas) y ENTER y" & vbCrLf
      ms = ms & "comprueba que se queda abierto y escuchando. En la última línea debería poner:  " & vbCrLf
      ms = ms & "Serving commview.dll|debug chan 1 on port 666" & vbCrLf
      ms = ms & "Si es así, parece que va mejor, minimízalo (NO LO CIERRES) y vuelve a arrancar Wlan Buster." & vbCrLf
      ms = ms & "Si no es así, algo va mal y aquí no te puedo ayudar más. Quizá en la página www.aiasec.net encuentres ayuda." & vbCrLf
      MsgBox ms, vbExclamation, "Algo va mal con el Airserv-NG"
      Exit Sub
   End If
   MinimizaAirserv
   Espera 500
   pb_iniciar.Enabled = True
   pb_scanner.Enabled = True
End Sub

Private Sub MinimizaAirserv()
   Dim sTitle As String
   Dim iHwnd As Long
   Dim ihTask As Long
   Dim iReturn As Long
   sTitle = App.Path & "\airserv-ng.exe"
   iHwnd = FindWindow(0&, sTitle)
   iReturn = ShowWindow(iHwnd, 6)
End Sub
Private Function EsperaAirserv() As Boolean
   Const TOut = 4000
   Dim T As Long
   T = GetTickCount()
   Do
      If (GetTickCount() - T) > TOut Then Exit Function
      If ExisteAirserv() Then
         EsperaAirserv = True
         Exit Function
      End If
      DoEvents
   Loop
End Function
Private Function ExisteAirserv() As Boolean
   Dim sTitle As String
   Dim iHwnd As Long
   Dim ihTask As Long
   Dim iReturn As Long
   sTitle = App.Path & "\airserv-ng.exe"
   iHwnd = FindWindow(0&, sTitle)
   If iHwnd > 0 Then
      ExisteAirserv = True
   End If
End Function
Private Function AirservFunciona() As Boolean
   Dim T As Long
   ws01.RemoteHost = "127.0.0.1"
   ws01.RemotePort = 666
   ws01.Connect
   T = GetTickCount()
   Do While ((GetTickCount() - T) < 3000) And ws01.State <> 7
      DoEvents
   Loop
   If ws01.State = 7 Then
      AirservFunciona = True
      ws01.Close
   End If
End Function

Private Sub Form_Load()
   Dim Aviso As String
   Dim v As String
   l_version.Caption = "Versión " & Mid$(Version, 1, 1) & "." & Mid$(Version, 2, 1) & "." & Mid$(Version, 3, 1)

   Aviso = "Aviso Legal" & vbCrLf
   Aviso = Aviso & "El aprovechamiento de los recursos servidos por un punto de acceso sin permiso del propietario de ese determinado punto de acceso es un delito." & vbCrLf
   
   Aviso = Aviso & "WLAN Buster ha sido concebido como una herramienta de estudio de redes así como una prueba de concepto de ciertas debilidades en la configuración de punto de acceso de Telefónica de España, S.A. y por lo tanto no debe ser utilizado como utilidad para acceder a recursos o datos ajenos." & vbCrLf
   
   Aviso = Aviso & "Los creadores de WLAN Buster se eximen de cualquier mal uso que se pueda hacer de este programa o la información descubierta por él." & vbCrLf
   
   Aviso = Aviso & "Si no estás de acuerdo con cualquiera de los puntos anteriores o tienes cualquier duda acerca del uso que puedes hacer de esta aplicación, por favor pulsa el botón de salir." & vbCrLf
   
   l_aviso.Caption = Aviso
   PCName = GetMachineName()
   If PCName = "pccris" Then Exit Sub
   On Error Resume Next
   Me.Caption = Me.Caption & " V " & Mid$(Version, 1, 1) & "." & Mid$(Version, 2, 1) & "." & Mid$(Version, 3, 1)
   v = Inet1.OpenURL("http://" & WebServer & "/www/version.php")
   v = Mid$(v, InStr(v, "<body>") + 8, 3)
   If Version < v Then
      Me.Caption = Me.Caption & " (Versión Actual: " & Mid$(v, 1, 1) & "." & Mid$(v, 2, 1) & "." & Mid$(v, 3, 1) & ")"
      MsgBox "Existe una versión más reciente de WLAN Buster." & vbCrLf & "Consulta el foro de la herramienta en http://www.aiasec.net/bb para encontrar instrucciones de descarga."
   End If
End Sub

Private Sub pb_disector_Click()
   On Error Resume Next
   Activado = True
   frm_lector.Show
End Sub

Private Sub pb_iniciar_Click()
'    frm_niveles.Show
   Activado = False
   frm_capturas.Show
'   frm_assoc.Show
   Unload Me
End Sub

Private Sub pb_salir_Click()
   CierraAirserv
   End
End Sub

Private Sub pb_scanner_Click()
   Activado = False
   frm_niveles.Show
   Unload Me
End Sub
Private Sub CierraAirserv()
   Dim sTitle As String
   Dim iHwnd As Long
   Dim ihTask As Long
   Dim iReturn As Long
   Dim T As Long
   On Error Resume Next
   AppActivate App.Path & "\airserv-ng.exe"
   DoEvents
   SendKeys "^C"
   sTitle = App.Path & "\airserv-ng.exe"
   iHwnd = FindWindow(0&, sTitle)
   iReturn = CloseWindow(iHwnd)
   iReturn = DestroyWindow(iHwnd)
   T = GetTickCount()
   Do While (GetTickCount() - T) < 1000
      DoEvents
   Loop
End Sub
