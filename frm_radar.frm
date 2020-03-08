VERSION 5.00
Begin VB.Form frm_radar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10725
   ClientLeft      =   5430
   ClientTop       =   1245
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   8925
   Begin VB.Label l_p 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WLAN_XX"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape sh_p 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   600
      Shape           =   2  'Oval
      Top             =   10080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape sh_aimer 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4320
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   6255
      Index           =   4
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   6255
      Index           =   3
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   6255
      Index           =   2
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   7935
      Index           =   1
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   480
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Height          =   8655
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   8655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   0
      Y2              =   9240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frm_radar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Long
Dim r As Long
Dim cp As Integer
Private Type Punto_TPE
   X As Long
   Y As Long
End Type

Private Sub Form_Load()
    d = 120
    r = 4380
End Sub
Private Sub CnvPunto(ByRef Punto As Punto_TPE)
    Dim X As Long
    Dim Y As Long
    If Punto.X > 1000 Then Punto.X = 1000
    If Punto.X < -1000 Then Punto.X = -1000
    If Punto.Y > 1000 Then Punto.Y = 1000
    If Punto.Y < -1000 Then Punto.Y = -1000
    
    X = d + r + (Punto.X / 1000 * r)
    Y = d + r - (Punto.Y / 1000 * r)
    Punto.X = X
    Punto.Y = Y
End Sub
Private Sub AñadeAP(ByVal ESSID As String, ByVal X As Integer, ByVal Y As Integer)
    Dim P As Punto_TPE
    P.X = X
    P.Y = Y
    CnvPunto P
'    p.X = p.X - l_p(0).Width / 2
'    p.Y = p.Y - sh_p(0).Height / 2
    cp = cp + 1
    Load l_p(cp)
    Load sh_p(cp)
    l_p(cp).Left = P.X - l_p(cp).Width / 2
    l_p(cp).Top = P.Y + sh_p(cp).Height / 2
    sh_p(cp).Left = P.X - sh_p(cp).Width / 2
    sh_p(cp).Top = P.Y - sh_p(cp).Height / 2
    l_p(cp).Caption = ESSID
    l_p(cp).Visible = True
    sh_p(cp).Visible = True
End Sub
Private Sub PosicionaAimer(ByVal Pos As String)
    Dim P As Punto_TPE
    Select Case Pos
       Case "N"
            P.X = 0
            P.Y = 1000
       Case "S"
            P.X = 0
            P.Y = -1000
       Case "E"
            P.X = 1000
            P.Y = 0
       Case "W"
            P.X = -1000
            P.Y = 0
       Case "C"
            P.X = 0
            P.Y = 0
    End Select
    CnvPunto P
    sh_aimer.Left = P.X - sh_aimer.Width / 2
    sh_aimer.Top = P.Y - sh_aimer.Width / 2
End Sub
