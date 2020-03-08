VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_tablon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   5295
   ClientTop       =   3060
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7215
   Begin VB.TextBox t_dip 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox t_sip 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox t_dmac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox t_smac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cd01 
      Left            =   9120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4200
      Width           =   6975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "···"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox t_fichero 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IP Destino"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IP Origen"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC Destino"
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
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC Origen"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fichero WLB"
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
      Left            =   120
      TabIndex        =   1
      Top             =   145
      Width           =   1215
   End
End
Attribute VB_Name = "frm_tablon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargaFichero(ByVal fichero As String)
   Dim f As Integer
   Dim HD As String * 42
   Dim DT As String
   f = FreeFile
   List1.Clear
   Open fichero For Binary Access Read As #f
   Do
      Get #f, , HD
      List1.AddItem HD
      If Left$(HD, 10) = String$(10, Chr$(0)) Then Exit Do
      DT = Space$(Asc(Mid$(HD, Len(HD) - 1, 1)) * 256 + Asc(Right$(HD, 1)))
      Get #f, , DT
   Loop
   Close (f)
End Sub

Private Sub Command1_Click()
   cd01.Filter = "*.WLB|*.WLB"
   cd01.FilterIndex = 1
   cd01.InitDir = App.Path
   cd01.ShowOpen
   If cd01.FileName > "" Then
      t_fichero = cd01.FileName
   End If
   Screen.MousePointer = 11
      CargaFichero t_fichero
   Screen.MousePointer = 0
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm_capturas.Show
End Sub

Private Sub List1_dblClick()
   Dim f As Integer
   Dim HD As String * 42
   Dim DT As String
   Dim ipk As Integer
   f = FreeFile

   Open t_fichero.Text For Binary Access Read As #f
   Do
      Get #f, , HD
      If Left$(HD, 10) = String$(10, Chr$(0)) Then Exit Do
      DT = Space$(Asc(Mid$(HD, Len(HD) - 1, 1)) * 256 + Asc(Right$(HD, 1)))
      Get #f, , DT
      If ipk = List1.ListIndex Then
         ProcesaCaptura HD, DT
         Exit Sub
      End If
      ipk = ipk + 1
   Loop
   Close (f)
End Sub
Private Sub ProcesaCaptura(ByVal HD As String, ByVal DT As String)
   Dim p As String
   Dim k As String
   Dim bs As String
   Dim es As String
   Dim smac As String
   Dim dmac As String
   Dim sIP As String
   Dim dIP As String
   
   p = Mid$(DT, 25)
   k = Mid$(HD, 28, 13)
   p = Desencripta(p, k)
   If p <> "?" Then
      sIP = IP(Mid$(p, 21, 4))
      t_sip = sIP
      dIP = IP(Mid$(p, 25, 4))
      t_dip = dIP
      smac = MAC(Mid$(DT, 17, 6))
      t_smac = smac
      dmac = MAC(Mid$(DT, 5, 6))
      t_dmac = dmac
      t_datos.Text = p
   End If
End Sub

