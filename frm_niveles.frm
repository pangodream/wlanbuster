VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_niveles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scanner  de Canales"
   ClientHeight    =   7440
   ClientLeft      =   5565
   ClientTop       =   2145
   ClientWidth     =   4815
   Icon            =   "frm_niveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   4815
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   12
      Left            =   4440
      TabIndex        =   58
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   840
      Top             =   7680
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1560
      Picture         =   "frm_niveles.frx":0442
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   8520
      Width           =   225
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1320
      Picture         =   "frm_niveles.frx":068C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   48
      Top             =   8520
      Width           =   225
   End
   Begin MSWinsockLib.Winsock ws02 
      Left            =   120
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   34
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   33
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   32
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   30
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   29
      Top             =   360
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   28
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton op_ch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   4200
         Picture         =   "frm_niveles.frx":08D6
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   7
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   28
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   29
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   30
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   31
         Left            =   120
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   4200
         Picture         =   "frm_niveles.frx":0B20
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   56
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   6
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   24
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   25
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   26
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   27
         Left            =   120
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   4200
         Picture         =   "frm_niveles.frx":0D6A
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   55
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   5
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   20
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   21
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   22
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   23
         Left            =   120
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   4200
         Picture         =   "frm_niveles.frx":0FB4
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   54
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   4
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   16
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   17
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   18
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   19
         Left            =   120
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   4200
         Picture         =   "frm_niveles.frx":11FE
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   53
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   3
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   12
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   13
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   14
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   15
         Left            =   120
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   4200
         Picture         =   "frm_niveles.frx":1448
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   52
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   8
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   9
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   10
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   11
         Left            =   120
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   4200
         Picture         =   "frm_niveles.frx":1692
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   51
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   1
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   4
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   5
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   6
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   7
         Left            =   120
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frm_wlan 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox p_wlan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   4200
         Picture         =   "frm_niveles.frx":18DC
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   50
         Top             =   480
         Width           =   225
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   3
         Left            =   120
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   120
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000080FF&
         Height          =   135
         Index           =   1
         Left            =   120
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape shgen 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   120
         Top             =   600
         Width           =   3855
      End
      Begin VB.Shape sh_wlan 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   120
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label l_bssid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:11:22:33:44:55"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A4075&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label l_essid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WLAN_XX"
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
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
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
      Index           =   20
      Left            =   4440
      TabIndex        =   59
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
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
      Index           =   19
      Left            =   4080
      TabIndex        =   47
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
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
      Index           =   18
      Left            =   3720
      TabIndex        =   46
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
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
      Index           =   17
      Left            =   3360
      TabIndex        =   45
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "9 "
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
      Index           =   16
      Left            =   3000
      TabIndex        =   44
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "8 "
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
      Index           =   15
      Left            =   2640
      TabIndex        =   43
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "7 "
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
      Index           =   14
      Left            =   2280
      TabIndex        =   42
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "6 "
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
      Index           =   13
      Left            =   1920
      TabIndex        =   41
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "5 "
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
      Index           =   12
      Left            =   1560
      TabIndex        =   40
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "4 "
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
      Index           =   11
      Left            =   1200
      TabIndex        =   39
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 "
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
      Index           =   10
      Left            =   840
      TabIndex        =   38
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "2 "
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
      Index           =   9
      Left            =   480
      TabIndex        =   37
      Top             =   120
      Width           =   255
   End
   Begin VB.Label l_essid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 "
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
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frm_niveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const w = 3855
Const Max = 100
Dim ps(0 To 7) As Boolean
Dim LastT(0 To 7) As Long
Private Sub Form_Load()
   If Not Conecta() Then
      MsgBox "No se puede conectar"
   End If
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
   HabilitaCanal 6
   
End Function

Private Sub SetLevel(ByVal bssid As String, ByVal ESSID As String, ByVal Level As Integer)
   Dim Idx As Integer
   Dim wi As Integer
   Dim sm As Long
   Dim ism As Integer
   Idx = IndexBssid(bssid)
   If Idx > -1 Then
      LastT(Idx) = GetTickCount()
      l_bssid(Idx).Caption = bssid
      l_essid(Idx).Caption = ESSID
      wi = (Level / Max) * w
      sh_wlan(Idx).Width = wi
      'Revisar en alguna versión, debería servir para hacer suaves los saltos del gauge
      'sm = sh_wlan(Idx).Width
      'For ism = sm To wi Step ((wi - sm) / 1)
      '    sh_wlan(Idx).Width = ism
      '    sh_wlan(Idx).Refresh
      'Next ism
      If wi >= w * 0.75 Then
         sh_wlan(Idx).BackColor = Verde
      ElseIf wi > w * 0.5 Then
         sh_wlan(Idx).BackColor = Naranja
      ElseIf wi > w * 0.25 Then
         sh_wlan(Idx).BackColor = Rojo
      Else
         sh_wlan(Idx).BackColor = 0
      End If
      If ps(Idx) Then
         ps(Idx) = False
         p_wlan(Idx).Picture = pic(0).Picture
      Else
         ps(Idx) = True
         p_wlan(Idx).Picture = pic(1).Picture
      End If
      frm_wlan(Idx).Visible = True
      sh_wlan(Idx).Refresh
   End If
End Sub
Private Function IndexBssid(ByVal bssid As String) As Integer
   Dim n As Integer
   Dim NoVisible As Integer
   IndexBssid = -1
   NoVisible = -1
   For n = 0 To 7
       If frm_wlan(n).Visible = False And NoVisible = -1 Then
          NoVisible = n
       End If
       If bssid = l_bssid(n) Then
          IndexBssid = n
          Exit Function
       End If
   Next n
   If NoVisible > -1 Then IndexBssid = NoVisible
End Function

Private Sub Form_Unload(Cancel As Integer)
   frm_pres.Show
End Sub

Private Sub op_ch_Click(Index As Integer)
    Dim n As Integer
    HabilitaCanal Index + 1
    For n = 0 To 7
        AnulaAP n
    Next n
End Sub
Private Sub AnulaAP(ByVal Idx As Integer)
    frm_wlan(Idx).Visible = False
    l_bssid(Idx).Caption = "00:11:22:33:44:55"
End Sub

Private Sub Timer1_Timer()
   Dim n As Integer
   Dim T As Long
   T = GetTickCount()
   For n = 0 To 7
       If T - LastT(n) > 10000 Then
          AnulaAP n
       End If
   Next n
End Sub

Private Sub ws02_DataArrival(ByVal bytesTotal As Long)
   Dim a As String
   ws02.GetData a
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
Private Sub ProcesaPaquete(ByVal HDR As String, ByVal PLD As String)
   Dim q As String
   Dim n As Integer
   Dim i As Integer
   Dim bssid As String
   Dim ESSID As String
   Dim pType As String
   Dim pwr As Byte
   Dim pFrmCtr As TPE_FrameControl
   
   On Error Resume Next
   Select Case Asc(Left$(HDR, 1))
      Case 5 'De punto de Acceso
         ''Debug.Print ASCIIHexa(Left$(PLD, 28))
         P = Mid$(PLD, 29)
         SetFrameControl Left$(P, 2), pFrmCtr
         If pFrmCtr.Type > 0 Then Exit Sub
         bssid = MAC(Mid$(P, 11, 6))
      Case Else
         Exit Sub
   End Select
   i = InStr(36, P, Chr$(1))
   If i > 0 Then
      ESSID = Mid$(P, 39, i - 39)
   End If
   pwr = Asc(Left$(PLD, 1))
   If UCase$(Left$(ESSID, 5)) = "WLAN_" And Len(ESSID) = 7 Then
      SetLevel bssid, ESSID, pwr
   End If
End Sub
Private Sub HabilitaCanal(ByVal Canal As Integer)
    ws02.SendData Chr$(3) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(4) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(Canal)
    ws02.SendData Chr$(8) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
End Sub


