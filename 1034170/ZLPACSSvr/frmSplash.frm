VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4365
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1005
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image imgPic 
      Height          =   2700
      Left            =   150
      Picture         =   "frmSplash.frx":5D0A2
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label LblProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1590
      TabIndex        =   7
      Top             =   1350
      Width           =   4650
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用权属于："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   6
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "产品开发商："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   5
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "技术支持商："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   4
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label lblGrant 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   3
      Top             =   2205
      Width           =   90
   End
   Begin VB.Label lbl技术支持商 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   2
      Top             =   2610
      Width           =   90
   End
   Begin VB.Label lbl开发商 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   1
      Top             =   3030
      Width           =   90
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   165
      Picture         =   "frmSplash.frx":5DB05
      Stretch         =   -1  'True
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "警告：本软件受软件保护法和软件使用许可证保护。未经授权许可，任何人不得复制、销售及解密此软件，否则将承担全部法律责任。"
      Height          =   465
      Left            =   1065
      TabIndex        =   0
      Top             =   3825
      Width           =   5490
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
