VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareSendCard 
   Caption         =   "���ѿ�����"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13125
   Icon            =   "frmSquareSendCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   13125
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   61
      Top             =   10200
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareSendCard.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18071
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   13680
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareSendCard.frx":70E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareSendCard.frx":743A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCard 
      BorderStyle     =   0  'None
      Height          =   9240
      Left            =   2400
      ScaleHeight     =   9240
      ScaleWidth      =   11010
      TabIndex        =   62
      Top             =   600
      Width           =   11010
      Begin VB.PictureBox picCardInfor 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   8700
         Left            =   120
         ScaleHeight     =   8700
         ScaleWidth      =   10410
         TabIndex        =   63
         Top             =   240
         Width           =   10410
         Begin VB.Frame fraBaseInfor 
            Caption         =   "����������Ϣ"
            Height          =   4215
            Left            =   75
            TabIndex        =   33
            Top             =   165
            Width           =   7725
            Begin VB.CommandButton cmdSel 
               Caption         =   "��"
               Height          =   270
               Index           =   2
               Left            =   7335
               TabIndex        =   17
               TabStop         =   0   'False
               Tag             =   "�쿨����"
               Top             =   2010
               Width           =   285
            End
            Begin VB.CommandButton cmdSel 
               Caption         =   "��"
               Height          =   270
               Index           =   1
               Left            =   3105
               TabIndex        =   14
               TabStop         =   0   'False
               Tag             =   "�쿨��"
               Top             =   2040
               Width           =   285
            End
            Begin VB.CommandButton cmdSel 
               Caption         =   "��"
               Height          =   270
               Index           =   0
               Left            =   7335
               TabIndex        =   11
               TabStop         =   0   'False
               Tag             =   "����ԭ��"
               Top             =   1605
               Width           =   285
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   21
               Left            =   5340
               TabIndex        =   31
               Top             =   3630
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   22
               Left            =   1095
               TabIndex        =   29
               Top             =   3630
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   18
               Left            =   5340
               TabIndex        =   16
               Top             =   2010
               Width           =   2295
            End
            Begin VB.CheckBox chk�Ƿ��ֵ 
               Caption         =   "�Ƿ��ֵ��"
               Height          =   450
               Left            =   5355
               TabIndex        =   4
               Top             =   690
               Width           =   1830
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   9
               Left            =   1095
               TabIndex        =   25
               Top             =   3225
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   8
               Left            =   5340
               TabIndex        =   23
               Top             =   2835
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   7
               Left            =   1095
               TabIndex        =   21
               Top             =   2835
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   6
               Left            =   1110
               TabIndex        =   19
               Top             =   2430
               Width           =   6525
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   4
               Left            =   1110
               TabIndex        =   13
               Top             =   2025
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   3
               Left            =   1125
               MaxLength       =   50
               TabIndex        =   10
               Top             =   1605
               Width           =   6525
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   5340
               PasswordChar    =   "*"
               TabIndex        =   8
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1110
               PasswordChar    =   "*"
               TabIndex        =   6
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   0
               Left            =   1125
               TabIndex        =   3
               Top             =   795
               Width           =   2280
            End
            Begin VB.ComboBox cbo������ 
               Height          =   300
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   345
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtp����Ч���� 
               Height          =   300
               Left            =   5340
               TabIndex        =   27
               Top             =   3240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   529
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   111280131
               CurrentDate     =   40156.0854282407
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   5
               Left            =   5340
               TabIndex        =   64
               Top             =   3240
               Width           =   2295
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   11
               Left            =   4530
               TabIndex        =   30
               Top             =   3690
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   21
               Left            =   495
               TabIndex        =   28
               Top             =   3690
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�쿨����(&M)"
               Height          =   180
               Index           =   20
               Left            =   4260
               TabIndex        =   15
               Top             =   2070
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "��ǰ���"
               Height          =   180
               Index           =   10
               Left            =   330
               TabIndex        =   24
               Top             =   3285
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   9
               Left            =   4530
               TabIndex        =   22
               Top             =   2895
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   8
               Left            =   510
               TabIndex        =   20
               Top             =   2895
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "��ע(&S)"
               Height          =   180
               Index           =   7
               Left            =   435
               TabIndex        =   18
               Top             =   2490
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "����Ч����(&D)"
               Height          =   180
               Index           =   6
               Left            =   4080
               TabIndex        =   26
               Top             =   3285
               Width           =   1170
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�쿨��(&D)"
               Height          =   180
               Index           =   5
               Left            =   255
               TabIndex        =   12
               Top             =   2085
               Width           =   810
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "����ԭ��(&Y)"
               Height          =   180
               Index           =   4
               Left            =   90
               TabIndex        =   9
               Top             =   1650
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "����ȷ��(&E)"
               Height          =   180
               Index           =   3
               Left            =   4260
               TabIndex        =   7
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "����(&W)"
               Height          =   180
               Index           =   2
               Left            =   435
               TabIndex        =   5
               Top             =   1245
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "����(&N)"
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   2
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "������(&T)"
               Height          =   180
               Index           =   0
               Left            =   270
               TabIndex        =   0
               Top             =   405
               Width           =   810
            End
         End
         Begin VB.Frame fra��ֵ 
            Caption         =   "����ֵ���"
            Height          =   705
            Left            =   90
            TabIndex        =   53
            Top             =   4440
            Width           =   7710
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   11
               Left            =   5325
               TabIndex        =   38
               Top             =   270
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   10
               Left            =   1110
               TabIndex        =   36
               Top             =   270
               Width           =   2295
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "ʵ�����۶�(&J)"
               Height          =   180
               Index           =   13
               Left            =   4080
               TabIndex        =   37
               Top             =   330
               Width           =   1170
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�����(&M)"
               Height          =   180
               Index           =   12
               Left            =   270
               TabIndex        =   35
               Top             =   330
               Width           =   810
            End
         End
         Begin VB.Frame fra��ֵ��� 
            Caption         =   "��ֵ��Ϣ���"
            Height          =   1950
            Left            =   90
            TabIndex        =   66
            Tag             =   "1"
            Top             =   5235
            Width           =   10125
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   27
               Left            =   7185
               MaxLength       =   30
               TabIndex        =   46
               Tag             =   "1"
               Top             =   1470
               Width           =   2280
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   14
               Left            =   7890
               TabIndex        =   72
               Top             =   330
               Width           =   2040
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   12
               Left            =   4320
               TabIndex        =   71
               Top             =   330
               Width           =   1635
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   13
               Left            =   1230
               TabIndex        =   70
               Top             =   323
               Width           =   1515
            End
            Begin VB.ComboBox cboStyle 
               Height          =   300
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Tag             =   "1"
               Top             =   1110
               Width           =   1965
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   25
               Left            =   4320
               MaxLength       =   50
               TabIndex        =   44
               Tag             =   "1"
               Top             =   1110
               Width           =   5640
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   1200
               MaxLength       =   20
               TabIndex        =   45
               Tag             =   "1"
               Top             =   1470
               Width           =   4905
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   24
               Left            =   4305
               TabIndex        =   42
               Top             =   698
               Width           =   5640
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   23
               Left            =   1230
               TabIndex        =   40
               Top             =   698
               Width           =   1695
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               Height          =   180
               Left            =   6360
               TabIndex        =   77
               Top             =   1530
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���γ�ֵ(&B)"
               Height          =   180
               Index           =   14
               Left            =   3150
               TabIndex        =   76
               Top             =   390
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ֵ����(&K)"
               Height          =   180
               Index           =   15
               Left            =   135
               TabIndex        =   75
               Top             =   390
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʵ�ʳ�ֵ�ɿ�(&I)"
               Height          =   180
               Index           =   16
               Left            =   6480
               TabIndex        =   74
               Top             =   390
               Width           =   1350
            End
            Begin VB.Label lblPer 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2790
               TabIndex        =   73
               Top             =   375
               Width           =   120
            End
            Begin VB.Label lblzffs 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "֧����ʽ"
               Height          =   240
               Left            =   405
               TabIndex        =   69
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label lblkhh 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               Height          =   240
               Left            =   3720
               TabIndex        =   68
               Top             =   1140
               Width           =   600
            End
            Begin VB.Label lblzh 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʺ�"
               Height          =   240
               Left            =   720
               TabIndex        =   67
               Top             =   1500
               Width           =   480
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "��ֵ˵��(&Z)"
               Height          =   180
               Index           =   23
               Left            =   3255
               TabIndex        =   41
               Top             =   765
               Width           =   990
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "�ɿ���(&R)"
               Height          =   180
               Index           =   22
               Left            =   315
               TabIndex        =   39
               Top             =   765
               Width           =   810
            End
         End
         Begin VB.Frame fra�ɿ� 
            Caption         =   "���νɿ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   120
            TabIndex        =   65
            Top             =   7440
            Width           =   10095
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   17
               Left            =   4320
               TabIndex        =   50
               Text            =   "12"
               ToolTipText     =   "���νɿ�"
               Top             =   390
               Width           =   2010
            End
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   16
               Left            =   825
               TabIndex        =   48
               Text            =   "12"
               ToolTipText     =   "ʵ�պϼ�"
               Top             =   405
               Width           =   2010
            End
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   15
               Left            =   7770
               TabIndex        =   52
               Text            =   "12"
               ToolTipText     =   "���νɿ�"
               Top             =   405
               Width           =   2010
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�ɿ�(&U)"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   19
               Left            =   3060
               TabIndex        =   49
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�ϼ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   18
               Left            =   105
               TabIndex        =   47
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "�Ҳ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   17
               Left            =   6960
               TabIndex        =   51
               Top             =   480
               Width           =   660
            End
         End
         Begin VB.Frame fra������� 
            Caption         =   "�������(&X)"
            Height          =   4980
            Left            =   7905
            TabIndex        =   34
            Top             =   180
            Width           =   2310
            Begin MSComctlLib.ListView lvwType 
               Height          =   4665
               Left            =   75
               TabIndex        =   32
               Top             =   240
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   8229
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FlatScrollBar   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Key             =   "����"
                  Object.Tag             =   "����"
                  Text            =   "����"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   480
      ScaleHeight     =   7755
      ScaleWidth      =   4350
      TabIndex        =   54
      Top             =   240
      Width           =   4350
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   20
         Left            =   1005
         TabIndex        =   60
         Top             =   525
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vsGrid 
         Height          =   5235
         Left            =   120
         TabIndex        =   55
         Top             =   930
         Width           =   4065
         _cx             =   7170
         _cy             =   9234
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   28
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareSendCard.frx":778E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picImg 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   56
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareSendCard.frx":7B98
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   19
         Left            =   1005
         TabIndex        =   58
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label lbl��ʼ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   195
         TabIndex        =   57
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   195
         TabIndex        =   59
         Top             =   555
         Width           =   720
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   180
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSquareSendCard.frx":80E6
      Left            =   210
      Top             =   345
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareSendCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintSucces As Integer
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mCardEditType As gCardEditType
Private mlng���ѿ�ID As Long, mblnFirst As Boolean, mlng�ӿڱ�� As Long
Private mblnNoClick As Boolean, mblnChange As Boolean
Private mfrmMain As Form
Private mrs��� As ADODB.Recordset
Private mrs���ѿ����� As ADODB.Recordset
Private mblnUnLoad As Boolean
Private Type TyCardInfor
    str���� As String
    dbl����� As Double
    dblʵ������ As Double
    dbl����ֵ As Double
    dbl�ۿ��� As Double
    int�ɷ��ֵ As Integer
    strЧ�� As String
    bln�����ֵ As Boolean  '���޸�ʱ��Ч,����Ե�ǰ�����ĳ�ֵ��¼��ֻ��һ��,�Ϳ������¸�����صĳ�ֵ��Ϣ,
    bln������ As Boolean    '�����˵�,�Ͳ��ܸ�����ֵ���Ժ���ֵ
    lng��ֵ���� As Long
End Type
Private mCardInfor As TyCardInfor
Private Enum mPaneID
    Pane_Cards = 1     '�������
    Pane_CardInfor = 2  '����Ϣ
End Enum
Private mPanSearch As Pane
Private Enum mtxtIdx
    idx_txt���� = 0
    idx_txt���� = 1
    idx_txtȷ������ = 2
    idx_txt����ԭ�� = 3
    idx_txt�쿨�� = 4
    idx_txt����Ч���� = 5
    idx_txt��ע = 6
    idx_txt������ = 7
    idx_txt�������� = 8
    idx_txt��ǰ��� = 9
    idx_txt����� = 10
    idx_txtʵ�����۶� = 11
    idx_txt���γ�ֵ = 12
    idx_txt��ֵ���� = 13
    idx_txtʵ�ʳ�ֵ�ɿ� = 14
    idx_txt�Ҳ� = 15
    idx_txtʵ�պϼ� = 16
    idx_txt���νɿ� = 17
    idx_txt�쿨���� = 18
    idx_txt��ʼ���� = 19
    idx_txt�������� = 20
    idx_txt������ = 22
    idx_txt����ʱ�� = 21
    idx_txt�ɿ��� = 23
    idx_txt��ֵ��ע = 24
    idx_txt������ = 25
    idx_txt�ʺ� = 26
    idx_txt������� = 27
End Enum
Private Enum mcmdIdx
    idx_cmd����ԭ�� = 0
    idx_cmd�쿨�� = 1
    idx_cmd�쿨���� = 2
End Enum
Private mlngSel����ID As Long
Private Enum mlblIdx
    idx_lbl�Ҳ� = 17
End Enum
Private Const mconMenu_Edit_Affirm = 225
Private Const FM_HEIGHT = 10000  '�༭���ڵĴ��ڴ�С
Private Const FM_WIDTH = 10935  '�༭���ڵĴ��ڴ�С
Private Const PIC_CARD_HEIGHT = 8700    '��Ƭ��ĸ߶�
Private Const PIC_CARD_WIDTH = 7050     '��Ƭ��Ŀ��
Private Type Ty_Para
    str����ǰ׺ As String
    lng���ų��� As Long
    bln�������� As Boolean
    bln�ɿ��ӡ As Boolean
End Type
Private mTy_MoudlePara As Ty_Para
Private mblnHaveOtherCard As Boolean '�Ƿ���������Ŀ�Ƭ(����ʱ)
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mdblʵ�պϼ� As Double
Private mobjKeyboard As Object
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����

Public Function zlShowCard(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, ByVal CardEditType As gCardEditType, ByVal lng�ӿڱ�� As Long, Optional lng���ѿ�ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,�鿴�ѷ��������ӷ������޸ķ�����Ϣ
    '����:
    '����:���˺�
    '����:2009-12-09 13:40:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng���ѿ�ID = lng���ѿ�ID: mCardEditType = CardEditType
    mlng�ӿڱ�� = lng�ӿڱ��

    With gTy_TestBug
        If CardEditType = gEd_���� Then
            .BytType = 1
        Else
            .BytType = 2
        End If
    End With
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2009-12-10 17:31:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet���ѿ��ӿ�()
    With rsTemp
        .Filter = 0
        .Find "���=" & mlng�ӿڱ��, , adSearchForward, 1
        If Not rsTemp.EOF Then
            With mTy_MoudlePara
                .str����ǰ׺ = Nvl(rsTemp!ǰ׺�ı�)
                .lng���ų��� = Val(Nvl(rsTemp!���ų���))
                .bln�������� = Val(Nvl(rsTemp!�Ƿ�����)) = 1
            End With
        End If
    End With
    txtEdit(mtxtIdx.idx_txt��ʼ����).MaxLength = mTy_MoudlePara.lng���ų���
    txtEdit(mtxtIdx.idx_txt��������).MaxLength = mTy_MoudlePara.lng���ų���
    txtEdit(mtxtIdx.idx_txt����).MaxLength = mTy_MoudlePara.lng���ų���

    With mTy_MoudlePara
        .bln�ɿ��ӡ = Val(zlDatabase.GetPara("�ɿ��ӡ", glngSys, mlngModule)) = 1
    End With
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-12-10 10:19:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Cards, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "����������Ϣ": mPanSearch.Options = PaneNoCloseable
        mPanSearch.Handle = picList.hWnd
        Set objPane = .CreatePane(mPaneID.Pane_CardInfor, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = "����Ϣ"
        If mCardEditType = gEd_���� Or mCardEditType = gEd_���� Then
            objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Else
            objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            mPanSearch.Closed = True
        End If

        objPane.Handle = picCard.hWnd
        objPane.MaxTrackSize.Width = picCard.Width \ Screen.TwipsPerPixelX
        objPane.MinTrackSize.Width = picCard.Width \ Screen.TwipsPerPixelX
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With

'    If mCardEditType = gEd_���� Then
'        zlRestoreDockPanceToReg Me, dkpMan, "����-����"
'    ElseIf mCardEditType = gEd_���� Then
'        zlRestoreDockPanceToReg Me, dkpMan, "����-����"
'    End If

End Function

Private Function CheckDepented() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ������
    '���:
    '����:
    '����:���ݹ����Ϸ�,����true, ���򷵻�False
    '����:���˺�
    '����:2009-12-09 14:28:26
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    Set mrs��� = zlGet�շ����
    If mrs���.RecordCount = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   û����ص��շ���Ŀ���,����ϵͳ����Ա��ϵ!"
        Exit Function
    End If

    With lvwType
        .ListItems.Clear
         Do While Not mrs���.EOF
            .ListItems.Add , "K" & Nvl(mrs���!����), Nvl(mrs���!����) & "-" & Nvl(mrs���!����)
            mrs���.MoveNext
         Loop
         mrs���.MoveFirst
    End With
    gstrSQL = "Select rownum as ID, ����,����, ȱʡ���, ȱʡ�ۿ�, ȱʡ��־ From ���ѿ�����"
    Set mrs���ѿ����� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If mrs���ѿ�����.RecordCount = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   û��������ص����ѿ�����,����[�ֵ����]������!"
        Exit Function
    End If

    zlComboxLoadFromRecodeset Me.Caption, mrs���ѿ�����, cbo������, True
    '����Ƿ���������ص�ˢ������
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng�ӿڱ��)

    CheckDepented = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-10 10:24:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup


    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With

    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.Visible = False

    '�����
    With cbsThis.KeyBindings
        .Add FALT, Asc("O"), mconMenu_Edit_Affirm
        .Add FALT, Asc("X"), conMenu_Edit_CardModify
     End With

    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched

    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): mcbrControl.BeginGroup = True

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoveCard, "�Ƴ���Ƭ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Apply_AllCard, "Ӧ��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Apply_AllColumn, "Ӧ���ڴ���"): mcbrControl.BeginGroup = True

        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_Affirm, "ȷ��  "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
    End With

    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ClearCtlData(Optional blnClearVsGridData As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2009-12-09 15:24:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control, lngID As Long, i As Long
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Then
            ctl.Text = ""
        End If
    Next
    For i = 1 To Me.lvwType.ListItems.count
        lvwType.ListItems(i).Checked = False
    Next
    stbThis.Panels(2).Text = ""
    dtp����Ч����.value = Null
    If blnClearVsGridData Then
        With vsGrid
            .Rows = 2
            .Clear 1
        End With
    End If
    Call SetDefaultValue
    Call Show��������
End Sub
Private Sub SetDefaultValue(Optional bln��� As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡֵ
    '����:���˺�
    '����:2009-12-09 16:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, blnDefault��ֵ As Boolean
    '����ȱʡ��ֵ
    lngID = cbo������.ItemData(cbo������.ListIndex)
    mrs���ѿ�����.Filter = 0
    If mrs���ѿ�����.RecordCount <> 0 Then mrs���ѿ�����.MoveFirst
    mrs���ѿ�����.Find "ID=" & lngID, , adSearchForward, 0

    If Not mrs���ѿ�����.EOF Then
'        If Val(txtEdit(idx_txt��ֵ����).Tag) <> 0 Or Val(txtEdit(idx_txt��ֵ����).Text) = 0 Then
            txtEdit(idx_txt��ֵ����).Text = Format(Val(Nvl(mrs���ѿ�����!ȱʡ�ۿ�, 100)), "0.00")
            txtEdit(idx_txt��ֵ����).Tag = txtEdit(idx_txt��ֵ����).Text
            txtEdit(idx_txtʵ�ʳ�ֵ�ɿ�).Text = Format(Val(Nvl(mrs���ѿ�����!ȱʡ�ۿ�, 100)) * Val(txtEdit(idx_txt���γ�ֵ).Text) / 100, "0.00")
'        End If
        If bln��� Then
            txtEdit(idx_txt�����).Text = Format(Val(Nvl(mrs���ѿ�����!ȱʡ���)), "0.00")
            txtEdit(idx_txtʵ�����۶�).Text = Format(Val(Nvl(mrs���ѿ�����!ȱʡ���)) * (txtEdit(idx_txt��ֵ����).Text / 100), "0.00")
        End If
        Call ModifyGridMoney
    Else
        Call Calc���
        Call Calcʵ�պϼ�
    End If
End Sub
Private Sub ModifyGridMoney()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����¸��������еĳ�ֵ����,������ֵ������
    '���ƣ����˺�
    '���ڣ�2010-03-26 13:54:24
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnDefault��ֵ  As Boolean
    blnDefault��ֵ = mCardInfor.bln�����ֵ And chk�Ƿ��ֵ.value = 1

    '�ȼ������
    Call Calc���
    If mCardEditType = gEd_���� Then
        With vsGrid
            If Split(.Cell(flexcpData, .Row, .ColIndex("����")) & ",", ",")(0) = txtEdit(mtxtIdx.idx_txt����).Text Then
                If blnDefault��ֵ Then
                    .TextMatrix(.Row, .ColIndex("��ֵ����")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ����).Text)
                    .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = Trim(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text)
                    .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text)
                End If
                .TextMatrix(.Row, .ColIndex("�����")) = Trim(txtEdit(idx_txt�����).Text)
                .TextMatrix(.Row, .ColIndex("ʵ������")) = Trim(txtEdit(idx_txtʵ�����۶�).Text)
            End If
            .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
        End With
    End If
    Call Calcʵ�պϼ�
End Sub


Private Sub Calcʵ�պϼ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʵ�պϼ�
    '����:���˺�
    '����:2009-12-09 16:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblʵ�պϼ� As Double, dbl��� As Double, i As Long
    Dim dbl�ɿ� As Double
    dblʵ�պϼ� = 0

    If mCardEditType = gEd_���� Then
        '�����ܺϼ�
        With vsGrid
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("����"))) <> "" Then
                    dbl�ɿ� = IIf(Val(.Cell(flexcpData, i, .ColIndex("�Ƿ��ֵ"))) <> 1, 0, Val(.TextMatrix(i, .ColIndex("ʵ�ʳ�ֵ�ɿ�")))) + Val(.TextMatrix(i, .ColIndex("ʵ������")))
                    dbl�ɿ� = dbl�ɿ� * Val(.TextMatrix(i, .ColIndex("��������")))
                    dblʵ�պϼ� = dblʵ�պϼ� + dbl�ɿ�
                End If
            Next
        End With
    ElseIf mCardEditType = gEd_�޸� Then
        dblʵ�պϼ� = dblʵ�պϼ� + IIf(chk�Ƿ��ֵ.value = 0, 0, Val(txtEdit(idx_txtʵ�ʳ�ֵ�ɿ�).Text)) + Val(txtEdit(idx_txtʵ�����۶�).Text)
    Else
        dblʵ�պϼ� = dblʵ�պϼ� + Val(txtEdit(idx_txtʵ�ʳ�ֵ�ɿ�).Text)
    End If
    mdblʵ�պϼ� = dblʵ�պϼ�
    txtEdit(idx_txtʵ�պϼ�).Text = Format(dblʵ�պϼ�, "0.00")
    Call SetLblCatpion
End Sub
Private Sub Calc���()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2009-12-09 16:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblʵ�պϼ� As Double, dbl��� As Double, i As Long
    If Not (mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Or mCardEditType = gEd_��ֵ) Then Exit Sub
    If mCardEditType = gEd_�޸� And (mCardInfor.bln������) Then
        Exit Sub
    End If
    If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
        dbl��� = Val(txtEdit(idx_txt�����).Text)
    End If
    dbl��� = dbl��� + IIf(chk�Ƿ��ֵ.value = 0, 0, Val(txtEdit(idx_txt���γ�ֵ).Text))
    '���㿨���:
    txtEdit(idx_txt��ǰ���).Text = Format(IIf(mCardEditType = gEd_�޸�, 0, mCardInfor.dbl�����) + dbl���, "0.00")
End Sub
Private Function zlFromCardNOGetDataToCtrl(ByVal strCardNo As String, Optional blnSetBase As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ���,��ȡ��Ӧ�Ŀ�Ƭ��Ϣ���ؼ�
    '���:strCardNo-��ǰ����
    '     blnSetBase-���û�������Ϣ(��:������,�������,�Ƿ���ֵ��)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-14 16:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    mlngSel����ID = 0
    mCardInfor.bln�����ֵ = False
    gstrSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,a.��Ч��,a.����ԭ��, a.����," & _
    "          a.������,a.�쿨��,to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��, " & _
    "          a.������,to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬,a.��ע, " & _
    "          to_char(a.������," & gOraFmtString.FM_��� & ") as ������ ," & _
    "          to_char(a.���۽��," & gOraFmtString.FM_��� & ") as ���۽�� ," & _
    "          to_char(a.��ֵ�ۿ���," & gOraFmtString.FM_�ۿ��� & ") as ��ֵ�ۿ��� ," & _
    "          to_char(a.���," & gOraFmtString.FM_��� & ") as ��� ," & _
    "          a.ͣ����,to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.�쿨����ID,decode(b.����,NULL,'' ,b.����||'-'||b.����) AS �쿨����,a.������� " & _
    "   From ���ѿ�Ŀ¼ A ,���ű� b" & _
    "   Where A.���� = [1] and A.�ӿڱ��=[2] And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��) and a.�쿨����id=b.Id(+) " & _
    "   Order by a.���"
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng�ӿڱ��)

    If rsTemp.EOF Then

        If blnSetBase = True Then
        Else
            ShowMsgbox "δ�ҵ���ص����ѿ���¼,�����Ѿ�������ɾ���˸����ѿ�,����!"
        End If
        If mCardEditType = gEd_��ֵ Then
            Call ClearCtlData: mlng���ѿ�ID = 0
         End If
        Exit Function
    End If
    If Val(Nvl(rsTemp!�ɷ��ֵ)) <> 1 And mCardEditType = gEd_��ֵ Then
        ShowMsgbox "��ǰ����Ϊ:" & Nvl(rsTemp!����) & " ���ǳ�ֵ��,���ܽ��г�ֵ"
        Call ClearCtlData: mlng���ѿ�ID = 0
        Exit Function
    End If
    mblnNoClick = True
    mlng���ѿ�ID = Val(rsTemp!id)
    With cbo������
        .ListIndex = -1
        strTemp = Nvl(rsTemp!������): blnFind = False
        For i = 0 To .ListCount - 1
            If .List(i) & ";" Like "*." & strTemp & ";" Then
                blnFind = True
                .ListIndex = i: Exit For
            End If
        Next
        If blnFind = False Then
            .AddItem strTemp
            .ListIndex = .NewIndex
        End If
    End With

    With lvwType
        strTemp = Nvl(rsTemp!�������): blnFind = False
        For i = 1 To .ListItems.count
            If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
              .ListItems(i).Checked = True
            Else
              .ListItems(i).Checked = False
            End If
        Next
    End With

    chk�Ƿ��ֵ.value = IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, 1, 0)
    mlngSel����ID = Val(Nvl(rsTemp!id))
    If blnSetBase = True Then
        If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
             Call SetDefaultValue
        End If
        mblnNoClick = False: zlFromCardNOGetDataToCtrl = True
        Exit Function
    End If
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����): txtEdit(mtxtIdx.idx_txtȷ������).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����ԭ��).Text = Nvl(rsTemp!����ԭ��)
    txtEdit(mtxtIdx.idx_txt����Ч����).Text = Format(rsTemp!��Ч��, "yyyy-MM-DD")
    If txtEdit(mtxtIdx.idx_txt����Ч����).Text >= "3000-01-01" Then
        txtEdit(mtxtIdx.idx_txt����Ч����).Text = ""
    End If
    If txtEdit(mtxtIdx.idx_txt����Ч����).Text <> "" Then
        dtp����Ч����.value = CDate(txtEdit(mtxtIdx.idx_txt����Ч����).Text)
    Else
        dtp����Ч����.value = Empty
    End If
    txtEdit(mtxtIdx.idx_txt��ע).Text = Nvl(rsTemp!��ע)
    txtEdit(mtxtIdx.idx_txt������).Text = Nvl(rsTemp!������)
    txtEdit(mtxtIdx.idx_txt��������).Text = Nvl(rsTemp!����ʱ��)
    txtEdit(mtxtIdx.idx_txt�쿨��).Text = Nvl(rsTemp!�쿨��)
    txtEdit(mtxtIdx.idx_txt��ǰ���).Text = Nvl(rsTemp!���)
    txtEdit(mtxtIdx.idx_txt�����).Text = Nvl(rsTemp!������)
    txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text = Nvl(rsTemp!���۽��)
    txtEdit(mtxtIdx.idx_txtʵ�����۶�).Tag = Nvl(rsTemp!���۽��)

    txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = ""
    txtEdit(mtxtIdx.idx_txt��ֵ����).Text = Nvl(rsTemp!��ֵ�ۿ���)
    txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = ""
    txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = ""
    txtEdit(mtxtIdx.idx_txt�Ҳ�).Text = ""
    txtEdit(mtxtIdx.idx_txtʵ�պϼ�).Text = ""
    txtEdit(mtxtIdx.idx_txt���νɿ�).Text = ""
    txtEdit(mtxtIdx.idx_txt�쿨����).Text = Nvl(rsTemp!�쿨����)
    txtEdit(mtxtIdx.idx_txt�쿨����).Tag = Nvl(rsTemp!�쿨����ID)

    If mCardEditType = gEd_���� Or mCardEditType = gEd_���� Then
        txtEdit(mtxtIdx.idx_txt������).Text = UserInfo.����
        txtEdit(mtxtIdx.idx_txt����ʱ��).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    With mCardInfor
        .str���� = Nvl(rsTemp!����)
        .int�ɷ��ֵ = IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, 1, 0)
        .strЧ�� = txtEdit(mtxtIdx.idx_txt����Ч����).Text
        .dbl����� = Val(Nvl(rsTemp!���))
        .dblʵ������ = Val(Nvl(rsTemp!���۽��))
        .dbl����ֵ = Val(Nvl(rsTemp!������))
        .dbl�ۿ��� = Val(Nvl(rsTemp!��ֵ�ۿ���))
        .lng��ֵ���� = 0
    End With
    '77845:���ϴ�,2014/9/15,��������ʱ�������޸�ʱ������Ķ���ֵ���
    mCardInfor.bln�����ֵ = False
    If chk�Ƿ��ֵ.value = 1 Then
        If mCardEditType = gEd_��ֵ Then
            mCardInfor.bln�����ֵ = InStr(1, mstrPrivs, ";��ֵ;") > 0
            txtEdit(mtxtIdx.idx_txt��ֵ����).Text = ""
            Call SetDefaultValue(False)
        End If
    End If

    mblnNoClick = False
    zlFromCardNOGetDataToCtrl = True
    Exit Function
Errhand:
mblnNoClick = False
    If ErrCenter = 1 Then Resume
End Function

Private Function LoadDatatoCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݵ��ؼ�
    '����:���˺�
    '����:2009-12-09 15:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean

    Err = 0: On Error GoTo Errhand:
    If mCardEditType = gEd_���� Then
        Call ClearCtlData
        With mCardInfor
            .str���� = ""
            .int�ɷ��ֵ = 0
            .strЧ�� = "3000-01-01"
            .dbl����� = 0
            .dblʵ������ = 0
            .dbl����ֵ = 0
            .dbl�ۿ��� = 0
        End With
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        mCardInfor.bln�����ֵ = InStr(1, mstrPrivs, ";��ֵ;") > 0
        txtEdit(mtxtIdx.idx_txt������).Text = UserInfo.����: txtEdit(mtxtIdx.idx_txt��������).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        dtp����Ч����.Visible = True: txtEdit(mtxtIdx.idx_txt����Ч����).Visible = False
        dtp����Ч����.MinDate = CDate(txtEdit(mtxtIdx.idx_txt��������).Text)
        dtp����Ч����.value = DateAdd("m", 1, dtp����Ч����.MinDate)
        dtp����Ч����.value = Null
        Call Set�ɷ��ֵ
        LoadDatatoCard = True: Exit Function
    
    End If
    If mCardEditType = gEd_��ֵ And mlng���ѿ�ID = 0 Then
        '��ֵ,��û����ֵ,��ֱ���˳���
        Call ClearCtlData
        With mCardInfor
            .str���� = ""
            .int�ɷ��ֵ = 0
            .strЧ�� = "3000-01-01"
            .dbl����� = 0
            .dblʵ������ = 0
            .dbl����ֵ = 0
            .dbl�ۿ��� = 0
        End With
        dtp����Ч����.value = Null
        Call Set�ɷ��ֵ
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        LoadDatatoCard = True: Exit Function
    End If
    If mCardEditType = gEd_���� And mlng���ѿ�ID = 0 Then
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        LoadDatatoCard = True: Exit Function
         Exit Function
    End If
    '�鿴����������
    gstrSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,a.��Ч��,a.����ԭ��, a.����," & _
    "          a.������,a.�쿨��,to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��, " & _
    "          a.������,to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬,a.��ע, " & _
    "          to_char(a.������," & gOraFmtString.FM_��� & ") as ������ ," & _
    "          to_char(a.���۽��," & gOraFmtString.FM_��� & ") as ���۽�� ," & _
    "          to_char(a.��ֵ�ۿ���," & gOraFmtString.FM_�ۿ��� & ") as ��ֵ�ۿ��� ," & _
    "          to_char(a.���," & gOraFmtString.FM_��� & ") as ��� ," & _
    "          a.ͣ����,to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.�쿨����ID,decode(b.����,null,'',b.����||'-'||b.����) AS �쿨����,a.������� " & _
    "   From ���ѿ�Ŀ¼ A,���ű� B " & _
    "   Where   a.�쿨����id=b.Id(+) and A.Id =[1]   " & _
    "   Order by a.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ѿ�ID)

    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ���ص����ѿ���¼,�����Ѿ�������ɾ���˸����ѿ�,����!"
        Exit Function
    End If

    If mCardEditType = gEd_��ֵ And IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, 1, 0) <> 1 Then
        '����ǳ�ֵ,���ֲ�������ֵ��,����ʾ���˳�
        ShowMsgbox "ע��:" & vbCrLf & "    �ÿ���������ֵ�����ܽ��г�ֵ������"
        Exit Function
    End If

    mblnNoClick = True
    mlngSel����ID = mlng���ѿ�ID
    chk�Ƿ��ֵ.value = IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, 1, 0)

    With cbo������
        .ListIndex = -1
        strTemp = Nvl(rsTemp!������): blnFind = False
        For i = 0 To .ListCount - 1
            If .List(i) & ";" Like "*." & strTemp & ";" Then
                blnFind = True
                .ListIndex = i: Exit For
            End If
        Next
        If blnFind = False Then
            .AddItem strTemp
            .ListIndex = .NewIndex
        End If
    End With

    With lvwType
        strTemp = Nvl(rsTemp!�������): blnFind = False
        For i = 1 To .ListItems.count
            If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
                .ListItems(i).Checked = True
            Else
                .ListItems(i).Checked = False
            End If
        Next
    End With
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����).Tag = Nvl(rsTemp!id)
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����): txtEdit(mtxtIdx.idx_txtȷ������).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����ԭ��).Text = Nvl(rsTemp!����ԭ��)
    txtEdit(mtxtIdx.idx_txt����Ч����).Text = Format(rsTemp!��Ч��, "yyyy-MM-DD")
    If txtEdit(mtxtIdx.idx_txt����Ч����).Text >= "3000-01-01" Then
        txtEdit(mtxtIdx.idx_txt����Ч����).Text = ""
    End If
    If txtEdit(mtxtIdx.idx_txt����Ч����).Text <> "" Then
        dtp����Ч����.value = CDate(txtEdit(mtxtIdx.idx_txt����Ч����).Text)
    Else
        dtp����Ч����.value = Null
    End If
    txtEdit(mtxtIdx.idx_txt��ע).Text = Nvl(rsTemp!��ע)
    txtEdit(mtxtIdx.idx_txt������).Text = Nvl(rsTemp!������)
    txtEdit(mtxtIdx.idx_txt��������).Text = Nvl(rsTemp!����ʱ��)
    txtEdit(mtxtIdx.idx_txt�쿨��).Text = Nvl(rsTemp!�쿨��)
    txtEdit(mtxtIdx.idx_txt��ǰ���).Text = Format(Val(Nvl(rsTemp!���)), gVbFmtString.FM_���)
    txtEdit(mtxtIdx.idx_txt�����).Text = Format(Val(Nvl(rsTemp!������)), gVbFmtString.FM_���)
    txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text = Format(Val(Nvl(rsTemp!���۽��)), gVbFmtString.FM_���)
    txtEdit(mtxtIdx.idx_txtʵ�����۶�).Tag = Val(Nvl(rsTemp!���۽��))

    txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = ""
    txtEdit(mtxtIdx.idx_txt��ֵ����).Text = Format(Val(Nvl(rsTemp!��ֵ�ۿ���)), "0.00")
    txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = ""
    txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = ""
    txtEdit(mtxtIdx.idx_txt�Ҳ�).Text = ""
    txtEdit(mtxtIdx.idx_txtʵ�պϼ�).Text = ""
    txtEdit(mtxtIdx.idx_txt���νɿ�).Text = ""
    txtEdit(mtxtIdx.idx_txt�쿨����).Text = Nvl(rsTemp!�쿨����)
    txtEdit(mtxtIdx.idx_txt�쿨����).Tag = Nvl(rsTemp!�쿨����ID)

    If mCardEditType = gEd_���� Or mCardEditType = gEd_�˿� Then
        txtEdit(mtxtIdx.idx_txt������).Text = UserInfo.����
        txtEdit(mtxtIdx.idx_txt����ʱ��).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mtxtIdx.idx_txt������).Text = Nvl(rsTemp!������)
        txtEdit(mtxtIdx.idx_txt����ʱ��).Text = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
        If txtEdit(mtxtIdx.idx_txt����ʱ��).Text >= "3000-01-01" Then txtEdit(mtxtIdx.idx_txt����ʱ��).Text = ""
    End If
    With mCardInfor
        .str���� = Nvl(rsTemp!����)
        .int�ɷ��ֵ = IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, 1, 0)
        .strЧ�� = txtEdit(mtxtIdx.idx_txt����Ч����).Text
        .dbl����� = Val(Nvl(rsTemp!���))
        .dblʵ������ = Val(Nvl(rsTemp!���۽��))
        .dbl����ֵ = Val(Nvl(rsTemp!������))
        .dbl�ۿ��� = Val(Nvl(rsTemp!��ֵ�ۿ���))
        .bln�����ֵ = InStr(1, mstrPrivs, ";��ֵ;") > 0 And (mCardEditType = gEd_��ֵ Or mCardEditType = gEd_����)
        .lng��ֵ���� = 0
    End With
    '77845:���ϴ�,2014/9/15,��������ʱ�������޸�ʱ������Ķ���ֵ���
    If mCardEditType = gEd_���� Then
        InsertIntoGrid txtEdit(mtxtIdx.idx_txt����).Text, False, True
    End If

    Call SetDefaultValue(False)
    Call Set�ɷ��ֵ

    Call SetLblCatpion

    mblnNoClick = False
    LoadDatatoCard = True
    Exit Function
Errhand:
    mblnNoClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub SetFormNOTResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô��岻���������С
    '����:���˺�
    '����:2009-12-10 15:15:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Me.Width = FM_WIDTH: Me.Height = FM_HEIGHT + IIf(mCardEditType = gEd_��ֵ, 300, 0)
    Call zlSetWindowsBroldStyle(Me)  '���������óɲ��ɵ�
End Sub
Private Function InsertIntoGrid(ByVal strCardNo As String, Optional blnEndCard As Boolean = False, Optional blnNotCardNo As Boolean = False, Optional blnModifyCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ָ����������
    '���:strCardNo-����
    '     blnEndCard=true:���ڽ������ı���ˢ��,false:�ڿ�ʼ���ı���ˢ��
    '     blnNotCardNo-�Ƿ񲻶�ȡԭ���Ŀ�Ƭ��Ϣ(�Ի�����Ч)
    '     blnModifyCard-�Ƿ��޸ĸĺ�
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-12-14 11:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, strCurCard As String, strCards As String, varData As Variant, i As Long, lng���� As Long
    Dim strCardRange As String

    Err = 0: On Error GoTo Errhand:
    lng���� = 1

    If mCardEditType = gEd_���� Then
        '����ʱ,�������ݿ����Ϣ��ȷ��
        If Not blnNotCardNo Then
            If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        End If
    Else
        Call zlFromCardNOGetDataToCtrl(strCardNo, True)
    End If
    
    '92796:���ϴ�,2016/1/20,�������ѿ����ų���С���޶�����
    '�ȼ���������Ƿ�Ϸ�:
    If CheckInput = False Then Exit Function
    If zlCommFun.ActualLen(strCardNo) > txtEdit(mtxtIdx.idx_txt����).MaxLength Then
        ShowMsgbox "���ų��Ȳ�����ȷ,����!"
        Exit Function
    End If

    '�ȼ�����ݵĺϷ���
    With vsGrid
        For lngRow = 1 To .Rows - 1
            '�ȿ��Ƿ��Ѿ�����ˢ������Ϣ
            strCurCard = Trim(.Cell(flexcpData, lngRow, .ColIndex("����")))
            If InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 And lngRow <> .Row Then
                '�����Ѿ�����,��Ӧ���ٲ���
                ShowMsgbox "ע��:" & vbCrLf & "    �ڵ�" & lngRow & "����,�Ѿ����ڴ����ѿ���(����:" & strCardNo & "),�����ټ���!"
                Exit Function
            End If
            If .Row = lngRow Then

                If strCurCard = strCardNo Then
                    'ͬ�е�,�϶������ٲ�����
                    InsertIntoGrid = True: Exit Function
                End If
                If InStr(1, strCurCard, ",") > 0 And InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 Then
                    '�����Ѿ�����,��Ӧ���ٲ���
                    ShowMsgbox "ע��:" & vbCrLf & "    �ڵ�" & lngRow & "����,�Ѿ����ڴ����ѿ���(����:" & strCardNo & "),�����ټ���!"
                    Exit Function

                End If
                If blnEndCard Then
                    '�����������Ƿ���ȷ
                    If Split(.TextMatrix(lngRow, .ColIndex("����")) & "��", "��")(1) = strCardNo Then
                         InsertIntoGrid = True: Exit Function
                    End If

                End If
            End If
        Next
        '�������һ����Χ��ˢ��,Ҳ��Ҫ���
        If blnEndCard Then
            strCurCard = Trim(.Cell(flexcpData, .Row, .ColIndex("����")))
            If InStr(1, strCurCard, ",") <= 0 Then  '�������һ����Χ,�͵����µ�ˢ����¼����
                If strCurCard < strCardNo Then
                  strCardRange = strCurCard & "��" & strCardNo

                  If zlCardNoRange(strCardRange, strCards) = False Then Exit Function
                  varData = Split(strCards, ","): lng���� = UBound(varData) + 1
                  For i = 0 To UBound(varData)
                    '����Ƿ����ظ���
                    For lngRow = 1 To .Rows - 1
                        '�ȿ��Ƿ��Ѿ�����ˢ������Ϣ
                        strCurCard = Trim(.Cell(flexcpData, lngRow, .ColIndex("����")))

                        If InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 And lngRow <> .Row Then
                            '�����Ѿ�����,��Ӧ���ٲ���
                            ShowMsgbox "ע��:" & vbCrLf & "    �ڵ�" & lngRow & "����,�Ѿ����ڴ����ѿ���(����:" & strCardNo & "),�����ټ���!"
                            Exit Function
                        End If
                    Next
                  Next
                    '����Ƿ��Ѿ������˴��ڵĿ�Ƭ��Ϣ,������ڵĿ�Ƭ��Ϣ,��Ҫ���:
                    '1. �����Ͳ�һ�µ�,���ܷ�.
                    '2. �Ƿ��ֵ��һ�µ�,Ҳ���ܷ�
                    '3.������Щ���(����:�Ѿ����˿���,�����ٴη���)
                    If ��鿨���Ƿ�Ϸ�(strCards, True, True) = False Then Exit Function
                Else
                    '��ʼ�Ŵ��ڽ�����, ҲֻĬ��Ϊ�µĺ�
                    blnEndCard = False
                End If
            Else
                blnEndCard = False
            End If
        End If

        If blnEndCard = False Then
            '����Ƿ��Ѿ������˴��ڵĿ�Ƭ��Ϣ,������ڵĿ�Ƭ��Ϣ,��Ҫ���:
            '1. �����Ͳ�һ�µ�,���ܷ�.
            '2. �Ƿ��ֵ��һ�µ�,Ҳ���ܷ�
            '3.������Щ���(����:�Ѿ����˿���,�����ٴη���)
            If ��鿨���Ƿ�Ϸ�(strCardNo, True, True) = False Then Exit Function
        End If
        mblnNoClick = True
        '��������������
        ' ���ݽ��������õ�ֵ���д���

        If blnEndCard Then
            '��ԭ���Ļ����ϼ��Ϸ�Χ
            .TextMatrix(.Row, .ColIndex("����")) = strCardRange & "   (��" & lng���� & "�ſ�)"
            .Cell(flexcpData, .Row, .ColIndex("����")) = strCards
        Else
            If .TextMatrix(.Row, .ColIndex("����")) <> "" And blnModifyCard = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .TextMatrix(.Row, .ColIndex("����")) = strCardNo
            .Cell(flexcpData, .Row, .ColIndex("����")) = strCardNo
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        End If
        .TextMatrix(.Row, .ColIndex("ID")) = mlngSel����ID
        .TextMatrix(.Row, .ColIndex("������")) = Mid(cbo������.Text, InStr(cbo������.Text, ".") + 1)
        .TextMatrix(.Row, .ColIndex("�Ƿ��ֵ")) = IIf(chk�Ƿ��ֵ.value = 1, "��", ""): .Cell(flexcpData, .Row, .ColIndex("�Ƿ��ֵ")) = IIf(chk�Ƿ��ֵ.value = 1, 1, 0)
        .TextMatrix(.Row, .ColIndex("����")) = "******************": .Cell(flexcpData, .Row, .ColIndex("����")) = Trim(txtEdit(mtxtIdx.idx_txt����).Text)
        .Cell(flexcpData, .Row, .ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text)
        .TextMatrix(.Row, .ColIndex("����ԭ��")) = Trim(txtEdit(mtxtIdx.idx_txt����ԭ��).Text)
        .TextMatrix(.Row, .ColIndex("�쿨��")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Text)
        .TextMatrix(.Row, .ColIndex("�쿨����")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨����).Text): .Cell(flexcpData, .Row, .ColIndex("�쿨����")) = Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)

        .TextMatrix(.Row, .ColIndex("��ע")) = Trim(txtEdit(mtxtIdx.idx_txt��ע).Text)
        .TextMatrix(.Row, .ColIndex("������")) = Trim(txtEdit(mtxtIdx.idx_txt������).Text)
        If Trim(.TextMatrix(.Row, .ColIndex("������"))) = "" Then .TextMatrix(.Row, .ColIndex("������")) = UserInfo.����
        .TextMatrix(.Row, .ColIndex("��������")) = Trim(txtEdit(mtxtIdx.idx_txt��������).Text)

        .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
        .TextMatrix(.Row, .ColIndex("����Ч��")) = Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM")
        .TextMatrix(.Row, .ColIndex("�����")) = Trim(txtEdit(mtxtIdx.idx_txt�����).Text)
        .TextMatrix(.Row, .ColIndex("ʵ������")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text)

        If chk�Ƿ��ֵ.value = 0 Then
            .TextMatrix(.Row, .ColIndex("��ֵ����")) = ""
            .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = ""
            .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = ""
        Else
            .TextMatrix(.Row, .ColIndex("��ֵ����")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ����).Text)
            .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = Trim(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text)
            .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text)
        End If
        .TextMatrix(.Row, .ColIndex("�������")) = Get�������
        .TextMatrix(.Row, .ColIndex("��ֵ˵��")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text)
        .TextMatrix(.Row, .ColIndex("��ֵ�ɿ���")) = Trim(txtEdit(mtxtIdx.idx_txt�ɿ���).Text)
        .TextMatrix(.Row, .ColIndex("��������")) = lng����
    End With
    Call Show��������
    '����Ƿ��������������Ϣ
    Call CheckOtherCard
    Call Calcʵ�պϼ�


    mblnNoClick = False
    InsertIntoGrid = True
    Exit Function
Errhand:
    mblnNoClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Show��������()
    Dim lngRow As Long, lngNum As Long
    If mCardEditType <> gEd_���� And mCardEditType <> gEd_���� Then Exit Sub
    lngNum = 0
    With vsGrid
        For lngRow = 1 To .Rows - 1
            lngNum = lngNum + Val(.TextMatrix(lngRow, .ColIndex("��������")))
        Next
    End With
    stbThis.Panels(2).Text = "���ι�" & IIf(mCardEditType <> gEd_����, "��", "����") & lngNum & "�ſ�Ƭ"
End Sub
Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô��ڴ�С
    '����:���˺�
    '����:2010-01-04 16:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
  Select Case mCardEditType
    Case gEd_����
    Case gEd_�޸�
        Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    Case gEd_��ֵ
         Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    Case gEd_����
    Case gEd_ȡ������
        Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    Case gEd_����
        Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    Case gEd_�˿�, gEd_ȡ���˿�
        Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    Case gEd_��ѯ
        Call SetFormNOTResize   '���óɲ��ɵ���С
        stbThis.Visible = False
    End Select
End Sub

Private Sub SetEditProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ñ༭����
    '����:���˺�
    '����:2009-12-09 16:54:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    txtEdit(mtxtIdx.idx_txt������).Enabled = False: txtEdit(mtxtIdx.idx_txt��������).Enabled = False
    txtEdit(mtxtIdx.idx_txtʵ�պϼ�).Locked = True: txtEdit(mtxtIdx.idx_txt�Ҳ�).Locked = True
    txtEdit(mtxtIdx.idx_txt��ǰ���).Enabled = False
    txtEdit(mtxtIdx.idx_txt�ɿ���).Enabled = False
    txtEdit(mtxtIdx.idx_txt������).Enabled = False: txtEdit(mtxtIdx.idx_txt����ʱ��).Enabled = False
    cmdSel(mcmdIdx.idx_cmd����ԭ��).Visible = False
    cmdSel(mcmdIdx.idx_cmd�쿨��).Visible = False
    cmdSel(mcmdIdx.idx_cmd�쿨����).Visible = False

    Select Case mCardEditType
    Case gEd_����
        txtEdit(mtxtIdx.idx_txt�����).Enabled = zlCheckPrivs(mstrPrivs, "������Ŀ����")
        txtEdit(mtxtIdx.idx_txtʵ�����۶�).Enabled = txtEdit(mtxtIdx.idx_txt�����).Enabled
        cmdSel(mcmdIdx.idx_cmd����ԭ��).Visible = True
        cmdSel(mcmdIdx.idx_cmd�쿨��).Visible = True
        cmdSel(mcmdIdx.idx_cmd�쿨����).Visible = True
        If mCardInfor.bln�����ֵ Then GoTo DoSetColor:
        fra�ɿ�.Visible = txtEdit(mtxtIdx.idx_txt�����).Enabled
        If txtEdit(mtxtIdx.idx_txt�����).Enabled = False Then picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height
       ' txtEdit(mtxtIdx.idx_txt�����).Enabled = False: txtEdit(mtxtIdx.idx_txtʵ�����۶�).Enabled = False
        txtEdit(mtxtIdx.idx_txt���γ�ֵ).Enabled = False: txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Enabled = False
        txtEdit(mtxtIdx.idx_txt��ֵ����).Enabled = False:

    Case gEd_�޸�
        txtEdit(mtxtIdx.idx_txt�����).Enabled = False ' zlCheckPrivs(mstrPrivs, "������Ŀ����") And Not mCardInfor.bln������
        txtEdit(mtxtIdx.idx_txtʵ�����۶�).Enabled = txtEdit(mtxtIdx.idx_txt�����).Enabled
        txtEdit(mtxtIdx.idx_txt�쿨��).Enabled = True: txtEdit(mtxtIdx.idx_txt�쿨����).Enabled = True
        txtEdit(mtxtIdx.idx_txt����ԭ��).Enabled = True:
        txtEdit(mtxtIdx.idx_txt����).Enabled = False
        cmdSel(mcmdIdx.idx_cmd����ԭ��).Visible = True
        cmdSel(mcmdIdx.idx_cmd�쿨��).Visible = True
        cmdSel(mcmdIdx.idx_cmd�쿨����).Visible = True

        chk�Ƿ��ֵ.Enabled = (Not mCardInfor.bln������) And mCardInfor.lng��ֵ���� <= 1
        '77845:���ϴ�,2014/9/15,��������ʱ�������޸�ʱ������Ķ���ֵ���
         fra��ֵ���.Visible = False
         fra�ɿ�.Visible = txtEdit(mtxtIdx.idx_txt�����).Enabled:

        If txtEdit(mtxtIdx.idx_txt�����).Enabled = False Then
            picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height - 300
            Me.Height = FM_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height - 300
        Else
            fra�ɿ�.Top = fra��ֵ���.Top
            picCardInfor.Height = PIC_CARD_HEIGHT - fra��ֵ���.Height - 300
            Me.Height = FM_HEIGHT - fra��ֵ���.Height - 300
        End If
        txtEdit(mtxtIdx.idx_txt���γ�ֵ).Enabled = False: txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Enabled = False
        txtEdit(mtxtIdx.idx_txt��ֵ����).Enabled = False

    Case gEd_��ֵ
         For Each ctl In Controls
            Select Case UCase(TypeName(ctl))
            Case "TEXTBOX", "COMBOBOX"
                If Not ctl Is txtEdit(mtxtIdx.idx_txt����) Then
                    ctl.Enabled = False
                End If
            Case "CHECKBOX", "LISTVIEW", "LISTBOX"
                ctl.Enabled = False
            End Select
         Next
        txtEdit(mtxtIdx.idx_txt����).Enabled = True
        txtEdit(mtxtIdx.idx_txtʵ�պϼ�).Enabled = True: txtEdit(mtxtIdx.idx_txt�Ҳ�).Enabled = True
        txtEdit(mtxtIdx.idx_txt���γ�ֵ).Enabled = True: txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Enabled = True
        txtEdit(mtxtIdx.idx_txt��ֵ����).Enabled = True: dtp����Ч����.Visible = False
        txtEdit(mtxtIdx.idx_txt�ɿ���).Enabled = True: txtEdit(mtxtIdx.idx_txt��ֵ��ע).Enabled = True
        Call Set�ɷ��ֵ

    Case gEd_����
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp����Ч����.Visible = False
        fra�ɿ�.Visible = False
        fra��ֵ���.Visible = False
        If Me.WindowState = 0 Then Me.Height = FM_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height
        picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height
        
        txtEdit(mtxtIdx.idx_txt����).Enabled = True: txtEdit(mtxtIdx.idx_txt��ʼ����).Visible = False: txtEdit(mtxtIdx.idx_txt��������).Visible = False
        lbl��ʼ����.Visible = False: lbl��.Visible = False
        Call picList_Resize
    Case gEd_ȡ������
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        fra�ɿ�.Visible = False
        dtp����Ч����.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - 300
        Me.Height = FM_HEIGHT - fra�ɿ�.Height - 300

    Case gEd_����
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        fra�ɿ�.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - 300
        Me.Height = FM_HEIGHT - fra�ɿ�.Height - 300
    Case gEd_�˿�, gEd_ȡ���˿�
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp����Ч����.Visible = False
        fra�ɿ�.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - 300
        Me.Height = FM_HEIGHT - fra�ɿ�.Height - 300
    Case gEd_��ѯ
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp����Ч����.Visible = False
        fra�ɿ�.Visible = False
        fra��ֵ���.Visible = False

        picCardInfor.Height = PIC_CARD_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height - 300
        Me.Height = FM_HEIGHT - fra�ɿ�.Height - fra��ֵ���.Height - 300
    End Select
    Call picCard_Resize
DoSetColor:
    Call SetCtlBackColor
End Sub

Private Sub cboStyle_Click()
    '����֧Ʊʱ,��������ɿλ
    Dim blnEnabled As Boolean
    If cboStyle.ListIndex = -1 Then Exit Sub
    Call Set֧����ʽEnabled
End Sub

Private Sub Set֧����ʽEnabled()
     '����֧Ʊʱ,��������ɿλ
    Dim blnEnabled As Boolean
    If cboStyle.ListIndex = -1 Then Exit Sub

    blnEnabled = cboStyle.ItemData(cboStyle.ListIndex) = 2 And (cboStyle.Text Like "*Ʊ*" Or cboStyle.Text Like "*��*")
    txtEdit(mtxtIdx.idx_txt�������).Enabled = blnEnabled
    txtEdit(mtxtIdx.idx_txt������).Enabled = blnEnabled
    txtEdit(mtxtIdx.idx_txt�ʺ�).Enabled = blnEnabled
    If Not blnEnabled Then txtEdit(mtxtIdx.idx_txt�������).Text = "": txtEdit(mtxtIdx.idx_txt������).Text = "": txtEdit(mtxtIdx.idx_txt�ʺ�).Text = ""
    Call SetCtlBackColor
End Sub

Private Sub cbo������_Click()
    If mblnNoClick Then Exit Sub
    mblnChange = True
    '��������ȱʡֵ
    If Not (mCardEditType = gEd_���� Or mCardEditType = gEd_�޸�) Then Exit Sub
    If mCardEditType = gEd_�޸� And mCardInfor.bln�����ֵ = False Then Exit Sub
    If mCardEditType = gEd_�޸� And mCardInfor.bln������ = True Then Exit Sub
    Call SetDefaultValue(True)
    Calcʵ�պϼ�
End Sub
Private Sub Set�ɷ��ֵ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿɷ��ֵ
    '����:���˺�
    '����:2009-12-17 15:11:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    fra��ֵ���.Enabled = chk�Ƿ��ֵ.value = 1 And IIf(mCardEditType = gEd_����, InStr(1, mstrPrivs, ";��ֵ;") > 0, mCardInfor.bln�����ֵ)
    txtEdit(mtxtIdx.idx_txt���γ�ֵ).Enabled = fra��ֵ���.Enabled
    txtEdit(mtxtIdx.idx_txt��ֵ��ע).Enabled = fra��ֵ���.Enabled
    txtEdit(mtxtIdx.idx_txt��ֵ����).Enabled = fra��ֵ���.Enabled
    txtEdit(mtxtIdx.idx_txt���νɿ�).Enabled = fra��ֵ���.Enabled Or (Val(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text) <> 0 And (mCardEditType = gEd_���� Or mCardEditType = gEd_�޸�))
    txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Enabled = fra��ֵ���.Enabled
    cboStyle.Enabled = fra��ֵ���.Enabled
    Call Set֧����ʽEnabled
End Sub
Private Sub SetCtlBackColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��Ŀɱ�����ɫ
    '����:���˺�
    '����:2009-12-17 15:14:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Or UCase(TypeName(ctl)) = "COMBOBOX" Then
            Call zl_SetCtlBackColor(ctl)
        End If
    Next
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab

End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_���� Then Exit Sub

    '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
    With vsGrid
        If .TextMatrix(.Row, .ColIndex("����")) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("������")) = Mid(cbo������.Text, InStr(cbo������.Text, ".") + 1)
    End With
 End Sub

Private Sub chk�Ƿ��ֵ_Click()
    If mblnNoClick Then Exit Sub
    If mCardEditType <> gEd_���� And mCardEditType <> gEd_�޸� Then Exit Sub
    Call Set�ɷ��ֵ
    If mCardEditType <> gEd_���� Then
        Call Calc���:         Calcʵ�պϼ�
        Call SetLblCatpion
        Exit Sub
    End If
    '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
    With vsGrid

        If Split(.Cell(flexcpData, .Row, .ColIndex("����")) & ",", ",")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("�Ƿ��ֵ")) = IIf(chk�Ƿ��ֵ.value = 1, "��", ""): .Cell(flexcpData, .Row, .ColIndex("�Ƿ��ֵ")) = IIf(chk�Ƿ��ֵ.value = 1, 1, 0)
        If chk�Ƿ��ֵ.value = 0 Then
            .TextMatrix(.Row, .ColIndex("��ֵ����")) = ""
            .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = ""
            .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = ""
        Else
            .TextMatrix(.Row, .ColIndex("��ֵ����")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ����).Text)
            .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = Trim(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text)
            .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text)
        End If
        .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
    End With
    Call Calc���:         Calcʵ�պϼ�
    Call SetLblCatpion
End Sub

Private Sub chk�Ƿ��ֵ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cmdSel_Click(Index As Integer)
    Dim lngID As Long, str���� As String, str���� As String
    Select Case cmdSel(Index).Tag
    Case "�쿨��"
        'ѡ����Ա
        lngID = Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)
        If Select��Աѡ����(Me, txtEdit(mtxtIdx.idx_txt�쿨��), "", lngID, , True) = False Then
              Exit Sub
        End If
        If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
            '�쿨�˾��ǽɿ���
            txtEdit(mtxtIdx.idx_txt�ɿ���).Text = txtEdit(mtxtIdx.idx_txt�쿨��).Text
            txtEdit(mtxtIdx.idx_txt�ɿ���).Tag = txtEdit(mtxtIdx.idx_txt�쿨��).Tag
        End If
        '��Ҫ��ȡȱʡ����:
        If zl_From��Ա��ȡȱʡ����(Val(txtEdit(mtxtIdx.idx_txt�쿨��).Tag), str����, str����, lngID) Then
            txtEdit(mtxtIdx.idx_txt�쿨����).Text = str���� & "-" & str����
            txtEdit(mtxtIdx.idx_txt�쿨����).Tag = lngID
        End If
    Case "�쿨����"
        'ѡ��ȱʡ����
        lngID = Val(txtEdit(mtxtIdx.idx_txt�쿨��).Tag)
        If Select����ѡ����(Me, txtEdit(mtxtIdx.idx_txt�쿨����), "", "", IIf(lngID = 0, False, True), "", 0, "����ѡ����", , , , , lngID) = False Then
            Exit Sub
        End If
    Case "����ԭ��"
        If zl_SelectAndNotAddItem(Me, txtEdit(mtxtIdx.idx_txt����ԭ��), "", "���÷���ԭ��", "���÷���ԭ��ѡ��", True, True) = False Then
            Exit Sub
        End If
    Case Else
    End Select
End Sub

Private Sub dtp����Ч����_Change()
    If mCardEditType <> gEd_���� Then Exit Sub

    '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("����")) & "��", "��")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("����Ч��")) = Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM")
    End With
End Sub

Private Sub dtp����Ч����_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode <> vbKeyReturn Then Exit Sub
     zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp����Ч����_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_���� Then Exit Sub

    '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
    '84792:���ϴ�,2015/7/17,���������жϲ���ȷ
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("����")) & "��", "��")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("����Ч��")) = Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM")
    End With
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    If CheckDepented = False Then Unload Me: Exit Sub
    If LoadDatatoCard = False Then
        mlng���ѿ�ID = 0
        If mCardEditType <> gEd_��ֵ Then
            '�����ٴγ�ֵ
            Unload Me: Exit Sub
        End If
    End If
    '������ʽ���������ڴ���Load���֮��
    Call SetWindowsSize
    Call SetEditProperty
    If mCardEditType = gEd_��ֵ Then
         zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt���γ�ֵ)
    End If
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnUnLoad = False
    Call CreateObjectKeyboard
    Call InitModulePara
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitPanel
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitVsGrid
    
    Call Load֧����ʽ
    Me.Caption = Switch(mCardEditType = gEd_��ѯ, "���ѿ���Ϣ��ѯ", mCardEditType = gEd_�˿�, "���ѿ��˿�", mCardEditType = gEd_�޸�, "���ѿ���Ϣ�޸�", mCardEditType = gEd_��ֵ, "���ѿ���ֵ����", mCardEditType = gEd_����, "���ѿ�����", mCardEditType = gEd_����, "���ѿ����չ���", mCardEditType = gEd_����, "���ѿ�����", mCardEditType = gEd_ȡ������, "���ѿ�ȡ������", True, "���ѿ�����")
    mstrTitle = Me.Caption
    RaisEffect picCardInfor, -1
    '����65902,������:�������ѿ��޸�����ķ�ʽ
    If mCardEditType = gEd_�޸� Then
        txtEdit(1).Enabled = False
        txtEdit(2).Enabled = False
        txtEdit(2).Visible = False
        lblEdit(3).Visible = False
    Else
        txtEdit(1).Enabled = True
        txtEdit(2).Enabled = True
        txtEdit(2).Visible = True
        lblEdit(3).Visible = True
    End If
    If mCardEditType = gEd_���� Or mCardEditType = gEd_���� Then
        RestoreWinState Me, App.ProductName, mstrTitle
    End If
End Sub

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-12-10 11:39:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����")) = "1|0"
        .ColData(.ColIndex("��־")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("��ǰ���")) = "1|0"
        .Clear 1
        .Rows = 2
    End With
End Sub

Private Function zlMoveCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƴ���ǰ��Ƭ��Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow  As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_���� And mCardEditType <> gEd_���� Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            .Clear 1
            .Cell(flexcpData, 1, 0, .Rows - 1, .Cols - 1) = ""
            Call FromGridToCtlData
            Call SetDefaultValue
            Call CheckOtherCard
            zlMoveCard = True
            Exit Function
        End If
        lngCurRow = .Row
        .RemoveItem lngCurRow
        If lngCurRow < .Rows - 1 Then
            lngCurRow = lngCurRow + 1
        Else
            lngCurRow = .Rows - 1
        End If
        If lngCurRow < 1 Then lngCurRow = 1
        If lngCurRow > 1 Then .Row = lngCurRow
    End With
    Call Show��������
    Call CheckOtherCard
    Call Calcʵ�պϼ�
    zlMoveCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume

End Function

Private Function zlAppColumnData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ӧ����������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow  As Long, strTemp As String, strTempData As String, lngCurCol As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_���� Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            Exit Function
        End If
        lngCurCol = .Col
        If .ColIndex("����") = lngCurCol Then Exit Function
        strTemp = .TextMatrix(.Row, lngCurCol)
        strTempData = .Cell(flexcpData, .Row, lngCurCol)
        For lngRow = 1 To .Rows - 1
            If lngRow <> .Row Then
                .TextMatrix(lngRow, lngCurCol) = strTemp
                .Cell(flexcpData, lngRow, lngCurCol) = strTempData
            End If
        Next
    End With
    zlAppColumnData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlAppAllCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ӧ����������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow  As Long, strTemp As String, strTempData As String, lngCol As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_���� Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            Exit Function
        End If
        For lngRow = 1 To .Rows - 1
            If lngRow <> .Row Then
                For lngCol = 0 To .Cols - 1
                    Select Case lngCol
                    Case .ColIndex("����"), .ColIndex("��־"), .ColIndex("��������")

                    Case Else
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(.Row, lngCol)
                        .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, .Row, lngCol)
                    End Select
                Next
            End If
        Next
    End With
    zlAppAllCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function



'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case mconMenu_Edit_Affirm 'ȷ��
            Call SaveData
        Case conMenu_File_Print '��ӡ
        Case conMenu_Edit_MoveCard   '�Ƴ���Ƭ
             If zlMoveCard = False Then Exit Sub
         Case conMenu_Apply_AllCard     'Ӧ��������
            If zlAppAllCard = False Then Exit Sub
         Case conMenu_Apply_AllColumn     'Ӧ���ڴ���
            If zlAppColumnData = False Then Exit Sub
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub SaveData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2009-12-11 14:00:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng������� As Long, lngID As Long, rsTemp As ADODB.Recordset, txtTemp As TextBox
    If isValied = False Then Exit Sub
    
    '65048:������,2013-10-30,����ʱǿ�Ƹ���¼����Ϣ,������������
    For Each txtTemp In txtEdit
        If Not (txtTemp.Index = 12 Or txtTemp.Index = 13 Or txtTemp.Index = 14 Or txtTemp.Index = 15 Or txtTemp.Index = 16 Or txtTemp.Index = 17) Then _
        Call txtEdit_Validate(txtTemp.Index, False)
    Next
    
    If mCardEditType = gEd_���� Then
        lng������� = zlDatabase.GetNextId("���ѿ�Ŀ¼")
        If SavePayCard(lng�������) = False Then Exit Sub
        If mTy_MoudlePara.bln�ɿ��ӡ Then
            '��ӡ�ɿ
            '���ܴ���δ�ɿ�����
            If InStr(1, mstrPrivs, ";���ѿ��շ��վ�;") <> 0 Then
                'If Val(txtEdit(mtxtIdx.idx_txtʵ�պϼ�)) <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "�������=" & lng�������, "�ɿ�=" & Val(txtEdit(mtxtIdx.idx_txt���νɿ�).Text), "�Ҳ�=" & Val(txtEdit(mtxtIdx.idx_txt�Ҳ�).Tag), "��ֵID=0", "ReportFormat=1", 2)
            End If
        End If
        Call ClearCtlData(True)
        Call zl_CtlSetFocus(Me.txtEdit(mtxtIdx.idx_txt����))
        mblnChange = False: mintSucces = mintSucces + 1
        Exit Sub
    End If

    If mCardEditType = gEd_�޸� Then
        '�޸Ĵ���
        If SaveModifyCard = False Then Exit Sub
        '��ӡ�ɿ
        '���ܴ���δ�ɿ�����
        If Val(txtEdit(mtxtIdx.idx_txtʵ�պϼ�)) <> 0 And mCardInfor.bln�����ֵ And mTy_MoudlePara.bln�ɿ��ӡ Then
            '���ܴ���δ�ɿ�����
            If InStr(1, mstrPrivs, ";���ѿ��շ��վ�;") <> 0 Then
                gstrSQL = "Select ������� From ���ѿ�Ŀ¼ where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ѿ�ID)
                If rsTemp.EOF = False Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "�������=" & Val(Nvl(rsTemp!�������)), "�ɿ�=" & Val(txtEdit(mtxtIdx.idx_txt���νɿ�).Text), "�Ҳ�=" & Val(txtEdit(mtxtIdx.idx_txt�Ҳ�).Text), "��ֵID=0", "ReportFormat=1", 2)
                End If
            End If

         '   Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "���ѿ�ID=" & mlng���ѿ�ID, 2)
        End If
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If mCardEditType = gEd_���� Then
        '���մ���
        If SaveCallBack = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_ȡ������ Then
        '���մ���
        If SaveCallBack(True) = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_�˿� Then
       '�˿��������
       If SaveBackCard(False) = False Then Exit Sub
        mblnChange = False: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_ȡ���˿� Then
       '�˿��������
       If SaveBackCard(True) = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If mCardEditType = gEd_��ֵ Then
        If mlng���ѿ�ID = 0 Then
            ShowMsgbox "ע��:" & vbCrLf & "    ����ָ�������ѿ�,���ܳ�ֵ!"
            Exit Sub
        End If

        If SaveInFull(lngID) = False Then Exit Sub
        '��ӡ�ɿ
        '���ܴ���δ�ɿ�����
        If Val(txtEdit(mtxtIdx.idx_txtʵ�պϼ�)) <> 0 And mCardInfor.bln�����ֵ And mTy_MoudlePara.bln�ɿ��ӡ Then
            If InStr(1, mstrPrivs, ";���ѿ��շ��վ�;") <> 0 Then
                'If Val(txtEdit(mtxtIdx.idx_txtʵ�պϼ�)) <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "��ֵID=" & lngID, "�ɿ�=" & Val(txtEdit(mtxtIdx.idx_txt���νɿ�).Text), "�Ҳ�=" & Val(txtEdit(mtxtIdx.idx_txt�Ҳ�).Text), "�������=0", "ReportFormat=2", 2)
            End If
        End If
        mblnChange = False:: mintSucces = mintSucces + 1
        If IIf(Val(zlDatabase.GetPara("������ֵ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            Call ClearCtlData(True): mlng���ѿ�ID = 0
            Call zl_CtlSetFocus(Me.txtEdit(mtxtIdx.idx_txt����))
            Call Set�ɷ��ֵ:             Calc���: Calcʵ�պϼ�
            Call SetEditProperty
        Else
            Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function SaveInFull(ByRef lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ֵ����
    '����:lngID-���ر��εĳ�ֵ��ID
    '����:��ֵ�ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-12-14 10:51:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��� As Long
    Err = 0: On Error GoTo Errhand:
    
    '61905:������,2013-10-29,֧Ʊ��ֵ��д��Ϣ�������
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt������).Text)) > 50 Then
        ShowMsgbox "ע��:" & vbCrLf & "    �����г��ȳ���,��������д!"
        txtEdit(mtxtIdx.idx_txt������).SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt�ʺ�).Text)) > 20 Then
        ShowMsgbox "ע��:" & vbCrLf & "    �ʺų��ȳ���,��������д!"
        txtEdit(mtxtIdx.idx_txt�ʺ�).SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt�������).Text)) > 30 Then
        ShowMsgbox "ע��:" & vbCrLf & "    ������볤�ȳ���,��������д!"
        txtEdit(mtxtIdx.idx_txt�������).SetFocus
        Exit Function
    End If
    '78084:���ϴ�,2014/9/18,��ֵ����ж�
    If Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text) = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "    ���γ�ֵ�Ľ��Ϊ��,��������д!"
        If txtEdit(mtxtIdx.idx_txt���γ�ֵ).Visible And txtEdit(mtxtIdx.idx_txt���γ�ֵ).Enabled Then txtEdit(mtxtIdx.idx_txt���γ�ֵ).SetFocus
        Exit Function
    End If
    
    lngID = zlDatabase.GetNextId("���ѿ���ֵ��¼")
    lng��� = GetMax��ֵ���(mlng���ѿ�ID)
    'Zl_���ѿ���ֵ��¼_Insert
    gstrSQL = "Zl_���ѿ���ֵ��¼_Insert("
    '  Id_In         In ���ѿ���ֵ��¼.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  ���ѿ�id_In   In ���ѿ���ֵ��¼.���ѿ�id%Type,
    gstrSQL = gstrSQL & "" & mlng���ѿ�ID & ","
    '  ���_In       In ���ѿ���ֵ��¼.���%Type,
    gstrSQL = gstrSQL & "" & lng��� & ","
    '  ��ֵ���_In   In ���ѿ���ֵ��¼.��ֵ���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text), 4) & ","
    '  ��ֵ�ۿ�_In   In ���ѿ���ֵ��¼.��ֵ�ۿ�%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt��ֵ����).Text), 4) & ","
    '  �ɿ���_In   In ���ѿ���ֵ��¼.�ɿ���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text), 4) & ","
    '  ��ֵʱ��_In   In ���ѿ���ֵ��¼.��ֵʱ��%Type,
    gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����Ա����_In In ���ѿ���ֵ��¼.����Ա����%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
    '  �ɿ���_In     In ���ѿ���ֵ��¼.�ɿ���%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt�ɿ���).Text) & "',"
    '  ��ע_In       In ���ѿ���ֵ��¼.��ע%Type
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text) & "',"
     '  ���㷽ʽ_IN  in ���ѿ���ֵ��¼.���㷽ʽ%Type
    gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & cboStyle.Text & "'", "NULL") & ","
    '  ������_IN  in ���ѿ���ֵ��¼.��λ������%Type
    gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt������).Text) & "'", "NULL") & ","
    
   '  �ʺ�_IN  in ���ѿ���ֵ��¼.��λ�ʺ�%Type
    gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt�ʺ�).Text) & "'", "NULL") & ","
    
   '  �������_IN  in ���ѿ���ֵ��¼.��λ�������%Type
    gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt�������).Text) & "'", "NULL") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveInFull = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function GetMax��ֵ���(ByVal lng���ѿ�ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ĳ�ֵ���
    '����:���˺�
    '����:2009-12-14 11:06:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select nvl( Max(���),0)+1 as ��ֵ��� From ���ѿ���ֵ��¼ where ���ѿ�ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ĳ�ֵ���", lng���ѿ�ID)
    GetMax��ֵ��� = Val(Nvl(rsTemp!��ֵ���))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveCallBack(Optional blnCancelCallBack As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���մ���
    '���:blnCancelCallBack-ȡ������
    '����:���˺�
    '����:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection, lngRow As Long, strIDs As String, blnHaveData As Boolean
    
    Err = 0: On Error GoTo Errhand:
    If blnCancelCallBack Then
        'Zl_���ѿ�Ŀ¼_Callback
        gstrSQL = "Zl_���ѿ�Ŀ¼_Callback("
        '  Ids_In       IN varchar2,
        gstrSQL = gstrSQL & "" & mlng���ѿ�ID & ","
        '  ������_In   In ���ѿ�Ŀ¼.������%Type,
        gstrSQL = gstrSQL & IIf(blnCancelCallBack = True, "NULL", "'" & UserInfo.���� & "'") & ","
        '  ����ʱ��_In In ���ѿ�Ŀ¼.����ʱ��%Type
        gstrSQL = gstrSQL & "NULL)"
        AddArray cllPro, gstrSQL
    Else
        '���ܴ����������ղ���,���Ҫ����ID�ļ���
        strIDs = ""
        blnHaveData = False
        With vsGrid
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("ID"))) <> 0 Then
                    If zlCommFun.ActualLen(strIDs) > 4000 Then
                        strIDs = Mid(2, strIDs)
                        'Zl_���ѿ�Ŀ¼_Callback
                        gstrSQL = "Zl_���ѿ�Ŀ¼_Callback("
                        '  Ids_In     IN varchar2,
                        gstrSQL = gstrSQL & "'" & strIDs & "',"
                        '  ������_In   In ���ѿ�Ŀ¼.������%Type,
                        gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                        '  ����ʱ��_In In ���ѿ�Ŀ¼.����ʱ��%Type
                        gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
                        AddArray cllPro, gstrSQL
                        strIDs = ""
                    End If
                    strIDs = strIDs & "," & Val(.TextMatrix(lngRow, .ColIndex("ID")))
                    blnHaveData = True
                End If
            Next
            If strIDs <> "" Then
                strIDs = Mid(strIDs, 2)
                'Zl_���ѿ�Ŀ¼_Callback
                gstrSQL = "Zl_���ѿ�Ŀ¼_Callback("
                '  Ids_In       IN varchar2,
                gstrSQL = gstrSQL & "'" & strIDs & "',"
                '  ������_In   In ���ѿ�Ŀ¼.������%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                '  ����ʱ��_In In ���ѿ�Ŀ¼.����ʱ��%Type
                gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
                AddArray cllPro, gstrSQL
            End If
            If blnHaveData = False Then
                ShowMsgbox "ע��:" & vbCrLf & "    ��û��ˢҪ���յ����ѿ�,����!"
                zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
                Exit Function
            End If
        End With
    End If
    
Err = 0: On Error GoTo Errhand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveCallBack = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Function SaveBackCard(Optional blnCancelBackCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���մ���
    '���:blnCancelBackCard-ȡ���˿�
    '����:���˺�
    '����:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Err = 0: On Error GoTo Errhand:
    'Zl_���ѿ�Ŀ¼_Backcard
    gstrSQL = "Zl_���ѿ�Ŀ¼_Backcard("
    '  Id_In       In ���ѿ�Ŀ¼.ID%Type,
    gstrSQL = gstrSQL & "" & mlng���ѿ�ID & ","
    '  ������_In   In ���ѿ�Ŀ¼.������%Type,
    gstrSQL = gstrSQL & IIf(blnCancelBackCard = True, "NULL", "'" & UserInfo.���� & "'") & ","
    '  ����ʱ��_In In ���ѿ�Ŀ¼.����ʱ��%Type
    gstrSQL = gstrSQL & IIf(blnCancelBackCard = True, "NULL", "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveBackCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


Private Function Get�������() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:���˺�
    '����:2009-12-11 15:36:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String, i As Long
    With lvwType
         For i = 1 To .ListItems.count
            If .ListItems.Item(i).Checked Then
                strType = strType & "," & Mid(.ListItems(i).Key, 2)
            End If
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get������� = strType
End Function

Private Function SaveModifyCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���濨Ƭ�޸���Ϣ
    '����:�޸ĳɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-12-11 14:28:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    'Zl_���ѿ�Ŀ¼_Update
    gstrSQL = "Zl_���ѿ�Ŀ¼_Update("
    '  Id_In         In ���ѿ�Ŀ¼.ID%Type,
    gstrSQL = gstrSQL & "" & mlng���ѿ�ID & ","
    '  ����_In       In ���ѿ�Ŀ¼.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt����).Text) & "',"
    '  ������_In     In ���ѿ�Ŀ¼.������%Type,
    gstrSQL = gstrSQL & "" & zl_FromComboxGetData(cbo������) & ","
    '  �������_In   In ���ѿ�Ŀ¼.�������%Type,
    gstrSQL = gstrSQL & "'" & Get������� & "',"
    '  �ɷ��ֵ_In   In ���ѿ�Ŀ¼.�ɷ��ֵ%Type,
    gstrSQL = gstrSQL & "" & IIf(chk�Ƿ��ֵ.value = 1, 1, 0) & ","
    '  ��Ч��_In     In ���ѿ�Ŀ¼.��Ч��%Type,
        If IsNull(dtp����Ч����.value) Then
            gstrSQL = gstrSQL & "NULL,"
        Else
            gstrSQL = gstrSQL & "to_date('" & Format(dtp����Ч����.value, "yyyy-mm-dd HH:MM") & "','yyyy-mm-dd hh24:mi'),"
        End If
    '  ����ԭ��_In   In ���ѿ�Ŀ¼.����ԭ��%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt����ԭ��).Text) & "',"
    '  ������_In     In ���ѿ�Ŀ¼.������%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt������).Text) & "',"
    '  �쿨��_In     In ���ѿ�Ŀ¼.�쿨��%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Text) & "',"
    '  �쿨����id_In In ���ѿ�Ŀ¼.�쿨����id%Type,
    gstrSQL = gstrSQL & "" & IIf(Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag) = 0, "NULL", Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)) & ","
    '  ����ʱ��_In   In ���ѿ�Ŀ¼.����ʱ��%Type,
    gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ��ע_In       In ���ѿ�Ŀ¼.��ע%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt��ע).Text) & "',"
    '  ���㷽ʽ_In   In ���ѿ�Ŀ¼.���㷽ʽ%Type,
    gstrSQL = gstrSQL & "'" & Trim(cboStyle.Text) & "',"
    '  ������_In   In ���ѿ�Ŀ¼.������%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt�����).Text), 4) & ","
    '  ���۽��_In   In ���ѿ�Ŀ¼.���۽��%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text), 4) & ","
    '  ��ֵ�ۿ���_In In ���ѿ�Ŀ¼.��ֵ�ۿ���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt��ֵ����).Text), 4) & ","
    '  ���_In       In ���ѿ�Ŀ¼.���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt��ǰ���).Text), 4) & ","
    '  ��ֵ���_In   In ���ѿ���ֵ��¼.��ֵ���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text), 4) & ","
    '  �ɿ���_In   In ���ѿ���ֵ��¼.�ɿ���%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text), 4) & ","
    '  ��ֵ˵��_In   In ���ѿ���ֵ��¼.��ע%Type
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text) & "',"
    ' n_���½�� Number:=0 --n_���½��:1-��Ҫ������ֵ����ֵ����:0-ֻ���¸�����Ϣ
    gstrSQL = gstrSQL & IIf(mCardInfor.bln�����ֵ, "1", "0") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    SaveModifyCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case mconMenu_Edit_Affirm 'ȷ��
        Control.Enabled = Not (mCardEditType = gEd_��ѯ)
        Control.Visible = Not (mCardEditType = gEd_��ѯ)
    
    Case conMenu_File_Print '��ӡ
    Case conMenu_Edit_MoveCard   '�Ƴ���Ƭ
        If mCardEditType <> gEd_���� And mCardEditType <> gEd_���� Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        Control.Enabled = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("����")) <> "" Or vsGrid.Rows > 2
    Case conMenu_Apply_AllCard     'Ӧ��������
        If mCardEditType <> gEd_���� Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        With vsGrid
            Control.Caption = "Ӧ����������Ƭ"
            Control.ToolTipText = "������Ϊ��" & .TextMatrix(.Row, .ColIndex("����")) & "������Ϣ Ӧ������������Ϣ"
            Control.Enabled = mblnHaveOtherCard 'ֻ�д�����������ʱ���Ż���ִ�����Ϣ
        End With
    Case conMenu_Apply_AllColumn     'Ӧ���ڴ���
        If mCardEditType <> gEd_���� Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        With vsGrid
            Select Case .Col
            Case .ColIndex("����"), .ColIndex("��������")
                 Control.Enabled = False: Exit Sub
            Case Else
                Control.Caption = "Ӧ���ڡ�" & Trim(.TextMatrix(0, .Col)) & "����"
                Control.ToolTipText = "����" & .TextMatrix(.Row, .Col) & "�� Ӧ���ڡ�" & Trim(.TextMatrix(0, .Col)) & "���е�������Ƭ��Ϣ"
            End Select
            Control.Enabled = mblnHaveOtherCard 'ֻ�д�����������ʱ���Ż���ִ�����Ϣ
        End With

    End Select
End Sub

Private Sub CheckOtherCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����������Ƭ��Ϣ
    '����:���˺�
    '����:2009-12-18 09:47:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mblnHaveOtherCard = False
    With vsGrid
        For i = 1 To .Rows - 1
            If i <> .Row And .TextMatrix(i, .ColIndex("����")) <> "" Then
                mblnHaveOtherCard = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Cards     '������������
        Item.Handle = picList
    Case mPaneID.Pane_CardInfor    '��ϸ����Ϣ
        Item.Handle = picCard.hWnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mCardEditType = gEd_���� Or mCardEditType = gEd_���� Then
        SaveWinState Me, App.ProductName, mstrTitle
    End If
'    If mCardEditType = gEd_���� Then
'        zlSaveDockPanceToReg Me, dkpMan, "����-����"
'    ElseIf mCardEditType = gEd_���� Then
'        zlSaveDockPanceToReg Me, dkpMan, "����-����"
'    End If
    If mblnChange = False Then Exit Sub
    If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lvwType_Click()
    mblnChange = True
End Sub

'65048:������,2013-10-30,����ʱ�������δ���������
Private Sub lvwType_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mCardEditType <> gEd_���� Then Exit Sub
    '84792:���ϴ�,2015/7/17,���������жϲ���ȷ
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("����")) & "��", "��")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("�������")) = Get�������
    End With
End Sub

Private Sub lvwType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
        If fra�ɿ�.Visible = False Then
            If MsgBox("���Ƿ���Ҫ������ص�" & IIf(mCardEditType = gEd_�޸�, "�޸�", "����") & "������?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            End If
            Exit Sub
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwType_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_���� Then Exit Sub
    '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
    '84792:���ϴ�,2015/7/17,���������жϲ���ȷ
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("����")) & "��", "��")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("�������")) = Get�������
    End With
End Sub

Private Sub mobjBrushCard_zlBrushCarding(ByVal strCardNo As String)
    'ˢ������
    If strCardNo = "" Then Exit Sub
    Call zlBrusCard(strCardNo, True)
        mobjBrushCard.zlSetAutoBrush False
End Sub

Private Sub picCard_Resize()
    Err = 0: On Error Resume Next
    Dim sngTop As Single, sngLeft As Long
    With picCard
        sngLeft = (.ScaleLeft + .ScaleWidth - picCardInfor.Width) \ 2
        sngLeft = IIf(sngLeft < 0, 0, sngLeft)
        sngTop = (.ScaleTop + .ScaleHeight - picCardInfor.Height) \ 2
        sngTop = IIf(sngTop < 0, 0, sngTop)
        picCardInfor.Move sngLeft, sngTop
    End With
End Sub

Private Sub picList_Resize()
    Dim sngWidth As Single
    Err = 0: On Error Resume Next
    With picList
        If .ScaleWidth - lbl��.Width - txtEdit(mtxtIdx.idx_txt��������).Width - txtEdit(mtxtIdx.idx_txt��ʼ����).Left - txtEdit(mtxtIdx.idx_txt��ʼ����).Width - 120 > 0 Then
            txtEdit(mtxtIdx.idx_txt��������).Top = txtEdit(mtxtIdx.idx_txt��ʼ����).Top
            lbl��.Top = lbl��ʼ����.Top
            lbl��.Left = txtEdit(mtxtIdx.idx_txt��ʼ����).Left + txtEdit(mtxtIdx.idx_txt��ʼ����).Width + 100
            txtEdit(mtxtIdx.idx_txt��������).Left = lbl��.Left + lbl��.Width + 20
        Else
            txtEdit(mtxtIdx.idx_txt��������).Left = txtEdit(mtxtIdx.idx_txt��ʼ����).Left
            lbl��.Left = lbl��ʼ����.Left

            txtEdit(mtxtIdx.idx_txt��������).Top = txtEdit(mtxtIdx.idx_txt��ʼ����).Top + txtEdit(mtxtIdx.idx_txt��ʼ����).Height + 50
            lbl��.Top = txtEdit(mtxtIdx.idx_txt��������).Top + (txtEdit(mtxtIdx.idx_txt��������).Height - lbl��.Height) \ 2
        End If
    
'        If .Width < 4350 Then
'            lbl��.Caption = "��������": lbl��.FontBold = False
'            txtEdit(mtxtIdx.idx_txt��������).Text
'            lbl��.Top = txtEdit(mtxtIdx.idx_txt��ʼ����).Top + txtEdit(mtxtIdx.idx_txt��ʼ����).Height + 50
'
'        Else
'            sngWidth = .ScaleWidth - (lbl����.Left + lbl����.Width) - (lbl��.Width * 3)
'            sngWidth = IIf(sngWidth < 3360, 3360, sngWidth)
'            sngWidth = sngWidth \ 2
'            txtEdit(mtxtIdx.idx_txt��ʼ����).Width = sngWidth
'            lbl��.Left = txtEdit(mtxtIdx.idx_txt��ʼ����).Left + txtEdit(mtxtIdx.idx_txt��ʼ����).Width + lbl��.Width
'            txtEdit(mtxtIdx.idx_txt��������).Left = lbl��.Left + (lbl��.Width * 2)
'            txtEdit(mtxtIdx.idx_txt��������).Width = sngWidth
'        End If
        vsGrid.Top = IIf(txtEdit(mtxtIdx.idx_txt��������).Visible = False, 0, txtEdit(mtxtIdx.idx_txt��������).Top + txtEdit(mtxtIdx.idx_txt��������).Height) + 50
        vsGrid.Left = .ScaleLeft
        vsGrid.Width = .ScaleWidth
        vsGrid.Height = .ScaleHeight - vsGrid.Top
    End With
End Sub
Private Function zlCardNoRange(ByVal strCardNoRange As String, ByRef strCardNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ���Ŀ��ŷ�Χ���ֽ����صĿ���
    '���:strCardNoRange-���ŷ�Χ
    '����:strCardNos-���ؿ�����(�ö��ŷ���)
    '����:
    '����:���˺�
    '����:2009-12-10 16:30:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    Dim strCardStartNO As String, strCardEndNO As String, strCurNo As String
    Dim lngCount As Long


    varData = Split(strCardNoRange & "��", "��")
    strCardStartNO = varData(0): strCardEndNO = varData(1)
    If strCardEndNO = "" Then strCardNos = strCardStartNO: GoTo GoExit:

'    If strSartCardText = mTy_MoudlePara.str����ǰ׺ Then
'        strCardStartNO = Mid(strCardStartNO, Len(mTy_MoudlePara.str����ǰ׺) + 1)
'    End If
'    If strEndCardText = mTy_MoudlePara.str����ǰ׺ Then
'        strCardEndNO = Mid(strCardEndNO, Len(mTy_MoudlePara.str����ǰ׺) + 1)
'    End If
    If strCardStartNO > strCardEndNO Then
        Exit Function
    End If

    strCurNo = strCardStartNO
    strCardNos = strCardStartNO         'mTy_MoudlePara.str����ǰ׺ &
    lngCount = 0
    Do While True
        If strCurNo >= strCardEndNO Then
            Exit Do
        End If
        strCurNo = zlCommFun.IncStr(strCurNo)
        strCardNos = strCardNos & "," & strCurNo  'mTy_MoudlePara.str����ǰ׺ &
        lngCount = lngCount + 1
        If lngCount > 1000 Then Exit Do
Loop
    If lngCount > 1000 Then
        MsgBox "ע�⿨�ŷ�Χ���ܴ���1000�����ܼ�������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
GoExit:
    zlCardNoRange = True: Exit Function
End Function

Private Function zl���ѿ�InsertSQL(ByVal lng������� As Long, ByVal strCardNo As String, ByVal str����ʱ�� As String, ByVal lngRow As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����SQL���
    '���:lng�������-��Ҫ�Ǳ���һ������ʱ�ķ������,�Ա��ӡ
    '����:
    '����:
    '����:���˺�
    '����:2009-12-11 09:21:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        ' Zl_���ѿ�Ŀ¼_Insert
        gstrSQL = "Zl_���ѿ�Ŀ¼_Insert("
        '   �ӿڱ��_IN      IN ���ѿ�Ŀ¼.�ӿڱ��%Type,

        gstrSQL = gstrSQL & "'" & mlng�ӿڱ�� & "',"
        '  ����_In       In Varchar2, --��,����
        gstrSQL = gstrSQL & "'" & strCardNo & "',"
        '  ������_In     In ���ѿ�Ŀ¼.������%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "',"
        '  ����_In       In ���ѿ�Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & zlCommFun.zlStringEncode(Trim(.Cell(flexcpData, lngRow, .ColIndex("����")))) & "',"
        '  �������_In   In ���ѿ�Ŀ¼.�������%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("�������"))) & "',"
        '  �ɷ��ֵ_In   In ���ѿ�Ŀ¼.�ɷ��ֵ%Type,
        gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("�Ƿ��ֵ"))) & ","
        '  ��Ч��_In     In ���ѿ�Ŀ¼.��Ч��%Type,
        If Trim(.TextMatrix(lngRow, .ColIndex("����Ч��"))) = "" Then
            gstrSQL = gstrSQL & "NULL,"
        Else
            gstrSQL = gstrSQL & "to_date('" & Trim(.TextMatrix(lngRow, .ColIndex("����Ч��"))) & "','yyyy-mm-dd hh24:mi'),"
        End If
        '  ����ԭ��_In   In ���ѿ�Ŀ¼.����ԭ��%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("����ԭ��"))) & "',"
        '  ������_In     In ���ѿ�Ŀ¼.������%Type,
        gstrSQL = gstrSQL & "'" & IIf(Trim(.TextMatrix(lngRow, .ColIndex("������"))) = "", UserInfo.����, Trim(.TextMatrix(lngRow, .ColIndex("������")))) & "',"
        '  �쿨��_In     In ���ѿ�Ŀ¼.�쿨��%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("�쿨��"))) & "',"
        '  �쿨����id_In In ���ѿ�Ŀ¼.�쿨����id%Type,
        gstrSQL = gstrSQL & "" & IIf(Val(.Cell(flexcpData, lngRow, .ColIndex("�쿨����"))) = 0, "NULL", Val(.Cell(flexcpData, lngRow, .ColIndex("�쿨����")))) & ","
        '  ����ʱ��_In   In ���ѿ�Ŀ¼.����ʱ��%Type,
        gstrSQL = gstrSQL & "to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ��ע_In       In ���ѿ�Ŀ¼.��ע%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("��ע"))) & "',"
        '  ������_In   In ���ѿ�Ŀ¼.������%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("�����"))), 4) & ","
        '  ���۽��_In   In ���ѿ�Ŀ¼.���۽��%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("ʵ������"))), 4) & ","
        '  ��ֵ�ۿ���_In In ���ѿ�Ŀ¼.��ֵ�ۿ���%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("��ֵ����"))) * IIf(chk�Ƿ��ֵ.value = 1, 1, 0), 4) & ","
        '  ���_In       In ���ѿ�Ŀ¼.���%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("��ǰ���"))), 4) & ","
        '    �������_IN   IN ���ѿ�Ŀ¼.�������%type,
        gstrSQL = gstrSQL & "" & lng������� & ","
        '  ��ֵ���_In   In ���ѿ���ֵ��¼.��ֵ���%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("���γ�ֵ"))) * IIf(chk�Ƿ��ֵ.value = 1, 1, 0), 4) & ","
        '  �ɿ���_In   In ���ѿ���ֵ��¼.�ɿ���%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("ʵ�ʳ�ֵ�ɿ�"))) * IIf(chk�Ƿ��ֵ.value = 1, 1, 0), 4) & ","
        '  ��ֵ˵��_In   In ���ѿ���ֵ��¼.��ע%Type
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("��ֵ˵��"))) & "',"
        ' ���㷽ʽIn   In ���ѿ���ֵ��¼.���㷽ʽ%Type
        gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("���㷽ʽ"))) & "'", "NULL") & ","
         
        '  ������_IN  in ���ѿ���ֵ��¼.��λ������%Type
        gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "'", "NULL") & ","

        '  �ʺ�_IN  in ���ѿ���ֵ��¼.��λ�ʺ�%Type
        gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("�ʺ�"))) & "'", "NULL") & ","

        '  �������_IN  in ���ѿ���ֵ��¼.��λ�������%Type
        gstrSQL = gstrSQL & IIf(chk�Ƿ��ֵ.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("�������"))) & "'", "NULL") & ")"
    End With
    AddArray cllPro, gstrSQL
    zl���ѿ�InsertSQL = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ��鿨���Ƿ�Ϸ�(ByVal strCardNos As String, Optional bln��鿨���� As Boolean = False, Optional bln����Ƿ���ֵ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ŀ����Ƿ�Ϸ�
    '���:strCardNos-���ż�(ÿ�������ö��ŷ���)
    '     bln��鿨����-��Ҫ��鿨����(�Է�����Ч)
    '     bln����Ƿ���ֵ-����Ƿ���ֵ(�Է�����Ч)
    '����:���˺�
    '����:2009-12-11 11:08:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strErrCardNo As String
    Dim lngRow As Long, strTable As String, strValue(0 To 10) As String, varData As Variant
    Dim lngCurRow As Long, i As Long, strCardNosTemp As String, j As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_���� And mCardEditType <> gEd_���� Then
        gstrSQL = "" & _
        "   Select ID,������,�ɷ��ֵ ,����,���,(Select Max(���) From ���ѿ�Ŀ¼ B where A.����=B.���� and A.�ӿڱ��=b.�ӿڱ��) as ������, " & _
        "       to_char(����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��, to_char(ͣ������ ,'yyyy-mm-dd hh24:mi:ss') as ͣ������ " & _
        "   From ���ѿ�Ŀ¼ A " & _
        "   Where A.id = [2]"
    Else
        If zlCommFun.ActualLen(strCardNos) > 1990 Then
            varData = Split(strCardNos, ",")
            strCardNosTemp = ""
            j = 3
            For i = 0 To UBound(varData)
                If j - 3 > 10 Then
                    strCardNosTemp = strCardNosTemp & "," & varData(i)
                Else
                    If zlCommFun.ActualLen(strCardNosTemp) > 1990 Then
                        strValue(j - 3) = Mid(strCardNosTemp, 2)
                        strTable = strTable & " UNION ALL Select Column_Value From Table(Cast(f_Str2list([" & j & "]) As Zltools.t_Strlist)) "
                        strCardNosTemp = ""
                        j = j + 1
                    End If
                    strCardNosTemp = strCardNosTemp & "," & varData(i)
                End If
            Next
            If j - 3 > 10 And strCardNosTemp <> "" Then
                strCardNosTemp = Mid(strCardNosTemp, 2)
                strCardNosTemp = "'" & Replace(strCardNosTemp, ",", "','") & "'"
                strTable = strTable & " UNION ALL   Select ���� From ���ѿ�Ŀ¼ Where ���� in(" & strCardNosTemp & ")"
            ElseIf strCardNosTemp <> "" Then
                strValue(j - 3) = Mid(strCardNosTemp, 2)
                strTable = strTable & " UNION ALL Select Column_Value From Table(Cast(f_Str2list([" & j & "]) As Zltools.t_Strlist)) "
            End If
            If strTable <> "" Then strTable = Mid(strTable, 11)
        Else
            strTable = "Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) "
        End If

        gstrSQL = "" & _
        "   Select  /*+ RULE */ ID,������,�ɷ��ֵ, ����,���,��� as ������,to_char(����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��, to_char(ͣ������ ,'yyyy-mm-dd hh24:mi:ss') as ͣ������ " & _
        "   From ���ѿ�Ŀ¼ A, (" & strTable & ") B " & _
        "   Where A.���� = B.Column_Value And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ�� ) and a.�ӿڱ��=[3]  "
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNos, mlng���ѿ�ID, mlng�ӿڱ��, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))

    strErrCardNo = ""
    Do While Not rsTemp.EOF
        '��鿨���Ƿ�Ϸ�
        Select Case mCardEditType
        Case gEd_����
            If Nvl(rsTemp!����ʱ��, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ�����ʹ�ã������ٷ���,����!"
                Exit Function
            End If
            If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ�ֹͣʹ�ã������ٷ���,����!"
                Exit Function
            End If
            lngCurRow = -1
            For lngRow = 1 To vsGrid.Rows - 1
                If "," & vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("����")) & "," Like "*," & strCardNos & "," Then
                    lngCurRow = lngRow: Exit For
                End If
            Next
            If lngCurRow = -1 Then
                If Val(Nvl(rsTemp!�ɷ��ֵ)) <> IIf(chk�Ƿ��ֵ.value = 1, 1, 0) Then
                    ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ�ԭ��Ϊ" & IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, "��ֵ��", "�ǳ�ֵ��") & "��������Ϊ" & IIf(chk�Ƿ��ֵ.value = 1, "��ֵ��", "�ǳ�ֵ��") & ",����!"
                    Exit Function
                End If
            Else
                If Val(Nvl(rsTemp!�ɷ��ֵ)) <> IIf((vsGrid.Cell(flexcpData, lngCurRow, vsGrid.ColIndex("�Ƿ��ֵ"))) = 1, 1, 0) Then
                    ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ�ԭ��Ϊ" & IIf(Val(Nvl(rsTemp!�ɷ��ֵ)) = 1, "��ֵ��", "�ǳ�ֵ��") & "��������Ϊ" & IIf(chk�Ƿ��ֵ.value = 1, "��ֵ��", "�ǳ�ֵ��") & ",����!"
                    Exit Function
                End If
            End If
            If Trim(Nvl(rsTemp!������)) <> Mid(cbo������.Text, InStr(cbo������.Text, ".") + 1) Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ�ԭ���Ŀ�����Ϊ" & Trim(Nvl(rsTemp!������)) & "��������Ϊ" & Mid(cbo������.Text, InStr(cbo������.Text, ".") + 1) & ",����!"
                Exit Function
            End If
        Case gEd_�޸�
            If Val(Nvl(rsTemp!���)) < Val(Nvl(rsTemp!������)) Then
                ShowMsgbox "ע��:" & vbCrLf & "   �����޸���ʷ������Ϣ(����Ϊ:" & Nvl(rsTemp!����) & ") ,����!"
                Exit Function
            End If
        Case gEd_��ֵ
            If Val(Nvl(rsTemp!���)) < Val(Nvl(rsTemp!������)) Then
                ShowMsgbox "ע��:" & vbCrLf & "   ���ܶ���ʷ���Ž��г�ֵ(����Ϊ:" & Nvl(rsTemp!����) & "),����!"
                Exit Function
            End If
            If Nvl(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ����գ������ٳ�ֵ,���ȷ�����,�ٽ��г�ֵ!"
                Exit Function
            End If
            If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ�ֹͣʹ�ã������ٳ�ֵ,����!"
                Exit Function
            End If
            If Val(Nvl(rsTemp!�ɷ��ֵ)) <> 1 Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ�������ֵ���������ٳ�ֵ,����!"
                Exit Function
            End If
        Case gEd_����
            If Val(Nvl(rsTemp!���)) < Val(Nvl(rsTemp!������)) Then
                ShowMsgbox "ע��:" & vbCrLf & "   ���ܶԻ�����ʷ����(����Ϊ:" & Nvl(rsTemp!����) & ") ,����!"
                Exit Function
            End If
            If Nvl(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ����գ������ٻ���,����!"
                Exit Function
            End If
            If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ�ֹͣʹ�ã����ܻ���,����!"
                Exit Function
            End If
        Case gEd_�˿�
            If Val(Nvl(rsTemp!���)) < Val(Nvl(rsTemp!������)) Then
                ShowMsgbox "ע��:" & vbCrLf & "   ���ܶԻ�����ʷ����(����Ϊ:" & Nvl(rsTemp!����) & ") ,����!"
                Exit Function
            End If
            If Nvl(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ����գ������ٻ���,����!"
                Exit Function
            End If
            If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "ע��:" & vbCrLf & "   ����Ϊ:" & Nvl(rsTemp!����) & " �����ѿ��ѱ�ֹͣʹ�ã����ܻ���,����!"
                Exit Function
            End If
        End Select
        rsTemp.MoveNext
    Loop

    If rsTemp.RecordCount = 0 Then
        If mCardEditType = gEd_�޸� Then
            ShowMsgbox "���ѿ������Ѿ�������ɾ���������޸Ŀ���Ϣ,����!"
            Exit Function
        End If
        If mCardEditType = gEd_��ֵ Then
            ShowMsgbox "���ѿ������Ѿ�������ɾ�������ܳ�ֵ,����!"
            Exit Function
        End If
        If mCardEditType = gEd_���� Then
            ShowMsgbox "���ѿ������Ѿ�������ɾ�������ܻ���,����!"
            Exit Function
        End If
        If mCardEditType = gEd_�˿� Then
            ShowMsgbox "���ѿ������Ѿ�������ɾ���������˿�,����!"
            Exit Function
        End If
    End If

    ��鿨���Ƿ�Ϸ� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckVsGridData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������е������Ƿ�Ϸ�
    '����:���˺�
    '����:2009-12-11 12:06:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, lngRow As Long, i As Long, strCurNos As String
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        For lngRow = 1 To .Rows - 1
            strTemp = Trim(.Cell(flexcpData, lngRow, .ColIndex("����")))
            If strTemp <> "" Then
                If zlCommFun.ActualLen(strTemp) > 4000 Then
                    Do While True
                        i = InStr(3900, strTemp, ",")
                        If i > 0 Then
                            strCurNos = Mid(strTemp, 1, i - 1)
                            strTemp = Mid(strTemp, i + 1)
                            If strCurNos <> "" Then
                                If ��鿨���Ƿ�Ϸ�(strCurNos) = False Then
                                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                                    Exit Function
                                End If
                            End If
                        Else
                            If strTemp <> "" Then
                                If ��鿨���Ƿ�Ϸ�(strTemp) = False Then
                                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                                    Exit Function
                                End If
                            End If
                            Exit Do
                        End If
                    Loop
                Else
                    If ��鿨���Ƿ�Ϸ�(strTemp) = False Then
                        .Row = lngRow: zlCtlSetFocus vsGrid, True
                        Exit Function
                    End If
                End If
            End If

            If mCardEditType = gEd_���� Then
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("����ԭ��")), 50, 0, "����ԭ��") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("��ע")), 100, 0, "��ע") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("��ֵ˵��")), 100, 0, "��ֵ˵��") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If Trim(.Cell(flexcpData, lngRow, .ColIndex("ID"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("����"))) Then
                    ShowMsgbox "�ڵ�" & lngRow & "�е��������벻��ȷ,���������ȷ�������Ƿ���ȷ!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("����"))), 20, 0, "����") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
'                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("ȷ������"))), 20, 0, "ȷ������") = False Then
'                    .Row = lngRow: zlCtlSetFocus vsGrid, True
'                    Exit Function
'                End If
                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("�쿨��"))), 20, 0, "�쿨��") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                '�����
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("�����"))), 16, True, False, 0, "�����") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("ʵ������"))), 16, True, False, 0, "ʵ������") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("��ֵ����"))), 3, True, False, 0, "��ֵ����") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("���γ�ֵ"))), 16, True, False, 0, "���γ�ֵ") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("ʵ�ʳ�ֵ�ɿ�"))), 16, True, False, 0, "ʵ�ʳ�ֵ�ɿ�") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If



                If Val(Trim(.TextMatrix(lngRow, .ColIndex("�����")))) < Val(Trim(.TextMatrix(lngRow, .ColIndex("ʵ������")))) Then
                    ShowMsgbox "ע��:" & vbCrLf & "������С��ʵ�����۶�,����!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If Val(.TextMatrix(lngRow, .ColIndex("���γ�ֵ"))) < Val(.TextMatrix(lngRow, .ColIndex("ʵ�ʳ�ֵ�ɿ�"))) Then
                    ShowMsgbox "ע��:" & vbCrLf & "���γ�ֵ����С��ʵ�ʳ�ֵ�ɿ�,����!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                 If Val(.TextMatrix(lngRow, .ColIndex("��ֵ����"))) > 100 Then
                     ShowMsgbox "ע��:" & vbCrLf & "��ֵ���ʲ��ܴ���100%,����!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                     Exit Function
                 End If
                 If Val(.TextMatrix(lngRow, .ColIndex("��ֵ����"))) < 0 Then
                     ShowMsgbox "ע��:" & vbCrLf & "��ֵ���ʲ���С��0,����!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                     Exit Function
                 End If

            End If
        Next
    End With
    CheckVsGridData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Check�ɿ����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ɿ����
    '����:���˺�
    '����:2009-12-11 13:56:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txtʵ�պϼ�).Text), 16, True, False, 0, "ʵ�պϼ�") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txtʵ�պϼ�), True
        Exit Function
    End If
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt���νɿ�).Text), 16, True, False, 0, "���νɿ�") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
        Exit Function
    End If
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt�Ҳ�).Text), 16, True, False, 0, "�Ҳ�") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt�Ҳ�), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt�Ҳ�).Tag) < 0 Then
        If Val(txtEdit(mtxtIdx.idx_txt���νɿ�).Text) <> 0 Then
            ShowMsgbox "ע��:" & vbCrLf & "    ������Ľɿ����(��Ӧ��:" & Trim(txtEdit(mtxtIdx.idx_txt�Ҳ�).Text) & ")�� ,���ܼ���!"
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
            Exit Function
        End If
        If mCardEditType = gEd_���� Then
            If MsgBox("ע��:" & vbCrLf & "   �㻹δ��ȡ�쿨�˵Ľɿ�,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
                Exit Function
            End If
        ElseIf mCardEditType = gEd_�޸� Then
            If MsgBox("ע��:" & vbCrLf & "   �㻹δ��ȡ�쿨�˵Ľɿ�,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
                Exit Function
            End If
        ElseIf mCardEditType = gEd_��ֵ Then
            If MsgBox("ע��:" & vbCrLf & "   �㻹δ��ȡ��صĽɿ�,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
                Exit Function
            End If
        Else
            If MsgBox("ע��:" & vbCrLf & "   �㻹δ��ȡ��صĽɿ�,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�), True
                Exit Function
            End If
        End If
    End If
    Check�ɿ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2009-12-14 15:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt����ԭ��).Text), 50, 0, "����ԭ��") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����ԭ��), True
        Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt��ע).Text), 100, 0, "��ע") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt��ע), True
        Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text), 100, 0, "��ֵ˵��") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt��ֵ��ע), True
        Exit Function
    End If
    '76213,���ϴ�,2014-08-12,������ʱ������������֤
    If mCardEditType <> gEd_�޸� And mCardEditType <> gEd_���� Then
        If Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text) <> Trim(txtEdit(mtxtIdx.idx_txt����).Text) Then
            ShowMsgbox "�������벻��ȷ,���������ȷ�������Ƿ���ȷ!"
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
            Exit Function
        End If
        If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt����).Text), 20, 0, "����") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
            Exit Function
        End If
        If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text), 20, 0, "ȷ������") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txtȷ������), True
            Exit Function
        End If
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Text), 20, 0, "�쿨��") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt�쿨��), True
        Exit Function
    End If
    If Trim(txtEdit(mtxtIdx.idx_txt�쿨����).Text) <> "" And Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag) = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��������쿨������������!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt�쿨����), True
        Exit Function
    End If
    '�����
   If CheckInput����� = False Then Exit Function
   If CheckInputʵ�����۶� = False Then Exit Function
   If CheckInputʵ�ʳ�ֵ�ɿ� = False Then Exit Function
   If CheckInput��ֵ���� = False Then Exit Function
   If CheckInput���γ�ֵ = False Then Exit Function
    CheckInput = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-11 10:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo Errhand:

    '��鿨��
    Select Case mCardEditType
    Case gEd_����
        If CheckVsGridData = False Then Exit Function
        If Check�ɿ���� = False Then Exit Function
    Case gEd_����
        If CheckVsGridData = False Then Exit Function
    Case gEd_�޸�
        If ��鿨���Ƿ�Ϸ�("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
            Exit Function
        End If
        If CheckInput() = False Then Exit Function

    Case gEd_��ֵ
        If mlng���ѿ�ID = 0 Then
            ShowMsgbox "δѡ��Ϸ������ѿ�,���ܳ�ֵ,����"
            Exit Function
        End If
        If ��鿨���Ƿ�Ϸ�("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
            Exit Function
        End If
       If Check�ɿ���� = False Then Exit Function
    Case gEd_����

    Case gEd_�˿�
        If ��鿨���Ƿ�Ϸ�("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����), True
            Exit Function
        End If
    End Select

    isValied = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SavePayCard(ByVal lng������� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢����Ϣ
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-10 16:14:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection, lngRow  As Long, lngID As Long, strCardNos As String, strTemp As String, strCurNos As String, i As Long
    Dim str����ʱ�� As String, varData As Variant
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Set cllPro = New Collection
    With vsGrid
        For lngRow = 1 To .Rows - 1
            strTemp = Trim(.Cell(flexcpData, lngRow, .ColIndex("����")))
            If strTemp <> "" Then
                If zlCommFun.ActualLen(strTemp) > 4000 Then
                    Do While True
                        i = InStr(3900, strTemp, ",")
                        If i > 0 Then
                            strCurNos = Mid(strTemp, 1, i - 1)
                            strTemp = Mid(strTemp, i + 1)
                            If strCurNos <> "" Then
                                If zl���ѿ�InsertSQL(lng�������, strCurNos, str����ʱ��, lngRow, cllPro) = False Then Exit Function
                            End If
                        Else
                            If strTemp <> "" Then
                                If zl���ѿ�InsertSQL(lng�������, strTemp, str����ʱ��, lngRow, cllPro) = False Then Exit Function
                            End If
                            Exit Do
                        End If
                    Loop
                Else
                    If zl���ѿ�InsertSQL(lng�������, strTemp, str����ʱ��, lngRow, cllPro) = False Then Exit Function
                End If
            End If
        Next
    End With
    If cllPro.count = 0 Then Exit Function
    Err = 0: On Error GoTo Errhand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SavePayCard = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_Change(Index As Integer)
    If mblnNoClick Then Exit Sub
    mblnChange = True
    '--����47855
    If Index = mtxtIdx.idx_txt���� Or Index = mtxtIdx.idx_txtȷ������ Then
        vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("����")) = "******************": vsGrid.Cell(flexcpData, vsGrid.Row, vsGrid.ColIndex("����")) = Trim(txtEdit(mtxtIdx.idx_txt����).Text)
        vsGrid.Cell(flexcpData, vsGrid.Row, vsGrid.ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text)  '���ȷ������
    End If
    txtEdit(Index).Tag = ""
    If Index = mtxtIdx.idx_txt��ʼ���� Then
        txtEdit(mtxtIdx.idx_txt��������) = ""
    ElseIf Index = mtxtIdx.idx_txt���νɿ� Or Index = mtxtIdx.idx_txtʵ�պϼ� Then
        Call SetLblCatpion
    ElseIf mtxtIdx.idx_txtʵ�����۶� = Index Then
        Call Set�ɷ��ֵ
    End If
End Sub

Public Sub SetLblCatpion()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ������Ҳ��ı���
    '���ƣ����˺�
    '���ڣ�2010-03-22 16:01:21
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim dbl�Ҳ� As Double
   dbl�Ҳ� = Val(txtEdit(mtxtIdx.idx_txt���νɿ�).Text) - mdblʵ�պϼ�
   If dbl�Ҳ� >= 0 Then
         txtEdit(mtxtIdx.idx_txt�Ҳ�).Text = Format(dbl�Ҳ�, "0.00")
         txtEdit(mtxtIdx.idx_txt�Ҳ�).ForeColor = &H80000008
         lblEdit(mlblIdx.idx_lbl�Ҳ�).Caption = "�Ҳ�"
         txtEdit(mtxtIdx.idx_txt�Ҳ�).Tag = txtEdit(mtxtIdx.idx_txt�Ҳ�).Text

   Else
         txtEdit(mtxtIdx.idx_txt�Ҳ�).Text = Format(Abs(dbl�Ҳ�), "0.00")
         txtEdit(mtxtIdx.idx_txt�Ҳ�).Tag = Format(dbl�Ҳ�, "0.00")
         txtEdit(mtxtIdx.idx_txt�Ҳ�).ForeColor = vbRed
         lblEdit(mlblIdx.idx_lbl�Ҳ�).Caption = "Ӧ��"
   End If

End Sub


Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt��ʼ����, mtxtIdx.idx_txt��������, mtxtIdx.idx_txt����

        If mtxtIdx.idx_txt�������� = Index Then
            gTy_TestBug.strStartNo = Trim(txtEdit(mtxtIdx.idx_txt��ʼ����))
        Else
            gTy_TestBug.strStartNo = ""
        End If
'        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")

    Case mtxtIdx.idx_txt��ע, mtxtIdx.idx_txt��ֵ��ע, mtxtIdx.idx_txt����ԭ��, mtxtIdx.idx_txt�ɿ���, mtxtIdx.idx_txt�쿨����, mtxtIdx.idx_txt�쿨��
        zlCommFun.OpenIme True
    Case mtxtIdx.idx_txt����
        Call OpenPassKeyboard(txtEdit(Index), False)
    Case mtxtIdx.idx_txtȷ������
        Call OpenPassKeyboard(txtEdit(Index), True)
    Case Else
        zlCommFun.OpenIme False
    End Select
    zlControl.TxtSelAll txtEdit(Index)
End Sub
Private Function zlBrusCard(ByVal strCardNo As String, Optional blnBrushCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '����:���˺�
    '����:2009-12-16 10:33:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnModifyCard As Boolean
    If Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt��ʼ����) Then
        intIndex = mtxtIdx.idx_txt��ʼ����
    ElseIf Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt��������) Then
        intIndex = mtxtIdx.idx_txt��������
    ElseIf Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt����) Then
        intIndex = mtxtIdx.idx_txt����
    Else
        zlBrusCard = True
        Exit Function
    End If
    txtEdit(intIndex).Text = strCardNo
    txtEdit(intIndex).Tag = strCardNo

    Select Case intIndex
    Case mtxtIdx.idx_txt��ʼ����, mtxtIdx.idx_txt��������

        If mCardEditType = gEd_���� Then
           If InsertIntoGrid(strCardNo, mtxtIdx.idx_txt�������� = intIndex, , False) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '��grid�ؼ�����textbox�ؼ��м���
                Call FromGridToCtlData
                '��Ϊ���ܴ����������������,��˾��ƶ�����һ�ؼ�����
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex)
           End If

        ElseIf mCardEditType = gEd_���� Then
           If InsertIntoGrid(strCardNo, False) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '��Ϊ���ܴ����������յ����,��˾��ƶ�����һ�ؼ�����
                '�����ǰ�Ŀ��ţ���Ҫ����������ˢ��
                txtEdit(intIndex).Text = ""
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        End If
    Case mtxtIdx.idx_txt����
        Select Case mCardEditType
        Case gEd_����
            '�ڱ༭������ʱ,��Ĭ��Ϊ�޸�ԭ���Ŀ���
            ' 1.��ǰ�ڿ��Ŵ��ֹ�����
            ' 2.ˢ�����Զ�����
            blnModifyCard = (intIndex = mtxtIdx.idx_txt����) And (blnBrushCard = False)

           If InsertIntoGrid(strCardNo, False, , blnModifyCard) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                 Exit Function
           Else
                '��Ϊ���ܴ����������������,��˾��ƶ�����һ�ؼ�����
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        Case gEd_��ֵ
            If zlFromCardNOGetDataToCtrl(strCardNo) = False Then
                zl_CtlSetFocus txtEdit(intIndex), True
                Call SetEditProperty
                Exit Function
            End If
            Call SetEditProperty
            zl_CtlSetFocus txtEdit(intIndex)
            zlCommFun.PressKey vbKeyTab
        Case gEd_����
           If InsertIntoGrid(strCardNo, mtxtIdx.idx_txt�������� = intIndex) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '��Ϊ���ܴ����������յ����,��˾��ƶ�����һ�ؼ�����
                '�����ǰ�Ŀ��ţ���Ҫ����������ˢ��
                'txtEdit(intIndex).Text = ""
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        End Select

    Case gEd_����
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_ȡ������
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_ȡ���˿�
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_�޸�
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case Else
    End Select

    zlBrusCard = True
End Function


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str���� As String, str���� As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt��ʼ����, mtxtIdx.idx_txt��������, mtxtIdx.idx_txt����
        'If txtEdit(Index) = "" Then Exit Sub
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
''        If txtEdit(Index).Text = "" Then
''            'ֱ�ӵ�����
''            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
''                Exit Sub
''            End If
''            txtEdit(Index).Text = strCardNo
''            txtEdit(Index).Tag = strCardNo
''        End If
''        Call zlBrusCard(Trim(txtEdit(Index)), False)
    Case mtxtIdx.idx_txt�쿨��
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub

        'ѡ����Ա
        lngID = Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text), lngID, , True, , , , , , , "") = False Then
            zlCommFun.PressKey vbKeyTab
        End If
        If mCardEditType = gEd_���� Then
            '�쿨�˾��ǽɿ���
            txtEdit(mtxtIdx.idx_txt�ɿ���).Text = txtEdit(mtxtIdx.idx_txt�쿨��).Text
            txtEdit(mtxtIdx.idx_txt�ɿ���).Tag = txtEdit(mtxtIdx.idx_txt�쿨��).Tag
        End If

        '��Ҫ��ȡȱʡ����:
        If zl_From��Ա��ȡȱʡ����(Val(txtEdit(mtxtIdx.idx_txt�쿨��).Tag), str����, str����, lngID) Then
            txtEdit(mtxtIdx.idx_txt�쿨����).Text = str���� & "-" & str����
            txtEdit(mtxtIdx.idx_txt�쿨����).Tag = lngID
        End If
        Exit Sub
    Case mtxtIdx.idx_txt�ɿ���
        '�ɿ���,��ȷ��,��ѡ��
        zlCommFun.PressKey vbKeyTab: Exit Sub
    Case mtxtIdx.idx_txt�쿨����
        'ѡ����
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        'ѡ��ȱʡ����
        lngID = Val(txtEdit(mtxtIdx.idx_txt�쿨��).Tag)
        If Select����ѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text), "", IIf(lngID = 0, False, True), "", 0, "����ѡ����", , , , , lngID) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_txt����ԭ��
        'ѡ�񷢿�ԭ��
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If zl_SelectAndNotAddItem(Me, txtEdit(Index), Trim(txtEdit(Index).Text), "���÷���ԭ��", "���÷���ԭ��ѡ��", True, True, , , , True) = False Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    Case mtxtIdx.idx_txtȷ������
        If Trim(txtEdit(mtxtIdx.idx_txt����).Text) <> "" Then
            If Trim(txtEdit(Index).Text) = "" Then
                ShowMsgbox "������ȷ������,����!"
                zl_CtlSetFocus txtEdit(Index)
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
            If Trim(txtEdit(Index).Text) <> Trim(txtEdit(mtxtIdx.idx_txt����).Text) Then
                ShowMsgbox "��������벻һ��,����!"
                zl_CtlSetFocus txtEdit(Index)
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
        End If
        If Trim(txtEdit(Index).Text) <> "" And Trim(txtEdit(mtxtIdx.idx_txt����).Text) = "" Then
            ShowMsgbox "����δ����,����!"
            zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt����)
            Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt�����
        If CheckInput����� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab

    Case mtxtIdx.idx_txtʵ�����۶�
        If CheckInputʵ�����۶� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�
        If CheckInputʵ�ʳ�ֵ�ɿ� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab

    Case mtxtIdx.idx_txt���γ�ֵ
        If CheckInput���γ�ֵ = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt��ֵ����
        If CheckInput��ֵ���� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt���νɿ�
        If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
            If MsgBox("���Ƿ���Ҫ������ص�" & IIf(mCardEditType = gEd_�޸�, "�޸�", "����") & "������?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            Else
                If txtEdit(mtxtIdx.idx_txt���νɿ�).Enabled And txtEdit(mtxtIdx.idx_txt���νɿ�).Visible Then
                    zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���νɿ�)
                End If
            End If
        End If
        If mCardEditType = gEd_��ֵ Then
            If MsgBox("���Ƿ���Ҫ������صĳ�ֵ������?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            End If
        End If
    Case Else
        If Index = mtxtIdx.idx_txt��ֵ��ע Then
            zlCommFun.PressKey vbKeyTab
        End If
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Function CheckInput��ֵ����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֵ�����Ƿ�Ϸ�
    '����:���˺�
    '����:2009-12-17 16:03:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt��ֵ����).Text), 3, True, False, 0, "��ֵ����") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt��ֵ����), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt��ֵ����).Text) > 100 Then
        ShowMsgbox "ע��:" & vbCrLf & "��ֵ���ʲ��ܴ���100%,����!"
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt��ֵ����)
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt��ֵ����)
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt��ֵ����).Text) < 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "��ֵ���ʲ���С��0,����!"
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt��ֵ����)
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt��ֵ����)
        Exit Function
    End If
    CheckInput��ֵ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInputʵ�����۶�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʵ�����۶�
    '����:
    '����:���˺�
    '����:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text), 16, True, False, 0, "ʵ������") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txtʵ�����۶�), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt�����).Text) < Val(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text) Then
        ShowMsgbox "ע��:" & vbCrLf & "������С��ʵ�����۶�,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txtʵ�����۶�), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txtʵ�����۶�)
        Exit Function
    End If
    CheckInputʵ�����۶� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput���γ�ֵ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鱾�γ�ֵ
    '����:
    '����:���˺�
    '����:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text), 16, True, False, 0, "���γ�ֵ") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���γ�ֵ), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text) < Val(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text) Then
        ShowMsgbox "ע��:" & vbCrLf & "���γ�ֵ����С��ʵ�ʳ�ֵ�ɿ�,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt���γ�ֵ), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt���γ�ֵ)
        Exit Function
    End If
    CheckInput���γ�ֵ = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckInputʵ�ʳ�ֵ�ɿ�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鱾�γ�ֵ
    '����:
    '����:���˺�
    '����:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text), 16, True, False, 0, "ʵ�ʳ�ֵ�ɿ�") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text) < Val(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text) Then
        ShowMsgbox "ע��:" & vbCrLf & "���γ�ֵ����С��ʵ�ʳ�ֵ�ɿ�,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�)
        Exit Function
    End If
    CheckInputʵ�ʳ�ֵ�ɿ� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput�����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鿨���
    '����:���˺�
    '����:2009-12-17 16:08:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:

    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt�����).Text), 16, True, False, 0, "�����") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt�����), True
        Exit Function
    End If

    If Val(txtEdit(mtxtIdx.idx_txt�����).Text) < Val(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text) Then
        ShowMsgbox "ע��:" & vbCrLf & "������С��ʵ�����۶�,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt�����), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt�����)
        Exit Function
    End If

    CheckInput����� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function



Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt���γ�ֵ, mtxtIdx.idx_txt���νɿ�, mtxtIdx.idx_txt��ֵ����, mtxtIdx.idx_txt��ǰ���, mtxtIdx.idx_txt�����, mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�, mtxtIdx.idx_txtʵ�����۶�, mtxtIdx.idx_txtʵ�պϼ�, mtxtIdx.idx_txt�Ҳ�
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m���ʽ)
    Case mtxtIdx.idx_txt����, mtxtIdx.idx_txt��ʼ����, mtxtIdx.idx_txt��������
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m�ı�ʽ)
        If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0

        Call BrushCard(txtEdit(Index), KeyAscii)

    Case Else
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m�ı�ʽ)
    End Select
End Sub
 Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '����:���˺�
    '����:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean

    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(objEdit, KeyAscii, mTy_MoudlePara.bln��������)
    If blnCard And Len(objEdit.Text) = objEdit.MaxLength - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        Call zlBrusCard(Trim(objEdit), False)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = idx_txt�쿨���� Then
        If txtEdit(idx_txt�쿨����).Tag = "" And txtEdit(idx_txt�쿨����).Text <> "" Then txtEdit(idx_txt�쿨����).Text = ""
    End If
    Select Case Index
    Case mtxtIdx.idx_txt����
        Call ClosePassKeyboard(txtEdit(Index))
    Case mtxtIdx.idx_txtȷ������
        Call ClosePassKeyboard(txtEdit(Index))
    End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If mCardEditType <> gEd_��ֵ And mCardEditType <> gEd_���� And mCardEditType <> gEd_�޸� Then Exit Sub

    Select Case Index
    Case mtxtIdx.idx_txt�����
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        Call Calc���
        If mCardEditType = gEd_��ֵ Or mCardEditType = gEd_�޸� Then Calcʵ�պϼ�
    Case mtxtIdx.idx_txtʵ�����۶�
        Call Calc���
        Call Set�ɷ��ֵ
        If mCardEditType = gEd_��ֵ Or mCardEditType = gEd_�޸� Then Calcʵ�պϼ�
    Case mtxtIdx.idx_txt���γ�ֵ
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = Format(Val(txtEdit(Index).Text) * (Round(Val(txtEdit(mtxtIdx.idx_txt��ֵ����)) / 100, 6)), "0.00")
        Call Calc���
        If mCardEditType = gEd_��ֵ Or mCardEditType = gEd_�޸� Then Calcʵ�պϼ�
    Case mtxtIdx.idx_txt��ֵ����
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = Format(Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text) * (Round(Val(txtEdit(mtxtIdx.idx_txt��ֵ����)) / 100, 4)), "0.00")
        Call Calc���
        If mCardEditType = gEd_��ֵ Or mCardEditType = gEd_�޸� Then Calcʵ�պϼ�
    Case mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        If Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text) <> 0 Then
            txtEdit(mtxtIdx.idx_txt��ֵ����).Text = Format((Round(Val(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text) / Val(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text), 6)) * 100, "0.00")
        Else
             txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text
        End If
        Call Calc���
        If mCardEditType = gEd_��ֵ Or mCardEditType = gEd_�޸� Then Calcʵ�պϼ�
    Case Else

    End Select

    If mCardEditType = gEd_���� Or mCardEditType = gEd_�޸� Then
        '�޸ĵĻ�,��Ҫͬ�����������е����ݲ���
        With vsGrid
            If Split(.Cell(flexcpData, .Row, .ColIndex("����")) & ",", ",")(0) <> txtEdit(mtxtIdx.idx_txt����).Text Then Exit Sub
            'ֻ����ͬ���ݲſ��Ը���
            Select Case Index
            Case idx_txt����
                '����
            Case idx_txt����
                '--����47855
                '.TextMatrix(.Row, .ColIndex("����")) = "******************": .Cell(flexcpData, .Row, .ColIndex("����")) = Trim(txtEdit(mtxtIdx.idx_txt����).Text)
                If Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text) <> Trim(txtEdit(mtxtIdx.idx_txt����).Text) And Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text) <> "" Then
                     ShowMsgbox "����������ȷ�����벻һ��,����"
                     txtEdit(mtxtIdx.idx_txtȷ������).SetFocus
                     'If Trim(txtEdit(mtxtIdx.idx_txt����).Text) <> "" Then Cancel = True: Exit Sub
                End If
            Case idx_txtȷ������
                   If Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text) <> Trim(txtEdit(mtxtIdx.idx_txt����).Text) Then
                        ShowMsgbox "����������ȷ�����벻һ��,����"
                        txtEdit(mtxtIdx.idx_txt����).SetFocus
                        'If Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text) <> "" Then Cancel = True: Exit Sub
                   End If
                   '--����47855
                   ' .Cell(flexcpData, .Row, .ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txtȷ������).Text)  '���ȷ������
            Case idx_txt����ԭ��
                .TextMatrix(.Row, .ColIndex("����ԭ��")) = Trim(txtEdit(mtxtIdx.idx_txt����ԭ��).Text)
            Case idx_txt�쿨��
                .TextMatrix(.Row, .ColIndex("�쿨��")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Text): .Cell(flexcpData, .Row, .ColIndex("�쿨��")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Tag)
                .TextMatrix(.Row, .ColIndex("�쿨����")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨����).Text): .Cell(flexcpData, .Row, .ColIndex("�쿨����")) = Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)
                .TextMatrix(.Row, .ColIndex("��ֵ�ɿ���")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Text): .Cell(flexcpData, .Row, .ColIndex("��ֵ�ɿ���")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨��).Tag)

            Case idx_txt����Ч����
            Case idx_txt��ע
                .TextMatrix(.Row, .ColIndex("��ע")) = Trim(txtEdit(mtxtIdx.idx_txt��ע).Text)
            Case idx_txt������
            Case idx_txt��������
            Case idx_txt��ǰ���
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
            Case idx_txt�����
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("�����")) = Trim(txtEdit(mtxtIdx.idx_txt�����).Text)
                .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
                Calcʵ�պϼ�
            Case idx_txtʵ�����۶�
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("ʵ������")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text)
                .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
                Calcʵ�պϼ�
            Case idx_txt���γ�ֵ, idx_txt��ֵ����, idx_txtʵ�ʳ�ֵ�ɿ�
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                If chk�Ƿ��ֵ.value = 0 Then
                    .TextMatrix(.Row, .ColIndex("��ֵ����")) = ""
                    .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = ""
                    .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = ""
                    .TextMatrix(.Row, .ColIndex("���㷽ʽ")) = cboStyle.Text
                Else
                    .TextMatrix(.Row, .ColIndex("��ֵ����")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ����).Text)
                    .TextMatrix(.Row, .ColIndex("���γ�ֵ")) = Trim(txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text)
                    .TextMatrix(.Row, .ColIndex("ʵ�ʳ�ֵ�ɿ�")) = Trim(txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text)
                    .TextMatrix(.Row, .ColIndex("���㷽ʽ")) = cboStyle.Text
                End If
                .TextMatrix(.Row, .ColIndex("��ǰ���")) = Trim(txtEdit(mtxtIdx.idx_txt��ǰ���).Text)
                Calcʵ�պϼ�
               Case idx_txt������, idx_txt�ʺ�, idx_txt�������
                If chk�Ƿ��ֵ.value = 0 Then
                    .TextMatrix(.Row, .ColIndex("������")) = ""
                    .TextMatrix(.Row, .ColIndex("�ʺ�")) = ""
                    .TextMatrix(.Row, .ColIndex("�������")) = ""
                    .TextMatrix(.Row, .ColIndex("���㷽ʽ")) = cboStyle.Text
                Else
                    .TextMatrix(.Row, .ColIndex("������")) = Trim(txtEdit(mtxtIdx.idx_txt������).Text)
                    .TextMatrix(.Row, .ColIndex("�ʺ�")) = Trim(txtEdit(mtxtIdx.idx_txt�ʺ�).Text)
                    .TextMatrix(.Row, .ColIndex("�������")) = Trim(txtEdit(mtxtIdx.idx_txt�������).Text)
                    .TextMatrix(.Row, .ColIndex("���㷽ʽ")) = cboStyle.Text
                End If
                 
            
            Case idx_txt�Ҳ�
            Case idx_txtʵ�պϼ�
            Case idx_txt���νɿ�
            Case idx_txt�쿨����
                .TextMatrix(.Row, .ColIndex("�쿨����")) = Trim(txtEdit(mtxtIdx.idx_txt�쿨����).Text): .Cell(flexcpData, .Row, .ColIndex("�쿨����")) = Val(txtEdit(mtxtIdx.idx_txt�쿨����).Tag)
            Case idx_txt��ʼ����
                Call txtEdit_KeyPress(idx_txt��ʼ����, 13)
            Case idx_txt��������
            Case idx_txt������
            Case idx_txt����ʱ��
            Case idx_txt�ɿ���
                .TextMatrix(.Row, .ColIndex("��ֵ�ɿ���")) = Trim(txtEdit(mtxtIdx.idx_txt�ɿ���).Text)
            Case idx_txt��ֵ��ע
                .TextMatrix(.Row, .ColIndex("��ֵ˵��")) = Trim(txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text)
            End Select
       End With
    End If
End Sub
Private Function FromGridToCtlData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���е������������в�������
    '����:���˺�
    '����:2009-12-17 17:42:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, lngRow As Long, strTemp As String, i As Long
    Dim strCards As String
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        lngRow = .Row
        If Trim(.TextMatrix(lngRow, .ColIndex("����"))) = "" Then
            Call ClearCtlData: Call SetDefaultValue(True): FromGridToCtlData = True: Exit Function
        End If
        
        '95344: ���ϴ�,2016/4/22,���ų��Ȳ���������£���ȷ��ȡ�����ĵ�һ�ſ���
        strCards = Split(Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "(", "(")(0)
        txtEdit(mtxtIdx.idx_txt����).Text = Split(strCards & "��", "��")(0)
        mblnNoClick = True
        With cbo������
            .ListIndex = -1
            strTemp = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("������")): blnFind = False
            For i = 0 To .ListCount - 1
                If .List(i) & ";" Like "*." & strTemp & ";" Then
                    blnFind = True
                    .ListIndex = i: Exit For
                End If
            Next
            If blnFind = False And strTemp <> "" Then
                .AddItem strTemp
                .ListIndex = .NewIndex
            End If
        End With

        With lvwType
            strTemp = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("�������")): blnFind = False
            For i = 1 To .ListItems.count
                If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
                    .ListItems(i).Checked = True
                Else
                    .ListItems(i).Checked = False
                End If
            Next
        End With

        chk�Ƿ��ֵ.value = IIf(Val(.Cell(flexcpData, lngRow, .ColIndex("�Ƿ��ֵ"))) = 1, 1, 0)
        txtEdit(mtxtIdx.idx_txt����).Text = .Cell(flexcpData, lngRow, .ColIndex("����"))
        txtEdit(mtxtIdx.idx_txtȷ������).Text = .Cell(flexcpData, lngRow, .ColIndex("ID"))
        txtEdit(mtxtIdx.idx_txt����ԭ��).Text = .TextMatrix(lngRow, .ColIndex("����ԭ��")): txtEdit(mtxtIdx.idx_txt����ԭ��).Tag = txtEdit(mtxtIdx.idx_txt����ԭ��).Text
        txtEdit(mtxtIdx.idx_txt�쿨��).Text = .TextMatrix(lngRow, .ColIndex("�쿨��")): txtEdit(mtxtIdx.idx_txt�쿨��).Tag = .Cell(flexcpData, lngRow, .ColIndex("�쿨��"))
        txtEdit(mtxtIdx.idx_txt�쿨����).Text = .TextMatrix(lngRow, .ColIndex("�쿨����")): txtEdit(mtxtIdx.idx_txt�쿨����).Tag = .Cell(flexcpData, lngRow, .ColIndex("�쿨����"))
        txtEdit(mtxtIdx.idx_txt��ע).Text = .TextMatrix(lngRow, .ColIndex("��ע"))

        txtEdit(mtxtIdx.idx_txt������).Text = .TextMatrix(lngRow, .ColIndex("������"))
        txtEdit(mtxtIdx.idx_txt��������).Text = .TextMatrix(lngRow, .ColIndex("��������"))
        txtEdit(mtxtIdx.idx_txt��ǰ���).Text = .TextMatrix(lngRow, .ColIndex("��ǰ���"))
        If .TextMatrix(lngRow, .ColIndex("����Ч��")) = "" Or IsDate(.TextMatrix(lngRow, .ColIndex("����Ч��"))) = False Then
            dtp����Ч����.value = Null
        Else
            dtp����Ч����.value = CDate(.TextMatrix(lngRow, .ColIndex("����Ч��")))
        End If
        txtEdit(mtxtIdx.idx_txt�����).Text = .TextMatrix(lngRow, .ColIndex("�����"))
        txtEdit(mtxtIdx.idx_txtʵ�����۶�).Text = .TextMatrix(lngRow, .ColIndex("ʵ������"))
        If chk�Ƿ��ֵ.value = 0 Then
            txtEdit(mtxtIdx.idx_txt��ֵ����).Text = ""
            txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = ""
            txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = ""
            Call SetDefaultValue(False)
        Else
            txtEdit(mtxtIdx.idx_txt��ֵ����).Text = .TextMatrix(lngRow, .ColIndex("��ֵ����"))
            txtEdit(mtxtIdx.idx_txt���γ�ֵ).Text = .TextMatrix(lngRow, .ColIndex("���γ�ֵ"))
            txtEdit(mtxtIdx.idx_txtʵ�ʳ�ֵ�ɿ�).Text = .TextMatrix(lngRow, .ColIndex("ʵ�ʳ�ֵ�ɿ�"))
        End If
        txtEdit(mtxtIdx.idx_txt��ֵ��ע).Text = .TextMatrix(lngRow, .ColIndex("��ֵ˵��"))
        txtEdit(mtxtIdx.idx_txt�ɿ���).Text = .TextMatrix(lngRow, .ColIndex("��ֵ�ɿ���"))
    End With
    mblnNoClick = False
    FromGridToCtlData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldCol <> NewCol Then
        cbsThis.RecalcLayout
    End If
    If OldRow = NewRow Then
        With vsGrid
            If Trim(txtEdit(mtxtIdx.idx_txt����).Text) <> "" Then
                Exit Sub
            End If
        End With
    End If
    If mblnNoClick Then Exit Sub
    Call FromGridToCtlData
End Sub


Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsGrid.ColIndex("��־") Then Cancel = True
End Sub

Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub


Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:lgf
    '����:2012-12-2 11:11:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, str���� As String


    If str���� = "" Then str���� = ",1,2" '1-�ֽ�,2.֧Ʊ
    str���� = Mid(str����, 2)
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����" & _
        " From ���㷽ʽ B" & _
        " Where Nvl(B.����,1) In(" & str���� & ")" & _
        " Order by B.����"

    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboStyle
        .Clear:
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboStyle.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

