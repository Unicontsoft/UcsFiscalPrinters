VERSION 5.00
Begin VB.Form frmTremolSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������� Zeka ��������"
   ClientHeight    =   6252
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8172
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTremolSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   8172
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   6
      Left            =   2250
      TabIndex        =   194
      Top             =   84
      Width           =   5775
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   11
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   46
         Top             =   4872
         Width           =   1776
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   10
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   45
         Top             =   4452
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   9
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   44
         Top             =   4032
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   8
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   43
         Top             =   3612
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   7
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   42
         Top             =   3192
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   6
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   41
         Top             =   2772
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   5
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   40
         Top             =   2352
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   3
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   38
         Top             =   1512
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   4
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   39
         Top             =   1932
         Width           =   3540
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Left            =   3948
         MaxLength       =   40
         TabIndex        =   47
         Top             =   4872
         Width           =   1692
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   2
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1092
         Width           =   3540
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   1
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   672
         Width           =   3540
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   48
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   0
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   252
         Width           =   3540
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 12:"
         Height          =   204
         Index           =   11
         Left            =   168
         TabIndex        =   253
         Top             =   4872
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 11:"
         Height          =   204
         Index           =   10
         Left            =   168
         TabIndex        =   252
         Top             =   4452
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 10:"
         Height          =   204
         Index           =   9
         Left            =   168
         TabIndex        =   251
         Top             =   4032
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 9:"
         Height          =   204
         Index           =   8
         Left            =   168
         TabIndex        =   250
         Top             =   3612
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 8:"
         Height          =   204
         Index           =   7
         Left            =   168
         TabIndex        =   249
         Top             =   3192
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 7:"
         Height          =   204
         Index           =   6
         Left            =   168
         TabIndex        =   248
         Top             =   2772
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 6:"
         Height          =   204
         Index           =   5
         Left            =   168
         TabIndex        =   247
         Top             =   2352
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 5:"
         Height          =   264
         Index           =   4
         Left            =   168
         TabIndex        =   220
         Top             =   1932
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 4:"
         Height          =   264
         Index           =   3
         Left            =   168
         TabIndex        =   198
         Top             =   1512
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 3:"
         Height          =   264
         Index           =   2
         Left            =   168
         TabIndex        =   197
         Top             =   1092
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 2:"
         Height          =   264
         Index           =   1
         Left            =   168
         TabIndex        =   196
         Top             =   672
         Width           =   1944
         WordWrap        =   -1  'True
      End
      Begin VB.Label labPmtType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 1:"
         Height          =   264
         Index           =   0
         Left            =   168
         TabIndex        =   195
         Top             =   252
         Width           =   1944
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   5
      Left            =   2250
      TabIndex        =   180
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtInvCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   990
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   36
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtInvEnd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   35
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtInvStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   34
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   183
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   182
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   181
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   0
      Left            =   2250
      TabIndex        =   125
      Top             =   90
      Width           =   5775
      Begin VB.CheckBox chkConnectRemember 
         Caption         =   "����������� ���������"
         Height          =   195
         Left            =   1620
         TabIndex        =   3
         Top             =   2070
         Width           =   2985
      End
      Begin VB.ComboBox cobConnectPort 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   1080
         Width           =   1635
      End
      Begin VB.ComboBox cobConnectSpeed 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "���������"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   4
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label labConnectCurrent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   129
         Top             =   270
         Width           =   5325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   128
         Top             =   1080
         Width           =   1785
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   195
         Left            =   180
         TabIndex        =   127
         Top             =   1530
         Width           =   1785
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   2
      Left            =   2250
      TabIndex        =   126
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtTaxOperPass 
         Height          =   285
         Left            =   2070
         TabIndex        =   14
         Text            =   "0"
         Top             =   5220
         Width           =   1095
      End
      Begin VB.TextBox txtTaxRegDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtTaxRegNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2070
         TabIndex        =   13
         Top             =   4590
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2070
         TabIndex        =   12
         Top             =   4230
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2070
         TabIndex        =   11
         Top             =   3870
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2070
         TabIndex        =   10
         Top             =   3510
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   15
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2070
         TabIndex        =   9
         Top             =   3150
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2070
         TabIndex        =   8
         Top             =   2790
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   7
         Top             =   2430
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2070
         TabIndex        =   6
         Top             =   2070
         Width           =   1095
      End
      Begin VB.TextBox txtTaxDecimals 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   5
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox txtTaxSerNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtTaxMemModule 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         Height          =   195
         Left            =   180
         TabIndex        =   246
         Top             =   5220
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� �����������:"
         Height          =   195
         Left            =   180
         TabIndex        =   240
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���. ����� � ���:"
         Height          =   195
         Left            =   180
         TabIndex        =   239
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   238
         Top             =   4590
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   237
         Top             =   4590
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   236
         Top             =   4230
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   235
         Top             =   4230
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   234
         Top             =   3870
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   233
         Top             =   3870
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   232
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   231
         Top             =   3510
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   144
         Top             =   3150
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   143
         Top             =   2790
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   142
         Top             =   2430
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   141
         Top             =   2070
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   140
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   139
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   138
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   137
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   136
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   133
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �� ������:"
         Height          =   195
         Left            =   180
         TabIndex        =   130
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   3
      Left            =   2250
      TabIndex        =   145
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdDateTransfer 
         Caption         =   "�� ��������"
         Height          =   375
         Left            =   2070
         TabIndex        =   18
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Timer tmrDate 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4860
         Top             =   540
      End
      Begin VB.TextBox txtDateCompDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txtDateCompTime 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox txtDateDate 
         Height          =   285
         Left            =   2070
         TabIndex        =   16
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtDateTime 
         Height          =   285
         Left            =   2070
         TabIndex        =   17
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   19
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ���:"
         Height          =   195
         Left            =   180
         TabIndex        =   149
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   148
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   147
         Top             =   1170
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� ���:"
         Height          =   195
         Left            =   180
         TabIndex        =   146
         Top             =   1530
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   4
      Left            =   2250
      TabIndex        =   152
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtHeadPingTimeout 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   32
         Top             =   4860
         Width           =   825
      End
      Begin VB.TextBox txtHeadRowChars 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   31
         Top             =   4500
         Width           =   825
      End
      Begin VB.ComboBox cobHeadBulstatName 
         Height          =   315
         Left            =   2070
         TabIndex        =   22
         Top             =   990
         Width           =   1635
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   0
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   30
         Top             =   3510
         Width           =   3525
      End
      Begin VB.TextBox txtHeadBulstatText 
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   23
         Top             =   990
         Width           =   1635
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   6
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   27
         Top             =   2430
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   8
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   29
         Top             =   3150
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   7
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   28
         Top             =   2790
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   5
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   26
         Top             =   2070
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   4
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   25
         Top             =   1710
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   3
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   24
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   1
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   20
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   2
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   21
         Top             =   630
         Width           =   3525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   33
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label76 
         Caption         =   "����� �� ���� �����:"
         Height          =   285
         Left            =   180
         TabIndex        =   242
         Top             =   4860
         Width           =   1905
      End
      Begin VB.Label Label75 
         Caption         =   "���� ������:"
         Height          =   285
         Left            =   180
         TabIndex        =   241
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screensaver:"
         Height          =   195
         Left            =   180
         TabIndex        =   221
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Left            =   3690
         TabIndex        =   162
         Top             =   990
         Width           =   285
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   195
         Left            =   180
         TabIndex        =   161
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 6:"
         Height          =   195
         Left            =   180
         TabIndex        =   160
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer:"
         Height          =   195
         Left            =   180
         TabIndex        =   159
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 7:"
         Height          =   195
         Left            =   180
         TabIndex        =   158
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 5:"
         Height          =   195
         Left            =   180
         TabIndex        =   157
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 4:"
         Height          =   195
         Left            =   180
         TabIndex        =   156
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 3:"
         Height          =   195
         Left            =   180
         TabIndex        =   155
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   154
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   153
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5685
      Index           =   11
      Left            =   2250
      TabIndex        =   229
      Top             =   180
      Width           =   5775
      Begin VB.CheckBox chkParam 
         Caption         =   "������ � ��������"
         Height          =   348
         Index           =   2
         Left            =   168
         TabIndex        =   80
         Top             =   1008
         Width           =   5472
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   14
         Left            =   4320
         TabIndex        =   87
         Top             =   5130
         Width           =   1275
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "������ � ���� �������� � ���������� ���"
         Height          =   348
         Index           =   9
         Left            =   168
         TabIndex        =   86
         Top             =   3024
         Width           =   5472
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "����� ��� � ����� �����"
         Height          =   348
         Index           =   7
         Left            =   168
         TabIndex        =   85
         Top             =   2688
         Width           =   5472
      End
      Begin VB.TextBox txtParamCashNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   78
         Top             =   270
         Width           =   825
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "��������� �������"
         Height          =   348
         Index           =   4
         Left            =   168
         TabIndex        =   82
         Top             =   1680
         Width           =   5472
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "����� �� ����"
         Height          =   348
         Index           =   1
         Left            =   168
         TabIndex        =   79
         Top             =   672
         Width           =   5472
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "����������� �����"
         Height          =   348
         Index           =   3
         Left            =   168
         TabIndex        =   81
         Top             =   1344
         Width           =   5472
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "������ � ������"
         Height          =   348
         Index           =   6
         Left            =   168
         TabIndex        =   84
         Top             =   2352
         Width           =   5472
      End
      Begin VB.CheckBox chkParam 
         Caption         =   "�������� ����� ��������"
         Height          =   348
         Index           =   5
         Left            =   168
         TabIndex        =   83
         Top             =   2016
         Width           =   5472
      End
      Begin VB.Label Label42 
         Caption         =   "No. �� ����:"
         Height          =   285
         Left            =   180
         TabIndex        =   230
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   13
      Left            =   2250
      TabIndex        =   184
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtCashComment 
         Height          =   285
         Left            =   2070
         TabIndex        =   92
         Top             =   2700
         Width           =   3525
      End
      Begin VB.TextBox txtCashOperPass 
         Height          =   285
         Left            =   2070
         TabIndex        =   94
         Text            =   "0"
         Top             =   5220
         Width           =   1095
      End
      Begin VB.TextBox txtCashOperNo 
         Height          =   285
         Left            =   2070
         TabIndex        =   93
         Text            =   "1"
         Top             =   4860
         Width           =   1095
      End
      Begin VB.ComboBox cobCashPayment 
         Height          =   315
         Left            =   2070
         TabIndex        =   88
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton optCashOut 
         Caption         =   "���������"
         Height          =   285
         Left            =   3528
         TabIndex        =   90
         Top             =   1980
         Width           =   1455
      End
      Begin VB.OptionButton optCashIn 
         Caption         =   "���������"
         Height          =   285
         Left            =   2070
         TabIndex        =   89
         Top             =   1980
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtCashSum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   91
         Top             =   2340
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����/�����"
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   95
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtCashOut 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtCashIn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   990
         Width           =   1545
      End
      Begin VB.TextBox txtCashTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   180
         TabIndex        =   245
         Top             =   2700
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         Height          =   195
         Left            =   180
         TabIndex        =   244
         Top             =   5220
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   180
         TabIndex        =   243
         Top             =   4860
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� �������:"
         Height          =   195
         Left            =   180
         TabIndex        =   228
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   180
         TabIndex        =   191
         Top             =   2340
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   190
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   188
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   186
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   14
      Left            =   2250
      TabIndex        =   192
      Top             =   90
      Width           =   5775
      Begin VB.CheckBox chkReportPayments2 
         Caption         =   "��������"
         Height          =   285
         Left            =   2520
         TabIndex        =   109
         Top             =   3330
         Width           =   1545
      End
      Begin VB.CheckBox chkReportPayments1 
         Caption         =   "��������"
         Height          =   285
         Left            =   2520
         TabIndex        =   104
         Top             =   2070
         Width           =   1545
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "����� �� ���"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   113
         Top             =   4860
         Width           =   5145
      End
      Begin VB.CheckBox chkReportOperClosure 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   111
         Top             =   4140
         Width           =   1725
      End
      Begin VB.CheckBox chkReportDepartments 
         Caption         =   "������������"
         Height          =   285
         Left            =   2520
         TabIndex        =   98
         Top             =   630
         Width           =   1905
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "����� ���������"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   110
         Top             =   3780
         Width           =   5145
      End
      Begin VB.CheckBox chkReportDetailed1 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   103
         Top             =   2070
         Width           =   1725
      End
      Begin VB.TextBox txtReportStart 
         Height          =   285
         Left            =   1800
         TabIndex        =   101
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox txtReportEnd 
         Height          =   285
         Left            =   3420
         TabIndex        =   102
         Top             =   1710
         Width           =   1095
      End
      Begin VB.CheckBox chkReportDetailed2 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   108
         Top             =   3330
         Width           =   1725
      End
      Begin VB.TextBox txtReportTD 
         Height          =   285
         Left            =   3420
         TabIndex        =   107
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtReportFD 
         Height          =   285
         Left            =   1800
         TabIndex        =   106
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   114
         Top             =   5220
         Width           =   1275
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "��������� �����"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   112
         Top             =   4500
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ���� �� �����"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   105
         Top             =   2520
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ����� �� �����"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   100
         Top             =   1350
         Width           =   5145
      End
      Begin VB.CheckBox chkReportItems 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   97
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportClosure 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   99
         Top             =   990
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "������ �������� �����"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   96
         Top             =   270
         Value           =   -1  'True
         Width           =   5145
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �����:"
         Height          =   195
         Left            =   900
         TabIndex        =   202
         Top             =   1710
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��: "
         Height          =   195
         Left            =   2520
         TabIndex        =   201
         Top             =   1710
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��: "
         Height          =   195
         Left            =   2520
         TabIndex        =   200
         Top             =   2880
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� ����:"
         Height          =   300
         Left            =   900
         TabIndex        =   199
         Top             =   2880
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   16
      Left            =   2250
      TabIndex        =   193
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdStatusReset 
         Caption         =   "�����"
         Height          =   375
         Left            =   168
         TabIndex        =   116
         Top             =   5208
         Width           =   1185
      End
      Begin VB.ListBox lstStatus 
         Height          =   4920
         IntegralHeight  =   0   'False
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   115
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   9
      Left            =   2250
      TabIndex        =   204
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdItemLoad 
         Caption         =   "���������"
         Height          =   375
         Left            =   2610
         TabIndex        =   59
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox txtItemDep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         TabIndex        =   65
         Top             =   2250
         Width           =   825
      End
      Begin VB.TextBox txtItemReport 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   3510
         Width           =   1545
      End
      Begin VB.TextBox txtItemTime 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   3870
         Width           =   1545
      End
      Begin VB.CommandButton cmdItemDelete 
         Caption         =   "���������"
         Height          =   375
         Left            =   4320
         TabIndex        =   66
         Top             =   4770
         Width           =   1275
      End
      Begin VB.CommandButton cmdItemNew 
         Caption         =   "���"
         Height          =   375
         Left            =   3960
         TabIndex        =   60
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox txtItemSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1545
      End
      Begin VB.TextBox txtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   62
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox txtItemPLU 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         MaxLength       =   5
         TabIndex        =   61
         Top             =   810
         Width           =   825
      End
      Begin VB.TextBox txtItemAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1545
      End
      Begin VB.ComboBox cobItemGroup 
         Height          =   315
         Left            =   1620
         TabIndex        =   63
         Top             =   1530
         Width           =   825
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   1620
         TabIndex        =   64
         Top             =   1890
         Width           =   3975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   13
         Left            =   4320
         TabIndex        =   67
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtItemNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   58
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������:"
         Height          =   195
         Left            =   180
         TabIndex        =   223
         Top             =   2250
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����. �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   222
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� ��:"
         Height          =   195
         Left            =   180
         TabIndex        =   212
         Top             =   3870
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         Height          =   195
         Left            =   180
         TabIndex        =   211
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   180
         TabIndex        =   210
         Top             =   1170
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLU:"
         Height          =   195
         Left            =   180
         TabIndex        =   209
         Top             =   810
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   208
         Top             =   1530
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         Height          =   195
         Left            =   180
         TabIndex        =   207
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   180
         TabIndex        =   206
         Top             =   1890
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   195
         Left            =   180
         TabIndex        =   205
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   17
      Left            =   2250
      TabIndex        =   203
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtLog 
         Height          =   5505
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   117
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   10
      Left            =   2250
      TabIndex        =   173
      Top             =   90
      Width           =   5775
      Begin VB.ComboBox cobLogoActive 
         Height          =   315
         Left            =   3780
         TabIndex        =   69
         Top             =   270
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "���������"
         Height          =   1815
         Left            =   180
         TabIndex        =   217
         Top             =   810
         Width           =   5415
         Begin VB.TextBox txtLogoIndex 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   76
            Top             =   1350
            Width           =   465
         End
         Begin VB.TextBox txtLogoHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   73
            Text            =   "140"
            Top             =   630
            Width           =   465
         End
         Begin VB.TextBox txtLogoWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   72
            Text            =   "512"
            Top             =   630
            Width           =   465
         End
         Begin VB.OptionButton optLogoStretch 
            Caption         =   "���������"
            Height          =   285
            Left            =   1800
            TabIndex        =   74
            Top             =   990
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optLogoCenter 
            Caption         =   "����������"
            Height          =   285
            Left            =   3420
            TabIndex        =   75
            Top             =   990
            Width           =   1545
         End
         Begin VB.CommandButton cmdLogoOpen 
            Caption         =   "�����"
            Height          =   375
            Left            =   3960
            TabIndex        =   71
            Top             =   270
            Width           =   1275
         End
         Begin VB.TextBox txtLogoTreshold 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   70
            Text            =   "50"
            Top             =   270
            Width           =   465
         End
         Begin VB.Label labLogoInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   2340
            TabIndex        =   227
            Top             =   1350
            Width           =   2985
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������������:"
            Height          =   195
            Left            =   180
            TabIndex        =   226
            Top             =   1350
            Width           =   1635
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            Height          =   285
            Left            =   2250
            TabIndex        =   225
            Top             =   630
            Width           =   285
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            Height          =   195
            Left            =   180
            TabIndex        =   224
            Top             =   630
            Width           =   1635
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���� �� �����:"
            Height          =   390
            Left            =   180
            TabIndex        =   219
            Top             =   270
            Width           =   1635
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   2340
            TabIndex        =   218
            Top             =   270
            Width           =   555
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox chkLogoPrint 
         Caption         =   "����� �������� ���� ����� header"
         Height          =   285
         Left            =   180
         TabIndex        =   68
         Top             =   270
         Width           =   4245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   77
         Top             =   5220
         Width           =   1275
      End
      Begin VB.PictureBox picLogoScroll 
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   180
         ScaleHeight     =   2448
         ScaleWidth      =   5412
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   2700
         Width           =   5415
         Begin VB.HScrollBar scbLogoHor 
            CausesValidation=   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   216
            TabStop         =   0   'False
            Top             =   2070
            Width           =   5415
         End
         Begin VB.PictureBox picLogo 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   0
            ScaleHeight     =   1632
            ScaleWidth      =   5232
            TabIndex        =   215
            TabStop         =   0   'False
            Top             =   0
            Width           =   5235
         End
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   8
      Left            =   2250
      TabIndex        =   164
      Top             =   90
      Width           =   5775
      Begin VB.ComboBox cobDepGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   55
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtDepTotalSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1545
      End
      Begin VB.TextBox txtDepSales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1545
      End
      Begin VB.TextBox txtDepNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.ListBox lstDeps 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   54
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   57
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtDepName 
         Height          =   285
         Left            =   2610
         TabIndex        =   56
         Top             =   1260
         Width           =   2985
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� �� ����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   172
         Top             =   2700
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   170
         Top             =   2340
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� �����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   168
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   167
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   165
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   7
      Left            =   2250
      TabIndex        =   174
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtOperPass2 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   52
         Top             =   4410
         Width           =   1545
      End
      Begin VB.TextBox txtOperPass 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   51
         Top             =   4050
         Width           =   1545
      End
      Begin VB.TextBox txtOperName 
         Height          =   285
         Left            =   2610
         TabIndex        =   50
         Top             =   900
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   53
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstOpers 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   49
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtOperNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   179
         Top             =   4410
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� ������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   178
         Top             =   4050
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   177
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   176
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   90
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   180
      Width           =   1275
   End
   Begin VB.ListBox lstCmds 
      Height          =   5685
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   0
      Top             =   168
      Width           =   2085
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   90
      TabIndex        =   163
      Top             =   5940
      Width           =   7935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTremolSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFP20 (c) 2008-2019 by wqweto@gmail.com
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Nastrojki na FP po Zeka protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmTremolSetup"

'=========================================================================
' API
'=========================================================================

Private Const EM_SCROLLCARET            As Long = &HB7

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const CAP_MSG               As String = "��������� Tremol ��������"
Private Const LNG_NUM_DEPS          As Long = 50
Private Const LNG_NUM_OPERS         As Long = 50
Private Const PROGID_PROTOCOL       As String = LIB_NAME & ".cTremolProtocol"
'--- strings
Private Const STR_SPEEDS            As String = "9600|19200"
Private Const STR_COMMANDS          As String = "������ �������|���������|    ������� ����������|    ���� � ���|    �������|    ������ �� �������|    ������ ��������|    ���������|    ������������|    ��������|    �������� ����|    ���������|��������|    ���������/���������|    ����� ������|�������������|    ������|    ������ �����������"
Private Const STR_GROUPS            As String = "�|�|�|�"
Private Const STR_BULSTAT_NAME      As String = "���|���|���|�������� No"
Private Const STR_STATUS_CONNECTING As String = "���������..."
Private Const STR_STATUS_SUCCESS_CONNECT As String = "������� %1"
Private Const STR_STATUS_FAILURE_CONNECT As String = "���� ������"
Private Const STR_STATUS_SAVING     As String = "���������..."
Private Const STR_STATUS_SUCCESS_SAVE As String = "������� ��������� �� %1 (%2 ���.)"
Private Const STR_STATUS_FETCHING   As String = "����������..."
Private Const STR_STATUS_SUCCESS_FETCH As String = "������� ���������� �� %1 (%2 ���.)"
Private Const STR_STATUS_NOT_IMPLEMENTED As String = "�� � �����������"
Private Const STR_STATUS_FETCH_DEP  As String = "���������� ����������� %1 �� " & LNG_NUM_DEPS & "..."
Private Const STR_STATUS_ENUM_PORTS As String = "���������� �� ������� ��������..."
Private Const STR_STATUS_FETCH_OPER As String = "���������� �������� %1 �� " & LNG_NUM_OPERS & "..."
Private Const STR_STATUS_PRINT      As String = "�����������..."
Private Const STR_STATUS_RESETTING  As String = "�����..."
Private Const STR_LOGO_DIMENSIONS   As String = "������ � �������: %1"
Private Const STR_LOGO_ASSIGNED     As String = " - �����������"
Private Const STR_FP_STATUSES       As String = "ST0.0 - �� �������� ���� �� ������ (��� ST3.0, ST3.1 ��� ST3.2 = 1)|ST0.1 - ����� � ���������� �� ������������ ��� ������� ����. ���|ST0.2 - ������� �������|ST0.3 - �������� ��������|ST0.4 - ���������� ����|ST0.5 - ������ � ������������ ����� RAM|ST0.6 - ��������� ������ � ���������|ST1.0 - ������ ������|ST1.1 - ���������� � ������������� �� ��������|ST1.2 - �������� �� ������������|ST1.3 - ������� ������ �����|ST1.4 - ������� ��������� �����|ST1.5 - ������� ����������� �����|ST1.6 - �� � ��������� ��������|" & _
                                                "ST2.0 - ������� �������� ���|ST2.1 - ������� �������� ���|ST2.2 - ���������� ����� ���|ST2.3 - � ��� � ����|ST2.4 - ���� � �������|ST2.5 - ����������|ST2.6 - ����������|ST3.0 - ������ ��|ST3.1 - ������ ��� ��|ST3.2 - ����� ��|ST3.3 - ������� 50 ��� ��-����� �������� ������ ��� ��|ST3.4 - ������ �� �������: ������ = 1, ���� = 0|ST3.5 - ������������|ST3.6 - �������� ������������ ����� �� ��� � ����� �� ��|" & _
                                                "ST4.0 - ����������� ��������� �� ����|ST4.1 - ��������� �������|ST4.2 - ������� �� �����������: 9600 = 1; 19200 = 0|ST4.3 - ��������� �� ��� �� ��: �������� = 1; ��������� = 0|ST4.4 - ����������� �������� �� ��������|ST4.5 - � ����� �� ���� � ����|ST4.6 - ����� �� ���� ��� ������ � ����|ST5.0 - ������ ��� �����|ST5.1 - ����������, ���� ��|ST5.2 - �� �� � ������� ������|ST5.3 - ����������|ST5.4 - ����������|ST5.5 � ������ ����|ST5.6 - �� � ��������������"
'--- messages
Private Const MSG_PASSWORDS_MISMATCH As String = "�������� �� ��������"
Private Const MSG_CANNOT_ACCESS_PRINTER_PROXY As String = "������ ��� ��������� �� ��������� �� ������ �� �������� ������� %1." & vbCrLf & vbCrLf & "%2"
Private Const MSG_ERROR_LOADING_IMAGE As String = "������ ��� ��������� �� �����������"
Private Const MSG_PASSWORD_ALREADY_USED As String = "�������� ���� �� �������� �� ���� ��������"
Private Const MSG_INVALID_LOGO_NO   As String = "���� �������� ����� �� ���� �� ������������"
'--- defaults
Private Const DEF_PING_TIMEOUT      As Long = 200
Private Const DEF_COMMENT_LEN       As Long = 30

Private m_oFP                   As cTremolProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private WithEvents m_oFPSink    As cTremolProtocol
Attribute m_oFPSink.VB_VarHelpID = -1
Private m_sLog                  As String
Private m_vDeps                 As Variant
Private m_vOpers                As Variant
Private m_vItems                As Variant
Private m_vAdminCash            As Variant
Private m_hLogo                 As Long
Private m_baLogoBW()            As Byte
Private m_lTimeout              As Long

Private Enum UcsCommands
    ucsCmdConnect
    ucsCmdSettings
        ucsCmdTaxInfo
        ucsCmdDateTime
        ucsCmdHeaderFooter
        ucsCmdInvoiceNo
        ucsCmdPaymentTypes
        ucsCmdOperators
        ucsCmdDepartments
        ucsCmdItems
        ucsCmdGraphicalLogo
        ucsCmdParameters
    ucsCmdOperations
        ucsCmdCashOper
        ucsCmdReports
    ucsCmdAdmin
        ucsCmdStatus
        ucsCmdLog
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    MsgBox MODULE_NAME & "." & sFunc & "(" & Erl & ")" & ": " & Err.Description, vbCritical
End Sub

'=========================================================================
' Properties
'=========================================================================

Private Property Get pvStatus() As String
    pvStatus = labStatus.Caption
End Property

Private Property Let pvStatus(sValue As String)
    labStatus.Caption = sValue
    Call UpdateWindow(Me.hWnd)
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit(DeviceString As String, sServer As String, OwnerForm As Object) As Boolean
    Const FUNC_NAME     As String = "frInit"
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sError          As String
    
    On Error GoTo EH
    vSplit = Split(DeviceString, ";")
    m_lTimeout = C_Lng(At(vSplit, 2))
    vSplit = Split(At(vSplit, 1), ",")
    Set m_oFP = pvGetPrinter(sServer, sError)
    If m_oFP Is Nothing Then
        If LenB(sError) <> 0 Then
            MsgBox Printf(MSG_CANNOT_ACCESS_PRINTER_PROXY, At(vSplit, 0) & IIf(LenB(sServer) <> 0, "@" & sServer, vbNullString), sError), vbExclamation
        End If
        GoTo QH
    End If
    On Error Resume Next '--- checked
    Set m_oFPSink = m_oFP
    On Error GoTo EH
    '--- init UI
    FixThemeSupport Controls
    For Each vElem In Split(STR_COMMANDS, "|")
        lstCmds.AddItem vElem
    Next
    For Each vElem In Split(STR_GROUPS, "|")
        cobDepGroup.AddItem vElem
        cobItemGroup.AddItem vElem
    Next
    For Each vElem In Split(STR_BULSTAT_NAME, "|")
        cobHeadBulstatName.AddItem vElem
    Next
    For lIdx = fraCommands.LBound To fraCommands.UBound
        If DispInvoke(fraCommands(lIdx), "Index", VbGet) Then
            fraCommands(lIdx).Visible = False
        End If
    Next
    cmdExit.Left = -cmdExit.Width
    '--- login
    pvStatus = STR_STATUS_ENUM_PORTS
    cobConnectPort.Clear
    For Each vElem In EnumSerialPorts
        cobConnectPort.AddItem vElem
    Next
    cobConnectPort.Text = At(vSplit, 0)
    chkConnectRemember.Value = -(LenB(cobConnectPort.Text) <> 0)
    If cobConnectPort.ListCount > 0 And Len(cobConnectPort.Text) = 0 Then
        cobConnectPort.ListIndex = 0
    End If
    cobConnectSpeed.Clear
    For Each vElem In Split(STR_SPEEDS, "|")
        cobConnectSpeed.AddItem vElem
    Next
    cobConnectSpeed.Text = At(vSplit, 1)
    If cobConnectSpeed.ListCount > 0 And Len(cobConnectSpeed.Text) = 0 Then
        cobConnectSpeed.ListIndex = 0
    End If
    labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
    lstCmds.ListIndex = ucsCmdConnect
    If chkConnectRemember.Value Then
        cmdSave(ucsCmdConnect).Value = True
    End If
    If OwnerForm Is Nothing Then
        Show vbModal
    Else
        On Error Resume Next '--- checked
        Show vbModal, OwnerForm
        If Err.Number <> 0 Then
            Show vbModal
        End If
        On Error GoTo EH
    End If
    '--- success
    frInit = True
QH:
    Unload Me
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetPrinter(sServer As String, sError As String) As cTremolProtocol
    On Error Resume Next '--- checked
    If LenB(sServer) = 0 Then
        With New cFiscalAdmin
            Set pvGetPrinter = .CreateObject(PROGID_PROTOCOL)
        End With
    Else
        Set pvGetPrinter = CreateObject(PROGID_PROTOCOL, sServer)
    End If
    sError = Err.Description
    On Error GoTo 0
End Function

Private Function pvFetchData(ByVal eCmd As UcsCommands) As Boolean
    Const FUNC_NAME     As String = "pvFetchData"
    Dim lIdx            As Long
    Dim vResult         As Variant
    Dim sText           As String
    Dim vElem           As Variant
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCmd <> ucsCmdConnect Then
        pvStatus = labConnectCurrent.Caption
        Exit Function
    End If
    Select Case eCmd
    Case ucsCmdConnect
        pvStatus = labConnectCurrent.Caption
    Case ucsCmdTaxInfo
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoDiagnostics, "0"), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        txtTaxMemModule.Text = At(vResult, 0)
        txtTaxSerNo.Text = At(vResult, 1)
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoBulstat), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        txtTaxRegNo.Text = At(vResult, 2)
        txtTaxRegDate.Text = At(vResult, 3)
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoDecimals), ";")
        txtTaxDecimals.Text = C_Lng(At(vResult, 0))
        vResult = Split(Replace(m_oFP.SendCommand(ucsZekCmdInfoTaxRates), "%", vbNullString), ";")
        For lIdx = 0 To 7
            txtTaxGroup(lIdx).Text = C_Lng(At(vResult, lIdx))
        Next
    Case ucsCmdDateTime
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoDateTime), " ")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        txtDateDate.Text = At(vResult, 0)
        txtDateTime.Text = At(vResult, 1)
        tmrDate_Timer
    Case ucsCmdHeaderFooter
        For lIdx = 0 To 8
            txtHeadHeader(lIdx).Text = Trim$(At(Split(m_oFP.SendCommand(ucsZekCmdInfoHeaderFooter, C_Str(lIdx)), ";"), 1))
            If LenB(m_oFP.LastError) <> 0 Then
                GoTo QH
            End If
        Next
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoBulstat), ";")
        cobHeadBulstatName.ListIndex = C_Lng(At(vResult, 1))
        txtHeadBulstatText.Text = Trim$(At(vResult, 0))
        txtHeadRowChars.Text = pvZfplibValue("FPLineWidth", DEF_COMMENT_LEN)
        txtHeadPingTimeout.Text = pvZfplibValue("PingTimeout", DEF_PING_TIMEOUT)
    Case ucsCmdInvoiceNo
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoInvoiceNo), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        txtInvStart.Text = C_Dbl(At(vResult, 0))
        txtInvEnd.Text = C_Dbl(At(vResult, 1))
        txtInvCurrent.Text = C_Dbl(At(vResult, 0))
    Case ucsCmdPaymentTypes
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoPaymentTypes), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        For lIdx = 0 To Limit(UBound(vResult), , txtPmtType.UBound + 1)
            If IsNumeric(vResult(lIdx)) Then
                Exit For
            End If
            txtPmtType(lIdx).Text = Trim$(vResult(lIdx))
            txtPmtType(lIdx).Tag = Trim$(vResult(lIdx))
        Next
        txtPmtRate.Text = Trim$(vResult(lIdx))
        txtPmtRate.Tag = Trim$(vResult(lIdx))
        txtPmtType(lIdx - 1).Width = txtPmtRate.Left - txtPmtType(lIdx - 1).Left - 60
        txtPmtRate.Top = txtPmtType(lIdx - 1).Top
        For lIdx = lIdx To txtPmtType.UBound
            txtPmtType(lIdx).Visible = False
            labPmtType(lIdx).Visible = False
        Next
    Case ucsCmdOperators
        vResult = m_oFP.SendCommand(ucsZekCmdInfoStatus)
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        If Not IsArray(m_vOpers) Then
            ReDim m_vOpers(0 To LNG_NUM_OPERS) As Variant
        End If
        For lIdx = 1 To UBound(m_vOpers)
            If Not IsArray(m_vOpers(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_OPER, lIdx)
                m_vOpers(lIdx) = Split(m_oFP.SendCommand(ucsZekCmdInfoOperator, C_Str(lIdx)), ";")
                If LenB(m_oFP.LastError) Then
                    ReDim Preserve m_vOpers(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            If lstOpers.ListCount < lIdx Then
                lstOpers.AddItem vbNullString
            End If
            sText = lIdx & ": " & At(m_vOpers(lIdx), 1)
            If lstOpers.List(lIdx - 1) <> sText Then
                lstOpers.List(lIdx - 1) = sText
            End If
        Next
        lstOpers_Click
        pvStatus = vbNullString
    Case ucsCmdDepartments
        vResult = m_oFP.SendCommand(ucsZekCmdInfoStatus)
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        If Not IsArray(m_vDeps) Then
            ReDim m_vDeps(0 To LNG_NUM_DEPS) As Variant
        End If
        For lIdx = 1 To UBound(m_vDeps)
            If Not IsArray(m_vDeps(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_DEP, lIdx)
                m_vDeps(lIdx) = Split(m_oFP.SendCommand(ucsZekCmdInfoDepartment, Right$("0" & C_Str(lIdx), 2)), ";")
                If LenB(m_oFP.LastError) <> 0 Then
                    ReDim Preserve m_vDeps(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            If lstDeps.ListCount < lIdx Then
                lstDeps.AddItem vbNullString
            End If
            vResult = m_vDeps(lIdx)
            sText = At(vResult, 1) & " (" & At(vResult, 0) & ")"
            If lstDeps.List(lIdx - 1) <> sText Then
                lstDeps.List(lIdx - 1) = sText
            End If
        Next
        lstDeps_Click
        pvStatus = vbNullString
    Case ucsCmdItems
        '--- do nothing
    Case ucsCmdGraphicalLogo
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoParameters), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        chkLogoPrint.Value = IIf(At(vResult, 1) = "1", vbChecked, vbUnchecked)
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoLogo, "?"), ";")
        cobLogoActive.Clear
        For lIdx = 1 To Len(At(vResult, 1))
            cobLogoActive.AddItem (lIdx - 1) & IIf(Mid$(At(vResult, 1), lIdx, 1) = "1", STR_LOGO_ASSIGNED, vbNullString)
        Next
        On Error Resume Next '--- checked
        cobLogoActive.ListIndex = C_Lng(At(vResult, 0))
        On Error GoTo EH
        pvApplyLogo
    Case ucsCmdParameters
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoParameters), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        txtParamCashNo.Text = C_Dbl(At(vResult, 0))
        For lIdx = 1 To 9
            If lIdx <> 8 Then
                chkParam(lIdx).Enabled = LenB(At(vResult, lIdx)) <> 0
                chkParam(lIdx).Value = -(At(vResult, lIdx) = "1")
            End If
        Next
    Case ucsCmdCashOper
        ReDim m_vAdminCash(0 To 2) As Variant
        m_vAdminCash(0) = Split(m_oFP.SendCommand(ucsZekCmdInfoRegisters, "0"), ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        m_vAdminCash(1) = Split(m_oFP.SendCommand(ucsZekCmdInfoRegisters, "2"), ";")
        m_vAdminCash(2) = Split(m_oFP.SendCommand(ucsZekCmdInfoRegisters, "3"), ";")
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoPaymentTypes), ";")
        cobCashPayment.Clear
        For Each vElem In vResult
            If LenB(Trim$(vElem)) = 0 Then
                Exit For
            End If
            cobCashPayment.AddItem Trim$(vElem)
        Next
        cobCashPayment.ListIndex = 0
    Case ucsCmdReports
        '--- do nothing
    Case ucsCmdStatus
        If lstStatus.ListCount = 0 Then
            For Each vElem In Split(STR_FP_STATUSES, "|")
                lstStatus.AddItem vElem
            Next
        End If
        vResult = Left$(m_oFP.SendCommand(ucsZekCmdInfoStatus) & String$(6, 0), 6)
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        lstStatus.Tag = vbNullString
        For lIdx = 0 To lstStatus.ListCount - 1
            lstStatus.Selected(lIdx) = ((Asc(Mid$(vResult, 1 + lIdx \ 7)) And 2 ^ (lIdx Mod 7)) <> 0)
            lIdx = lIdx + 1
        Next
        lstStatus.ListIndex = 0
        lstStatus.Tag = "Locked"
    Case ucsCmdLog
        m_sLog = Right$(m_sLog, 32000)
        txtLog.Text = m_sLog
        txtLog.SelStart = Len(m_sLog)
        pvStatus = labConnectCurrent.Caption
    Case Else
        pvStatus = STR_STATUS_NOT_IMPLEMENTED
        Exit Function
    End Select
    '--- success
    pvFetchData = True
    Exit Function
QH:
    If LenB(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
    End If
    Exit Function
EH:
    If pvShowError() Then
        Exit Function
    End If
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvSaveData(ByVal eCommand As UcsCommands) As Boolean
    Const FUNC_NAME     As String = "pvSaveData"
    Dim vResult         As Variant
    Dim eCmd            As UcsTremolCommandsEnum
    Dim lIdx            As Long
    Dim sData           As String
    Dim dDate           As Date
    Dim lSize           As Long
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCommand <> ucsCmdConnect Then
        Exit Function
    End If
    Select Case eCommand
    Case ucsCmdConnect
        pvStatus = STR_STATUS_CONNECTING
        If m_oFP.Init(cobConnectPort.Text & "," & C_Lng(cobConnectSpeed.Text), m_lTimeout) Then
            On Error Resume Next '--- checked
            m_oFP.SendCommand ucsZekCmdInfoTransaction
            If pvShowError() Then
                On Error GoTo EH
                labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
                Caption = CAP_MSG
            Else
                On Error GoTo EH
                labConnectCurrent.Caption = Printf(STR_STATUS_SUCCESS_CONNECT, m_oFP.Device)
                Caption = m_oFP.Device & " - " & CAP_MSG
                '--- flush cache
                m_vDeps = Empty
                m_vOpers = Empty
                m_vItems = Empty
                m_vAdminCash = Empty
                m_sLog = vbNullString
                lstCmds.ListIndex = ucsCmdTaxInfo
            End If
        Else
            labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
            Caption = CAP_MSG
        End If
    Case ucsCmdTaxInfo
        m_oFP.SendCommand ucsZekCmdInitDecimals, Pad(txtTaxOperPass.Text, 6) & ";" & txtTaxDecimals.Text
        If LenB(m_oFP.LastError) <> 0 Then
            MsgBox m_oFP.LastError, vbExclamation
        End If
        For lIdx = 0 To 7
            sData = sData & ";" & txtTaxGroup(lIdx).Text & "%"
        Next
        m_oFP.SendCommand ucsZekCmdInitTaxRates, Pad(txtTaxOperPass.Text, 6) & sData
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
    Case ucsCmdDateTime
        dDate = C_Date(txtDateDate.Text) + C_Date(txtDateTime.Text)
        m_oFP.SendCommand ucsZekCmdInitDateTime, Format$(dDate, "dd\-MM\-yy") & " " & Format$(dDate, "hh\:nn\:ss")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
    Case ucsCmdHeaderFooter
        For lIdx = 0 To 8
            vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoHeaderFooter, C_Str(lIdx)), ";")
            If lIdx = 0 Then
                lSize = Len(txtHeadHeader(0).Text)
            Else
                lSize = C_Lng(txtHeadRowChars.Text) + 2
            End If
            m_oFP.SendCommand ucsZekCmdInitHeaderFooter, C_Str(lIdx) & ";" & Pad(CenterText(txtHeadHeader(lIdx).Text, lSize), Len(At(vResult, 1)))
            If LenB(m_oFP.LastError) <> 0 Then
                GoTo QH
            End If
        Next
        pvZfplibSet "FPLineWidth", C_Lng(txtHeadRowChars.Text)
        pvZfplibSet "PingTimeout", C_Lng(txtHeadPingTimeout.Text)
    Case ucsCmdInvoiceNo
        m_oFP.SendCommand ucsZekCmdInitInvoiceNo, Pad(txtInvStart.Text, -10) & ";" & Pad(txtInvEnd.Text, -10)
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
    Case ucsCmdPaymentTypes
        For lIdx = 2 To 11
            If txtPmtType(lIdx).Top = txtPmtRate.Top Then
                If txtPmtType(lIdx).Text <> txtPmtType(lIdx).Tag Or txtPmtRate.Text <> txtPmtRate.Tag Then
                    m_oFP.SendCommand ucsZekCmdInitPaymentType, lIdx & ";" & Pad(txtPmtType(lIdx).Text, 10) & _
                        ";" & Format$(C_Dbl(txtPmtRate.Text), "0000.00000")
                    If LenB(m_oFP.LastError) <> 0 Then
                        GoTo QH
                    End If
                End If
            ElseIf txtPmtType(lIdx).Text <> txtPmtType(lIdx).Tag Then
                m_oFP.SendCommand ucsZekCmdInitPaymentType, lIdx & ";" & Pad(txtPmtType(lIdx).Text, 10)
                If LenB(m_oFP.LastError) <> 0 Then
                    GoTo QH
                End If
            End If
        Next
    Case ucsCmdOperators
        If txtOperPass.Text <> txtOperPass2.Text Then
            MsgBox MSG_PASSWORDS_MISMATCH, vbExclamation
            Exit Function
        End If
        m_oFP.SendCommand ucsZekCmdInitOperator, txtOperNo.Text & ";" & Pad(txtOperName.Text, 20) & ";" & Pad(txtOperPass.Text, 4)
        If m_oFP.StatusNo = 2 Then
            MsgBox MSG_PASSWORD_ALREADY_USED, vbExclamation
            GoTo QH
        End If
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        m_vOpers = Empty
    Case ucsCmdDepartments
        m_oFP.SendCommand ucsZekCmdInitDepartment, Pad(txtDepNo.Text, -2) & ";" & Pad(txtDepName.Text, 20) & ";" & cobDepGroup.Text
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        m_vDeps = Empty
    Case ucsCmdItems
        m_oFP.SendCommand ucsZekCmdInitItem, txtItemPLU.Text & ";" & Pad(txtItemName.Text, 20) & ";" & txtItemPrice.Text & ";" & cobItemGroup.Text & ";" & IIf(LenB(txtItemDep.Text) <> 0, Chr$(&H80 + C_Lng(txtItemDep.Text)), Chr$(&H80))
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        pvLoadItem Split(m_oFP.SendCommand(ucsZekCmdInfoItem, Right$(String$(5, "0") & txtItemPLU.Text, 5)), ";")
    Case ucsCmdGraphicalLogo
        If m_hLogo <> 0 Then
            If LenB(txtLogoIndex.Text) = 0 Then
                txtLogoIndex.SetFocus
                MsgBox MSG_INVALID_LOGO_NO, vbExclamation
                Exit Function
            End If
            If C_Lng(txtLogoIndex.Text) = 0 Then
                sData = StrConv(m_baLogoBW, vbUnicode)
                m_oFP.SendCommand ucsZekCmdInitLogo, sData
            Else
                sData = (C_Lng(txtLogoIndex.Text) - 1) & StrConv(m_baLogoBW, vbUnicode)
                m_oFP.SendCommand ucsZekCmdInitLogoByNum, sData
            End If
            If LenB(m_oFP.LastError) <> 0 Then
                GoTo QH
            End If
        End If
        vResult = Split(m_oFP.SendCommand(ucsZekCmdInfoParameters), ";")
        vResult(1) = IIf(chkLogoPrint.Value = vbChecked, "1", "0")
        m_oFP.SendCommand ucsZekCmdInitParameters, Join(vResult, ";")
        If LenB(m_oFP.LastError) <> 0 Then
            GoTo QH
        End If
        If cobLogoActive.ListIndex >= 0 Then
            m_oFP.SendCommand ucsZekCmdInfoLogo, cobLogoActive.ListIndex
            If LenB(m_oFP.LastError) <> 0 Then
                GoTo QH
            End If
        End If
    Case ucsCmdCashOper
        sData = txtCashOperNo.Text & ";" & Pad(txtCashOperPass.Text, 4) & ";" & cobCashPayment.ListIndex & ";" & IIf(optCashOut.Value, "-", vbNullString) & txtCashSum.Text & "@" & Left$(txtCashComment.Text, 34)
        m_oFP.SendCommand ucsZekCmdAdminCashDebitCredit, sData
        If LenB(m_oFP.LastError) Then
            GoTo QH
        End If
    Case ucsCmdParameters
        sData = Pad(txtParamCashNo.Text, -4)
        For lIdx = 1 To 9
            If lIdx <> 8 Then
                sData = sData & ";" & IIf(chkParam(lIdx).Value = vbChecked, "1", "0")
            Else
                sData = sData & ";0"
            End If
        Next
        m_oFP.SendCommand ucsZekCmdInitParameters, sData & ";"
        If LenB(m_oFP.LastError) Then
            GoTo QH
        End If
    Case ucsCmdReports
        pvStatus = STR_STATUS_PRINT
        If optReportType(0).Value Then
            eCmd = Switch(chkReportItems.Value = vbChecked, ucsZekCmdPrintReportDailyItems, _
                chkReportDepartments.Value = vbChecked, ucsZekCmdPrintReportDailyDepartments, _
                True, ucsZekCmdPrintReportDaily)
            m_oFP.SendCommand eCmd, IIf(chkReportClosure.Value = vbChecked, "Z", "X")
        ElseIf optReportType(1).Value Then
            eCmd = IIf(chkReportDetailed1.Value = vbChecked, ucsZekCmdPrintReportByNumberDetailed, ucsZekCmdPrintReportByNumberShort)
            m_oFP.SendCommand eCmd, Pad(txtReportStart.Text, -4) & ";" & Pad(txtReportEnd.Text, -4) & IIf(chkReportPayments1.Value = vbChecked, ";P", vbNullString)
        ElseIf optReportType(2).Value Then
            eCmd = IIf(chkReportDetailed1.Value = vbChecked, ucsZekCmdPrintReportByDateDetailed, ucsZekCmdPrintReportByDateShort)
            m_oFP.SendCommand eCmd, Format$(C_Date(txtReportFD.Text), "ddmmyy") & ";" & Format$(C_Date(txtReportTD.Text), "ddmmyy") & IIf(chkReportPayments2.Value = vbChecked, ";P", vbNullString)
        ElseIf optReportType(3).Value Then
            m_oFP.SendCommand ucsZekCmdPrintReportByOperators, IIf(chkReportOperClosure.Value = vbChecked, "Z", "X") & ";0"
        ElseIf optReportType(4).Value Then
            m_oFP.SendCommand ucsZekCmdPrintReportSpecial
        ElseIf optReportType(5).Value Then ' EKL
            m_oFP.SendCommand ucsZekCmdPrintReportDaily, "E"
        End If
        If LenB(m_oFP.LastError) Then
            GoTo QH
        End If
        pvStatus = vbNullString
    End Select
    '--- success
    pvSaveData = True
    Exit Function
QH:
    If LenB(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
    End If
    Exit Function
EH:
    If pvShowError() Then
        Exit Function
    End If
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvShowError() As Boolean
    If Len(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
        pvStatus = m_oFP.LastError
        pvShowError = True
    End If
End Function

Private Sub pvLoadItem(vSplit As Variant)
    txtItemPLU.Text = Trim$(At(vSplit, 0))
    txtItemName.Text = Trim$(At(vSplit, 1))
    txtItemPrice.Text = Trim$(At(vSplit, 2))
    cobItemGroup.Text = Trim$(At(vSplit, 3))
    txtItemSum.Text = Trim$(At(vSplit, 4))
    txtItemAmount.Text = Trim$(At(vSplit, 5))
    txtItemReport.Text = Trim$(At(vSplit, 6))
    txtItemTime.Text = Trim$(At(vSplit, 7))
    If LenB(Trim$(At(vSplit, 8))) <> 0 Then
        If Asc(Trim$(At(vSplit, 8))) - &H80 <> 0 Then
            txtItemDep.Text = Asc(Trim$(At(vSplit, 8))) - &H80
        Else
            txtItemDep.Text = vbNullString
        End If
    Else
        txtItemDep.Text = vbNullString
    End If
End Sub

Private Sub pvApplyLogo()
    If m_hLogo <> 0 Then
        m_baLogoBW = ConvertToBW(m_hLogo, LimitLong(C_Lng(txtLogoWidth.Text), 1, 1000), LimitLong(C_Lng(txtLogoHeight.Text), 1, 1000), LimitLong(C_Lng(txtLogoTreshold.Text), 1, 99) * 256 \ 100, optLogoCenter.Value)
        Set picLogo.Picture = pvLoadBmp(m_baLogoBW)
        labLogoInfo.Caption = Printf(STR_LOGO_DIMENSIONS, UBound(m_baLogoBW) + 1)
        picLogo.Visible = True
        picLogo.Width = ScaleX(picLogo.Picture.Width, vbHimetric, vbTwips)
        picLogo.Height = ScaleY(picLogo.Picture.Height, vbHimetric, vbTwips)
        If picLogo.Width > picLogoScroll.Width Then
            scbLogoHor.Top = LimitLong(picLogo.Height, 0, picLogoScroll.Height - scbLogoHor.Height)
            scbLogoHor.Max = picLogo.Width - picLogoScroll.Width
            scbLogoHor.SmallChange = picLogoScroll.Width / 20
            scbLogoHor.LargeChange = picLogoScroll.Width / 4
            scbLogoHor.Visible = True
        Else
            scbLogoHor.Value = 0
            scbLogoHor.Visible = False
        End If
        scbLogoHor_Change
    Else
        labLogoInfo.Caption = vbNullString
        picLogo.Visible = False
        scbLogoHor.Value = 0
        scbLogoHor.Visible = False
    End If
End Sub

Private Function pvZfplibValue(sRegValue As String, ByVal lDefValue As Long) As Long
    Dim lValue          As Long
    
    lValue = C_Lng(RegReadString(HKEY_LOCAL_MACHINE, "Software\Tremol\ZFPLib", sRegValue))
    If lValue = 0 Then
        lValue = lDefValue
    End If
    pvZfplibValue = lValue
End Function

Private Sub pvZfplibSet(sRegValue As String, ByVal lValue As Long)
    RegWriteValue HKEY_LOCAL_MACHINE, "Software\Tremol\ZFPLib", sRegValue, lValue
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub chkReportDepartments_Click()
    chkReportItems.Value = vbUnchecked
End Sub

Private Sub chkReportItems_Click()
    chkReportDepartments.Value = vbUnchecked
End Sub

Private Sub cmdStatusReset_Click()
    Const FUNC_NAME     As String = "cmdStatusReset_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    If Not m_oFP.IsConnected Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_RESETTING
    m_oFP.SendCommand ucsZekCmdNonFiscalClose
    m_oFP.SendCommand ucsZekCmdFiscalCancel
    m_oFP.SendCommand ucsZekCmdFiscalPayAndClose
    pvStatus = vbNullString
    pvFetchData ucsCmdStatus
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cobCashPayment_Click()
    If cobCashPayment.ListIndex >= 0 Then
        txtCashTotal.Text = Format$(C_Dbl(At(m_vAdminCash(0), cobCashPayment.ListIndex + 1)), "0.00")
        txtCashIn.Text = Format$(C_Dbl(At(m_vAdminCash(1), cobCashPayment.ListIndex + 1)), "0.00")
        txtCashOut.Text = Format$(C_Dbl(At(m_vAdminCash(2), cobCashPayment.ListIndex + 1)), "0.00")
    Else
        txtCashTotal.Text = vbNullString
        txtCashIn.Text = vbNullString
        txtCashOut.Text = vbNullString
    End If
End Sub

Private Sub lstCmds_Click()
    Const FUNC_NAME     As String = "lstCmds_Click"
    Dim lIdx            As Long
    Dim lVisibleFrame   As Long
    Dim dblTimer        As Double
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    dblTimer = TimerEx
    If lstCmds.ListIndex = ucsCmdSettings Or lstCmds.ListIndex = ucsCmdOperations Or lstCmds.ListIndex = ucsCmdAdmin Then
        lVisibleFrame = -1
        GoTo QH
    End If
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            lVisibleFrame = -1
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_FETCHING
    If pvFetchData(lstCmds.ListIndex) Then
        If pvStatus = STR_STATUS_FETCHING Or LenB(pvStatus) = 0 Then
            pvStatus = Printf(STR_STATUS_SUCCESS_FETCH, Trim(lstCmds.List(lstCmds.ListIndex)), Format$(TimerEx - dblTimer, "0.000"))
        End If
        lVisibleFrame = lstCmds.ListIndex
    Else
        lVisibleFrame = -1
        If pvStatus = STR_STATUS_FETCHING Then
            pvStatus = vbNullString
        End If
    End If
QH:
    For lIdx = fraCommands.LBound To fraCommands.UBound
        If DispInvoke(fraCommands(lIdx), "Index", VbGet) Then
            fraCommands(lIdx).Visible = (lIdx = lVisibleFrame)
        End If
    Next
    tmrDate.Enabled = (lVisibleFrame = ucsCmdDateTime)
    Call SendMessage(txtLog.hWnd, EM_SCROLLCARET, 0, ByVal 0&)
    For lIdx = cmdSave.LBound To cmdSave.UBound
        If DispInvoke(cmdSave(lIdx), "Index", VbGet) Then
            If cmdSave(lIdx).Visible Then
                cmdSave(lIdx).Default = True
            End If
        End If
    Next
    On Error GoTo EH
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME & "(" & lstCmds.ListIndex & ")"
    Resume Next
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Const FUNC_NAME     As String = "cmdSave_Click"
    Dim dblTimer        As Double
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    dblTimer = TimerEx
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_SAVING
    If pvSaveData(lstCmds.ListIndex) Then
        If pvStatus = STR_STATUS_SAVING Then
            pvStatus = STR_STATUS_SAVING & " " & STR_STATUS_FETCHING
        End If
        If pvFetchData(lstCmds.ListIndex) Then
            If pvStatus = STR_STATUS_SAVING & " " & STR_STATUS_FETCHING Then
                pvStatus = Printf(STR_STATUS_SUCCESS_SAVE, Trim(lstCmds.List(lstCmds.ListIndex)), Format$(TimerEx - dblTimer, "0.000"))
            End If
        End If
    End If
QH:
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME & "(" & lstCmds.ListIndex & ")"
    Resume Next
End Sub

Private Sub lstDeps_Click()
    Const FUNC_NAME     As String = "lstDeps_Click"
    Dim vResult         As Variant
    
    On Error GoTo EH
    If lstDeps.ListIndex >= 0 Then
        txtDepNo.Text = lstDeps.ListIndex + 1
        vResult = m_vDeps(lstDeps.ListIndex + 1)
    Else
        txtDepNo.Text = vbNullString
    End If
    cobDepGroup.Text = At(vResult, 2)
    txtDepName.Text = Trim$(At(vResult, 1))
    txtDepSales.Text = Trim$(At(vResult, 4))
    txtDepTotalSum.Text = Trim$(At(vResult, 3))
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lstOpers_Click()
    Const FUNC_NAME     As String = "lstOpers_Click"
    Dim vResult         As Variant
    
    On Error GoTo EH
    If lstOpers.ListIndex >= 0 Then
        txtOperNo.Text = lstOpers.ListIndex + 1
        vResult = m_vOpers(lstOpers.ListIndex + 1)
    Else
        txtOperNo.Text = vbNullString
    End If
    txtOperName.Text = Trim$(At(vResult, 1))
    txtOperPass.Text = Trim$(At(vResult, 2))
    txtOperPass2.Text = vbNullString
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cmdItemLoad_Click()
    Const FUNC_NAME As String = "cmdItemLoad_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    If Not m_oFP.IsConnected Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvLoadItem Split(m_oFP.SendCommand(ucsZekCmdInfoItem, Right$(String$(5, "0") & txtItemNo.Text, 5)), ";")
    pvStatus = vbNullString
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cmdItemNew_Click()
    pvLoadItem Empty
End Sub

Private Sub cmdItemDelete_Click()
    '
End Sub

Private Sub lstStatus_Click()
    Const FUNC_NAME     As String = "lstStatus_Click"
    
    On Error GoTo EH
    If lstStatus.ListIndex >= 0 Then
        lstStatus.ToolTipText = lstStatus.List(lstStatus.ListIndex)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lstStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "lstStatus_KeyDown"
    Dim lIdx            As Long
    Dim sText           As String
    
    On Error GoTo EH
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        For lIdx = 0 To lstStatus.ListCount - 1
            sText = sText & IIf(lstStatus.Selected(lIdx), "[x] ", "[ ] ") & lstStatus.List(lIdx) & vbCrLf
        Next
        Clipboard.Clear
        Clipboard.SetText sText
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lstStatus_ItemCheck(Item As Integer)
    Const FUNC_NAME     As String = "lstStatus_ItemCheck"
    
    On Error GoTo EH
    If lstStatus.Tag = "Locked" Then
        lstStatus.Selected(Item) = Not lstStatus.Selected(Item)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub scbLogoHor_Change()
    picLogo.Left = -scbLogoHor.Value
End Sub

Private Sub scbLogoHor_Scroll()
    scbLogoHor_Change
End Sub

Private Sub tmrDate_Timer()
    txtDateCompDate.Text = Format$(Now, "dd\-MM\-yyyy")
    txtDateCompTime.Text = Format$(Now, "hh\:nn\:ss")
End Sub

Private Sub cmdDateTransfer_Click()
    txtDateDate.Text = txtDateCompDate.Text
    txtDateTime.Text = txtDateCompTime.Text
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtLogoHeight_Change()
    pvApplyLogo
End Sub

Private Sub txtLogoTreshold_Change()
    pvApplyLogo
End Sub

Private Sub txtLogoWidth_Change()
    pvApplyLogo
End Sub

Private Sub optLogoCenter_Click()
    txtLogoTreshold_Change
End Sub

Private Sub optLogoStretch_Click()
    txtLogoTreshold_Change
End Sub

Private Function pvLoadBmp(baData() As Byte) As StdPicture
    Dim nFile           As Integer
    Dim sFile           As String
    
    sFile = GetErrorTempPath() & "\~tmp" & Timer * 100 & ".bmp"
    On Error Resume Next '--- checked
    SetAttr sFile, vbArchive
    Kill sFile
    On Error GoTo 0
    nFile = FreeFile
    Open sFile For Binary As nFile
    Put nFile, , baData
    Close nFile
    Set pvLoadBmp = LoadPicture(sFile)
    On Error Resume Next '--- checked
    SetAttr sFile, vbArchive
    Kill sFile
    On Error GoTo 0
End Function
    
Private Sub cmdLogoOpen_Click()
    Const FUNC_NAME     As String = "cmdLogoOpen_Click"
    Const STR_TITLE     As String = "Logo"
    Const STR_FILTER    As String = "Graphical files (*.bmp;*.gif;*.png;*.jpg)|*.bmp;*.gif;*.png;*.jpg|All files (*.*)|*.*"
    Dim sFile           As String
    
    On Error GoTo EH
    If OpenSaveDialog(Me.hWnd, STR_FILTER, STR_TITLE, sFile) Then
        If m_hLogo <> 0 Then
            GdipReleaseImage m_hLogo
        End If
        Erase m_baLogoBW
        m_hLogo = GdipLoadImage(sFile)
        If m_hLogo = 0 Then
            MsgBox MSG_ERROR_LOADING_IMAGE, vbExclamation
        End If
        pvApplyLogo
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFPSink_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFPSink_CommandComplete"
    
    On Error GoTo EH
    m_sLog = m_sLog & lCmd & IIf(LenB(sData) <> 0, "<-" & sData, vbNullString) & IIf(LenB(sResult) <> 0, "->" & sResult, vbNullString) & vbCrLf
    If LenB(m_oFP.LastError) <> 0 Then
        m_sLog = m_sLog & m_oFP.LastError & vbCrLf
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

