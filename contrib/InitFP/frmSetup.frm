VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������"
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
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   8172
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   16
      Left            =   2250
      TabIndex        =   171
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdStatusReset 
         Caption         =   "�����"
         Height          =   375
         Left            =   180
         TabIndex        =   248
         Top             =   5220
         Width           =   1185
      End
      Begin VB.PictureBox picTab2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2625
         Left            =   90
         ScaleHeight     =   2628
         ScaleWidth      =   5592
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   2430
         Width           =   5595
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "������� �������"
            Height          =   285
            Index           =   6
            Left            =   90
            TabIndex        =   231
            Top             =   1710
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "��������� �������"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   230
            Top             =   1440
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "�������� �����"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   229
            Top             =   1170
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "Sw1.4"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   228
            Top             =   900
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "Sw1.3"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   227
            Top             =   630
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "������������� header"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   226
            Top             =   360
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusDip 
            Caption         =   "���������� header/footer"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   225
            Top             =   90
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "������ ����� ��"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   224
            Top             =   1980
            Width           =   1995
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "���� ��"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   223
            Top             =   2250
            Width           =   1995
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "����� �������� ��"
            Height          =   285
            Index           =   3
            Left            =   2880
            TabIndex        =   222
            Top             =   90
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "����� ��"
            Height          =   285
            Index           =   4
            Left            =   2880
            TabIndex        =   221
            Top             =   360
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "�������� ������"
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   220
            Top             =   630
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "�� ���� �� ������"
            Height          =   285
            Index           =   8
            Left            =   2880
            TabIndex        =   219
            Top             =   900
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "����������� ��"
            Height          =   285
            Index           =   9
            Left            =   2880
            TabIndex        =   218
            Top             =   1170
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "�������� �����"
            Height          =   285
            Index           =   11
            Left            =   2880
            TabIndex        =   217
            Top             =   1440
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "�������� ������� ������"
            Height          =   285
            Index           =   12
            Left            =   2880
            TabIndex        =   216
            Top             =   1710
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusMemory 
            Caption         =   "����������� ����� �� ��"
            Height          =   285
            Index           =   13
            Left            =   2880
            TabIndex        =   215
            Top             =   1980
            Width           =   2625
         End
      End
      Begin VB.PictureBox picTab1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2175
         Left            =   90
         ScaleHeight     =   2172
         ScaleWidth      =   5592
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   180
         Width           =   5595
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "����������� ������"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   213
            Top             =   90
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "��������� �������"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   212
            Top             =   360
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "�������� ���������"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   211
            Top             =   630
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "�� �� ��������"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   210
            Top             =   900
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "������������ ���������"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   209
            Top             =   1170
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "���� ������"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   208
            Top             =   1440
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "����������"
            Height          =   285
            Index           =   8
            Left            =   2880
            TabIndex        =   207
            Top             =   90
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "��������� �������� �����"
            Height          =   285
            Index           =   9
            Left            =   90
            TabIndex        =   206
            Top             =   1710
            Width           =   2715
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "���������� �����"
            Height          =   285
            Index           =   10
            Left            =   2880
            TabIndex        =   205
            Top             =   360
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "����������� �����"
            Height          =   285
            Index           =   12
            Left            =   2880
            TabIndex        =   204
            Top             =   630
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "���� ������"
            Height          =   285
            Index           =   16
            Left            =   2880
            TabIndex        =   203
            Top             =   900
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "������� �������� ���"
            Height          =   285
            Index           =   19
            Left            =   2880
            TabIndex        =   202
            Top             =   1170
            Width           =   2625
         End
         Begin VB.CheckBox chkStatusStatus 
            Caption         =   "������� �������� ���"
            Height          =   285
            Index           =   21
            Left            =   2880
            TabIndex        =   201
            Top             =   1440
            Width           =   2625
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����������"
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   85
         Top             =   5220
         Width           =   1275
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   10
      Left            =   2250
      TabIndex        =   140
      Top             =   90
      Width           =   5775
      Begin VB.Frame Frame1 
         Caption         =   "���������"
         Height          =   1185
         Left            =   180
         TabIndex        =   241
         Top             =   3780
         Width           =   5415
         Begin VB.OptionButton optLogoStretch 
            Caption         =   "���������"
            Height          =   285
            Left            =   1800
            TabIndex        =   247
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optLogoCenter 
            Caption         =   "����������"
            Height          =   285
            Left            =   180
            TabIndex        =   246
            Top             =   720
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.CommandButton cmdLogoOpen 
            Caption         =   "�����"
            Height          =   375
            Left            =   3960
            TabIndex        =   243
            Top             =   270
            Width           =   1275
         End
         Begin VB.TextBox txtLogoTreshold 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   242
            Text            =   "50"
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���� �� �����:"
            Height          =   390
            Left            =   180
            TabIndex        =   245
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
            TabIndex        =   244
            Top             =   270
            Width           =   555
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox chkLogoPrint 
         Caption         =   "����� �������� ���� ����� header"
         Height          =   285
         Left            =   180
         TabIndex        =   239
         Top             =   270
         Width           =   4245
      End
      Begin VB.PictureBox picLogoScroll 
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   180
         ScaleHeight     =   2448
         ScaleWidth      =   5412
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   720
         Width           =   5415
         Begin VB.HScrollBar scbLogoHor 
            CausesValidation=   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   238
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
            TabIndex        =   237
            TabStop         =   0   'False
            Top             =   0
            Width           =   5235
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   235
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label labLogoInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   195
         Left            =   180
         TabIndex        =   240
         Top             =   3330
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   5
      Left            =   2250
      TabIndex        =   157
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtInvCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   990
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   39
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtInvEnd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   38
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtInvStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   37
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   161
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
         TabIndex        =   159
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
         TabIndex        =   158
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   13
      Left            =   2250
      TabIndex        =   170
      Top             =   90
      Width           =   5775
      Begin VB.OptionButton optReportType 
         Caption         =   "����� ���������"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   80
         Top             =   3420
         Width           =   5145
      End
      Begin VB.CheckBox chkReportDetailed1 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   75
         Top             =   1710
         Width           =   1725
      End
      Begin VB.TextBox txtReportStart 
         Height          =   285
         Left            =   1800
         TabIndex        =   73
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txtReportEnd 
         Height          =   285
         Left            =   3420
         TabIndex        =   74
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CheckBox chkReportDetailed2 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   79
         Top             =   2970
         Width           =   1725
      End
      Begin VB.TextBox txtReportTD 
         Height          =   285
         Left            =   3420
         TabIndex        =   78
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtReportFD 
         Height          =   285
         Left            =   1800
         TabIndex        =   77
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   83
         Top             =   5220
         Width           =   1275
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "��������� ���� �� ������"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   81
         Top             =   3780
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ���� �� �����"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   76
         Top             =   2160
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ����� �� �����"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   72
         Top             =   990
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "������� ������"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   82
         Top             =   4140
         Width           =   5145
      End
      Begin VB.CheckBox chkReportDepartments 
         Caption         =   "������������"
         Height          =   285
         Left            =   3960
         TabIndex        =   71
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportItems 
         Caption         =   "��������"
         Height          =   285
         Left            =   2430
         TabIndex        =   70
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportClosure 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   69
         Top             =   630
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "������ �������� �����"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   68
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
         TabIndex        =   187
         Top             =   1350
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
         TabIndex        =   186
         Top             =   1350
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
         TabIndex        =   185
         Top             =   2520
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� ����:"
         Height          =   300
         Left            =   900
         TabIndex        =   184
         Top             =   2520
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   17
      Left            =   2250
      TabIndex        =   188
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtLog 
         Height          =   5505
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   0
      Left            =   2250
      TabIndex        =   87
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
         Top             =   1530
         Width           =   1785
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   15
      Left            =   2250
      TabIndex        =   172
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtDiagFirmware 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtDiagChecksum 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtDiagSwitches 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   84
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firmware:"
         Height          =   195
         Left            =   180
         TabIndex        =   178
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   177
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� Sw1..Sw4:"
         Height          =   195
         Left            =   180
         TabIndex        =   176
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   12
      Left            =   2250
      TabIndex        =   162
      Top             =   90
      Width           =   5775
      Begin VB.OptionButton optCashOut 
         Caption         =   "�����"
         Height          =   285
         Left            =   3150
         TabIndex        =   65
         Top             =   1620
         Width           =   1455
      End
      Begin VB.OptionButton optCashIn 
         Caption         =   "����"
         Height          =   285
         Left            =   2070
         TabIndex        =   64
         Top             =   1620
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtCashSum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   66
         Top             =   1980
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����/�����"
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   67
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtCashOut 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   990
         Width           =   1545
      End
      Begin VB.TextBox txtCashIn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtCashTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   180
         TabIndex        =   169
         Top             =   1980
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   168
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   166
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� ����:"
         Height          =   195
         Left            =   180
         TabIndex        =   164
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   9
      Left            =   2250
      TabIndex        =   189
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtItemTime 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2985
      End
      Begin VB.CommandButton cmdItemDelete 
         Caption         =   "���������"
         Height          =   375
         Left            =   4320
         TabIndex        =   62
         Top             =   3960
         Width           =   1275
      End
      Begin VB.CommandButton cmdItemNew 
         Caption         =   "���"
         Height          =   375
         Left            =   2970
         TabIndex        =   61
         Top             =   3960
         Width           =   1275
      End
      Begin VB.TextBox txtItemSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   3510
         Width           =   1545
      End
      Begin VB.TextBox txtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   9
         TabIndex        =   58
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtItemPLU 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   57
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtItemAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1545
      End
      Begin VB.ComboBox cobItemGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   59
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   2610
         MaxLength       =   25
         TabIndex        =   60
         Top             =   1980
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   13
         Left            =   4320
         TabIndex        =   63
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstItems 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   56
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������� ��:"
         Height          =   195
         Left            =   2610
         TabIndex        =   233
         Top             =   2520
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   199
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   197
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLU:"
         Height          =   195
         Left            =   2610
         TabIndex        =   196
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� �����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   195
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   194
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   192
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   191
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   8
      Left            =   2250
      TabIndex        =   128
      Top             =   90
      Width           =   5775
      Begin VB.ComboBox cobDepGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   52
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
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txtDepRecSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   136
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
         TabIndex        =   134
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
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtDepName2 
         Height          =   285
         Left            =   2610
         MaxLength       =   36
         TabIndex        =   54
         Top             =   1890
         Width           =   2985
      End
      Begin VB.ListBox lstDeps 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   51
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   55
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtDepName 
         Height          =   285
         Left            =   2610
         MaxLength       =   31
         TabIndex        =   53
         Top             =   1260
         Width           =   2985
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� �� ����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   139
         Top             =   3060
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   137
         Top             =   2700
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� ��������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   135
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
         TabIndex        =   133
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
         TabIndex        =   132
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� �����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   130
         Top             =   1620
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   129
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   7
      Left            =   2250
      TabIndex        =   141
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdOperReset 
         Caption         =   "��������"
         Height          =   375
         Left            =   4320
         TabIndex        =   47
         Top             =   3150
         Width           =   1275
      End
      Begin VB.TextBox txtOperPass2 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   49
         Top             =   4410
         Width           =   1545
      End
      Begin VB.TextBox txtOperPass 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   48
         Top             =   4050
         Width           =   1545
      End
      Begin VB.TextBox txtOperSurcharge 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   2430
         Width           =   1545
      End
      Begin VB.TextBox txtOperVoid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1545
      End
      Begin VB.TextBox txtOperFiscal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtOperSells 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1545
      End
      Begin VB.TextBox txtOperDisc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1545
      End
      Begin VB.TextBox txtOperName 
         Height          =   285
         Left            =   2610
         MaxLength       =   24
         TabIndex        =   46
         Top             =   900
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   50
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstOpers 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   45
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
         TabIndex        =   142
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
         TabIndex        =   156
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
         TabIndex        =   155
         Top             =   4050
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   154
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   153
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   150
         Top             =   1350
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   149
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   148
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������:"
         Height          =   195
         Left            =   2610
         TabIndex        =   144
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
         TabIndex        =   143
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   6
      Left            =   2250
      TabIndex        =   179
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   3
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   43
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   2
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   42
         Top             =   990
         Width           =   3525
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   1
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   41
         Top             =   630
         Width           =   3525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   44
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   0
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   40
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 4:"
         Height          =   195
         Left            =   180
         TabIndex        =   183
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 3:"
         Height          =   195
         Left            =   180
         TabIndex        =   182
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   181
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   180
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   4
      Left            =   2250
      TabIndex        =   116
      Top             =   90
      Width           =   5775
      Begin VB.CheckBox chkHeadFormatInvoice 
         Caption         =   "����������� ���� ������� �� ������ �������� ������"
         Height          =   285
         Left            =   180
         TabIndex        =   33
         Top             =   4680
         Width           =   5415
      End
      Begin VB.TextBox txtHeadBulstatName 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   19
         Top             =   990
         Width           =   1635
      End
      Begin VB.TextBox txtHeadBulstatText 
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   20
         Top             =   990
         Width           =   1635
      End
      Begin VB.TextBox txtHeadHeader6 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   24
         Top             =   2430
         Width           =   3525
      End
      Begin VB.TextBox txtHeadRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Top             =   4320
         Width           =   1275
      End
      Begin VB.CheckBox chkHeadAdvanceHeader 
         Caption         =   "������������� header (����� ������)"
         Height          =   285
         Left            =   180
         TabIndex        =   35
         Top             =   5400
         Width           =   4065
      End
      Begin VB.CheckBox chkHeadVat 
         Caption         =   "����� ��� � ���������� �������� ���"
         Height          =   285
         Left            =   180
         TabIndex        =   34
         Top             =   5040
         Width           =   4065
      End
      Begin VB.CheckBox chkHeadSumDivider 
         Caption         =   "������������ ����� ����� ���� ����"
         Height          =   285
         Left            =   180
         TabIndex        =   29
         Top             =   3960
         Width           =   5415
      End
      Begin VB.CheckBox chkHeadEmptyFooter 
         Caption         =   "������ ����� ���� footer"
         Height          =   285
         Left            =   2970
         TabIndex        =   28
         Top             =   3600
         Width           =   2715
      End
      Begin VB.CheckBox chkHeadEmptyHeader 
         Caption         =   "������ ����� ���� header"
         Height          =   285
         Left            =   180
         TabIndex        =   27
         Top             =   3600
         Width           =   2715
      End
      Begin VB.TextBox txtHeadFooter2 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   26
         Top             =   3150
         Width           =   3525
      End
      Begin VB.TextBox txtHeadFooter1 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   25
         Top             =   2790
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader5 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   23
         Top             =   2070
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader4 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   22
         Top             =   1710
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader3 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   21
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader1 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   17
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader2 
         Height          =   285
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   18
         Top             =   630
         Width           =   3525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   36
         Top             =   5220
         Width           =   1275
      End
      Begin VB.CheckBox chkHeadRateEUR 
         Caption         =   "����: "
         Height          =   285
         Left            =   2970
         TabIndex        =   31
         Top             =   4320
         Width           =   1185
      End
      Begin VB.CheckBox chkHeadSumEUR 
         Caption         =   "����� ���� ���� � EUR"
         Height          =   285
         Left            =   180
         TabIndex        =   30
         Top             =   4320
         Width           =   3435
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Left            =   3690
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   123
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   3
      Left            =   2250
      TabIndex        =   109
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdDateTransfer 
         Caption         =   "�� ��������"
         Height          =   375
         Left            =   2070
         TabIndex        =   15
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
         TabIndex        =   115
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
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox txtDateDate 
         Height          =   285
         Left            =   2070
         TabIndex        =   13
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtDateTime 
         Height          =   285
         Left            =   2070
         TabIndex        =   14
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   16
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������� ���:"
         Height          =   195
         Left            =   180
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
         Top             =   1530
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   2
      Left            =   2250
      TabIndex        =   88
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   12
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtTaxGroup4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   11
         Top             =   3510
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   10
         Top             =   3150
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   2790
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   8
         Top             =   2430
         Width           =   1095
      End
      Begin VB.TextBox txtTaxRates 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   7
         Top             =   2070
         Width           =   1095
      End
      Begin VB.TextBox txtTaxCurrency 
         Height          =   285
         Left            =   2070
         TabIndex        =   6
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox txtTaxDecimals 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   5
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txtTaxCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.TextBox txtTaxSerNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   95
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
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   108
         Top             =   3510
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   107
         Top             =   3150
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   106
         Top             =   2790
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   105
         Top             =   2430
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   104
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   103
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   102
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �:"
         Height          =   195
         Left            =   180
         TabIndex        =   101
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� ������:"
         Height          =   195
         Left            =   180
         TabIndex        =   100
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� �������:"
         Height          =   195
         Left            =   180
         TabIndex        =   99
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   98
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   195
         Left            =   180
         TabIndex        =   96
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   94
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
         TabIndex        =   92
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
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   180
      Width           =   1275
   End
   Begin VB.ListBox lstCmds 
      Height          =   5685
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2085
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   90
      TabIndex        =   127
      Top             =   5940
      Width           =   7935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "frmSetup"

'=========================================================================
' API
'=========================================================================

Private Const EM_SCROLLCARET                    As Long = &HB7
Private Const OFN_HIDEREADONLY          As Long = &H4&
Private Const OFN_EXTENSIONDIFFERENT    As Long = &H400
Private Const OFN_CREATEPROMPT          As Long = &H2000&
Private Const OFN_EXPLORER              As Long = &H80000
Private Const OFN_LONGNAMES             As Long = &H200000
Private Const OFN_ENABLESIZING          As Long = &H800000

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OPENFILENAME
    lStructSize         As Long     ' size of type/structure
    hWndOwner           As Long     ' Handle of owner window
    hInstance           As Long
    lpstrFilter         As Long     ' Filters used to select files
    lpstrCustomFilter   As Long
    nMaxCustomFilter    As Long
    nFilterIndex        As Long     ' index of Filter to start with
    lpstrFile           As Long     ' Holds filepath and name
    nMaxFile            As Long     ' Maximum Filepath and name length
    lpstrFileTitle      As Long     ' Filename
    nMaxFileTitle       As Long     ' Max Length of filename
    lpstrInitialDir     As Long     ' Starting Directory
    lpstrTitle          As Long     ' Title of window
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
    pvReserved          As Long
    dwReserved          As Long
    FlagsEx             As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_NUM_DEPS          As Long = 50
Private Const LNG_NUM_OPERS         As Long = 24
Private Const FORMAT_CURRENCY       As String = "#0.00######"
Private Const LNG_LOGO_FORECOLOR    As Long = &H800000
'--- strings
Private Const STR_SPEEDS            As String = "115200|38400|9600|19200|57600"
Private Const STR_COMMANDS          As String = "������ �������|���������|    ������� ����������|    ���� � ���|    Header � footer|    ������ �� �������|    ������������ ������|    ���������|    ������������|    ��������|    �������� ����|��������|    ���� � �����|    ����� ������|�������������|    �����������|    ������|    ������ �����������"
Private Const STR_GROUPS            As String = "�|�|�|�"
Private Const STR_NA                As String = "N/A"
Private Const STR_STATUS_CONNECTING As String = "���������..."
Private Const STR_STATUS_SUCCESS_CONNECT As String = "������� %1"
Private Const STR_STATUS_FAILURE_CONNECT As String = "���� ������"
Private Const STR_STATUS_SAVING     As String = "���������..."
Private Const STR_STATUS_SUCCESS_SAVE As String = "������� ��������� �� %1 (%2 ���.)"
Private Const STR_STATUS_FETCHING   As String = "����������..."
Private Const STR_STATUS_SUCCESS_FETCH As String = "������� ���������� �� %1 (%2 ���.)"
Private Const STR_STATUS_NOT_IMPLEMENTED As String = "�� � �����������"
Private Const STR_STATUS_NO_DEP_SELECTED As String = "������ ������ �����������"
Private Const STR_STATUS_FETCH_DEP  As String = "���������� ����������� %1 �� " & LNG_NUM_DEPS & "..."
Private Const STR_STATUS_ENUM_PORTS As String = "���������� �� ������� ��������..."
Private Const STR_STATUS_FETCH_OPER As String = "���������� ��������� %1 �� " & LNG_NUM_OPERS & "..."
Private Const STR_STATUS_NO_OPER_SELECTED As String = "������ ������ ��������"
Private Const STR_STATUS_OPER_RESETTING As String = "��������..."
Private Const STR_STATUS_OPER_SUCCESS_RESET As String = "������� �������� �� �������� %1"
Private Const STR_STATUS_REFRESH    As String = "�����������..."
Private Const STR_STATUS_PRINT      As String = "�����������..."
Private Const STR_STATUS_NO_ITEM_GROUP As String = "���� �������� ����� �� �������"
Private Const STR_STATUS_NO_ITEM_PLU As String = "���� ������� PLU �� �������"
Private Const STR_STATUS_NO_ITEM_PRICE As String = "���� �������� ���� �� �������"
Private Const STR_STATUS_NO_ITEM_NAME As String = "���� �������� ������������ �� �������"
Private Const STR_STATUS_ITEM_FAILURE_ADD As String = "���������� ��������. %1 �������� ������� ��������"
Private Const STR_STATUS_ITEM_DELETING As String = "���������..."
Private Const STR_STATUS_ITEM_SUCCESS_DELETE As String = "������� ��������� �� ������� PLU %1"
Private Const STR_STATUS_FETCH_LOGO As String = "���������� �� ��� %1..."
Private Const STR_STATUS_SAVE_LOGO  As String = "����� �� ��� %1/%2..."
Private Const STR_OPER_PASS_PROMPT  As String = "������ �� �������� %1"
Private Const STR_OPER_PASS_CAPTION As String = "������ �� ������"
Private Const STR_LOGO_DIMENSIONS   As String = "������ �� �������: %1x%2"
Private Const STR_STATUS_RESETTING  As String = "�����..."
'--- messages
Private Const MSG_INVALID_PASSWORD  As String = "���������� ������" & vbCrLf & vbCrLf & "�������� �� ������� �� 4 �� 6 �����"
Private Const MSG_PASSWORDS_MISMATCH As String = "�������� �� ��������"
Private Const MSG_REJECTED_PASSWORD As String = "��������� ������ �� ��������"
Private Const MSG_REQUEST_CANCELLED As String = "�������� � ��������"
Private Const MSG_CONFIRM_ITEM_DELETE As String = "������� �� �� �������� ������� PLU %1?"

Private WithEvents m_oFP        As cIslProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private m_sLog                  As String
Private m_vDeps                 As Variant
Private m_vOpers                As Variant
Private m_vItems                As Variant
Private m_vLogo                 As Variant
Private m_picLogo               As StdPicture

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
    ucsCmdOperations
        ucsCmdCashOper
        ucsCmdReports
    ucsCmdAdmin
        ucsCmdDiagnostics
        ucsCmdStatus
        ucsCmdLog
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    MsgBox MODULE_NAME & "." & sFunc & ": " & Error, vbCritical
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
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

Private Property Get pvLogoPixel(ByVal lX As Long, ByVal lY As Long) As Boolean
    On Error GoTo EH
    pvLogoPixel = ((CLng("&H" & Mid(m_vLogo(lY), 1 + 2 * (lX \ 8), 2)) And (2 ^ (7 - lX Mod 8))) <> 0)
    Exit Property
EH:
    Debug.Print lY, lX, Mid(m_vLogo(lY), 1 + 2 * (lX \ 8), 2)
    Resume Next
End Property

Private Property Let pvLogoPixel(ByVal lX As Long, ByVal lY As Long, ByVal bValue As Boolean)
    Dim lValue          As Long
    
    lValue = C_Lng("&H" & Mid(m_vLogo(lY), 1 + 2 * (lX \ 8), 2))
    If bValue Then
        lValue = lValue Or (2 ^ (7 - lX Mod 8))
    Else
        lValue = lValue And (Not 2 ^ (7 - lX Mod 8))
    End If
    Mid(m_vLogo(lY), 1 + 2 * (lX \ 8), 2) = Right("0" & Hex(lValue), 2)
End Property

Private Property Get OsVersion() As Long
    Static lVersion     As Long
    Dim uVer            As OSVERSIONINFO
    
    If lVersion = 0 Then
        uVer.dwOSVersionInfoSize = Len(uVer)
        If GetVersionEx(uVer) Then
            lVersion = uVer.dwMajorVersion * 100 + uVer.dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property

'=========================================================================
' Methods
'=========================================================================

Private Property Get pvLock(oCtl As Object) As Boolean
    On Error Resume Next
    pvLock = oCtl.Locked
    If Err.Number Then
        pvLock = Not oCtl.Enabled
    End If
End Property

Private Property Let pvLock(oCtl As Object, ByVal bValue As Boolean)
    On Error Resume Next
    oCtl.Locked = bValue
    If Err.Number Then
        oCtl.Enabled = Not bValue
    Else
        oCtl.BackColor = IIf(bValue, vbButtonFace, vbWindowBackground)
    End If
End Property

Private Function pvFetchData(ByVal eCmd As UcsCommands) As Boolean
    Const FUNC_NAME     As String = "pvFetchData"
    Dim lIdx            As Long
    Dim vResult         As Variant
    Dim sText           As String
    Dim lRow            As Long
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCmd <> ucsCmdConnect And lstCmds.ListIndex <> ucsCmdStatus Then
        pvStatus = labConnectCurrent.Caption
        Exit Function
    End If
    Select Case eCmd
    Case ucsCmdConnect
        pvStatus = labConnectCurrent.Caption
    Case ucsCmdTaxInfo
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoDiagnostics, "0"), ",")
        If UBound(vResult) < 0 Then
            vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoDiagnostics, vbNullString), ",")
        End If
        txtTaxMemModule.Text = At(vResult, 5)
        txtTaxSerNo.Text = At(vResult, 4)
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInitDecimals, vbNullString), ",")
        txtTaxDecimals.Text = C_Lng(At(vResult, 1))
        txtTaxCurrency.Text = Trim(At(vResult, 2))
        txtTaxRates.Text = C_Lng(At(vResult, 3))
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoTaxRates, vbNullString), ",")
        txtTaxGroup1.Text = C_Lng(At(vResult, 0))
        txtTaxGroup2.Text = C_Lng(At(vResult, 1))
        txtTaxGroup3.Text = C_Lng(At(vResult, 2))
        txtTaxGroup4.Text = C_Lng(At(vResult, 3))
    Case ucsCmdDateTime
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoDateTime, vbNullString), " ")
        txtDateDate.Text = At(vResult, 0)
        txtDateTime.Text = At(vResult, 1)
        tmrDate_Timer
    Case ucsCmdHeaderFooter
        txtHeadHeader1.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I0")
        pvLock(txtHeadHeader1) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadHeader2.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I1")
        pvLock(txtHeadHeader2) = m_oFP.Status(ucsIslStbPrintingError)
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoBulstat, vbNullString), ",")
        txtHeadBulstatName.Text = At(vResult, 1)
        txtHeadBulstatText.Text = At(vResult, 0)
        pvLock(txtHeadBulstatName) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(txtHeadBulstatText) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadHeader3.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I2")
        pvLock(txtHeadHeader3) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadHeader4.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I3")
        pvLock(txtHeadHeader4) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadHeader5.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I4")
        pvLock(txtHeadHeader5) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadHeader6.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I5")
        pvLock(txtHeadHeader6) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadFooter1.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I6")
        pvLock(txtHeadFooter1) = m_oFP.Status(ucsIslStbPrintingError)
        txtHeadFooter2.Text = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "I7")
        pvLock(txtHeadFooter2) = m_oFP.Status(ucsIslStbPrintingError)
        chkHeadFormatInvoice.Value = -(m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IA") = "1")
        pvLock(chkHeadFormatInvoice) = m_oFP.Status(ucsIslStbPrintingError)
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IE"), ",")
        chkHeadSumEUR.Value = -(At(vResult, 0) = "1")
        chkHeadRateEUR.Value = chkHeadSumEUR.Value
        txtHeadRate.Text = Trim(At(vResult, 1))
        pvLock(chkHeadSumEUR) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(chkHeadRateEUR) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(txtHeadRate) = m_oFP.Status(ucsIslStbPrintingError)
        chkHeadAdvanceHeader.Value = -(m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IH") = "1")
        pvLock(chkHeadAdvanceHeader) = m_oFP.Status(ucsIslStbPrintingError)
        vResult = m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IP")
        chkHeadEmptyHeader.Value = -(Mid(vResult, 1, 1) = "1")
        chkHeadEmptyFooter.Value = -(Mid(vResult, 3, 1) = "1")
        chkHeadSumDivider.Value = -(Mid(vResult, 4, 1) = "1")
        pvLock(chkHeadEmptyHeader) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(chkHeadEmptyFooter) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(chkHeadSumDivider) = m_oFP.Status(ucsIslStbPrintingError)
        chkHeadVat.Value = -(m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IT") = "1")
        pvLock(chkHeadVat) = m_oFP.Status(ucsIslStbPrintingError)
    Case ucsCmdInvoiceNo
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInitInvoiceNo, vbNullString), ",")
        txtInvStart.Text = C_Dbl(At(vResult, 0))
        txtInvEnd.Text = C_Dbl(At(vResult, 1))
        txtInvCurrent.Text = C_Dbl(At(vResult, 2))
        pvLock(txtInvStart) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(txtInvEnd) = m_oFP.Status(ucsIslStbPrintingError)
        pvLock(txtInvCurrent) = m_oFP.Status(ucsIslStbPrintingError)
    Case ucsCmdPaymentTypes
        txtPmtType(0).Text = m_oFP.SendCommand(ucsIslCmdInitPaymentType, "I")
        If m_oFP.Status(ucsIslStbPrintingError) Then
            txtPmtType(0).Text = m_oFP.SendCommand(ucsIslCmdExtendedInitText, "R61")
        End If
        pvLock(txtPmtType(0)) = m_oFP.Status(ucsIslStbPrintingError)
        txtPmtType(1).Text = m_oFP.SendCommand(ucsIslCmdInitPaymentType, "J")
        If m_oFP.Status(ucsIslStbPrintingError) Then
            txtPmtType(1).Text = m_oFP.SendCommand(ucsIslCmdExtendedInitText, "R62")
        End If
        pvLock(txtPmtType(1)) = m_oFP.Status(ucsIslStbPrintingError)
        txtPmtType(2).Text = m_oFP.SendCommand(ucsIslCmdInitPaymentType, "K")
        If m_oFP.Status(ucsIslStbPrintingError) Then
            txtPmtType(2).Text = m_oFP.SendCommand(ucsIslCmdExtendedInitText, "R63")
        End If
        pvLock(txtPmtType(2)) = m_oFP.Status(ucsIslStbPrintingError)
        txtPmtType(3).Text = m_oFP.SendCommand(ucsIslCmdInitPaymentType, "L")
        If m_oFP.Status(ucsIslStbPrintingError) Then
            txtPmtType(3).Text = m_oFP.SendCommand(ucsIslCmdExtendedInitText, "R64")
        End If
        pvLock(txtPmtType(3)) = m_oFP.Status(ucsIslStbPrintingError)
    Case ucsCmdOperators
        If Not IsArray(m_vOpers) Then
            ReDim m_vOpers(0 To LNG_NUM_OPERS) As Variant
        End If
        For lIdx = 1 To UBound(m_vOpers)
            If Not IsArray(m_vOpers(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_OPER, lIdx)
                m_vOpers(lIdx) = Split(m_oFP.SendCommand(ucsIslCmdInfoOperator, C_Str(lIdx)), ",")
                If m_oFP.Status(ucsIslStbPrintingError) Then
                    ReDim Preserve m_vOpers(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            If lstOpers.ListCount < lIdx Then
                lstOpers.AddItem vbNullString
            End If
            sText = lIdx & ": " & At(m_vOpers(lIdx), 5)
            If lstOpers.List(lIdx - 1) <> sText Then
                lstOpers.List(lIdx - 1) = sText
            End If
        Next
        lstOpers_Click
        pvStatus = vbNullString
    Case ucsCmdDepartments
        If Not IsArray(m_vDeps) Then
            ReDim m_vDeps(0 To LNG_NUM_DEPS)
        End If
        For lIdx = 1 To UBound(m_vDeps)
            If Not IsArray(m_vDeps(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_DEP, lIdx)
                m_vDeps(lIdx) = Split(m_oFP.SendCommand(ucsIslCmdInfoDepartment, C_Str(lIdx)), ",")
                If m_oFP.Status(ucsIslStbPrintingError) Then
                    ReDim Preserve m_vDeps(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            If lstDeps.ListCount < lIdx Then
                lstDeps.AddItem vbNullString
            End If
            vResult = m_vDeps(lIdx)
            If Left(At(vResult, 0), 1) = "P" Then
                sText = At(Split(At(vResult, 4), vbLf), 0) & " (" & Mid(At(vResult, 0), 2) & ")"
            Else
                sText = lIdx & ": " & STR_NA
            End If
            If lstDeps.List(lIdx - 1) <> sText Then
                lstDeps.List(lIdx - 1) = sText
            End If
        Next
        lstDeps_Click
        pvStatus = vbNullString
    Case ucsCmdItems
        If Not IsArray(m_vItems) Then
            lstItems.Clear
            ReDim m_vItems(0 To 0)
            lIdx = 0
            vResult = Split(m_oFP.SendCommand(ucsIslCmdInitItem, "F"), ",")
            Do While vResult(0) <> "F"
                lIdx = lIdx + 1
                ReDim Preserve m_vItems(0 To lIdx) As Variant
                m_vItems(lIdx) = vResult
                vResult = Split(m_oFP.SendCommand(ucsIslCmdInitItem, "N"), ",")
            Loop
        End If
        For lIdx = 1 To UBound(m_vItems)
            If Not IsArray(m_vItems(lIdx)) Then
                m_vItems(lIdx) = Split(m_oFP.SendCommand(ucsIslCmdInitItem, "R" & m_vItems(lIdx)), ",")
            End If
            vResult = m_vItems(lIdx)
            If lstItems.ListCount < lIdx Then
                lstItems.AddItem vbNullString
            End If
            sText = At(vResult, 1) & ": " & At(vResult, 7)
            If lstItems.List(lIdx - 1) <> sText Then
                lstItems.List(lIdx - 1) = sText
            End If
        Next
        lstItems_Click
    Case ucsCmdGraphicalLogo
        chkLogoPrint.Value = IIf(m_oFP.SendCommand(ucsIslCmdInitHeaderFooter, "IL") = "1", vbChecked, vbUnchecked)
        pvLock(chkLogoPrint) = m_oFP.Status(ucsIslStbPrintingError)
        If Not IsArray(m_vLogo) Then
            ReDim m_vLogo(0 To 1000)
            For lRow = 0 To UBound(m_vLogo)
                pvStatus = Printf(STR_STATUS_FETCH_LOGO, lRow + 1)
                m_vLogo(lRow) = m_oFP.SendCommand(ucsIslCmdInitLogo, "R" & lRow)
                If m_oFP.Status(ucsIslStbPrintingError) Then
                    If lRow > 0 Then
                        ReDim Preserve m_vLogo(0 To lRow - 1) As Variant
                    Else
                        '--- daisy FP
                        vResult = Split(m_oFP.SendCommand(ucsIslCmdExtendedInitSetting, vbNullString), ",")
                        ReDim Preserve m_vLogo(0 To C_Lng(At(vResult, 1, 64)) - 1) As Variant
                        For lIdx = 0 To UBound(m_vLogo)
                            m_vLogo(lIdx) = String(C_Lng(At(vResult, 0, 64)) / 4, "0")
                        Next
                    End If
                    Exit For
                End If
                '--- note: bug in firmware byte to hex routine: 0xA - 1 = "@" instead of "9"
                m_vLogo(lRow) = Replace(m_vLogo(lRow), "@", "9")
            Next
            picLogo.Width = Len(m_vLogo(0)) * 4 * Screen.TwipsPerPixelX
            picLogo.Height = (1 + UBound(m_vLogo)) * Screen.TwipsPerPixelY
            If picLogo.Width > picLogoScroll.Width Then
                scbLogoHor.Top = picLogo.Height
                scbLogoHor.Visible = True
                scbLogoHor.Max = picLogo.Width - picLogoScroll.Width
                scbLogoHor.SmallChange = picLogoScroll.Width / 20
                scbLogoHor.LargeChange = picLogoScroll.Width / 4
            Else
                scbLogoHor.Visible = False
            End If
            picLogoScroll.Height = picLogo.Height + scbLogoHor.Height
            labLogoInfo.Top = picLogoScroll.Top + picLogoScroll.Height + 60
            labLogoInfo.Caption = Printf(STR_LOGO_DIMENSIONS, picLogo.Width \ Screen.TwipsPerPixelX, picLogo.Height \ Screen.TwipsPerPixelY)
            pvStatus = vbNullString
        End If
        For lRow = 0 To UBound(m_vLogo)
            For lIdx = 0 To Len(m_vLogo(lRow)) * 4 - 1
                Call SetPixel(picLogo.hDC, lIdx, lRow, IIf(pvLogoPixel(lIdx, lRow), LNG_LOGO_FORECOLOR, vbWhite))
            Next
        Next
    Case ucsCmdCashOper
        vResult = Split(m_oFP.SendCommand(ucsIslCmdFiscalServiceDeposit, vbNullString), ",")
        txtCashTotal.Text = Format(C_Dbl(At(vResult, 1)) / 100, FORMAT_CURRENCY)
        txtCashIn.Text = Format(C_Dbl(At(vResult, 2)) / 100, FORMAT_CURRENCY)
        txtCashOut.Text = Format(C_Dbl(At(vResult, 3)) / 100, FORMAT_CURRENCY)
        pvLock(txtCashSum) = m_oFP.Status(ucsIslStbPrintingError)
    Case ucsCmdReports
        '--- do nothing
    Case ucsCmdStatus
        On Error Resume Next
        For lIdx = chkStatusStatus.LBound To chkStatusStatus.ubound
            chkStatusStatus(lIdx).Value = -m_oFP.Status(2 ^ lIdx)
        Next
        For lIdx = chkStatusDip.LBound To chkStatusDip.ubound
            chkStatusDip(lIdx).Value = -m_oFP.Dip(2 ^ lIdx)
        Next
        For lIdx = chkStatusMemory.LBound To chkStatusMemory.ubound
            chkStatusMemory(lIdx).Value = -m_oFP.Memory(2 ^ lIdx)
        Next
        On Error GoTo EH
    Case ucsCmdDiagnostics
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoDiagnostics, "1"), ",")
        txtDiagFirmware.Text = At(vResult, 0)
        txtDiagChecksum.Text = At(vResult, 1)
        txtDiagSwitches.Text = At(vResult, 2)
    Case ucsCmdLog
        m_sLog = Right(m_sLog, 32000)
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
    Dim sPass           As String
    Dim bCheckPass      As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCommand <> ucsCmdConnect Then
        Exit Function
    End If
    Select Case eCommand
    Case ucsCmdConnect
        '--- value might be not be found
        On Error Resume Next
        DeleteSetting App.Title, "Connect", "Port"
        On Error GoTo EH
        pvStatus = STR_STATUS_CONNECTING
        If m_oFP.Init("Port=" & cobConnectPort.Text & ";Speed=" & C_Lng(cobConnectSpeed.Text) & ";Timeout=3000") Then
            On Error Resume Next
            m_oFP.SendCommand ucsIslCmdInfoTransaction, vbNullString
            If pvShowError() Then
                On Error GoTo EH
                labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
                Caption = App.Title
            Else
                On Error GoTo EH
                labConnectCurrent.Caption = Printf(STR_STATUS_SUCCESS_CONNECT, m_oFP.Device)
                Caption = m_oFP.Device & " - " & App.Title
                '--- save conn info
                If chkConnectRemember.Value Then
                    SaveSetting App.Title, "Connect", "Port", cobConnectPort.Text
                    SaveSetting App.Title, "Connect", "Speed", cobConnectSpeed.Text
                End If
                '--- flush cache
                m_vDeps = Empty
                m_vOpers = Empty
                m_vItems = Empty
                m_vLogo = Empty
                m_sLog = vbNullString
                lstCmds.ListIndex = ucsCmdTaxInfo
            End If
        Else
            labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
            Caption = App.Title
        End If
    Case ucsCmdTaxInfo
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInitDecimals, vbNullString), ",")
        m_oFP.SendCommand ucsIslCmdInitDecimals, At(vResult, 0) & "," & C_Lng(txtTaxDecimals.Text) & "," & txtTaxCurrency.Text & " ," & C_Lng(txtTaxRates.Text)
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInfoTaxRates, vbNullString), ",")
        vResult(0) = C_Lng(txtTaxGroup1.Text)
        vResult(1) = C_Lng(txtTaxGroup2.Text)
        vResult(2) = C_Lng(txtTaxGroup3.Text)
        vResult(3) = C_Lng(txtTaxGroup4.Text)
        m_oFP.SendCommand ucsIslCmdInitTaxRates, Join(vResult, ",")
    Case ucsCmdDateTime
        m_oFP.SendCommand ucsIslCmdInitDateTime, txtDateDate.Text & " " & txtDateTime.Text
    Case ucsCmdHeaderFooter
        If Not pvLock(txtHeadHeader1) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "0" & RTrim(txtHeadHeader1.Text)
        End If
        If Not pvLock(txtHeadHeader2) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "1" & RTrim(txtHeadHeader2.Text)
        End If
        If Not pvLock(txtHeadBulstatText) Then
            m_oFP.SendCommand ucsIslCmdInitBulstat, RTrim(txtHeadBulstatText.Text) & "," & RTrim(txtHeadBulstatName.Text)
        End If
        If Not pvLock(txtHeadHeader3) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "2" & RTrim(txtHeadHeader3.Text)
        End If
        If Not pvLock(txtHeadHeader4) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "3" & RTrim(txtHeadHeader4.Text)
        End If
        If Not pvLock(txtHeadHeader5) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "4" & RTrim(txtHeadHeader5.Text)
        End If
        If Not pvLock(txtHeadHeader6) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "5" & RTrim(txtHeadHeader6.Text)
        End If
        If Not pvLock(txtHeadFooter1) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "6" & RTrim(txtHeadFooter1.Text)
        End If
        If Not pvLock(txtHeadFooter2) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "7" & RTrim(txtHeadFooter2.Text)
        End If
        If Not pvLock(chkHeadFormatInvoice) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "A" & chkHeadFormatInvoice.Value
        End If
        If Not pvLock(chkHeadSumEUR) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "E" & chkHeadSumEUR.Value & IIf(chkHeadRateEUR.Value, "," & txtHeadRate.Text, vbNullString)
        End If
        If Not pvLock(chkHeadAdvanceHeader) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "H" & chkHeadAdvanceHeader.Value
        End If
        If Not pvLock(chkHeadEmptyHeader) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "P" & chkHeadEmptyHeader.Value & "0" & chkHeadEmptyFooter.Value & chkHeadSumDivider.Value
        End If
        If Not pvLock(chkHeadVat) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "T" & chkHeadVat.Value
        End If
    Case ucsCmdDepartments
        If LenB(txtDepNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_DEP_SELECTED
            Exit Function
        End If
        m_oFP.SendCommand ucsIslCmdInitDepartment, txtDepNo.Text & "," & cobDepGroup.Text & "," & txtDepName.Text & IIf(LenB(txtDepName2.Text) <> 0, vbLf & txtDepName2.Text, vbNullString)
        '--- force refetch of department info
        m_vDeps(C_Lng(txtDepNo.Text)) = Empty
    Case ucsCmdOperators
        If LenB(txtOperNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_OPER_SELECTED
            Exit Function
        End If
        If LenB(txtOperPass.Text) <> 0 Then
            If Not pvIsPassCorrect(txtOperPass.Text) Then
                MsgBox MSG_INVALID_PASSWORD, vbExclamation
                Exit Function
            End If
            If txtOperPass.Text <> txtOperPass2.Text Then
                MsgBox MSG_PASSWORDS_MISMATCH, vbExclamation
                Exit Function
            End If
        End If
        sPass = InputBox(Printf(STR_OPER_PASS_PROMPT, txtOperNo.Text), STR_OPER_PASS_CAPTION)
        If StrPtr(sPass) = 0 Then
            Exit Function
        ElseIf Not pvIsPassCorrect(sPass) Then
            MsgBox MSG_INVALID_PASSWORD, vbExclamation
            Exit Function
        End If
        bCheckPass = True
        m_oFP.SendCommand ucsIslCmdInitOperatorName, txtOperNo.Text & "," & sPass & "," & txtOperName.Text
        If LenB(txtOperPass.Text) <> 0 Then
            m_oFP.SendCommand ucsIslCmdInitOperatorPassword, txtOperNo.Text & "," & sPass & "," & txtOperPass.Text
        End If
        bCheckPass = False
        '--- force refetch of oper info
        m_vOpers(C_Lng(txtOperNo.Text)) = Empty
    Case ucsCmdInvoiceNo
        If Not pvLock(txtInvStart) Then
            m_oFP.SendCommand ucsIslCmdInitInvoiceNo, txtInvStart.Text & "," & txtInvEnd.Text
        End If
    Case ucsCmdCashOper
        If Not pvLock(txtCashSum) And C_Dbl(txtCashSum.Text) <> 0 Then
            vResult = Split(m_oFP.SendCommand(ucsIslCmdFiscalServiceDeposit, IIf(optCashOut.Value, -1, 1) * Abs(C_Dbl(txtCashSum.Text))), ",")
            If At(vResult, 0) <> "P" Then
                MsgBox MSG_REQUEST_CANCELLED, vbExclamation
                pvStatus = MSG_REQUEST_CANCELLED
                Exit Function
            End If
        End If
    Case ucsCmdReports
        pvStatus = STR_STATUS_PRINT
        If optReportType(0).Value Then
            If chkReportItems.Value = vbChecked And chkReportDepartments.Value = vbChecked Then
                eCmd = ucsIslCmdPrintDailyReportItemsDepartments
            ElseIf chkReportItems.Value = vbChecked Then
                eCmd = ucsIslCmdPrintDailyReportItems
            ElseIf chkReportDepartments.Value = vbChecked Then
                eCmd = ucsIslCmdPrintDailyReportDepartments
            Else
                eCmd = ucsIslCmdPrintDailyReport
            End If
            '--- "rychno" razpechatwane na elektronna kontrolna lenta
            If chkReportClosure.Value = vbChecked Then
                vResult = Split(m_oFP.SendCommand(ucsIslCmdInitEcTape, "I"), ",")
                '--- print
                For lIdx = 1 To C_Lng(At(vResult, 1))
                    m_oFP.SendCommand ucsIslCmdInitEcTape, IIf(lIdx = 1, "PS", "CS")
                    If lIdx = C_Lng(At(vResult, 1)) Then
                        '--- erase
                        m_oFP.SendCommand ucsIslCmdInitEcTape, "E"
                    End If
                Next
            End If
            vResult = m_oFP.SendCommand(eCmd, IIf(chkReportClosure.Value = vbChecked, "0", "2") & "N")
            If m_oFP.Status(ucsIslStbPrintingError) Then
                '--- daisy: pechat po depatamenti
                If eCmd = ucsIslCmdPrintDailyReportDepartments Then
                    vResult = m_oFP.SendCommand(ucsIslCmdPrintDailyReport, IIf(chkReportClosure.Value = vbChecked, "8", "9") & "N")
                ElseIf eCmd = ucsIslCmdPrintDailyReportItemsDepartments Then
                    vResult = m_oFP.SendCommand(ucsIslCmdPrintDailyReportItems, IIf(chkReportClosure.Value = vbChecked, "8", "9") & "N")
                End If
            End If
        ElseIf optReportType(2).Value Then '--- by number
            If chkReportDetailed1.Value Then
                eCmd = ucsIslCmdPrintReportByNumberDetailed
            Else
                eCmd = ucsIslCmdPrintReportByNumberShort
            End If
            vResult = m_oFP.SendCommand(eCmd, txtReportStart.Text & "," & txtReportEnd.Text)
        ElseIf optReportType(3).Value Then '--- by date
            If chkReportDetailed2.Value Then
                eCmd = ucsIslCmdPrintReportByDateDetailed
            Else
                eCmd = ucsIslCmdPrintReportByDateShort
            End If
            vResult = m_oFP.SendCommand(eCmd, txtReportFD.Text & "," & txtReportTD.Text)
        ElseIf optReportType(5).Value Then '--- by operator
            vResult = m_oFP.SendCommand(ucsIslCmdPrintReportByOperators, vbNullString)
        End If
        pvStatus = vbNullString
    Case ucsCmdStatus
        pvStatus = STR_STATUS_REFRESH
        vResult = m_oFP.SendCommand(ucsIslCmdInfoStatus, "W")
        pvStatus = vbNullString
    Case ucsCmdDiagnostics
        pvStatus = STR_STATUS_PRINT
        vResult = m_oFP.SendCommand(ucsIslCmdPrintDiagnostics, vbNullString)
        pvStatus = vbNullString
    Case ucsCmdPaymentTypes
        If Not pvLock(txtPmtType(0)) Then
            m_oFP.SendCommand ucsIslCmdInitPaymentType, "I," & txtPmtType(0).Text
            If m_oFP.Status(ucsIslStbPrintingError) Then
                m_oFP.SendCommand ucsIslCmdExtendedInitText, "P61," & txtPmtType(0).Text
            End If
        End If
        If Not pvLock(txtPmtType(1)) Then
            m_oFP.SendCommand ucsIslCmdInitPaymentType, "J," & txtPmtType(1).Text
            If m_oFP.Status(ucsIslStbPrintingError) Then
                m_oFP.SendCommand ucsIslCmdExtendedInitText, "P62," & txtPmtType(1).Text
            End If
        End If
        If Not pvLock(txtPmtType(2)) Then
            m_oFP.SendCommand ucsIslCmdInitPaymentType, "K," & txtPmtType(2).Text
            If m_oFP.Status(ucsIslStbPrintingError) Then
                m_oFP.SendCommand ucsIslCmdExtendedInitText, "P63," & txtPmtType(2).Text
            End If
        End If
        If Not pvLock(txtPmtType(3)) Then
            m_oFP.SendCommand ucsIslCmdInitPaymentType, "L," & txtPmtType(3).Text
            If m_oFP.Status(ucsIslStbPrintingError) Then
                m_oFP.SendCommand ucsIslCmdExtendedInitText, "P64," & txtPmtType(3).Text
            End If
        End If
    Case ucsCmdItems
        If LenB(cobItemGroup.Text) = 0 Then
            pvStatus = STR_STATUS_NO_ITEM_GROUP
            Exit Function
        End If
        If LenB(txtItemPLU.Text) = 0 Then
            pvStatus = STR_STATUS_NO_ITEM_PLU
            Exit Function
        End If
        If LenB(txtItemPrice.Text) = 0 Then
            pvStatus = STR_STATUS_NO_ITEM_PRICE
            Exit Function
        End If
        If LenB(txtItemName.Text) = 0 Then
            pvStatus = STR_STATUS_NO_ITEM_NAME
            Exit Function
        End If
        If LenB(txtItemNo.Text) = 0 Then
            lIdx = UBound(m_vItems) + 1
            ReDim Preserve m_vItems(0 To lIdx)
        Else
            lIdx = lstItems.ListIndex + 1
        End If
        vResult = Split(m_oFP.SendCommand(ucsIslCmdInitItem, "P" & cobItemGroup.Text & txtItemPLU.Text & "," & txtItemPrice.Text & "," & txtItemName.Text), ",")
        If At(vResult, 0) = "F" Then
            pvStatus = Printf(STR_STATUS_ITEM_FAILURE_ADD, At(vResult, 1))
            Exit Function
        End If
        m_vItems(lIdx) = txtItemPLU.Text
    Case ucsCmdGraphicalLogo
        If Not pvLock(chkLogoPrint) Then
            m_oFP.SendCommand ucsIslCmdInitHeaderFooter, "L" & chkLogoPrint.Value
        End If
        If Not m_picLogo Is Nothing Then
            For lIdx = 0 To UBound(m_vLogo)
                pvStatus = Printf(STR_STATUS_SAVE_LOGO, lIdx + 1, UBound(m_vLogo) + 1)
                m_oFP.SendCommand ucsIslCmdInitLogo, lIdx & "," & m_vLogo(lIdx)
            Next
        End If
    End Select
    '--- success
    pvSaveData = True
    Exit Function
EH:
    If bCheckPass Then
        If m_oFP.Status(ucsIslStbInvalidFiscalMode) Then
            MsgBox MSG_REJECTED_PASSWORD, vbExclamation
            Exit Function
        End If
    End If
    If pvShowError() Then
        Exit Function
    End If
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    At = sDefault
    At = C_Str(vData(lIdx))
    On Error GoTo 0
End Function

Private Function pvShowError() As Boolean
    If Len(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
        pvStatus = m_oFP.LastError
        pvShowError = True
    End If
    If m_oFP.Status(ucsIslStbPrintingError) Then
        MsgBox m_oFP.ErrorText, vbExclamation
        pvStatus = m_oFP.ErrorText
        pvShowError = True
    End If
End Function

Private Function pvIsPassCorrect(sPass As String) As Boolean
    Dim lIdx            As Long
    Dim lChar           As Long
    
    If Len(sPass) >= 4 And Len(sPass) <= 6 Then
        For lIdx = 1 To Len(sPass)
            lChar = Asc(Mid(sPass, lIdx, 1))
            If lChar < 48 Or lChar > 57 Then '--- 48 = '0', 57 = '9'
                Exit Function
            End If
        Next
        pvIsPassCorrect = True
    End If
End Function

Private Sub pvApplyLogo(oPic As StdPicture, ByVal lTreshold As Long, ByVal bStretch As Boolean)
    Const FUNC_NAME     As String = "pvApplyLogo"
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lRGB            As Long
    
    On Error GoTo EH
    Set picLogo.Picture = Nothing
    If bStretch Then
        oPic.Render picLogo.hDC, 0, 0, _
            ScaleX(picLogo.Width, vbTwips, vbPixels), ScaleY(picLogo.Height, vbTwips, vbPixels), _
            0, oPic.Height, oPic.Width, -oPic.Height, ByVal 0
    Else
        oPic.Render picLogo.hDC, _
            (ScaleX(picLogo.Width, vbTwips, vbPixels) - ScaleX(oPic.Width, vbHimetric, vbPixels)) \ 2, _
            (ScaleY(picLogo.Height, vbTwips, vbPixels) - ScaleY(oPic.Height, vbHimetric, vbPixels)) \ 2, _
            ScaleX(oPic.Width, vbHimetric, vbPixels), ScaleY(oPic.Height, vbHimetric, vbPixels), _
            0, oPic.Height, oPic.Width, -oPic.Height, ByVal 0
    End If
    lTreshold = lTreshold * 256 / 100
    If lTreshold <= 0 Or lTreshold > 255 Then
        lTreshold = 128
    End If
    For lRow = 0 To UBound(m_vLogo)
        For lCol = 0 To Len(m_vLogo(0)) * 4 - 1
            lRGB = GetPixel(picLogo.hDC, lCol, lRow)
            '--- calc luminance
            lRGB = (lRGB And &HFF&) * 0.299 + ((lRGB \ &H100&) And &HFF&) * 0.587 + ((lRGB \ &H10000) And &HFF&) * 0.114
            pvLogoPixel(lCol, lRow) = (lRGB < lTreshold)
            Call SetPixel(picLogo.hDC, lCol, lRow, IIf(pvLogoPixel(lCol, lRow), LNG_LOGO_FORECOLOR, vbWhite))
        Next
    Next
    picLogo.Refresh
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function Printf(ByVal sText As String, ParamArray A()) As String
    Dim lI          As Long
    
    For lI = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lI - LBound(A) + 1), A(lI))
    Next
    Printf = sText
End Function

Private Function C_Lng(v As Variant) As Long
    On Error Resume Next
    C_Lng = CLng(v)
    On Error GoTo 0
End Function

Private Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Private Function C_Dbl(v As Variant) As Double
    On Error Resume Next
    C_Dbl = CDbl(v)
    On Error GoTo 0
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub lstCmds_Click()
    Const FUNC_NAME     As String = "lstCmds_Click"
    Dim lIdx            As Long
    Dim lVisibleFrame   As Long
    Dim dblTimer        As Double
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    dblTimer = Timer
    If lstCmds.ListIndex = ucsCmdSettings Or lstCmds.ListIndex = ucsCmdOperations Or lstCmds.ListIndex = ucsCmdAdmin Then
        lVisibleFrame = -1
        GoTo QH
    End If
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect And lstCmds.ListIndex <> ucsCmdStatus Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next
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
            pvStatus = Printf(STR_STATUS_SUCCESS_FETCH, Trim(lstCmds.List(lstCmds.ListIndex)), Round(Timer - dblTimer, 2))
        End If
        lVisibleFrame = lstCmds.ListIndex
    Else
        lVisibleFrame = -1
        If pvStatus = STR_STATUS_FETCHING Then
            pvStatus = vbNullString
        End If
    End If
QH:
    '--- might have missing entries in control array
    On Error Resume Next
    For lIdx = fraCommands.LBound To fraCommands.ubound
        fraCommands(lIdx).Visible = (lIdx = lVisibleFrame)
    Next
    On Error GoTo EH
    tmrDate.Enabled = (lVisibleFrame = ucsCmdDateTime)
    Call SendMessage(txtLog.hWnd, EM_SCROLLCARET, 0, ByVal 0&)
    '--- might have missing entries in control array
    On Error Resume Next
    For lIdx = cmdSave.LBound To cmdSave.ubound
        If Not cmdSave(lIdx).Visible Then
        Else
            cmdSave(lIdx).Default = True
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
    dblTimer = Timer
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next
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
                pvStatus = Printf(STR_STATUS_SUCCESS_SAVE, Trim(lstCmds.List(lstCmds.ListIndex)), Round(Timer - dblTimer, 2))
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

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    InitGlobals
    Set m_oFP = New cIslProtocol
    '--- init UI
    For Each vElem In Split(STR_COMMANDS, "|")
        lstCmds.AddItem vElem
    Next
    For Each vElem In Split(STR_GROUPS, "|")
        cobDepGroup.AddItem vElem
        cobItemGroup.AddItem vElem
    Next
    On Error Resume Next
    For lIdx = fraCommands.LBound To fraCommands.ubound
        fraCommands(lIdx).Visible = False
    Next
    On Error GoTo EH
    cmdExit.Left = -cmdExit.Width
    '--- login
    Visible = True
    pvStatus = STR_STATUS_ENUM_PORTS
    cobConnectPort.Clear
    For Each vElem In EnumSerialPorts
        cobConnectPort.AddItem vElem
    Next
    cobConnectPort.Text = GetSetting(App.Title, "Connect", "Port", vbNullString)
    chkConnectRemember.Value = -(LenB(cobConnectPort.Text) <> 0)
    If cobConnectPort.ListCount > 0 And Len(cobConnectPort.Text) = 0 Then
        cobConnectPort.ListIndex = 0
    End If
    cobConnectSpeed.Clear
    For Each vElem In Split(STR_SPEEDS, "|")
        cobConnectSpeed.AddItem vElem
    Next
    cobConnectSpeed.Text = GetSetting(App.Title, "Connect", "Speed", vbNullString)
    If cobConnectSpeed.ListCount > 0 And Len(cobConnectSpeed.Text) = 0 Then
        cobConnectSpeed.ListIndex = 0
    End If
    labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
    lstCmds.ListIndex = ucsCmdConnect
    If chkConnectRemember.Value Then
        cmdSave(ucsCmdConnect).Value = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
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
    cobDepGroup.Text = Mid(At(vResult, 0), 2)
    txtDepName.Text = At(Split(At(vResult, 4), vbLf), 0)
    txtDepName2.Text = At(Split(At(vResult, 4), vbLf), 1)
    txtDepSales.Text = At(vResult, 1)
    txtDepRecSum.Text = At(vResult, 2)
    txtDepTotalSum.Text = At(vResult, 3)
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
    txtOperName.Text = At(vResult, 5)
    txtOperFiscal.Text = At(vResult, 0)
    txtOperSells.Text = At(vResult, 1)
    txtOperDisc.Text = At(vResult, 2)
    txtOperSurcharge.Text = At(vResult, 3)
    txtOperVoid.Text = At(vResult, 4)
    txtOperPass.Text = vbNullString
    txtOperPass2.Text = vbNullString
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lstItems_Click()
    Const FUNC_NAME     As String = "lstItems_Click"
    Dim vResult         As Variant
    
    On Error GoTo EH
    If lstItems.ListIndex >= 0 Then
        txtItemNo.Text = lstItems.ListIndex + 1
        vResult = m_vItems(lstItems.ListIndex + 1)
    Else
        txtItemNo.Text = vbNullString
    End If
    cobItemGroup.Text = At(vResult, 3)
    If LenB(cobItemGroup.Text) = 0 Then
        cobItemGroup.ListIndex = 1
    End If
    txtItemPLU.Text = At(vResult, 1)
    txtItemPLU.Locked = (LenB(txtItemNo.Text) <> 0)
    txtItemPLU.BackColor = IIf(LenB(txtItemNo.Text) <> 0, vbButtonFace, vbWindowBackground)
    txtItemPrice.Text = At(vResult, 4)
    txtItemName.Text = At(vResult, 7)
    txtItemAmount.Text = At(vResult, 5)
    txtItemSum.Text = At(vResult, 6)
    txtItemTime.Text = At(vResult, 2)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub picLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "picLogo_MouseDown"
    Dim lRow            As Long
    Dim lCol            As Long
    
    On Error GoTo EH
    lRow = Y / Screen.TwipsPerPixelY
    lCol = X / Screen.TwipsPerPixelX
    pvLogoPixel(lCol, lRow) = Not pvLogoPixel(lCol, lRow)
    Call SetPixel(picLogo.hDC, lCol, lRow, IIf(pvLogoPixel(lCol, lRow), LNG_LOGO_FORECOLOR, vbWhite))
    picLogo.Refresh
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
    txtDateCompDate.Text = Format(Now, "dd-MM-yy")
    txtDateCompTime.Text = Format(Now, "hh:mm:ss")
End Sub

Private Sub cmdDateTransfer_Click()
    txtDateDate.Text = txtDateCompDate.Text
    txtDateTime.Text = txtDateCompTime.Text
End Sub

Private Sub cmdItemNew_Click()
    lstItems.ListIndex = -1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtLogoTreshold_Change()
    If Not m_picLogo Is Nothing Then
        pvApplyLogo m_picLogo, C_Lng(txtLogoTreshold.Text), optLogoStretch.Value
    End If
End Sub

Private Sub optLogoCenter_Click()
    txtLogoTreshold_Change
End Sub

Private Sub optLogoStretch_Click()
    txtLogoTreshold_Change
End Sub

Private Sub cmdLogoOpen_Click()
    Const FUNC_NAME     As String = "cmdLogoOpen_Click"
    Const STR_TITLE     As String = "Logo"
    Const STR_FILTER    As String = "Graphical files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|All files (*.*)|*.*"
    Dim uOFN            As OPENFILENAME
    Dim sFilter         As String
    Dim sTitle          As String
    Dim sBuffer         As String
    Dim sFile           As String
    
    On Error GoTo EH
    sFilter = StrConv(Replace(STR_FILTER, "|", vbNullChar), vbFromUnicode)
    sTitle = StrConv(STR_TITLE, vbFromUnicode)
    sBuffer = String(1000, 0)
    If OsVersion >= 500 Then
        uOFN.lStructSize = Len(uOFN)
    Else
        uOFN.lStructSize = Len(uOFN) - 12
    End If
    uOFN.Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT Or OFN_EXPLORER Or OFN_ENABLESIZING
    uOFN.hWndOwner = Me.hWnd
    uOFN.lpstrFilter = StrPtr(sFilter)
    uOFN.nFilterIndex = 1
    uOFN.lpstrTitle = StrPtr(sTitle)
    uOFN.lpstrFile = StrPtr(sBuffer)
    uOFN.nMaxFile = Len(sBuffer)
    If GetOpenFileName(uOFN) Then
        sFile = StrConv(sBuffer, vbUnicode)
        sFile = Left(sFile, InStr(sFile, Chr$(0)) - 1)
        picLogo.BackColor = vbWhite
        Set m_picLogo = LoadPicture(sFile)
        pvApplyLogo m_picLogo, C_Lng(txtLogoTreshold.Text), optLogoStretch.Value
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cmdOperReset_Click()
    Const FUNC_NAME     As String = "cmdOperReset_Click"
    Dim sPass           As String
    
    On Error GoTo EH
    If LenB(txtOperNo.Text) = 0 Then
        Exit Sub
    End If
    sPass = InputBox(Printf(STR_OPER_PASS_PROMPT, txtOperNo.Text), STR_OPER_PASS_CAPTION)
    If StrPtr(sPass) = 0 Then
        Exit Sub
    ElseIf Not pvIsPassCorrect(sPass) Then
        MsgBox MSG_INVALID_PASSWORD, vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If Not m_oFP.IsConnected Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_OPER_RESETTING
    m_oFP.SendCommand ucsIslCmdInitOperatorReset, txtOperNo.Text & "," & sPass
    m_vOpers(C_Lng(txtOperNo.Text)) = Empty
    pvFetchData ucsCmdOperators
    pvStatus = Printf(STR_STATUS_OPER_SUCCESS_RESET, txtOperNo.Text)
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    If pvShowError() Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cmdItemDelete_Click()
    Const FUNC_NAME     As String = "cmdItemDelete_Click"
    Dim sPLU            As String
    
    On Error GoTo EH
    If LenB(txtItemPLU.Text) = 0 Then
        Exit Sub
    End If
    sPLU = txtItemPLU.Text
    If MsgBox(Printf(MSG_CONFIRM_ITEM_DELETE, sPLU), vbQuestion Or vbYesNo) = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If Not m_oFP.IsConnected Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_ITEM_DELETING
    m_oFP.SendCommand ucsIslCmdInitItem, "D" & sPLU
    m_vItems = Empty
    pvFetchData ucsCmdItems
    pvStatus = Printf(STR_STATUS_ITEM_SUCCESS_DELETE, sPLU)
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    If pvShowError() Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cmdStatusReset_Click()
    Const FUNC_NAME     As String = "cmdStatusReset_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    If Not m_oFP.IsConnected Then
        pvStatus = STR_STATUS_CONNECTING
        On Error Resume Next
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_RESETTING
'    If Left(m_oFP.SendCommand(ucsIslCmdInfoTransaction), 1) = "1" Then
'        If m_oFP.Status(ucsIslStbFiscalPrinting) Then
            '--- note: when printing invoice, if no contragent info set then cancel fails!
            m_oFP.SendCommand ucsIslCmdFiscalCgInfo, "0000000000"
            '--- note: FP3530 moje da anulira winagi, FP550F ne moje
            m_oFP.SendCommand ucsIslCmdFiscalCancel, vbNullString
            '--- zaradi FP550F
            m_oFP.SendCommand ucsIslCmdFiscalClose, vbNullString
'        Else
            m_oFP.SendCommand ucsIslCmdNonFiscalClose, vbNullString
'        End If
'    End If
    pvFetchData ucsCmdStatus
    pvStatus = vbNullString
    If m_oFP.IsConnected Then
        m_oFP.Disconnect
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    If pvShowError() Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFP_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFP_CommandComplete"
    
    On Error GoTo EH
    m_sLog = m_sLog & lCmd & IIf(LenB(sData) <> 0, "<-" & sData, vbNullString) & IIf(LenB(sResult) <> 0, "->" & sResult, vbNullString) & vbCrLf
    If LenB(m_oFP.LastError) <> 0 Then
        m_sLog = m_sLog & m_oFP.LastError & vbCrLf
    End If
    If m_oFP.Status(ucsIslStbPrintingError) Then
        m_sLog = m_sLog & m_oFP.StatusText & vbCrLf & m_oFP.DipText & vbCrLf & m_oFP.MemoryText & vbCrLf
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

