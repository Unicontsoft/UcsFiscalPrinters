VERSION 5.00
Begin VB.Form frmIslSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������� ICL ��������"
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
   Icon            =   "frmIslSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   8172
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   5
      Left            =   2250
      TabIndex        =   174
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtSettNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4044
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1452
         Width           =   825
      End
      Begin VB.ListBox lstSettings 
         Height          =   4152
         IntegralHeight  =   0   'False
         Left            =   168
         TabIndex        =   39
         Top             =   1428
         Width           =   2265
      End
      Begin VB.TextBox txtSettValue 
         Height          =   285
         Left            =   2604
         MaxLength       =   24
         TabIndex        =   41
         Top             =   2076
         Width           =   2985
      End
      Begin VB.TextBox txtInvCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   990
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   42
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtInvEnd 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2070
         TabIndex        =   38
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtInvStart 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2070
         TabIndex        =   37
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����:"
         Height          =   192
         Left            =   2604
         TabIndex        =   238
         Top             =   1452
         Width           =   1908
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   204
         Left            =   2604
         TabIndex        =   237
         Top             =   1812
         Width           =   1908
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� �����:"
         Height          =   195
         Left            =   180
         TabIndex        =   178
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
         TabIndex        =   176
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
         TabIndex        =   175
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   7
      Left            =   2250
      TabIndex        =   158
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdOperReset 
         Caption         =   "��������"
         Height          =   375
         Left            =   4320
         TabIndex        =   62
         Top             =   3150
         Width           =   1275
      End
      Begin VB.TextBox txtOperPass2 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   64
         Top             =   4410
         Width           =   1545
      End
      Begin VB.TextBox txtOperPass 
         Height          =   285
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   63
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1545
      End
      Begin VB.TextBox txtOperName 
         Height          =   285
         Left            =   2610
         MaxLength       =   24
         TabIndex        =   61
         Top             =   900
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   65
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstOpers 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   60
         Top             =   252
         Width           =   2265
      End
      Begin VB.TextBox txtOperNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   159
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   167
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   161
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
         TabIndex        =   160
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   4
      Left            =   2250
      TabIndex        =   133
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   6
      Left            =   2250
      TabIndex        =   196
      Top             =   84
      Width           =   5775
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   7
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   58
         Top             =   2790
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   6
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   56
         Top             =   2430
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   5
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   54
         Top             =   2070
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   4
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   52
         Top             =   1710
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   3
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   50
         Top             =   1350
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   2
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   48
         Top             =   990
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   1
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   46
         Top             =   630
         Width           =   891
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   300
         Index           =   0
         Left            =   4704
         MaxLength       =   40
         TabIndex        =   44
         Top             =   270
         Width           =   891
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   4
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   51
         Top             =   1710
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   5
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   53
         Top             =   2070
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   6
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   55
         Top             =   2430
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Height          =   300
         Index           =   7
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   57
         Top             =   2790
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   3
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   2
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   990
         Width           =   2520
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   1
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   630
         Width           =   2520
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   59
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtPmtType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Index           =   0
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   270
         Width           =   2520
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 5:"
         Height          =   195
         Left            =   180
         TabIndex        =   236
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 6:"
         Height          =   195
         Left            =   180
         TabIndex        =   235
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 7:"
         Height          =   195
         Left            =   180
         TabIndex        =   234
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 8:"
         Height          =   195
         Left            =   180
         TabIndex        =   233
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��� ������� 4:"
         Height          =   195
         Left            =   180
         TabIndex        =   200
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
         TabIndex        =   199
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
         TabIndex        =   198
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
         TabIndex        =   197
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   16
      Left            =   2250
      TabIndex        =   188
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdStatusReset 
         Caption         =   "�����"
         Height          =   375
         Left            =   180
         TabIndex        =   101
         Top             =   5220
         Width           =   1185
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����������"
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   102
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstStatus 
         Height          =   4920
         IntegralHeight  =   0   'False
         ItemData        =   "frmIslSetup.frx":000C
         Left            =   84
         List            =   "frmIslSetup.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   100
         Top             =   168
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   15
      Left            =   2250
      TabIndex        =   189
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtDiagFirmware 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtDiagChecksum 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtDiagSwitches 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   99
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firmware:"
         Height          =   195
         Left            =   180
         TabIndex        =   195
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
         TabIndex        =   194
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
         TabIndex        =   193
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   12
      Left            =   2250
      TabIndex        =   179
      Top             =   90
      Width           =   5775
      Begin VB.OptionButton optCashOut 
         Caption         =   "���������"
         Height          =   285
         Left            =   3528
         TabIndex        =   80
         Top             =   1620
         Width           =   1455
      End
      Begin VB.OptionButton optCashIn 
         Caption         =   "���������"
         Height          =   285
         Left            =   2070
         TabIndex        =   79
         Top             =   1620
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtCashSum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   81
         Top             =   1980
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����/�����"
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   82
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtCashOut 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   184
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
         TabIndex        =   182
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
         TabIndex        =   180
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
         TabIndex        =   186
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
         TabIndex        =   185
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
         TabIndex        =   183
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
         TabIndex        =   181
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   10
      Left            =   2250
      TabIndex        =   157
      Top             =   90
      Width           =   5775
      Begin VB.Frame Frame1 
         Caption         =   "���������"
         Height          =   1185
         Left            =   180
         TabIndex        =   226
         Top             =   3780
         Width           =   5415
         Begin VB.OptionButton optLogoStretch 
            Caption         =   "���������"
            Height          =   285
            Left            =   1800
            TabIndex        =   232
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optLogoCenter 
            Caption         =   "����������"
            Height          =   285
            Left            =   180
            TabIndex        =   231
            Top             =   720
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.CommandButton cmdLogoOpen 
            Caption         =   "�����"
            Height          =   375
            Left            =   3960
            TabIndex        =   228
            Top             =   270
            Width           =   1275
         End
         Begin VB.TextBox txtLogoTreshold 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   227
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
            TabIndex        =   230
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
            TabIndex        =   229
            Top             =   270
            Width           =   555
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox chkLogoPrint 
         Caption         =   "����� �������� ���� ����� header"
         Height          =   285
         Left            =   180
         TabIndex        =   224
         Top             =   270
         Width           =   4245
      End
      Begin VB.PictureBox picLogoScroll 
         BorderStyle     =   0  'None
         Height          =   2445
         Left            =   180
         ScaleHeight     =   2448
         ScaleWidth      =   5412
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   720
         Width           =   5415
         Begin VB.HScrollBar scbLogoHor 
            CausesValidation=   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   223
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
            TabIndex        =   222
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
         TabIndex        =   220
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label labLogoInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   195
         Left            =   180
         TabIndex        =   225
         Top             =   3330
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   8
      Left            =   2250
      TabIndex        =   145
      Top             =   90
      Width           =   5775
      Begin VB.ComboBox cobDepGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   67
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
         TabIndex        =   155
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
         TabIndex        =   153
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
         TabIndex        =   151
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
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtDepName2 
         Height          =   285
         Left            =   2610
         MaxLength       =   36
         TabIndex        =   69
         Top             =   1890
         Width           =   2985
      End
      Begin VB.ListBox lstDeps 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   66
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   70
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtDepName 
         Height          =   285
         Left            =   2610
         MaxLength       =   31
         TabIndex        =   68
         Top             =   1260
         Width           =   2985
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� �� ����:"
         Height          =   195
         Left            =   2610
         TabIndex        =   156
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
         TabIndex        =   154
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
         TabIndex        =   152
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   147
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
         TabIndex        =   146
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   13
      Left            =   2250
      TabIndex        =   187
      Top             =   90
      Width           =   5775
      Begin VB.OptionButton optReportType 
         Caption         =   "����� ���������"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   95
         Top             =   3420
         Width           =   5145
      End
      Begin VB.CheckBox chkReportDetailed1 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   90
         Top             =   1710
         Width           =   1725
      End
      Begin VB.TextBox txtReportStart 
         Height          =   285
         Left            =   1800
         TabIndex        =   88
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txtReportEnd 
         Height          =   285
         Left            =   3420
         TabIndex        =   89
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CheckBox chkReportDetailed2 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   94
         Top             =   2970
         Width           =   1725
      End
      Begin VB.TextBox txtReportTD 
         Height          =   285
         Left            =   3420
         TabIndex        =   93
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtReportFD 
         Height          =   285
         Left            =   1800
         TabIndex        =   92
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   98
         Top             =   5220
         Width           =   1275
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "��������� ���� �� ������"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   96
         Top             =   3780
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ���� �� ����� (DDMMYY, MMYY, YY)"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   91
         Top             =   2160
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "���������� ����� �� ����� �� ���� (4 �����)"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   87
         Top             =   990
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "������� ������"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   97
         Top             =   4140
         Width           =   5145
      End
      Begin VB.CheckBox chkReportDepartments 
         Caption         =   "������������"
         Height          =   285
         Left            =   3960
         TabIndex        =   86
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportItems 
         Caption         =   "��������"
         Height          =   285
         Left            =   2430
         TabIndex        =   85
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportClosure 
         Caption         =   "��������"
         Height          =   285
         Left            =   900
         TabIndex        =   84
         Top             =   630
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "������ �������� �����"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   83
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
         TabIndex        =   204
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
         Top             =   2520
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   9
      Left            =   2250
      TabIndex        =   206
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtItemTime 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2985
      End
      Begin VB.CommandButton cmdItemDelete 
         Caption         =   "���������"
         Height          =   375
         Left            =   4320
         TabIndex        =   77
         Top             =   3960
         Width           =   1275
      End
      Begin VB.CommandButton cmdItemNew 
         Caption         =   "���"
         Height          =   375
         Left            =   2970
         TabIndex        =   76
         Top             =   3960
         Width           =   1275
      End
      Begin VB.TextBox txtItemSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   3510
         Width           =   1545
      End
      Begin VB.TextBox txtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   9
         TabIndex        =   73
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtItemPLU 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   72
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtItemAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1545
      End
      Begin VB.ComboBox cobItemGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   74
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   2610
         MaxLength       =   25
         TabIndex        =   75
         Top             =   1980
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�����"
         Height          =   375
         Index           =   13
         Left            =   4320
         TabIndex        =   78
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstItems 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   71
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
         TabIndex        =   207
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
         TabIndex        =   218
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
         TabIndex        =   216
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   209
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
         TabIndex        =   208
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   2
      Left            =   2250
      TabIndex        =   105
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
         TabIndex        =   114
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
         TabIndex        =   112
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
         TabIndex        =   110
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label labTaxCountry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   195
         Left            =   180
         TabIndex        =   113
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
         TabIndex        =   111
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
         TabIndex        =   109
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   3
      Left            =   2250
      TabIndex        =   126
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
         Top             =   1530
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   17
      Left            =   2250
      TabIndex        =   205
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtLog 
         Height          =   5505
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   0
      Left            =   2250
      TabIndex        =   104
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
         Top             =   1530
         Width           =   1785
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   90
      TabIndex        =   219
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
      TabIndex        =   144
      Top             =   5940
      Width           =   7935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmIslSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFP20 (c) 2008-2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Nastrojki na FP po ISL protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmIslSetup"

'=========================================================================
' API
'=========================================================================

Private Const EM_SCROLLCARET                    As Long = &HB7

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const CAP_MSG               As String = "��������� ISL ��������"
Private Const LNG_NUM_DEPS          As Long = 100
Private Const LNG_NUM_OPERS         As Long = 100
Private Const LNG_NUM_SETTINGS      As Long = 100
Private Const FORMAT_CURRENCY       As String = "#0.00######"
Private Const LNG_LOGO_FORECOLOR    As Long = &H800000
Private Const PROGID_PROTOCOL       As String = LIB_NAME & ".cIslProtocol"
'--- strings
Private Const STR_SPEEDS            As String = "9600|19200|38400|57600|115200"
Private Const STR_COMMANDS          As String = "������ �������|���������|    ������� ����������|    ���� � ���|    �������|    ������ �� �������|    ������ ��������|    ���������|    ������������|    ��������|    �������� ����|��������|    ���������/���������|    ����� ������|�������������|    �����������|    ������|    ������ �����������"
Private Const STR_COUNTRIES         As String = "|�����|������|�������|������|�����|�������|��������|7|8|�������"
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
Private Const STR_STATUS_OPER_FETCH As String = "���������� �������� %1 �� " & LNG_NUM_OPERS & "..."
Private Const STR_STATUS_OPER_RESETTING As String = "��������..."
Private Const STR_STATUS_OPER_SUCCESS_RESET As String = "������� �������� �� �������� %1"
Private Const STR_STATUS_SETT_FETCH As String = "���������� ��������� %1 �� " & LNG_NUM_SETTINGS & "..."
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
Private Const STR_PAYMENT_TYPES     As String = "� ����|� ������� �����|� ���|� �������� �����"
Private Const STR_TAXCOUNTRY        As String = "�������:|Sw1..Sw8:"
Private Const STR_FP_STATUSES       As String = "0.7 ����������|0.6 ������� � ������� �� ��������|0.5 ���� ������ (OR �� ������ ������, ��������� � #)|0.4 # ���������� �� ���������� ���������� ��� ������������|0.3 �� � ������� ��������� �������|0.2 ���������� �� � ���������|0.1 # ����� �� ���������� ������� � ���������|0.0 # ���������� ����� ���� ����������� ������|1.7 ����������|1.6 ���������� ������� �������� �� ��������|1.5 ������� � �������� ��� �� ����� �� �������� �� 90 ������� �����|1.4 # ���������� � ����������� �� ������������ �� ������������ ����� (RAM) ���� ���������|1.3 # ����� ������� (���������� �� ������ ����� � � ��������� RESET)|1.2 # ��������� � ���������� �� ������������ �����|1.1 # ������������ �� ��������� �� � ��������� � ������� �������� �����|1.0 ��� ���������� �� ��������� �� � �������� ���������� �� ����� ������ �� ������" & _
                                                "|2.7 ����������|2.6 ��������� �� �����|2.5 ������� � �������� ���|2.4 ������ ���� �� ���� (��-����� �� 10 MB �� ���� ��������)|2.3 ������� � �������� ���|2.2 ���� �� ���� (��-����� �� 1 MB �� ���� ��������)|2.1 �������� � ����� ������|2.0 # �������� � ��������|3.7 ����������|3.6 ��������� �� Sw7|3.5 ��������� �� Sw6|3.4 ��������� �� Sw5|3.3 ��������� �� Sw4|3.2 ��������� �� Sw3|3.1 ��������� �� Sw2|3.0 ��������� �� Sw1" & _
                                                "|4.7 ����������|4.6 �� �� ��������|4.5 OR �� ������ ������, ��������� � * �� ������� 4 � 5|4.4 * ���������� ����� � �����|4.3 ��� ����� �� ��-����� �� 50 ������ ��� ��|4.2 �������� �� ������������ ����� �� �������� � ����� �� ���������� �����|4.1 ������� � ���|4.0 * ��� ������ ��� ����� ��� ���������� �����|5.7 ����������|5.6 �� �� ��������|5.5 �������� �� ������� ������|5.4 �������� �� ���� ������ ��������� ������|5.3 ��������� � ��� �������� �����|5.2 * ���������� ����� ��� ���������� ����� �� � �������|5.1 ���������� ����� � �����������|5.0 * ���������� ����� � ���������� � ����� READONLY (���������)"
'--- messages
Private Const MSG_INVALID_PASSWORD  As String = "���������� ������" & vbCrLf & vbCrLf & "�������� �� ������� �� 4 �� 6 �����"
Private Const MSG_PASSWORDS_MISMATCH As String = "�������� �� ��������"
Private Const MSG_REJECTED_PASSWORD As String = "��������� ������ �� ��������"
Private Const MSG_REQUEST_CANCELLED As String = "�������� � ��������"
Private Const MSG_CONFIRM_ITEM_DELETE As String = "������� �� �� �������� ������� PLU %1?"
Private Const MSG_CANNOT_ACCESS_PRINTER_PROXY As String = "������ ��� ��������� �� ��������� �� ������ �� �������� ������� %1." & vbCrLf & vbCrLf & "%2"

Private m_oFP                   As cIslProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private WithEvents m_oFPSink    As cIslProtocol
Attribute m_oFPSink.VB_VarHelpID = -1
Private m_sLog                  As String
Private m_vDeps                 As Variant
Private m_vOpers                As Variant
Private m_vItems                As Variant
Private m_vLogo                 As Variant
Private m_vSettings             As Variant
Private m_picLogo               As StdPicture
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

Private Property Get pvLogoPixel(ByVal lX As Long, ByVal lY As Long) As Boolean
    On Error GoTo EH
    pvLogoPixel = ((CLng("&H" & Mid$(m_vLogo(lY), 1 + 2 * (lX \ 8), 2)) And (2 ^ (7 - lX Mod 8))) <> 0)
    Exit Property
EH:
    Debug.Print lY, lX, Mid$(m_vLogo(lY), 1 + 2 * (lX \ 8), 2)
    Resume Next
End Property

Private Property Let pvLogoPixel(ByVal lX As Long, ByVal lY As Long, ByVal bValue As Boolean)
    Dim lValue          As Long
    
    lValue = C_Lng("&H" & Mid$(m_vLogo(lY), 1 + 2 * (lX \ 8), 2))
    If bValue Then
        lValue = lValue Or (2 ^ (7 - lX Mod 8))
    Else
        lValue = lValue And (Not 2 ^ (7 - lX Mod 8))
    End If
    Mid$(m_vLogo(lY), 1 + 2 * (lX \ 8), 2) = Right$("0" & Hex$(lValue), 2)
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

Private Function pvGetPrinter(sServer As String, sError As String) As cIslProtocol
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
    Dim lRow            As Long
    Dim vElem           As Variant
    Dim lWidth          As Long
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCmd <> ucsCmdConnect And lstCmds.ListIndex <> ucsCmdStatus Then
        pvStatus = labConnectCurrent.Caption
        Exit Function
    End If
    Select Case eCmd
    Case ucsCmdConnect
        pvStatus = labConnectCurrent.Caption
    Case ucsCmdTaxInfo
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0"), ",")
        txtTaxMemModule.Text = At(vResult, 5)
        txtTaxSerNo.Text = At(vResult, 4)
        If Len(At(vResult, 3)) <= 2 Then
            txtTaxCountry.Text = At(Split(STR_COUNTRIES, "|"), C_Lng(At(vResult, 3)) + 1)
            labTaxCountry.Caption = Split(STR_TAXCOUNTRY, "|")(0)
        Else
            txtTaxCountry.Text = At(vResult, 3)
            labTaxCountry.Caption = Split(STR_TAXCOUNTRY, "|")(1)
        End If
        m_oFP.Exceptions = False
        vResult = Split(m_oFP.SendCommand(ucsFpcInitDecimals), ",")
        m_oFP.Exceptions = True
        If UBound(vResult) > 0 Then
            txtTaxDecimals.Text = C_Lng(At(vResult, 1))
            txtTaxCurrency.Text = Trim(At(vResult, 2))
            txtTaxRates.Text = C_Lng(At(vResult, 3))
        Else
            LockControl(txtTaxDecimals) = True
            LockControl(txtTaxCurrency) = True
            LockControl(txtTaxRates) = True
        End If
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoTaxRates), ",")
        txtTaxGroup1.Text = C_Lng(At(vResult, 0))
        txtTaxGroup2.Text = C_Lng(At(vResult, 1))
        txtTaxGroup3.Text = C_Lng(At(vResult, 2))
        txtTaxGroup4.Text = C_Lng(At(vResult, 3))
    Case ucsCmdDateTime
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoDateTime), " ")
        txtDateDate.Text = At(vResult, 0)
        txtDateTime.Text = At(vResult, 1)
        tmrDate_Timer
    Case ucsCmdHeaderFooter
        m_oFP.Exceptions = False
        txtHeadHeader1.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I0")
        LockControl(txtHeadHeader1) = m_oFP.Status(ucsStbPrintingError)
        txtHeadHeader2.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I1")
        LockControl(txtHeadHeader2) = m_oFP.Status(ucsStbPrintingError)
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoBulstat), ",")
        txtHeadBulstatName.Text = At(vResult, 1)
        txtHeadBulstatText.Text = At(vResult, 0)
        LockControl(txtHeadBulstatName) = m_oFP.Status(ucsStbPrintingError)
        LockControl(txtHeadBulstatText) = m_oFP.Status(ucsStbPrintingError)
        txtHeadHeader3.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I2")
        LockControl(txtHeadHeader3) = m_oFP.Status(ucsStbPrintingError)
        txtHeadHeader4.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I3")
        LockControl(txtHeadHeader4) = m_oFP.Status(ucsStbPrintingError)
        txtHeadHeader5.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I4")
        LockControl(txtHeadHeader5) = m_oFP.Status(ucsStbPrintingError)
        txtHeadHeader6.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I5")
        LockControl(txtHeadHeader6) = m_oFP.Status(ucsStbPrintingError)
        txtHeadFooter1.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
        LockControl(txtHeadFooter1) = m_oFP.Status(ucsStbPrintingError)
        txtHeadFooter2.Text = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
        LockControl(txtHeadFooter2) = m_oFP.Status(ucsStbPrintingError)
        chkHeadFormatInvoice.Value = -(m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IA") = "1")
        LockControl(chkHeadFormatInvoice) = m_oFP.Status(ucsStbPrintingError)
        vResult = Split(m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IE"), ",")
        chkHeadSumEUR.Value = -(At(vResult, 0) = "1")
        chkHeadRateEUR.Value = chkHeadSumEUR.Value
        txtHeadRate.Text = Trim(At(vResult, 1))
        LockControl(chkHeadSumEUR) = m_oFP.Status(ucsStbPrintingError)
        LockControl(chkHeadRateEUR) = m_oFP.Status(ucsStbPrintingError)
        LockControl(txtHeadRate) = m_oFP.Status(ucsStbPrintingError)
        chkHeadAdvanceHeader.Value = -(m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IH") = "1")
        LockControl(chkHeadAdvanceHeader) = m_oFP.Status(ucsStbPrintingError)
        vResult = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IP")
        chkHeadEmptyHeader.Value = -(Mid$(vResult, 1, 1) = "1")
        chkHeadEmptyFooter.Value = -(Mid$(vResult, 3, 1) = "1")
        chkHeadSumDivider.Value = -(Mid$(vResult, 4, 1) = "1")
        LockControl(chkHeadEmptyHeader) = m_oFP.Status(ucsStbPrintingError)
        LockControl(chkHeadEmptyFooter) = m_oFP.Status(ucsStbPrintingError)
        LockControl(chkHeadSumDivider) = m_oFP.Status(ucsStbPrintingError)
        chkHeadVat.Value = -(m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IT") = "1")
        LockControl(chkHeadVat) = m_oFP.Status(ucsStbPrintingError)
        m_oFP.Exceptions = True
    Case ucsCmdInvoiceNo
        m_oFP.Exceptions = False
        vResult = Split(m_oFP.SendCommand(ucsFpcInitInvoiceNo), ",")
        txtInvStart.Text = C_Dbl(At(vResult, 0))
        txtInvEnd.Text = C_Dbl(At(vResult, 1))
        txtInvCurrent.Text = C_Dbl(At(vResult, 2))
        LockControl(txtInvStart) = m_oFP.Status(ucsStbPrintingError)
        LockControl(txtInvEnd) = m_oFP.Status(ucsStbPrintingError)
        LockControl(txtInvCurrent) = m_oFP.Status(ucsStbPrintingError)
        If m_oFP.IsDaisy Then
            If Not IsArray(m_vSettings) Then
                ReDim m_vSettings(0 To LNG_NUM_SETTINGS) As Variant
            End If
            For lIdx = 1 To UBound(m_vSettings)
                sText = vbNullString
                If Not IsArray(m_vSettings(lIdx)) Then
                    pvStatus = Printf(STR_STATUS_SETT_FETCH, lIdx)
                    vResult = Split(m_oFP.SendCommand(ucsFpcExtendedInitSetting, "N" & C_Str(lIdx)), vbTab)
                    If m_oFP.Status(ucsStbPrintingError) Or LenB(m_oFP.LastError) <> 0 Then
                        vResult = Split(m_oFP.SendCommand(ucsFpcExtendedInitSetting, "R" & C_Str(lIdx)), vbTab)
                    End If
                    If m_oFP.Status(ucsStbPrintingError) Or LenB(m_oFP.LastError) <> 0 Then
                        ReDim Preserve m_vSettings(0 To lIdx - 1) As Variant
                        Exit For
                    End If
                    sText = lIdx & ": " & At(vResult, 1)
                    m_vSettings(lIdx) = Split(At(vResult, 0), ",")
                End If
                If lstSettings.ListCount < lIdx Then
                    lstSettings.AddItem sText
                ElseIf lstSettings.List(lIdx - 1) <> sText And LenB(sText) <> 0 Then
                    lstSettings.List(lIdx - 1) = sText
                End If
            Next
            lstSettings_Click
            pvStatus = vbNullString
        ElseIf m_oFP.IsIncotex Then
            vResult = Split(m_oFP.SendCommand(ucsFpcInfoSums), ",")
            If Val(At(vResult, -1)) <> 0 Then
                txtInvStart.Text = Val(At(vResult, -1)) - 1
                txtInvEnd.Text = txtInvStart.Text
                txtInvCurrent.Text = txtInvStart.Text
            End If
            LockControl(txtInvStart) = False
            LockControl(txtInvEnd) = False
            LockControl(txtInvCurrent) = False
        Else
            LockControl(lstSettings) = True
            LockControl(txtSettValue) = True
        End If
        m_oFP.Exceptions = True
    Case ucsCmdPaymentTypes
        m_oFP.Exceptions = False
        If m_oFP.IsDaisy Then
            For lIdx = 0 To 7
                txtPmtType(lIdx).Text = m_oFP.SendCommand(ucsFpcExtendedInitText, "R" & (60 + lIdx))
                vResult = Split(m_oFP.SendCommand(ucsFpcExtendedInitCurrencyRate, "R" & (lIdx)), vbTab)
                txtPmtRate(lIdx).Text = At(vResult, 1)
                LockControl(txtPmtType(lIdx)) = lIdx < 1
                LockControl(txtPmtRate(lIdx)) = lIdx < 1
                txtPmtRate(lIdx).Visible = lIdx > 0
            Next
        ElseIf m_oFP.IsIncotex Then
            vResult = Split(STR_PAYMENT_TYPES, "|")
            For lIdx = 0 To 7
                If lIdx < 4 Then
                    LockControl(txtPmtType(lIdx)) = True
                    If lIdx < 1 Then
                        txtPmtType(lIdx).Text = At(vResult, lIdx)
                    ElseIf lIdx < 3 Then
                        txtPmtType(lIdx).Text = At(Split(m_oFP.SendCommand(ucsFpcExtendedInitText, "R" & (9 + lIdx)), ","), 1)
                    Else
                        txtPmtType(lIdx).Text = vbNullString
                    End If
                ElseIf lIdx < 6 Then
                    txtPmtType(lIdx).Text = At(Split(m_oFP.SendCommand(ucsFpcExtendedInitText, "R" & (8 + lIdx)), ","), 1)
                Else
                    LockControl(txtPmtType(lIdx)) = True
                    txtPmtType(lIdx).Text = vbNullString
                End If
                txtPmtRate(lIdx).Visible = False
            Next
        Else
            vResult = Split(STR_PAYMENT_TYPES, "|")
            For lIdx = 0 To 3
                txtPmtType(lIdx).Text = At(vResult, lIdx)
                LockControl(txtPmtType(lIdx)) = True
                txtPmtRate(lIdx).Visible = False
            Next
            For lIdx = 4 To 7
                txtPmtType(lIdx).Text = m_oFP.SendCommand(ucsFpcInitPaymentType, Chr$(69 + lIdx))
                LockControl(txtPmtType(lIdx)) = m_oFP.Status(ucsStbPrintingError)
                txtPmtRate(lIdx).Visible = False
            Next
        End If
        m_oFP.Exceptions = True
    Case ucsCmdOperators
        m_oFP.Exceptions = False
        If Not IsArray(m_vOpers) Then
            ReDim m_vOpers(0 To LNG_NUM_OPERS) As Variant
        End If
        For lIdx = 1 To UBound(m_vOpers)
            If Not IsArray(m_vOpers(lIdx)) Then
                pvStatus = Printf(STR_STATUS_OPER_FETCH, lIdx)
                m_vOpers(lIdx) = Split(m_oFP.SendCommand(ucsFpcInfoOperator, C_Str(lIdx)), ",")
                If m_oFP.Status(ucsStbPrintingError) Or LenB(m_oFP.LastError) <> 0 Then
                    ReDim Preserve m_vOpers(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            sText = lIdx & ": " & At(m_vOpers(lIdx), 5)
            If lstOpers.ListCount < lIdx Then
                lstOpers.AddItem sText
            ElseIf lstOpers.List(lIdx - 1) <> sText Then
                lstOpers.List(lIdx - 1) = sText
            End If
        Next
        lstOpers_Click
        pvStatus = vbNullString
        m_oFP.Exceptions = True
    Case ucsCmdDepartments
        m_oFP.Exceptions = False
        If Not IsArray(m_vDeps) Then
            ReDim m_vDeps(0 To LNG_NUM_DEPS) As Variant
        End If
        For lIdx = 1 To UBound(m_vDeps)
            If Not IsArray(m_vDeps(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_DEP, lIdx)
                m_vDeps(lIdx) = Split(m_oFP.SendCommand(ucsFpcInfoDepartment, C_Str(lIdx)), ",")
                If m_oFP.Status(ucsStbPrintingError) Or LenB(m_oFP.LastError) <> 0 Then
                    ReDim Preserve m_vDeps(0 To lIdx - 1) As Variant
                    Exit For
                End If
            End If
            If lstDeps.ListCount < lIdx Then
                lstDeps.AddItem vbNullString
            End If
            vResult = m_vDeps(lIdx)
            If Left$(At(vResult, 0), 1) = "P" Then
                sText = At(Split(At(vResult, 4), vbLf), 0) & " (" & Mid$(At(vResult, 0), 2) & ")"
            Else
                sText = lIdx & ": " & STR_NA
            End If
            If lstDeps.List(lIdx - 1) <> sText Then
                lstDeps.List(lIdx - 1) = sText
            End If
        Next
        lstDeps_Click
        pvStatus = vbNullString
        m_oFP.Exceptions = True
    Case ucsCmdItems
        If Not IsArray(m_vItems) Then
            lstItems.Clear
            ReDim m_vItems(0 To 0) As Variant
            lIdx = 0
            vResult = Split(m_oFP.SendCommand(ucsFpcInitItem, "F"), ",")
            Do While vResult(0) <> "F"
                lIdx = lIdx + 1
                ReDim Preserve m_vItems(0 To lIdx) As Variant
                m_vItems(lIdx) = vResult
                vResult = Split(m_oFP.SendCommand(ucsFpcInitItem, "N"), ",")
            Loop
        End If
        For lIdx = 1 To UBound(m_vItems)
            If Not IsArray(m_vItems(lIdx)) Then
                m_vItems(lIdx) = Split(m_oFP.SendCommand(ucsFpcInitItem, "R" & m_vItems(lIdx)), ",")
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
        m_oFP.Exceptions = False
        vResult = Split(m_oFP.SendCommand(ucsFpcInitHeaderFooter, "IL"), ",")
        chkLogoPrint.Value = IIf(At(vResult, UBound(vResult)) = "1", vbChecked, vbUnchecked)
        LockControl(chkLogoPrint) = m_oFP.Status(ucsStbPrintingError)
        If Not IsArray(m_vLogo) Then
            If m_oFP.IsDaisy Or m_oFP.IsIncotex Then
                vResult = Split(m_oFP.SendCommand(ucsFpcExtendedInfoConsts), ",")
                lWidth = C_Lng(At(vResult, 0, 64))  '--- P1      Horizonatal size of graphical logo in pixels.
                lRow = C_Lng(At(vResult, 1, 64))    '--- P2      Vertical size of graphical logo in pixels.
            Else
                lRow = 1000
            End If
            If lRow > 0 Then
                ReDim m_vLogo(0 To lRow - 1) As Variant
            Else
                m_vLogo = Array()
            End If
            For lRow = 0 To UBound(m_vLogo)
                pvStatus = Printf(STR_STATUS_FETCH_LOGO, lRow + 1)
                m_vLogo(lRow) = m_oFP.SendCommand(ucsFpcInitLogo, "R" & lRow)
                If m_oFP.Status(ucsStbPrintingError) Or LenB(m_oFP.LastError) <> 0 Or LenB(m_vLogo(lRow)) = 0 Then
                    If lRow > 0 Then
                        ReDim Preserve m_vLogo(0 To lRow - 1) As Variant
                    Else
                        If m_oFP.IsDaisy Then
                            For lIdx = 0 To UBound(m_vLogo)
                                m_vLogo(lIdx) = String$(lWidth / 4, "0")
                            Next
                        Else
                            m_vLogo = Array()
                        End If
                    End If
                    Exit For
                End If
                '--- note: bug in firmware byte to hex routine: 0xA - 1 = "@" instead of "9"
                m_vLogo(lRow) = Replace(m_vLogo(lRow), "@", "9")
            Next
            If UBound(m_vLogo) >= 0 Then
                picLogo.Width = Len(m_vLogo(0)) * 4 * Screen.TwipsPerPixelX
            End If
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
        m_oFP.Exceptions = True
    Case ucsCmdCashOper
        m_oFP.Exceptions = False
        vResult = Split(m_oFP.SendCommand(ucsFpcFiscalServiceDeposit), ",")
        txtCashTotal.Text = Format$(C_Dbl(At(vResult, 1)) / 100, FORMAT_CURRENCY)
        txtCashIn.Text = Format$(C_Dbl(At(vResult, 2)) / 100, FORMAT_CURRENCY)
        txtCashOut.Text = Format$(C_Dbl(At(vResult, 3)) / 100, FORMAT_CURRENCY)
        LockControl(txtCashSum) = m_oFP.Status(ucsStbPrintingError)
        m_oFP.Exceptions = True
    Case ucsCmdReports
        '--- do nothing
    Case ucsCmdStatus
        If lstStatus.ListCount = 0 Then
            For Each vElem In Split(STR_FP_STATUSES, "|")
                lstStatus.AddItem vElem
            Next
        End If
        lstStatus.Tag = vbNullString
        For lIdx = 0 To lstStatus.ListCount - 1
            lRow = (lIdx \ 8) * 8 + (7 - (lIdx Mod 8))
            If lRow < 24 Then
                lstStatus.Selected(lIdx) = m_oFP.Status(2 ^ lRow)
            ElseIf lRow < 32 Then
                lstStatus.Selected(lIdx) = m_oFP.Dip(2 ^ (lRow - 24))
            Else
                lstStatus.Selected(lIdx) = m_oFP.Memory(2 ^ (lRow - 32))
            End If
        Next
        lstStatus.ListIndex = 0
        lstStatus.Tag = "Locked"
    Case ucsCmdDiagnostics
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoDiagnostics, "1"), ",")
        txtDiagFirmware.Text = At(vResult, 0)
        txtDiagChecksum.Text = At(vResult, 1)
        txtDiagSwitches.Text = At(vResult, 2)
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
        pvStatus = STR_STATUS_CONNECTING
        If m_oFP.Init("Port=" & cobConnectPort.Text & ";Speed=" & C_Lng(cobConnectSpeed.Text) & ";Timeout=" & m_lTimeout) Then
            On Error Resume Next '--- checked
            m_oFP.SendCommand ucsFpcInfoTransaction
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
                m_vLogo = Empty
                m_sLog = vbNullString
                lstCmds.ListIndex = ucsCmdTaxInfo
            End If
        Else
            labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
            Caption = CAP_MSG
        End If
    Case ucsCmdTaxInfo
        vResult = Split(m_oFP.SendCommand(ucsFpcInitDecimals), ",")
        m_oFP.SendCommand ucsFpcInitDecimals, At(vResult, 0) & "," & C_Lng(txtTaxDecimals.Text) & "," & txtTaxCurrency.Text & " ," & C_Lng(txtTaxRates.Text)
        vResult = Split(m_oFP.SendCommand(ucsFpcInfoTaxRates), ",")
        vResult(0) = C_Lng(txtTaxGroup1.Text)
        vResult(1) = C_Lng(txtTaxGroup2.Text)
        vResult(2) = C_Lng(txtTaxGroup3.Text)
        vResult(3) = C_Lng(txtTaxGroup4.Text)
        m_oFP.SendCommand ucsFpcInitTaxRates, Join(vResult, ",")
    Case ucsCmdDateTime
        m_oFP.SendCommand ucsFpcInitDateTime, txtDateDate.Text & " " & txtDateTime.Text
    Case ucsCmdHeaderFooter
        m_oFP.Exceptions = False
        If Not LockControl(txtHeadHeader1) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "0" & RTrim(txtHeadHeader1.Text)
        End If
        If Not LockControl(txtHeadHeader2) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "1" & RTrim(txtHeadHeader2.Text)
        End If
        If Not LockControl(txtHeadBulstatText) Then
            m_oFP.SendCommand ucsFpcInitBulstat, RTrim(txtHeadBulstatText.Text) & "," & RTrim(txtHeadBulstatName.Text)
        End If
        If Not LockControl(txtHeadHeader3) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "2" & RTrim(txtHeadHeader3.Text)
        End If
        If Not LockControl(txtHeadHeader4) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "3" & RTrim(txtHeadHeader4.Text)
        End If
        If Not LockControl(txtHeadHeader5) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "4" & RTrim(txtHeadHeader5.Text)
        End If
        If Not LockControl(txtHeadHeader6) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "5" & RTrim(txtHeadHeader6.Text)
        End If
        If Not LockControl(txtHeadFooter1) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "6" & RTrim(txtHeadFooter1.Text)
        End If
        If Not LockControl(txtHeadFooter2) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "7" & RTrim(txtHeadFooter2.Text)
        End If
        If Not LockControl(chkHeadFormatInvoice) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "A" & chkHeadFormatInvoice.Value
        End If
        If Not LockControl(chkHeadSumEUR) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "E" & chkHeadSumEUR.Value & IIf(chkHeadRateEUR.Value, "," & txtHeadRate.Text, vbNullString)
        End If
        If Not LockControl(chkHeadAdvanceHeader) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "H" & chkHeadAdvanceHeader.Value
        End If
        If Not LockControl(chkHeadEmptyHeader) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "P" & chkHeadEmptyHeader.Value & "0" & chkHeadEmptyFooter.Value & chkHeadSumDivider.Value
        End If
        If Not LockControl(chkHeadVat) Then
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "T" & chkHeadVat.Value
        End If
        m_oFP.Exceptions = True
    Case ucsCmdDepartments
        If LenB(txtDepNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_DEP_SELECTED
            Exit Function
        End If
        m_oFP.SendCommand ucsFpcInitDepartment, txtDepNo.Text & "," & cobDepGroup.Text & "," & txtDepName.Text & IIf(LenB(txtDepName2.Text) <> 0, vbLf & txtDepName2.Text, vbNullString)
        '--- force refetch of department info
        m_vDeps(C_Lng(txtDepNo.Text)) = Empty
    Case ucsCmdOperators
        If LenB(txtOperNo.Text) = 0 Then
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
        If m_oFP.IsEcr Then
            vResult = Split(m_oFP.SendCommand(ucsFpcEcrReadRow, "2;" & txtOperNo.Text & ";"), ";")
            ValueAt(vResult, 2) = txtOperName.Text
            If LenB(txtOperPass.Text) <> 0 Then
                ValueAt(vResult, 3) = txtOperPass.Text
            End If
            vResult = m_oFP.SendCommand(ucsFpcEcrWriteRow, Join(vResult, ";"))
        Else
            sPass = InputBox(Printf(STR_OPER_PASS_PROMPT, txtOperNo.Text), STR_OPER_PASS_CAPTION, m_oFP.GetDefaultPassword(txtOperNo.Text))
            If StrPtr(sPass) = 0 Then
                Exit Function
            ElseIf Not pvIsPassCorrect(sPass) Then
                MsgBox MSG_INVALID_PASSWORD, vbExclamation
                Exit Function
            End If
            bCheckPass = True
            m_oFP.SendCommand ucsFpcInitOperatorName, txtOperNo.Text & "," & sPass & "," & txtOperName.Text
            If LenB(txtOperPass.Text) <> 0 Then
                m_oFP.SendCommand ucsFpcInitOperatorPassword, txtOperNo.Text & "," & sPass & "," & txtOperPass.Text
            End If
            bCheckPass = False
        End If
        '--- force refetch of oper info
        m_vOpers(C_Lng(txtOperNo.Text)) = Empty
    Case ucsCmdInvoiceNo
        If Not LockControl(txtInvStart) Then
            If m_oFP.IsIncotex Then
                m_oFP.SendCommand ucsFpcExtendedInitInvoiceNo, txtInvStart.Text & vbLf & txtInvEnd.Text
            Else
                m_oFP.SendCommand ucsFpcInitInvoiceNo, txtInvStart.Text & "," & txtInvEnd.Text
            End If
        End If
        If m_oFP.IsDaisy And LenB(txtSettNo.Text) <> 0 Then
            m_oFP.SendCommand ucsFpcExtendedInitSetting, "P" & txtSettNo.Text & "," & txtSettValue
            m_vSettings(C_Lng(txtSettNo.Text)) = Empty
        End If
    Case ucsCmdCashOper
        If Not LockControl(txtCashSum) And C_Dbl(txtCashSum.Text) <> 0 Then
            vResult = Split(m_oFP.SendCommand(ucsFpcFiscalServiceDeposit, IIf(optCashOut.Value, -1, 1) * Abs(C_Dbl(txtCashSum.Text))), ",")
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
                eCmd = ucsFpcPrintDailyReportItemsDepartments
            ElseIf chkReportItems.Value = vbChecked Then
                eCmd = ucsFpcPrintDailyReportItems
            ElseIf chkReportDepartments.Value = vbChecked Then
                eCmd = ucsFpcPrintDailyReportDepartments
            Else
                eCmd = ucsFpcPrintDailyReport
            End If
            m_oFP.Exceptions = False
            '--- "rychno" razpechatwane na elektronna kontrolna lenta
            If chkReportClosure.Value = vbChecked Then
                vResult = Split(m_oFP.SendCommand(ucsFpcInitEcTape, "I"), ",")
                '--- print
                For lIdx = 1 To C_Lng(At(vResult, 1))
                    m_oFP.SendCommand ucsFpcInitEcTape, IIf(lIdx = 1, "PS", "CS")
                    If lIdx = C_Lng(At(vResult, 1)) Then
                        '--- erase
                        m_oFP.SendCommand ucsFpcInitEcTape, "E"
                    End If
                Next
            End If
            vResult = m_oFP.SendCommand(eCmd, IIf(chkReportClosure.Value = vbChecked, "0", "2") & "N")
            If m_oFP.Status(ucsStbPrintingError) Then
                '--- daisy: pechat po depatamenti
                If eCmd = ucsFpcPrintDailyReportDepartments Then
                    vResult = m_oFP.SendCommand(ucsFpcPrintDailyReport, IIf(chkReportClosure.Value = vbChecked, "8", "9") & "N")
                ElseIf eCmd = ucsFpcPrintDailyReportItemsDepartments Then
                    vResult = m_oFP.SendCommand(ucsFpcPrintDailyReportItems, IIf(chkReportClosure.Value = vbChecked, "8", "9") & "N")
                End If
            End If
            m_oFP.Exceptions = True
        ElseIf optReportType(2).Value Then '--- by number
            If chkReportDetailed1.Value Then
                eCmd = ucsFpcPrintReportByNumberDetailed
            Else
                eCmd = ucsFpcPrintReportByNumberShort
            End If
            vResult = m_oFP.SendCommand(eCmd, txtReportStart.Text & "," & txtReportEnd.Text)
        ElseIf optReportType(3).Value Then '--- by date
            If chkReportDetailed2.Value Then
                eCmd = ucsFpcPrintReportByDateDetailed
            Else
                eCmd = ucsFpcPrintReportByDateShort
            End If
            vResult = m_oFP.SendCommand(eCmd, txtReportFD.Text & IIf(LenB(txtReportTD.Text) <> 0, "," & txtReportTD.Text, vbNullString))
        ElseIf optReportType(5).Value Then '--- by operator
            vResult = m_oFP.SendCommand(ucsFpcPrintReportByOperators)
        End If
        pvStatus = vbNullString
    Case ucsCmdStatus
        pvStatus = STR_STATUS_REFRESH
        vResult = m_oFP.SendCommand(ucsFpcInfoStatus, "W")
        pvStatus = vbNullString
    Case ucsCmdDiagnostics
        pvStatus = STR_STATUS_PRINT
        vResult = m_oFP.SendCommand(ucsFpcPrintDiagnostics)
        pvStatus = vbNullString
    Case ucsCmdPaymentTypes
        If m_oFP.IsDaisy Then
            For lIdx = 0 To 7
                If Not LockControl(txtPmtType(lIdx)) And LenB(txtPmtType(lIdx).Text) <> 0 Then
                    m_oFP.SendCommand ucsFpcExtendedInitText, "P" & (60 + lIdx) & "," & txtPmtType(lIdx).Text
                    m_oFP.SendCommand ucsFpcExtendedInitCurrencyRate, "P" & (lIdx) & "," & txtPmtType(lIdx).Text & vbTab & txtPmtRate(lIdx)
                End If
            Next
        ElseIf m_oFP.IsIncotex Then
            For lIdx = 4 To 7
                If Not LockControl(txtPmtType(lIdx)) And LenB(txtPmtType(lIdx).Text) <> 0 Then
                    m_oFP.SendCommand ucsFpcExtendedInitText, "P" & (8 + lIdx) & "," & txtPmtType(lIdx).Text
                End If
            Next
        Else
            m_oFP.Exceptions = False
            For lIdx = 4 To 7
                If Not LockControl(txtPmtType(lIdx)) Then
                    m_oFP.SendCommand ucsFpcInitPaymentType, Chr$(69 + lIdx) & "," & txtPmtType(lIdx).Text
                End If
            Next
            m_oFP.Exceptions = True
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
            ReDim Preserve m_vItems(0 To lIdx) As Variant
        Else
            lIdx = lstItems.ListIndex + 1
        End If
        vResult = Split(m_oFP.SendCommand(ucsFpcInitItem, "P" & cobItemGroup.Text & txtItemPLU.Text & "," & txtItemPrice.Text & "," & txtItemName.Text), ",")
        If At(vResult, 0) = "F" Then
            pvStatus = Printf(STR_STATUS_ITEM_FAILURE_ADD, At(vResult, 1))
            Exit Function
        End If
        m_vItems(lIdx) = txtItemPLU.Text
    Case ucsCmdGraphicalLogo
        If Not LockControl(chkLogoPrint) Then
            m_oFP.Exceptions = False
            m_oFP.SendCommand ucsFpcInitHeaderFooter, "L" & chkLogoPrint.Value
            If m_oFP.Status(ucsStbPrintingError) Then
                m_oFP.SendCommand ucsFpcInitHeaderFooter, "L" & (UBound(m_vLogo) + 1) & "," & chkLogoPrint.Value
                If m_oFP.Status(ucsStbPrintingError) Then
                    If m_oFP.IsEcr Then
                        vResult = Split(m_oFP.SendCommand(ucsFpcEcrReadRow, "5;1;"), ";")
                        ValueAt(vResult, 6) = chkLogoPrint.Value
                        vResult = m_oFP.SendCommand(ucsFpcEcrWriteRow, Join(vResult, ";"))
                    End If
                End If
            End If
            m_oFP.Exceptions = False
        End If
        If Not m_picLogo Is Nothing Then
            For lIdx = 0 To UBound(m_vLogo)
                pvStatus = Printf(STR_STATUS_SAVE_LOGO, lIdx + 1, UBound(m_vLogo) + 1)
                m_oFP.SendCommand ucsFpcInitLogo, lIdx & "," & m_vLogo(lIdx)
            Next
        End If
    End Select
    '--- success
    pvSaveData = True
    Exit Function
EH:
    If bCheckPass Then
        If m_oFP.Status(ucsStbInvalidFiscalMode) Then
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

Private Function pvShowError() As Boolean
    If Len(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
        pvStatus = m_oFP.LastError
        pvShowError = True
    End If
    If m_oFP.Status(ucsStbPrintingError) Then
        MsgBox m_oFP.ErrorText, vbExclamation
        pvStatus = m_oFP.ErrorText
        pvShowError = True
    End If
End Function

Private Function pvIsPassCorrect(sPass As String) As Boolean
    Dim lIdx            As Long
    Dim lChar           As Long
    
    If Len(sPass) >= 1 And Len(sPass) <= 6 Then
        For lIdx = 1 To Len(sPass)
            lChar = Asc(Mid$(sPass, lIdx, 1))
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
    dblTimer = TimerEx
    If lstCmds.ListIndex = ucsCmdSettings Or lstCmds.ListIndex = ucsCmdOperations Or lstCmds.ListIndex = ucsCmdAdmin Then
        lVisibleFrame = -1
        GoTo QH
    End If
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect And lstCmds.ListIndex <> ucsCmdStatus Then
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
    cobDepGroup.Text = Mid$(At(vResult, 0), 2)
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

Private Sub lstSettings_Click()
    Const FUNC_NAME     As String = "lstSettings_Click"
    Dim vResult         As Variant
    
    On Error GoTo EH
    If lstSettings.ListIndex >= 0 Then
        txtSettNo.Text = lstSettings.ListIndex + 1
        vResult = m_vSettings(lstSettings.ListIndex + 1)
    Else
        txtOperNo.Text = vbNullString
    End If
    txtSettValue.Text = At(vResult, 1)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
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
    txtDateCompDate.Text = Format$(Now, "dd\-MM\-yy")
    txtDateCompTime.Text = Format$(Now, "hh\:nn\:ss")
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
    Dim sFile           As String
    
    On Error GoTo EH
    If OpenSaveDialog(Me.hWnd, STR_FILTER, STR_TITLE, sFile) Then
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
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_OPER_RESETTING
    m_oFP.SendCommand ucsFpcInitOperatorReset, txtOperNo.Text & "," & sPass
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
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_ITEM_DELETING
    m_oFP.SendCommand ucsFpcInitItem, "D" & sPLU
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
        On Error Resume Next '--- checked
        m_oFP.Connect
        On Error GoTo EH
        If pvShowError() Then
            GoTo QH
        End If
    End If
    pvStatus = STR_STATUS_RESETTING
    m_oFP.Exceptions = False
'    If Left$(m_oFP.SendCommand(ucsFpcInfoTransaction), 1) = "1" Then
'        If m_oFP.Status(ucsStbFiscalPrinting) Then
            '--- note: when printing invoice, if no contragent info set then cancel fails!
            m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000" & vbTab & "0"
            '--- note: FP3530 moje da anulira winagi, FP550F ne moje
            m_oFP.SendCommand ucsFpcFiscalCancel
            '--- zaradi FP550F
            m_oFP.SendCommand ucsFpcFiscalClose
'        Else
            m_oFP.SendCommand ucsFpcNonFiscalClose
'        End If
'    End If
    m_oFP.Exceptions = True
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

Private Sub m_oFPSink_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFPSink_CommandComplete"
    
    On Error GoTo EH
    m_sLog = m_sLog & lCmd & IIf(LenB(sData) <> 0, "<-" & sData, vbNullString) & IIf(LenB(sResult) <> 0, "->" & sResult, vbNullString) & vbCrLf
    If LenB(m_oFP.LastError) <> 0 Then
        m_sLog = m_sLog & m_oFP.LastError & vbCrLf
    End If
    If m_oFP.Status(ucsStbPrintingError) Then
        m_sLog = m_sLog & m_oFP.StatusText & vbCrLf & m_oFP.DipText & vbCrLf & m_oFP.MemoryText & vbCrLf
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

