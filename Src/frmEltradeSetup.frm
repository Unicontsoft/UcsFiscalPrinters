VERSION 5.00
Begin VB.Form frmEltradeSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки ELTRADE протокол"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEltradeSetup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   14
      Left            =   2250
      TabIndex        =   197
      Top             =   90
      Width           =   5775
      Begin VB.CheckBox chkReportDepartments 
         Caption         =   "Департаменти"
         Height          =   285
         Left            =   3960
         TabIndex        =   90
         Top             =   630
         Width           =   1725
      End
      Begin VB.CheckBox chkReportItems 
         Caption         =   "Артикули"
         Height          =   285
         Left            =   2430
         TabIndex        =   89
         Top             =   630
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Дневен финансов отчет"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   87
         Top             =   270
         Value           =   -1  'True
         Width           =   5145
      End
      Begin VB.CheckBox chkReportClosure 
         Caption         =   "Нулиране"
         Height          =   285
         Left            =   900
         TabIndex        =   88
         Top             =   630
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Данъчни ставки"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   101
         Top             =   4140
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Периодичен отчет по номер на запис"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   91
         Top             =   990
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Периодичен отчет по дата на запис"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   95
         Top             =   2160
         Width           =   5145
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Натрупани суми за период"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   100
         Top             =   3780
         Width           =   5145
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Печат"
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   102
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtReportFD 
         Height          =   285
         Left            =   1800
         TabIndex        =   96
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtReportTD 
         Height          =   285
         Left            =   3420
         TabIndex        =   97
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox chkReportDetailed2 
         Caption         =   "Детайлен"
         Height          =   285
         Left            =   900
         TabIndex        =   98
         Top             =   2970
         Width           =   1725
      End
      Begin VB.TextBox txtReportEnd 
         Height          =   285
         Left            =   3420
         TabIndex        =   93
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txtReportStart 
         Height          =   285
         Left            =   1800
         TabIndex        =   92
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CheckBox chkReportDetailed1 
         Caption         =   "Детайлен"
         Height          =   285
         Left            =   900
         TabIndex        =   94
         Top             =   1710
         Width           =   1725
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Отчет оператори"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   99
         Top             =   3420
         Width           =   5145
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "От дата:"
         Height          =   300
         Left            =   900
         TabIndex        =   201
         Top             =   2520
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "до: "
         Height          =   195
         Left            =   2520
         TabIndex        =   200
         Top             =   2520
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "до: "
         Height          =   195
         Left            =   2520
         TabIndex        =   199
         Top             =   1350
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "От номер:"
         Height          =   195
         Left            =   900
         TabIndex        =   198
         Top             =   1350
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   13
      Left            =   2250
      TabIndex        =   194
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtCashTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton optCashOut 
         Caption         =   "Износ"
         Height          =   285
         Left            =   3150
         TabIndex        =   84
         Top             =   1620
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Внос/износ"
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   86
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtCashSum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   85
         Top             =   1980
         Width           =   1545
      End
      Begin VB.OptionButton optCashIn 
         Caption         =   "Внос"
         Height          =   285
         Left            =   2070
         TabIndex        =   83
         Top             =   1620
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Наличност каса:"
         Height          =   195
         Left            =   180
         TabIndex        =   196
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сума:"
         Height          =   195
         Left            =   180
         TabIndex        =   195
         Top             =   1980
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   16
      Left            =   2250
      TabIndex        =   202
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   2430
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   2790
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   3150
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   1710
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2070
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. на маса:"
         Height          =   195
         Left            =   180
         TabIndex        =   213
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сер. No. на сметка:"
         Height          =   195
         Left            =   180
         TabIndex        =   212
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Фактура:"
         Height          =   195
         Left            =   180
         TabIndex        =   211
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Позиция на ключ:"
         Height          =   195
         Left            =   180
         TabIndex        =   210
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. на оператор:"
         Height          =   195
         Left            =   180
         TabIndex        =   209
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тип бележка:"
         Height          =   195
         Left            =   180
         TabIndex        =   208
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. на разписка:"
         Height          =   195
         Left            =   180
         TabIndex        =   205
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Последна команда:"
         Height          =   195
         Left            =   180
         TabIndex        =   204
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сума:"
         Height          =   195
         Left            =   180
         TabIndex        =   203
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   10
      Left            =   2250
      TabIndex        =   206
      Top             =   90
      Width           =   5775
      Begin VB.ListBox lstYesNoParams 
         Height          =   4335
         ItemData        =   "frmEltradeSetup.frx":000C
         Left            =   180
         List            =   "frmEltradeSetup.frx":0040
         Style           =   1  'Checkbox
         TabIndex        =   215
         Top             =   270
         Width           =   5415
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   112
         Top             =   5220
         Width           =   1275
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   11
      Left            =   2250
      TabIndex        =   216
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtKeysWithNumber 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   219
         Top             =   630
         Width           =   4335
      End
      Begin VB.TextBox txtKeysNoNumber 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   218
         Top             =   270
         Width           =   4335
      End
      Begin VB.CommandButton cmdKeysReset 
         Caption         =   "Ресет"
         Height          =   375
         Left            =   180
         TabIndex        =   217
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Със номер:"
         Height          =   195
         Left            =   180
         TabIndex        =   221
         Top             =   630
         Width           =   1185
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Без номер:"
         Height          =   195
         Left            =   180
         TabIndex        =   220
         Top             =   270
         Width           =   1185
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   17
      Left            =   2250
      TabIndex        =   207
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtLog 
         Height          =   5505
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   113
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   7
      Left            =   2250
      TabIndex        =   167
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtOperResto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1545
      End
      Begin VB.TextBox txtOperDisc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1545
      End
      Begin VB.TextBox txtOperSells 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1545
      End
      Begin VB.TextBox txtOperFiscal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtOperVoid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1545
      End
      Begin VB.TextBox txtOperSurcharge 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2430
         Width           =   1545
      End
      Begin VB.TextBox txtOperNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.ListBox lstOpers 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   53
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   63
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtOperName 
         Height          =   285
         Left            =   2610
         MaxLength       =   12
         TabIndex        =   55
         Top             =   900
         Width           =   2985
      End
      Begin VB.TextBox txtOperPass 
         Height          =   285
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   61
         Top             =   4050
         Width           =   1545
      End
      Begin VB.TextBox txtOperPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4050
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   62
         Top             =   4410
         Width           =   1545
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Върнати:"
         Height          =   195
         Left            =   2610
         TabIndex        =   177
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Отстъпки:"
         Height          =   195
         Left            =   2610
         TabIndex        =   176
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Продажби:"
         Height          =   195
         Left            =   2610
         TabIndex        =   175
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Фискални бонове:"
         Height          =   195
         Left            =   2610
         TabIndex        =   174
         Top             =   1350
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Корекции:"
         Height          =   195
         Left            =   2610
         TabIndex        =   173
         Top             =   2790
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Надбавки:"
         Height          =   195
         Left            =   2610
         TabIndex        =   172
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Номер:"
         Height          =   195
         Left            =   2610
         TabIndex        =   171
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Наименование:"
         Height          =   195
         Left            =   2610
         TabIndex        =   170
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Нова парола:"
         Height          =   195
         Left            =   2610
         TabIndex        =   169
         Top             =   4050
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Пак парола:"
         Height          =   195
         Left            =   2610
         TabIndex        =   168
         Top             =   4410
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   8
      Left            =   2250
      TabIndex        =   178
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtDepName 
         Height          =   285
         Left            =   2610
         MaxLength       =   31
         TabIndex        =   67
         Top             =   1260
         Width           =   2985
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   71
         Top             =   5220
         Width           =   1275
      End
      Begin VB.ListBox lstDeps 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   64
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtDepNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtDepSales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1545
      End
      Begin VB.TextBox txtDepItemGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   40
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1710
         Width           =   825
      End
      Begin VB.TextBox txtDepTotalSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1545
      End
      Begin VB.ComboBox cobDepGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   66
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Наименование:"
         Height          =   195
         Left            =   2610
         TabIndex        =   184
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Номер:"
         Height          =   195
         Left            =   2610
         TabIndex        =   183
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Данъчна група:"
         Height          =   195
         Left            =   2610
         TabIndex        =   182
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Количество:"
         Height          =   195
         Left            =   2610
         TabIndex        =   181
         Top             =   2340
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група:"
         Height          =   195
         Left            =   2610
         TabIndex        =   180
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Оборот:"
         Height          =   195
         Left            =   2610
         TabIndex        =   179
         Top             =   2700
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   9
      Left            =   2250
      TabIndex        =   185
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtItemSoldQuo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1545
      End
      Begin VB.ListBox lstItemFlags 
         Height          =   1365
         IntegralHeight  =   0   'False
         ItemData        =   "frmEltradeSetup.frx":01FC
         Left            =   2610
         List            =   "frmEltradeSetup.frx":020F
         Style           =   1  'Checkbox
         TabIndex        =   222
         Top             =   2700
         Width           =   2985
      End
      Begin VB.TextBox txtItemDep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   9
         TabIndex        =   78
         Top             =   2340
         Width           =   825
      End
      Begin VB.TextBox txtItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   270
         Width           =   825
      End
      Begin VB.ListBox lstItems 
         Height          =   5325
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   72
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   13
         Left            =   4320
         TabIndex        =   81
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   2610
         MaxLength       =   25
         TabIndex        =   76
         Top             =   1620
         Width           =   2985
      End
      Begin VB.ComboBox cobItemGroup 
         Height          =   315
         Left            =   4050
         TabIndex        =   75
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtItemAvailable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   4140
         Width           =   1545
      End
      Begin VB.TextBox txtItemGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   77
         Top             =   1980
         Width           =   825
      End
      Begin VB.TextBox txtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   9
         TabIndex        =   74
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtItemTurnover 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4500
         Width           =   1545
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Продадено кол."
         Height          =   195
         Left            =   2610
         TabIndex        =   224
         Top             =   4860
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Номер:"
         Height          =   195
         Left            =   2610
         TabIndex        =   193
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Наименование:"
         Height          =   195
         Left            =   2610
         TabIndex        =   192
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Налични:"
         Height          =   195
         Left            =   2610
         TabIndex        =   191
         Top             =   4140
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Данъчна група:"
         Height          =   195
         Left            =   2610
         TabIndex        =   190
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група:"
         Height          =   195
         Left            =   2610
         TabIndex        =   189
         Top             =   1980
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Цена:"
         Height          =   195
         Left            =   2610
         TabIndex        =   188
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Оборот:"
         Height          =   195
         Left            =   2610
         TabIndex        =   187
         Top             =   4500
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Подгрупа:"
         Height          =   195
         Left            =   2610
         TabIndex        =   186
         Top             =   2340
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   6
      Left            =   2250
      TabIndex        =   155
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtPmtRate 
         Height          =   285
         Index           =   3
         Left            =   4770
         MaxLength       =   40
         TabIndex        =   51
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   285
         Index           =   2
         Left            =   4770
         MaxLength       =   40
         TabIndex        =   49
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   285
         Index           =   1
         Left            =   4770
         MaxLength       =   40
         TabIndex        =   47
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtPmtRate 
         Height          =   285
         Index           =   0
         Left            =   4770
         MaxLength       =   40
         TabIndex        =   45
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   0
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   44
         Top             =   270
         Width           =   1995
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   52
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   1
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   46
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   2
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   48
         Top             =   990
         Width           =   1995
      End
      Begin VB.TextBox txtPmtType 
         Height          =   285
         Index           =   3
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   50
         Top             =   1350
         Width           =   1995
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Курс:"
         Height          =   195
         Left            =   4050
         TabIndex        =   163
         Top             =   1350
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Курс:"
         Height          =   195
         Left            =   4050
         TabIndex        =   162
         Top             =   990
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Курс:"
         Height          =   195
         Left            =   4050
         TabIndex        =   161
         Top             =   630
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Курс:"
         Height          =   195
         Left            =   4050
         TabIndex        =   160
         Top             =   270
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тип плащане 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   159
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тип плащане 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   158
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тип плащане 3:"
         Height          =   195
         Left            =   180
         TabIndex        =   157
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тип плащане 4:"
         Height          =   195
         Left            =   180
         TabIndex        =   156
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   5
      Left            =   2250
      TabIndex        =   164
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtInvStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   41
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtInvEnd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2070
         TabIndex        =   42
         Top             =   630
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   43
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Начален номер:"
         Height          =   195
         Left            =   180
         TabIndex        =   166
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Брой номера:"
         Height          =   195
         Left            =   180
         TabIndex        =   165
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   4
      Left            =   2250
      TabIndex        =   146
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   8
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   29
         Top             =   1350
         Width           =   1725
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   1
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   26
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   0
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   24
         Top             =   270
         Width           =   3525
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   38
         ToolTipText     =   "Получер (bold)"
         Top             =   3510
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   36
         ToolTipText     =   "Получер (bold)"
         Top             =   3150
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   34
         ToolTipText     =   "Получер (bold)"
         Top             =   2430
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   32
         ToolTipText     =   "Получер (bold)"
         Top             =   2070
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   30
         ToolTipText     =   "Получер (bold)"
         Top             =   1710
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   25
         ToolTipText     =   "Получер (bold)"
         Top             =   630
         Width           =   555
      End
      Begin VB.CheckBox chkHeadHeader 
         Caption         =   "ПЧ"
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   23
         ToolTipText     =   "Получер (bold)"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   40
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   2
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   31
         Top             =   1710
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   3
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   33
         Top             =   2070
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   4
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   35
         Top             =   2430
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   5
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   37
         Top             =   3150
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   6
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   39
         Top             =   3510
         Width           =   3525
      End
      Begin VB.TextBox txtHeadHeader 
         Height          =   285
         Index           =   7
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   27
         Top             =   990
         Width           =   1725
      End
      Begin VB.TextBox txtHeadBulstatText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3870
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   28
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Идент. No."
         Height          =   195
         Left            =   180
         TabIndex        =   214
         Top             =   1350
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   154
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   153
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 3:"
         Height          =   195
         Left            =   180
         TabIndex        =   152
         Top             =   1710
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 4:"
         Height          =   195
         Left            =   180
         TabIndex        =   151
         Top             =   2070
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header 5:"
         Height          =   195
         Left            =   180
         TabIndex        =   150
         Top             =   2430
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer 1:"
         Height          =   195
         Left            =   180
         TabIndex        =   149
         Top             =   3150
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer 2:"
         Height          =   195
         Left            =   180
         TabIndex        =   148
         Top             =   3510
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "БУЛСТАТ:"
         Height          =   195
         Left            =   180
         TabIndex        =   147
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   3
      Left            =   2250
      TabIndex        =   141
      Top             =   90
      Width           =   5775
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   22
         Top             =   5220
         Width           =   1275
      End
      Begin VB.TextBox txtDateTime 
         Height          =   285
         Left            =   2070
         TabIndex        =   20
         Top             =   1530
         Width           =   1635
      End
      Begin VB.TextBox txtDateDate 
         Height          =   285
         Left            =   2070
         TabIndex        =   19
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtDateCompTime 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox txtDateCompDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
      End
      Begin VB.Timer tmrDate 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4860
         Top             =   540
      End
      Begin VB.CommandButton cmdDateTransfer 
         Caption         =   "От системна"
         Height          =   375
         Left            =   2070
         TabIndex        =   21
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Принтер час:"
         Height          =   195
         Left            =   180
         TabIndex        =   145
         Top             =   1530
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Принтер дата:"
         Height          =   195
         Left            =   180
         TabIndex        =   144
         Top             =   1170
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Системна дата:"
         Height          =   195
         Left            =   180
         TabIndex        =   143
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Системен час:"
         Height          =   195
         Left            =   180
         TabIndex        =   142
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   2
      Left            =   2250
      TabIndex        =   121
      Top             =   90
      Width           =   5775
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2070
         TabIndex        =   15
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2070
         TabIndex        =   14
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2070
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2070
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtTaxMemModule 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtTaxSerNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   990
         Width           =   3525
      End
      Begin VB.TextBox txtTaxModel 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2070
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2070
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtTaxGroup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2070
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Запис"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   16
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   140
         Top             =   3960
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група З:"
         Height          =   195
         Left            =   180
         TabIndex        =   139
         Top             =   3960
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   138
         Top             =   3600
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група Ж:"
         Height          =   195
         Left            =   180
         TabIndex        =   137
         Top             =   3600
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   136
         Top             =   3240
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група Е:"
         Height          =   195
         Left            =   180
         TabIndex        =   135
         Top             =   3240
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   134
         Top             =   2880
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група Д:"
         Height          =   195
         Left            =   180
         TabIndex        =   133
         Top             =   2880
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Фискална памет:"
         Height          =   195
         Left            =   180
         TabIndex        =   132
         Top             =   630
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сериен номер:"
         Height          =   195
         Left            =   180
         TabIndex        =   131
         Top             =   990
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Модел:"
         Height          =   195
         Left            =   180
         TabIndex        =   130
         Top             =   270
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група А:"
         Height          =   195
         Left            =   180
         TabIndex        =   129
         Top             =   1440
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група Б:"
         Height          =   195
         Left            =   180
         TabIndex        =   128
         Top             =   1800
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група В:"
         Height          =   195
         Left            =   180
         TabIndex        =   127
         Top             =   2160
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Група Г:"
         Height          =   195
         Left            =   180
         TabIndex        =   126
         Top             =   2520
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   125
         Top             =   1440
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   124
         Top             =   1800
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   123
         Top             =   2160
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   122
         Top             =   2520
         Width           =   375
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCommands 
      Height          =   5775
      Index           =   0
      Left            =   2250
      TabIndex        =   115
      Top             =   90
      Width           =   5775
      Begin VB.CheckBox chkConnectRemember 
         Caption         =   "Автоматично свързване"
         Height          =   195
         Left            =   1620
         TabIndex        =   2
         Top             =   2070
         Width           =   2985
      End
      Begin VB.ComboBox cobConnectPort 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Top             =   1080
         Width           =   1635
      End
      Begin VB.ComboBox cobConnectSpeed 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Свързване"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   3
         Top             =   5220
         Width           =   1275
      End
      Begin VB.Label labConnectCurrent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   118
         Top             =   270
         Width           =   5325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сериен порт:"
         Height          =   195
         Left            =   180
         TabIndex        =   117
         Top             =   1080
         Width           =   1785
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Скорост:"
         Height          =   195
         Left            =   180
         TabIndex        =   116
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
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   180
      Width           =   1275
   End
   Begin VB.ListBox lstCmds 
      Height          =   5685
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   114
      Top             =   180
      Width           =   2085
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   90
      TabIndex        =   120
      Top             =   5940
      Width           =   7935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEltradeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/frmEltradeSetup.frm 7     4.07.11 15:48 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Nastrojka na ECR po Eltrade protocol
'
' $Log: /UcsFiscalPrinter/Src/frmEltradeSetup.frm $
' 
' 7     4.07.11 15:48 Wqw
' REF: err handling
'
' 6     4.05.11 19:48 Wqw
' REF: fiscal memory report by dates
'
' 5     23.02.11 17:10 Wqw
' REF: po UI
'
' 4     22.02.11 13:53 Wqw
' REF: polzwa cFiscalAdmin za class factory na protocol-a
'
' 3     22.02.11 10:33 Wqw
' REF: show s owner moje da fail-ne
'
' 2     22.02.11 10:06 Wqw
' REF: polzwa string functions
'
' 1     21.02.11 13:42 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmEltradeSetup"

'=========================================================================
' API
'=========================================================================

Private Const EM_SCROLLCARET            As Long = &HB7

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const CAP_MSG               As String = "Настройки ELTRADE протокол"
Private Const LNG_NUM_DEPS          As Long = 10
Private Const LNG_NUM_OPERS         As Long = 30
Private Const LNG_NUM_ITEMS         As Long = 100
Private Const PROGID_PROTOCOL       As String = LIB_NAME & ".cEltradeProtocol"
'--- strings
Private Const STR_COMMANDS          As String = "Връзка принтер|Настройки|    ДДС групи|    Дата и час|    Клишета|    Номера на фактури|    Типове плащания|    Оператори|    Департаменти|    Артикули|    Параметри|    Клавиши|Операции|    Внос и износ|    Печат отчети|Администрация|    Последна операция|    Журнал комуникация"
Private Const STR_SPEEDS            As String = "9600|19200"
Private Const STR_GROUPS            As String = "А|Б|В|Г"
Private Const STR_STATUS_ENUM_PORTS As String = "Изброяване на налични принтери..."
Private Const STR_STATUS_FAILURE_CONNECT As String = "Няма връзка"
Private Const STR_STATUS_CONNECTING As String = "Свързване..."
Private Const STR_STATUS_FETCHING   As String = "Получаване..."
Private Const STR_STATUS_SUCCESS_FETCH As String = "Успешно получаване на %1 (%2 сек.)"
Private Const STR_STATUS_SAVING     As String = "Запазване..."
Private Const STR_STATUS_SUCCESS_SAVE As String = "Успешно запазване на %1 (%2 сек.)"
Private Const STR_STATUS_SUCCESS_CONNECT As String = "Свързан %1"
Private Const STR_STATUS_FETCH_OPER As String = "Получаване опрератор %1 от " & LNG_NUM_OPERS & "..."
Private Const STR_STATUS_NO_OPER_SELECTED As String = "Липсва избран оператор"
Private Const STR_STATUS_FETCH_DEP  As String = "Получаване департамент %1 от " & LNG_NUM_DEPS & "..."
Private Const STR_STATUS_NO_DEP_SELECTED As String = "Липсва избран департамент"
Private Const STR_STATUS_FETCH_ITEM As String = "Получаване артикул %1 от " & LNG_NUM_ITEMS & "..."
Private Const STR_STATUS_NO_ITEM_SELECTED As String = "Липсва избран артикул"
Private Const STR_STATUS_RESETTING  As String = "Ресет..."
Private Const STR_ERROR_VATRATE_SEND As String = "Грешка при установяване на ДДС ставки: 0x%1"
Private Const STR_ERROR_DATETIME_FORMAT As String = "Некоректен формат на дата и час"
Private Const STR_ERROR_DATETIME_SEND As String = "Грешка при установяване на дата и час: 0x%1"
Private Const STR_ERROR_HDRLINES_SEND As String = "Грешка при установяване на заглавни линии: 0x%1"
Private Const MSG_PASSWORDS_MISMATCH As String = "Паролите не съвпадат"
Private Const MSG_CANNOT_ACCESS_PRINTER_PROXY As String = "Грешка при създаване на компонент за достъп до фискален принтер %1." & vbCrLf & vbCrLf & "%2"
Private Const MSG_INVALID_DATE      As String = "Невалидна дата"

Private m_oFP                   As cEltradeProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private WithEvents m_oFPSink    As cEltradeProtocol
Attribute m_oFPSink.VB_VarHelpID = -1
Private m_lRowChars             As Long
Private m_vDeps                 As Variant
Private m_vOpers                As Variant
Private m_vItems                As Variant
Private m_vLogo                 As Variant
Private m_sLog                  As String
Private m_lTimeout              As Long
Private m_lCashDeskNo           As Long

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
        ucsCmdYesNoParams
        ucsCmdKeys
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
    Mid$(m_vLogo(lY), 1 + 2 * (lX \ 8), 2) = Right$("0" & Hex(lValue), 2)
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
    m_lCashDeskNo = C_Lng(At(vSplit, 3))
    vSplit = Split(At(vSplit, 1), ",")
    Set m_oFP = pvGetPrinter(sServer, sError)
    If m_oFP Is Nothing Then
        If LenB(sError) <> 0 Then
            MsgBox Printf(MSG_CANNOT_ACCESS_PRINTER_PROXY, At(vSplit, 0) & IIf(LenB(sServer) <> 0, "@" & sServer, vbNullString), sError), vbExclamation
        End If
        GoTo QH
    End If
    On Error Resume Next
    Set m_oFPSink = m_oFP
    On Error GoTo EH
    '--- init UI
    FixThemeSupport Controls
    For Each vElem In Split(STR_GROUPS, "|")
        cobDepGroup.AddItem vElem
        cobItemGroup.AddItem vElem
    Next    '--- init UI
    For Each vElem In Split(STR_COMMANDS, "|")
        lstCmds.AddItem vElem
    Next
    On Error Resume Next
    For lIdx = fraCommands.LBound To fraCommands.UBound
        fraCommands(lIdx).Visible = False
    Next
    On Error GoTo EH
    cmdExit.Left = -cmdExit.Width
    '--- login
    pvStatus = STR_STATUS_ENUM_PORTS
    cobConnectPort.Clear
    For Each vElem In m_oFP.EnumPorts
        cobConnectPort.AddItem vElem
    Next
    cobConnectPort.Text = At(vSplit, 0) ' GetSetting(CAP_MSG, "Connect", "Port", vbNullString)
    chkConnectRemember.Value = -(LenB(cobConnectPort.Text) <> 0)
    If cobConnectPort.ListCount > 0 And Len(cobConnectPort.Text) = 0 Then
        cobConnectPort.ListIndex = 0
    End If
    cobConnectSpeed.Clear
    For Each vElem In Split(STR_SPEEDS, "|")
        cobConnectSpeed.AddItem vElem
    Next
    cobConnectSpeed.Text = At(vSplit, 1) ' GetSetting(CAP_MSG, "Connect", "Speed", vbNullString)
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
        On Error Resume Next
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

Private Function pvGetPrinter(sServer As String, sError As String) As cEltradeProtocol
    On Error Resume Next
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

Private Function pvShowError() As Boolean
    If Len(m_oFP.LastError) <> 0 Then
        MsgBox m_oFP.LastError, vbExclamation
        pvStatus = m_oFP.LastError
        pvShowError = True
    End If
End Function

Private Function pvFetchData(ByVal eCmd As UcsCommands) As Boolean
    Const FUNC_NAME     As String = "pvFetchData"
    Dim lIdx            As Long
    Dim sResult         As String
    Dim sText           As String
    Dim dblSum          As Double
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCmd <> ucsCmdConnect Then
        pvStatus = labConnectCurrent.Caption
        Exit Function
    End If
    Select Case eCmd
    Case ucsCmdConnect
        pvStatus = labConnectCurrent.Caption
    Case ucsCmdTaxInfo
        sResult = m_oFP.SendCommand(ucsEltCmdInfoEcrParams)
        Select Case pvPeek(sResult, 2)
        Case 1
            txtTaxModel.Text = "A100"
        Case 2
            txtTaxModel.Text = "A300"
        Case 4
            txtTaxModel.Text = "A500"
        Case 8
            txtTaxModel.Text = "A800"
        Case 16
            txtTaxModel.Text = "A600"
        Case Else
            txtTaxModel.Text = "Model 0x" & Hex(pvPeek(sResult, 2))
        End Select
        Select Case pvPeek(sResult, 1)
        Case &H30, &H70, &HF0
            sResult = "70a"
        Case &HF4
            sResult = "70x"
        Case Else
            sResult = "Unknown 0x" & Hex(pvPeek(sResult, 1))
        End Select
        txtTaxModel.Text = txtTaxModel.Text & " (" & sResult & ")"
        txtTaxMemModule.Text = pvFromPbcd(m_oFP.SendCommand(ucsEltCmdInfoFiscalNumber))
        txtTaxSerNo.Text = m_oFP.SendCommand(ucsEltCmdInfoFpNumber)
        sResult = m_oFP.SendCommand(ucsEltCmdInfoVatGroups)
        For lIdx = 0 To 7
            txtTaxGroup(lIdx).Text = pvPeek(sResult, 0 + 2 * lIdx, 2) / 100#
        Next
    Case ucsCmdDateTime
        sResult = pvFromPbcd(m_oFP.SendCommand(ucsEltCmdInfoDateTime))
        txtDateDate.Text = Mid$(sResult, 5, 2) & "-" & Mid$(sResult, 7, 2) & "-" & Mid$(sResult, 9, 2)
        txtDateTime.Text = Mid$(sResult, 3, 2) & ":" & Mid$(sResult, 1, 2)
    Case ucsCmdHeaderFooter
        txtHeadBulstatText.Text = m_oFP.SendCommand(ucsEltCmdInfoBulstat)
        sResult = m_oFP.SendCommand(ucsEltCmdInfoHeadingLines)
        For lIdx = 0 To 6
            If pvPeek(sResult, lIdx * (m_lRowChars + 1)) <> 0 Then
                chkHeadHeader(lIdx).Value = IIf(pvPeek(sResult, lIdx * (m_lRowChars + 1)) And &H10, vbChecked, vbUnchecked)
                txtHeadHeader(lIdx).Text = Trim$(pvFromMik(Mid$(sResult, lIdx * (m_lRowChars + 1) + 2, m_lRowChars)))
            End If
        Next
        txtHeadHeader(7).Text = Trim$(pvFromMik(Mid$(sResult, 7 * (m_lRowChars + 1) + 2, m_lRowChars \ 2)))
        txtHeadHeader(8).Text = Trim$(pvFromMik(Mid$(sResult, 7 * (m_lRowChars + 1) + 2 + m_lRowChars \ 2, m_lRowChars - m_lRowChars \ 2)))
    Case ucsCmdInvoiceNo
        sResult = m_oFP.SendCommand(ucsEltCmdInfoInvoiceNo)
        txtInvStart.Text = pvPeek(Mid$(sResult, 5, 2) & sResult, 0, 6)
        txtInvEnd.Text = pvPeek(sResult, 6, 2)
    Case ucsCmdPaymentTypes
        sResult = m_oFP.SendCommand(ucsEltCmdInfoPaymentTypes)
        For lIdx = 0 To 3
            txtPmtRate(lIdx).Text = pvPeek(sResult, 0 + 2 * lIdx, 2)
            txtPmtType(lIdx).Text = RTrim$(pvFromMik(Mid$(sResult, 9 + 6 * lIdx, 6)))
        Next
    Case ucsCmdOperators
        If Not IsArray(m_vOpers) Then
            ReDim m_vOpers(0 To LNG_NUM_OPERS) As Variant
        End If
        For lIdx = 1 To UBound(m_vOpers)
            If Not IsArray(m_vOpers(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_OPER, lIdx)
                sResult = m_oFP.SendCommand(ucsEltCmdInfoOperator, pvPoke(lIdx, 2))
                If Len(sResult) >= 18 Then
                    m_vOpers(lIdx) = Array(pvPeek(sResult, 0, 2), pvSPeek(sResult, 2, 12), pvSPeek(sResult, 14, 4), Empty)
                    If LenB(At(m_vOpers(lIdx), 1)) <> 0 Then
                        m_vOpers(lIdx)(3) = m_oFP.SendCommand(ucsEltCmdInfoTurnoverOperator, pvPoke(lIdx, 2))
                    End If
                End If
            End If
            If lstOpers.ListCount < lIdx Then
                lstOpers.AddItem vbNullString
            End If
            sText = At(m_vOpers(lIdx), 0) & ": " & At(m_vOpers(lIdx), 1)
            If lstOpers.List(lIdx - 1) <> sText Then
                lstOpers.List(lIdx - 1) = sText
            End If
        Next
        lstOpers_Click
        pvStatus = vbNullString
    Case ucsCmdDepartments
        If Not IsArray(m_vDeps) Then
            ReDim m_vDeps(0 To LNG_NUM_DEPS) As Variant
        End If
        For lIdx = 1 To UBound(m_vDeps)
            If Not IsArray(m_vDeps(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_DEP, lIdx)
                sResult = m_oFP.SendCommand(ucsEltCmdInfoDepartment, pvPoke(lIdx - 1, 2))
                If Len(sResult) >= 30 Then
                    If pvPeek(12, 1) = 0 Then
                        m_vDeps(lIdx) = Array(pvPeek(sResult, 0, 2) + 1, pvSPeek(sResult, 13, 18), _
                            pvPeek(sResult, 6, 1), pvPeek(sResult, 7, 1), pvPeek(sResult, 8, 4), _
                            pvPeek(sResult, 32, 4))
                    Else
                        m_vDeps(lIdx) = Array(pvPeek(sResult, 0, 2) + 1, pvSPeek(sResult, 12, 12), _
                            pvPeek(sResult, 6, 1), pvPeek(sResult, 7, 1), pvPeek(sResult, 8, 4), _
                            pvPeek(sResult, 26, 4))
                    End If
                End If
            End If
            If lstDeps.ListCount < lIdx Then
                lstDeps.AddItem vbNullString
            End If
            sText = At(m_vDeps(lIdx), 0) & ": " & At(m_vDeps(lIdx), 1)
            If lstDeps.List(lIdx - 1) <> sText Then
                lstDeps.List(lIdx - 1) = sText
            End If
        Next
        lstDeps_Click
        pvStatus = vbNullString
    Case ucsCmdItems
        If Not IsArray(m_vItems) Then
            ReDim m_vItems(0 To LNG_NUM_ITEMS) As Variant
        End If
        For lIdx = 1 To UBound(m_vItems)
            If Not IsArray(m_vItems(lIdx)) Then
                pvStatus = Printf(STR_STATUS_FETCH_ITEM, lIdx)
                sResult = m_oFP.SendCommand(ucsEltCmdInfoItem, pvPoke(lIdx, 2))
'                If LenB(pvSPeek(sResult, 16, 12)) = 0 Then
'                    Exit For
'                End If
                '-- Item array indexs:
                '-- 0 - Item No
                '-- 1 - Item Name
                '-- 2 - Single price
                '-- 3 - Available
                '-- 4 - Item group
                '-- 5 - VAT group
                '-- 6 - Turnover
                '-- 7 - Type of payment
                '-- 8 - Subgroup
                '-- 9 - Flags
                '-- 10 - Quantity sold
                If pvPeek(sResult, 16, 1) <> 0 Then
                    m_vItems(lIdx) = Array(pvPeek(sResult, 0, 2), pvSPeek(sResult, 16, 12), _
                        pvPeek(sResult, 2, 4), pvPeek(sResult, 6, 4), _
                        pvPeek(sResult, 10, 1), pvPeek(sResult, 11, 1), _
                        pvPeek(sResult, 12, 4), pvPeek(sResult, 28, 1), _
                        pvPeek(sResult, 29, 1), pvPeek(sResult, 30, 1), _
                        pvPeek(sResult, 31, 4))
                Else
                    m_vItems(lIdx) = Array(pvPeek(sResult, 0, 2), pvSPeek(sResult, 17, 18), _
                        pvPeek(sResult, 2, 4), pvPeek(sResult, 6, 4), _
                        pvPeek(sResult, 10, 1), pvPeek(sResult, 11, 1), _
                        pvPeek(sResult, 12, 4), -1, _
                        pvPeek(sResult, 35, 1), pvPeek(sResult, 36, 1), _
                        pvPeek(sResult, 37, 4))
                End If
            End If
            If lstItems.ListCount < lIdx Then
                lstItems.AddItem vbNullString
            End If
            sText = At(m_vItems(lIdx), 0) & ": " & At(m_vItems(lIdx), 1)
            If lstItems.List(lIdx - 1) <> sText Then
                lstItems.List(lIdx - 1) = sText
            End If
        Next
        lstItems_Click
        pvStatus = vbNullString
    Case ucsCmdCashOper
        sResult = m_oFP.SendCommand(ucsEltCmdInfoTurnoverVatGroups)
        For lIdx = 0 To 7
            dblSum = dblSum + pvPeek(sResult, lIdx * 4, 4) / 100#
        Next
        txtCashTotal.Text = Format$(dblSum, "0.00")
    Case ucsCmdReports
        pvStatus = labConnectCurrent.Caption
    Case ucsCmdStatus
        sResult = m_oFP.SendCommand(ucsEltCmdInfoStatus)
        txtStatus(0).Text = Format$(pvPeek(sResult, 0, 4) / 100#, "0.00")
        For lIdx = 1 To 8
            txtStatus(lIdx).Text = pvPeek(sResult, 2 + lIdx * 2, 2)
        Next
    Case ucsCmdYesNoParams
        sResult = m_oFP.SendCommand(ucsEltCmdInfoYesNoParams)
        For lIdx = 0 To 14
            lstYesNoParams.Selected(lIdx) = pvPeek(sResult, lIdx * 2, 2) <> 0
        Next
        lstYesNoParams.ListIndex = 0
    Case ucsCmdKeys
        sResult = m_oFP.SendCommand(ucsEltCmdInfoKeyFunctions)
        txtKeysNoNumber.Text = pvDumpHex(Mid$(sResult, 1, Len(sResult) \ 2))
        txtKeysWithNumber.Text = pvDumpHex(Mid$(sResult, 1 + Len(sResult) \ 2, Len(sResult)))
    Case ucsCmdLog
        m_sLog = Right$(m_sLog, 32000)
        txtLog.Text = m_sLog
        txtLog.SelStart = Len(m_sLog)
        pvStatus = labConnectCurrent.Caption
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
    Dim sResult         As String
    Dim lIdx            As Long
    Dim dDate           As Date
    Dim vSplit          As Variant
    Dim sData           As String
    Dim lReceipt        As Long
    
    On Error GoTo EH
    If Not m_oFP.IsConnected And eCommand <> ucsCmdConnect Then
        Exit Function
    End If
    Select Case eCommand
    Case ucsCmdConnect
        '--- value might be not be found
        On Error Resume Next
        DeleteSetting CAP_MSG, "Connect", "Port"
        On Error GoTo EH
        pvStatus = STR_STATUS_CONNECTING
        If m_oFP.Init(cobConnectPort.Text & "," & C_Lng(cobConnectSpeed.Text), m_lTimeout, m_lCashDeskNo) Then
            On Error Resume Next
            m_oFP.SendCommand ucsEltCmdInfoDeviceStatus
            If pvShowError() Then
                On Error GoTo EH
                labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
                Caption = CAP_MSG
            Else
                On Error GoTo EH
                sResult = m_oFP.SendCommand(ucsEltCmdInfoEcrParams)
                m_lRowChars = pvPeek(sResult, 3, 1)
                If m_lRowChars = 0 Then
                    sResult = m_oFP.SendCommand(ucsEltCmdInfoHeadingLines)
                    m_lRowChars = Len(sResult) / 8 - 1
                End If
                labConnectCurrent.Caption = Printf(STR_STATUS_SUCCESS_CONNECT, m_oFP.Device)
                Caption = m_oFP.Device & " - " & CAP_MSG
                '--- save conn info
'                If chkConnectRemember.Value Then
'                    SaveSetting CAP_MSG, "Connect", "Port", cobConnectPort.Text
'                    SaveSetting CAP_MSG, "Connect", "Speed", cobConnectSpeed.Text
'                End If
                '--- flush cache
                m_vDeps = Empty
                lstDeps.Clear
                m_vOpers = Empty
                lstOpers.Clear
                m_vItems = Empty
                lstItems.Clear
'                m_vLogo = Empty
                m_sLog = vbNullString
                lstCmds.ListIndex = ucsCmdTaxInfo
            End If
        Else
            labConnectCurrent.Caption = STR_STATUS_FAILURE_CONNECT
            Caption = CAP_MSG
        End If
    Case ucsCmdTaxInfo
        sData = pvPoke(&HFFFF&, 2)
        For lIdx = 0 To 7
            sData = sData & pvPoke(C_Lng(C_Dbl(txtTaxGroup(lIdx)) * 100), 2)
        Next
        sData = sData & Mid$(m_oFP.SendCommand(ucsEltCmdInfoHeadingLines), 1, 2 * (m_lRowChars + 1))
        sResult = m_oFP.SendCommand(ucsEltCmdInitVatGroups, sData)
        If pvPeek(sResult, 0, 2) <> &HFFFF& Then
            pvStatus = Printf(STR_ERROR_VATRATE_SEND, Hex(pvPeek(sResult, 0, 2)))
'            GoTo QH
        End If
    Case ucsCmdDateTime
        vSplit = Split(txtDateDate.Text, "-")
        dDate = DateSerial(vSplit(2), vSplit(1), vSplit(0)) + CDate(txtDateTime.Text)
        If dDate = 0 Then
            pvStatus = STR_ERROR_DATETIME_FORMAT
            GoTo QH
        End If
        sData = Chr$("&H" & Day(dDate)) & Chr$("&H" & Month(dDate)) & Chr$("&H" & (Year(dDate) Mod 100)) & _
            Chr$("&H" & Weekday(dDate, vbMonday)) & Chr$("&H" & Minute(dDate)) & Chr$("&H" & Hour(dDate))
        sResult = m_oFP.SendCommand(ucsEltCmdInitDateTime, sData)
        If pvPeek(sResult, 0, 2) <> &HFFFF& Then
            pvStatus = Printf(STR_ERROR_DATETIME_SEND, Hex(pvPeek(sResult, 0, 2)))
            GoTo QH
        End If
    Case ucsCmdHeaderFooter
        sData = pvPoke(&HFFFF&, 2) & Mid$(m_oFP.SendCommand(ucsEltCmdInfoVatGroups), 1, 16)
        For lIdx = 0 To 1
            If chkHeadHeader(lIdx).Value = vbChecked Then
                sData = sData & Chr$(&H90)
            Else
                sData = sData & Chr$(&H80)
            End If
            sData = sData & pvSPoke(txtHeadHeader(lIdx).Text, m_lRowChars, vbCenter)
        Next
        sResult = m_oFP.SendCommand(ucsEltCmdInitVatGroups, sData)
        If pvPeek(sResult, 0, 2) <> &HFFFF& Then
            pvStatus = Printf(STR_ERROR_HDRLINES_SEND, Hex(pvPeek(sResult, 0, 2)))
'            GoTo QH
        End If
        sData = vbNullString
        For lIdx = 2 To 6
            If LenB(txtHeadHeader(lIdx).Text) = 0 Then
                sData = sData & Chr$(0)
            ElseIf chkHeadHeader(lIdx).Value = vbChecked Then
                sData = sData & Chr$(&H90)
            Else
                sData = sData & Chr$(&H80)
            End If
            sData = sData & pvSPoke(txtHeadHeader(lIdx).Text, m_lRowChars, vbCenter)
        Next
        '--- special row: Bulstat/Ident No
        sData = sData & Chr$(&H80)
        sData = sData & pvSPoke(txtHeadHeader(7).Text & " ", m_lRowChars \ 2, vbRightJustify) & pvSPoke(txtHeadHeader(8).Text & " ", m_lRowChars - m_lRowChars \ 2, vbRightJustify)
        sResult = m_oFP.SendCommand(ucsEltCmdInitHeadingLines, sData)
        If pvPeek(sResult, 0, 2) <> &HFFFF& Then
            pvStatus = Printf(STR_ERROR_HDRLINES_SEND, Hex(pvPeek(sResult, 0, 2)))
'            GoTo QH
        End If
    Case ucsCmdInvoiceNo
        sData = pvPoke(txtInvStart.Text, 6)
        sData = Mid$(sData, 3, 4) & Mid$(sData, 1, 2) & pvPoke(txtInvEnd.Text, 2)
        Call m_oFP.SendCommand(ucsEltCmdInitInvoiceNo, sData)
        m_oFP.WaitDevice
    Case ucsCmdPaymentTypes
        For lIdx = 0 To 3
            sData = sData & pvPoke(C_Lng(txtPmtRate(lIdx).Text), 2)
        Next
        For lIdx = 0 To 3
            sData = sData & pvSPoke(txtPmtType(lIdx).Text, 6)
        Next
        sResult = m_oFP.SendCommand(ucsEltCmdInitPaymentTypes, sData)
    Case ucsCmdOperators
        If C_Lng(txtOperNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_OPER_SELECTED
            GoTo QH
        End If
        If LenB(txtOperName.Text) <> 0 Then
            If Left$(txtOperPass.Text, 4) <> Left$(txtOperPass2.Text, 4) Then
                txtOperPass2.SetFocus
                txtOperPass2.SelStart = 0
                txtOperPass2.SelLength = Len(txtOperPass2.Text)
                MsgBox MSG_PASSWORDS_MISMATCH, vbExclamation
                GoTo QH
            End If
        End If
        sData = pvPoke(C_Lng(txtOperNo.Text), 2) & pvFromMik(pvSPoke(txtOperName.Text, 12)) & _
            pvSPoke(IIf(LenB(txtOperName.Text) <> 0, txtOperPass.Text, vbNullString), 4, lFillChar:=0) & pvPoke(0, 2)
        sResult = m_oFP.SendCommand(ucsEltCmdInitOperator, sData)
        m_vOpers(C_Lng(txtOperNo.Text)) = Empty
    Case ucsCmdDepartments
        If C_Lng(txtDepNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_DEP_SELECTED
            GoTo QH
        End If
        sData = m_oFP.SendCommand(ucsEltCmdInfoDepartment, pvPoke(C_Lng(txtDepNo.Text) - 1, 2))
        Mid$(sData, 7, 1) = Chr$(C_Lng(txtDepItemGroup.Text))
        Mid$(sData, 8, 1) = Chr$(cobDepGroup.ListIndex)
        If Asc(Mid$(sData, 13, 1)) <> 0 Then
            Mid$(sData, 13, 12) = pvSPoke(txtDepName.Text, 12)
        Else
            Mid$(sData, 14, 18) = pvSPoke(txtDepName.Text, 18)
        End If
        sResult = m_oFP.SendCommand(ucsEltCmdInitDepartment, sData)
        m_vDeps(C_Lng(txtDepNo.Text)) = Empty
    Case ucsCmdItems
        If C_Lng(txtItemNo.Text) = 0 Then
            pvStatus = STR_STATUS_NO_ITEM_SELECTED
            GoTo QH
        End If
        lIdx = -lstItemFlags.Selected(0) * 2 ^ 0 + _
               -lstItemFlags.Selected(1) * 2 ^ 1 + _
               -lstItemFlags.Selected(2) * 2 ^ 2 + _
               -lstItemFlags.Selected(3) * 2 ^ 3 + _
               -lstItemFlags.Selected(4) * 2 ^ 5
        sData = m_oFP.SendCommand(ucsEltCmdInfoItem, pvPoke(C_Lng(txtItemNo.Text), 2))
        Mid$(sData, 3, 4) = pvPoke(C_Lng(C_Dbl(txtItemPrice.Text) * 100), 4)
        Mid$(sData, 11, 1) = pvPoke(LimitLong(C_Lng(txtItemGroup.Text) - 1, 0, 99), 1)
        Mid$(sData, 12, 1) = pvPoke(cobItemGroup.ListIndex, 1)
        If Len(sData) >= 41 Then
            Mid$(sData, 18, 18) = pvSPoke(txtItemName.Text, 18)
            Mid$(sData, 36, 1) = pvPoke(LimitLong(C_Lng(txtItemDep.Text) - 1, 0, 99), 1)
            Mid$(sData, 37, 1) = Chr$(lIdx)
        Else
            Mid$(sData, 17, 12) = pvSPoke(txtItemName.Text, 12)
            Mid$(sData, 30, 1) = pvPoke(LimitLong(C_Lng(txtItemDep.Text) - 1, 0, 99), 1)
            Mid$(sData, 31, 1) = Chr$(lIdx)
        End If
        sResult = m_oFP.SendCommand(ucsEltCmdInitItem, sData)
        m_vItems(C_Lng(txtItemNo.Text)) = Empty
    Case ucsCmdCashOper
'        If C_Dbl(txtCashSum.Text) <> 0 Then
'            sResult = m_oFP.SendCommand(ucsEltCmdInfoKeyFunctions)
'            txtCashSum.Text = vbNullString
'        End If
    Case ucsCmdReports
        If optReportType(0).Value Then
            If chkReportItems.Value = vbUnchecked And chkReportDepartments.Value = vbUnchecked Then
                If chkReportClosure.Value = vbUnchecked Then
                    '--- X report
                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(2))
                    lReceipt = pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2)
                    sData = Chr$(&H20) & Chr$(&H1C) & Chr$(&H20)
                    sResult = m_oFP.SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
                    For lIdx = 1 To 150
                        If lReceipt <> pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2) Then
                            Exit For
                        End If
                        m_oFP.WaitDevice 100
                    Next
                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(0))
                Else
                    '--- Z report
                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(4))
                    lReceipt = pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2)
                    sData = Chr$(&H20)
                    sResult = m_oFP.SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
                    For lIdx = 1 To 150
                        If lReceipt <> pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2) Then
                            Exit For
                        End If
                        m_oFP.WaitDevice 100
                    Next
                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(0))
                End If
            ElseIf chkReportItems.Value = vbChecked And chkReportDepartments.Value = vbUnchecked Then
                If chkReportClosure.Value = vbUnchecked Then
                    '--- po artikulri bez nulirane
'                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(2))
'                    sData = Chr$(&H20)
'                    sResult = m_oFP.SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
'                    sData = Chr$(32) & Chr$(7) & Chr$(&H20)
'                    sResult = m_oFP.SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
'                    m_oFP.WaitDevice 5000
'                    sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(0))
                Else
                
                End If
            End If
        ElseIf optReportType(2).Value Then
        ElseIf optReportType(3).Value Then
            If C_Date(txtReportFD.Text) = 0 Then
                txtReportFD.SetFocus
                MsgBox MSG_INVALID_DATE, vbExclamation
                GoTo QH
            End If
            If C_Date(txtReportTD.Text) = 0 Then
                txtReportTD.SetFocus
                MsgBox MSG_INVALID_DATE, vbExclamation
                GoTo QH
            End If
            sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(2))
            lReceipt = pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2)
            sData = IIf(chkReportDetailed2.Value = vbChecked, "1", "11") & Chr$(&H1F) & _
                Format$(C_Date(txtReportFD.Text), "ddmmyy") & Chr$(&H1A) & _
                Format$(C_Date(txtReportTD.Text), "ddmmyy") & Chr$(&H1B)
            sResult = m_oFP.SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
            For lIdx = 1 To 150
                If lReceipt <> pvPeek(m_oFP.SendCommand(ucsEltCmdInfoStatus), 6, 2) Then
                    Exit For
                End If
                m_oFP.WaitDevice 100
            Next
            sResult = m_oFP.SendCommand(ucsEltCmdInitKeylock, Chr$(0))
        End If
    Case ucsCmdYesNoParams
        sData = m_oFP.SendCommand(ucsEltCmdInfoYesNoParams)
        For lIdx = 0 To 14
            Mid$(sData, lIdx * 2 + 1, 2) = pvPoke(lstYesNoParams.Selected(lIdx), 2)
        Next
        sResult = m_oFP.SendCommand(ucsEltCmdInitYesNoParams, sData)
    End Select
    '--- success
    pvSaveData = True
QH:
    Exit Function
EH:
    If pvShowError() Then
        Exit Function
    End If
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Dim lI              As Long
    
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

Private Function pvPeek(sText As String, ByVal lOffset As Long, Optional ByVal lSize As Long = 1) As Currency
    Dim sTemp           As String
    
    sTemp = StrReverse(Mid$(sText, lOffset + 1, lSize))
    Call CopyMemory(pvPeek, ByVal sTemp, Len(sTemp))
    pvPeek = pvPeek * 10000@
End Function

Private Function pvPoke(ByVal cValue As Currency, Optional ByVal lSize As Long = 1) As String
    Dim sTemp           As String
    
    cValue = cValue / 10000@
    sTemp = String$(8, 0)
    Call CopyMemory(ByVal sTemp, cValue, 8)
    pvPoke = Right$(StrReverse(sTemp), lSize)
End Function

Private Function pvSPeek(sText As String, ByVal lOffset As Long, ByVal lSize As Long) As String
    pvSPeek = pvFromMik(RTrim$(Replace$(Mid$(sText, lOffset + 1, lSize), Chr$(0), vbNullString)))
End Function

Private Function pvSPoke(sText As String, ByVal lSize As Long, Optional ByVal eAlign As AlignmentConstants = vbLeftJustify, Optional ByVal lFillChar As Long = 32) As String
    Select Case eAlign
    Case vbLeftJustify
        pvSPoke = Left$(sText, lSize)
        pvSPoke = pvSPoke & String$(lSize - Len(pvSPoke), lFillChar)
    Case vbRightJustify
        pvSPoke = Right$(sText, lSize)
        pvSPoke = String$(lSize - Len(pvSPoke), lFillChar) & pvSPoke
    Case vbCenter
        pvSPoke = Left$(sText, lSize)
        pvSPoke = String$((lSize - Len(pvSPoke) + 1) \ 2, lFillChar) & pvSPoke
        pvSPoke = pvSPoke & String$(lSize - Len(pvSPoke), lFillChar)
    End Select
    pvSPoke = pvToMik(pvSPoke)
End Function

Private Function pvFromPbcd(sText As String) As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 0 To UBound(baText)
        pvFromPbcd = pvFromPbcd & Right$("0" & Hex(baText(lIdx)), 2)
    Next
End Function

Private Function pvFromMik(sText As String) As String
    Dim lIdx            As Long
    Dim lChar           As Long
    
    pvFromMik = sText
    For lIdx = 1 To Len(sText)
        lChar = Asc(Mid$(sText, lIdx, 1))
        If lChar >= &H80 And lChar < &H80 + 64 Then
            Mid$(pvFromMik, lIdx, 1) = Chr$(lChar + 64)
        End If
    Next
End Function

Private Function pvToMik(sText As String) As String
    Dim lIdx            As Long
    Dim lChar           As Long
    
    pvToMik = sText
    For lIdx = 1 To Len(sText)
        lChar = Asc(Mid$(sText, lIdx, 1))
        If lChar >= &HC0 And lChar < &HC0 + 64 Then
            Mid$(pvToMik, lIdx, 1) = Chr$(lChar - 64)
        End If
    Next
End Function

Private Function pvDumpHex(sText As String, Optional sSeparator As String = " ") As String
    Dim lIdx        As Long
    
    For lIdx = 1 To Len(sText)
        pvDumpHex = pvDumpHex & sSeparator & Right$("0" & Hex(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
    pvDumpHex = Mid$(pvDumpHex, 1 + Len(sSeparator), Len(pvDumpHex))
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub cmdKeysReset_Click()
    Const FUNC_NAME     As String = "cmdKeysReset_Click"
    Dim sData           As String
    
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
    sData = Chr$(&H14) & Chr$(&HD) & Chr$(&H19) & Chr$(&H16) & Chr$(&H16) & Chr$(&H16) & Chr$(&H4) & Chr$(&H8) & _
            Chr$(&H1B) & Chr$(&H3) & Chr$(&H7) & Chr$(&H1C) & Chr$(&H2) & Chr$(&H6) & Chr$(&H1) & Chr$(&H5) & _
            Chr$(&HE) & Chr$(&HF) & _
            Chr$(&H14) & Chr$(&HD) & Chr$(&H18) & Chr$(&H13) & Chr$(&H12) & Chr$(&H1D) & Chr$(&H4) & Chr$(&H8) & _
            Chr$(&H1A) & Chr$(&H3) & Chr$(&H7) & Chr$(&H16) & Chr$(&H2) & Chr$(&H6) & Chr$(&H1) & Chr$(&H5) & _
            Chr$(&HE) & Chr$(&HF)
    m_oFP.SendCommand ucsEltCmdInitKeyFunctions, sData
    pvFetchData ucsCmdKeys
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
    If Not m_oFP.IsConnected And lstCmds.ListIndex <> ucsCmdConnect Then
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
            pvStatus = Printf(STR_STATUS_SUCCESS_FETCH, Trim$(lstCmds.List(lstCmds.ListIndex)), Round(Timer - dblTimer, 2))
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
    For lIdx = fraCommands.LBound To fraCommands.UBound
        fraCommands(lIdx).Visible = (lIdx = lVisibleFrame)
    Next
    On Error GoTo EH
    tmrDate.Enabled = (lVisibleFrame = ucsCmdDateTime)
    Call SendMessage(txtLog.hWnd, EM_SCROLLCARET, 0, ByVal 0&)
    '--- might have missing entries in control array
    On Error Resume Next
    For lIdx = cmdSave.LBound To cmdSave.UBound
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
                pvStatus = Printf(STR_STATUS_SUCCESS_SAVE, Trim$(lstCmds.List(lstCmds.ListIndex)), Round(Timer - dblTimer, 2))
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
        vResult = m_vDeps(lstDeps.ListIndex + 1)
    End If
    txtDepNo.Text = At(vResult, 0)
    txtDepName.Text = At(vResult, 1)
    If LenB(txtDepNo.Text) <> 0 Then
        cobDepGroup.ListIndex = C_Lng(At(vResult, 3))
        txtDepItemGroup.Text = C_Lng(At(vResult, 2))
        txtDepSales.Text = Format$(C_Lng(At(vResult, 5)) / 1000#, "0.000")
        txtDepTotalSum.Text = Format$(C_Lng(At(vResult, 4)) / 100#, "0.00")
    Else
        cobDepGroup.ListIndex = -1
        txtDepItemGroup.Text = vbNullString
        txtDepSales.Text = vbNullString
        txtDepTotalSum.Text = vbNullString
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub lstItems_Click()
    Const FUNC_NAME     As String = "lstItems_Click"
    Dim vResult         As Variant
    Dim lFlags          As Long
    
    On Error GoTo EH
    If lstItems.ListIndex >= 0 Then
        vResult = m_vItems(lstItems.ListIndex + 1)
    End If
    txtItemNo.Text = At(vResult, 0)
    txtItemName.Text = At(vResult, 1)
    If LenB(txtItemNo.Text) <> 0 Then
        txtItemPrice.Text = Format$(C_Lng(At(vResult, 2)) / 100#, "0.00")
        cobItemGroup.ListIndex = C_Lng(At(vResult, 5))
        txtItemGroup.Text = C_Lng(At(vResult, 4)) + 1
        txtItemDep.Text = C_Lng(At(vResult, 8)) + 1
        txtItemAvailable.Text = Format$(C_Lng(At(vResult, 3)) / 1000#, "0.000")
        txtItemTurnover.Text = Format$(C_Lng(At(vResult, 6)) / 1000#, "0.000")
        txtItemSoldQuo.Text = Format$(C_Lng(At(vResult, 10)) / 100#, "0.00")
        lFlags = C_Lng(At(vResult, 9))
    Else
        txtItemPrice.Text = vbNullString
        cobItemGroup.ListIndex = -1
        txtItemGroup.Text = vbNullString
        txtItemDep.Text = vbNullString
        txtItemAvailable.Text = vbNullString
        txtItemTurnover.Text = vbNullString
        txtItemSoldQuo.Text = vbNullString
    End If
    lstItemFlags.Selected(0) = (lFlags And 2 ^ 0) <> 0
    lstItemFlags.Selected(1) = (lFlags And 2 ^ 1) <> 0
    lstItemFlags.Selected(2) = (lFlags And 2 ^ 2) <> 0
    lstItemFlags.Selected(3) = (lFlags And 2 ^ 3) <> 0
    lstItemFlags.Selected(4) = (lFlags And 2 ^ 5) <> 0
    lstItemFlags.ListIndex = 0
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
        vResult = m_vOpers(lstOpers.ListIndex + 1)
    End If
    txtOperNo.Text = At(vResult, 0)
    txtOperName.Text = At(vResult, 1)
    txtOperPass.Text = At(vResult, 2)
    txtOperPass2.Text = vbNullString
    If C_Str(pvPeek(At(vResult, 3), 0, 2)) = At(vResult, 0) Then
        txtOperFiscal.Text = pvPeek(At(vResult, 3), 2, 2)
        txtOperSells.Text = Format$((pvPeek(At(vResult, 3), 28, 4) + pvPeek(At(vResult, 3), 32, 4) + pvPeek(At(vResult, 3), 36, 4) + pvPeek(At(vResult, 3), 40, 4)) / 100#, "0.00")
        txtOperDisc.Text = Format$(pvPeek(At(vResult, 3), 20, 4) / 100#, "0.00")
        txtOperSurcharge.Text = Format$(pvPeek(At(vResult, 3), 24, 4) / 100#, "0.00")
        txtOperVoid.Text = Format$(pvPeek(At(vResult, 3), 48, 4) / 100#, "0.00")
        txtOperResto.Text = Format$(pvPeek(At(vResult, 3), 44, 4) / 100#, "0.00")
    Else
        txtOperFiscal.Text = vbNullString
        txtOperSells.Text = vbNullString
        txtOperDisc.Text = vbNullString
        txtOperSurcharge.Text = vbNullString
        txtOperVoid.Text = vbNullString
        txtOperResto.Text = vbNullString
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub tmrDate_Timer()
    txtDateCompDate.Text = Format$(Now, "dd-MM-yy")
    txtDateCompTime.Text = Format$(Now, "hh:mm:ss")
End Sub

Private Sub cmdDateTransfer_Click()
    txtDateDate.Text = txtDateCompDate.Text
    txtDateTime.Text = txtDateCompTime.Text
End Sub

Private Sub m_oFPSink_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFPSink_CommandComplete"
    
    On Error GoTo EH
    m_sLog = m_sLog & pvDumpHex(Chr$(lCmd)) & IIf(LenB(sData) <> 0, "<-" & pvDumpHex(sData), vbNullString) & IIf(LenB(sResult) <> 0, "->" & pvDumpHex(sResult), vbNullString) & vbCrLf
    If LenB(m_oFP.LastError) <> 0 Then
        m_sLog = m_sLog & m_oFP.LastError & vbCrLf
    End If
'    If m_oFP.Status(ucsStbPrintingError) Then
'        m_sLog = m_sLog & m_oFP.StatusText & vbCrLf & m_oFP.DipText & vbCrLf & m_oFP.MemoryText & vbCrLf
'    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
