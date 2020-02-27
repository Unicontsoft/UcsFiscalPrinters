VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDatecsProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' UcsFP20 (c) 2008-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Implementation of Datecs X models protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cDatecsProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsDatecsStatusBitsEnum
    ucsDcsStbSyntaxError = 2 ^ 0
    ucsDcsStbInvalidCommand = 2 ^ 1
    ucsDcsStbClockNotSet = 2 ^ 2
    ucsDcsStbFailurePrintingMechanism = 2 ^ 4
    ucsDcsStbGeneralError = 2 ^ 5
    ucsDcsStbCoverOpen = 2 ^ 6
    ucsDcsStbSumOverflow = 2 ^ (8 + 0)
    ucsDcsStbCommandNotPermitted = 2 ^ (8 + 1)
    ucsDcsStbOutOfPaper = 2 ^ (16 + 0)
    ucsDcsStbNearOutOfPaper = 2 ^ (16 + 1)
    ucsDcsStbEJFull = 2 ^ (16 + 2)
    ucsDcsStbFiscalReceiptOpen = 2 ^ (16 + 3)
    ucsDcsStbNearEJFull = 2 ^ (16 + 4)
    ucsDcsStbNonFiscalReceiptOpen = 2 ^ (16 + 5)
End Enum

Public Enum UcsDatecsFiscalMemoryEnum
    ucsDcsFmeAccessError = 2 ^ 0
    ucsDcsFmeTaxNumberNotSet = 2 ^ 1
    ucsDcsFmeSerialNoSet = 2 ^ 2
    ucsDcsFmeNearOutOnMemory = 2 ^ 3
    ucsDcsFmeOutOfMemory = 2 ^ 4
    ucsDcsFmeMemoryError = 2 ^ 5
    ucsDcsFmeFormatted = 2 ^ (8 + 1)
    ucsDcsFmeDeviceFiscalized = 2 ^ (8 + 3)
    ucsDcsFmeVatRatesAreSetOnce = 2 ^ (8 + 4)
End Enum

Public Enum UcsDatecsCommandsEnum
    ucsDcsCmdNoop = -1
    '--- init
    ucsDcsCmdInitConnectionMode = 45
    ucsDcsCmdInitDateTime = 61
    ucsDcsCmdInitInvoiceNo = 66
    ucsDcsCmdInitBulstat = 98
    ucsDcsCmdInitOperatorPassword = 101
    ucsDcsCmdInitItem = 107
    ucsDcsCmdInitOpenDrawer = 106
    '--- transaction
    ucsDcsCmdNonFiscalOpen = 38
    ucsDcsCmdNonFiscalClose = 39
    ucsDcsCmdNonFiscalText = 42
    ucsDcsCmdFiscalReversal = 43
    ucsDcsCmdLineFeed = 44
    ucsDcsCmdFiscalOpen = 48
    ucsDcsCmdFiscalSell = 49
    ucsDcsCmdFiscalSubtotal = 51
    ucsDcsCmdFiscalPayment = 53
    ucsDcsCmdFiscalText = 54
    ucsDcsCmdFiscalClose = 56
    ucsDcsCmdFiscalCgInfo = 57
    ucsDcsCmdFiscalCancel = 60
    ucsDcsCmdFiscalServiceDeposit = 70
    '--- info
    ucsDcsCmdInfoDateTime = 62
    ucsDcsCmdInfoLastFiscalBlock = 64
    ucsDcsCmdInfoTotalsByTaxGroups = 65
    ucsDcsCmdInfoFreeMemoryRecords = 68
    ucsDcsCmdInfoStatus = 74
    ucsDcsCmdInfoTransaction = 76
    ucsDcsCmdInfoLastFiscalDateTime = 86
    ucsDcsCmdInfoDepartment = 88
    ucsDcsCmdInfoDiagnostics = 90
    ucsDcsCmdInfoBulstat = 99
    ucsDcsCmdInfoTotalsByPayments = 110
    ucsDcsCmdInfoOperator = 112
    ucsDcsCmdInfoMemoryDump = 116
    ucsDcsCmdInfoDevice = 123
    ucsDcsCmdInfoEJ = 125
    ucsDcsCmdInfoProgramming = 255                 '-- command above &H7F
    '--- print
    ucsDcsCmdPrintReportTaxRates = 50
    ucsDcsCmdPrintDailyReport = 69
    ucsDcsCmdPrintDiagnostics = 71
    ucsDcsCmdPrintBarcode = 84
    ucsDcsCmdPrintSeparator = 92
    ucsDcsCmdPrintReportByDateDetailed = 94
    ucsDcsCmdPrintReportByNumberShort = 95
    ucsDcsCmdPrintReportByOperators = 105
    ucsDcsCmdPrintDuplicate = 109
    ucsDcsCmdPrintReportByItems = 111
    '--- paper
    ucsDcsCmdPaperLinefeed = 44
    ucsDcsCmdPaperCut = 46
    '--- display
    ucsDcsCmdDisplayClear = 33
    ucsDcsCmdDisplaySecondLine = 35
    ucsDcsCmdDisplayFirstLine = 47
    ucsDcsCmdDisplayClock = 63
    '--- misc
    ucsDcsCmdMiscRingSound = 80
End Enum

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1
Private Const EV_ERR                        As Long = &H80

Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

Private Type OVERLAPPED
    Internal                    As Long
    InternalHigh                As Long
    Offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "Timeout waiting for response|Invalid response length: %1|Invalid response checksum: %1|Response missing postamble: %1|Response missing preamble: %1|No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Error %1|Memory %1|Unsupported payment type %1|Cannot print|Wrong invoice number %1 generated|Unsupported discount type %1"
Private Const STR_STATUSES              As String = "Syntax error|Command code is invalid|The real time clock is not synchronized|3|Failure in printing mechanism|5|Cover is open|7|Overflow during command execution|Command is not permitted|2|3|4|5|6|7|End of paper|Near paper end|EJ is full|Fiscal receipt is open|EJ nearly full|Nonfiscal receipt is open"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Command code is invalid|2|3|Failure in printing mechanism|5|6|7|0|Command is not permitted|2|3|4|5|6|7|End of paper"
Private Const STR_MEMORY                As String = "Error when trying to access data stored in the FM|Tax number is set|Serial number and number of FM are set|There is space for less then 60 reports in Fiscal memory|Fiscal memory is full|FM error|Fiscal memory is not found or damaged|7|0|FM is formatted|2|Device is fiscalized|VAT are set at least once"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|SURCHARGE %1|DISCOUNT %1|SUBTOTAL|Sales %1"
Private Const STR_ERROR_CODES           As String = "-100001|ERR_IO|-100002|ERR_CHECKSUM|-100003|ERR_END_OF_DATA|-100004|ERR_NOTFOUND|-100005|ERR_NO_RECORDS|-100006|ERR_ABORTED|-100007|ERR_WRONG_MODE|-100008|ERR_NOT_READY|-100009|ERR_NOTHING_TO_PRINT|-100100|ERR_FM_BUSY|-100101|ERR_FM_FAILURE|-100102|ERR_FM_WRITE_PROTECTED|-100103|ERR_FM_WRONG_ADDRESS|-100104|ERR_FM_WRONG_SIZE|-100105|ERR_FM_NOT_CONNECTED|-100106|ERR_FM_WRONG_CHECK_SUM|-100107|ERR_FM_BLOCK_IS_EMPTY|-100108|ERR_FM_MAX_NUMBER|-100109|ERR_FM_WRONG_RANGE|-100110|ERR_FM_EMPTY_RANGE|-100111|ERR_FM_NEW_MODULE|-100112|ERR_FM_NOT_EMPTY|-100113|ERR_FM_NOT_EQUAL|-100114|ERR_FM_FULL|-100115|ERR_FM_NEED_UPDATE|-100116|ERR_FM_BLOCKED|-100400|ERR_LTP_VCCERR|-100401|ERR_LTP_SVPERR|-100402|ERR_LTP_STHERR|-100403|ERR_LTP_PESENS|-100404|ERR_LTP_HDSENS|-100405|ERR_LTP_NESENS|-100406|ERR_LTP_MKSENS|-100407|ERR_LTP_CUTERR|-100408|ERR_LTP_PR_ERR|-100409|ERR_LTP_PR_BUSY|-100410|ERR_LTP_BZLPDEC" & _
                                                    "|-100411|ERR_LTP_BZLCLMP|-100412|ERR_LTP_CHARGE_MODE|-100413|ERR_LTP_INZERR_MODE|-100414|ERR_LTP_MOTOR_OVERRUN|-100500|ERR_PROGRAM_SELF_CHECK_ERROR|-100501|ERR_SRAM_ERROR|-100502|ERR_FLASH_ERROR|-100503|ERR_SDCARD_ERROR|-100504|ERR_INVALID_MSG_FILE|-100505|ERR_FM_ERROR|-100506|ERR_NO_RAM_BATTERY|-100507|ERR_SAM_ERROR|-100508|ERR_RTC_ERROR|-100509|ERR_PROGRAM_EXRAM_CHECK_ERROR|-100510|ERR_SDCARD_WRONG_SIZE|-100511|ERR_TPM_ERROR|-101000|ERR_NO_HEAP_MEMORY|-101001|ERR_FILE_MANIPULATE|-101003|ERR_REJECTED|-101004|ERR_BAD_INPUT|-101005|ERR_IAP|-101006|ERR_NOT_POSSIBLE|-101007|ERR_TMOUT|-101007|ERR_TIMEOUT|-101008|ERR_INVALID_TIME|-101009|ERR_CANCELLED|-101010|ERR_INVALID_FORMAT|-101011|ERR_INVALID_DATA|-101012|ERR_PARSE_ERROR|-101013|ERR_HARDWARE_CONFIGURATION|-101014|ERR_ACCESS_DENIED|-101015|ERR_BAD_DATA_LENGTH|-101016|ERR_VERIFY_Z|-101017|ERR_NO_PERMISSION|-101500|ERR_NO_UPDATE" & _
                                                    "|-101501|ERR_UPDATE_IN_PROGRESS|-102000|ERR_LOW_BATTERY|-102001|ERR_LOW_BATTERY_WARNING|-102002|ERR_OPER_WRONG_PASSWORD|-102003|ERR_IDNUMBER_IS_EMPTY|-102004|ERR_NOT_FOUND_BLUETOOTH|-102005|ERR_DISPLAY_DISCONNECTED|-102006|ERR_PRINTER_DISCONNECTED|-102007|ERR_SD_NOT_PRESENT|-102008|ERR_SD2_NOT_PRESENT|-102009|ERR_VAT_RATES_ARE_EMPTY|-102010|ERR_HEADER_IS_EMPTY|-102011|ERR_ZDDS_NUM_IS_EMPTY|-102012|ERR_FMNUMBER_IS_EMPTY|-102013|ERR_SERVICEMAN_NAME_IS_EMPTY|-102014|ERR_SERVICEMAN_ID_IS_EMPTY|-102015|ERR_TAXOFFICE_ID_IS_EMPTY|-102016|ERR_WRONG_FORMAT|-102017|ERR_TAXNUMBER_IS_EMPTY|-102018|ERR_WRONG_IDNUMBER|-102019|ERR_DATETIME_EARLIER_THAN_PREV_Z|-102020|ERR_NEED_SOFTWARE_PASSWORD|-103000|ERR_PLUDB_NOT_FOUND|-103001|ERR_PLUDB_PLUCODE_EXISTS|-103002|ERR_PLUDB_BARCODE_EXISTS|-103003|ERR_PLUDB_FULL|-103004|ERR_P_HAVE_TURNOVER|-103005|ERR_PLUDB_NAME_EXISTS|-103006|ERR_PLUDB_NAMES_NOT_UNIQUE" & _
                                                    "|-103007|ERR_PLUDB_FORMAT_INCOMPATIBLE|-103008|ERR_PLUDB_CAN_NOT_OPEN|-104000|ERR_NEED_Z_REPORT|-104001|ERR_NEED_SERVICE_JUMPER|-104002|ERR_NEED_SERVICE_PASSWORD|-104003|ERR_FORBIDEN|-104004|ERR_NEED_SERVICE_INTERVENTION|-104005|ERR_NEED_ALL_CLEARING_REPORTS|-104006|ERR_Z_REPORT_CLOSED|-104007|ERR_NEED_MONTH_REPORT|-104008|ERR_NEED_YEAR_REPORT|-104009|ERR_NEED_BACKUP|-104010|ERR_NEED_ALL_PAIDOUT|-104011|ERR_NEED_OPERATOR_Z_REPORT|-104012|ERR_NEED_GROUP_Z_REPORT|-104013|ERR_NEED_VAT_CHANGES|-105000|ERR_EJ_NO_RECORDS|-105001|ERR_CANNOT_ADD_TO_EJ|-105002|ERR_EJ_WRONG_MAC_RECORD|-105003|ERR_EJ_IMMPOSSIBLE_TO_CHK_MAC_RECORD|-105004|ERR_EJ_BAD_RECORDS|-105005|ERR_EJ_CAN_NOT_GENERATE_MAC|-105006|ERR_EJ_WRONG_TYPE_TO_SIGN|-105007|ERR_EJ_ALREADY_SIGNED|-105008|ERR_EJ_NOT_FROM_THIS_DEVICE|-105009|ERR_EJ_NEAR_FULL|-105010|ERR_EJ_FULL|-105011|ERR_EJ_WRONG_FORMAT|-105012|ERR_EJ_NOT_READY" & _
                                                    "|-105013|ERR_EJ_NEED_NEW|-106000|ERR_R_FIRM_NOTEXIST|-106001|ERR_FIRMDB_FIRMCODE_EXISTS|-106002|ERR_FIRMDB_EIK_EXISTS|-106003|ERR_FIRMDB_FULL|-106004|ERR_FIRMDB_NOT_FOUND|-107001|ERR_INVALID_CERTIFICATE|-107002|ERR_VALID_CERT_EXISTS|-107003|ERR_CERT_UNPACKING_FAILED|-107004|ERR_CERT_WRONG_PASSWORD|-107005|ERR_CERT_FILE_WRITE|-107006|ERR_CERT_FILE_READ|-107007|ERR_CERT_NOT_FOUND|-108000|ERR_R_DISC_NOTEXIST|-108001|ERR_DISCDB_DISCCODE_EXISTS|-108002|ERR_DISCDB_BARCODE_EXISTS|-108003|ERR_DISCDB_FULL|-108004|ERR_DISCDB_NOT_FOUND|-109981|ERR_SMARTCARD_NOT_PRESENT|-109982|ERR_SMARTCARD_CONFIG|-109983|ERR_SMARTCARD_COMMUNICATION|-109984|ERR_SMARTCARD_CARD_FAULT|-109985|ERR_SMARTCARD_MODULE_ERROR|-109986|ERR_SMARTCARD_WRONG_ID|-110100|ERR_DEVICE_COMM_ERROR|-110101|ERR_DEVICE_WRONG_STRUCT|-110102|ERR_DEVICE_STFLAG_ACTIVE|-110103|ERR_DEVICE_INVALID_DATA|-110104|ERR_DEVICE_NOT_FISCALIZED" & _
                                                    "|-110105|ERR_DEVICE_ALREADY_FISCALIZED|-110106|ERR_DEVICE_IN_SERVICE_MODE|-110107|ERR_DEVICE_PASSED_SERVICE_DATE|-110108|ERR_DEVICE_DAY_IS_OPEN|-110109|ERR_DEVICE_DAY_IS_CLOSED|-110110|ERR_DEVICE_WRONG_NUMBERS|-110111|ERR_DEVICE_ADMIN_ONLY|-110112|ERR_DEVICE_UNFISCALIZED|-110200|ERR_NAP_OPEN_SESSION|-110201|ERR_NAP_PREPARE_DATA|-110202|ERR_NAP_SEND_DATA|-110203|ERR_NAP_RECV_DATA|-110204|ERR_NAP_EMPTY_DATA|-110205|ERR_NAP_NEGATIVE_ANSWER|-110206|ERR_NAP_WRONG_ANSWER_FORMAT|-110207|ERR_NAP_HOSTDI_ZERRO|-110208|ERR_NAP_EXCEPTION|-110209|ERR_NAP_NOTPERSONALIZED|-110209|ERR_NAP_NOTREGISTERED|-110210|ERR_NAP_BLOCKED_72H|-110211|ERR_NAP_BLOCKED_NO_MODEM_LAN|-110212|ERR_NAP_BUSY|-110213|ERR_NAP_REGISTERED|-110214|ERR_NAP_WRONG_PSTYPE|-110215|ERR_NAP_DEREG_ON_SERVER|-110216|ERR_NAP_WRONG_IMSI|-110217|ERR_NAP_BLOCKED_MAX_ZERRORS|-110218|ERR_NAP_WRONG_FDTYPE|-110219|ERR_NAP_BLOCKED_BY_SERVER" & _
                                                    "|-110220|ERR_NAP_BLOCKED_ERROR_FROM_SERVER|-110221|ERR_NAP_NO_SERVER_ADDRESS|-110222|ERR_NAP_NO_REGISTRATIONS_POSSIBLE|-110223|ERR_NAP_INVALID_OPERATOR_INN|-110224|ERR_NAP_INVALID_SERVER_INN|-110225|ERR_NAP_BLOCKED_MAX_SELLERRORS|-110226|ERR_NAP_BLOCKED_24H|-110300|ERR_WORK_INVALID_FILE|-110301|ERR_WORK_INVALID_PARAM|-110500|ERR_MODEM_CTRL|-110501|ERR_MODEM_NO_SIM|-110502|ERR_MODEM_PIN|-110503|ERR_MODEM_ATTACH|-110504|ERR_MODEM_PPP|-110505|ERR_MODEM_CONFIG|-110506|ERR_MODEM_WAIT_INIT|-110507|ERR_MODEM_NOTREADY|-110508|ERR_MODEM_REMOVE_SIM|-110509|ERR_MODEM_CELL_FOUND|-110510|ERR_MODEM_CELL_NOTFOUND|-110511|ERR_MODEM_LOT_DAYS_FAIL|-110601|ERR_MODEM_CONNECT_AP|-110700|ERR_NET_DNS_RESOLVE|-110701|ERR_NET_SOCKET|-110702|ERR_NET_CONNECTION|-110703|ERR_NET_CONFIG|-110704|ERR_NET_SOCKET_CONNECTED|-110705|ERR_NET_SSL_ERROR|-110706|ERR_NET_HTTP_ERROR|-110800|ERR_DT_OK|-110801|ERR_DT_UNKNOWN_ID" & _
                                                    "|-110802|ERR_DT_INVALID_TOKEN|-110803|ERR_DT_PROTOCOL_ERROR|-110804|ERR_DT_UNKNOWN_COMMAND|-110805|ERR_DT_UNSUPPORTED_COMMAND|-110806|ERR_DT_INVALID_CONFIGURATION|-110807|ERR_DT_SSL_IS_NOT_ALLOWED|-110808|ERR_DT_INVALID_REQUEST_NUMBER|-110809|ERR_DT_INVALID_RETRY_REQUEST|-110810|ERR_DT_CANT_CANCEL_TICKET|-110811|ERR_DT_OPEN_SHIFT_TIMEOUT_EXPIRED|-110812|ERR_DT_INVALID_LOGIN_PASSWORD|-110813|ERR_DT_INCORRECT_REQUEST_DATA|-110814|ERR_DT_NOT_ENOUGH_CASH|-110815|ERR_DT_BLOCKED|-110854|ERR_DT_SERVICE_TEMPORARILY_UNAVAILABLE|-110855|ERR_DT_UNKNOWN_ERROR|-111000|ERR_R_CLEAR|-111001|ERR_R_NOCLEAR|-111002|ERR_R_SYNTAX|-111003|ERR_R_NPOSSIBLE|-111004|ERR_R_PLU_NOTEXIST|-111005|ERR_R_PLU_VAT_DISABLE|-111006|ERR_R_PLU_QTY_PRC|-111007|ERR_R_PLU_NO_PRC|-111008|ERR_R_PLU_GRP_RANGE|-111009|ERR_R_PLU_DEP_RANGE|-111010|ERR_R_BAR_NOTEXIST|-111011|ERR_R_OVF_TOTAL|-111012|ERR_R_OVF_QTY|-111013|ERR_R_ECR_OVR" & _
                                                    "|-111014|ERR_R_BILL_TL_OVR|-111015|ERR_R_OPEN_BON|-111016|ERR_R_CLOSED_BON|-111017|ERR_R_PAY_NOCASH|-111018|ERR_R_PAY_STARTED|-111019|ERR_R_OVF_TRZ_BUFF|-111020|ERR_R_NO_TRANSACTIONS|-111021|ERR_R_NEGATIVE_SUMVAT|-111022|ERR_R_PYFOREIGN_HAVERESTO|-111023|ERR_R_TRZ_NOT_EXIST|-111024|ERR_R_END_OF_24_HOUR_PERIOD|-111025|ERR_R_NO_VALID_INVOICE|-111026|ERR_R_POS_TERM_CANCELED|-111027|ERR_R_POS_TERM_APPROVED|-111028|ERR_R_POS_TERM_NOT_APPROVED|-111029|ERR_R_POS_TERM_CONN_ERR|-111030|ERR_R_PLU_QTY_PRC_TOO_LOW|-111031|ERR_R_VALUE_TOO_BIG|-111032|ERR_R_VALUE_BAD|-111033|ERR_R_PRICE_TOO_BIG|-111034|ERR_R_PRICE_BAD|-111035|ERR_R_ALL_VOID_SELECTED|-111036|ERR_R_ONLY_ALL_VOID_IS_POSSIBLE|-111040|ERR_R_REST_NOFREESPC_SELLS|-111041|ERR_R_REST_NOFREESPCFORNEWACNT|-111042|ERR_R_REST_ACCOUNT_IS_OPENED|-111043|ERR_R_REST_WRONG_INDEX|-111044|ERR_R_REST_ACNT_IS_NOTFOUND|-111045|ERR_R_REST_NOT_PERMITTED" & _
                                                    "|-111046|ERR_R_OPEN_NONFISCALBON|-111047|ERR_R_OPEN_FISCALBON|-111048|ERR_R_BUYERS_TIN_IS_ENTERED|-111049|ERR_R_BUYERS_TIN_IS_NOT_ENTERED|-111050|ERR_R_PAY_NOT_STARTED|-111051|ERR_R_BON_TYPE_MISMATCH|-111052|ERR_R_REACH_BON_TL_LIMIT|-111053|ERR_R_CASH_NO_MULT_MIN_COIN|-111054|ERR_R_PAY_BIG_AMOUNT|-111055|ERR_R_PAY_VOUCHER_NEED_INPUT_SUM|-111056|ERR_R_PAY_VOUCHER_NEED_SURCHARGE|-111057|ERR_R_PAY_FOREIGN_DISABLED|-111058|ERR_R_PAY_FOREIGN_IMPOSSIBLE|-111059|ERR_R_PAY_FOREIGN_SMALL_AMOUNT|-111060|ERR_R_SAFE_OPEN_DISABLED|-111061|ERR_R_PAY_FORBIDDEN|-111062|ERR_R_PERC_KEY_FORBIDDEN|-111063|ERR_R_AMOUNT_BIGGER_BILLAMOUNT|-111064|ERR_R_AMOUNT_SMALLER_BILLAMOUNT|-111065|ERR_R_ZERO_BILLAMOUNT|-111066|ERR_R_ALL_VOID_EXECUTED|-111067|ERR_R_OPEN_STORNOBON|-111068|ERR_R_PAY_ZERO_AMOUNT|-111069|ERR_R_PLU_PRICETYPE_RANGE|-111070|ERR_R_PLU_PRICETYPE_LINKED|-111071|ERR_R_PLU_PRICETYPE_NEGATIVE" & _
                                                    "|-111072|ERR_R_MORE_THAN_ONE_VAT|-111073|ERR_R_PINPAD|-111074|ERR_R_WRONG_BUYERS_DATA|-111075|ERR_R_VAT_SYSTEM_DISABLE|-111076|ERR_R_OPER_NOT_LOGGED_IN|-111077|ERR_R_WRONG_DATE_FM|-111078|ERR_R_CORR_DATA_NOT_ENTERED|-111079|ERR_R_FRACTIONAL_QTY|-111080|ERR_R_OUT_OF_STOCK|-111081|ERR_R_STL_NEEDED|-111082|ERR_R_PACK_NOTEXIST|-111083|ERR_R_PLU_UNIT_NOTEXIST|-111084|ERR_R_PLU_CATEGORY_NOTEXIST|-111085|ERR_R_DEP_WRONG_NAME|-111086|ERR_R_BANK_TERM_NOT_CONFIGURED|-111087|ERR_R_SIGN_PAY_INCORECT|-111088|ERR_R_SIGN_INCORRECT|-111089|ERR_R_PLU_OVER_MAX_PRC|-111090|ERR_R_PLU_FIX_PRC|-111091|ERR_R_SIGN_AGENT_INCORECT|-111092|ERR_R_PAY_VOUCHER_RESTO|-111093|ERR_R_PAY_ADVANCE_BIG|-111094|ERR_R_PAY_STORNO_RESTO|-111095|ERR_R_NOT_EXCISE_PLU_WITH_EXCISE_STAMP|-111096|ERR_R_EXCISE_PLU_WITHOUT_EXCISE_STAMP|-111097|ERR_R_EXCISE_PLU_FORBIDDEN|-111500|ERR_PINPAD_NONE|-111501|ERR_PINPAD_GENERAL" & _
                                                    "|-111502|ERR_PINPAD_INVALID_COMMAND|-111503|ERR_PINPAD_INVALID_PARAM|-111504|ERR_PINPAD_INVALID_ADDRESS|-111505|ERR_PINPAD_INVALID_VALUE|-111506|ERR_PINPAD_INVALID_LENGTH|-111507|ERR_PINPAD_NOT_PERMIT|-111508|ERR_PINPAD_NO_DATA|-111509|ERR_PINPAD_TIMEOUT|-111510|ERR_PINPAD_INVALID_KEY_NUMBER|-111511|ERR_PINPAD_INVALID_KEY_ATTRIBUTES|-111512|ERR_PINPAD_INVALID_DEVICE|-111513|ERR_PINPAD_NOT_SUPPORT|-111514|ERR_PINPAD_PIN_LIMIT|-111515|ERR_PINPAD_FLASH|-111516|ERR_PINPAD_HARDWARE|-111517|ERR_PINPAD_INVALID_CRC|-111518|ERR_PINPAD_CANCEL|-111519|ERR_PINPAD_INVALID_SIGNATURE|-111520|ERR_PINPAD_INVALID_HEADER|-111521|ERR_PINPAD_INVALID_PASSWORD|-111522|ERR_PINPAD_INVALID_KEY_FORMAT|-111523|ERR_PINPAD_SCR|-111524|ERR_PINPAD_HAL|-111525|ERR_PINPAD_INVALID_KEY|-111526|ERR_PINPAD_NO_PIN_DATA|-111527|ERR_PINPAD_INVALID_REMINDER|-111528|ERR_PINPAD_NOT_INIT|-111529|ERR_PINPAD_LIMIT" & _
                                                    "|-111530|ERR_PINPAD_INVALID_SEQUENCE|-111531|ERR_PINPAD_NO_PERMITION|-111532|ERR_PINPAD_NO_TMK|-111533|ERR_PINPAD_INVALID_KEK|-111534|ERR_PINPAD_DUPLICATE_KEY|-111535|ERR_PINPAD_KEYBOARD|-111536|ERR_PINPAD_KEYBOARD_NOT_CALIBRATED|-111537|ERR_PINPAD_KEYBOARD_FAILED|-111538|ERR_PINPAD_DEVICE_BUSY|-111539|ERR_PINPAD_TAMPERED|-111540|ERR_PINPAD_EMSR|-111541|ERR_PINPAD_ACCEPT|-111542|ERR_PINPAD_INVALID_PAN|-111543|ERR_PINPAD_NOT_ENOUGH_MEMORY|-111544|ERR_PINPAD_EMV|-111545|ERR_PINPAD_CRYPTOGRAPHY|-111546|ERR_PINPAD_COMMUNICATION|-111547|ERR_PINPAD_INVALID_VERSION|-111548|ERR_PINPAD_NOPAPER|-111549|ERR_PINPAD_OVERHEATED|-111550|ERR_PINPAD_NOT_CONNECTED|-111551|ERR_PINPAD_USE_CHIP|-111552|ERR_PINPAD_END_DAY|-111554|ERR_PINPAD_BOR_ERR|-111555|ERR_PINPAD_NO_CONN|-111556|ERR_PINPAD_ECR|-111557|ERR_PINPAD_NOT_CONF|-111558|ERR_PINPAD_SAME_TRANS|-111559|ERR_PINPAD_RECEIPT|-111560|ERR_PINPAD_FP_TRANS" & _
                                                    "|-111561|ERR_PINPAD_NOT_CONF_TYPE|-111700|ERR_PINPAD_INV_AMOUNT|-111701|ERR_PINPAD_TRN_NOT_FOUND|-111702|ERR_PINPAD_FILE_EMPTY|-111703|ERR_PINPAD_MAX_CASHBACK|-111800|ERR_SCALE_NOT_RESPOND|-111801|ERR_SCALE_NOT_CALCULATED|-111802|ERR_SCALE_WRONG_RESPONSE|-111803|ERR_SCALE_ZERO_WEIGHT|-111804|ERR_SCALE_NEGATIVE_WEIGHT|-111805|ERR_SCALE_T_WRONG_INTF|-111806|ERR_SCALE_T_CONNECT|-111807|ERR_SCALE_SEND|-111808|ERR_SCALE_RECEIVE|-111809|ERR_SCALE_FILE_GENERATE|-111810|ERR_SCALE_NOT_CONFIG|-111900|ERR_NTP_NO_COMM|-111901|ERR_NTP_EARLIER_DATETIME|-111902|ERR_NTP_WRONG_IP|-112000|ERR_FP_INVALID_COMMAND|-112001|ERR_FP_INVALID_SYNTAX|-112002|ERR_FP_COMMAND_NOT_PERMITTED|-112003|ERR_FP_OVERFLOW|-112004|ERR_FP_WRONG_DATE_TIME|-112005|ERR_FP_NEEDED_MODE_PC|-112006|ERR_FP_NO_PAPER|-112007|ERR_FP_COVER_IS_OPEN|-112008|ERR_FP_PRINTER_FAILURE|-112100|_ERR_FP_SYNTAX_PARAM_BEGIN|-112101|ERR_FP_SYNTAX_PARAM_1" & _
                                                    "|-112102|ERR_FP_SYNTAX_PARAM_2|-112103|ERR_FP_SYNTAX_PARAM_3|-112104|ERR_FP_SYNTAX_PARAM_4|-112105|ERR_FP_SYNTAX_PARAM_5|-112106|ERR_FP_SYNTAX_PARAM_6|-112107|ERR_FP_SYNTAX_PARAM_7|-112108|ERR_FP_SYNTAX_PARAM_8|-112109|ERR_FP_SYNTAX_PARAM_9|-112110|ERR_FP_SYNTAX_PARAM_10|-112111|ERR_FP_SYNTAX_PARAM_11|-112112|ERR_FP_SYNTAX_PARAM_12|-112113|ERR_FP_SYNTAX_PARAM_13|-112114|ERR_FP_SYNTAX_PARAM_14|-112115|ERR_FP_SYNTAX_PARAM_15|-112116|ERR_FP_SYNTAX_PARAM_16|-112199|_ERR_FP_SYNTAX_PARAM_END|-112200|_ERR_FP_BAD_PARAM_BEGIN|-112201|ERR_FP_BAD_PARAM_1|-112202|ERR_FP_BAD_PARAM_2|-112203|ERR_FP_BAD_PARAM_3|-112204|ERR_FP_BAD_PARAM_4|-112205|ERR_FP_BAD_PARAM_5|-112206|ERR_FP_BAD_PARAM_6|-112207|ERR_FP_BAD_PARAM_7|-112208|ERR_FP_BAD_PARAM_8|-112209|ERR_FP_BAD_PARAM_9|-112210|ERR_FP_BAD_PARAM_10|-112211|ERR_FP_BAD_PARAM_11|-112212|ERR_FP_BAD_PARAM_12|-112213|ERR_FP_BAD_PARAM_13" & _
                                                    "|-112214|ERR_FP_BAD_PARAM_14|-112215|ERR_FP_BAD_PARAM_15|-112216|ERR_FP_BAD_PARAM_16|-112299|_ERR_FP_BAD_PARAM_END|-112900|_ERR_RANGE_EM_BEGIN|-112999|_ERR_RANGE_EM_END|-113000|ERR_FLASH_WRONG_ID|-113001|ERR_FLASH_WRONG_SIZE|-114000|ERR_POS_TERM_CHAN_CLOSED|-118000|ERR_ECRSRV_NO_SOCKET_OPENED|-118001|ERR_ECRSRV_SET_IS_NOT_TAKEN|-118002|ERR_ECRSRV_WRONG_PARAM|-118003|ERR_ECRSRV_NOT_SEND|-118004|ERR_ECRSRV_RECV_TMOUT|-118005|ERR_ECRSRV_SOCKET_CLOSED|-118006|ERR_ECRSRV_UNKNOWN|-118007|ERR_ECRSRV_FORBIDEN|-120000|ERR_PGM_NAME_NOT_UNIQUE|-120001|ERR_PGM_OPER_PASS_NOT_UNIQUE|-120002|ERR_PGM_DATETIME_OUT_OF_RANGE_MIN|-120003|ERR_PGM_DATETIME_OUT_OF_RANGE_MAX|-121000|ERR_SCANNER_GENERAL|-121001|ERR_COURIER_EIK_INVALID|-170000|ERR_USB_HOST_INIT|-170001|ERR_USB_NO_DEVICE|-170002|ERR_USB_NO_FILESYSTEM|-170003|ERR_USB_FILE_OPEN|-170004|ERR_USB_FILE_COPY|-170005|ERR_FILE_UNPACK" & _
                                                    "|-171000|ERR_RENTAL_NOT_FOUND|-171001|ERR_RENTAL_FULL|-171002|ERR_RENTAL_POSITION_OCCUPIED|-171003|ERR_RENTAL_POSITION_FREE|-171004|ERR_RENTAL_SUBSCRIPTION_ACTIVE"
Private Const DEF_SEP                   As String = vbTab
Private Const DEF_AUTODETECTTIMEOUT     As Long = 52
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 100
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_CONNECTIONTIMEOUT     As Long = 3000
Private Const DEF_TIMEOUT               As Long = 15000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_PASSWORD              As String = "0000"
Private Const DEF_SOCKET_PORT           As Long = 4999
Private Const DEF_ROW_CHARS             As Long = 42
Private Const DEF_ITEM_CHARS            As Long = 72
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_CODEPAGE              As Long = 1251
Private Const DEF_MAX_CGINFO            As String = "36|36|36|36|36|16|16"
Private Const DEF_MAX_OPERATOR_NAME     As Long = 32
Private Const DEF_SEQUENCE              As Long = &H20
Private Const LNG_MAX_SEQUENCE          As Long = &H7F
Private Const LNG_OPEN_IMPULSE          As Long = 250
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_REPORT        As String = "dd\-mm\-yy"
Private Const ERR_OPER_WRONG_PASSWORD   As Long = -102002
Private Const MAX_HEADER_LINES          As Long = 10
Private Const MAX_FOOTER_LINES          As Long = 10

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lConnectionTimeout        As Long
Private m_lCommandTimeout           As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsDatecsStatusBitsEnum
Private m_eMemory                   As UcsDatecsFiscalMemoryEnum
Private m_vDeviceNumbers            As Variant
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues

Private Enum UcsInternalErrors
    [__ucsErrTimeoutWaitingForResponse]
    ucsErrInvalidResponseLength
    ucsErrInvalidResponseChecksum
    ucsErrResponseMissingPostamble
    ucsErrResponseMissingPreamble
    ucsErrNoReceiptStarted
    ucsErrInvalidOperatorPassword
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
    ucsErrInvalidReportType
    ucsErrUnknownError
    ucsErrUnknownMemory
    ucsErrUnsupportedPaymentType
    ucsErrCannotPrint
    ucsErrWrongInvoiceNo
    ucsErrUnsupportedDiscType
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtSurcharge
    ucsTxtDiscount
    ucsTxtSubTotal
    ucsTxtPluSales
End Enum

Private Enum UcsProtocolMarkers
    SYM_PREAMBLE = &H1
    SYM_TERMINATOR = &H3
    SYM_SEPARATOR = &H4
    SYM_POSTAMBLE = &H5
    SYM_NAK = &H15
    SYM_SYN = &H16
    SYM_BASEDATA = &H20
End Enum

Private Enum UcsMaxCgInfoEnum
    ucsMciOperatorName
    ucsMciCgPrsReceive
    ucsMciCgName
    ucsMciCgCity
    ucsMciCgAddress
    ucsMciCgTaxNo
    ucsMciCgVatNo
    [_ucsMciLast]
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WaitOver            As OVERLAPPED
    EventMask           As Long
End Type

Private Type UcsConfigValues
    DefaultPassword     As String
    CodePage            As Long
    IsEcr               As VbTriState
    MaxCgInfo(0 To [_ucsMciLast] - 1) As Long
    MaxOperatorName     As Long
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    Logger.Log vbLogEventTypeDebug, MODULE_NAME, sFunction, sText
End Sub

Private Sub DebugDataDump(sFunction As String, sPrefix As String, sData As String)
    Logger.DataDump MODULE_NAME, sFunction, sPrefix, sData
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get ConnectionTimeout() As Long
    ConnectionTimeout = m_lConnectionTimeout
End Property

Property Let ConnectionTimeout(ByVal lValue As Long)
    m_lConnectionTimeout = lValue
End Property

Property Get CommandTimeout() As Long
    CommandTimeout = m_lCommandTimeout
End Property

Property Let CommandTimeout(ByVal lValue As Long)
    m_lCommandTimeout = lValue
End Property

Property Get Status(ByVal eValue As UcsDatecsStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterErrors))
End Property

Property Get Memory(ByVal eValue As UcsDatecsFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = pvGetFlagsText(m_eMemory, m_uConfig.LocalizedText(ucsFscLciMemoryTexts))
End Property

Property Get LastErrNo() As UcsFiscalErrorsEnum
    LastErrNo = m_uPrintData.LastErrNo
End Property

Property Get LastError() As String
    LastError = m_uPrintData.LastError
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get IsConnected() As Boolean
    IsConnected = Not m_oConnector Is Nothing
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = SplitOrReindex(sValue, "|")
    Select Case eIdx
    Case ucsFscLciInternalErrors
        m_uPrintData.LocalizedText.ErrNoReceiptStarted = pvInternal(ucsErrNoReceiptStarted)
    Case ucsFscLciReceiptTexts
        m_uPrintData.LocalizedText.TxtSurcharge = pvText(ucsTxtSurcharge)
        m_uPrintData.LocalizedText.TxtDiscount = pvText(ucsTxtDiscount)
        m_uPrintData.LocalizedText.TxtPluSales = pvText(ucsTxtPluSales)
    End Select
End Property

Property Get IsEcr() As Boolean
    If m_uConfig.IsEcr = vbUseDefault Then
        m_uConfig.IsEcr = IIf(Left$(GetDeviceModel(), 1) <> "F", vbTrue, vbFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = vbTrue)
End Property

'= private ===============================================================

Private Property Get pvRowChars() As Long
    Const FUNC_NAME     As String = "pvRowChars"
    
    On Error GoTo EH
    If m_uPrintData.Config.RowChars = 0 Then
        m_uPrintData.Config.RowChars = C_Lng(pvSettings("PrintColumns"))
        If m_uPrintData.Config.RowChars <= 0 Then
            m_uPrintData.Config.RowChars = DEF_ROW_CHARS
        End If
    End If
    pvRowChars = m_uPrintData.Config.RowChars
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvCommentChars() As Long
    pvCommentChars = m_uPrintData.Config.CommentChars
    If pvCommentChars = 0 Then
        pvCommentChars = pvRowChars - 2
    End If
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uPrintData.Config.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = DEF_ITEM_CHARS
    End If
End Property

Private Property Get pvSettings(sName As String, Optional ByVal Index As Long = -1) As String
    Const CONF_NAME     As String = "Settings"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    pvGetCommandConfig CONF_NAME, "InfoProgramming", eCmd, ucsDcsCmdInfoProgramming, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & DEF_SEP, _
        "Index", lIndex, 1
    sParam = Printf(sParam, sName, IIf(Index >= 0, Index, vbNullString))
    pvSettings = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
End Property

Private Property Let pvSettings(sName As String, Optional ByVal Index As Long = -1, sValue As String)
    Const CONF_NAME     As String = "Settings"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    pvGetCommandConfig CONF_NAME, "InitProgramming", eCmd, ucsDcsCmdInfoProgramming, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP, _
        "Index", lIndex, 1
    sParam = Printf(sParam, sName, IIf(Index >= 0, Index, vbNullString), sValue)
    SendCommand eCmd, sParam
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim lIdx            As Long
    Dim baProbe()       As Byte
    Dim sReplyText      As String
    Dim lTimeout        As Long
    Dim lDeviceInfoTimeout As Long
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim aPorts()        As UcsAutodetect
    Dim hEvents()       As Long
    Dim lEventsCount    As Long
    Dim lResult         As Long
    Dim sProtocol       As String
    Dim sModel          As String
    Dim sFirmware       As String
    Dim sDeviceSerialNo As String
    Dim sFiscalMemoryNo As String
    Dim vSpeeds         As Variant
    Dim sDefProbe       As String
    Dim baBuffer()      As Byte
    Dim sBuffer         As String
    
    On Error GoTo EH
    vRetVal = vPorts
    If Not IsArray(vPorts) Then
        GoTo QH
    End If
    If UBound(vPorts) < 0 Then
        GoTo QH
    End If
    sDefProbe = Chr$(SYM_PREAMBLE) & Chr$(&H30) & Chr$(&H30) & Chr$(&H32) & Chr$(&H3A) & Chr$(&H21) & Chr$(&H30) & Chr$(&H30) & Chr$(&H34) & Chr$(&H3A) & Chr$(SYM_POSTAMBLE) & Chr$(&H30) & Chr$(&H31) & Chr$(&H3C) & Chr$(&H30) & Chr$(SYM_TERMINATOR)
    baProbe = ToAscii(GetConfigValue(STR_PROTOCOL_DATECS_X, "ProbeText", sDefProbe), m_uConfig.CodePage)
    sReplyText = GetConfigValue(STR_PROTOCOL_DATECS_X, "ReplyText", Chr$(SYM_PREAMBLE)) ' PREAMBLE (&H1)
    lTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_X, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_X, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_DATECS_X, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
    If IsObject(vSpeeds) Then
        vSpeeds = vSpeeds.Items()
    End If
    For Each vSpeed In vSpeeds
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        lEventsCount = 0
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not IsArray(vRetVal(lIdx)) And LenB(At(vPorts, lIdx)) <> 0 Then
                    Set .Connector = InitDeviceConnector(vRetVal(lIdx) & "," & vSpeed, lTimeout, LocalizedText(ucsFscLciConnectorErrors))
                    If Not .Connector Is Nothing Then
                        .Connector.WriteData baProbe, 1
                        .WaitOver.hEvent = CreateEvent(0, 1, 0, 0)
                        Call SetCommMask(.Connector.hComm, EV_RXCHAR Or EV_ERR)
                        lResult = WaitCommEvent(.Connector.hComm, .EventMask, .WaitOver)
                        If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                            hEvents(lEventsCount) = .WaitOver.hEvent
                            lEventsCount = lEventsCount + 1
                        End If
                    End If
                End If
            End With
        Next
        If lEventsCount > 0 Then
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing And .EventMask <> 0 Then
                    If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                        sBuffer = Replace(Replace(StrConv(baBuffer, vbUnicode), Chr$(SYM_NAK), vbNullString), Chr$(SYM_SYN), vbNullString)
                        If LenB(sBuffer) = 0 Then
                            If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                                sBuffer = Replace(Replace(StrConv(baBuffer, vbUnicode), Chr$(SYM_NAK), vbNullString), Chr$(SYM_SYN), vbNullString)
                            End If
                        End If
                        If Left$(sBuffer, Len(sReplyText)) = sReplyText Then
                            DebugLog FUNC_NAME, "Auto-detecting " & .Connector.Device & " reply 0x" & ToHexDump(sBuffer)
                            If pvGetDeviceInfo(.Connector, lDeviceInfoTimeout, Empty, sProtocol, sModel, sFirmware, sDeviceSerialNo, sFiscalMemoryNo) Then
                                vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel, sFirmware)
                                DebugLog FUNC_NAME, "Auto-detected vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", sProtocol=" & sProtocol & ", sModel=" & sModel & ", sFirmware=" & sFirmware & ", sDeviceSerialNo=" & sDeviceSerialNo & ", sFiscalMemoryNo=" & sFiscalMemoryNo
                            End If
                        Else
                            DebugLog FUNC_NAME, "Discarding vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", " & .Connector.Device & " reply 0x" & ToHexDump(sBuffer)
                        End If
                    End If
                    Set .Connector = Nothing
                    If .WaitOver.hEvent <> 0 Then
                        Call CloseHandle(.WaitOver.hEvent)
                    End If
                End If
            End With
        Next
    Next
QH:
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    Dim sDeviceSerialNo As String
    Dim sFiscalMemoryNo As String
    Dim sError          As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    pvSetLastError vbNullString
    Set oOptions = ParseDeviceString(sDeviceString)
    m_lConnectionTimeout = C_Lng(JsonItem(oOptions, "ConnectionTimeout"))
    If m_lConnectionTimeout <= 0 Then
        m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    End If
    m_lCommandTimeout = C_Lng(JsonItem(oOptions, "Timeout"))
    If m_lCommandTimeout <= 0 Then
        m_lCommandTimeout = DEF_TIMEOUT
    End If
    m_sDevice = ToConnectorDevice(oOptions, DEF_SOCKET_PORT, Me)
    If Not Connect() Then
        GoTo QH
    End If
    sDeviceSerialNo = ToDeviceSerialNo(JsonItem(oOptions, "DeviceSerialNo"))
    If LenB(sDeviceSerialNo) = 0 Then
        If Not pvGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, vbNullString, vbNullString, vbNullString, sDeviceSerialNo, sFiscalMemoryNo, Error:=sError) Then
            pvSetLastError sError
            GoTo QH
        End If
        JsonItem(oOptions, "DeviceSerialNo") = sDeviceSerialNo
        JsonItem(oOptions, "FiscalMemoryNo") = sFiscalMemoryNo
    Else
        m_vDeviceNumbers = sDeviceSerialNo & "|" & Trim$(JsonItem(oOptions, "FiscalMemoryNo"))
    End If
    sDeviceString = ToDeviceString(oOptions)
    '--- load config values
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sDeviceSerialNo, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .CommentChars = GetConfigNumber(sDeviceSerialNo, "CommentChars", C_Lng(JsonItem(oOptions, "CommentChars")))
        .ItemChars = GetConfigNumber(sDeviceSerialNo, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sDeviceSerialNo, "AbsoluteDiscount", JsonBoolItem(oOptions, "AbsoluteDiscount", True)))
        .NegativePrices = C_Bool(GetConfigValue(sDeviceSerialNo, "NegativePrices", JsonBoolItem(oOptions, "NegativePrices", True)))
        .MinDiscount = GetConfigNumber(sDeviceSerialNo, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sDeviceSerialNo, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sDeviceSerialNo, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
    End With
    With m_uConfig
        .DefaultPassword = GetConfigValue(sDeviceSerialNo, "DefaultPassword", JsonItem(oOptions, "DefaultPassword"))
        .CodePage = GetConfigNumber(sDeviceSerialNo, "CodePage", Znl(C_Lng(JsonItem(oOptions, "CodePage")), DEF_CODEPAGE))
        .IsEcr = GetConfigNumber(sDeviceSerialNo, "IsEcr", Znl(C_Lng(JsonItem(oOptions, "IsEcr")), vbUseDefault))
        vSplit = Split(DEF_MAX_CGINFO, "|")
        For lIdx = LBound(.MaxCgInfo) To UBound(.MaxCgInfo)
            .MaxCgInfo(lIdx) = GetConfigNumber(sDeviceSerialNo, "MaxCgInfo" & lIdx, C_Lng(vSplit(lIdx)))
        Next
        .MaxOperatorName = GetConfigNumber(sDeviceSerialNo, "MaxOperatorName", Znl(C_Lng(JsonItem(oOptions, "MaxOperatorName")), DEF_MAX_OPERATOR_NAME))
        Set .ConfigCommands = GetConfigCollection(sDeviceSerialNo, "Commands")
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "m_sDevice=" & m_sDevice & ", m_lConnectionTimeout=" & m_lConnectionTimeout
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lConnectionTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
    If m_oConnector Is Nothing Then
        DebugLog FUNC_NAME, "InitDeviceConnector failed, m_uPrintData.LastError=" & m_uPrintData.LastError
        GoTo QH
    End If
    '--- success
    Connect = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Disconnect() As Boolean
    Const FUNC_NAME     As String = "Disconnect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "TypeOf m_oConnector=" & TypeName(m_oConnector)
    Set m_oConnector = Nothing
    '--- success
    Disconnect = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsDatecsCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Const ERR_OPER_WRONG_PASSWORD As Long = -102002
    Dim sError          As String
    Dim lPos            As Long
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsDcsCmdNoop Then
        Exit Function
    End If
    SendCommand = pvTransmitCommand(m_oConnector, m_lCommandTimeout, eCmd, Data, sError)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    If Status(ucsDcsStbGeneralError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
        GoTo QH
    End If
    If Memory(ucsDcsFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory)))
        GoTo QH
    End If
    lPos = InStr(SendCommand, DEF_SEP)
    If lPos > 1 Then
        sError = Left$(SendCommand, lPos - 1)
        Select Case C_Lng(sError)
        Case 0
            '--- do nothing
        Case ERR_OPER_WRONG_PASSWORD
            pvSetLastError pvToErrorCode(sError), ucsFerInvalidPassword
        Case Else
            pvSetLastError pvToErrorCode(sError)
        End Select
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDuplicate( _
            Optional DuplicateNo As String, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDuplicate"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    If LenB(DuplicateNo) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsDcsCmdPrintDuplicate, sParam, vbNullString
    Else
        pvGetCommandConfig CONF_NAME, "InfoEJ" & sParam, eCmd, ucsDcsCmdInfoEJ, sParam, "3" & DEF_SEP & "%1" & DEF_SEP & "0" & DEF_SEP
        sParam = Printf(sParam, DuplicateNo)
    End If
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        pvSetLastError pvInternal(ucsErrCannotPrint)
        GoTo QH
    End If
    '--- success
    PrintDuplicate = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDailyXReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDailyXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintDailyReport, sParam, "X" & DEF_SEP
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintDailyReport, sParam, "G" & DEF_SEP
    Case ucsFscRstDailyByDepartments
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintDailyReport, sParam, "D" & DEF_SEP
    Case ucsFscRstDailyByOperators
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintReportByOperators, sParam, DEF_SEP & DEF_SEP & "0" & DEF_SEP
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        pvSetLastError pvInternal(ucsErrCannotPrint)
        GoTo QH
    End If
    '--- success
    PrintDailyXReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDailyZReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDailyZReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintDailyReport, sParam, "Z" & DEF_SEP
    Case ucsFscRstDailyByOperators
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsDcsCmdPrintReportByOperators, sParam, DEF_SEP & DEF_SEP & "1" & DEF_SEP
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        pvSetLastError pvInternal(ucsErrCannotPrint)
        GoTo QH
    End If
    '--- success
    PrintDailyZReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintMonthlyReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            ByVal StartDate As Date, _
            ByVal EndDate As Date, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintMonthlyReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sFormat         As String
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstMonthlyShort
        '--- {Type}<SEP>{Start}<SEP>{End}<SEP>
        sParam = "0" & DEF_SEP & "%1" & DEF_SEP & "%2" & DEF_SEP
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsDcsCmdPrintReportByDateDetailed, sParam, sParam, _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case ucsFscRstMonthlyDetailed
        sParam = "1" & DEF_SEP & "%1" & DEF_SEP & "%2" & DEF_SEP
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsDcsCmdPrintReportByDateDetailed, sParam, sParam, _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        pvSetLastError pvInternal(ucsErrCannotPrint)
        GoTo QH
    End If
    '--- success
    PrintMonthlyReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintServiceDeposit( _
            ByVal Amount As Double, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date, _
            Optional ReceiptAmount As Double) As Variant
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal(0 To 2) As Double
    Dim lIdx            As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim sPrevReceiptNo  As String
    Dim bPrint          As Boolean
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    bPrint = Abs(Amount) > DBL_EPSILON
    If bPrint Then
        sPrevReceiptNo = GetLastReceiptNo()
        ReceiptDate = GetClock()
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
    End If
    '--- %1={Type}<SEP>%2={Amount}<SEP>
    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP
    pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit" & -bPrint, eCmd, ucsDcsCmdFiscalServiceDeposit, sParam, sParam, _
        "Format", sFormat, FORMAT_BASE_2
    sParam = Printf(sParam, IIf(Amount > -DBL_EPSILON, "0", "1"), SafeFormat(Abs(Amount), sFormat))
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vSplit) <= 0 Then
        GoTo QH
    End If
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    If bPrint Then
        ReceiptNo = GetLastReceiptNo()
        ReceiptAmount = Amount
        If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
            GoTo QH
        End If
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = ParseSum(At(vSplit, lIdx + 1))
    Next
    '--- idiosyncrasy: negative total on service withdraws
    vRetVal(2) = -vRetVal(2)
    '--- success
    PrintServiceDeposit = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsDcsCmdInfoDateTime, sParam, vbNullString, _
        "Index", lIndex, 1
    GetClock = ParseExtendedDate(Replace(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex), "DST", vbNullString))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean
    Const FUNC_NAME     As String = "SetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsDcsCmdInitDateTime, sParam, "%1" & DEF_SEP, _
        "Format", sFormat, FORMAT_DATE_EXTENDED
    SendCommand eCmd, Printf(sParam, Format$(NewDate, sFormat))
    '--- success
    SetClock = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNo() As String
    Const FUNC_NAME     As String = "GetLastReceiptNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim vSplit          As Variant
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDevice", eCmd, ucsDcsCmdInfoTransaction, sParam, vbNullString, _
        "Index", lIndex, 2
    vSplit = Split(pvTransmitCommand(m_oConnector, m_lCommandTimeout, eCmd, sParam), DEF_SEP)
    GetLastReceiptNo = At(vSplit, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNo() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsDcsCmdInitInvoiceNo, sParam, vbNullString, _
        "Index", lIndex, 3
    GetLastInvoiceNo = At(Split(pvTransmitCommand(m_oConnector, m_lCommandTimeout, eCmd, sParam), DEF_SEP), lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    Const FUNC_NAME     As String = "GetLastQRCodeInfo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lDateIndex      As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsDcsCmdInfoStatus, sParam, "0" & DEF_SEP, _
        "Index", lIndex, 5, _
        "DateIndex", lDateIndex, 6
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vSplit) <= 0 Then
        GoTo QH
    End If
    LastReceiptDate = ParseExtendedDate(At(vSplit, lDateIndex))
    '--- success
    GetLastQRCodeInfo = RTrim$(At(vSplit, lIndex))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceSerialNo() As String
    Const FUNC_NAME     As String = "GetDeviceSerialNo"
    
    On Error GoTo EH
    If Not IsArray(m_vDeviceNumbers) Then
        GetDeviceSerialNo = At(Split(C_Str(m_vDeviceNumbers), "|"), 0)
    End If
    If LenB(GetDeviceSerialNo) = 0 Then
        pvGetDeviceInfo m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, vbNullString, vbNullString, vbNullString, GetDeviceSerialNo, vbNullString
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalMemoryNo() As String
    Const FUNC_NAME     As String = "GetFiscalMemoryNo"
    
    On Error GoTo EH
    If Not IsArray(m_vDeviceNumbers) Then
        GetFiscalMemoryNo = At(Split(C_Str(m_vDeviceNumbers), "|"), 1)
    End If
    If LenB(GetFiscalMemoryNo) = 0 Then
        pvGetDeviceInfo m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, vbNullString, vbNullString, vbNullString, vbNullString, GetFiscalMemoryNo
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, GetDeviceProtocol, vbNullString, vbNullString, vbNullString, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, vbNullString, GetDeviceModel, vbNullString, vbNullString, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceFirmware() As String
    Const FUNC_NAME     As String = "GetDeviceFirmware"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lCommandTimeout, m_vDeviceNumbers, vbNullString, vbNullString, GetDeviceFirmware, vbNullString, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNo() As String
    Const FUNC_NAME     As String = "GetTaxNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "TAXnumber"
    GetTaxNo = RTrim$(pvSettings(sParam))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "TAXlabel"
    GetTaxCaption = RTrim$(pvSettings(sParam))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_HEADER_LINES Then
        pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "Header"
        GetHeaderText = RTrim$(pvStripBold(pvSettings(sParam, Index - 1)))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_FOOTER_LINES Then
        pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "Footer"
        GetFooterText = RTrim$(pvStripBold(pvSettings(sParam, Index - 1)))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    
    On Error GoTo EH
    If PmtType >= ucsFscPmtCash And PmtType <= ucsFscPmtCash Then
        GetPaymentName = pvText(ucsTxtPaymentCash + PmtType - 1)
    Else
        GetPaymentName = pvToDevicePmtName(PmtType)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo QH
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsDcsCmdInitOpenDrawer, sParam, "%1" & DEF_SEP
    SendCommand eCmd, Printf(sParam, LNG_OPEN_IMPULSE)
    '--- success
    OpenDrawer = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eTransCmd       As UcsDatecsCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lStatusClosed   As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsDcsCmdInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 1, _
        "StatusClosed", lStatusClosed, 0
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lStatusClosed Then
        '--- nothing to do
        CancelReceipt = True
        GoTo QH
    End If
    If Status(ucsDcsStbFiscalReceiptOpen) Then
        pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsDcsCmdFiscalCancel, sParam, vbNullString
        SendCommand eCmd, sParam
        If Status(ucsDcsStbFiscalReceiptOpen) Then
            pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsDcsCmdFiscalPayment, sParam, DEF_SEP & DEF_SEP & DEF_SEP
            SendCommand eCmd, sParam
            '--- {Seller}<SEP>{Receiver}<SEP>{Buyer}<SEP>{Address1}<SEP>{Address2}<SEP>{TypeTAXN}<SEP>{TAXN}<SEP>{VATN}<SEP>
            pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsDcsCmdFiscalCgInfo, sParam, DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & "0" & DEF_SEP & "0000000000" & DEF_SEP & DEF_SEP
            SendCommand eCmd, sParam
            pvGetCommandConfig CONF_NAME, "FiscalClose", eCmd, ucsDcsCmdFiscalClose, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
    Else
        pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsDcsCmdNonFiscalClose, sParam, vbNullString
        SendCommand eCmd, sParam
    End If
    '--- success (or failure)
    CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lStatusClosed
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    Const FUNC_NAME     As String = "GetDeviceStatus"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsDcsCmdInfoStatus, sParam, vbNullString
    pvTransmitCommand m_oConnector, m_lCommandTimeout, eCmd, sParam, m_uPrintData.LastError
    If LenB(m_uPrintData.LastError) <> 0 Then
        '--- do nothing
    ElseIf Status(ucsDcsStbGeneralError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
    ElseIf Memory(ucsDcsFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory)))
    End If
    StatusText = m_uPrintData.LastError
    pvSetLastError vbNullString
    GetDeviceStatus = IIf(LenB(StatusText) <> 0, ucsFscStaFailed, ucsFscStaReady)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPaymentTypes() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal(0 To MAX_PMT_TYPE - MIN_PMT_TYPE) As Variant
    Dim lIdx            As Long
    Dim lPmtIndex       As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments0", eCmd, ucsDcsCmdInfoTotalsByPayments, sParam, "0" & DEF_SEP, _
        "Offset", lOffset, 1
    vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments1", eCmd, ucsDcsCmdInfoTotalsByPayments, sParam, "1" & DEF_SEP, _
        "RevOffset", lRevOffset, 1
    vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPmtSum) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        lPmtIndex = pvToDevicePmtIndex(lIdx + 1)
        If lPmtIndex >= 0 Then
            vRetVal(lIdx) = Array(pvToDevicePmtName(lIdx + 1), ParseSum(At(vPmtSum, lPmtIndex + lOffset)), ParseSum(At(vRevSum, lPmtIndex + lRevOffset)))
        End If
    Next
    '--- success
    GetTotalsByPaymentTypes = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByTaxGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByTaxGroups"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vTaxRates       As Variant
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", 0, 0, sParam, "valVat", _
        "Offset", lOffset, 0
    ReDim vTaxRates(0 To MAX_TAX_GROUP - 1) As String
    For lIdx = 0 To UBound(vTaxRates)
        vTaxRates(lIdx) = pvSettings(sParam, lIdx + lOffset)
    Next
    pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups0", eCmd, ucsDcsCmdInfoTotalsByTaxGroups, sParam, "0" & DEF_SEP, _
        "Offset", lOffset, 2
    vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups2", eCmd, ucsDcsCmdInfoTotalsByTaxGroups, sParam, "2" & DEF_SEP, _
        "RevOffset", lRevOffset, 2
    vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPmtSum) <= 0 And UBound(vRevSum) <= 0 Then
        GetTotalsByTaxGroups = Array()
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = Array(ParseSum(At(vTaxRates, lIdx)), ParseSum(At(vPmtSum, lIdx + lOffset)), ParseSum(At(vRevSum, lIdx + lRevOffset)))
    Next
    '--- success
    GetTotalsByTaxGroups = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsStats() As Variant
    Const FUNC_NAME     As String = "GetTotalsStats"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lReceipts       As Long
    Dim dLastDate       As Date
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDevice", eCmd, ucsDcsCmdInfoDevice, sParam, "3" & DEF_SEP, _
        "Index", lIndex, 3
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    lReceipts = C_Lng(At(vSplit, lIndex))
    dLastDate = ParseExtendedDate(At(vSplit, lIndex + 1))
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetLocalizedCommand( _
            sFunction As String, _
            sKey As String, _
            Optional ByVal cmd As UcsDatecsCommandsEnum, _
            Optional Param As String, _
            Optional More As Variant)
    Const FUNC_NAME     As String = "SetLocalizedCommand"
    Dim sMerged         As String
    Dim lIdx            As Long
    
    On Error GoTo EH
    If m_uConfig.LocalizedCommands Is Nothing Then
        Set m_uConfig.LocalizedCommands = New Collection
    End If
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    With m_uConfig.LocalizedCommands
        If cmd <> 0 Then
            .Add cmd, sMerged
        End If
        If LenB(Param) <> 0 Then
            .Add Param, sMerged & "Param"
        End If
        If IsArray(More) Then
            For lIdx = 0 To UBound(More) Step 2
                .Add More(lIdx + 1), sMerged & More(lIdx)
            Next
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub ClearLocalizedCommand(sFunction As String, sKey As String, Optional More As Variant)
    Const FUNC_NAME     As String = "ClearLocalizedCommand"
    Dim sMerged         As String
    Dim lIdx            As Long
    
    On Error GoTo EH
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    If SearchCollection(m_uConfig.LocalizedCommands, sMerged) Then
        m_uConfig.LocalizedCommands.Remove sMerged
    End If
    If SearchCollection(m_uConfig.LocalizedCommands, sMerged & "Param") Then
        m_uConfig.LocalizedCommands.Remove sMerged & "Param"
    End If
    If IsArray(More) Then
        For lIdx = 0 To UBound(More)
            If SearchCollection(m_uConfig.LocalizedCommands, sMerged & More(lIdx)) Then
                m_uConfig.LocalizedCommands.Remove sMerged & More(lIdx)
            End If
        Next
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function GetOperatorName(OperatorCode As String) As String
    Const FUNC_NAME     As String = "GetOperatorName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "OperName"
    GetOperatorName = pvSettings(sParam, C_Lng(OperatorCode))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    If LenB(m_uConfig.DefaultPassword) <> 0 Then
        GetDefaultPassword = m_uConfig.DefaultPassword
    ElseIf IsEcr Then
        GetDefaultPassword = OperatorCode
    Else
        GetDefaultPassword = DEF_PASSWORD
    End If
End Function

Public Function GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    Const FUNC_NAME     As String = "GetDepartmentName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings1", 0, 0, sParam, "Dept_name"
    GetDepartmentName = pvSettings(sParam, C_Lng(DepartmentCode))
    pvGetCommandConfig CONF_NAME, "InfoSettings2", 0, 0, sParam, "Dept_vat"
    TaxGroup = C_Val(pvSettings(sParam, C_Lng(DepartmentCode)))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxRates() As Variant
    Const FUNC_NAME     As String = "GetTaxRates"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    Dim lOffset         As Long
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Double
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", 0, 0, sParam, "valVat", _
        "Offset", lOffset, 0
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = ParseSum(pvSettings(sParam, lIdx + lOffset))
        If vRetVal(lIdx) >= 100 Then
            vRetVal(lIdx) = -1
        End If
    Next
    '--- success
    GetTaxRates = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvTransmitCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsDatecsCommandsEnum, _
            Optional Data As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvTransmitCommand"
    Dim sSend           As String
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim sBuffer         As String
    Dim dblTimerEx      As Double
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "eCmd=" & eCmd & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    '--- prepare command
    sSend = Data
    sSend = Chr$(SYM_PREAMBLE) & pvToWireLong(SYM_BASEDATA + 10 + Len(sSend)) & Chr$(m_lSequence) & pvToWireLong(eCmd) & sSend & Chr$(SYM_POSTAMBLE)
    '--- increment sequence
    m_lSequence = IIf(m_lSequence >= LNG_MAX_SEQUENCE, DEF_SEQUENCE, m_lSequence + 1)
    '--- append checksum and terminator
    sSend = sSend & pvCheckSum(sSend) & Chr$(SYM_TERMINATOR)
    '--- send request
    dblTimerEx = TimerEx
RetrySend:
    If TimerEx > dblTimerEx + lTimeout / 1000# Then
        Error = pvConnectorError(ucsErrTimeoutWaitingForResponse)
        GoTo QH
    End If
    DebugDataDump FUNC_NAME, "[SEND ] ", sSend
    baText = ToAscii(sSend, m_uConfig.CodePage)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog FUNC_NAME, "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    sText = vbNullString
    Do
        If TimerEx > dblTimerEx + lTimeout / 1000# Then
            Error = pvConnectorError(ucsErrTimeoutWaitingForResponse)
            GoTo QH
        End If
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog FUNC_NAME, "ReadData failed, Error=" & Error
            GoTo QH
        End If
        sBuffer = FromAscii(baText, m_uConfig.CodePage)
        DebugDataDump FUNC_NAME, "[RECV ] ", sBuffer
        '--- check for NAK (&H15)
        If InStr(sBuffer, Chr$(SYM_NAK)) > 0 Then
            DebugLog FUNC_NAME, "Will retry send on NAK received, sBuffer=0x" & ToHexDump(sBuffer)
            GoTo RetrySend
        End If
        '--- remove SYN (&H16)
        sText = sText & Replace(sBuffer, Chr$(SYM_SYN), vbNullString)
        '--- restart timeout timer
        If InStr(sBuffer, Chr$(SYM_SYN)) > 0 Then
            dblTimerEx = TimerEx
        End If
    Loop While Right$(sText, 1) <> Chr$(SYM_TERMINATOR)
RetryParse:
    '--- response complete w/ TERMINATOR (&H3)
    Debug.Assert Right$(sText, 1) = Chr$(SYM_TERMINATOR)
    If Left$(sText, 1) = Chr$(SYM_PREAMBLE) Then
        lIdx = InStr(sText, Chr$(SYM_POSTAMBLE))
        If lIdx < 10 Then
            Error = Printf(pvInternal(ucsErrResponseMissingPostamble), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        lSize = pvFromWireLong(Mid$(sText, 2, 4)) - SYM_BASEDATA
        '--- check message length
        If lSize < 10 Or lSize >= lIdx Then
            Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        '--- check sum
        If pvCheckSum(Mid$(sText, 1, lSize + 1)) <> Mid$(sText, lSize + 2, 4) Then
            Error = Printf(pvInternal(ucsErrInvalidResponseChecksum), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        '--- check for status info
        lIdx = InStrRev(sText, Chr$(SYM_SEPARATOR), lIdx)
        If lIdx > 10 Then
            pvSetStatus Mid$(sText, lIdx + 1, 8)
            lSize = lIdx - 1
        Else
            pvSetStatus vbNullString '--- clear
        End If
        pvTransmitCommand = Mid$(sText, 11, lSize - 10)
        If Status(ucsDcsStbGeneralError) Then
            Error = Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
        End If
        DebugLog FUNC_NAME, "RetVal=" & pvTransmitCommand & _
            IIf(Status(ucsDcsStbGeneralError), ", ErrorText=" & Error, vbNullString) & _
            IIf(Memory(ucsDcsFmeMemoryError), ", MemoryText=" & Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory))), vbNullString)
    Else
        '--- check for GSM tax terminal interferences: "AT+CPIN?\r\n" or "AT+CIMI\r\n"
        lIdx = InStr(sText, Chr$(SYM_PREAMBLE))
        If lIdx > 1 Then
            DebugLog FUNC_NAME, "Trim bogus symbols, Prefix=0x" & ToHexDump(Left$(sText, lIdx - 1))
            sText = Mid$(sText, lIdx)
            GoTo RetryParse
        End If
        Error = Printf(pvInternal(ucsErrResponseMissingPreamble), "0x" & ToHexDump(sText))
        pvSetStatus vbNullString '--- clear
    End If
QH:
    RaiseEvent CommandComplete(eCmd, Data, pvTransmitCommand)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCheckSum(sText As String) As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    
    baText = ToAscii(sText, m_uConfig.CodePage)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    pvCheckSum = pvToWireLong(lSum)
End Function

Private Function pvToWireLong(ByVal lValue As Long) As String
    pvToWireLong = Chr$((lValue \ &H1000 Mod &H10) + &H30) & Chr$((lValue \ &H100 Mod &H10) + &H30) & Chr$((lValue \ &H10 Mod &H10) + &H30) & Chr$((lValue \ &H1 Mod &H10) + &H30)
End Function

Private Function pvFromWireLong(baValue() As Byte) As Long
    If baValue(0) >= &H30 And baValue(2) >= &H30 And baValue(4) >= &H30 And baValue(6) >= &H30 Then
        pvFromWireLong = (baValue(0) - &H30) * &H1000 + (baValue(2) - &H30) * &H100 + (baValue(4) - &H30) * &H10 + (baValue(6) - &H30)
    End If
End Function

Private Sub pvSetStatus(sStatus As String)
    Const FUNC_NAME     As String = "pvSetStatus"
    Dim baStatus()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo EH
    baStatus = ToAscii(sStatus, m_uConfig.CodePage)
    For lIdx = 0 To UBound(baStatus)
        If (baStatus(lIdx) And &H80) = 0 Then
            baStatus(lIdx) = 0
        End If
    Next
    If UBound(baStatus) >= 2 Then
        Call CopyMemory(m_eStatus, baStatus(0), 3)
    Else
        m_eStatus = 0
    End If
    If UBound(baStatus) >= 5 Then
        Call CopyMemory(m_eMemory, baStatus(4), 2)
    Else
        m_eMemory = 0
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvGetFlagsText(ByVal lFlags As Long, vTexts As Variant) As String
    Const FUNC_NAME     As String = "pvGetFlagsText"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To UBound(vTexts)
        If (lFlags And (2 ^ lIdx)) <> 0 And Len(vTexts(lIdx)) > 1 Then
            pvGetFlagsText = IIf(LenB(pvGetFlagsText) <> 0, pvGetFlagsText & ", ", vbNullString) & vTexts(lIdx)
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    Const FUNC_NAME     As String = "pvInternal"
    
    On Error GoTo EH
    pvInternal = At(m_uConfig.LocalizedText(ucsFscLciInternalErrors), lIdx)
    If LenB(pvInternal) = 0 Then
        pvInternal = At(Split(STR_INTERNAL, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvText(ByVal lIdx As UcsReceiptTextsEnum) As String
    Const FUNC_NAME     As String = "pvText"
    
    On Error GoTo EH
    pvText = At(m_uConfig.LocalizedText(ucsFscLciReceiptTexts), lIdx)
    If LenB(pvText) = 0 Then
        pvText = At(Split(STR_RECEIPT_TEXTS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvConnectorError(ByVal lIdx As UcsConnectorErrorsEnum) As String
    Const FUNC_NAME     As String = "pvConnectorError"
    
    On Error GoTo EH
    pvConnectorError = At(m_uConfig.LocalizedText(ucsFscLciConnectorErrors), lIdx)
    If LenB(pvConnectorError) = 0 Then
        pvConnectorError = At(Split(STR_CONNECTOR_ERRORS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem               As Variant
    Dim eCmd                As UcsDatecsCommandsEnum
    Dim sParam              As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "FiscalText", eCmd, ucsDcsCmdFiscalText, sParam, "%1" & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, Printf(sParam, C_Str(vElem))
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Next
    Else
        SendCommand eCmd, Printf(sParam, C_Str(vText))
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
    End If
    '--- success
    pvPrintLine = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintReceipt(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Const CONF_NAME     As String = "PrintReceipt"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lStatusClosed   As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim sFormatZero     As String
    Dim sFormatDate     As String
    Dim lAmountIndex    As Long
    Dim lPaidIndex      As Long
    Dim sDiscType       As String
    Dim sCodeType       As String
    Dim sCodeData       As String
    Dim dLastDate       As Date
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            If .InitDisablePrinting Then
                pvGetCommandConfig CONF_NAME, "InitDisablePrinting", eCmd, ucsDcsCmdInitConnectionMode, sParam, "0" & DEF_SEP
                SendCommand eCmd, sParam
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
            End If
            pvGetCommandConfig CONF_NAME, "InfoOperator", 0, 0, sParam, "OperName"
            sData = pvSettings(sParam, C_Lng(.InitOperatorCode) - 1)
            If LenB(.InitOperatorName) <> 0 Then
                If .InitOperatorName <> sData Then
                    pvSettings(sParam, C_Lng(.InitOperatorCode) - 1) = Left$(.InitOperatorName, m_uConfig.MaxOperatorName)
                End If
            Else
                .InitOperatorName = RTrim$(sData)
            End If
            '--- for operator errors: if linked receipt in current Z-report -> print service deposit for non-cash payments
            '---                      else -> print service deposit for total receipt amount
            If LenB(.InitRevData(ucsRevReceiptNo)) <> 0 And .InitRevData(ucsRevType) = ucsFscRevOperatorError Then
                If .InitRevData(ucsRevFiscalMemoryNo) = GetFiscalMemoryNo() Then
                    pvGetCommandConfig CONF_NAME, "InfoDevice", eCmd, ucsDcsCmdInfoDevice, sParam, "3" & DEF_SEP, _
                        "Index", lIndex, 4
                    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
                    dLastDate = ParseExtendedDate(At(vSplit, lIndex))
                Else
                    dLastDate = .InitRevData(ucsRevReceiptDate)
                End If
                vSplit = PpdGetTotalsByPaymentTypes(uRow, lRowCount)
                dblTotal = C_Dbl(At(vSplit, 0)) - IIf(.InitRevData(ucsRevReceiptDate) > dLastDate, C_Dbl(At(vSplit, ucsFscPmtCash)), 0)
                If Abs(dblTotal) > DBL_EPSILON Then
                    PrintServiceDeposit dblTotal
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End If
            End If
            '--- setup invoice no
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    pvGetCommandConfig CONF_NAME, "InfoInvoiceRangeBeg", 0, 0, sParam, "InvoiceRangeBeg"
                    pvSettings(sParam) = .InitInvData(ucsInvDocNo)
                    pvGetCommandConfig CONF_NAME, "InfoInvoiceRangeEnd", 0, 0, sParam, "InvoiceRangeEnd"
                    pvSettings(sParam) = .InitInvData(ucsInvDocNo)
                End Select
            End If
            '--- setup fiscal receipt
            Select Case .InitReceiptType
            Case ucsFscRcpReversal, ucsFscRcpCreditNote
                '--- %1={OpCode}<SEP>%2={OpPwd}<SEP>{TillNmb}<SEP>%3={Storno}<SEP>%4={DocNum}<SEP>%5={DateTime}<SEP>%6={FMNumber}
                '---    <SEP>{Invoice}<SEP>%7={ToInvoice}<SEP>%8={Reason}<SEP>%9={NSale}<SEP>
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "1" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & _
                    IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%7" & DEF_SEP & "%8", DEF_SEP & DEF_SEP & DEF_SEP) & _
                    DEF_SEP & "%9" & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalReversal" & .InitReceiptType, eCmd, ucsDcsCmdFiscalReversal, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, .InitRevData(ucsRevType), _
                    .InitRevData(ucsRevReceiptNo), Format$(.InitRevData(ucsRevReceiptDate), sFormatDate), .InitRevData(ucsRevFiscalMemoryNo), _
                    .InitRevData(ucsRevInvoiceNo), .InitRevData(ucsRevReason), .InitUniqueSaleNo)
            Case Else
                '--- %1={OpCode}<SEP>%2={OpPwd}%3=(<SEP>{NSale})<SEP>{TillNmb}<SEP>{Invoice}<SEP>
                sParam = "%1" & DEF_SEP & "%2%3" & DEF_SEP & "1" & DEF_SEP & IIf(.InitReceiptType = ucsFscRcpInvoice, "I", vbNullString) & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsDcsCmdFiscalOpen, sParam, sParam
                sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & .InitUniqueSaleNo, vbNullString))
            End Select
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            '--- assert invoice number generated
            Select Case .InitReceiptType
            Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                uCtx.InvoiceNo = GetLastInvoiceNo()
                If LenB(.InitInvData(ucsInvDocNo)) <> 0 And LenB(uCtx.InvoiceNo) <> 0 Then
                    If Val(.InitInvData(ucsInvDocNo)) <> Val(uCtx.InvoiceNo) Then
                        pvSetLastError Printf(pvInternal(ucsErrWrongInvoiceNo), uCtx.InvoiceNo)
                        GoTo QH
                    End If
                End If
            End Select
        Case ucsRowPlu
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluPrice * .PluQuantity, DEF_PRICE_SCALE)
            If .DiscType = ucsFscDscPlu Then
                dblDisc = Round(dblTotal * .DiscValue / 100#, DEF_PRICE_SCALE)
            ElseIf .DiscType = ucsFscDscPluAbs Then
                dblDisc = Round(.DiscValue, DEF_PRICE_SCALE)
            Else
                dblDisc = 0
            End If
            uCtx.GrpTotal(.PluTaxGroup) = Round(uCtx.GrpTotal(.PluTaxGroup) + dblTotal + dblDisc, DEF_PRICE_SCALE)
            vSplit = WrapText(.PluItemName, Clamp(pvCommentChars, , pvItemChars))
            For lIdx = 0 To UBound(vSplit) - 1
                If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                    GoTo QH
                End If
            Next
            sData = At(vSplit, UBound(vSplit))
            sDiscType = vbNullString
            Select Case .DiscType
            Case ucsFscDscPlu
                If Abs(.PluPrice) > DBL_EPSILON Then
                    sDiscType = Switch(.DiscValue > DBL_EPSILON, "1", .DiscValue < -DBL_EPSILON, "2", True, vbNullString)
                End If
            Case ucsFscDscPluAbs
                sDiscType = Switch(.DiscValue > DBL_EPSILON, "3", .DiscValue < -DBL_EPSILON, "4", True, vbNullString)
            End Select
            '--- %1={PluName}<SEP>%2={TaxCd}<SEP>%3={Price}<SEP>%4={Quantity}<SEP>%5={DiscountType}<SEP>%6={DiscountValue}<SEP>%7={Department}<SEP>%8=[ {Unit}<SEP> ]
            sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & DEF_SEP & "%7" & DEF_SEP & "%8" & DEF_SEP
            pvGetCommandConfig CONF_NAME, "FiscalSell" & .PrintRowType, eCmd, ucsDcsCmdFiscalSell, sParam, sParam, _
                "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                "FormatQty", sFormatQty, FORMAT_BASE_3, _
                "FormatDisc", sFormatDisc, FORMAT_BASE_2
            sParam = Printf(sParam, sData, .PluTaxGroup, SafeFormat(.PluPrice, sFormatPrice), _
                IIf(Abs(.PluQuantity - 1) > DBL_EPSILON, SafeFormat(.PluQuantity, sFormatQty), vbNullString), _
                sDiscType, _
                IIf(LenB(sDiscType) <> 0, SafeFormat(Abs(.DiscValue), sFormatDisc), vbNullString), _
                .PluDepartmentNo, _
                Zn(.PluUnitOfMeasure, " "))
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRcpNonfiscal
                Select Case .DiscType
                Case ucsFscDscSubtotal
                    pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal" & .DiscType, 0, 0, vbNullString, vbNullString, _
                        "FormatDisc", sFormatDisc, FORMAT_BASE_2, _
                        "FormatAmount", sFormatAmount, FORMAT_BASE_2
                    If Not pvPrintLine(AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvCommentChars)) Then
                        GoTo QH
                    End If
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                            If Not pvPrintLine(AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & SafeFormat(Abs(.DiscValue), sFormatDisc) & "%"), _
                                    SafeFormat(dblTotal, sFormatAmount), pvCommentChars)) Then
                                GoTo QH
                            End If
                        End If
                    Next
                Case Else
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedDiscType), .DiscType)
                    GoTo QH
                End Select
            Case Else
                sDiscType = vbNullString
                Select Case .DiscType
                Case ucsFscDscSubtotal
                    sDiscType = Switch(.DiscValue > DBL_EPSILON, "1", .DiscValue < -DBL_EPSILON, "2", True, vbNullString)
                Case ucsFscDscSubtotalAbs
                    sDiscType = Switch(.DiscValue > DBL_EPSILON, "3", .DiscValue < -DBL_EPSILON, "4", True, vbNullString)
                Case Else
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedDiscType), .DiscType)
                    GoTo QH
                End Select
                '--- %1={Print}<SEP>{Display}<SEP>%2={DiscountType}<SEP>%3={DiscountValue}<SEP>
                sParam = "%1" & DEF_SEP & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsDcsCmdFiscalSubtotal, sParam, sParam, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
                sParam = Printf(sParam, IIf(LenB(sDiscType) = 0, "1", vbNullString), _
                    sDiscType, _
                    IIf(LenB(sDiscType) <> 0, SafeFormat(.DiscValue, sFormatDisc), vbNullString))
                SendCommand eCmd, sParam
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
            End Select
            Select Case .DiscType
            Case ucsFscDscSubtotal
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, DEF_PRICE_SCALE)
                    End If
                Next
            Case ucsFscDscSubtotalAbs
                '--- ToDo: fix for multiple tax groups
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + .DiscValue, DEF_PRICE_SCALE)
                        Exit For
                    End If
                Next
            End Select
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapText(.LineText, pvCommentChars)
            Else
                vSplit = Array(Left$(.LineText, pvCommentChars))
            End If
            If Not pvPrintLine(vSplit) Then
                GoTo QH
            End If
        Case ucsRowBarcode
            If .BarcodeHeight > 0 Then
                pvGetCommandConfig CONF_NAME, "InfoSettingsBarCodeHeight", 0, 0, sParam, "BarCodeHeight"
                pvSettings(sParam) = Clamp(.BarcodeHeight \ 7, 1, 10)
                pvGetCommandConfig CONF_NAME, "InfoSettingsBarcodeName", 0, 0, sParam, "BarcodeName"
                pvSettings(sParam) = 0
            End If
            sCodeData = .BarcodeText
            Select Case .BarcodeType
            Case ucsFscBrcEan8
                sCodeType = "1"
                sCodeData = Pad(sCodeData, 7, "0")
            Case ucsFscBrcEan13
                sCodeType = "2"
                sCodeData = Pad(sCodeData, 12, "0")
            Case ucsFscBrcCode128
                sCodeType = "3"
            Case ucsFscBrcQRcode
                sCodeType = "4"
            End Select
            '--- %1={Type}<SEP>%2={Data}<SEP>[ {QRcodeSize}<SEP> ]
            sParam = "%1" & DEF_SEP & "%2" & DEF_SEP
            pvGetCommandConfig CONF_NAME, "PrintBarcode", eCmd, ucsDcsCmdPrintBarcode, sParam, sParam
            sParam = Printf(sParam, sCodeType, sCodeData)
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowPayment
            pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsDcsCmdInfoTransaction, sParam, vbNullString, _
                "Index", lIndex, 1, _
                "StatusClosed", lStatusClosed, 0, _
                "AmountIndex", lAmountIndex, 4, _
                "PaidIndex", lPaidIndex, 5
            vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
            If C_Lng(At(vSplit, lIndex)) = lStatusClosed Then
                pvSetLastError pvInternal(ucsErrCannotPrint)
                GoTo QH
            End If
            If .PmtType <> 0 Then
                If Abs(.PmtAmount) < DBL_EPSILON Then
                    eCmd = ucsDcsCmdNoop
                ElseIf uRow(0).InitReceiptType = ucsFscRcpReversal Or uRow(0).InitReceiptType = ucsFscRcpCreditNote Then
                    eCmd = ucsDcsCmdNoop
                Else
                    '--- %1={PaidMode}<SEP>%2={Amount}<SEP>{Type}<SEP>
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & DEF_SEP
                    pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsDcsCmdFiscalPayment, sParam, sParam, _
                        "FormatAmount", sFormatAmount, FORMAT_BASE_2, _
                        "FormatZero", sFormatZero, "\"
                    lIndex = pvToDevicePmtIndex(.PmtType)
                    If lIndex < 0 Then
                        pvSetLastError Printf(pvInternal(ucsErrUnsupportedPaymentType), .PmtType)
                        GoTo QH
                    End If
                    sParam = Printf(sParam, lIndex, SafeFormat(.PmtAmount, IIf(Abs(.PmtAmount) < DBL_EPSILON, sFormatZero, sFormatAmount)))
                End If
            Else
                '--- check if anything left for payment
                If C_Dbl(At(vSplit, lAmountIndex)) > C_Dbl(At(vSplit, lPaidIndex)) _
                        Or LenB(At(vSplit, lAmountIndex)) = 0 _
                        Or Not uCtx.PmtPrinted Then
                    sParam = DEF_SEP & DEF_SEP & DEF_SEP
                    pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsDcsCmdFiscalPayment, sParam, sParam
                    sParam = Printf(sParam, pvToDevicePmtIndex(ucsFscPmtCash))
                    SendCommand eCmd, sParam
                    If LenB(m_uPrintData.LastError) = 0 Then
                        uCtx.PmtPrinted = True
                    End If
                End If
                Select Case .PrintRowType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    '--- %1={Seller}<SEP>%2={Receiver}<SEP>%3={Buyer}<SEP>%4={Address1}<SEP>%5={Address2}<SEP>%6={TypeTAXN}<SEP>%7={TAXN}<SEP>%8={VATN}<SEP>
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & DEF_SEP & "%7" & DEF_SEP & "%8" & DEF_SEP
                    pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsDcsCmdFiscalCgInfo, sParam, sParam
                    With uRow(0)
                        sParam = Printf(sParam, Left$(.InitOperatorName, m_uConfig.MaxCgInfo(ucsMciOperatorName)), _
                            Left$(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxCgInfo(ucsMciCgPrsReceive)), _
                            Left$(.InitInvData(ucsInvCgName), m_uConfig.MaxCgInfo(ucsMciCgName)), _
                            Left$(.InitInvData(ucsInvCgCity), m_uConfig.MaxCgInfo(ucsMciCgCity)), _
                            Left$(.InitInvData(ucsInvCgAddress), m_uConfig.MaxCgInfo(ucsMciCgAddress)), _
                            .InitInvData(ucsInvCgTaxNoType), _
                            Left$(.InitInvData(ucsInvCgTaxNo), m_uConfig.MaxCgInfo(ucsMciCgTaxNo)), _
                            Left$(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxCgInfo(ucsMciCgVatNo)))
                    End With
                    SendCommand eCmd, sParam
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End Select
                uCtx.ReceiptNo = GetLastReceiptNo()
                uCtx.ReceiptDate = GetClock()
                uCtx.ReceiptAmount = C_Dbl(At(vSplit, lAmountIndex))
                '--- reversal type: 2 - operator error, 3 - refund, 4 - tax base reduction
                If C_Lng(At(vSplit, lIndex)) >= 2 And C_Lng(At(vSplit, lIndex)) <= 4 Then
                    uCtx.ReceiptAmount = -uCtx.ReceiptAmount
                End If
                If uCtx.PmtPrinted Or SumArray(uCtx.GrpTotal) > DBL_EPSILON Then
                    pvGetCommandConfig CONF_NAME, "FiscalClose" & .PrintRowType, eCmd, ucsDcsCmdFiscalClose, sParam, vbNullString
                Else
                    pvGetCommandConfig CONF_NAME, "FiscalCancel" & .PrintRowType, eCmd, ucsIslCmdFiscalCancel, sParam, vbNullString
                End If
            End If
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            If eCmd <> ucsDcsCmdNoop Then
                uCtx.PmtPrinted = True
            End If
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
QH:
    If uRow(0).InitDisablePrinting Then
        pvGetCommandConfig CONF_NAME, "InitEnablePrinting", eCmd, ucsDcsCmdInitConnectionMode, sParam, "1" & DEF_SEP
        '--- use pvTransmitCommand to preserve LastError
        pvTransmitCommand m_oConnector, m_lCommandTimeout, eCmd, sParam
    End If
    If m_uPrintData.LastErrNo = ucsFerInvalidPassword Then
        pvSetLastError m_uPrintData.LastError & pvInternal(ucsErrInvalidOperatorPassword), ucsFerInvalidPassword
    End If
    If LenB(m_uPrintData.LastError) <> 0 And uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_uPrintData.LastError), m_uPrintData.LastErrNo
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvGetDeviceInfo( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            vResult As Variant, _
            sProtocol As String, _
            sModel As String, _
            sFirmware As String, _
            sDeviceSerialNo As String, _
            sFiscalMemoryNo As String, _
            Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lNameIndex      As Long
    Dim lSerialIndex    As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsDcsCmdInfoDiagnostics, sParam, DEF_SEP, _
        "NameIndex", lNameIndex, 1, _
        "SerialIndex", lSerialIndex, 7
    If Not IsArray(vResult) Then
        vSplit = Split(pvTransmitCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_SEP)
        If UBound(vSplit) >= lSerialIndex Then
            vResult = vSplit
        End If
    End If
    If LenB(At(vResult, lSerialIndex)) <> 0 Then
        '--- check serial
        sProtocol = STR_PROTOCOL_DATECS_X
        sModel = RTrim$(At(vResult, lNameIndex))
        sFirmware = RTrim$(At(vResult, lNameIndex + 1) & " " & At(vResult, lNameIndex + 2) & " " & At(vResult, lNameIndex + 3))
        sDeviceSerialNo = RTrim$(At(vResult, lSerialIndex))
        sFiscalMemoryNo = RTrim$(At(vResult, lSerialIndex + 1))
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvStripBold(sText As String) As String
    Const FUNC_NAME     As String = "pvStripBold"
    
    On Error GoTo EH
    pvStripBold = preg_replace(GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Pattern", "/\^/i"), _
        sText, GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Replace", vbNullString))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Const FUNC_NAME     As String = "pvToDevicePmtIndex"
    
    On Error GoTo EH
    pvToDevicePmtIndex = -1         ' unsupported
    Select Case PmtType
    Case ucsFscPmtCash
        pvToDevicePmtIndex = 0      ' 0=SCash (б апни)
    Case ucsFscPmtCheque
        pvToDevicePmtIndex = 3      ' 1=SChecks (вей)
    Case ucsFscPmtCoupon
        pvToDevicePmtIndex = 4      ' 2=ST (йсонм)
    Case ucsFscPmtVoucher
        pvToDevicePmtIndex = 5      ' 3=SOT (бюсвеп)
    Case ucsFscPmtCard
        pvToDevicePmtIndex = 2      ' 7=SCards (деа.йюпрю)
    Case ucsFscPmtBank
'        pvToDevicePmtIndex = 1      ' 7=SCards (йпедхр)
        pvToDevicePmtIndex = -1     ' unsupported
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToDevicePmtName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "pvToDevicePmtName"
    Const CONF_NAME     As String = "GetPaymentName"
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    lIndex = pvToDevicePmtIndex(PmtType)
    pvGetCommandConfig CONF_NAME, "InfoSettings" & PmtType, 0, 0, sParam, "PayName", _
        "PmtType", lIndex, lIndex
    If lIndex >= 0 Then
        pvToDevicePmtName = RTrim$(pvStripBold(pvSettings(sParam, lIndex)))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvGetCommandConfig( _
            sFunction As String, _
            sKey As String, _
            eCmd As UcsDatecsCommandsEnum, _
            ByVal eCmdDefault As UcsDatecsCommandsEnum, _
            sParam As String, _
            sParamDefault As String, _
            ParamArray A() As Variant)
    Const FUNC_NAME     As String = "pvGetCommandConfig"
    Dim lIdx            As Long
    
    On Error GoTo EH
    eCmd = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey, eCmdDefault)
    sParam = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & "Param", sParamDefault)
    For lIdx = 0 To UBound(A) Step 3
        A(lIdx + 1) = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & A(lIdx + 0), A(lIdx + 2))
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvSetLastError(sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1)
    Const FUNC_NAME     As String = "pvSetLastError"
    
    On Error GoTo EH
    If ErrNum < 0 Then
        m_uPrintData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_uPrintData.LastErrNo = ErrNum
    End If
    m_uPrintData.LastError = sError
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvToErrorCode(sCode As String) As String
    Const FUNC_NAME     As String = "pvToErrorCode"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = Split(STR_ERROR_CODES, "|")
    For lIdx = 0 To UBound(vSplit) - 1 Step 2
        If sCode = At(vSplit, lIdx) Then
            pvToErrorCode = At(vSplit, lIdx + 1)
            Exit Function
        End If
    Next
    pvToErrorCode = "Error " & sCode
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    m_lCommandTimeout = DEF_TIMEOUT
    m_lSequence = DEF_SEQUENCE
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS_ONLY
    LocalizedText(ucsFscLciMemoryTexts) = STR_MEMORY
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_DATECS_X, "Commands")
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddBarcode(ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, Text As String, Optional ByVal Height As Long) As Boolean
    IDeviceProtocol_AddBarcode = PpdAddBarcode(m_uPrintData, BarcodeType, Text, Height)
End Function

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = PpdAddDiscount(m_uPrintData, DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = PpdAddLine(m_uPrintData, Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal PmtType As UcsFiscalPaymentTypeEnum, Optional PmtName As String, Optional ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = PpdAddPayment(m_uPrintData, PmtType, PmtName, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, Optional ByVal Quantity As Double = 1, Optional ByVal TaxGroup As Long = 2, Optional UnitOfMeasure As String, Optional ByVal DepartmentNo As Long) As Boolean
    IDeviceProtocol_AddPLU = PpdAddPLU(m_uPrintData, Name, Price, Quantity, TaxGroup, UnitOfMeasure, DepartmentNo)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double) As Variant
    IDeviceProtocol_PrintServiceDeposit = PrintServiceDeposit(Amount, ReceiptNo, ReceiptDate, ReceiptAmount)
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_PrintDuplicate(Optional DuplicateNo As String, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDuplicate = PrintDuplicate(DuplicateNo, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = Disconnect()
End Function

Private Function IDeviceProtocol_PrintReceipt(Optional ResumeToken As String, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double, Optional InvoiceNo As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        IDeviceProtocol_PrintReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
        With m_uPrintData.ExecCtx
            ReceiptNo = .ReceiptNo
            ReceiptDate = .ReceiptDate
            ReceiptAmount = .ReceiptAmount
            InvoiceNo = .InvoiceNo
        End With
    End If
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = pvRowChars
End Function

Private Function IDeviceProtocol_GetCommentTextMaxLength() As Long
    IDeviceProtocol_GetCommentTextMaxLength = pvCommentChars
End Function

Private Function IDeviceProtocol_GetItemNameMaxLength() As Long
    IDeviceProtocol_GetItemNameMaxLength = pvItemChars
End Function

Private Function IDeviceProtocol_GetClock() As Date
    IDeviceProtocol_GetClock = GetClock()
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    IDeviceProtocol_GetOperatorName = GetOperatorName(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    IDeviceProtocol_GetDefaultPassword = GetDefaultPassword(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    IDeviceProtocol_GetDeviceFirmware = GetDeviceFirmware()
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(StatusText)
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String
    IDeviceProtocol_GetFiscalMemoryNo = GetFiscalMemoryNo()
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = GetFooterText(Index)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = GetHeaderText(Index)
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = LastErrNo
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String
    IDeviceProtocol_GetLastReceiptNo = GetLastReceiptNo()
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String
    IDeviceProtocol_GetLastInvoiceNo = GetLastInvoiceNo()
End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    IDeviceProtocol_GetLastQRCodeInfo = GetLastQRCodeInfo(LastReceiptDate)
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeToken() As String
    IDeviceProtocol_GetResumeToken = PpdGetResumeToken(m_uPrintData)
End Function

Private Function IDeviceProtocol_GetDeviceSerialNo() As String
    IDeviceProtocol_GetDeviceSerialNo = GetDeviceSerialNo()
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String
    IDeviceProtocol_GetTaxCaption = GetTaxCaption()
End Function

Private Function IDeviceProtocol_GetTaxNo() As String
    IDeviceProtocol_GetTaxNo = GetTaxNo()
End Function

Private Function IDeviceProtocol_GetTotalsByPaymentTypes() As Variant
    IDeviceProtocol_GetTotalsByPaymentTypes = GetTotalsByPaymentTypes()
End Function

Private Function IDeviceProtocol_GetTotalsByTaxGroups() As Variant
    IDeviceProtocol_GetTotalsByTaxGroups = GetTotalsByTaxGroups()
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    IDeviceProtocol_GetTotalsStats = GetTotalsStats()
End Function

Private Function IDeviceProtocol_Init(DeviceString As String) As Boolean
    IDeviceProtocol_Init = Init(DeviceString)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = OpenDrawer()
End Function

Private Function IDeviceProtocol_PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintMonthlyReport = PrintMonthlyReport(ReportType, StartDate, EndDate, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_PrintDailyXReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyXReport = PrintDailyXReport(ReportType, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_PrintDailyZReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyZReport = PrintDailyZReport(ReportType, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = SetClock(NewDate)
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    Dim sOperCode       As String
    Dim sOperPass       As String
    
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.CommentChars = pvCommentChars
    m_uPrintData.Config.ItemChars = pvItemChars
    sOperCode = Zn(OperatorCode, DEF_OPERNO)
    sOperPass = Zn(OperatorPassword, GetDefaultPassword(sOperCode))
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, sOperCode, OperatorName, sOperPass, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    IDeviceProtocol_GetDepartmentName = GetDepartmentName(DepartmentCode, TaxGroup)
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant
    IDeviceProtocol_GetTaxRates = GetTaxRates()
End Function
