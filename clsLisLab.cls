VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FS As Form
Private TxDia As VB.TextBox
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow

Private Const PasVe = "_nDAFkEch8a1xlNuo*"
Private Const PasOw = "*mSAFjEvh2a5xlMud!"

'-------------------- LL22 --------------------------

Private Const LL_LINK_HPOS_NONE = 0                                 ' 0x0000
Private Const LL_LINK_HPOS_START = 1                                ' 0x0001
Private Const LL_LINK_HPOS_END = 2                                  ' 0x0002
Private Const LL_LINK_HPOS_ABS = 3                                  ' 0x0003
Private Const LL_LINK_HPOS_MASK = 7                                 ' 0x0007
Private Const LL_LINK_VPOS_NONE = 0                                 ' 0x0000
Private Const LL_LINK_VPOS_START = 16                               ' 0x0010
Private Const LL_LINK_VPOS_END = 32                                 ' 0x0020
Private Const LL_LINK_VPOS_ABS = 48                                 ' 0x0030
Private Const LL_LINK_VPOS_MASK = 112                               ' 0x0070
Private Const LL_LINK_HSIZE_NONE = 0                                ' 0x0000
Private Const LL_LINK_HSIZE_PROP = 256                              ' 0x0100
Private Const LL_LINK_HSIZE_INV = 512                               ' 0x0200
Private Const LL_LINK_HSIZE_MASK = 768                              ' 0x0300
Private Const LL_LINK_VSIZE_NONE = 0                                ' 0x0000
Private Const LL_LINK_VSIZE_PROP = 4096                             ' 0x1000
Private Const LL_LINK_VSIZE_INV = 8192                              ' 0x2000
Private Const LL_LINK_VSIZE_MASK = 12288                            ' 0x3000
Private Const LL_LINK_SIZEPOS_MASK = 13175                          ' 0x3377
Private Const LL_LINK_SIZEOFPARENT = 16384                          ' 0x4000
Private Const LL_DESIGNERPRINTCALLBACK_PREVIEW_START = 1            ' 1
Private Const LL_DESIGNERPRINTCALLBACK_PREVIEW_ABORT = 2            ' 2
Private Const LL_DESIGNERPRINTCALLBACK_PREVIEW_FINALIZE = 3         ' 3
Private Const LL_DESIGNERPRINTCALLBACK_PREVIEW_QUEST_JOBSTATE = 4   ' 4
Private Const LL_DESIGNERPRINTCALLBACK_EXPORT_START = 5             ' 5
Private Const LL_DESIGNERPRINTCALLBACK_EXPORT_ABORT = 6             ' 6
Private Const LL_DESIGNERPRINTCALLBACK_EXPORT_FINALIZE = 7          ' 7
Private Const LL_DESIGNERPRINTCALLBACK_EXPORT_QUEST_JOBSTATE = 8    ' 8
Private Const LL_DESIGNERPRINTTHREAD_STATE_STOPPED = 0              ' 0
Private Const LL_DESIGNERPRINTTHREAD_STATE_SUSPENDED = 1            ' 1
Private Const LL_DESIGNERPRINTTHREAD_STATE_RUNNING = 2              ' 2
Private Const LL_INVOKEACTION_FLAG_SAVE_NO_PROJECTDESCRIPTIONCHECK = 1          ' 0x0001
Private Const LL_DRILLDOWN_START = 1                                ' 1
Private Const LL_DRILLDOWN_FINALIZE = 2                             ' 2
Private Const LL_PRINTJOB_FINALIZE = 3                              ' 3
Private Const LL_JOBOPENFLAG_NOLLXPRELOAD = 4096                    ' 0x00001000
Private Const LL_JOBOPENFLAG_ONLYEXACTLANGUAGE = 8192               ' 0x00002000
Private Const LL_DEBUG_CMBTLL = 1                                   ' 0x0001
Private Const LL_DEBUG_CMBTDWG = 2                                  ' 0x0002
Private Const LL_DEBUG_CMBTLL_NOCALLBACKS = 4                       ' 0x0004
Private Const LL_DEBUG_CMBTLL_NOSTORAGE = 8                         ' 0x0008
Private Const LL_DEBUG_CMBTLL_NOWAITDLG = 16                        ' 0x0010
Private Const LL_DEBUG_CMBTLL_NOSYSINFO = 32                        ' 0x0020
Private Const LL_DEBUG_CMBTLL_LOGTOFILE = 64                        ' 0x0040
Private Const LL_DEBUG_CMBTLS = 128                                 ' 0x0080
Private Const LL_DEBUG_PRNINFO = 256                                ' 0x0100
Private Const LL_DEBUG_CMBTLL_OBJECTSTATES = 1024                   ' 0x0400
Private Const LL_DEBUG_NOPRIVACYDATA = 2048                         ' 0x0800
Private Const LL_DEBUG_FORCE_SYSINFO = 4096                         ' 0x1000
Private Const LL_DEBUG_EVAL2HOSTEXPRESSION = 8192                   ' 0x2000
Private Const LL_DEBUG_EXTENDED_DEBWINFLAGS = -2147483648#           ' 0x80000000
Private Const LL_VERSION_MAJOR = 1                                  ' 1
Private Const LL_VERSION_MINOR = 2                                  ' 2
Private Const LL_CMND_DRAW_USEROBJ = 0                              ' 0
Private Const LL_CMND_EDIT_USEROBJ = 1                              ' 1
Private Const LL_CMND_GETSIZE_USEROBJ = 2                           ' 2
Private Const LL_CMND_GETSIZE_USEROBJ_SCM = 3                       ' 3
Private Const LL_CMND_GETSIZE_USEROBJ_PIXEL = 4                     ' 4
Private Const LL_CMND_TABLELINE = 10                                ' 10
Private Const LL_TABLE_LINE_HEADER = 0                              ' 0
Private Const LL_TABLE_LINE_BODY = 1                                ' 1
Private Const LL_TABLE_LINE_FOOTER = 2                              ' 2
Private Const LL_TABLE_LINE_FILL = 3                                ' 3
Private Const LL_TABLE_LINE_GROUP = 4                               ' 4
Private Const LL_TABLE_LINE_GROUPFOOTER = 5                         ' 5
Private Const LL_CMND_TABLEFIELD = 11                               ' 11
Private Const LL_TABLE_FIELD_HEADER = 0                             ' 0
Private Const LL_TABLE_FIELD_BODY = 1                               ' 1
Private Const LL_TABLE_FIELD_FOOTER = 2                             ' 2
Private Const LL_TABLE_FIELD_FILL = 3                               ' 3
Private Const LL_TABLE_FIELD_GROUP = 4                              ' 4
Private Const LL_TABLE_FIELD_GROUPFOOTER = 5                        ' 5
Private Const LL_CMND_EVALUATE = 12                                 ' 12
Private Const LL_CMND_OBJECT = 20                                   ' 20
Private Const LL_CMND_PAGE = 21                                     ' 21
Private Const LL_CMND_PROJECT = 22                                  ' 22
Private Const LL_CMND_DRAW_GROUP_BEGIN = 23                         ' 23
Private Const LL_CMND_DRAW_GROUP_END = 24                           ' 24
Private Const LL_CMND_DRAW_GROUPLINE = 25                           ' 25
Private Const LL_RSP_GROUP_IMT = 0                                  ' 0
Private Const LL_RSP_GROUP_NEXTPAGE = 1                             ' 1
Private Const LL_RSP_GROUP_OK = 2                                   ' 2
Private Const LL_RSP_GROUP_DRAWFOOTER = 3                           ' 3
Private Const LL_CMND_HELP = 30                                     ' 30
Private Const LL_CMND_ENABLEMENU = 31                               ' 31
Private Const LL_CMND_MODIFYMENU = 32                               ' 32
Private Const LL_CMND_SELECTMENU = 33                               ' 33
Private Const LL_CMND_GETVIEWERBUTTONSTATE = 34                     ' 34
Private Const LL_CMND_VARHELPTEXT = 35                              ' 35
Private Const LL_INFO_METER = 37                                    ' 37
Private Const LL_METERJOB_LOAD = 1                                  ' 1
Private Const LL_METERJOB_SAVE = 2                                  ' 2
Private Const LL_METERJOB_CONSISTENCYCHECK = 3                      ' 3
Private Const LL_METERJOB_PASS2 = 4                                 ' 4
Private Const LL_METERJOB_PASS3 = 5                                 ' 5
Private Const LL_NTFY_FAILSFILTER = 1000                            ' 1000
Private Const LL_NTFY_VIEWERBTNCLICKED = 38                         ' 38
Private Const LL_CMND_DLGEXPR_VARBTN = 39                           ' 39
Private Const LL_CMND_HOSTPRINTER = 40                              ' 40
Private Const LL_PRN_CREATE_DC = 1                                  ' 1
Private Const LL_PRN_DELETE_DC = 2                                  ' 2
Private Const LL_PRN_SET_ORIENTATION = 3                            ' 3
Private Const LL_PRN_GET_ORIENTATION = 4                            ' 4
Private Const LL_PRN_EDIT = 5                                       ' 5
Private Const LL_PRN_GET_DEVICENAME = 6                             ' 6
Private Const LL_PRN_GET_DRIVERNAME = 7                             ' 7
Private Const LL_PRN_GET_PORTNAME = 8                               ' 8
Private Const LL_PRN_RESET_DC = 9                                   ' 9
Private Const LL_PRN_COMPARE_PRINTER = 10                           ' 10
Private Const LL_PRN_GET_PHYSPAGE = 11                              ' 11
Private Const LL_PRN_SET_PHYSPAGE = 12                              ' 12
Private Const LL_PRN_GET_PAPERFORMAT = 13                           ' 13
Private Const LL_PRN_SET_PAPERFORMAT = 14                           ' 14
Private Const LL_OEM_TOOLBAR_START = 41                             ' 41
Private Const LL_OEM_TOOLBAR_END = 50                               ' 50
Private Const LL_NTFY_EXPRERROR = 51                                ' 51
Private Const LL_CMND_CHANGE_DCPROPERTIES_CREATE = 52               ' 52
Private Const LL_CMND_CHANGE_DCPROPERTIES_DOC = 53                  ' 53
Private Const LL_CMND_CHANGE_DCPROPERTIES_PAGE = 54                 ' 54
Private Const LL_CMND_CHANGE_DCPROPERTIES_PREPAGE = 56              ' 56
Private Const LL_CMND_MODIFY_METAFILE = 57                          ' 57
Private Const LL_INFO_PRINTJOBSUPERVISION = 58                      ' 58
Private Const LL_CMND_DELAYEDVALUE = 59                             ' 59
Private Const LL_CMND_SUPPLY_USERDATA = 60                          ' 60
Private Const LL_CMND_SAVEFILENAME = 61                             ' 61
Private Const LL_QUERY_IS_VARIABLE_OR_FIELD = 62                    ' 62
Private Const LL_NTFY_PROJECTLOADED = 63                            ' 63
Private Const LL_QUERY_DESIGNERACTIONSTATE = 64                     ' 64
Private Const LL_NTFY_DESIGNERREADY = 65                            ' 65
Private Const LL_NTFY_DESIGNERPRINTJOB = 66                         ' 66
Private Const LL_NTFY_VIEWERDRILLDOWN = 67                          ' 67
Private Const LL_NTFY_QUEST_DRILLDOWNDENIED = 68                    ' 68
Private Const LL_QUERY_DRILLDOWN_ADDITIONAL_BASETABLES_FOR_VARIABLES = 69         ' 69
Private Const LL_QUERY_DRILLDOWN_ADDITIONAL_TABLES = 70             ' 70
Private Const LL_NTFY_DRILLDOWN_DESIGNERACTION = 71                 ' 71
Private Const LL_NTFY_INPLACEDESIGNER_START = 72                    ' 72
Private Const LL_NTFY_INPLACEDESIGNER_END = 73                      ' 73
Private Const LL_QUERY_OWN_MENU = 74                                ' 74
Private Const LL_CMND_UPDATE_MENU = 75                              ' 75
Private Const LL_NTFY_FRAMEHANDLE = 76                              ' 76
Private Const LL_QUERY_DEFAULTPROJECTSTREAM = 77                    ' 77
Private Const LL_NTFY_QUEST_RP_REALDATAJOBDENIED = 78               ' 78
Private Const LL_NTFY_QUEST_EXPANDABLEREGIONSJOBDENIED = 79         ' 79
Private Const LL_NTFY_QUEST_INTERACTIVESORTINGJOBDENIED = 80         ' 80
Private Const LL_QUERY_EXPR2HOSTEXPRESSION = 81                     ' 81
Private Const LL_NTFY_REPORTPARAMETERS_COLLECTED = 82               ' 82
Private Const LL_NTFY_EXPORTERPAGEFINISHED = 83                     ' 83
Private Const LL_NTFY_HYPERLINK = 84                                ' 84
Private Const LL_NTFY_SAVEREPORTSTATEITEM = 85                      ' 85
Private Const LL_NTFY_RESTOREREPORTSTATEITEM = 86                   ' 86
Private Const LL_NTFY_PROGRESS = 87                                 ' 87
Private Const LL_NTFY_TRIGGEREDJOBINUITHREAD = 88                   ' 88
Private Const LL_NTFY_PLEASE_TRANSLATE = 89                         ' 89
Private Const LL_NTFY_PREVIEW_PRINT_START = 99                      ' 99
Private Const LL_NTFY_PREVIEW_PRINT_PAGE = 100                      ' 100
Private Const LL_NTFY_PREVIEW_PRINT_END = 101                       ' 101
Private Const LL_NTFY_EMF_PAGE = 102                                ' 102
Private Const LL_QUERY_FILENAME_FOR_EXPORTJOB = 103                 ' 103
Private Const LL_QUERY_OBJECT_NOT_SUPPORTED = 104                   ' 104
Private Const LL_QUERY_REPLACE_FILESYSTEMITEM = 105                 ' 105
Private Const LL_QUERY_HOSTIMPORT = 106                             ' 106
Private Const OBJECT_LABEL = 1                                      ' 1
Private Const OBJECT_LIST = 2                                       ' 2
Private Const OBJECT_CARD = 3                                       ' 3
Private Const LL_PROJECT_LABEL = 1                                  ' 1
Private Const LL_PROJECT_LIST = 2                                   ' 2
Private Const LL_PROJECT_CARD = 3                                   ' 3
Private Const LL_PROJECT_TOC = 4                                    ' 4
Private Const LL_PROJECT_IDX = 5                                    ' 5
Private Const LL_PROJECT_GTC = 6                                    ' 6
Private Const LL_PROJECT_LAST = 6                                   ' 6
Private Const LL_PROJECT_MASK = 15                                  ' 0x000f
Private Const LL_OBJ_MARKER = 0                                     ' 0
Private Const LL_OBJ_TEXT = 1                                       ' 1
Private Const LL_OBJ_RECT = 2                                       ' 2
Private Const LL_OBJ_LINE = 3                                       ' 3
Private Const LL_OBJ_BARCODE = 4                                    ' 4
Private Const LL_OBJ_DRAWING = 5                                    ' 5
Private Const LL_OBJ_TABLE = 6                                      ' 6
Private Const LL_OBJ_TEMPLATE = 7                                   ' 7
Private Const LL_OBJ_ELLIPSE = 8                                    ' 8
Private Const LL_OBJ_GROUP = 9                                      ' 9
Private Const LL_OBJ_RTF = 10                                       ' 10
Private Const LL_OBJ_LLX = 11                                       ' 11
Private Const LL_OBJ_INPUT = 12                                     ' 12
Private Const LL_OBJ_LAST = 12                                      ' 12
Private Const LL_OBJ_REPORTCONTAINER = 253                          ' 253
Private Const LL_OBJ_PROJECT = 254                                  ' 254
Private Const LL_OBJ_PAGE = 255                                     ' 255
Private Const LL_DELAYEDVALUE = -2147483648#                         ' 0x80000000
Private Const LL_TYPEMASK = 2147418112                              ' 0x7fff0000
Private Const LL_TABLE_FIELDTYPEMASK = 63488                        ' 0x0000f800
Private Const LL_SUBTYPEMASK = 1023                                 ' 0x000003ff
Private Const LL_TYPEFLAGS = -2147420160                             ' 0x8000f800
Private Const LL_HANDLESTREAM_TYPE_MASK = 1024                      ' 0x00000400
Private Const LL_HANDLESTREAM_BLOB = 0                              ' 0x00000000
Private Const LL_HANDLESTREAM_UTF8 = 0                              ' 0x00000000
Private Const LL_HANDLESTREAM_UTF16 = 1024                          ' 0x00000400
Private Const LL_TABLE_FOOTERFIELD = 32768                          ' 0x00008000
Private Const LL_TABLE_GROUPFIELD = 16384                           ' 0x00004000
Private Const LL_TABLE_HEADERFIELD = 8192                           ' 0x00002000
Private Const LL_TABLE_BODYFIELD = 4096                             ' 0x00001000
Private Const LL_TABLE_GROUPFOOTERFIELD = 2048                      ' 0x00000800
Private Const LL_BARCODE = 1073741824                               ' 0x40000000
Private Const LL_BARCODE_METHODMASK = 255                           ' 0x000000ff
Private Const LL_BARCODE_WITHTEXT = 256                             ' 0x00000100
Private Const LL_BARCODE_WITHOUTTEXT = 512                          ' 0x00000200
Private Const LL_BARCODE_TEXTDONTCARE = 0                           ' 0x00000000
Private Const LL_BARCODE_EAN13 = 1073741824                         ' 0x40000000
Private Const LL_BARCODE_EAN8 = 1073741825                          ' 0x40000001
Private Const LL_BARCODE_GTIN13 = 1073741824                        ' 0x40000000
Private Const LL_BARCODE_GTIN8 = 1073741825                         ' 0x40000001
Private Const LL_BARCODE_UPCA = 1073741826                          ' 0x40000002
Private Const LL_BARCODE_UPCE = 1073741827                          ' 0x40000003
Private Const LL_BARCODE_3OF9 = 1073741828                          ' 0x40000004
Private Const LL_BARCODE_25INDUSTRIAL = 1073741829                  ' 0x40000005
Private Const LL_BARCODE_25INTERLEAVED = 1073741830                 ' 0x40000006
Private Const LL_BARCODE_25DATALOGIC = 1073741831                   ' 0x40000007
Private Const LL_BARCODE_25MATRIX = 1073741832                      ' 0x40000008
Private Const LL_BARCODE_POSTNET = 1073741833                       ' 0x40000009
Private Const LL_BARCODE_FIM = 1073741834                           ' 0x4000000A
Private Const LL_BARCODE_CODABAR = 1073741835                       ' 0x4000000B
Private Const LL_BARCODE_EAN128 = 1073741836                        ' 0x4000000C
Private Const LL_BARCODE_GS1_128 = 1073741836                       ' 0x4000000C
Private Const LL_BARCODE_CODE128 = 1073741837                       ' 0x4000000D
Private Const LL_BARCODE_DP_LEITCODE = 1073741838                   ' 0x4000000E
Private Const LL_BARCODE_DP_IDENTCODE = 1073741839                  ' 0x4000000F
Private Const LL_BARCODE_GERMAN_PARCEL = 1073741840                 ' 0x40000010
Private Const LL_BARCODE_CODE93 = 1073741841                        ' 0x40000011
Private Const LL_BARCODE_MSI = 1073741842                           ' 0x40000012
Private Const LL_BARCODE_CODE11 = 1073741843                        ' 0x40000013
Private Const LL_BARCODE_MSI_10_CD = 1073741844                     ' 0x40000014
Private Const LL_BARCODE_MSI_10_10 = 1073741845                     ' 0x40000015
Private Const LL_BARCODE_MSI_11_10 = 1073741846                     ' 0x40000016
Private Const LL_BARCODE_MSI_PLAIN = 1073741847                     ' 0x40000017
Private Const LL_BARCODE_EAN14 = 1073741848                         ' 0x40000018
Private Const LL_BARCODE_GTIN14 = 1073741848                        ' 0x40000018
Private Const LL_BARCODE_UCC14 = 1073741849                         ' 0x40000019
Private Const LL_BARCODE_CODE39 = 1073741850                        ' 0x4000001A
Private Const LL_BARCODE_CODE39_CRC43 = 1073741851                  ' 0x4000001B
Private Const LL_BARCODE_PZN = 1073741852                           ' 0x4000001C
Private Const LL_BARCODE_CODE39_EXT = 1073741853                    ' 0x4000001D
Private Const LL_BARCODE_JAPANESE_POSTAL = 1073741854               ' 0x4000001E
Private Const LL_BARCODE_RM4SCC = 1073741855                        ' 0x4000001F
Private Const LL_BARCODE_RM4SCC_CRC = 1073741856                    ' 0x40000020
Private Const LL_BARCODE_SSCC = 1073741857                          ' 0x40000021
Private Const LL_BARCODE_ISBN = 1073741858                          ' 0x40000022
Private Const LL_BARCODE_GS1 = 1073741859                           ' 0x40000023
Private Const LL_BARCODE_GS1_TRUNCATED = 1073741860                 ' 0x40000024
Private Const LL_BARCODE_GS1_STACKED = 1073741861                   ' 0x40000025
Private Const LL_BARCODE_GS1_STACKED_OMNI = 1073741862              ' 0x40000026
Private Const LL_BARCODE_GS1_LIMITED = 1073741863                   ' 0x40000027
Private Const LL_BARCODE_GS1_EXPANDED = 1073741864                  ' 0x40000028
Private Const LL_BARCODE_INTELLIGENT_MAIL = 1073741865              ' 0x40000029
Private Const LL_BARCODE_PZN8 = 1073741866                          ' 0x4000002A
Private Const LL_BARCODE_CODE128_FULL = 1073741867                  ' 0x4000002B
Private Const LL_BARCODE_EAN128_FULL = 1073741868                   ' 0x4000002C
Private Const LL_BARCODE_CODABLOCK_F = 1073741869                   ' 0x4000002D
Private Const LL_BARCODE_LLXSTART = 1073741888                      ' 0x40000040
Private Const LL_BARCODE_PDF417 = 1073741888                        ' 0x40000040
Private Const LL_BARCODE_MAXICODE = 1073741889                      ' 0x40000041
Private Const LL_BARCODE_MAXICODE_UPS = 1073741890                  ' 0x40000042
Private Const LL_BARCODE_DATAMATRIX = 1073741892                    ' 0x40000044
Private Const LL_BARCODE_AZTEC = 1073741893                         ' 0x40000045
Private Const LL_BARCODE_QRCODE = 1073741894                        ' 0x40000046
Private Const LL_BARCODE_DATAMATRIX_PREMIUMADRESS = 1073741895      ' 0x40000047
Private Const LL_BARCODE_MICROPDF417 = 1073741896                   ' 0x40000048
Private Const LL_DRAWING = 536870912                                ' 0x20000000
Private Const LL_DRAWING_METHODMASK = 255                           ' 0x000000ff
Private Const LL_DRAWING_HMETA = 536870913                          ' 0x20000001
Private Const LL_DRAWING_USEROBJ = 536870914                        ' 0x20000002
Private Const LL_DRAWING_USEROBJ_DLG = 536870915                    ' 0x20000003
Private Const LL_DRAWING_HBITMAP = 536870916                        ' 0x20000004
Private Const LL_DRAWING_HICON = 536870917                          ' 0x20000005
Private Const LL_DRAWING_HEMETA = 536870918                         ' 0x20000006
Private Const LL_DRAWING_HDIB = 536870919                           ' 0x20000007
Private Const LL_META_MAXX = 10000                                  ' 10000
Private Const LL_META_MAXY = 10000                                  ' 10000
Private Const LL_TEXT = 268435456                                   ' 0x10000000
Private Const LL_TEXT_ALLOW_WORDWRAP = 268435456                    ' 0x10000000
Private Const LL_TEXT_DENY_WORDWRAP = 268435457                     ' 0x10000001
Private Const LL_TEXT_FORCE_WORDWRAP = 268435458                    ' 0x10000002
Private Const LL_NUMERIC = 134217728                                ' 0x08000000
Private Const LL_NUMERIC_LOCALIZED = 134217729                      ' 0x08000001
Private Const LL_NUMERIC_INTEGER = 134217730                        ' 0x08000002
Private Const LL_DATE = 67108864                                    ' 0x04000000
Private Const LL_DATE_METHODMASK = 255                              ' 0x000000ff
Private Const LL_DATE_DELPHI_1 = 67108865                           ' 0x04000001
Private Const LL_DATE_DELPHI = 67108866                             ' 0x04000002
Private Const LL_DATE_MS = 67108866                                 ' 0x04000002
Private Const LL_DATE_OLE = 67108866                                ' 0x04000002
Private Const LL_DATE_VFOXPRO = 67108867                            ' 0x04000003
Private Const LL_DATE_DMY = 67108868                                ' 0x04000004
Private Const LL_DATE_MDY = 67108869                                ' 0x04000005
Private Const LL_DATE_YMD = 67108870                                ' 0x04000006
Private Const LL_DATE_YYYYMMDD = 67108871                           ' 0x04000007
Private Const LL_DATE_LOCALIZED = 67108872                          ' 0x04000008
Private Const LL_DATE_JULIAN = 67108873                             ' 0x04000009
Private Const LL_DATE_CLARION = 67108874                            ' 0x0400000a
Private Const LL_DATE_YMD_AUTO = 67108880                           ' 0x04000010
Private Const LL_BOOLEAN = 33554432                                 ' 0x02000000
Private Const LL_RTF = 16777216                                     ' 0x01000000
Private Const LL_HTML = 8388608                                     ' 0x00800000
Private Const LL_PDF = 4194304                                      ' 0x00400000
Private Const LL_INPUTOBJECT = 2097152                              ' 0x00200000
Private Const LL_LLXOBJECT = 1048576                                ' 0x00100000
Private Const LL_SUBTABLELIST = 524288                              ' 0x00080000
Private Const LL_RETURN_ERROR_IF_FILE_NOT_FOUND = 65536             ' 0x00010000
Private Const LL_FIXEDNAME = 32768                                  ' 0x00008000
Private Const LL_NOSAVEAS = 16384                                   ' 0x00004000
Private Const LL_DESIGNER_OVER_CHILD = 8192                         ' 0x00002000
Private Const LL_EXPRCONVERTQUIET = 4096                            ' 0x00001000
Private Const LL_NONAMEINTITLE = 2048                               ' 0x00000800
Private Const LL_PRVOPT_PRN_USEDEFAULT = 0                          ' 0x00000000
Private Const LL_PRVOPT_PRN_ASKPRINTERIFNEEDED = 1                  ' 0x00000001
Private Const LL_PRVOPT_PRN_ASKPRINTERALWAYS = 2                    ' 0x00000002
Private Const LL_PRVOPT_PRN_ALWAYSUSEDEFAULT = 3                    ' 0x00000003
Private Const LL_PRVOPT_PRN_ASSIGNMASK = 3                          ' 0x00000003
Private Const LL_PRVOPT_FLAG_STANDALONEVIEWER = 16                  ' 0x00000010
Private Const LL_OPTION_COPIES = 0                                  ' 0
Private Const LL_OPTION_STARTPAGE = 1                               ' 1
Private Const LL_OPTION_PAGE = 1                                    ' 1
Private Const LL_OPTION_OFFSET = 2                                  ' 2
Private Const LL_OPTION_COPIES_SUPPORTED = 3                        ' 3
Private Const LL_OPTION_FIRSTPAGE = 5                               ' 5
Private Const LL_OPTION_LASTPAGE = 6                                ' 6
Private Const LL_OPTION_JOBPAGES = 7                                ' 7
Private Const LL_OPTION_PRINTORDER = 8                              ' 8
Private Const LL_PRNOPT_COPIES = 0                                  ' 0
Private Const LL_COPIES_HIDE = -32768                               ' -32768
Private Const LL_PRNOPT_STARTPAGE = 1                               ' 1
Private Const LL_PRNOPT_PAGE = 1                                    ' 1
Private Const LL_PAGE_HIDE = -32768                                 ' -32768
Private Const LL_PRNOPT_OFFSET = 2                                  ' 2
Private Const LL_PRNOPT_COPIES_SUPPORTED = 3                        ' 3
Private Const LL_PRNOPT_UNITS = 4                                   ' 4
Private Const LL_UNITS_MM_DIV_10 = 0                                ' 0
Private Const LL_UNITS_INCH_DIV_100 = 1                             ' 1
Private Const LL_UNITS_INCH_DIV_1000 = 2                            ' 2
Private Const LL_UNITS_SYSDEFAULT_LORES = 3                         ' 3
Private Const LL_UNITS_SYSDEFAULT = 4                               ' 4
Private Const LL_UNITS_MM_DIV_100 = 5                               ' 5
Private Const LL_UNITS_MM_DIV_1000 = 6                              ' 6
Private Const LL_UNITS_SYSDEFAULT_HIRES = 7                         ' 7
Private Const LL_PRNOPT_FIRSTPAGE = 5                               ' 5
Private Const LL_PRNOPT_LASTPAGE = 6                                ' 6
Private Const LL_PRNOPT_JOBPAGES = 7                                ' 7
Private Const LL_PRNOPT_PRINTORDER = 8                              ' 8
Private Const LL_PRINTORDER_HORZ_LTRB = 0                           ' 0
Private Const LL_PRINTORDER_VERT_LTRB = 1                           ' 1
Private Const LL_PRINTORDER_HORZ_RBLT = 2                           ' 2
Private Const LL_PRINTORDER_VERT_RBLT = 3                           ' 3
Private Const LL_PRNOPT_DEFPRINTERINSTALLED = 11                    ' 11
Private Const LL_PRNOPT_PRINTDLG_DESTMASK = 12                      ' 12
Private Const LL_DESTINATION_PRN = 1                                ' 1
Private Const LL_DESTINATION_PRV = 2                                ' 2
Private Const LL_DESTINATION_FILE = 4                               ' 4
Private Const LL_DESTINATION_EXTERN = 8                             ' 8
Private Const LL_DESTINATION_MSFAX = 16                             ' 16
Private Const LL_PRNOPT_PRINTDLG_DEST = 13                          ' 13
Private Const LL_PRNOPT_PRINTDLG_ONLYPRINTERCOPIES = 14             ' 14
Private Const LL_PRNOPT_JOBID = 17                                  ' 17
Private Const LL_PRNOPT_PAGEINDEX = 18                              ' 18
Private Const LL_PRNOPT_USES2PASS = 19                              ' 19
Private Const LL_PRNOPT_PAGERANGE_USES_ABSOLUTENUMBER = 20          ' 20
Private Const LL_PRNOPT_USEMEMORYMETAFILE = 22                      ' 22
Private Const LL_PRNOPT_PARTIALPREVIEW = 23                         ' 23
Private Const LL_PRNOPT_ADDITIONALPAGES_FOR_TOTAL = 24              ' 24
Private Const LL_PRNOPT_HAS_TOTALPAGES = 25                         ' 25
Private Const LL_PRNOPTSTR_PRINTDST_FILENAME = 0                    ' 0
Private Const LL_PRNOPTSTR_EXPORTDESCR = 1                          ' 1
Private Const LL_PRNOPTSTR_EXPORT = 2                               ' 2
Private Const LL_PRNOPTSTR_PRINTJOBNAME = 3                         ' 3
Private Const LL_PRNOPTSTR_PRESTARTDOCESCSTRING = 4                 ' 4
Private Const LL_PRNOPTSTR_POSTENDDOCESCSTRING = 5                  ' 5
Private Const LL_PRNOPTSTR_PRESTARTPAGEESCSTRING = 6                ' 6
Private Const LL_PRNOPTSTR_POSTENDPAGEESCSTRING = 7                 ' 7
Private Const LL_PRNOPTSTR_PRESTARTPROJECTESCSTRING = 8             ' 8
Private Const LL_PRNOPTSTR_POSTENDPROJECTESCSTRING = 9              ' 9
Private Const LL_PRNOPTSTR_PAGERANGES = 10                          ' 10
Private Const LL_PRNOPTSTR_ISSUERANGES = 11                         ' 11
Private Const LL_PRNOPTSTR_PREVIEWTITLE = 12                        ' 12
Private Const LL_PRINT_V1POINTX = 0                                 ' 0x00000000
Private Const LL_PRINT_NORMAL = 256                                 ' 0x00000100
Private Const LL_PRINT_PREVIEW = 512                                ' 0x00000200
Private Const LL_PRINT_STORAGE = 512                                ' 0x00000200
Private Const LL_PRINT_FILE = 1024                                  ' 0x00000400
Private Const LL_PRINT_USERSELECT = 2048                            ' 0x00000800
Private Const LL_PRINT_EXPORT = 2048                                ' 0x00000800
Private Const LL_PRINT_MODEMASK = 3840                              ' 0x00000f00
Private Const LL_PRINT_MULTIPLE_JOBS = 4096                         ' 0x00001000
Private Const LL_PRINT_KEEPJOB = 8192                               ' 0x00002000
Private Const LL_PRINT_IS_DOM_CALLER = 16384                        ' 0x00004000
Private Const LL_PRINT_DOM_NOCREATEDC = 65536                       ' 0x00010000
Private Const LL_PRINT_DOM_NOOBJECTLOAD = 131072                    ' 0x00020000
Private Const LL_PRINT_REMOVE_UNUSED_VARS = 32768                   ' 0x00008000
Private Const LL_BOXTYPE_BOXTYPEMASK = 255                          ' 0x000000ff
Private Const LL_BOXTYPE_NONE = 255                                 ' 0x000000ff
Private Const LL_BOXTYPE_FLAG_ALLOWSUSPEND = 1073741824             ' 0x40000000
Private Const LL_BOXTYPE_FLAG_USEMARQUEE = -2147483648#              ' 0x80000000
Private Const LL_BOXTYPE_NORMALMETER = 0                            ' 0
Private Const LL_BOXTYPE_BRIDGEMETER = 1                            ' 1
Private Const LL_BOXTYPE_NORMALWAIT = 2                             ' 2
Private Const LL_BOXTYPE_BRIDGEWAIT = 3                             ' 3
Private Const LL_BOXTYPE_EMPTYWAIT = 4                              ' 4
Private Const LL_BOXTYPE_EMPTYABORT = 5                             ' 5
Private Const LL_BOXTYPE_STDWAIT = 6                                ' 6
Private Const LL_BOXTYPE_STDABORT = 7                               ' 7
Private Const LL_BOXTYPE_MAX = 7                                    ' 7
Private Const LL_FILE_ALSONEW = 32768                               ' 0x00008000
Private Const LL_SELECTFILEDLGTITLE_USE_OSFILENAME = 65536          ' 0x00010000
Private Const LL_FILE_FORCE_OS_DIALOG = 131072                      ' 0x00020000
Private Const LL_FCTPARATYPE_DOUBLE = 1                             ' 0x0001
Private Const LL_FCTPARATYPE_DATE = 2                               ' 0x0002
Private Const LL_FCTPARATYPE_STRING = 4                             ' 0x0004
Private Const LL_FCTPARATYPE_BOOL = 8                               ' 0x0008
Private Const LL_FCTPARATYPE_DRAWING = 16                           ' 0x0010
Private Const LL_FCTPARATYPE_BARCODE = 32                           ' 0x0020
Private Const LL_FCTPARATYPE_ALL = 63                               ' 0x003f
Private Const LL_FCTPARATYPE_PARA1 = 32769                          ' 0x8001
Private Const LL_FCTPARATYPE_PARA2 = 32770                          ' 0x8002
Private Const LL_FCTPARATYPE_PARA3 = 32771                          ' 0x8003
Private Const LL_FCTPARATYPE_PARA4 = 32772                          ' 0x8004
Private Const LL_FCTPARATYPE_SAME = 32831                           ' 0x803f
Private Const LL_FCTPARATYPE_MASK = 36863                           ' 0x8fff
Private Const LL_FCTPARATYPEFLAG_NONULLCHECK = 65536                ' 0x00010000
Private Const LL_FCTPARATYPEFLAG_MULTIDIM_ALLOWED = 131072          ' 0x00020000
Private Const LL_FCTPARATYPEFLAG_RAW = 524288                       ' 0x00080000
Private Const LL_FCTPARATYPEFLAG_RAW_WITH_SYNTAXCHECK = 262144      ' 0x00040000
Private Const LL_FCTPARATYPEFLAG_KEEP_LINEBREAKS = 1048576          ' 0x00100000
Private Const LL_FCTPARATYPEFLAG_EXECUTE_ON_SYNTAXCHECK_AT_LOADTIME = 2097152    ' 0x00200000
Private Const LL_EXPRTYPE_DOUBLE = 1                                ' 1
Private Const LL_EXPRTYPE_DATE = 2                                  ' 2
Private Const LL_EXPRTYPE_STRING = 3                                ' 3
Private Const LL_EXPRTYPE_BOOL = 4                                  ' 4
Private Const LL_EXPRTYPE_DRAWING = 5                               ' 5
Private Const LL_EXPRTYPE_BARCODE = 6                               ' 6
Private Const LL_OPTION_NEWEXPRESSIONS = 0                          ' 0
Private Const LL_OPTION_ONLYONETABLE = 1                            ' 1
Private Const LL_OPTION_TABLE_COLORING = 2                          ' 2
Private Const LL_COLORING_LL = 0                                    ' 0
Private Const LL_COLORING_PROGRAM = 1                               ' 1
Private Const LL_COLORING_DONTCARE = 2                              ' 2
Private Const LL_OPTION_SUPERVISOR = 3                              ' 3
Private Const LL_OPTION_UNITS = 4                                   ' 4
Private Const LL_OPTION_TABSTOPS = 5                                ' 5
Private Const LL_TABS_DELETE = 0                                    ' 0
Private Const LL_TABS_EXPAND = 1                                    ' 1
Private Const LL_OPTION_CALLBACKMASK = 6                            ' 6
Private Const LL_CB_PAGE = 1073741824                               ' 0x40000000
Private Const LL_CB_PROJECT = 536870912                             ' 0x20000000
Private Const LL_CB_OBJECT = 268435456                              ' 0x10000000
Private Const LL_CB_HELP = 134217728                                ' 0x08000000
Private Const LL_CB_TABLELINE = 67108864                            ' 0x04000000
Private Const LL_CB_TABLEFIELD = 33554432                           ' 0x02000000
Private Const LL_CB_QUERY_IS_VARIABLE_OR_FIELD = 16777216           ' 0x01000000
Private Const LL_OPTION_CALLBACKPARAMETER = 7                       ' 7
Private Const LL_OPTION_HELPAVAILABLE = 8                           ' 8
Private Const LL_OPTION_SORTVARIABLES = 9                           ' 9
Private Const LL_OPTION_SUPPORTPAGEBREAK = 10                       ' 10
Private Const LL_OPTION_SHOWPREDEFVARS = 11                         ' 11
Private Const LL_OPTION_USEHOSTPRINTER = 13                         ' 13
Private Const LL_OPTION_EXTENDEDEVALUATION = 14                     ' 14
Private Const LL_OPTION_TABREPRESENTATIONCODE = 15                  ' 15
Private Const LL_OPTION_SHOWSTATE = 16                              ' 16
Private Const LL_OPTION_METRIC = 18                                 ' 18
Private Const LL_OPTION_ADDVARSTOFIELDS = 19                        ' 19
Private Const LL_OPTION_MULTIPLETABLELINES = 20                     ' 20
Private Const LL_OPTION_CONVERTCRLF = 21                            ' 21
Private Const LL_OPTION_WIZ_FILENEW = 22                            ' 22
Private Const LL_OPTION_RETREPRESENTATIONCODE = 23                  ' 23
Private Const LL_OPTION_PRVZOOM_PERC = 25                           ' 25
Private Const LL_OPTION_PRVRECT_LEFT = 26                           ' 26
Private Const LL_OPTION_PRVRECT_TOP = 27                            ' 27
Private Const LL_OPTION_PRVRECT_WIDTH = 28                          ' 28
Private Const LL_OPTION_PRVRECT_HEIGHT = 29                         ' 29
Private Const LL_OPTION_STORAGESYSTEM = 30                          ' 30
Private Const LL_STG_COMPAT4 = 0                                    ' 0
Private Const LL_STG_STORAGE = 1                                    ' 1
Private Const LL_OPTION_COMPRESSSTORAGE = 31                        ' 31
Private Const LL_STG_COMPRESS_THREADED = 32768                      ' 0x00008000
Private Const LL_STG_COMPRESS_UNTHREADED = 65536                    ' 0x00010000
Private Const LL_OPTION_NOPARAMETERCHECK = 32                       ' 32
Private Const LL_OPTION_NONOTABLECHECK = 33                         ' 33
Private Const LL_OPTION_DRAWFOOTERLINEONPRINT = 34                  ' 34
Private Const LL_OPTION_PRVZOOM_LEFT = 35                           ' 35
Private Const LL_OPTION_PRVZOOM_TOP = 36                            ' 36
Private Const LL_OPTION_PRVZOOM_WIDTH = 37                          ' 37
Private Const LL_OPTION_PRVZOOM_HEIGHT = 38                         ' 38
Private Const LL_OPTION_SPACEOPTIMIZATION = 40                      ' 40
Private Const LL_OPTION_REALTIME = 41                               ' 41
Private Const LL_OPTION_AUTOMULTIPAGE = 42                          ' 42
Private Const LL_OPTION_USEBARCODESIZES = 43                        ' 43
Private Const LL_OPTION_MAXRTFVERSION = 44                          ' 44
Private Const LL_OPTION_VARSCASESENSITIVE = 46                      ' 46
Private Const LL_OPTION_DELAYTABLEHEADER = 47                       ' 47
Private Const LL_OPTION_OFNDIALOG_EXPLORER = 48                     ' 48
Private Const LL_OPTION_OFN_NOPLACESBAR = 1073741824                ' 0x40000000
Private Const LL_OPTION_EMFRESOLUTION = 49                          ' 49
Private Const LL_OPTION_SETCREATIONINFO = 50                        ' 50
Private Const LL_OPTION_XLATVARNAMES = 51                           ' 51
Private Const LL_OPTION_LANGUAGE = 52                               ' 52
Private Const LL_OPTION_PHANTOMSPACEREPRESENTATIONCODE = 54         ' 54
Private Const LL_OPTION_LOCKNEXTCHARREPRESENTATIONCODE = 55         ' 55
Private Const LL_OPTION_EXPRSEPREPRESENTATIONCODE = 56              ' 56
Private Const LL_OPTION_DEFPRINTERINSTALLED = 57                    ' 57
Private Const LL_OPTION_CALCSUMVARSONINVISIBLELINES = 58            ' 58
Private Const LL_OPTION_NOFOOTERPAGEWRAP = 59                       ' 59
Private Const LL_OPTION_IMMEDIATELASTPAGE = 64                      ' 64
Private Const LL_OPTION_LCID = 65                                   ' 65
Private Const LL_OPTION_TEXTQUOTEREPRESENTATIONCODE = 66            ' 66
Private Const LL_OPTION_SCALABLEFONTSONLY = 67                      ' 67
Private Const LL_OPTION_NOTIFICATIONMESSAGEHWND = 68                ' 68
Private Const LL_OPTION_DEFDEFFONT = 69                             ' 69
Private Const LL_OPTION_CODEPAGE = 70                               ' 70
Private Const LL_OPTION_FORCEFONTCHARSET = 71                       ' 71
Private Const LL_OPTION_COMPRESSRTF = 72                            ' 72
Private Const LL_OPTION_ALLOW_LLX_EXPORTERS = 74                    ' 74
Private Const LL_OPTION_SUPPORTS_PRNOPTSTR_EXPORT = 75              ' 75
Private Const LL_OPTION_DEBUGFLAG = 76                              ' 76
Private Const LL_OPTION_SKIPRETURNATENDOFRTF = 77                   ' 77
Private Const LL_OPTION_INTERCHARSPACING = 78                       ' 78
Private Const LL_OPTION_INCLUDEFONTDESCENT = 79                     ' 79
Private Const LL_OPTION_RESOLUTIONCOMPATIBLETO9X = 80               ' 80
Private Const LL_OPTION_USECHARTFIELDS = 81                         ' 81
Private Const LL_OPTION_OFNDIALOG_NOPLACESBAR = 82                  ' 82
Private Const LL_OPTION_SKETCH_COLORDEPTH = 83                      ' 83
Private Const LL_OPTION_FINAL_TRUE_ON_LASTPAGE = 84                 ' 84
Private Const LL_OPTION_INTERCHARSPACING_FORCED = 86                ' 86
Private Const LL_OPTION_RTFAUTOINCREMENT = 87                       ' 87
Private Const LL_OPTION_UNITS_DEFAULT = 88                          ' 88
Private Const LL_OPTION_NO_MAPI = 89                                ' 89
Private Const LL_OPTION_TOOLBARSTYLE = 90                           ' 90
Private Const LL_OPTION_TOOLBARSTYLE_STANDARD = 0                   ' 0
Private Const LL_OPTION_TOOLBARSTYLE_OFFICEXP = 1                   ' 1
Private Const LL_OPTION_TOOLBARSTYLE_OFFICE2003 = 2                 ' 2
Private Const LL_OPTION_TOOLBARSTYLEMASK = 15                       ' 0x0f
Private Const LL_OPTION_TOOLBARSTYLEFLAG_GRADIENT = 128             ' 0x80
Private Const LL_OPTION_TOOLBARSTYLEFLAG_DOCKABLE = 64              ' 0x40
Private Const LL_OPTION_TOOLBARSTYLEFLAG_CANCLOSE = 32              ' 0x20
Private Const LL_OPTION_TOOLBARSTYLEFLAG_SHRINK_TO_FIT = 16         ' 0x10
Private Const LL_OPTION_MENUSTYLE = 91                              ' 91
Private Const LL_OPTION_MENUSTYLE_STANDARD_WITHOUT_BITMAPS = 0          ' 0
Private Const LL_OPTION_MENUSTYLE_STANDARD = 1                      ' 1
Private Const LL_OPTION_MENUSTYLE_OFFICEXP = 2                      ' 2
Private Const LL_OPTION_MENUSTYLE_OFFICE2003 = 3                    ' 3
Private Const LL_OPTION_RULERSTYLE = 92                             ' 92
Private Const LL_OPTION_RULERSTYLE_FLAT = 16                        ' 0x10
Private Const LL_OPTION_RULERSTYLE_GRADIENT = 128                   ' 0x80
Private Const LL_OPTION_STATUSBARSTYLE = 93                         ' 93
Private Const LL_OPTION_STATUSBARSTYLE_STANDARD = 0                 ' 0
Private Const LL_OPTION_STATUSBARSTYLE_OFFICEXP = 1                 ' 1
Private Const LL_OPTION_STATUSBARSTYLE_OFFICE2003 = 2               ' 2
Private Const LL_OPTION_TABBARSTYLE = 94                            ' 94
Private Const LL_OPTION_TABBARSTYLE_STANDARD = 0                    ' 0
Private Const LL_OPTION_TABBARSTYLE_OFFICEXP = 1                    ' 1
Private Const LL_OPTION_TABBARSTYLE_OFFICE2003 = 2                  ' 2
Private Const LL_OPTION_DROPWINDOWSTYLE = 95                        ' 95
Private Const LL_OPTION_DROPWINDOWSTYLE_STANDARD = 0                ' 0
Private Const LL_OPTION_DROPWINDOWSTYLE_OFFICEXP = 1                ' 1
Private Const LL_OPTION_DROPWINDOWSTYLE_OFFICE2003 = 2              ' 2
Private Const LL_OPTION_DROPWINDOWSTYLEMASK = 15                    ' 0x0f
Private Const LL_OPTION_DROPWINDOWSTYLEFLAG_CANCLOSE = 32           ' 0x20
Private Const LL_OPTION_INTERFACEWRAPPER = 96                       ' 96
Private Const LL_OPTION_FONTQUALITY = 97                            ' 97
Private Const LL_OPTION_FONTPRECISION = 98                          ' 98
Private Const LL_OPTION_UISTYLE = 99                                ' 99
Private Const LL_OPTION_UISTYLE_STANDARD = 0                        ' 0
Private Const LL_OPTION_UISTYLE_OFFICEXP = 1                        ' 1
Private Const LL_OPTION_UISTYLE_OFFICE2003 = 2                      ' 2
Private Const LL_OPTION_NOFILEVERSIONUPGRADEWARNING = 100           ' 100
Private Const LL_OPTION_UPDATE_FOOTER_ON_DATALINEBREAK_AT_FIRST_LINE = 101        ' 101
Private Const LL_OPTION_ESC_CLOSES_PREVIEW = 102                    ' 102
Private Const LL_OPTION_VIEWER_ASSUMES_TEMPFILE = 103               ' 103
Private Const LL_OPTION_CALC_USED_VARS = 104                        ' 104
Private Const LL_OPTION_NOPRINTJOBSUPERVISION = 106                 ' 106
Private Const LL_OPTION_CALC_SUMVARS_ON_PARTIAL_LINES = 107         ' 107
Private Const LL_OPTION_BLACKNESS_SCM = 108                         ' 108
Private Const LL_OPTION_PROHIBIT_USERINTERACTION = 109              ' 109
Private Const LL_OPTION_PERFMON_INSTALL = 110                       ' 110
Private Const LL_OPTION_RESERVED111 = 111                           ' 111
Private Const LL_OPTION_VARLISTBUCKETCOUNT = 112                    ' 112
Private Const LL_OPTION_MSFAXALLOWED = 113                          ' 113
Private Const LL_OPTION_AUTOPROFILINGTICKS = 114                    ' 114
Private Const LL_OPTION_PROJECTBACKUP = 115                         ' 115
Private Const LL_OPTION_ERR_ON_FILENOTFOUND = 116                   ' 116
Private Const LL_OPTION_NOFAXVARS = 117                             ' 117
Private Const LL_OPTION_NOMAILVARS = 118                            ' 118
Private Const LL_OPTION_PATTERNRESCOMPATIBILITY = 119               ' 119
Private Const LL_OPTION_NODELAYEDVALUECACHING = 120                 ' 120
Private Const LL_OPTION_FEATURE = 1000                              ' 1000
Private Const LL_OPTION_FEATURE_CLEARALL = 0                        ' 0
Private Const LL_OPTION_FEATURE_SUPPRESS_JPEG_DISPLAY = 1           ' 1
Private Const LL_OPTION_FEATURE_SUPPRESS_JPEG_CREATION = 2          ' 2
Private Const LL_OPTION_VARLISTDISPLAY = 121                        ' 121
Private Const LL_OPTION_VARLISTDISPLAY_VARSORT_DECLARATIONORDER = 0          ' 0x0000
Private Const LL_OPTION_VARLISTDISPLAY_VARSORT_ALPHA = 1            ' 0x0001
Private Const LL_OPTION_VARLISTDISPLAY_VARSORT_MASK = 15            ' 0x000f
Private Const LL_OPTION_VARLISTDISPLAY_FOLDERPOS_DECLARATIONORDER = 0          ' 0x0000
Private Const LL_OPTION_VARLISTDISPLAY_FOLDERPOS_ALPHA = 16         ' 0x0010
Private Const LL_OPTION_VARLISTDISPLAY_FOLDERPOS_TOP = 32           ' 0x0020
Private Const LL_OPTION_VARLISTDISPLAY_FOLDERPOS_BOTTOM = 48         ' 0x0030
Private Const LL_OPTION_VARLISTDISPLAY_FOLDERPOS_MASK = 240         ' 0x00f0
Private Const LL_OPTION_VARLISTDISPLAY_LLFOLDERPOS_BOTTOM = 256        ' 0x0100
Private Const LL_OPTION_WORKAROUND_RTFBUG_EMPTYFIRSTPAGE = 122        ' 122
Private Const LL_OPTION_FORMULASTRINGCOMPARISONS_CASESENSITIVE = 123        ' 123
Private Const LL_OPTION_FIELDS_IN_PROJECTPARAMETERS = 124           ' 124
Private Const LL_OPTION_CHECKWINDOWTHREADEDNESS = 125               ' 125
Private Const LL_OPTION_ISUSED_WILDCARD_AT_START = 126              ' 126
Private Const LL_OPTION_ROOT_MUST_BE_MASTERTABLE = 127              ' 127
Private Const LL_OPTION_DLLTYPE = 128                               ' 128
Private Const LL_OPTION_DLLTYPE_32BIT = 1                           ' 0x0001
Private Const LL_OPTION_DLLTYPE_64BIT = 2                           ' 0x0002
Private Const LL_OPTION_DLLTYPE_BITMASK = 15                        ' 0x000f
Private Const LL_OPTION_DLLTYPE_SDBCS = 16                          ' 0x0010
Private Const LL_OPTION_DLLTYPE_UNICODE = 32                        ' 0x0020
Private Const LL_OPTION_DLLTYPE_CHARSET = 240                       ' 0x00f0
Private Const LL_OPTION_HLIBRARY = 129                              ' 129
Private Const LL_OPTION_INVERTED_PAGEORIENTATION = 130              ' 130
Private Const LL_OPTION_ENABLE_STANDALONE_DATACOLLECTING_OBJECTS = 131        ' 131
Private Const LL_OPTION_USERVARS_ARE_CODESNIPPETS = 132             ' 132
Private Const LL_OPTION_STORAGE_ADD_SUMMARYINFORMATION = 133        ' 133
Private Const LL_OPTION_INCREMENTAL_PREVIEW = 135                   ' 135
Private Const LL_OPTION_RELAX_AT_SHUTDOWN = 136                     ' 136
Private Const LL_OPTION_NOPRINTERPATHCHECK = 137                    ' 137
Private Const LL_OPTION_SUPPORT_HUGESTORAGEFS = 138                 ' 138
Private Const LL_OPTION_NOAUTOPROPERTYCORRECTION = 139              ' 139
Private Const LL_OPTION_NOVARLISTRESET_ON_RESETPROJECTSTATE = 140        ' 140
Private Const LL_OPTION_DESIGNERPREVIEWPARAMETER = 141              ' 141
Private Const LL_OPTION_RESERVED142 = 142                           ' 142
Private Const LL_OPTION_DESIGNEREXPORTPARAMETER = 143               ' 143
Private Const LL_OPTION_DESIGNERPRINT_SINGLETHREADED = 144          ' 144
Private Const LL_OPTION_ALLOW_COMMENTS_IN_FORMULA = 145             ' 145
Private Const LL_OPTION_USE_MLANG_LINEBREAKALGORITHM = 146          ' 146
Private Const LL_OPTION_USE_JPEG_OR_PNG_OPTIMIZATION = 147          ' 147
Private Const LL_OPTION_ENABLE_IMAGESMOOTHING = 148                 ' 148
Private Const LL_OPTION_MAXRTFVERSION_AVAILABLE = 159               ' 159
Private Const LL_OPTION_CONDREPRESENTATIONCODES_LIKE_ANSI = 160        ' 160
Private Const LL_OPTION_NULL_IS_NONDESTRUCTIVE = 161                ' 161
Private Const LL_OPTION_DRILLDOWNPARAMETER = 162                    ' 162
Private Const LL_OPTION_ROUNDINGSTRATEGY = 163                      ' 163
Private Const LL_ROUNDINGSTRATEGY_BANKERSROUNDING = 0               ' 0
Private Const LL_ROUNDINGSTRATEGY_ARITHMETIC_SYMMETRIC = 1          ' 1
Private Const LL_OPTION_RESERVED164 = 164                           ' 164
Private Const LL_OPTION_RESERVED165 = 165                           ' 165
Private Const LL_OPTION_PICTURE_TRANSPARENCY_IS_WHITE = 166         ' 166
Private Const LL_OPTION_FLOATPRECISION = 167                        ' 167
Private Const LL_OPTION_SUPPRESS_LRUENTRY = 168                     ' 168
Private Const LL_OPTION_FORCEFIRSTGROUPHEADER = 169                 ' 169
Private Const LL_OPTION_SUPPORT_PDFINPUTFIELDS = 170                ' 170
Private Const LL_OPTION_ENHANCED_SKIPRETURNATENDOFRTF = 171         ' 171
Private Const LL_OPTION_HIERARCHICALDATASOURCE = 172                ' 172
Private Const LL_OPTION_FORCE_HEADER_EVEN_ON_LARGE_FOOTERLINES = 173        ' 173
Private Const LL_OPTION_PRINTERDEVICEOPTIMIZATION = 174             ' 174
Private Const LL_OPTION_RTFHEIGHTSCALINGPERCENTAGE = 175            ' 175
Private Const LL_OPTION_FORCE_DEFAULT_PRINTER_IN_PREVIEW = 176        ' 176
Private Const LL_OPTION_SAVE_PROJECT_IN_UTF8 = 178                  ' 178
Private Const LL_DRILLDOWNFILTERSTRATEGY_ALLOW_ONLY_SUBTABLES = 0          ' 0
Private Const LL_DRILLDOWNFILTERSTRATEGY_ALLOW_ALL_TABLES = 1          ' 1
Private Const LL_DRILLDOWNFILTERSTRATEGY_ALLOW_SUBTABLES_AND_UNRELATED = 2          ' 2
Private Const LL_DRILLDOWNFILTERSTRATEGY_ALLOW_SUBTABLES_AND_USERDEFINED = 3          ' 3
Private Const LL_DRILLDOWNFILTERSTRATEGY_MASK = 15                  ' 0x0f
Private Const LL_DRILLDOWNFILTERFLAG_OFFER_BASERECORD_AS_VARIABLES = 16         ' 0x10
Private Const LL_OPTION_DRILLDOWN_DATABASEFILTERING = 179           ' 179
Private Const LL_OPTION_SUPPRESS_TASKBARBUTTON_PROGRESSSTATE = 180        ' 180
Private Const LL_OPTION_PRINTDLG_DEVICECHANGE_KEEPS_DEVMODESETTINGS = 181        ' 181
Private Const LL_OPTION_DRILLDOWN_SUPPORTS_EMBEDDING = 182          ' 182
Private Const LL_VARLISTCLEARSTRATEGY_EMPTY_LIST = 0                ' 0
Private Const LL_VARLISTCLEARSTRATEGY_SET_NULL = 1                  ' 1
Private Const LL_VARLISTCLEARSTRATEGY_SET_DEFAULT = 2               ' 2
Private Const LL_OPTION_VARLISTCLEARSTRATEGY_ON_DEFINE_START = 183        ' 183
Private Const LL_OPTION_RESERVED184 = 184                           ' 184
Private Const LL_OPTION_KEEP_EMPTY_SUM_VARS = 185                   ' 185
Private Const LL_OPTION_RESERVED187 = 187                           ' 187
Private Const LL_OPTION_DEFAULTDECSFORSTR = 188                     ' 188
Private Const LL_OPTION_RESETPROJECTSTATE_FORCES_NEW_PRINTJOB = 189        ' 189
Private Const LL_OPTION_DEFINEXXXSTART_COMPATIBLE_TO_PRE15 = 190        ' 190
Private Const LL_OPTION_RESETPROJECTSTATE_FORCES_NEW_DC = 191        ' 191
Private Const LL_OPTION_BITMAP_RESOLUTION_FOR_PREVIEW = 192         ' 192
Private Const LL_OPTION_DRAW_EMPTY_CHARTOBJECTS = 193               ' 193
Private Const LL_OPTION_PREVIOUS_DEFAULTS_TO_NULL = 194             ' 194
Private Const LL_OPTION_FORCE_IMAGEEMBEDDING = 195                  ' 195
Private Const LL_OPTION_VARKEY_MAP_SHARP_S_TO_SS = 196              ' 196
Private Const LL_OPTION_NO_LAYERED_WINDOWS = 197                    ' 197
Private Const LL_OPTION_SCALED_PERCENTAGEFORMATTER = 198            ' 198
Private Const LL_OPTION_USE_ANTIALIAS = 199                         ' 199
Private Const LL_OPTION_FORCETABLELINECALLBACK = 200                ' 200
Private Const LL_OPTION_EXPORTCONSUMER = 201                        ' 201
Private Const LL_OPTION_TOC_IDX_ITEMID = 202                        ' 202
Private Const LL_OPTION_FORCED2PASSMODE = 203                       ' 203
Private Const LL_OPTION_SETVAR_ONLY_SETS_IF_CONTAINER_PRINTS = 204        ' 204
Private Const LL_OPTION_SHOW_PREVIEW_AFTER_PRINT_END = 206          ' 206
Private Const LL_OPTION_PROPLIST_COMBOBOX_SCROLL_WRAPS = 207        ' 207
Private Const LL_OPTION_ALWAYS_CALC_GROUPCHANGE_CONDITION = 208        ' 208
Private Const LL_OPTION_NULLHANDLING_SUPPORTED_IN_ENHMODE = 209        ' 209
Private Const LL_OPTION_RESETPROJECTSTATE_FORCES_NEW_PREVIEWJOB = 210        ' 210
Private Const LL_OPTION_USE_LEGACY_WORDWRAPPINGALGORITHM = 211        ' 211
Private Const LL_OPTION_PREVIEW_USES_PRINTTHREAD = 212              ' 212
Private Const LL_OPTION_LL_SUPPLIES_MESSAGELOOP_WHILE_PRINTING_TO_PREVIEW = 213        ' 213
Private Const LL_OPTION_PRINTERDCCACHE_TIMEOUT_SEC = 214            ' 214
Private Const LL_OPTION_DESIGNER_RIBBONBACKGROUNDCOLOR = 215        ' 215
Private Const LL_OPTION_INTERNAL_EMFCLEANUP = 216                   ' 216
Private Const LL_OPTION_RIBBON_DEFAULT_ENABLEDSTATE = 217           ' 217
Private Const LL_OPTION_PRVFILEVERSION = 218                        ' 218
Private Const LL_OPTION_TRY_REDUCE_BMPSIZE_BY_CONVERTING_TO_PNG_OR_JPEG = 219        ' 219
Private Const LL_OPTION_NO_IMAGEFILEOPTIMIZATION = 220              ' 220
Private Const LL_OPTION_NO_ENFORCED_GROUPFOOTERPRIORITY_FOR_LAST_GROUPFOOTER = 221        ' 221
Private Const LL_OPTION_ALLOW_COMBINED_COLLECTING_OF_DATA_FOR_COLLECTIONCONTROLS = 222        ' 222
Private Const LL_OPTION_SUPPRESS_LOADERRORMESSAGES = 223            ' 223
Private Const LL_OPTION_IGNOREFORMULARESULTMISMATCH_AT_LOADTIME = 224        ' 224
Private Const LL_OPTION_MAX_SIZE_OF_PROJECTINFOCACHE = 225          ' 225
Private Const LL_OPTION_NO_CORRECTION_OF_UNICODE_RTF = 226          ' 226
Private Const LL_OPTION_MAY_RELEASE_UNNECESSARY_PROPS_AT_PRINTTIME = 227        ' 227
Private Const LL_OPTION_DO_NOT_RESTORE_PREVSTATE_ON_FILTER_MISMATCH = 228        ' 228
Private Const LL_OPTION_SUPPORT_USERDEFINED_REPORTPAGELAYOUT = 229        ' 229
Private Const LL_OPTION_DESIGNER_RIBBONTEXTCOLOR = 230              ' 230
Private Const LL_PARTSHARINGFLAG_VARIABLES_TOC = 1                  ' 0x01
Private Const LL_PARTSHARINGFLAG_VARIABLES_IDX = 2                  ' 0x02
Private Const LL_PARTSHARINGFLAG_VARIABLES_GTC = 4                  ' 0x04
Private Const LL_OPTION_PARTSHARINGFLAGS = 231                      ' 231
Private Const LL_OPTION_PIECHARTORDER_COMPATIBLE_TO_PRE19 = 232        ' 232
Private Const LL_OPTION_DATABASESTRUCTURE_SORT_DECLARATIONORDER = 233        ' 233
Private Const LL_OPTION_REPORT_PARAMETERS_REALDATAJOBPARAMETER = 234        ' 234
Private Const LL_OPTION_EXPANDABLE_REGIONS_REALDATAJOBPARAMETER = 235        ' 235
Private Const LL_OPTION_IMPROVED_TABLELINEANCHORING = 236           ' 236
Private Const LL_OPTION_INTERACTIVESORTING_REALDATAJOBPARAMETER = 237        ' 237
Private Const LL_OPTION_TEMPFILESTRATEGY = 238                      ' 238
Private Const LL_TEMPFILESTRATEGY_SPEED = 0                         ' 0
Private Const LL_TEMPFILESTRATEGY_SIZE = 1                          ' 1
Private Const LL_TEMPFILESTRATEGY_SECURITY = 2                      ' 2
Private Const LL_OPTION_RTF_WHITE_BACKGROUND_IS_TRANSPARENT = 239        ' 239
Private Const LL_OPTION_NO_DOTTED_LINE_ON_SECONDARY_AXIS = 240        ' 240
Private Const LL_OPTION_NO_PREVIOUS_VARLIST = 241                   ' 241
Private Const LL_OPTION_COMMIT_FILE_ON_SAVE = 242                   ' 242
Private Const LL_OPTION_DO_NOT_RTRIM_CELLTEXT = 243                 ' 243
Private Const LL_OPTION_ALLOW_FCT_TEXTWIDTH = 244                   ' 244
Private Const LL_OPTION_PASTEOBJECTS_TO_FIRST_VISIBLE_LAYER = 245        ' 245
Private Const LL_OPTION_EMPTY_FILE_TRIGGERS_PROJECT_WIZARD = 246        ' 246
Private Const LL_OPTION_DELAY_UPDATE_REMAININGTABLESPACE = 247        ' 247
Private Const LL_OPTION_WIZARD_ADDS_ORGNAME_TO_UI = 248             ' 248
Private Const LL_OPTION_PROHIBIT_EXTERNAL_FILES = 249               ' 249
Private Const LL_OPTION_DRAWINGS_INLINED = 250                      ' 250
Private Const LL_OPTION_SERIALIZE_PRINTAPI = 251                    ' 251
Private Const LL_OPTION_PROJECTFILELOCKTIMEOUT_IN_MS = 252          ' 252
Private Const LL_OPTION_ILLDATAPROVIDER = 253                       ' 253
Private Const LL_OPTION_RTF_SUPPORTS_PARABREAKOPTIONS = 254         ' 254
Private Const LL_OPTION_FORCE_PDFEMBEDDING = 255                    ' 255
Private Const LL_OPTION_IGNORE_NONSCALEABLEFONTPROPERTIES = 256        ' 256
Private Const LL_DATAPROVIDERTHREADNESS_NONE = 0                    ' 0
Private Const LL_DATAPROVIDERTHREADNESS_ONE_INSTANCE_PER_THREAD = 1          ' 1
Private Const LL_DATAPROVIDERTHREADNESS_DONTCARE = 2                ' 2
Private Const LL_OPTION_DATAPROVIDER_THREADEDNESS = 257             ' 257
Private Const LL_OPTION_SUBREPORT_BASE = 258                        ' 258
Private Const LL_OPTION_SUBREPORT_CLIENT = 259                      ' 259
Private Const LL_OPTION_NO_IPICTURE_SUPPORT = 260                   ' 260
Private Const LL_OPTION_FORCE_JPEG_RECOMPRESSION = 261              ' 261
Private Const LL_OPTION_TEXTWRAP_TOLERANCE_PERC = 262               ' 262
Private Const LL_OPTION_NO_USERVARCHECK_ON_LOAD = 263               ' 263
Private Const LL_OPTION_TOC_IDX_PAGE = 264                          ' 264
Private Const LL_OPTION_RTF_TAB_KEY_IS_TAB_FORMATTER = 265          ' 265
Private Const LL_OPTION_VARLISTDISPLAY_LL_FOLDER_AT_END = 266        ' 266
Private Const LL_OPTION_DOM_DO_NOT_KILL_EMPTY_TABLE = 267           ' 267
Private Const LL_OPTION_ENABLE_INPUTOBJECTS_IN_TABLES = 268         ' 268
Private Const LL_OPTION_MAX_ENTRIES_FOR_AUTOCOMPLETE = 269          ' 269
Private Const LL_OPTION_DEFAULT_FOR_SHADOWPAGEWRAP = 270            ' 270
Private Const LL_OPTION_MAX_UNDO_STEPS = 271                        ' 271
Private Const LL_OPTION_HTML_USE_MAILFORMAT = 272                   ' 272
Private Const LL_OPTION_CLIP_LABELS_TO_PROJECTAREA = 273            ' 273
Private Const LL_OPTION_RESETPROJECTSTATE_FORCES_NEW_EXPORTJOB = 274        ' 274
Private Const LL_OPTION_SCRIPTENGINE_ENABLED = 276                  ' 276
Private Const LL_OPTION_SCRIPTENGINE_TIMEOUTMS = 277                ' 277
Private Const LL_OPTION_SCRIPTENGINE_AUTOEXECUTE = 278              ' 278
Private Const LL_OPTION_SHAPEFILE_TIMEOUTMS = 279                   ' 279
Private Const LL_OPTION_COUNTALLPRINTEDDATA_LASTPRINT = 280         ' 280
Private Const LL_OPTION_SAVE_AS_ACTS_AS_EXPORT = 281                ' 281
Private Const LL_OPTION_RESETPROJECTSTATE_TRIGGERS_NEW_SHEET = 282        ' 282
Private Const LL_OPTION_HIDE_EXPORT_TAB_FROM_LAYOUT_CONFIG = 283        ' 283
Private Const LL_OPTION_USE_VARLIST_NAMESORTINDEXCACHE = 284        ' 284
Private Const LL_OPTION_NOCONTRASTOPTIMIZATION = 285                ' 285
Private Const LL_OPTION_AUTORECOVERY_DISABLED = 286                 ' 286
Private Const LL_OPTION_AUTORECOVERY_SAVEOPTIONS = 287              ' 287
Private Const LL_AUTORECOVERY_SAVEOPTIONS_NEWFILE = 1               ' 1
Private Const LL_AUTORECOVERY_SAVEOPTIONS_OVERWRITE = 2             ' 2
Private Const LL_OPTION_LINK_PRINTERQUEUES = 288                    ' 288
Private Const LL_OPTION_FORCE_RTFMERGING = 289                      ' 289
Private Const LL_OPTION_W201512300001 = 290                         ' 290
Private Const LL_OPTION_CALCLINEHEIGHT_COMPATIBLE_TO_19 = 291        ' 291
Private Const LL_OPTION_FORCE_UNIQUE_PARAMETERUISTRING = 292        ' 292
Private Const LL_OPTION_AUTOMATICFOOTER = 293                       ' 293
Private Const LL_OPTION_SUPPORT_PREDEFINED_COLORS = 294             ' 294
Private Const LL_OPTION_FAVORITE_SETTINGS = 295                     ' 295
Private Const LL_FAVORITES_ENABLE_FAVORITES_BY_DEFAULT = 1          ' 0x0001
Private Const LL_FAVORITES_HIDE_FAVORITES_BUTTON = 2                ' 0x0002
Private Const LL_OPTION_NEWMODE_EXPRSTARTREPRESENTATIONCODE = 296        ' 296
Private Const LL_OPTION_NEWMODE_EXPRENDREPRESENTATIONCODE = 297        ' 297
Private Const LL_OPTION_RESERVED_298 = 298                          ' 298
Private Const LL_OPTION_ILLREPOSITORY = 299                         ' 299
Private Const LL_OPTION_VARLISTLOOKUP_ALLOWS_GLOBALNAME = 300        ' 300
Private Const LL_OPTION_FORCESAVEDESIGNSCHEME = 301                 ' 301
Private Const LL_OPTION_REPOSITORY_SINGLEPROJECTMODE = 302          ' 302
Private Const LL_OPTION_ANIMATIONS_DISABLED = 303                   ' 303
Private Const LL_OPTION_UPDATE_INTERACTIONINFO_PER_PAGE = 304        ' 304
Private Const LL_OPTION_IS_REPORTSERVERDESIGNER = 305               ' 305
Private Const LL_OPTIONSTR_LABEL_PRJEXT = 0                         ' 0
Private Const LL_OPTIONSTR_LABEL_PRVEXT = 1                         ' 1
Private Const LL_OPTIONSTR_LABEL_PRNEXT = 2                         ' 2
Private Const LL_OPTIONSTR_CARD_PRJEXT = 3                          ' 3
Private Const LL_OPTIONSTR_CARD_PRVEXT = 4                          ' 4
Private Const LL_OPTIONSTR_CARD_PRNEXT = 5                          ' 5
Private Const LL_OPTIONSTR_LIST_PRJEXT = 6                          ' 6
Private Const LL_OPTIONSTR_LIST_PRVEXT = 7                          ' 7
Private Const LL_OPTIONSTR_LIST_PRNEXT = 8                          ' 8
Private Const LL_OPTIONSTR_LLXPATHLIST = 12                         ' 12
Private Const LL_OPTIONSTR_SHORTDATEFORMAT = 13                     ' 13
Private Const LL_OPTIONSTR_DECIMAL = 14                             ' 14
Private Const LL_OPTIONSTR_THOUSAND = 15                            ' 15
Private Const LL_OPTIONSTR_CURRENCY = 16                            ' 16
Private Const LL_OPTIONSTR_EXPORTS_AVAILABLE = 17                   ' 17
Private Const LL_OPTIONSTR_EXPORTS_ALLOWED = 18                     ' 18
Private Const LL_OPTIONSTR_DEFDEFFONT = 19                          ' 19
Private Const LL_OPTIONSTR_EXPORTFILELIST = 20                      ' 20
Private Const LL_OPTIONSTR_VARALIAS = 21                            ' 21
Private Const LL_OPTIONSTR_MAILTO = 24                              ' 24
Private Const LL_OPTIONSTR_MAILTO_CC = 25                           ' 25
Private Const LL_OPTIONSTR_MAILTO_BCC = 26                          ' 26
Private Const LL_OPTIONSTR_MAILTO_SUBJECT = 27                      ' 27
Private Const LL_OPTIONSTR_SAVEAS_PATH = 28                         ' 28
Private Const LL_OPTIONSTR_LABEL_PRJDESCR = 29                      ' 29
Private Const LL_OPTIONSTR_CARD_PRJDESCR = 30                       ' 30
Private Const LL_OPTIONSTR_LIST_PRJDESCR = 31                       ' 31
Private Const LL_OPTIONSTR_LLFILEDESCR = 32                         ' 32
Private Const LL_OPTIONSTR_PROJECTPASSWORD = 33                     ' 33
Private Const LL_OPTIONSTR_FAX_RECIPNAME = 34                       ' 34
Private Const LL_OPTIONSTR_FAX_RECIPNUMBER = 35                     ' 35
Private Const LL_OPTIONSTR_FAX_QUEUENAME = 36                       ' 36
Private Const LL_OPTIONSTR_FAX_SENDERNAME = 37                      ' 37
Private Const LL_OPTIONSTR_FAX_SENDERCOMPANY = 38                   ' 38
Private Const LL_OPTIONSTR_FAX_SENDERDEPT = 39                      ' 39
Private Const LL_OPTIONSTR_FAX_SENDERBILLINGCODE = 40               ' 40
Private Const LL_OPTIONSTR_FAX_AVAILABLEQUEUES = 42                 ' 42
Private Const LL_OPTIONSTR_LOGFILEPATH = 43                         ' 43
Private Const LL_OPTIONSTR_LICENSINGINFO = 44                       ' 44
Private Const LL_OPTIONSTR_PRINTERALIASLIST = 45                    ' 45
Private Const LL_OPTIONSTR_PREVIEWFILENAME = 46                     ' 46
Private Const LL_OPTIONSTR_EXPORTS_ALLOWED_IN_PREVIEW = 47          ' 47
Private Const LL_OPTIONSTR_HELPFILENAME = 48                        ' 48
Private Const LL_OPTIONSTR_NULLVALUE = 49                           ' 49
Private Const LL_OPTIONSTR_DEFAULT_EXPORT = 50                      ' 50
Private Const LL_OPTIONSTR_ORIGINALPROJECTFILENAME = 51             ' 51
Private Const LL_OPTIONSTR_HIERARCHICALDATASOURCE_ROOT = 52         ' 52
Private Const LL_OPTIONSTR_PRINTERDEFINITIONFILENAME = 53           ' 53
Private Const LL_OPTIONSTR_DOCINFO_DATATYPE = 54                    ' 54
Private Const LL_OPTIONSTR_IDX_PRJEXT = 55                          ' 55
Private Const LL_OPTIONSTR_IDX_PRVEXT = 56                          ' 56
Private Const LL_OPTIONSTR_IDX_PRNEXT = 57                          ' 57
Private Const LL_OPTIONSTR_TOC_PRJDESCR = 58                        ' 58
Private Const LL_OPTIONSTR_IDX_PRJDESCR = 59                        ' 59
Private Const LL_OPTIONSTR_TOC_PRJEXT = 60                          ' 60
Private Const LL_OPTIONSTR_TOC_PRVEXT = 61                          ' 61
Private Const LL_OPTIONSTR_TOC_PRNEXT = 62                          ' 62
Private Const LL_OPTIONSTR_DEFAULTSCHEME = 63                       ' 63
Private Const LL_OPTIONSTR_DEFAULTPROJECTNAME = 64                  ' 64
Private Const LL_OPTIONSTR_GTC_PRJEXT = 65                          ' 65
Private Const LL_OPTIONSTR_GTC_PRVEXT = 66                          ' 66
Private Const LL_OPTIONSTR_GTC_PRNEXT = 67                          ' 67
Private Const LL_OPTIONSTR_GTC_PRJDESCR = 68                        ' 68
Private Const LL_OPTIONSTR_ERRORTEXT_FROM_EXPORT = 69               ' 69
Private Const LL_OPTIONSTR_DEFAULTPRJDESCR = 70                     ' 70
Private Const LL_OPTIONSTR_DEFAULTPRINTER = 71                      ' 71
Private Const LL_OPTIONSTR_QUERY_LICENSINGINFO = 72                 ' 72
Private Const LL_OPTIONSTR_RESERVED73 = 73                          ' 73
Private Const LL_OPTIONSTR_REPRESENTATION_BOOL_TRUE = 74            ' 74
Private Const LL_OPTIONSTR_REPRESENTATION_BOOL_FALSE = 75           ' 75
Private Const LL_OPTIONSTR_DEFAULT_FILENAME_FOR_SAVEAS = 76         ' 76
Private Const LL_OPTIONSTR_LABEL_PRJDESCR_SINGULAR = 77             ' 77
Private Const LL_OPTIONSTR_LIST_PRJDESCR_SINGULAR = 78              ' 78
Private Const LL_OPTIONSTR_CARD_PRJDESCR_SINGULAR = 79              ' 79
Private Const LL_OPTIONSTR_TOC_PRJDESCR_SINGULAR = 80               ' 80
Private Const LL_OPTIONSTR_IDX_PRJDESCR_SINGULAR = 81               ' 81
Private Const LL_OPTIONSTR_GTC_PRJDESCR_SINGULAR = 82               ' 82
Private Const LL_OPTIONSTR_DEFAULTIMAGEPATH = 83                    ' 83
Private Const LL_SYSCOMMAND_MINIMIZE = -1                           ' -1
Private Const LL_SYSCOMMAND_MAXIMIZE = -2                           ' -2
Private Const LL_DLGBOXMODE_3DBUTTONS = 32768                       ' 0x8000
Private Const LL_DLGBOXMODE_3DFRAME2 = 16384                        ' 0x4000
Private Const LL_DLGBOXMODE_3DFRAME = 4096                          ' 0x1000
Private Const LL_DLGBOXMODE_NOBITMAPS = 8192                        ' 0x2000
Private Const LL_DLGBOXMODE_DONTCARE = 0                            ' 0x0000
Private Const LL_DLGBOXMODE_SAA = 1                                 ' 0x0001
Private Const LL_DLGBOXMODE_ALT1 = 2                                ' 0x0002
Private Const LL_DLGBOXMODE_ALT2 = 3                                ' 0x0003
Private Const LL_DLGBOXMODE_ALT3 = 4                                ' 0x0004
Private Const LL_DLGBOXMODE_ALT4 = 5                                ' 0x0005
Private Const LL_DLGBOXMODE_ALT5 = 6                                ' 0x0006
Private Const LL_DLGBOXMODE_ALT6 = 7                                ' 0x0007
Private Const LL_DLGBOXMODE_ALT7 = 8                                ' 0x0008
Private Const LL_DLGBOXMODE_ALT8 = 9                                ' 0x0009
Private Const LL_DLGBOXMODE_ALT9 = 10                               ' 0x000A
Private Const LL_DLGBOXMODE_ALT10 = 11                              ' 0x000B
Private Const LL_DLGBOXMODE_TOOLTIPS98 = 2048                       ' 0x0800
Private Const LL_CTL_ADDTOSYSMENU = 4                               ' 0x00000004
Private Const LL_CTL_ALSOCHILDREN = 16                              ' 0x00000010
Private Const LL_CTL_CONVERTCONTROLS = 65536                        ' 0x00010000
Private Const LL_GROUP_ALWAYSFOOTER = 1073741824                    ' 0x40000000
Private Const LL_PRINTERCONFIG_SAVE = 1                             ' 1
Private Const LL_PRINTERCONFIG_RESTORE = 2                          ' 2
Private Const LL_PRJTYPE_OPTION_FORCEDEFAULTSETTINGS = 32768        ' 0x8000
Private Const LL_PRJTYPE_OPTION_CREATEPARTSFROMPROJECT = 16384      ' 0x4000
Private Const LL_RTFTEXTMODE_RTF = 0                                ' 0x0000
Private Const LL_RTFTEXTMODE_PLAIN = 1                              ' 0x0001
Private Const LL_RTFTEXTMODE_EVALUATED = 0                          ' 0x0000
Private Const LL_RTFTEXTMODE_RAW = 2                                ' 0x0002
Private Const LL_RTFTEXTFLAG_ALL = 0                                ' 0x0000
Private Const LL_RTFTEXTFLAG_SELECTION = 4                          ' 0x0004
Private Const LL_ERR_BAD_JOBHANDLE = -1                             ' -1
Private Const LL_ERR_TASK_ACTIVE = -2                               ' -2
Private Const LL_ERR_BAD_OBJECTTYPE = -3                            ' -3
Private Const LL_ERR_BAD_PROJECTTYPE = -3                           ' -3
Private Const LL_ERR_PRINTING_JOB = -4                              ' -4
Private Const LL_ERR_NO_BOX = -5                                    ' -5
Private Const LL_ERR_ALREADY_PRINTING = -6                          ' -6
Private Const LL_ERR_NOT_YET_PRINTING = -7                          ' -7
Private Const LL_ERR_NO_PROJECT = -10                               ' -10
Private Const LL_ERR_NO_PRINTER = -11                               ' -11
Private Const LL_ERR_PRINTING = -12                                 ' -12
Private Const LL_ERR_EXPORTING = -13                                ' -13
Private Const LL_ERR_NEEDS_VB = -14                                 ' -14
Private Const LL_ERR_BAD_PRINTER = -15                              ' -15
Private Const LL_ERR_NO_PREVIEWMODE = -16                           ' -16
Private Const LL_ERR_NO_PREVIEWFILES = -17                          ' -17
Private Const LL_ERR_PARAMETER = -18                                ' -18
Private Const LL_ERR_BAD_EXPRESSION = -19                           ' -19
Private Const LL_ERR_BAD_EXPRMODE = -20                             ' -20
Private Const LL_ERR_NO_TABLE = -21                                 ' -21
Private Const LL_ERR_CFGNOTFOUND = -22                              ' -22
Private Const LL_ERR_EXPRESSION = -23                               ' -23
Private Const LL_ERR_CFGBADFILE = -24                               ' -24
Private Const LL_ERR_BADOBJNAME = -25                               ' -25
Private Const LL_ERR_NOOBJECT = -26                                 ' -26
Private Const LL_ERR_UNKNOWNOBJECT = -27                            ' -27
Private Const LL_ERR_NO_TABLEOBJECT = -28                           ' -28
Private Const LL_ERR_NO_OBJECT = -29                                ' -29
Private Const LL_ERR_NO_TEXTOBJECT = -30                            ' -30
Private Const LL_ERR_UNKNOWN = -31                                  ' -31
Private Const LL_ERR_BAD_MODE = -32                                 ' -32
Private Const LL_ERR_CFGBADMODE = -33                               ' -33
Private Const LL_ERR_ONLYWITHONETABLE = -34                         ' -34
Private Const LL_ERR_UNKNOWNVARIABLE = -35                          ' -35
Private Const LL_ERR_UNKNOWNFIELD = -36                             ' -36
Private Const LL_ERR_UNKNOWNSORTORDER = -37                         ' -37
Private Const LL_ERR_NOPRINTERCFG = -38                             ' -38
Private Const LL_ERR_SAVEPRINTERCFG = -39                           ' -39
Private Const LL_ERR_RESERVED = -40                                 ' -40
Private Const LL_ERR_NOVALIDPAGES = -41                             ' -41
Private Const LL_ERR_NOTINHOSTPRINTERMODE = -42                     ' -42
Private Const LL_ERR_NOTFINISHED = -43                              ' -43
Private Const LL_ERR_BUFFERTOOSMALL = -44                           ' -44
Private Const LL_ERR_BADCODEPAGE = -45                              ' -45
Private Const LL_ERR_CANNOTCREATETEMPFILE = -46                     ' -46
Private Const LL_ERR_NODESTINATION = -47                            ' -47
Private Const LL_ERR_NOCHART = -48                                  ' -48
Private Const LL_ERR_TOO_MANY_CONCURRENT_PRINTJOBS = -49            ' -49
Private Const LL_ERR_BAD_WEBSERVER_LICENSE = -50                    ' -50
Private Const LL_ERR_NO_WEBSERVER_LICENSE = -51                     ' -51
Private Const LL_ERR_INVALIDDATE = -52                              ' -52
Private Const LL_ERR_DRAWINGNOTFOUND = -53                          ' -53
Private Const LL_ERR_NOUSERINTERACTION = -54                        ' -54
Private Const LL_ERR_BADDATABASESTRUCTURE = -55                     ' -55
Private Const LL_ERR_UNKNOWNPROPERTY = -56                          ' -56
Private Const LL_ERR_INVALIDOPERATION = -57                         ' -57
Private Const LL_ERR_PROPERTY_ALREADY_DEFINED = -58                 ' -58
Private Const LL_ERR_CFGFOUND = -59                                 ' -59
Private Const LL_ERR_SAVECFG = -60                                  ' -60
Private Const LL_ERR_WRONGTHREAD = -61                              ' -61
Private Const LL_ERR_NO_SUCH_INFORMATION = -62                      ' -62
Private Const LL_ERR_SINK_ALREADY_PRESENT = -63                     ' -63
Private Const LL_ERR_SINK_NOT_PRESENT = -64                         ' -64
Private Const LL_ERR_ACCESS_DENIED = -65                            ' -65
Private Const LL_ERR_USER_ABORTED = -99                             ' -99
Private Const LL_ERR_BAD_DLLS = -100                                ' -100
Private Const LL_ERR_NO_LANG_DLL = -101                             ' -101
Private Const LL_ERR_NO_MEMORY = -102                               ' -102
Private Const LL_ERR_EXCEPTION = -104                               ' -104
Private Const LL_ERR_LICENSEVIOLATION = -105                        ' -105
Private Const LL_ERR_NOT_SUPPORTED_IN_THIS_OS = -106                ' -106
Private Const LL_ERR_NO_MORE_DATA = -107                            ' -107
Private Const LL_WRN_REPORTPARAMETERS_COLLECTION_FINISHED = -994    ' -994
Private Const LL_WRN_ISNULL = -995                                  ' -995
Private Const LL_WRN_TABLECHANGE = -996                             ' -996
Private Const LL_WRN_PRINTFINISHED = -997                           ' -997
Private Const LL_WRN_REPEAT_DATA = -998                             ' -998
Private Const LL_CHAR_TEXTQUOTE = 1                                 ' 1
Private Const LL_CHAR_PHANTOMSPACE = 8203                           ' 0x200b
Private Const LL_CHAR_LOCK = 8288                                   ' 0x2060
Private Const LL_CHAR_NEWLINE = 182                                 ' 182
Private Const LL_CHAR_EXPRSEP = 164                                 ' 164
Private Const LL_CHAR_TAB = 247                                     ' 247
Private Const LL_CHAR_EAN128NUL = 255                               ' 255
Private Const LL_CHAR_EAN128FNC1 = 254                              ' 254
Private Const LL_CHAR_EAN128FNC2 = 253                              ' 253
Private Const LL_CHAR_EAN128FNC3 = 252                              ' 252
Private Const LL_CHAR_EAN128FNC4 = 251                              ' 251
Private Const LL_CHAR_CODE93NUL = 255                               ' 255
Private Const LL_CHAR_CODE93EXDOLLAR = 254                          ' 254
Private Const LL_CHAR_CODE93EXPERC = 253                            ' 253
Private Const LL_CHAR_CODE93EXSLASH = 252                           ' 252
Private Const LL_CHAR_CODE93EXPLUS = 251                            ' 251
Private Const LL_CHAR_CODE39NUL = 255                               ' 255
Private Const LL_DLGEXPR_VAREXTBTN_ENABLE = 1                       ' 0x00000001
Private Const LL_DLGEXPR_VAREXTBTN_DOMODAL = 2                      ' 0x00000002
Private Const LL_LLX_EXTENSIONTYPE_EXPORT = 1                       ' 1
Private Const LL_LLX_EXTENSIONTYPE_BARCODE = 2                      ' 2
Private Const LL_LLX_EXTENSIONTYPE_OBJECT = 3                       ' 3
Private Const LL_LLX_EXTENSIONTYPE_WIZARD = 4                       ' 4
Private Const LL_LLX_EXTENSIONTYPEFLAG_FORCE_PUBLIC = 65536         ' 0x00010000
Private Const LL_LLX_EXTENSIONTYPEFLAG_FORCE_PRIVATE = 131072       ' 0x00020000
Private Const LL_LLX_EXTENSIONTYPE_TYPEMASK = 15                    ' 0x0000000f
Private Const LL_DECLARECHARTROW_FOR_OBJECTS = 1                    ' 0x00000001
Private Const LL_DECLARECHARTROW_FOR_TABLECOLUMNS = 2               ' 0x00000002
Private Const LL_DECLARECHARTROW_FOR_TABLECOLUMNS_FOOTERS = 4          ' 0x00000004
Private Const LL_GETCHARTOBJECTCOUNT_CHARTOBJECTS = 1               ' 1
Private Const LL_GETCHARTOBJECTCOUNT_CHARTOBJECTS_BEFORE_TABLE = 2          ' 2
Private Const LL_GETCHARTOBJECTCOUNT_CHARTCOLUMNS = 3               ' 3
Private Const LL_GETCHARTOBJECTCOUNT_CHARTCOLUMNS_FOOTERS = 4          ' 4
Private Const LL_VARIANTFLAG_NEUTRAL = 0                            ' 0x00000000
Private Const LL_VARIANTFLAG_USE_JULIAN_DATE = 1                    ' 0x00000001
Private Const LL_GRIPT_DIM_SCM = 1                                  ' 1
Private Const LL_GRIPT_DIM_PERC = 2                                 ' 2
Private Const LL_PARAMETERFLAG_PUBLIC = 0                           ' 0x00000000
Private Const LL_PARAMETERFLAG_SAVEDEFAULT = 65536                  ' 0x00010000
Private Const LL_PARAMETERFLAG_PRIVATE = 1073741824                 ' 0x40000000
Private Const LL_PARAMETERFLAG_FORMULA = 0                          ' 0x00000000
Private Const LL_PARAMETERFLAG_VALUE = 536870912                    ' 0x20000000
Private Const LL_PARAMETERFLAG_GLOBAL = 0                           ' 0x00000000
Private Const LL_PARAMETERFLAG_LOCAL = 268435456                    ' 0x10000000
Private Const LL_PARAMETERFLAG_MASK = -65536                        ' 0xffff0000
Private Const LL_PARAMETERTYPE_USER = 0                             ' 0
Private Const LL_PARAMETERTYPE_FAX = 1                              ' 1
Private Const LL_PARAMETERTYPE_MAIL = 2                             ' 2
Private Const LL_PARAMETERTYPE_LLINTERNAL = 4                       ' 4
Private Const LL_PARAMETERTYPE_MASK = 15                            ' 0x0000000f
Private Const LL_PRJOPEN_AM_READWRITE = 1073741824                  ' 0x40000000
Private Const LL_PRJOPEN_AM_READONLY = 0                            ' 0x00000000
Private Const LL_PRJOPEN_AM_MASK = 1073741824                       ' 0x40000000
Private Const LL_PRJOPEN_CD_OPEN_EXISTING = 0                       ' 0x00000000
Private Const LL_PRJOPEN_CD_CREATE_ALWAYS = 268435456               ' 0x10000000
Private Const LL_PRJOPEN_CD_CREATE_NEW = 536870912                  ' 0x20000000
Private Const LL_PRJOPEN_CD_OPEN_ALWAYS = 805306368                 ' 0x30000000
Private Const LL_PRJOPEN_CD_MASK = 805306368                        ' 0x30000000
Private Const LL_PRJOPEN_EM_IGNORE_FORMULAERRORS = 134217728        ' 0x08000000
Private Const LL_PRJOPEN_EM_MASK = 134217728                        ' 0x08000000
Private Const LL_PRJOPEN_FLG_NOINITPRINTER = 67108864               ' 0x04000000
Private Const LL_PRJOPEN_FLG_NOOBJECTLOAD = 33554432                ' 0x02000000
Private Const LL_PRJOPEN_FLG_RESERVED = 16777216                    ' 0x01000000
Private Const LL_ASSOCIATEPREVIEWCONTROLFLAG_DELETE_ON_CLOSE = 1          ' 0x0001
Private Const LL_ASSOCIATEPREVIEWCONTROLFLAG_HANDLE_IS_ATTACHINFO = 2          ' 0x0002
Private Const LL_ASSOCIATEPREVIEWCONTROLFLAG_PRV_REPLACE = 0          ' 0x0000
Private Const LL_ASSOCIATEPREVIEWCONTROLFLAG_PRV_ADD_TO_CONTROL_STACK = 4          ' 0x0004
Private Const LL_ASSOCIATEPREVIEWCONTROLFLAG_PRV_ADD_TO_CONTROL_IN_TAB = 8          ' 0x0008
Private Const LL_ASSOCIATEPREVIEWCONTROLMASK_ATTACHLOCATION = 60         ' 0x003c
Private Const LL_DESFILEOPEN_OPEN_EXISTING = 0                      ' 0x00000000
Private Const LL_DESFILEOPEN_CREATE_ALWAYS = 268435456              ' 0x10000000
Private Const LL_DESFILEOPEN_CREATE_NEW = 536870912                 ' 0x20000000
Private Const LL_DESFILEOPEN_OPEN_ALWAYS = 805306368                ' 0x30000000
Private Const LL_DESFILEOPEN_OPEN_IMPORT = 1073741824               ' 0x40000000
Private Const LL_DESFILEOPEN_OPENMASK = 1879048192                  ' 0x70000000
Private Const LL_DESFILEOPENFLAG_SUPPRESS_SAVEDIALOG = 1            ' 0x00000001
Private Const LL_DESFILEOPENFLAG_SUPPRESS_SAVE = 2                  ' 0x00000002
Private Const LL_DESFILESAVE_DEFAULT = 0                            ' 0x00000000
Private Const LLDESADDACTIONFLAG_ADD_TO_TOOLBAR = 536870912         ' 0x20000000
Private Const LLDESADDACTION_MENUITEM_APPEND = 0                    ' 0x00000000
Private Const LLDESADDACTION_MENUITEM_INSERT = 268435456            ' 0x10000000
Private Const LLDESADDACTION_MENUITEM_POSITIONMASK = 268435456      ' 0x10000000
Private Const LLDESADDACTION_ACCEL_CONTROL = 65536                  ' 0x00010000
Private Const LLDESADDACTION_ACCEL_SHIFT = 131072                   ' 0x00020000
Private Const LLDESADDACTION_ACCEL_ALT = 262144                     ' 0x00040000
Private Const LLDESADDACTION_ACCEL_VIRTKEY = 524288                 ' 0x00080000
Private Const LLDESADDACTION_ACCEL_KEYMODIFIERMASK = 983040         ' 0x000f0000
Private Const LLDESADDACTION_ACCEL_KEYCODEMASK = 65535              ' 0x0000ffff
Private Const LL_DESIGNEROPTSTR_PROJECTFILENAME = 1                 ' 1
Private Const LL_DESIGNEROPTSTR_WORKSPACETITLE = 2                  ' 2
Private Const LL_DESIGNEROPTSTR_PROJECTDESCRIPTION = 3              ' 3
Private Const LL_USEDIDENTIFIERSFLAG_VARIABLES = 1                  ' 0x0001
Private Const LL_USEDIDENTIFIERSFLAG_FIELDS = 2                     ' 0x0002
Private Const LL_USEDIDENTIFIERSFLAG_CHARTFIELDS = 4                ' 0x0004
Private Const LL_USEDIDENTIFIERSFLAG_TABLES = 8                     ' 0x0008
Private Const LL_USEDIDENTIFIERSFLAG_RELATIONS = 16                 ' 0x0010
Private Const LL_USEDIDENTIFIERSFLAG_FILES = 32                     ' 0x0020
Private Const LL_TEMPFILENAME_DEFAULT = 0                           ' 0x0000
Private Const LL_TEMPFILENAME_ENSURELONGPATH = 1                    ' 0x0001
Private Const LL_DICTIONARY_TYPE_STATIC = 1                         ' 1
Private Const LL_DICTIONARY_TYPE_IDENTIFIER = 2                     ' 2
Private Const LL_DICTIONARY_TYPE_TABLE = 3                          ' 3
Private Const LL_DICTIONARY_TYPE_RELATION = 4                       ' 4
Private Const LL_DICTIONARY_TYPE_SORTORDER = 5                      ' 5
Private Const LL_DICTIONARY_TYPE_ALL = 0                            ' 0
Private Const LL_UILANGUAGETYPE_NORMAL = 1                          ' 0x00000001
Private Const LL_UILANGUAGETYPE_TRIAL = 2                           ' 0x00000002
Private Const LL_ADDTABLEOPT_SUPPORTSSTACKEDSORTORDERS = 1          ' 0x00000001
Private Const LL_ADDTABLEOPT_SUPPORTSADVANCEDFILTERING = 2          ' 0x00000002
Private Const LL_INPLACEDESIGNERINTERACTION_QUERY_CANCLOSE = 1          ' 1
Private Const LL_EXPRXLATRESULT_OPTIMAL = 0                         ' 0x00000000
Private Const LL_EXPRXLATRESULT_PARTIAL = 1                         ' 0x00000001
Private Const LL_EXPRXLATRESULT_FAIL = 2                            ' 0x00000002
Private Const LL_EXPRXLATRESULT_MASK = 7                            ' 0x00000007
Private Const LLJOBOPENCOPYEXFLAG_NO_COPY_FIELDLIST = 1             ' 0x0001
Private Const LLJOBOPENCOPYEXFLAG_NO_COPY_DBSTRUCTS = 2             ' 0x0002
Private Const LLJOBOPENCOPYEXFLAG_NO_COPY_XLATTABLES = 4            ' 0x0004

Private Const CMBTLANG_DEFAULT = -1
Private Const CMBTLANG_GERMAN = 0
Private Const CMBTLANG_ENGLISH = 1
Private Const CMBTLANG_ARABIC = 2
Private Const CMBTLANG_AFRIKAANS = 3
Private Const CMBTLANG_ALBANIAN = 4
Private Const CMBTLANG_BASQUE = 5
Private Const CMBTLANG_BULGARIAN = 6
Private Const CMBTLANG_BYELORUSSIA = 7
Private Const CMBTLANG_CATALAN = 8
Private Const CMBTLANG_CHINESE = 9
Private Const CMBTLANG_CROATIAN = 10
Private Const CMBTLANG_CZECH = 11
Private Const CMBTLANG_DANISH = 12
Private Const CMBTLANG_DUTCH = 13
Private Const CMBTLANG_ESTONIAN = 14
Private Const CMBTLANG_FAEROESE = 15
Private Const CMBTLANG_FARSI = 16
Private Const CMBTLANG_FINNISH = 17
Private Const CMBTLANG_FRENCH = 18
Private Const CMBTLANG_GREEK = 19
Private Const CMBTLANG_HEBREW = 20
Private Const CMBTLANG_HUNGARIAN = 21
Private Const CMBTLANG_ICELANDIC = 22
Private Const CMBTLANG_INDONESIAN = 23
Private Const CMBTLANG_ITALIAN = 24
Private Const CMBTLANG_JAPANESE = 25
Private Const CMBTLANG_KOREAN = 26
Private Const CMBTLANG_LATVIAN = 27
Private Const CMBTLANG_LITHUANIAN = 28
Private Const CMBTLANG_NORWEGIAN = 29
Private Const CMBTLANG_POLISH = 30
Private Const CMBTLANG_PORTUGUESE = 31
Private Const CMBTLANG_ROMANIAN = 32
Private Const CMBTLANG_RUSSIAN = 33
Private Const CMBTLANG_SLOVAK = 34
Private Const CMBTLANG_SLOVENIAN = 35
Private Const CMBTLANG_SERBIAN = 36
Private Const CMBTLANG_SPANISH = 37
Private Const CMBTLANG_SWEDISH = 38
Private Const CMBTLANG_THAI = 39
Private Const CMBTLANG_TURKISH = 40
Private Const CMBTLANG_UKRAINIAN = 41
Private Const CMBTLANG_SERBIAN_LATIN = 42
Private Const CMBTLANG_CHINESE_TRADITIONAL = 48

Private Type DEVMODEA
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Long ' Basic Win32 definition had an error here!
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
 
Private Type RECT ' comment this out if compiler complains
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 
Private Type scLlExtFct
    nSize As Long
    sContents As String
    bEvaluate As Long
    sNewValue As String
    bError As Long
    sError As String
End Type
 
Private Type scLlObject
    nSize As Long
    sName As String
    nType As Long
    bPreDraw As Long
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
End Type
 
Private Type scLlPage
    nSize As Long
    bDesignerPreview As Long
    bPreDraw As Long
    hRefDC As Long
    hPaintDC As Long
End Type
 
Private Type scLlProject
    nSize As Long
    bDesignerPreview As Long
    bPreDraw As Long
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
End Type
 
Private Type scLlDrawUserObj
    nSize As Long
    sName As String
    sContents As String
    lPara As Long
    lpPtr As Long
    hPara As Long
    bIsotropic As Long
    sParameters As String
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
    nPaintMode As Long
End Type
 
Private Type scLlEditUserObj
    nSize As Long
    sName As String
    lPara As Long
    lpPtr As Long
    hPara As Long
    bIsotropic As Long
    hwnd As Long
    sParameters As String
    nParaBufSize As Long
End Type

Private Type scLlTableLine
    nSize As Long
    nType As Long
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
    nPageLine As Long
    nLine As Long
    bZebra As Long
    rcSpacing As RECT
End Type
 
Private Type scllTableField
    nSize As Long
    nType As Long
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
    nIndex As Long
    rcSpacing  As RECT
End Type
 
Private Type scLlColumn
    nSize As Long
    xPos As Long
    nWidth As Long
    fntColumn As Long
End Type
 
Private Type scLlGroup
    nSize As Long
    hRefDC As Long
    hPaintDC As Long
    rcPaint As RECT
    fntHeader As Long
    fntBody As Long
    fntFooter As Long
    nHeight As Long
    bPaint As Long
    bFrameLeft As Long
    bFrameRight As Long
    lParam As Long
    sParam As String
    nColumns As Long
    lpColumns As Long     ' sigh.... VB users: forget it.
End Type
 
Private Type scLlDlgExprVarExt
    nSize As Long
    nFunction As Long
    hWndDialog As Long
    sPage As String
    sValue As String
    bFields As Long
    nMask As Long
End Type
 
Private Type scLlToolbar
    nSize As Long
    hwnd As Long
    nType As Long
    hMenu As Long
    nParam As Long
End Type
 
Private Type scLlPrintJobInfo
    nSize As Long
    hLlJob As Long
    szDevice As String * 80
    dwJobID As Long
    dwState As Long
End Type

'-------------------- Sotrage --------------------------

Private Const LL_STGSYS_VERSION_LL18 = 2                            ' 2
Private Const LL_ERR_STG_NOSTORAGE = -1000                          ' -1000
Private Const LL_ERR_STG_BADVERSION = -1001                         ' -1001
Private Const LL_ERR_STG_READ = -1002                               ' -1002
Private Const LL_ERR_STG_WRITE = -1003                              ' -1003
Private Const LL_ERR_STG_UNKNOWNSYSTEM = -1004                      ' -1004
Private Const LL_ERR_STG_BADHANDLE = -1005                          ' -1005
Private Const LL_ERR_STG_ENDOFLIST = -1006                          ' -1006
Private Const LL_ERR_STG_BADJOB = -1007                             ' -1007
Private Const LL_ERR_STG_ACCESSDENIED = -1008                       ' -1008
Private Const LL_ERR_STG_BADSTORAGE = -1009                         ' -1009
Private Const LL_ERR_STG_CANNOTGETMETAFILE = -1010                  ' -1010
Private Const LL_ERR_STG_OUTOFMEMORY = -1011                        ' -1011
Private Const LL_ERR_STG_SEND_FAILED = -1012                        ' -1012
Private Const LL_ERR_STG_DOWNLOAD_PENDING = -1013                   ' -1013
Private Const LL_ERR_STG_DOWNLOAD_FAILED = -1014                    ' -1014
Private Const LL_ERR_STG_WRITE_FAILED = -1015                       ' -1015
Private Const LL_ERR_STG_UNEXPECTED = -1016                         ' -1016
Private Const LL_ERR_STG_CANNOTCREATEFILE = -1017                   ' -1017
Private Const LL_ERR_STG_UNKNOWN_CONVERTER = -1018                  ' -1018
Private Const LL_ERR_STG_INET_ERROR = -1019                         ' -1019
Private Const LL_ERR_STG_NOTFOUND = -1020                           ' -1020
Private Const LL_WRN_STG_UNFAXED_PAGES = -1100                      ' -1100
Private Const LS_OPTION_HAS16BITPAGES = 200                         ' 200
Private Const LS_OPTION_BOXTYPE = 201                               ' 201
Private Const LS_OPTION_UNITS = 203                                 ' 203
Private Const LS_OPTION_PRINTERCOUNT = 204                          ' 204
Private Const LS_OPTION_ISSTORAGE = 205                             ' 205
Private Const LS_OPTION_EMFRESOLUTION = 206                         ' 206
Private Const LS_OPTION_JOB = 207                                   ' 207
Private Const LS_OPTION_TOTALPAGES = 208                            ' 208
Private Const LS_OPTION_PAGESWITHFAXNUMBER = 209                    ' 209
Private Const LS_OPTION_HASINPUTOBJECTS = 210                       ' 210
Private Const LS_OPTION_HASFORCEDINPUTOBJECTS = 211                 ' 211
Private Const LS_OPTION_INPUTOBJECTSFINISHED = 212                  ' 212
Private Const LS_OPTION_HASHYPERLINKS = 213                         ' 213
Private Const LS_OPTION_USED_PRINTERCOUNT = 214                     ' 214
Private Const LS_OPTION_PAGENUMBER = 0                              ' 0
Private Const LS_OPTION_COPIES = 1                                  ' 1
Private Const LS_OPTION_PRN_ORIENTATION = 2                         ' 2
Private Const LS_OPTION_PHYSPAGE = 3                                ' 3
Private Const LS_OPTION_PRN_PIXELSOFFSET_X = 4                      ' 4
Private Const LS_OPTION_PRN_PIXELSOFFSET_Y = 5                      ' 5
Private Const LS_OPTION_PRN_PIXELS_X = 6                            ' 6
Private Const LS_OPTION_PRN_PIXELS_Y = 7                            ' 7
Private Const LS_OPTION_PRN_PIXELSPHYSICAL_X = 8                    ' 8
Private Const LS_OPTION_PRN_PIXELSPHYSICAL_Y = 9                    ' 9
Private Const LS_OPTION_PRN_PIXELSPERINCH_X = 10                    ' 10
Private Const LS_OPTION_PRN_PIXELSPERINCH_Y = 11                    ' 11
Private Const LS_OPTION_PRN_INDEX = 12                              ' 12
Private Const LS_OPTION_PRN_PAPERTYPE = 13                          ' 13
Private Const LS_OPTION_PRN_PAPERSIZE_X = 14                        ' 14
Private Const LS_OPTION_PRN_PAPERSIZE_Y = 15                        ' 15
Private Const LS_OPTION_PRN_FORCE_PAPERSIZE = 16                    ' 16
Private Const LS_OPTION_STARTNEWSHEET = 17                          ' 17
Private Const LS_OPTION_ISSUEINDEX = 18                             ' 18
Private Const LS_OPTION_STARTNEWJOB = 19                            ' 19
Private Const LS_OPTION_PROJECTNAME = 100                           ' 100
Private Const LS_OPTION_JOBNAME = 101                               ' 101
Private Const LS_OPTION_PRTNAME = 102                               ' 102
Private Const LS_OPTION_PRTDEVICE = 103                             ' 103
Private Const LS_OPTION_PRTPORT = 104                               ' 104
Private Const LS_OPTION_USER = 105                                  ' 105
Private Const LS_OPTION_CREATION = 106                              ' 106
Private Const LS_OPTION_CREATIONAPP = 107                           ' 107
Private Const LS_OPTION_CREATIONDLL = 108                           ' 108
Private Const LS_OPTION_CREATIONUSER = 109                          ' 109
Private Const LS_OPTION_FAXPARA_QUEUE = 110                         ' 110
Private Const LS_OPTION_FAXPARA_RECIPNAME = 111                     ' 111
Private Const LS_OPTION_FAXPARA_RECIPNUMBER = 112                   ' 112
Private Const LS_OPTION_FAXPARA_SENDERNAME = 113                    ' 113
Private Const LS_OPTION_FAXPARA_SENDERCOMPANY = 114                 ' 114
Private Const LS_OPTION_FAXPARA_SENDERDEPT = 115                    ' 115
Private Const LS_OPTION_FAXPARA_SENDERBILLINGCODE = 116             ' 116
Private Const LS_OPTION_FAX_AVAILABLEQUEUES = 118                   ' 118
Private Const LS_OPTION_PRINTERALIASLIST = 119                      ' 119
Private Const LS_OPTION_PRTDEVMODE = 120                            ' 120
Private Const LS_OPTION_USED_PRTDEVICE = 121                        ' 121
Private Const LS_OPTION_USED_PRTDEVMODE = 122                       ' 122
Private Const LS_OPTION_REGIONNAME = 123                            ' 123
Private Const LS_PRINTFLAG_FIT = 1                                  ' 0x00000001
Private Const LS_PRINTFLAG_STACKEDCOPIES = 2                        ' 0x00000002
Private Const LS_PRINTFLAG_TRYPRINTERCOPIES = 4                     ' 0x00000004
Private Const LS_PRINTFLAG_SHOWDIALOG = 8                           ' 0x00000008
Private Const LS_PRINTFLAG_METER = 16                               ' 0x00000010
Private Const LS_PRINTFLAG_ABORTABLEMETER = 32                      ' 0x00000020
Private Const LS_PRINTFLAG_METERMASK = 112                          ' 0x00000070
Private Const LS_PRINTFLAG_USEDEFPRINTERIFNULL = 128                ' 0x00000080
Private Const LS_PRINTFLAG_FAX = 256                                ' 0x00000100
Private Const LS_PRINTFLAG_OVERRIDEPROJECTCOPYCOUNT = 512           ' 0x00000200
Private Const LS_VIEWERCONTROL_QUERY_CHARWIDTH = 1                  ' 1
Private Const LS_VIEWERCONTROL_CLEAR = 1025                         ' WM_USER+1
Private Const LS_VIEWERCONTROL_SET_HANDLE_EX = 1026                 ' WM_USER+2
Private Const LS_VIEWERCONTROL_SET_HANDLE = 1027                    ' WM_USER+3
Private Const LS_VIEWERCONTROLSETHANDLEFLAG_ADD = 256               ' 0x0100
Private Const LS_VIEWERCONTROLSETHANDLEFLAG_DELETE_ON_CLOSE = 512        ' 0x0200
Private Const LS_VIEWERCONTROL_GET_HANDLE = 1028                    ' WM_USER+4
Private Const LS_VIEWERCONTROL_SET_FILENAME = 1029                  ' WM_USER+5
Private Const LS_STGFILEOPEN_READONLY = 0                           ' 0x00000000
Private Const LS_STGFILEOPEN_READWRITE = 1                          ' 0x00000001
Private Const LS_STGFILEOPEN_FORCE_NO_READWRITE = 2                 ' 0x00000002
Private Const LS_STGFILEOPEN_DELETE_ON_CLOSE = 4                    ' 0x00000004
Private Const LS_STGFILEOPENFLAG_ADD = 256                          ' 0x00000100
Private Const LS_VIEWERCONTROL_SET_OPTION = 1030                    ' WM_USER+6
Private Const LS_OPTION_MESSAGE = 0                                 ' 0
Private Const LS_OPTION_PRINTERASSIGNMENT = 1                       ' 1
Private Const LS_PRNASSIGNMENT_USEDEFAULT = 0                       ' 0x00000000
Private Const LS_PRNASSIGNMENT_ASKPRINTERIFNEEDED = 1               ' 0x00000001
Private Const LS_PRNASSIGNMENT_ASKPRINTERALWAYS = 2                 ' 0x00000002
Private Const LS_PRNASSIGNMENT_ALWAYSUSEDEFAULT = 3                 ' 0x00000003
Private Const LS_OPTION_TOOLBAR = 2                                 ' 2
Private Const LS_OPTION_SKETCHBAR = 3                               ' 3
Private Const LS_OPTION_SKETCHBARWIDTH = 4                          ' 4
Private Const LS_OPTION_TOOLBARSTYLE = 5                            ' 5
Private Const LS_OPTION_TOOLBARSTYLE_STANDARD = 0                   ' 0
Private Const LS_OPTION_TOOLBARSTYLE_OFFICEXP = 1                   ' 1
Private Const LS_OPTION_TOOLBARSTYLE_OFFICE2003 = 2                 ' 2
Private Const LS_OPTION_TOOLBARSTYLEMASK = 15                       ' 0x0f
Private Const LS_OPTION_TOOLBARSTYLEFLAG_GRADIENT = 128             ' 0x80
Private Const LS_OPTION_CODEPAGE = 7                                ' 7
Private Const LS_OPTION_SAVEASFILEPATH = 8                          ' 8
Private Const LS_OPTION_USERDATA = 9                                ' 9
Private Const LS_OPTION_BGCOLOR = 10                                ' 10
Private Const LS_OPTION_ASYNC_DOWNLOAD = 11                         ' 11
Private Const LS_OPTION_LANGUAGE = 12                               ' 12
Private Const LS_OPTION_ASSUME_TEMPFILE = 13                        ' 13
Private Const LS_OPTION_IOLECLIENTSITE = 14                         ' 14
Private Const LS_OPTION_TOOLTIPS = 15                               ' 15
Private Const LS_OPTION_AUTOSAVE = 16                               ' 16
Private Const LS_OPTION_CHANGEDFLAG = 17                            ' 17
Private Const LS_OPTION_SHOW_UNPRINTABLE_AREA = 18                  ' 18
Private Const LS_OPTION_NOUIRESET = 19                              ' 19
Private Const LS_OPTION_NAVIGATIONBAR = 20                          ' 20
Private Const LS_OPTION_NAVIGATIONBARWIDTH = 21                     ' 21
Private Const LS_OPTION_IN_PREVIEWPANE = 22                         ' 22
Private Const LS_OPTION_IN_LLVIEWER = 23                            ' 23
Private Const LS_OPTION_TABBARSTYLE = 24                            ' 24
Private Const LS_OPTION_TABBARSTYLE_STANDARD = 0                    ' 0
Private Const LS_OPTION_TABBARSTYLE_OFFICEXP = 1                    ' 1
Private Const LS_OPTION_TABBARSTYLE_OFFICE2003 = 2                  ' 2
Private Const LS_OPTION_DESIGNERPREVIEW = 25                        ' 25
Private Const LS_OPTION_MOUSEMODE = 26                              ' 26
Private Const LS_OPTION_MOUSEMODE_MOVE = 1                          ' 1
Private Const LS_OPTION_MOUSEMODE_ZOOM = 2                          ' 2
Private Const LS_OPTION_ALLOW_RBUTTONUSAGE = 27                     ' 27
Private Const LS_OPTION_TOOLBGCOLOR = 28                            ' 28
Private Const LS_VIEWERCONTROL_GET_OPTION = 1031                    ' WM_USER+7
Private Const LS_VIEWERCONTROL_QUERY_ENDSESSION = 1032              ' WM_USER+8
Private Const LS_VIEWERCONTROL_GET_ZOOM = 1033                      ' WM_USER+9
Private Const LS_VIEWERCONTROL_SET_ZOOM = 1034                      ' WM_USER+10
Private Const LS_VIEWERCONTROL_GET_ZOOMED = 1035                    ' WM_USER+11
Private Const LS_VIEWERCONTROL_POP_ZOOM = 1036                      ' WM_USER+12
Private Const LS_VIEWERCONTROL_RESET_ZOOM = 1037                    ' WM_USER+13
Private Const LS_VIEWERCONTROL_SET_ZOOM_TWICE = 1038                ' WM_USER+14
Private Const LS_VIEWERCONTROL_SET_PAGE = 1044                      ' WM_USER+20
Private Const LS_VIEWERCONTROL_GET_PAGE = 1045                      ' WM_USER+21
Private Const LS_VIEWERCONTROL_GET_PAGECOUNT = 1046                 ' WM_USER+22
Private Const LS_VIEWERCONTROL_GET_PAGECOUNT_FAXPAGES = 1047        ' WM_USER+23
Private Const LS_VIEWERCONTROL_GET_JOB = 1048                       ' WM_USER+24
Private Const LS_VIEWERCONTROL_GET_JOBPAGEINDEX = 1049              ' WM_USER+25
Private Const LS_VIEWERCONTROL_GET_METAFILE = 1050                  ' WM_USER+26
Private Const LS_VIEWERCONTROL_GET_ENABLED = 1051                   ' WM_USER+27
Private Const LS_VCITEM_SEARCH_FIRST = 0                            ' 0
Private Const LS_VCITEM_SEARCH_NEXT = 1                             ' 1
Private Const LS_VCITEM_SEARCH_OPTS = 2                             ' 2
Private Const LS_VCITEM_SEARCHACTIONMASK = 4095                     ' 0x00000fff
Private Const LS_VCITEM_SEARCHFLAG_CASEINSENSITIVE = 32768          ' 0x00008000
Private Const LS_VCITEM_SEARCHFLAG_UTF16 = 16384                    ' 0x00004000
Private Const LS_VCITEM_SEARCHFLAGMASK = -4096                      ' 0xfffff000
Private Const LS_VCITEM_SAVE_AS_FILE = 3                            ' 3
Private Const LS_VCITEM_SEND_AS_MAIL = 4                            ' 4
Private Const LS_VCITEM_SEND_AS_FAX = 5                             ' 5
Private Const LS_VCITEM_PRINT_ONE = 6                               ' 6
Private Const LS_VCITEM_PRINT_ALL = 7                               ' 7
Private Const LS_VCITEM_PAGENUMBER = 8                              ' 8
Private Const LS_VCITEM_ZOOM = 9                                    ' 9
Private Const LS_VCITEM_THEATERMODE = 10                            ' 10
Private Const LS_VCITEM_PREVSTG = 11                                ' 11
Private Const LS_VCITEM_NEXTSTG = 12                                ' 12
Private Const LS_VCITEM_SEARCH_DONE = 13                            ' 13
Private Const LS_VCITEM_FIRSTPAGE = 14                              ' 14
Private Const LS_VCITEM_NEXTPAGE = 15                               ' 15
Private Const LS_VCITEM_PREVIOUSPAGE = 16                           ' 16
Private Const LS_VCITEM_LASTPAGE = 17                               ' 17
Private Const LS_VCITEM_MOUSEMODE_MOVE = 18                         ' 18
Private Const LS_VCITEM_MOUSEMODE_ZOOM = 19                         ' 19
Private Const LS_VIEWERCONTROL_GET_SEARCHSTATE = 1052               ' WM_USER+28
Private Const LS_VIEWERCONTROL_SEARCH = 1053                        ' WM_USER+29
Private Const LS_VIEWERCONTROL_SEARCHDLGACTIVE = 1054               ' WM_USER+30
Private Const LS_VIEWERCONTROL_PRINT_CURRENT = 1055                 ' WM_USER+31
Private Const LS_VIEWERCONTROL_PRINT_ALL = 1056                     ' WM_USER+32
Private Const LS_VIEWERCONTROL_PRINT_TO_FAX = 1057                  ' WM_USER+33
Private Const LS_VIEWERCONTROL_UPDATE_TOOLBAR = 1059                ' WM_USER+35
Private Const LS_VIEWERCONTROL_GET_TOOLBAR = 1060                   ' WM_USER+36
Private Const LS_VIEWERCONTROL_SAVE_TO_FILE = 1061                  ' WM_USER+37
Private Const LS_VIEWERCONTROL_SEND_AS_MAIL = 1063                  ' WM_USER+39
Private Const LS_VIEWERCONTROL_SET_OPTIONSTR = 1064                 ' WM_USER+40
Private Const LS_VIEWERCONTROL_GET_OPTIONSTR = 1065                 ' WM_USER+41
Private Const LS_VIEWERCONTROL_GET_OPTIONSTRLEN = 1066              ' WM_USER+42
Private Const LS_VIEWERCONTROL_SET_NTFYCALLBACK = 1067              ' WM_USER+43
Private Const LS_VIEWERCONTROL_GET_NTFYCALLBACK = 1068              ' WM_USER+44
Private Const LS_VIEWERCONTROL_GET_TOOLBARBUTTONSTATE = 1069        ' WM_USER+45
Private Const LS_VIEWERCONTROL_SET_FOCUS = 1070                     ' WM_USER+46
Private Const LS_VCSF_PREVIEW = 1                                   ' 1
Private Const LS_VCSF_SKETCHLIST = 2                                ' 2
Private Const LS_VCSF_TOC = 3                                       ' 3
Private Const LS_VIEWERCONTROL_ADDTOOLBARITEM = 1071                ' WM_USER+47
Private Const LS_VIEWERCONTROL_INTERNAL_CHECKERRORLIST = 1072       ' WM_USER+48
Private Const LS_VIEWERCONTROL_SET_THEATERMODE = 1073               ' WM_USER+49
Private Const LS_VIEWERCONTROL_SET_THEATERFLIPDELAY = 1074          ' WM_USER+50
Private Const LS_VIEWERCONTROL_SET_THEATERFLIPMODE = 1075           ' WM_USER+51
Private Const LS_VCTFM_NONE = 0                                     ' 0
Private Const LS_VCTFM_LINEAR = 1                                   ' 1
Private Const LS_VCTFM_FADE = 2                                     ' 2
Private Const LS_VCTFM_WHEEL = 3                                    ' 3
Private Const LS_VIEWERCONTROL_SELECT_THEATERXFORM = 1076           ' WM_USER+52
Private Const LS_VIEWERCONTROL_NTFY_PRVFSCHANGED = 1077             ' WM_USER+53
Private Const LS_VIEWERCONTROL_SET_PROGRESSINFO = 1078              ' WM_USER+54
Private Const LS_VIEWERCONTROL_GET_FILENAME = 1079                  ' WM_USER+55
Private Const LS_VIEWERCONTROL_QUERY_PRVFS_COMPLETE = 1080          ' WM_USER+56
Private Const LS_VIEWERCONTROL_ONSIZEMOVE = 1081                    ' WM_USER+57
Private Const LS_VIEWERCONTROL_NTFY_SHOW = 1082                     ' WM_USER+58
Private Const LS_VIEWERCONTROL_GET_IDEVICEINFO = 1083               ' WM_USER+59
Private Const LS_VIEWERCONTROL_REMOVEFAILURETOOLTIPS = 1084         ' WM_USER+60
Private Const LS_VIEWERCONTROL_SET_LLNTFYSINK = 1085                ' WM_USER+61
Private Const LS_VIEWERCONTROL_OPEN_PREVSTG = 1086                  ' WM_USER+62
Private Const LS_VIEWERCONTROL_OPEN_NEXTSTG = 1087                  ' WM_USER+63
Private Const LS_VIEWERCONTROL_GET_THEATERSTATE = 1088              ' WM_USER+64
Private Const LS_VIEWERCONTROL_SET_PROGRESSINFO_INTERNAL = 1089       ' WM_USER+65
Private Const LS_VIEWERCONTROL_SUPPORTS_REALDATAEXPORT = 1090       ' WM_USER+66
Private Const LS_VIEWERCONTROL_GET_THIS = 1091                      ' WM_USER+67
Private Const LS_VIEWERCONTROL_SEARCH_LINK = 1092                   ' WM_USER+68
Private Const LS_SEARCH_LINK_HYPERLINK = 0                          ' 0
Private Const LS_VIEWERCONTROL_QUERY_DRILLDOWN_ACTIVE = 1093        ' WM_USER+69
Private Const LS_VIEWERCONTROL_CMND_ABORT_DRILLDOWN_JOBS = 1094       ' WM_USER+70
Private Const LS_VIEWERCONTROL_NTFY_PAGELOADED = 1                  ' 1
Private Const LS_VIEWERCONTROL_NTFY_UPDATETOOLBAR = 2               ' 2
Private Const LS_VIEWERCONTROL_NTFY_PRINT_START = 3                 ' 3
Private Const LS_VIEWERCONTROL_NTFY_PRINT_PAGE = 4                  ' 4
Private Const LS_VIEWERCONTROL_NTFY_PRINT_END = 5                   ' 5
Private Const LS_VIEWERCONTROL_NTFY_TOOLBARUPDATE = 6               ' 6
Private Const LS_VIEWERCONTROL_NTFY_EXITBTNPRESSED = 7              ' 7
Private Const LS_VIEWERCONTROL_NTFY_BTNPRESSED = 8                  ' 8
Private Const LS_VIEWERCONTROL_QUEST_BTNSTATE = 9                   ' 9
Private Const LS_VIEWERCONTROL_NTFY_ERROR = 10                      ' 10
Private Const LS_VIEWERCONTROL_NTFY_MAIL_SENT = 11                  ' 11
Private Const LS_VIEWERCONTROL_NTFY_DOWNLOADFINISHED = 12           ' 12
Private Const LS_VIEWERCONTROL_NTFY_KEYBOARDMESSAGE = 13            ' 13
Private Const LS_VIEWERCONTROL_NTFY_VIEWCHANGED = 14                ' 14
Private Const LS_VIEWERCONTROL_CMND_SAVEDATA = 15                   ' 15
Private Const LS_VIEWERCONTROL_NTFY_DATACHANGED = 16                ' 16
Private Const LS_VIEWERCONTROL_NTFY_PROGRESS = 17                   ' 17
Private Const LS_VIEWERCONTROL_QUEST_SUPPORTCONTINUATION = 18         ' 18
Private Const LS_VIEWERCONTROL_CMND_CONTINUE = 19                   ' 19
Private Const LS_VIEWERCONTROL_NTFY_VIEWERDRILLDOWN = 20            ' 20
Private Const LS_VIEWERCONTROL_QUEST_DRILLDOWNSUPPORT = 21          ' 21
Private Const LS_VIEWERCONTROL_QUEST_HOST_WANTS_KEY = 22            ' 22
Private Const LS_VIEWERCONTROL_INTERNALSYNC = 23                    ' 23
Private Const LS_VIEWERCONTROL_NTFY_RP_REALDATAJOB = 24             ' 24
Private Const LS_VIEWERCONTROL_QUEST_RP_REALDATAJOBSUPPORT = 25         ' 25
Private Const LS_VIEWERCONTROL_QUEST_PROJECTFILENAME = 26           ' 26
Private Const LS_VIEWERCONTROL_QUEST_ORGPROJECTFILENAME = 27         ' 27
Private Const LS_VIEWERCONTROL_NTFY_EXPANDABLEREGIONSJOB = 28         ' 28
Private Const LS_VIEWERCONTROL_NTFY_INTERACTIVESORTINGJOB = 30         ' 30
Private Const LS_VIEWERCONTROL_QUEST_ANYREALDATAJOBSUPPORT = 32         ' 32
Private Const LS_VIEWERCONTROL_NTFY_HYPERLINK = 33                  ' 33
Private Const LS_VIEWERCONTROL_NTFY_ZOOMCHANGED = 34                ' 34
Private Const LS_MAILCONFIG_GLOBAL = 1                              ' 0x0001
Private Const LS_MAILCONFIG_USER = 2                                ' 0x0002
Private Const LS_MAILCONFIG_PROVIDER = 4                            ' 0x0004
Private Const LS_DIO_CHECKBOX = 0                                   ' 0
Private Const LS_DIO_PUSHBUTTON = 1                                 ' 1
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_NONE = 0                   ' 0x0000
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_NEW_WINDOW = 1             ' 0x0001
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_NAVIGATION = 2             ' 0x0002
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_JAVA = 4                   ' 0x0004
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_SCRIPTING = 8              ' 0x0008
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_ACTIVEX = 16               ' 0x0010
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_ONLINE = 32                ' 0x0020
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_BROWSERCONTEXTMENU = 64         ' 0x0040
Private Const LSMAILVIEW_HTMLRIGHT_ALLOW_PRINT = 128                ' 0x0080

'-------------------- LL22 --------------------------

Private Declare Function LlJobOpen Lib "CMLL22.DLL" (ByVal nLanguage As Long) As Long
Private Declare Function LlJobOpenLCID Lib "CMLL22.DLL" (ByVal nLCID As Long) As Long
Private Declare Sub LlJobClose Lib "CMLL22.DLL" (ByVal hLlJob As Long)
Private Declare Sub LlSetDebug Lib "CMLL22.DLL" (ByVal nOnOff As Long)
Private Declare Function LlGetVersion Lib "CMLL22.DLL" (ByVal nCmd As Long) As Long
Private Declare Function LlGetNotificationMessage Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlSetNotificationMessage Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMessage As Long) As Long
Private Declare Function LlDefineField Lib "CMLL22.DLL" Alias "LlDefineFieldA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal lpbufContents As String) As Long
Private Declare Function LlDefineFieldExt Lib "CMLL22.DLL" Alias "LlDefineFieldExtA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal lpbufContents As String, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDefineFieldExtHandle Lib "CMLL22.DLL" Alias "LlDefineFieldExtHandleA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal hContents As Long, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Sub LlDefineFieldStart Lib "CMLL22.DLL" (ByVal hLlJob As Long)
Private Declare Function LlDefineVariable Lib "CMLL22.DLL" Alias "LlDefineVariableA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal lpbufContents As String) As Long
Private Declare Function LlDefineVariableExt Lib "CMLL22.DLL" Alias "LlDefineVariableExtA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal lpbufContents As String, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDefineVariableExtHandle Lib "CMLL22.DLL" Alias "LlDefineVariableExtHandleA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal hContents As Long, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDefineVariableName Lib "CMLL22.DLL" Alias "LlDefineVariableNameA" (ByVal hLlJob As Long, ByVal pszVarName As String) As Long
Private Declare Sub LlDefineVariableStart Lib "CMLL22.DLL" (ByVal hLlJob As Long)
Private Declare Function LlDefineSumVariable Lib "CMLL22.DLL" Alias "LlDefineSumVariableA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal lpbufContents As String) As Long
Private Declare Function LlDefineLayout Lib "CMLL22.DLL" Alias "LlDefineLayoutA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal pszTitle As String, ByVal nObjType As Long, ByVal pszObjName As String) As Long
Private Declare Function LlDlgEditLine Lib "CMLL22.DLL" Alias "LlDlgEditLineA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal lpBuf As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDlgEditLineEx Lib "CMLL22.DLL" Alias "LlDlgEditLineExA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal pszBuffer As String, ByVal nBufSize As Long, ByVal nParaTypes As Long, ByVal pszTitle As String, ByVal bTable As Integer, ByVal pvReserved As Any) As Long
Private Declare Function LlPreviewSetTempPath Lib "CMLL22.DLL" Alias "LlPreviewSetTempPathA" (ByVal hLlJob As Long, ByVal pszPath As String) As Long
Private Declare Function LlPreviewDeleteFiles Lib "CMLL22.DLL" Alias "LlPreviewDeleteFilesA" (ByVal hLlJob As Long, ByVal pszObjName As String, ByVal pszPath As String) As Long
Private Declare Function LlPreviewDisplay Lib "CMLL22.DLL" Alias "LlPreviewDisplayA" (ByVal hLlJob As Long, ByVal pszObjName As String, ByVal pszPath As String, ByVal Wnd As Long) As Long
Private Declare Function LlPreviewDisplayEx Lib "CMLL22.DLL" Alias "LlPreviewDisplayExA" (ByVal hLlJob As Long, ByVal pszObjName As String, ByVal pszPath As String, ByVal Wnd As Long, ByVal nOptions As Long, ByVal pOptions As Any) As Long
Private Declare Function LlPrint Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintAbort Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintCheckLineFit Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Integer
Private Declare Function LlPrintEnd Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nPages As Long) As Long
Private Declare Function LlPrintFields Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintFieldsEnd Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintGetCurrentPage Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintGetItemsPerPage Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintGetItemsPerTable Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintGetRemainingItemsPerTable Lib "CMLL22.DLL" Alias "LlPrintGetRemainingItemsPerTableA" (ByVal hLlJob As Long, ByVal pszField As String) As Long
Private Declare Function LlPrintGetRemItemsPerTable Lib "CMLL22.DLL" Alias "LlPrintGetRemItemsPerTableA" (ByVal hLlJob As Long, ByVal pszField As String) As Long
Private Declare Function LlPrintGetOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nIndex As Long) As Long
Private Declare Function LlPrintGetPrinterInfo Lib "CMLL22.DLL" Alias "LlPrintGetPrinterInfoA" (ByVal hLlJob As Long, ByVal pszPrn As String, ByVal nPrnLen As Long, ByVal pszPort As String, ByVal nPortLen As Long) As Long
Private Declare Function LlPrintOptionsDialog Lib "CMLL22.DLL" Alias "LlPrintOptionsDialogA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal pszText As String) As Long
Private Declare Function LlPrintSelectOffsetEx Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hwnd As Long) As Long
Private Declare Function LlPrintSetBoxText Lib "CMLL22.DLL" Alias "LlPrintSetBoxTextA" (ByVal hLlJob As Long, ByVal szText As String, ByVal nPercentage As Long) As Long
Private Declare Function LlPrintSetOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal nValue As Long) As Long
Private Declare Function LlPrintUpdateBox Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintStart Lib "CMLL22.DLL" Alias "LlPrintStartA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjName As String, ByVal nPrintOptions As Long, ByVal nReserved As Long) As Long
Private Declare Function LlPrintWithBoxStart Lib "CMLL22.DLL" Alias "LlPrintWithBoxStartA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjName As String, ByVal nPrintOptions As Long, ByVal nBoxType As Long, ByVal hwnd As Long, ByVal pszTitle As String) As Long
Private Declare Function LlPrinterSetup Lib "CMLL22.DLL" Alias "LlPrinterSetupA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal nObjType As Long, ByVal pszObjName As String) As Long
Private Declare Function LlSelectFileDlgTitleEx Lib "CMLL22.DLL" Alias "LlSelectFileDlgTitleExA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal pszTitle As String, ByVal nObjType As Long, ByVal pszObjName As String, ByVal nBufSize As Long, ByVal pReserved As Any) As Long
Private Declare Sub LlSetDlgboxMode Lib "CMLL22.DLL" (ByVal nMode As Long)
Private Declare Function LlGetDlgboxMode Lib "CMLL22.DLL" () As Long
Private Declare Function LlExprParse Lib "CMLL22.DLL" Alias "LlExprParseA" (ByVal hLlJob As Long, ByVal lpExprText As String, ByVal bIncludeFields As Integer) As Long
Private Declare Function LlExprType Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long) As Long
Private Declare Sub LlExprError Lib "CMLL22.DLL" Alias "LlExprErrorA" (ByVal hLlJob As Long, ByVal pszBuf As String, ByVal nBufSize As Long)
Private Declare Sub LlExprFree Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long)
Private Declare Function LlExprEvaluate Lib "CMLL22.DLL" Alias "LlExprEvaluateA" (ByVal hLlJob As Long, ByVal lpExpr As Long, ByVal pszBuf As String, ByVal nBufSize As Long) As Long
Private Declare Function LlExprGetUsedVars Lib "CMLL22.DLL" Alias "LlExprGetUsedVarsA" (ByVal hLlJob As Long, ByVal lpExpr As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlSetOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMode As Long, ByVal nValue As Long) As Long
Private Declare Function LlGetOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMode As Long) As Long
Private Declare Function LlSetOptionString Lib "CMLL22.DLL" Alias "LlSetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String) As Long
Private Declare Function LlGetOptionString Lib "CMLL22.DLL" Alias "LlGetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlPrintSetOptionString Lib "CMLL22.DLL" Alias "LlPrintSetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String) As Long
Private Declare Function LlPrintGetOptionString Lib "CMLL22.DLL" Alias "LlPrintGetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDesignerProhibitAction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMenuID As Long) As Long
Private Declare Function LlDesignerProhibitFunction Lib "CMLL22.DLL" Alias "LlDesignerProhibitFunctionA" (ByVal hLlJob As Long, ByVal pszFunction As String) As Long
Private Declare Function LlPrintEnableObject Lib "CMLL22.DLL" Alias "LlPrintEnableObjectA" (ByVal hLlJob As Long, ByVal pszObjectName As String, ByVal bEnable As Integer) As Long
Private Declare Function LlSetFileExtensions Lib "CMLL22.DLL" Alias "LlSetFileExtensionsA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjectExt As String, ByVal pszPrinterExt As String, ByVal pszSketchExt As String) As Long
Private Declare Function LlPrintGetTextCharsPrinted Lib "CMLL22.DLL" Alias "LlPrintGetTextCharsPrintedA" (ByVal hLlJob As Long, ByVal pszObjectName As String) As Long
Private Declare Function LlPrintGetFieldCharsPrinted Lib "CMLL22.DLL" Alias "LlPrintGetFieldCharsPrintedA" (ByVal hLlJob As Long, ByVal pszObjectName As String, ByVal pszField As String) As Long
Private Declare Function LlPrintIsVariableUsed Lib "CMLL22.DLL" Alias "LlPrintIsVariableUsedA" (ByVal hLlJob As Long, ByVal pszVarName As String) As Long
Private Declare Function LlPrintIsFieldUsed Lib "CMLL22.DLL" Alias "LlPrintIsFieldUsedA" (ByVal hLlJob As Long, ByVal pszFieldName As String) As Long
Private Declare Function LlPrintOptionsDialogTitle Lib "CMLL22.DLL" Alias "LlPrintOptionsDialogTitleA" (ByVal hLlJob As Long, ByVal hwnd As Long, ByVal pszTitle As String, ByVal pszText As String) As Long
Private Declare Function LlSetPrinterToDefault Lib "CMLL22.DLL" Alias "LlSetPrinterToDefaultA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjName As String) As Long
Private Declare Function LlDefineSortOrderStart Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlDefineSortOrder Lib "CMLL22.DLL" Alias "LlDefineSortOrderA" (ByVal hLlJob As Long, ByVal pszIdentifier As String, ByVal pszText As String) As Long
Private Declare Function LlPrintGetSortOrder Lib "CMLL22.DLL" Alias "LlPrintGetSortOrderA" (ByVal hLlJob As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDefineGrouping Lib "CMLL22.DLL" Alias "LlDefineGroupingA" (ByVal hLlJob As Long, ByVal pszSortorder As String, ByVal pszIdentifier As String, ByVal pszText As String) As Long
Private Declare Function LlPrintGetGrouping Lib "CMLL22.DLL" Alias "LlPrintGetGroupingA" (ByVal hLlJob As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlAddCtlSupport Lib "CMLL22.DLL" Alias "LlAddCtlSupportA" (ByVal hwnd As Long, ByVal nFlags As Long, ByVal pszInifile As String) As Long
Private Declare Function LlPrintBeginGroup Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lParam As Long, ByVal lpParam As Any) As Long
Private Declare Function LlPrintEndGroup Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lParam As Long, ByVal lpParam As Any) As Long
Private Declare Function LlPrintGroupLine Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lParam As Long, ByVal lpParam As Any) As Long
Private Declare Function LlPrintGroupHeader Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lParam As Long) As Long
Private Declare Function LlPrintGetFilterExpression Lib "CMLL22.DLL" Alias "LlPrintGetFilterExpressionA" (ByVal hLlJob As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlPrintWillMatchFilter Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlPrintDidMatchFilter Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlGetFieldContents Lib "CMLL22.DLL" Alias "LlGetFieldContentsA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlGetVariableContents Lib "CMLL22.DLL" Alias "LlGetVariableContentsA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlGetSumVariableContents Lib "CMLL22.DLL" Alias "LlGetSumVariableContentsA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlGetUserVariableContents Lib "CMLL22.DLL" Alias "LlGetUserVariableContentsA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlGetVariableType Lib "CMLL22.DLL" Alias "LlGetVariableTypeA" (ByVal hLlJob As Long, ByVal pszName As String) As Long
Private Declare Function LlGetFieldType Lib "CMLL22.DLL" Alias "LlGetFieldTypeA" (ByVal hLlJob As Long, ByVal pszName As String) As Long
Private Declare Function LlPrintGetColumnInfo Lib "CMLL22.DLL" Alias "LlPrintGetColumnInfoA" (ByVal hLlJob As Long, ByVal pszObjectName As String, ByVal nCol As Long, pCol As scLlColumn) As Long
Private Declare Function LlSetPrinterDefaultsDir Lib "CMLL22.DLL" Alias "LlSetPrinterDefaultsDirA" (ByVal hLlJob As Long, ByVal pszDir As String) As Long
Private Declare Function LlCreateSketch Lib "CMLL22.DLL" Alias "LlCreateSketchA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal lpszObjName As String) As Long
Private Declare Function LlViewerProhibitAction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMenuID As Long) As Long
Private Declare Function LlPrintCopyPrinterConfiguration Lib "CMLL22.DLL" Alias "LlPrintCopyPrinterConfigurationA" (ByVal hLlJob As Long, ByVal lpszFilename As String, ByVal nFunction As Long) As Long
Private Declare Function LlSetPrinterInPrinterFile Lib "CMLL22.DLL" Alias "LlSetPrinterInPrinterFileA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjName As String, ByVal nPrinterIndex As Long, ByVal pszPrinter As String, pDevMode As DEVMODEA) As Long
Private Declare Function LlRTFCreateObject Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlRTFDeleteObject Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long) As Long
Private Declare Function LlRTFSetText Lib "CMLL22.DLL" Alias "LlRTFSetTextA" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal pszText As String) As Long
Private Declare Function LlRTFGetTextLength Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal nFlags As Long) As Long
Private Declare Function LlRTFGetText Lib "CMLL22.DLL" Alias "LlRTFGetTextA" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal nFlags As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlRTFEditObject Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal hwnd As Long, ByVal hPrnDC As Long, ByVal nProjectType As Long, ByVal bModal As Integer) As Long
Private Declare Function LlRTFCopyToClipboard Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long) As Long
Private Declare Function LlRTFDisplay Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal hDC As Long, pRC As RECT, ByVal bRestart As Integer, pnState As Long) As Long
Private Declare Function LlRTFEditorProhibitAction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal nControlID As Long) As Long
Private Declare Function LlRTFEditorInvokeAction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal nControlID As Long) As Long
Private Declare Sub LlDebugOutput Lib "CMLL22.DLL" Alias "LlDebugOutputA" (ByVal nIndent As Long, ByVal pszText As String)
Private Declare Function LlEnumGetFirstVar Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nFlags As Long) As Long
Private Declare Function LlEnumGetFirstField Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nFlags As Long) As Long
Private Declare Function LlEnumGetNextEntry Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nPos As Long, ByVal nFlags As Long) As Long
Private Declare Function LlEnumGetEntry Lib "CMLL22.DLL" Alias "LlEnumGetEntryA" (ByVal hLlJob As Long, ByVal nPos As Long, ByVal pszNameBuf As String, ByVal nNameBufSize As Long, ByVal pszContBuf As String, ByVal nContBufSize As Long, pHandle As Long, pType As Long) As Long
Private Declare Function LlPrintResetObjectStates Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlXSetParameter Lib "CMLL22.DLL" Alias "LlXSetParameterA" (ByVal hLlJob As Long, ByVal nExtensionType As Long, ByVal pszExtensionName As String, ByVal pszKey As String, ByVal pszValue As String) As Long
Private Declare Function LlXGetParameter Lib "CMLL22.DLL" Alias "LlXGetParameterA" (ByVal hLlJob As Long, ByVal nExtensionType As Long, ByVal pszExtensionName As String, ByVal pszKey As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlPrintResetProjectState Lib "CMLL22.DLL" (ByVal hJob As Long) As Long
Private Declare Sub LlDefineChartFieldStart Lib "CMLL22.DLL" (ByVal hLlJob As Long)
Private Declare Function LlDefineChartFieldExt Lib "CMLL22.DLL" Alias "LlDefineChartFieldExtA" (ByVal hLlJob As Long, ByVal pszVarName As String, ByVal pszContents As String, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlPrintDeclareChartRow Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nFlags As Long) As Long
Private Declare Function LlPrintGetChartObjectCount Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nType As Long) As Long
Private Declare Function LlPrintIsChartFieldUsed Lib "CMLL22.DLL" Alias "LlPrintIsChartFieldUsedA" (ByVal hLlJob As Long, ByVal pszFieldName As String) As Long
Private Declare Function LlGetChartFieldContents Lib "CMLL22.DLL" Alias "LlGetChartFieldContentsA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlEnumGetFirstChartField Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nFlags As Long) As Long
Private Declare Function LlExprEvaluateVar Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long, pVar As Variant, ByVal nFlags As Long) As Long
Private Declare Function LlExprTypeVar Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long) As Long
Private Declare Function LlGetPrinterFromPrinterFile Lib "CMLL22.DLL" Alias "LlGetPrinterFromPrinterFileA" (ByVal hJob As Long, ByVal nObjType As Long, ByVal pszObjectName As String, ByVal nPrinter As Long, ByVal pszPrinter As String, pnPrinterBufSize As Long, pDevMode As DEVMODEA, pnDevModeBufSize As Long) As Long
Private Declare Function LlPrintGetRemainingSpacePerTable Lib "CMLL22.DLL" Alias "LlPrintGetRemainingSpacePerTableA" (ByVal hLlJob As Long, ByVal pszField As String, ByVal nDimension As Long) As Long
Private Declare Function LlSetDefaultProjectParameter Lib "CMLL22.DLL" Alias "LlSetDefaultProjectParameterA" (ByVal hLlJob As Long, ByVal pszParameter As String, ByVal pszValue As String, ByVal nFlags As Long) As Long
Private Declare Function LlGetDefaultProjectParameter Lib "CMLL22.DLL" Alias "LlGetDefaultProjectParameterA" (ByVal hLlJob As Long, ByVal pszParameter As String, ByVal pszBuffer As String, ByVal nBufSize As Long, pnFlags As Long) As Long
Private Declare Function LlPrintSetProjectParameter Lib "CMLL22.DLL" Alias "LlPrintSetProjectParameterA" (ByVal hLlJob As Long, ByVal pszParameter As String, ByVal pszValue As String, ByVal nFlags As Long) As Long
Private Declare Function LlPrintGetProjectParameter Lib "CMLL22.DLL" Alias "LlPrintGetProjectParameterA" (ByVal hLlJob As Long, ByVal pszParameter As String, ByVal bEvaluated As Integer, ByVal pszBuffer As String, ByVal nBufSize As Long, pnFlags As Long) As Long
Private Declare Function LlExprContainsVariable Lib "CMLL22.DLL" Alias "LlExprContainsVariableA" (ByVal hLlJob As Long, ByVal hExpr As Long, ByVal pszVariable As String) As Long
Private Declare Function LlExprIsConstant Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hExpr As Long) As Long
Private Declare Function LlProfileStart Lib "CMLL22.DLL" Alias "LlProfileStartA" (ByVal hThread As Long, ByVal pszDescr As String, ByVal pszFilename As String, ByVal nTicksMS As Long) As Long
Private Declare Sub LlProfileEnd Lib "CMLL22.DLL" (ByVal hThread As Long)
Private Declare Function LlDbAddTable Lib "CMLL22.DLL" Alias "LlDbAddTableA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszDisplayName As String) As Long
Private Declare Function LlDbAddTableRelation Lib "CMLL22.DLL" Alias "LlDbAddTableRelationA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszParentTableID As String, ByVal pszRelationID As String, ByVal pszRelationDisplayName As String) As Long
Private Declare Function LlDbAddTableSortOrder Lib "CMLL22.DLL" Alias "LlDbAddTableSortOrderA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszSortOrderID As String, ByVal pszSortOrderDisplayName As String) As Long
Private Declare Function LlPrintDbGetCurrentTable Lib "CMLL22.DLL" Alias "LlPrintDbGetCurrentTableA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal nTableIDLength As Long, ByVal bCompletePath As Integer) As Long
Private Declare Function LlPrintDbGetCurrentTableRelation Lib "CMLL22.DLL" Alias "LlPrintDbGetCurrentTableRelationA" (ByVal hJob As Long, ByVal pszRelationID As String, ByVal nRelationIDLength As Long) As Long
Private Declare Function LlPrintDbGetCurrentTableSortOrder Lib "CMLL22.DLL" Alias "LlPrintDbGetCurrentTableSortOrderA" (ByVal hJob As Long, ByVal pszSortOrderID As String, ByVal nSortOrderIDLength As Long) As Long
Private Declare Function LlDbDumpStructure Lib "CMLL22.DLL" (ByVal hJob As Long) As Long
Private Declare Function LlPrintDbGetRootTableCount Lib "CMLL22.DLL" (ByVal hJob As Long) As Long
Private Declare Function LlDbSetMasterTable Lib "CMLL22.DLL" Alias "LlDbSetMasterTableA" (ByVal hJob As Long, ByVal pszTableID As String) As Long
Private Declare Function LlDbGetMasterTable Lib "CMLL22.DLL" Alias "LlDbGetMasterTableA" (ByVal hJob As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlXSetExportParameter Lib "CMLL22.DLL" Alias "LlXSetExportParameterA" (ByVal hLlJob As Long, ByVal pszExtensionName As String, ByVal pszKey As String, ByVal pszValue As String) As Long
Private Declare Function LlXGetExportParameter Lib "CMLL22.DLL" Alias "LlXGetExportParameterA" (ByVal hLlJob As Long, ByVal pszExtensionName As String, ByVal pszKey As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlXlatName Lib "CMLL22.DLL" Alias "LlXlatNameA" (ByVal hLlJob As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDefineVariableVar Lib "CMLL22.DLL" Alias "LlDefineVariableVarA" (ByVal hLlJob As Long, ByVal pszName As String, pValue As Variant, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDefineFieldVar Lib "CMLL22.DLL" Alias "LlDefineFieldVarA" (ByVal hLlJob As Long, ByVal pszName As String, pValue As Variant, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDefineChartFieldVar Lib "CMLL22.DLL" Alias "LlDefineChartFieldVarA" (ByVal hLlJob As Long, ByVal pszName As String, pValue As Variant, ByVal lPara As Long, ByVal lpPtr As Any) As Long
Private Declare Function LlDesignerProhibitEditingObject Lib "CMLL22.DLL" Alias "LlDesignerProhibitEditingObjectA" (ByVal hLlJob As Long, ByVal pszObject As String) As Long
Private Declare Function LlGetUsedIdentifiers Lib "CMLL22.DLL" Alias "LlGetUsedIdentifiersA" (ByVal hLlJob As Long, ByVal pszProjectName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlExprGetUsedVarsEx Lib "CMLL22.DLL" Alias "LlExprGetUsedVarsExA" (ByVal hLlJob As Long, ByVal lpExpr As Long, ByVal pszBuffer As String, ByVal nBufSize As Long, ByVal OrgName As Integer) As Long
Private Declare Function LlDomGetProject Lib "CMLL22.DLL" (ByVal hLlJob As Long, phDOMObj As Long) As Long
Private Declare Function LlDomGetProperty Lib "CMLL22.DLL" Alias "LlDomGetPropertyA" (ByVal hDOMObj As Long, ByVal pszName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDomSetProperty Lib "CMLL22.DLL" Alias "LlDomSetPropertyA" (ByVal hDOMObj As Long, ByVal pszName As String, ByVal pszValue As String) As Long
Private Declare Function LlDomGetObject Lib "CMLL22.DLL" Alias "LlDomGetObjectA" (ByVal hDOMObj As Long, ByVal pszName As String, phDOMSubObj As Long) As Long
Private Declare Function LlDomGetSubobjectCount Lib "CMLL22.DLL" (ByVal hDOMObj As Long, pnCount As Long) As Long
Private Declare Function LlDomGetSubobject Lib "CMLL22.DLL" (ByVal hDOMObj As Long, ByVal nPosition As Long, phDOMSubObj As Long) As Long
Private Declare Function LlDomCreateSubobject Lib "CMLL22.DLL" Alias "LlDomCreateSubobjectA" (ByVal hDOMObj As Long, ByVal nPosition As Long, ByVal pszType As String, phDOMSubObj As Long) As Long
Private Declare Function LlDomDeleteSubobject Lib "CMLL22.DLL" (ByVal hDOMObj As Long, ByVal nPosition As Long) As Long
Private Declare Function LlProjectOpen Lib "CMLL22.DLL" Alias "LlProjectOpenA" (ByVal hLlJob As Long, ByVal nObjType As Long, ByVal pszObjName As String, ByVal nOpenMode As Long) As Long
Private Declare Function LlProjectSave Lib "CMLL22.DLL" Alias "LlProjectSaveA" (ByVal hLlJob As Long, ByVal pszObjName As String) As Long
Private Declare Function LlProjectClose Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlAssociatePreviewControl Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hWndControl As Long, ByVal nFlags As Long) As Long
Private Declare Function LlGetErrortext Lib "CMLL22.DLL" Alias "LlGetErrortextA" (ByVal nError As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlSetPreviewOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nOption As Long, ByVal nValue As Long) As Long
Private Declare Function LlGetPreviewOption Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nOption As Long, pnValue As Long) As Long
Private Declare Function LlDesignerInvokeAction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nMenuID As Long) As Long
Private Declare Function LlDesignerRefreshWorkspace Lib "CMLL22.DLL" (ByVal hLlJob As Long) As Long
Private Declare Function LlDesignerFileOpen Lib "CMLL22.DLL" Alias "LlDesignerFileOpenA" (ByVal hLlJob As Long, ByVal pszFilename As String, ByVal nFlags As Long) As Long
Private Declare Function LlDesignerFileSave Lib "CMLL22.DLL" Alias "LlDesignerFileSaveA" (ByVal hLlJob As Long, ByVal nFlags As Long) As Long
Private Declare Function LlDesignerAddAction Lib "CMLL22.DLL" Alias "LlDesignerAddActionA" (ByVal hLlJob As Long, ByVal nID As Long, ByVal nFlags As Long, ByVal pszMenuText As String, ByVal pszMenuHierarchy As String, ByVal pszTooltipText As String, ByVal nIcon As Long, ByVal pvReserved As Any) As Long
Private Declare Function LlDesignerGetOptionString Lib "CMLL22.DLL" Alias "LlDesignerGetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDesignerSetOptionString Lib "CMLL22.DLL" Alias "LlDesignerSetOptionStringA" (ByVal hLlJob As Long, ByVal nIndex As Long, ByVal pszBuffer As String) As Long
Private Declare Function LlJobOpenCopy Lib "CMLL22.DLL" (ByVal hJob As Long) As Long
Private Declare Function LlGetProjectParameter Lib "CMLL22.DLL" Alias "LlGetProjectParameterA" (ByVal hLlJob As Long, ByVal pszProjectName As String, ByVal pszParameter As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlConvertBLOBToString Lib "CMLL22.DLL" Alias "LlConvertBLOBToStringA" (pBytes As Integer, ByVal nBytes As Long, ByVal pszBuffer As String, ByVal nBufSize As Long, ByVal bWithCompression As Integer) As Long
Private Declare Function LlConvertStringToBLOB Lib "CMLL22.DLL" Alias "LlConvertStringToBLOBA" (ByVal pszText As String, pBytes As Integer, ByVal nBytes As Long) As Long
Private Declare Function LlDbAddTableRelationEx Lib "CMLL22.DLL" Alias "LlDbAddTableRelationExA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszParentTableID As String, ByVal pszRelationID As String, ByVal pszRelationDisplayName As String, ByVal pszKeyField As String, ByVal pszParentKeyField As String) As Long
Private Declare Function LlDbAddTableSortOrderEx Lib "CMLL22.DLL" Alias "LlDbAddTableSortOrderExA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszSortOrderID As String, ByVal pszSortOrderDisplayName As String, ByVal pszField As String) As Long
Private Declare Function LlGetUsedIdentifiersEx Lib "CMLL22.DLL" Alias "LlGetUsedIdentifiersExA" (ByVal hLlJob As Long, ByVal pszProjectName As String, ByVal nIdentifierTypes As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlGetTempFileName Lib "CMLL22.DLL" Alias "LlGetTempFileNameA" (ByVal pszPrefix As String, ByVal pszExt As String, ByVal pszBuffer As String, ByVal nBufSize As Long, ByVal nOptions As Long) As Long
Private Declare Function LlGetDebug Lib "CMLL22.DLL" () As Long
Private Declare Function LlRTFEditorGetRTFControlHandle Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal hRTF As Long) As Long
Private Declare Function LlGetDefaultPrinter Lib "CMLL22.DLL" Alias "LlGetDefaultPrinterA" (ByVal pszPrinter As String, pnPrinterBufSize As Long, pDevMode As DEVMODEA, pnDevModeBufSize As Long, ByVal nOptions As Long) As Long
Private Declare Function LlLocAddDictionaryEntry Lib "CMLL22.DLL" Alias "LlLocAddDictionaryEntryA" (ByVal hLlJob As Long, ByVal nLCID As Long, ByVal pszKey As String, ByVal pszValue As String, ByVal nType As Long) As Long
Private Declare Function LlLocAddDesignLCID Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nLCID As Long) As Long
Private Declare Function LlIsUILanguageAvailable Lib "CMLL22.DLL" (ByVal nLanguage As Long, ByVal nTypesToLookFor As Long) As Long
Private Declare Function LlIsUILanguageAvailableLCID Lib "CMLL22.DLL" (ByVal nLCID As Long, ByVal nTypesToLookFor As Long) As Long
Private Declare Function LlDbAddTableEx Lib "CMLL22.DLL" Alias "LlDbAddTableExA" (ByVal hJob As Long, ByVal pszTableID As String, ByVal pszDisplayName As String, ByVal nOptions As Long) As Long
Private Declare Function LlRTFSetTextEx Lib "CMLL22.DLL" Alias "LlRTFSetTextExA" (ByVal hLlJob As Long, ByVal hRTF As Long, ByVal nFlags As Long, ByVal pszText As String) As Long
Private Declare Function LlInplaceDesignerInteraction Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nAction As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LlGetProjectDescription Lib "CMLL22.DLL" Alias "LlGetProjectDescriptionA" (ByVal hLlJob As Long, ByVal pszProjectName As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlPrintDbGetCurrentTableFilter Lib "CMLL22.DLL" (ByVal hJob As Long, pvFilter As Variant, pvParams As Variant) As Long
Private Declare Function LlExprTranslateToHostExpression Lib "CMLL22.DLL" Alias "LlExprTranslateToHostExpressionA" (ByVal hJob As Long, ByVal hExpr As Long, pvFilter As Variant, pvParams As Variant, ByVal pszCurrentTableName As String, pnFlags As Long) As Long
Private Declare Function LlSRTriggerExport Lib "CMLL22.DLL" (ByVal hJob As Long, ByVal hSessionJob As Long, ByVal pszID As String, ByVal pszExportFormat As String, ByVal nFlags As Long) As Long
Private Declare Function LlUtilsGetVariantFromProfContentsInternal Lib "CMLL22.DLL" (ByVal pszText As String, v As Variant) As Long
Private Declare Function LlUtilsGetProfContentsFromVariantInternal Lib "CMLL22.DLL" (v As Variant, vAsString As Variant) As Long
Private Declare Function LlExprGetUsedFunctions Lib "CMLL22.DLL" Alias "LlExprGetUsedFunctionsA" (ByVal hLlJob As Long, ByVal lpExpr As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlDesignerTriggerJobInUIThread Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal nUserData As Long) As Long
Private Declare Function LlGetUsedIdentifiersExV Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal pszProjectName As String, ByVal nIdentifierTypes As Long, pvArray As Variant) As Long
Private Declare Function LlDomGetPropertyV Lib "CMLL22.DLL" (ByVal hDOMObj As Long, ByVal pszName As String, pvValue As Variant) As Long
Private Declare Function LlExprGetUsedFunctionsV Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long, pvArray As Variant) As Long
Private Declare Function LlExprGetUsedVarsExV Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal lpExpr As Long, pvArray As Variant, ByVal OrgName As Integer) As Long
Private Declare Function LlGetTableRelationToActiveTable Lib "CMLL22.DLL" (ByVal hLlJob As Long, ByVal pszTable As String, ByVal OrgName As Integer, pvValue As Variant) As Long
Private Declare Function LlJobOpenCopyEx Lib "CMLL22.DLL" (ByVal hJob As Long, ByVal nFlags As Long) As Long

'-------------------- Sotrage --------------------------

Private Declare Function LlStgsysStorageOpen Lib "CMLS22.DLL" Alias "LlStgsysStorageOpenA" (ByVal lpszFilename As String, ByVal pszTempPath As String, ByVal bReadOnly As Integer, ByVal bOneJobTranslation As Integer) As Long
Private Declare Sub LlStgsysStorageClose Lib "CMLS22.DLL" (ByVal hStg As Long)
Private Declare Function LlStgsysGetAPIVersion Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysGetFileVersion Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysGetFilename Lib "CMLS22.DLL" Alias "LlStgsysGetFilenameA" (ByVal hStg As Long, ByVal nJob As Long, ByVal nFile As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlStgsysGetJobCount Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysSetJob Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nJob As Long) As Long
Private Declare Function LlStgsysGetJob Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysGetPageCount Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysGetJobOptionValue Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nOption As Long) As Long
Private Declare Function LlStgsysGetPageOptionValue Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nPageIndex As Long, ByVal nOption As Long) As Long
Private Declare Function LlStgsysGetPageOptionString Lib "CMLS22.DLL" Alias "LlStgsysGetPageOptionStringA" (ByVal hStg As Long, ByVal nPageIndex As Long, ByVal nOption As Long, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlStgsysSetPageOptionString Lib "CMLS22.DLL" Alias "LlStgsysSetPageOptionStringA" (ByVal hStg As Long, ByVal nPageIndex As Long, ByVal nOption As Long, ByVal pszBuffer As String) As Long
Private Declare Function LlStgsysAppend Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal hStgToAppend As Long) As Long
Private Declare Function LlStgsysDeleteJob Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nPageIndex As Long) As Long
Private Declare Function LlStgsysDeletePage Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nPageIndex As Long) As Long
Private Declare Function LlStgsysGetPageMetafile Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nPageIndex As Long) As Long
Private Declare Function LlStgsysDestroyMetafile Lib "CMLS22.DLL" (ByVal hMF As Long) As Long
Private Declare Function LlStgsysDrawPage Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal hDC As Long, ByVal hPrnDC As Long, ByVal bAskPrinter As Integer, pRC As RECT, ByVal nPageIndex As Long, ByVal bFit As Integer, ByVal pReserved As Any) As Long
Private Declare Function LlStgsysGetLastError Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysDeleteFiles Lib "CMLS22.DLL" (ByVal hStg As Long) As Long
Private Declare Function LlStgsysPrint Lib "CMLS22.DLL" Alias "LlStgsysPrintA" (ByVal hStg As Long, ByVal pszPrinterName1 As String, ByVal pszPrinterName2 As String, ByVal nStartPageIndex As Long, ByVal nEndPageIndex As Long, ByVal nCopies As Long, ByVal nFlags As Long, ByVal pszMessage As String, ByVal hWndParent As Long) As Long
Private Declare Function LlStgsysStoragePrint Lib "CMLS22.DLL" Alias "LlStgsysStoragePrintA" (ByVal lpszFilename As String, ByVal pszTempPath As String, ByVal pszPrinterName1 As String, ByVal pszPrinterName2 As String, ByVal nStartPageIndex As Long, ByVal nEndPageIndex As Long, ByVal nCopies As Long, ByVal nFlags As Long, ByVal pszMessage As String, ByVal hWndParent As Long) As Long
Private Declare Function LlStgsysGetPagePrinter Lib "CMLS22.DLL" Alias "LlStgsysGetPagePrinterA" (ByVal hStg As Long, ByVal nPageIndex As Long, ByVal pszDeviceName As String, ByVal nDeviceNameSize As Long, phDevMode As Long) As Long
Private Declare Sub LsSetDebug Lib "CMLS22.DLL" (ByVal bOn As Integer)
Private Declare Function LsGetViewerControlClassName Lib "CMLS22.DLL" Alias "LsGetViewerControlClassNameA" () As String
Private Declare Function LsGetViewerControlDefaultMessage Lib "CMLS22.DLL" () As Long
Private Declare Function LsCreateViewerControlOverParent Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal hParentControl As Long) As Long
Private Declare Function LlStgsysGetJobOptionStringEx Lib "CMLS22.DLL" Alias "LlStgsysGetJobOptionStringExA" (ByVal hStg As Long, ByVal pszKey As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LlStgsysSetJobOptionStringEx Lib "CMLS22.DLL" Alias "LlStgsysSetJobOptionStringExA" (ByVal hStg As Long, ByVal pszKey As String, ByVal pszBuffer As String) As Long
Private Declare Function LsConversionJobOpen Lib "CMLS22.DLL" Alias "LsConversionJobOpenA" (ByVal hWndParent As Long, ByVal nLanguage As Long, ByVal pszFormat As String) As Long
Private Declare Function LsConversionJobClose Lib "CMLS22.DLL" (ByVal hCnvJob As Long) As Long
Private Declare Function LsConversionConfigurationDlg Lib "CMLS22.DLL" (ByVal hCnvJob As Long, ByVal hWndParent As Long) As Long
Private Declare Function LsConversionSetOptionString Lib "CMLS22.DLL" Alias "LsConversionSetOptionStringA" (ByVal hCnvJob As Long, ByVal pszKey As String, ByVal pszData As String) As Long
Private Declare Function LsConversionGetOptionString Lib "CMLS22.DLL" Alias "LsConversionGetOptionStringA" (ByVal hCnvJob As Long, ByVal pszKey As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LsConversionConvertEMFToFile Lib "CMLS22.DLL" Alias "LsConversionConvertEMFToFileA" (ByVal hCnvJob As Long, ByVal hEMF As Long, ByVal pszFilename As String) As Long
Private Declare Function LsConversionConvertStgToFile Lib "CMLS22.DLL" Alias "LsConversionConvertStgToFileA" (ByVal hCnvJob As Long, ByVal hStg As Long, ByVal pszFilename As String) As Long
Private Declare Function LlStgsysStorageConvert Lib "CMLS22.DLL" Alias "LlStgsysStorageConvertA" (ByVal pszStgFilename As String, ByVal pszDstFilename As String, ByVal pszFormat As String) As Long
Private Declare Function LlStgsysConvert Lib "CMLS22.DLL" Alias "LlStgsysConvertA" (ByVal hStg As Long, ByVal pszDstFilename As String, ByVal pszFormat As String) As Long
Private Declare Function LsMailConfigurationDialog Lib "CMLS22.DLL" Alias "LsMailConfigurationDialogA" (ByVal hWndParent As Long, ByVal pszSubkey As String, ByVal nFlags As Long, ByVal nLanguage As Long) As Long
Private Declare Function LsMailJobOpen Lib "CMLS22.DLL" (ByVal nLanguage As Long) As Long
Private Declare Function LsMailJobClose Lib "CMLS22.DLL" (ByVal hJob As Long) As Long
Private Declare Function LsMailSetOptionString Lib "CMLS22.DLL" Alias "LsMailSetOptionStringA" (ByVal hJob As Long, ByVal pszKey As String, ByVal pszValue As String) As Long
Private Declare Function LsMailGetOptionString Lib "CMLS22.DLL" Alias "LsMailGetOptionStringA" (ByVal hJob As Long, ByVal pszKey As String, ByVal pszBuffer As String, ByVal nBufSize As Long) As Long
Private Declare Function LsMailSendFile Lib "CMLS22.DLL" (ByVal hJob As Long, ByVal hWndParent As Long) As Long
Private Declare Function LlStgsysStorageCreate Lib "CMLS22.DLL" Alias "LlStgsysStorageCreateA" (ByVal lpszFilename As String, ByVal pszTempPath As String, ByVal hRefDC As Long, prcArea As RECT, ByVal bPhysicalPage As Integer) As Long
Private Declare Function LlStgsysAppendEMF Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal hEMFToAppend As Long, ByVal hRefDC As Long, prcArea As RECT, ByVal bPhysicalPage As Integer) As Long
Private Declare Function LsProfileStart Lib "CMLS22.DLL" Alias "LsProfileStartA" (ByVal hThread As Long, ByVal pszDescr As String, ByVal pszFilename As String, ByVal nTicksMS As Long) As Long
Private Declare Sub LsProfileEnd Lib "CMLS22.DLL" (ByVal hThread As Long)
Private Declare Function LsMailView Lib "CMLS22.DLL" Alias "LsMailViewA" (ByVal hWndParent As Long, ByVal pszMailFile As String, ByVal nRights As Long, ByVal nLanguage As Long) As Long
Private Declare Function LsInternalCreateViewerControlOverParent13 Lib "CMLS22.DLL" (ByVal hParentControl As Long, ByVal nFlags As Long) As Long
Private Declare Function LsInternalGetViewerControlFromParent13 Lib "CMLS22.DLL" (ByVal hParentControl As Long) As Long
Private Declare Sub LsSetDlgboxMode Lib "CMLS22.DLL" (ByVal nMode As Long)
Private Declare Function LsGetDlgboxMode Lib "CMLS22.DLL" () As Long
Private Declare Function LsGetViewerControlClassNameEx Lib "CMLS22.DLL" () As String
Private Declare Function LsGetDebug Lib "CMLS22.DLL" () As Integer
Private Declare Sub LsJobClose Lib "CMLS22.DLL" (ByVal hLsJob As Long)
Private Declare Function LlStgsysSetUILanguage Lib "CMLS22.DLL" (ByVal hStg As Long, ByVal nCMBTLanguage As Long) As Integer

'-------------------- DevMode --------------------------

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODEA) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Const HORZSIZE = 4 'Horizontal size in millimeters
Private Const VERTSIZE = 6 'Vertical size in millimeters
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_SPECVERSION = &H320
    
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_SCALE = &H10&
Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_PRINTQUALITY = &H400&
Private Const DM_COLOR = &H800&
Private Const DM_DUPLEX = &H1000&
Private Const DM_YRESOLUTION = &H2000&
Private Const DM_TTOPTION = &H4000&
Private Const DM_COLLATE As Long = &H8000
Private Const DM_FORMNAME As Long = &H10000

'orientation selections
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2

'paper selections
Private Const DMPAPER_LETTER = 1
Private Const DMPAPER_FIRST = DMPAPER_LETTER     '  Letter 8 1/2 x 11 in
Private Const DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
Private Const DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
Private Const DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
Private Const DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
Private Const DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
Private Const DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
Private Const DMPAPER_A3 = 8                     '  A3 297 x 420 mm
Private Const DMPAPER_A4 = 9                     '  A4 210 x 297 mm
Private Const DMPAPER_A5 = 11                    '  A5 148 x 210 mm
Private Const DMPAPER_A6 = 70                    '  DIN A6 Blatt
Private Const DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
Private Const DMPAPER_B4 = 12                    '  B4 250 x 354
Private Const DMPAPER_B5 = 13                    '  B5 182 x 257 mm
Private Const DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
Private Const DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
Private Const DMPAPER_10X14 = 16                 '  10x14 in
Private Const DMPAPER_11X17 = 17                 '  11x17 in
Private Const DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
Private Const DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
Private Const DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
Private Const DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
Private Const DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
Private Const DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
Private Const DMPAPER_CSHEET = 24                '  C size sheet
Private Const DMPAPER_DSHEET = 25                '  D size sheet
Private Const DMPAPER_ESHEET = 26                '  E size sheet
Private Const DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
Private Const DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
Private Const DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
Private Const DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
Private Const DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
Private Const DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
Private Const DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
Private Const DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
Private Const DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
Private Const DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
Private Const DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
Private Const DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
Private Const DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in

Private Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN
Private Const DMPAPER_USER = 256

'bin selections
Private Const DMBIN_UPPER = 1
Private Const DMBIN_FIRST = DMBIN_UPPER

Private Const DMBIN_ONLYONE = 1
Private Const DMBIN_LOWER = 2
Private Const DMBIN_MIDDLE = 3
Private Const DMBIN_MANUAL = 4
Private Const DMBIN_ENVELOPE = 5
Private Const DMBIN_ENVMANUAL = 6
Private Const DMBIN_AUTO = 7
Private Const DMBIN_TRACTOR = 8
Private Const DMBIN_SMALLFMT = 9
Private Const DMBIN_LARGEFMT = 10
Private Const DMBIN_LARGECAPACITY = 11
Private Const DMBIN_CASSETTE = 14
Private Const DMBIN_LAST = DMBIN_CASSETTE

Private Const DMBIN_USER = 256 'device specific bins start here

'print qualities
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
Private Const DMRES_HIGH = (-4)

'color enable/disable for color printers
Private Const DMCOLOR_MONOCHROME = 1
Private Const DMCOLOR_COLOR = 2

'duplex enable
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3

'TrueType options
Private Const DMTT_BITMAP = 1    'print TT fonts as graphics
Private Const DMTT_DOWNLOAD = 2  'download TT fonts as soft fonts
Private Const DMTT_SUBDEV = 3    'substitute device fonts for TT fonts

'Collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

'dmDisplayFlags flags
Private Const DM_GRAYSCALE = &H1
Private Const DM_INTERLACED = &H2

'-------------------- Ende --------------------------

Private Const TeTit = "Keine Daten"
Private Const TeMai = "Keine zu druckenden Daten"
Private Const TeInh = "Der Auftrag enthält keine Daten, die gedruckt oder exportiert werden könnten. Bitte prüfen Sie zum Beispiel, ob die richtige Rechnung gewählt wurde."
Private Const TeFus = "Beim Drucken eines Krankenblattes kann es vorkommen, dass dieses für den gewählten Zeitraum, keine Einträge enthält."

Private mHwnd As Long   'Jobhandle
Private eHwnd As Long   'Jobhandle Email
Private IdxNr As Long   'Indexnummer
Private ManNr As Long   'Mandantennummer
Private MitNr As Long   'Mitarbeiternummer
Private ArzNr As Long   'Hausarztnummer
Private MaRec As Long   'Rechnung zur Mahnung
Private AnSal As Single 'Anfangssaldo
Private MaGeb As Single 'Mahngebühren
Private ForNa As String 'Formularname
Private FiNam As String 'Dateiname und Pfad
Private TmPfa As String 'Pfadangabe Temp
Private DrNam As String 'Druckername
Private ExFmt As String 'Exportformat
Private StVer As String 'Startverzeichnis
Private DaNam As String 'Exportdateiname
Private EmAdr As String 'Emailadresse
Private EmBet As String 'Emailbetreff
Private EmTex As String 'Emailtext
Private EmNam As String 'Email Absendername
Private Krite As String 'Kriterienausdruck
Private KopTe As String 'Kopftex
Private FaxNr As String 'Faxnummer
Private EmSep As Boolean 'Emailempfängerseparation
Private DrVor As Boolean 'Druckvorschau
Private Gesam As Boolean 'Alle Datensätze
Private ChaOb As Boolean 'Chartobject
Private ManVo As Boolean 'Mandant vorhanden
Private MitVo As Boolean 'Mitarbeiter vorhanden
Private ArzVo As Boolean 'Verordner vorhanden
Private DrSof As Boolean 'Sofortdruck
Private DrKop As Boolean 'Druckkopie
Private DrDia As Boolean 'Druckdialog
Private Wiede As Boolean 'Eintrag Wiederholen
Private Proto As Boolean 'Protokolleinträge
Private EmDia As Boolean 'Emaildialog
Private OhGra As Integer 'Ohne Hintergrundgrafik
Private EmSen As Integer 'Emailversand
Private WieAn As Integer 'Anzahl Wiederholungen
Private ZeRau As Integer 'Zeitraumtyp
Private Kopie As Integer 'Kopien
Private Grupp As Integer 'Gruppierung
Private FoIdx As Integer

Private RS30 As ADODB.Recordset
Private RS31 As ADODB.Recordset
Private RS32 As ADODB.Recordset
Private RS33 As ADODB.Recordset
Private RS34 As ADODB.Recordset
Private RS35 As ADODB.Recordset
Private RS36 As ADODB.Recordset
Private RS37 As ADODB.Recordset
Private CM1 As ADODB.Command
Private FL1 As ADODB.Field

Private clFil As clsFile
Public Property Let AnfSal(Flag As Single)
    AnSal = Flag
End Property
Public Property Get AnfSal() As Single
    AnfSal = AnSal
End Property
Public Property Let ArztNr(Flag As Long)
    ArzNr = Flag
End Property
Public Property Get ArztNr() As Long
    ArztNr = ArzNr
End Property
Public Property Let ArztVo(Flag As Boolean)
    ArzVo = Flag
End Property
Public Property Get ArztVo() As Boolean
    ArztVo = ArzVo
End Property
Public Property Let ForIdx(Flag As Integer)
    FoIdx = Flag
End Property
Public Property Get ForIdx() As Integer
    ForIdx = FoIdx
End Property
Public Property Let MandNr(Flag As Long)
    ManNr = Flag
End Property
Public Property Get MandNr() As Long
    MandNr = ManNr
End Property
Public Property Let MitaNr(Flag As Long)
    MitNr = Flag
End Property
Public Property Get MitaNr() As Long
    MitaNr = MitNr
End Property
Public Property Let MandVo(Flag As Boolean)
    ManVo = Flag
End Property
Public Property Get MandVo() As Boolean
    MandVo = ManVo
End Property
Public Property Let MitaVo(Flag As Boolean)
    MitVo = Flag
End Property
Public Property Get MitaVo() As Boolean
    MitaVo = MitVo
End Property
Public Property Let ChaObj(Flag As Boolean)
    ChaOb = Flag
End Property
Public Property Get ChaObj() As Boolean
    ChaObj = ChaOb
End Property
Public Property Let DatNam(Flag As String)
    DaNam = Flag
End Property
Public Property Get DatNam() As String
    DatNam = DaNam
End Property
Public Property Get DruDia() As Boolean
    DruDia = DrDia
End Property
Public Property Let DruDia(Flag As Boolean)
    DrDia = Flag
End Property
Public Property Let DruNam(Flag As String)
    DrNam = Flag
End Property
Public Property Get DruNam() As String
    DruNam = DrNam
End Property
Public Property Get DruSof() As Boolean
    DruSof = DrSof
End Property
Public Property Let DruSof(Flag As Boolean)
    DrSof = Flag
End Property
Public Property Let DruVor(Flag As Boolean)
    DrVor = Flag
End Property
Public Property Get DruVor() As Boolean
    DruVor = DrVor
End Property
Public Property Let EmlAdr(Flag As String)
    EmAdr = Flag
End Property
Public Property Get EmlAdr() As String
    EmlAdr = EmAdr
End Property
Public Property Get EmlBet() As String
    EmlBet = EmBet
End Property

Public Property Let EmlBet(Flag As String)
    EmBet = Flag
End Property
Public Property Get EmlDia() As Boolean
    EmlDia = EmDia
End Property
Public Property Let EmlDia(Flag As Boolean)
    EmDia = Flag
End Property
Public Property Get EmlSen() As Integer
    EmlSen = EmSen
End Property
Public Property Let EmlSen(Flag As Integer)
    EmSen = Flag
End Property

Public Property Let EmlNam(Flag As String)
    EmNam = Flag
End Property
Public Property Get EmlNam() As String
    EmlNam = EmNam
End Property
Public Property Let EmlSeP(Flag As Boolean)
    EmSep = Flag
End Property
Public Property Get EmlSeP() As Boolean
    EmlSeP = EmSep
End Property
Public Property Let EmlTex(Flag As String)
    EmTex = Flag
End Property
Public Property Get EmlTex() As String
    EmlTex = EmTex
End Property
Public Property Let ExpFmt(Flag As String)
    ExFmt = Flag
End Property
Public Property Get ExpFmt() As String
    ExpFmt = ExFmt
End Property
Public Property Let FaxNum(Flag As String)
    FaxNr = Flag
End Property
Public Property Get FaxNum() As String
    FaxNum = FaxNr
End Property
Public Property Let FilNam(Flag As String)
    FiNam = Flag
End Property
Public Property Get FilNam() As String
    FilNam = FiNam
End Property
Public Property Let ForNam(Flag As String)
    ForNa = Flag
End Property
Public Property Get ForNam() As String
    ForNam = ForNa
End Property
Public Property Let Gesamt(Flag As Boolean)
    Gesam = Flag
End Property
Public Property Get Gesamt() As Boolean
    Gesamt = Gesam
End Property
Public Property Let Gruppe(Flag As Integer)
    Grupp = Flag
End Property
Public Property Get Gruppe() As Integer
    Gruppe = Grupp
End Property
Public Property Let IndxNr(Flag As Long)
    IdxNr = Flag
End Property
Public Property Get IndxNr() As Long
    IndxNr = IdxNr
End Property
Public Property Let Kopien(Flag As Integer)
    Kopie = Flag
End Property
Public Property Get Kopien() As Integer
    Kopien = Kopie
End Property
Public Property Let KopTex(Flag As String)
    KopTe = Flag
End Property
Public Property Get KopTex() As String
    KopTex = KopTe
End Property
Public Property Let Kriter(Flag As String)
    Krite = Flag
End Property
Public Property Get Kriter() As String
    Kriter = Krite
End Property
Public Property Let MahGeb(Flag As Single)
    MaGeb = Flag
End Property
Public Property Get MahGeb() As Single
    MahGeb = MaGeb
End Property
Public Property Let MahRec(Flag As Long)
    MaRec = Flag
End Property
Public Property Get MahRec() As Long
    MahRec = MaRec
End Property
Public Property Let OhnGra(Flag As Integer)
    OhGra = Flag
End Property
Public Property Get OhnGra() As Integer
    OhnGra = OhGra
End Property
Private Sub LDeFel(ByVal RpRow As XtremeReportControl.ReportRow, ByVal Flag As Integer)
On Error GoTo DeErr
'Definiert die Formularvariablen für ListLabel

Dim RetWe As Long
Dim TyStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo3 = FM.repCont3
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6

LlDefineFieldStart mHwnd

Select Case Flag
Case 1:
    Set RpCls = RpCo1.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                If RpCol.Tag = 1 Then
                    TyStr = LL_BOOLEAN
                Else
                    If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) Then
                        If InStr(1, RpRow.Record(RpCol.ItemIndex).Value, ",", vbTextCompare) > 0 Then
                            TyStr = LL_NUMERIC
                        Else
                            TyStr = LL_TEXT
                        End If
                    Else
                        TyStr = LL_TEXT
                    End If
                End If
                If TyStr = LL_NUMERIC Then
                    RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, CSng(RpRow.Record(RpCol.ItemIndex).Value), TyStr, vbNullString)
                Else
                    RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
                End If
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, vbNullString, LL_TEXT, vbNullString)
            End If
        End If
    Next RpCol
Case 2:
    Set RpCls = RpCo2.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If AdAry(RpCol.Index, RpRow.Index) <> vbNullString Then
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, AdAry(RpCol.Index, RpRow.Index), TyStr, vbNullString)
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 3:
    Set RpCls = RpCo3.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 4:
    Set RpCls = RpCo4.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                If RpCol.Tag = 1 Then
                    TyStr = LL_BOOLEAN
                Else
                    If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) Then
                        If InStr(1, RpRow.Record(RpCol.ItemIndex).Value, ",", vbTextCompare) > 0 Then
                            TyStr = LL_NUMERIC
                        Else
                            TyStr = LL_TEXT
                        End If
                    Else
                        TyStr = LL_TEXT
                    End If
                End If
                If TyStr = LL_NUMERIC Then
                    RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, CSng(RpRow.Record(RpCol.ItemIndex).Value), TyStr, vbNullString)
                Else
                    RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
                End If
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, vbNullString, LL_TEXT, vbNullString)
            End If
        End If
    Next RpCol
Case 5:
    Set RpCls = RpCo5.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 6:
    Set RpCls = RpCo6.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineFieldExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
End Select

If Grupp > 0 Then RetWe = LlDefineFieldExt(mHwnd, "Grupp", Grupp, LL_NUMERIC, vbNullString)

Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LDeFel " & Err.Number
Resume Next

End Sub
Private Sub LDeVar(ByVal RpRow As XtremeReportControl.ReportRow, ByVal Flag As Integer)
On Error GoTo DeErr
'Definiert die Formularvariablen für ListLabel

Dim RetWe As Long
Dim TyStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo8 = FM.repCont8
Set RpCo3 = FM.repCont3
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6

LlDefineVariableStart mHwnd

Select Case Flag
Case 1:
    Set RpCls = RpCo1.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
                Else
                    RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, vbNullString, TyStr, vbNullString)
                End If
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, vbNullString, TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 2:
    Set RpCls = RpCo2.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If AdAry(RpCol.Index, RpRow.Index) <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, AdAry(RpCol.Index, RpRow.Index), TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 3:
    Set RpCls = RpCo3.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 4:
    Set RpCls = RpCo4.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 5:
    Set RpCls = RpCo5.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 6:
    Set RpCls = RpCo6.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
Case 8:
    Set RpCls = RpCo8.Columns
    For Each RpCol In RpCls
        If RpRow.GroupRow = False Then
            If RpCol.Tag = 1 Then
                TyStr = LL_BOOLEAN
            Else
                TyStr = LL_TEXT
            End If
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, RpRow.Record(RpCol.ItemIndex).Value, TyStr, vbNullString)
            Else
                RetWe = LlDefineVariableExt(mHwnd, RpCol.Caption, "", TyStr, vbNullString)
            End If
        End If
    Next RpCol
End Select

If ManVo = True Then 'Mandanten vorhanden
    If ManNr > 0 Then
        RetWe = LlDefineVariableExt(mHwnd, "AnwVorname", GlThe(ManNr, 1), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwName", GlThe(ManNr, 2), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwStraße", GlThe(ManNr, 3), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwPLZ", GlThe(ManNr, 4), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwOrt", GlThe(ManNr, 5), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTelefon", GlThe(ManNr, 6), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTelefax", GlThe(ManNr, 7), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBank", GlThe(ManNr, 8), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ", GlThe(ManNr, 9), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKonto", GlThe(ManNr, 10), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwSteuer", GlThe(ManNr, 11), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBeruf", GlThe(ManNr, 12), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTitel", GlThe(ManNr, 14), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIKNr", GlThe(ManNr, 15), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwLANR", GlThe(ManNr, 15), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwEmail", GlThe(ManNr, 16), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwInternet", GlThe(ManNr, 17), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN", GlThe(ManNr, 18), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwPraxis", GlThe(ManNr, 19), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBank2", GlThe(ManNr, 20), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ2", GlThe(ManNr, 21), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKonto2", GlThe(ManNr, 22), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN2", GlThe(ManNr, 23), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwSprechzeiten", GlThe(ManNr, 24), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBIC", GlThe(ManNr, 31), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBIC2", GlThe(ManNr, 32), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGID", GlThe(ManNr, 33), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwAbsender", GlThe(ManNr, 36), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGirocode", GlThe(ManNr, 10), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGLN", GlThe(ManNr, 40), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwZSR", GlThe(ManNr, 41), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBunLan", GlThe(ManNr, 42), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKVBezirk", GlThe(ManNr, 43), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKanton", GlThe(ManNr, 44), LL_TEXT, vbNullString)
    Else
        LLReg
    End If
Else
    LLReg
End If

If GlMiV = True Then 'Mitarbeiter vorhanden
    If MitVo = True Then 'Mitarbeiter vorhanden
        If MitNr > 0 Then
            RetWe = LlDefineVariableExt(mHwnd, "MitNr", GlMiK(MitNr, 2), LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitName", GlMiK(MitNr, 1), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitNachname", GlMiK(MitNr, 3), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitVorname", GlMiK(MitNr, 4), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitEmail", GlMiK(MitNr, 22), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", GlMiK(MitNr, 29), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", GlMiK(MitNr, 30), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitGLN", GlMiK(MitNr, 31), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitZSR", GlMiK(MitNr, 32), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", GlMiK(MitNr, 33), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitStraße", GlMiK(MitNr, 34), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", GlMiK(MitNr, 35), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitOrt", GlMiK(MitNr, 36), LL_TEXT, vbNullString)
        Else
            RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
End If

If GlArV = True Then 'Verordner vorhanden
    If ArzVo = True Then 'Verordner vorhanden
        If ArzNr > 0 Then
            RetWe = LlDefineVariableExt(mHwnd, "ArzNr", GlArz(ArzNr, 0), LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", GlArz(ArzNr, 1), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", GlArz(ArzNr, 2), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", GlArz(ArzNr, 6), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", GlArz(ArzNr, 7), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzName", GlArz(ArzNr, 8), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", GlArz(ArzNr, 9), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", GlArz(ArzNr, 10), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", GlArz(ArzNr, 11), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", GlArz(ArzNr, 12), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", GlArz(ArzNr, 12), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", GlArz(ArzNr, 13), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", GlArz(ArzNr, 14), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", GlArz(ArzNr, 15), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", GlArz(ArzNr, 16), LL_TEXT, vbNullString)
        Else
            RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
End If

RetWe = LlDefineVariableExt(mHwnd, "Lizenziert", CInt(GlLiz), LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Druckkopie", DrKop, LL_BOOLEAN, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Katalog", GlFri, LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Bilderordner", GlBPf, LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Dokumentenordner", GlDox, LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "OhneGrafik", OhGra, LL_NUMERIC, vbNullString)

If GlTSN <> vbNullString Then
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", GlTSN, LL_TEXT, vbNullString)
Else
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", "1234", LL_TEXT, vbNullString)
End If

If EmSen > 0 Then
    RetWe = LlDefineVariableExt(mHwnd, "Emailversand", True, LL_BOOLEAN, vbNullString)
Else
    RetWe = LlDefineVariableExt(mHwnd, "Emailversand", False, LL_BOOLEAN, vbNullString)
End If

Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCo8 = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LDeVar " & Err.Number
Resume Next

End Sub
Private Sub LLDeB()
On Error GoTo DeErr
'Übergibt Behandlerdaten an Formulardesigner

Dim RetWe As Long

LlDefineVariableStart mHwnd

If ManVo = True Then 'Mandanten vorhanden
    If ManNr > 0 Then
        RetWe = LlDefineVariableExt(mHwnd, "AnwVorname", GlThe(ManNr, 1), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwName", GlThe(ManNr, 2), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwStraße", GlThe(ManNr, 3), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwPLZ", GlThe(ManNr, 4), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwOrt", GlThe(ManNr, 5), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTelefon", GlThe(ManNr, 6), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTelefax", GlThe(ManNr, 7), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBank", GlThe(ManNr, 8), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ", GlThe(ManNr, 9), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKonto", GlThe(ManNr, 10), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwSteuer", GlThe(ManNr, 11), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBeruf", GlThe(ManNr, 12), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwTitel", GlThe(ManNr, 14), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIKNr", GlThe(ManNr, 15), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwLANR", GlThe(ManNr, 15), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwEmail", GlThe(ManNr, 16), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwInternet", GlThe(ManNr, 17), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN", GlThe(ManNr, 18), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwPraxis", GlThe(ManNr, 19), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBank2", GlThe(ManNr, 20), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ2", GlThe(ManNr, 21), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKonto2", GlThe(ManNr, 22), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN2", GlThe(ManNr, 23), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwSprechzeiten", GlThe(ManNr, 24), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBIC", GlThe(ManNr, 31), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBIC2", GlThe(ManNr, 32), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGID", GlThe(ManNr, 33), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwAbsender", GlThe(ManNr, 36), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGirocode", GlThe(ManNr, 10), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwGLN", GlThe(ManNr, 40), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwZSR", GlThe(ManNr, 41), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwBunLan", GlThe(ManNr, 42), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKVBezirk", GlThe(ManNr, 43), LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "AnwKanton", GlThe(ManNr, 44), LL_TEXT, vbNullString)
    Else
        LLReg
    End If
Else
    LLReg
End If

If GlMiV = True Then 'Mitarbeiter vorhanden
    If MitVo = True Then 'Mitarbeiter vorhanden
        If MitNr > 0 Then
            RetWe = LlDefineVariableExt(mHwnd, "MitNr", GlMiK(MitNr, 2), LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitName", GlMiK(MitNr, 1), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitNachname", GlMiK(MitNr, 3), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitVorname", GlMiK(MitNr, 4), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitEmail", GlMiK(MitNr, 22), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", GlMiK(MitNr, 29), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", GlMiK(MitNr, 30), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitGLN", GlMiK(MitNr, 31), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitZSR", GlMiK(MitNr, 32), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", GlMiK(MitNr, 33), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitStraße", GlMiK(MitNr, 34), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", GlMiK(MitNr, 35), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitOrt", GlMiK(MitNr, 36), LL_TEXT, vbNullString)
        Else
            RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
End If

If GlArV = True Then 'Verordner vorhanden
    If ArzVo = True Then 'Verordner vorhanden
        If ArzNr > 0 Then
            RetWe = LlDefineVariableExt(mHwnd, "ArzNr", GlArz(ArzNr, 0), LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", GlArz(ArzNr, 1), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", GlArz(ArzNr, 2), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", GlArz(ArzNr, 6), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", GlArz(ArzNr, 7), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzName", GlArz(ArzNr, 8), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", GlArz(ArzNr, 9), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", GlArz(ArzNr, 10), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", GlArz(ArzNr, 11), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", GlArz(ArzNr, 12), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", GlArz(ArzNr, 12), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", GlArz(ArzNr, 13), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", GlArz(ArzNr, 14), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", GlArz(ArzNr, 15), LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", GlArz(ArzNr, 16), LL_TEXT, vbNullString)
        Else
            RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
End If

RetWe = LlDefineVariableExt(mHwnd, "Lizenziert", CInt(GlLiz), LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Druckkopie", DrKop, LL_BOOLEAN, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Katalog", GlFri, LL_NUMERIC, vbNullString)

If GlTSN <> vbNullString Then
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", GlTSN, LL_TEXT, vbNullString)
Else
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", "1234", LL_TEXT, vbNullString)
End If

If KopTe <> vbNullString Then
    RetWe = LlDefineVariableExt(mHwnd, "Kopftext", KopTe, LL_TEXT, vbNullString)
End If

RetWe = LlDefineVariableExt(mHwnd, "Übertrag", AnSal, LL_TEXT, vbNullString)

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLDeB " & Err.Number
Resume Next

End Sub
Private Sub LLDeC(ByRef RST As ADODB.Recordset)
On Error GoTo DeErr
'Definiert die Chartvariablen für ListLabel

Dim RetWe As Long
Dim TyStr As String

LlDefineChartFieldStart mHwnd

For Each FL1 In RST.Fields
    If FL1.Type <> 128 Then
        Select Case FL1.Type
        Case adBigInt: TyStr = LL_NUMERIC
        Case adBoolean: TyStr = LL_BOOLEAN
        Case adChar: TyStr = LL_TEXT
        Case adVarChar: TyStr = LL_TEXT
        Case adCurrency: TyStr = LL_NUMERIC
        Case adDate: TyStr = LL_TEXT
        Case adDecimal: TyStr = LL_NUMERIC
        Case adDouble: TyStr = LL_NUMERIC
        Case adInteger: TyStr = LL_NUMERIC
        Case adNumeric: TyStr = LL_NUMERIC
        Case adSingle: TyStr = LL_NUMERIC
        Case adSmallInt: TyStr = LL_NUMERIC
        Case Else: TyStr = LL_TEXT
        End Select
        If FL1.Value <> vbNullString Then
            RetWe = LlDefineChartFieldExt(mHwnd, FL1.Name, FL1.Value, TyStr, vbNullString)
        Else
            If TyStr = LL_NUMERIC Then
                RetWe = LlDefineChartFieldExt(mHwnd, FL1.Name, 0, TyStr, vbNullString)
            Else
                RetWe = LlDefineChartFieldExt(mHwnd, FL1.Name, "", TyStr, vbNullString)
            End If
        End If
    End If
Next FL1

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLDeC " & Err.Number
Resume Next

End Sub
Private Sub LLDeF(ByRef RST As ADODB.Recordset, Optional ByVal DauDg As String)
On Error GoTo DeErr
'Definiert die Feldvariablen für ListLabel

Dim RetWe As Long
Dim TyStr As String

LlDefineFieldStart mHwnd

For Each FL1 In RST.Fields
    If FL1.Type <> 128 Then
        Select Case FL1.Type
        Case adBigInt: TyStr = LL_NUMERIC
        Case adBoolean: TyStr = LL_BOOLEAN
        Case adChar: TyStr = LL_TEXT
        Case adVarChar: TyStr = LL_TEXT
        Case adCurrency: TyStr = LL_NUMERIC
        Case adDate: TyStr = LL_TEXT
        Case adDecimal: TyStr = LL_NUMERIC
        Case adDouble: TyStr = LL_NUMERIC
        Case adInteger: TyStr = LL_NUMERIC
        Case adNumeric: TyStr = LL_NUMERIC
        Case adSingle: TyStr = LL_NUMERIC
        Case adSmallInt: TyStr = LL_NUMERIC
        Case Else: TyStr = LL_TEXT
        End Select
        If FL1.Value <> vbNullString Then
            If LCase(FL1.Name) = "analog" Then
                If GlAnl = True Then
                    RetWe = LlDefineFieldExt(mHwnd, FL1.Name, FL1.Value, TyStr, vbNullString)
                Else
                    RetWe = LlDefineFieldExt(mHwnd, FL1.Name, 0, TyStr, vbNullString)
                End If
            Else
                RetWe = LlDefineFieldExt(mHwnd, FL1.Name, FL1.Value, TyStr, vbNullString)
            End If
        Else
            If TyStr = LL_NUMERIC Then
                RetWe = LlDefineFieldExt(mHwnd, FL1.Name, 0, TyStr, vbNullString)
            Else
                RetWe = LlDefineFieldExt(mHwnd, FL1.Name, "", TyStr, vbNullString)
            End If
        End If
    End If
Next FL1

If Grupp > 0 Then
    RetWe = LlDefineFieldExt(mHwnd, "Grupp", Grupp, LL_NUMERIC, vbNullString)
End If

If DauDg <> vbNullString Then
    RetWe = LlDefineFieldExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLDeF " & Err.Number
Resume Next

End Sub
Private Sub LLDeV(ByRef RST As ADODB.Recordset, Optional ByVal DauDg As String, Optional ByVal PrGen As Boolean = False, Optional ByVal TaStr As String, Optional ByVal EmSnd As Boolean = False)
On Error GoTo DeErr
'Definiert die Formularvariablen für ListLabel

Dim PatNr As Long
Dim MaNum As Long
Dim MiNum As Long
Dim ArNum As Long
Dim PaNum As Long
Dim RetWe As Long
Dim TmGui As String
Dim TyStr As String
Dim AnzPo As Integer
Dim AktPo As Integer
Dim TheVor As Boolean

LlDefineVariableStart mHwnd

For Each FL1 In RST.Fields
    If FL1.Type <> 128 Then
        Select Case FL1.Type
        Case adBigInt: TyStr = LL_NUMERIC
        Case adBoolean: TyStr = LL_BOOLEAN
        Case adChar: TyStr = LL_TEXT
        Case adVarChar: TyStr = LL_TEXT
        Case adCurrency: TyStr = LL_NUMERIC
        Case adDate: TyStr = LL_TEXT
        Case adDecimal: TyStr = LL_NUMERIC
        Case adDouble: TyStr = LL_NUMERIC
        Case adInteger: TyStr = LL_NUMERIC
        Case adNumeric: TyStr = LL_NUMERIC
        Case adSingle: TyStr = LL_NUMERIC
        Case adSmallInt: TyStr = LL_NUMERIC
        Case Else: TyStr = LL_TEXT_ALLOW_WORDWRAP
        End Select
        If FL1.Value <> vbNullString Then
            If FL1.Name = "Diagnose" Then
                Select Case ForNa
                Case "KarEt2": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "KraBla": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "EMRAbr": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "Rechnu": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechKo": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechLa": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "Auslan": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechKu": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechAl": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "Besche": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechFa": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechSt": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "RechSk": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "Angebo": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "Gutsch": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "StAnge": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "StGuts": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "AnBeBo": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "AnFrBo": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                Case "AnFxBo": RetWe = LlDefineVariableExt(mHwnd, "Diagnose", FL1.Value & vbCrLf & DauDg, LL_TEXT, vbNullString)
                End Select
            Else
                RetWe = LlDefineVariableExt(mHwnd, FL1.Name, FL1.Value, TyStr, vbNullString)
            End If
        Else
            If FL1.Name = "Diagnose" Then
                If DauDg <> vbNullString Then
                    Select Case ForNa
                    Case "KarEt2": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "KraBla": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "EMRAbr": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "Rechnu": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechKo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechLa": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "Auslan": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechKu": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechAl": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "Besche": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechFa": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechSt": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "RechSk": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "Angebo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "Gutsch": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "StAnge": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "StGuts": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "AnBeBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "AnFrBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    Case "AnFxBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, DauDg, LL_TEXT, vbNullString)
                    End Select
                Else
                    Select Case ForNa
                    Case "KarEt2": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "KraBla": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "EMRAbr": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "Rechnu": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechKo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechLa": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "Auslan": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechKu": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechAl": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "Besche": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechFa": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechSt": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "RechSk": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "Angebo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "Gutsch": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "StAnge": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "StGuts": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "AnBeBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "AnFrBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    Case "AnFxBo": RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, LL_TEXT, vbNullString)
                    End Select
                End If
            End If
            If FL1.Name <> "Diagnose" Then
                If TyStr = LL_NUMERIC Then
                    RetWe = LlDefineVariableExt(mHwnd, FL1.Name, 0, TyStr, vbNullString)
                Else
                    RetWe = LlDefineVariableExt(mHwnd, FL1.Name, vbNullString, TyStr, vbNullString)
                End If
            End If
        End If
    End If
Next FL1

Select Case ForNa
Case "EMRAbr": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "Rechnu": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "RechKo": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "RechLa": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "Auslan": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "RechKu": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "RechAl": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "RechSt": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "Angebo": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
Case "Gutsch": RetWe = LlDefineVariableExt(mHwnd, "Matrixcode", TaStr, LL_BARCODE_DATAMATRIX, vbNullString)
End Select

If ManVo = True Then 'Mandanten vorhanden
    Select Case ForNa
    Case "LabKom":
        If RST.Fields("IP0").Value <> vbNullString Then
            If RST.Fields("IP0").Value > 0 Then
                MaNum = RST.Fields("IP0").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        MiNum = GlMiA(GlSmI, 2)
    Case "LabBef":
        If RST.Fields("IP0").Value <> vbNullString Then
            If RST.Fields("IP0").Value > 0 Then
                MaNum = RST.Fields("IP0").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        MiNum = GlMiA(GlSmI, 2)
    Case "LabExp":
        If RST.Fields("IP0").Value <> vbNullString Then
            If RST.Fields("IP0").Value > 0 Then
                MaNum = RST.Fields("IP0").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        MiNum = GlMiA(GlSmI, 2)
    Case "LabAuf":
        If RST.Fields("IP0").Value <> vbNullString Then
            If RST.Fields("IP0").Value > 0 Then
                MaNum = RST.Fields("IP0").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        MiNum = GlMiA(GlSmI, 2)
    Case "EiMahn": 'Einzelmahnung
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDT").Value <> vbNullString Then
            If RST.Fields("IDT").Value > 0 Then
                MiNum = RST.Fields("IDT").Value
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "SaMahn": 'Sammelmahnung
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDT").Value <> vbNullString Then
            If RST.Fields("IDT").Value > 0 Then
                MiNum = RST.Fields("IDT").Value
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "StMahn": 'Einzelmahnung (Alternativ)
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDT").Value <> vbNullString Then
            If RST.Fields("IDT").Value > 0 Then
                MiNum = RST.Fields("IDT").Value
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "AdrEti":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "PhoEti":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "VerEti":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "KarEt1":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "KarEt2":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "DiaEti":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "KraBla":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "AnFrBo":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "ReUber":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "AnBeBo":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "AnFxBo":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "TerPat": 'Terminzettel
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("PatNr").Value <> vbNullString Then
                PatNr = RST.Fields("PatNr").Value
            End If
        End If
    Case "TerSer":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("PatNr").Value <> vbNullString Then
                PatNr = RST.Fields("PatNr").Value
            End If
        End If
    Case "Rechnu":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechKo":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechLa":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechKu":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechAl":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechFa":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "RechSt": 'Standardrechnung (Alternativ)
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "Angebo":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case "Gutsch":
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If RST.Fields("IDO").Value <> vbNullString Then
            If RST.Fields("IDO").Value > 0 Then
                ArNum = RST.Fields("IDO").Value
            Else
                ArNum = 0
            End If
        Else
            ArNum = 0
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    Case Else:
        If RST.Fields("IDP").Value <> vbNullString Then
            If RST.Fields("IDP").Value > 0 Then
                MaNum = RST.Fields("IDP").Value 'Fast immer [IDP]
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
        If RST.Fields("IDM").Value <> vbNullString Then
            If RST.Fields("IDM").Value > 0 Then
                MiNum = RST.Fields("IDM").Value 'Fast immer [IDM]
            Else
                MiNum = GlMiA(GlSmI, 2)
            End If
        Else
            MiNum = GlMiA(GlSmI, 2)
        End If
        If PrGen = True Then
            If RST.Fields("ID0").Value <> vbNullString Then
                PatNr = RST.Fields("ID0").Value
            End If
        End If
    End Select
    If MaNum > 0 Then
        AnzPo = UBound(GlThe)
        For AktPo = 1 To AnzPo
            If GlThe(AktPo, 0) = MaNum Then
                RetWe = LlDefineVariableExt(mHwnd, "AnwVorname", GlThe(AktPo, 1), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwName", GlThe(AktPo, 2), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwStraße", GlThe(AktPo, 3), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwPLZ", GlThe(AktPo, 4), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwOrt", GlThe(AktPo, 5), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwTelefon", GlThe(AktPo, 6), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwTelefax", GlThe(AktPo, 7), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBank", GlThe(AktPo, 8), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ", GlThe(AktPo, 9), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwKonto", GlThe(AktPo, 10), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwSteuer", GlThe(AktPo, 11), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBeruf", GlThe(AktPo, 12), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwTitel", GlThe(AktPo, 14), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwIKNr", GlThe(AktPo, 15), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwLANR", GlThe(AktPo, 15), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwEmail", GlThe(AktPo, 16), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwInternet", GlThe(AktPo, 17), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN", GlThe(AktPo, 18), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwPraxis", GlThe(AktPo, 19), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBank2", GlThe(AktPo, 20), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ2", GlThe(AktPo, 21), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwKonto2", GlThe(AktPo, 22), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN2", GlThe(AktPo, 23), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwSprechzeiten", GlThe(AktPo, 24), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBIC", GlThe(AktPo, 31), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBIC2", GlThe(AktPo, 32), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwGID", GlThe(AktPo, 33), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwAbsender", GlThe(AktPo, 36), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwGirocode", GlThe(AktPo, 10), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwGLN", GlThe(AktPo, 40), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwZSR", GlThe(AktPo, 41), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwBunLan", GlThe(AktPo, 42), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwKVBezirk", GlThe(AktPo, 43), LL_TEXT, vbNullString)
                RetWe = LlDefineVariableExt(mHwnd, "AnwKanton", GlThe(AktPo, 44), LL_TEXT, vbNullString)
                TheVor = True
                Exit For
            End If
        Next AktPo
        If TheVor = False Then
            LLReg
        End If
    Else
        LLReg
    End If
Else
    LLReg
End If

If GlMiV = True Then 'Mitarbeiter vorhanden
    If MitVo = True Then 'Mitarbeiter vorhanden
        If MiNum > 0 Then
            AnzPo = UBound(GlMiK)
            For AktPo = 1 To AnzPo
                If GlMiK(AktPo, 2) = MiNum Then
                    RetWe = LlDefineVariableExt(mHwnd, "MitNr", GlMiK(AktPo, 2), LL_NUMERIC, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitName", GlMiK(AktPo, 1), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitNachname", GlMiK(AktPo, 3), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitVorname", GlMiK(AktPo, 4), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitEmail", GlMiK(AktPo, 22), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", GlMiK(AktPo, 29), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", GlMiK(AktPo, 30), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitGLN", GlMiK(AktPo, 31), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitZSR", GlMiK(AktPo, 32), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", GlMiK(AktPo, 33), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitStraße", GlMiK(AktPo, 34), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", GlMiK(AktPo, 35), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "MitOrt", GlMiK(AktPo, 36), LL_TEXT, vbNullString)
                    Exit For
                End If
            Next AktPo
        Else
            RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "MitNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitAbsender", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitStraße", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitPLZ", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "MitOrt", vbNullString, LL_TEXT, vbNullString)
End If

If GlArV = True Then 'Verordner vorhanden
    If ArzVo = True Then 'Verordner vorhanden
        If ArNum > 1 Then
            AnzPo = UBound(GlArz)
            For AktPo = 1 To AnzPo
                If GlArz(AktPo, 0) = ArNum Then
                    RetWe = LlDefineVariableExt(mHwnd, "ArzNr", GlArz(AktPo, 0), LL_NUMERIC, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", GlArz(AktPo, 1), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", GlArz(AktPo, 2), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", GlArz(AktPo, 6), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", GlArz(AktPo, 7), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzName", GlArz(AktPo, 8), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", GlArz(AktPo, 9), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", GlArz(AktPo, 10), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", GlArz(AktPo, 11), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", GlArz(AktPo, 12), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", GlArz(AktPo, 12), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", GlArz(AktPo, 13), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", GlArz(AktPo, 14), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", GlArz(AktPo, 15), LL_TEXT, vbNullString)
                    RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", GlArz(AktPo, 16), LL_TEXT, vbNullString)
                    Exit For
                End If
            Next AktPo
        Else
            RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
            RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
        End If
    Else
        RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
        RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
    End If
Else
    RetWe = LlDefineVariableExt(mHwnd, "ArzNr", 0, LL_NUMERIC, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzVorname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzNachname", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefon", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTelefax", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzName", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzTitel", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzEmail", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBeruf", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzLANR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzGLN", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzZSR", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzBunLan", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKVBez", vbNullString, LL_TEXT, vbNullString)
    RetWe = LlDefineVariableExt(mHwnd, "ArzKanton", vbNullString, LL_TEXT, vbNullString)
End If

Select Case ForNa
Case "EiMahn": RetWe = LlDefineVariableExt(mHwnd, "Mahngebühr", MaGeb, LL_NUMERIC, vbNullString)
Case "SaMahn": RetWe = LlDefineVariableExt(mHwnd, "Mahngebühr", MaGeb, LL_NUMERIC, vbNullString)
Case "StMahn": RetWe = LlDefineVariableExt(mHwnd, "Mahngebühr", MaGeb, LL_NUMERIC, vbNullString)
Case "EMRAbr": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "Rechnu": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechKo": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechLa": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "Auslan": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechKu": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechAl": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "Besche": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechFa": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechSt": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "RechSk": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "Angebo": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "Gutsch": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "StAnge": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "StGuts": RetWe = LlDefineVariableExt(mHwnd, "ICDDauerdiagnose", DauDg, LL_TEXT, vbNullString)
Case "AnBeBo": RetWe = LlDefineVariableExt(mHwnd, "Zusammenfassung", DauDg, LL_TEXT, vbNullString)
Case "AnFrBo": RetWe = LlDefineVariableExt(mHwnd, "Zusammenfassung", DauDg, LL_TEXT, vbNullString)
Case "AnFxBo": RetWe = LlDefineVariableExt(mHwnd, "Zusammenfassung", DauDg, LL_TEXT, vbNullString)
End Select

If KopTe <> vbNullString Then
    RetWe = LlDefineVariableExt(mHwnd, "Kopftext", KopTe, LL_TEXT, vbNullString)
End If
If EmSen > 0 Then
    RetWe = LlDefineVariableExt(mHwnd, "Emailversand", True, LL_BOOLEAN, vbNullString)
Else
    RetWe = LlDefineVariableExt(mHwnd, "Emailversand", EmSnd, LL_BOOLEAN, vbNullString)
End If

RetWe = LlDefineVariableExt(mHwnd, "Lizenziert", CInt(GlLiz), LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Druckkopie", DrKop, LL_BOOLEAN, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Katalog", GlFri, LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Bilderordner", GlBPf, LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Dokumentenordner", GlDox, LL_TEXT, vbNullString)

If GlTSN <> vbNullString Then
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", GlTSN, LL_TEXT, vbNullString)
Else
    RetWe = LlDefineVariableExt(mHwnd, "TSESerien", "1234", LL_TEXT, vbNullString)
End If

If PrGen = True Then 'soll Kopie als PDF erstellt und ins Krankenblatt hinzugefpgt werden?
    TmGui = CreateID("P")
    GlNeK = GlKoX
    With GlNeK
        .PatNr = PatNr
        .IdxNr = 0
        .EiDat = Format$(Date, "dd.mm.yyyy")
        .EiZei = TimeValue(Now)
        .EiTyp = 105 'Bilddatei
        .KoStr = DaNam
        .KoGui = TmGui
        .TeStr = DaNam
        .NeuEi = False
        .Mitar = MiNum
    End With
    K_Einf
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLDeV " & Err.Number
Resume Next

End Sub
Public Function LLDrIn() As String
On Error GoTo DeErr
'Gibt den Druckernamern des Projektes aus

Dim RetWe As Long
Dim DaExt As String
Dim DruNa As String * 32 'Druckername
Dim DevMo As DEVMODEA

DaExt = Right$(FiNam, 3)

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt

Select Case LCase(DaExt):
Case "blg": RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DruNa, 32, DevMo, 32)
Case "lbl": RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LABEL, FiNam, -1, DruNa, 32, DevMo, 32)
Case "lst": RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DruNa, 32, DevMo, 32)
Case "crd": RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_CARD, FiNam, -1, DruNa, 32, DevMo, 32)
Case "asw": RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DruNa, 32, DevMo, 32)
End Select

LlJobClose mHwnd

LLDrIn = DruNa

Exit Function

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLDrIn " & Err.Number
Resume Next

End Function
Public Sub LLEdit()
On Error GoTo DeErr
'Öffnet den Formulardesigner

Dim RetWe As Long
Dim DauDg As String
Dim DaExt As String
Dim MxCod As String
Dim IdStr As Variant
Dim StaEd As Boolean
Dim KeiPa As Boolean 'Keine Paramater
Dim SQL1, SQL2, SQL3 As String
Dim Mld1, Mld2, Mld3, Tit1 As String
Dim Frage As Integer

Set FM = frmFormular

Set clFil = New clsFile

Mld1 = "Um den Formulardesigner nutzen zu können, müssen Beispieldaten vorhanden sein. Es sind noch keine entsprechenden Daten in Ihrem Programm vorhanden."
Mld2 = "Es ist keine oder eine falsche Formularvorlage zugeordnet. Möchten Sie eine neue Formularvorlage erstellen?"
Mld3 = "Bitte wählen Sie einen Namen für eine Formularvorlage"
Tit1 = "Formulardesigner"
DaExt = Right$(FiNam, 3)

Select Case ForNa
Case "DiaEti": 'Diagnoseetikett
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "KarEt1": 'Krankenblattetikett1
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "KarEt2": 'Krankenaktenausdruck
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "PhoEti": 'Kontaktdaten-Etikett
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "VerEti": 'Versicherungsetikett
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "StaEt2": 'Spezial-Etikett 2
        SQL1 = "qryPrSimAd3"
        SQL2 = "qryPrSimAd3"
        SQL3 = "qryPrSimAd3"
        KeiPa = True
Case "EMRAbr": 'EMR Abrechnungsformular
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2a"
        SQL3 = "qryPrSimRe3"
Case "Rechnu": 'Standardrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "Auslan": 'Auslandsrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "RechKo": 'Kompaktrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe5"
        SQL3 = "qryPrSimRe3"
Case "RechLa": 'Laborrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "RechKu": 'Kurzrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe5"
        SQL3 = "qryPrSimRe3"
Case "RechAl": 'Steuermixrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
        SQL3 = "qryPrSimRe3"
Case "Besche": 'Steuermixrechnung (Alternativ)
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
        SQL3 = "qryPrSimRe3"
Case "RechFa": 'Faxrechnung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe5"
        SQL3 = "qryPrSimRe3"
Case "RechSt": 'Standardrechnung (Alternativ)
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "StAnge": 'Stampit Angebot
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "StGuts": 'Gutschrift (Alternativ)
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "Gutsch": 'Gutschrift
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "RechSk": 'Kompaktrechnung (Alternativ)
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe5"
        SQL3 = "qryPrSimRe3"
Case "KraBla": 'Krankenblatt
        SQL1 = "qryPrSimKr1"
        SQL2 = "qryPrSimKr2"
        SQL3 = "qryPrSimKr3"
Case "AnBeBo": 'Anamneseberichtbogen
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn2"
        SQL3 = "qryPrAn3"
Case "AnFrBo": 'Anamnesefragebogen
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
        SQL3 = "qryPrAn3"
Case "AnFxBo": 'Anamnesefaxfragebogen
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
        SQL3 = "qryPrAn3"
Case "TagPro": 'Tagesprotokoll
        SQL1 = "qrySimAbKra07"
        SQL2 = "qrySimAbKra07"
        SQL3 = "qrySimAbKra07"
        Grupp = 1
        KeiPa = True
Case "Angebo": 'Kostenvoranschlag
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "EiMahn": 'Einzelmahnung
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
        SQL3 = "qryPrSimOP3"
Case "SaMahn": 'Sammelmahnung
        SQL1 = "qryPrSimOP4"
        SQL2 = "qryPrSimOP5"
        SQL3 = "qryPrSimOP6"
Case "StMahn": 'Einzelmahnung (Alternativ)
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
        SQL3 = "qryPrSimOP3"
Case "Uberwe": 'Überweisung
        SQL1 = "qryPrSimRe1"
        SQL3 = "qryPrSimRe3"
Case "UberEU": 'Überweisung_EU
        SQL1 = "qryPrSimRe1"
        SQL3 = "qryPrSimRe3"
Case "Liefer": 'Quittung
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        SQL3 = "qryPrSimRe3"
Case "LabAuf": 'Laborauftrag
        SQL1 = "qryPrLaAu1"
        SQL2 = "qryPrLaAu2"
        SQL3 = "qryPrLaAu3"
Case "LabKom": 'Komplettbericht
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
        SQL3 = "qryPrLaBe3"
Case "LabBef": 'Laborbefund
        SQL1 = "qryPrLaBe1"
        SQL3 = "qryPrLaBe3"
Case "LabFax": 'Laborfax
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
        SQL3 = "qryPrLaBe3"
Case "LabExp": 'Exportbericht
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
        SQL3 = "qryPrLaBe3"
Case "BuJour": 'Buchungsjournal
        SQL2 = "qryPrBuJo3"
        SQL3 = "qryPrBuJo3"
        KopTe = "Zeitraum"
        AnSal = "1"
        KeiPa = True
Case "BuKass": 'Buchführung
        SQL2 = "qryPrBuJo3"
        SQL3 = "qryPrBuJo3"
        KopTe = "Zeitraum"
        AnSal = "1"
        KeiPa = True
Case "EiUbMo": 'Einnahmenüberschussrechnung
        SQL2 = "qryPrBuEr3"
        SQL3 = "qryPrBuEr1"
        KopTe = "Zeitraum"
        KeiPa = True
Case "EiUbQu": 'Einnahmenüberschussrechnung
        SQL2 = "qryPrBuEr3"
        SQL3 = "qryPrBuEr4"
        KopTe = "Zeitraum"
        KeiPa = True
Case "EiUbJa": 'Einnahmenüberschussrechnung
        SQL2 = "qryPrBuEr3"
        SQL3 = "qryPrBuEr2"
        KopTe = "Zeitraum"
        KeiPa = True
Case "EiUbGr": 'Einnahmenüberschussrechnung (Gruppiert)
        SQL2 = "qryPrBuEr3"
        SQL3 = "qryPrBuEr1"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaKoMo": 'Saldenliste der Sachkonten
        SQL2 = "qryPrBuKo1"
        SQL3 = "qryPrBuKo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaKoQu": 'Saldenliste der Sachkonten
        SQL2 = "qryPrBuKo4"
        SQL3 = "qryPrBuKo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaKoJa": 'Saldenliste der Sachkonten
        SQL2 = "qryPrBuKo2"
        SQL3 = "qryPrBuKo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaGeMo": 'Saldenliste der Geldkonten
        SQL2 = "qryPrBuBa1"
        SQL3 = "qryPrBuBa3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaGeQu": 'Saldenliste der Geldkonten
        SQL2 = "qryPrBuBa4"
        SQL3 = "qryPrBuBa3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "SaGeJa": 'Saldenliste der Geldkonten
        SQL2 = "qryPrBuBa2"
        SQL3 = "qryPrBuBa3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "KonAus": 'Einzelkontenauswertung
        SQL2 = "qryPrBuJo1"
        SQL3 = "qryPrBuJo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "KonGru": 'Einzelkontenauswertung (Gruppiert Sachkonten)
        SQL2 = "qryPrBuJo2"
        SQL3 = "qryPrBuJo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "KonGel": 'Einzelkontenauswertung (Gruppiert Geldkonten)
        SQL2 = "qryPrBuJo4"
        SQL3 = "qryPrBuJo3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "ReUmsa": 'Rechnungsumsatz
        SQL2 = "qryPrReUm3"
        SQL3 = "qryPrReUm3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "KomAus": 'Kompaktauswertung
        SQL2 = "qryPrBuKom"
        SQL3 = "qryPrBuKom"
        KopTe = "Zeitraum"
        KeiPa = True
Case "UmStMo": 'Umsatzsteuerübersicht
        SQL2 = "qryPrBuSt1"
        SQL3 = "qryPrBuSt3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "UmStQu": 'Umsatzsteuerübersicht
        SQL2 = "qryPrBuSt4"
        SQL3 = "qryPrBuSt3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "UmStJa": 'Umsatzsteuerübersicht
        SQL2 = "qryPrBuSt2"
        SQL3 = "qryPrBuSt3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "PatUms": 'Patientenumsatzliste
        SQL2 = "qryPrSimTaPa3"
        SQL3 = "qryPrSimTaPa3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "TagUms": 'Tagesumsatzliste
        SQL2 = "qryPrSimTag1"
        SQL3 = "qryPrSimTag1"
        KopTe = "Zeitraum"
        KeiPa = True
Case "LabUms": 'Steuermixumsatzliste
        SQL2 = "qryPrSimTag5"
        SQL3 = "qryPrSimTag5"
        KopTe = "Zeitraum"
        KeiPa = True
Case "TagSum": 'Tagessummen
        SQL2 = "qryPrSimTag2"
        SQL3 = "qryPrSimTag2"
        KopTe = "Zeitraum"
        KeiPa = True
Case "RecSum": 'Rechnungssummen
        SQL2 = "qryPrSimTag3"
        SQL3 = "qryPrSimTag3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "StaDia": 'Diagnosestatistik
        SQL2 = "qryPrSimDia1b"
        SQL3 = "qryPrSimDia1b"
        KopTe = "Zeitraum"
        KeiPa = True
Case "StaGeb": 'Gebührenstatistik
        SQL2 = "qryPrSimGeb1"
        SQL3 = "qryPrSimGeb1"
        KopTe = "Zeitraum"
        KeiPa = True
Case "BuKont": 'Buchungskonten
        SQL2 = "qryPrKon3"
        SQL3 = "qryPrKon3"
        KeiPa = True
Case "ReUber": 'Patientensammelrechnung
        SQL1 = "qryPrSimReL3"
        SQL2 = "qryPrSimReL3"
        SQL3 = "qryPrSimReL3"
        KopTe = "Zeitraum"
        KeiPa = True
Case "AbrLis": 'Abrechnungsliste
        SQL2 = "qryPrAbrLi3"
        SQL3 = "qryPrAbrLi3"
        KeiPa = True
Case "BuList": 'Buchungsliste
        SQL2 = "qryPrBuLis3"
        SQL3 = "qryPrBuLis3"
        KeiPa = True
Case "ReList": 'Rechnungsexportliste
        SQL2 = "qryPrSimRe1"
        SQL3 = "qryPrSimRe3"
Case "ResUbe": 'Rechnugsschnellübersicht
        SQL2 = "qryPrSimRe9"
        SQL3 = "qryPrSimRe3"
        KeiPa = True
        DauDg = "ICDDauerdiagnose"
Case "TerExp": 'Terminliste
        SQL2 = "qryPrTerLis3"
        SQL3 = "qryPrTerLis3"
        KeiPa = True
Case "KatLis":
        SQL2 = "qryPrKata"
        SQL3 = "qryPrKata"
        KeiPa = True
Case "KetLis":
        SQL2 = "qryPrKette"
        SQL3 = "qryPrKette"
        KeiPa = True
Case "InvLis":
        SQL2 = "qryKat12C"
        SQL3 = "qryKat12C"
        KeiPa = True
Case "BesLis":
        SQL2 = "qryKat12C"
        SQL3 = "qryKat12C"
        KeiPa = True
Case "TerPat": 'Terminzettel
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm1"
        SQL3 = "qryPrTerm3"
Case "TerSer":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm1"
        SQL3 = "qryPrTerm3"
Case "TeUbNo":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "TeUbEu":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "QuiTer":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "QuiTem":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "TeUbAN":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "TeUbAE":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm5"
        SQL3 = "qryPrTerm3"
Case "TeEiMa":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm1"
        SQL3 = "qryPrTerm3"
Case "TeEiSa":
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm1"
        SQL3 = "qryPrTerm3"
Case "KasPro":
        SQL1 = "qryPrBuZal1"
        SQL2 = "qryPrBuZal1"
        SQL3 = "qryPrBuZal3"
        KopTe = "Kassenname"
End Select

If SQL3 = vbNullString Then
    WindowMess Mld1, Dial4, Tit1, FM.hwnd
    Exit Sub
End If

Set RS30 = New ADODB.Recordset
With RS30
    .CursorLocation = adUseClient
    .Source = SQL3
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdTableDirect
End With

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt

RetWe = LlDesignerProhibitAction(mHwnd, LL_SYSCOMMAND_MINIMIZE)
If GlRDP = True Then
    'RetWe = LlDesignerProhibitAction(mHwnd, 524) 'Datei > Neu
    'RetWe = LlDesignerProhibitAction(mHwnd, 726) 'Datei > Importieren
    'RetWe = LlDesignerProhibitAction(mHwnd, 509) 'Datei > Drucken/Exportieren
    'RetWe = LlDesignerProhibitAction(mHwnd, 519) 'Datei > Öffnen
    'RetWe = LlDesignerProhibitAction(mHwnd, 515) 'Datei > Speichern unter
    'RetWe = LlDesignerProhibitAction(mHwnd, 561) 'Objekte > Einfügen > Formularvorlage
    'RetWe = LlDesignerProhibitAction(mHwnd, 527) 'Objekte > Einfügen > Bild
    'RetWe = LlDesignerProhibitAction(mHwnd, 799)
End If

If GlLiz = False Then
    RetWe = LlDesignerProhibitEditingObject(mHwnd, "Summenzeile")
End If

If RS30.RecordCount > 0 Then
    'WICHTIG! Immer das erste Feld ist der Parameter
    If RS30.Fields(0).Type = adBigInt Then
        IdStr = CLng(RS30.Fields(0).Value)
    Else
        IdStr = RS30.Fields(0).Value
    End If

    Select Case ForNa
    Case "EMRAbr":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "Rechnu":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "RechKo":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "RechLa":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "Auslan":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "RechKu":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "RechAl":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "RechSt":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "Angebo":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    Case "Gutsch":
            DauDg = S_DrDi(IdStr)
            MxCod = S_ReMx(IdStr)
    End Select

    Select Case LCase(DaExt):
    Case "blg": 'Variablen und Listen
        RetWe = LlSetFileExtensions(mHwnd, LL_PROJECT_LIST, "blg", "lsp", "lsv")

        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        If KeiPa = True Then
            Set RS31 = DBCmRe0(SQL1)
        Else
            Set RS31 = DBCmRe1(SQL1, "@IdStr", IdStr)
        End If
        Set RS32 = New ADODB.Recordset
        RS32.CursorLocation = adUseClient
        If KeiPa = True Then
            Set RS32 = DBCmRe0(SQL2)
        Else
            Set RS32 = DBCmRe1(SQL2, "@IdStr", IdStr)
        End If
        If RS31.RecordCount > 0 Then
            If RS32.RecordCount > 0 Then
                StaEd = True
            End If
        End If
    Case "lbl": 'Nur Variablen
        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        If KeiPa = True Then
            Set RS31 = DBCmRe0(SQL1)
        Else
            Set RS31 = DBCmRe1(SQL1, "@IdStr", IdStr)
        End If
        If RS31.RecordCount > 0 Then
            StaEd = True
        End If
    Case "lst": 'Nur Listen
        Set RS32 = New ADODB.Recordset
        RS32.CursorLocation = adUseClient
        If KeiPa = True Then
            Set RS32 = DBCmRe0(SQL2)
        Else
            Set RS32 = DBCmRe1(SQL2, "@IdStr", IdStr)
        End If
        If RS32.RecordCount > 0 Then
            StaEd = True
        End If
    Case "crd": 'Nur Variablen
        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        If KeiPa = True Then
            Set RS31 = DBCmRe0(SQL1)
        Else
            Set RS31 = DBCmRe1(SQL1, "@IdStr", IdStr)
        End If
        If RS31.RecordCount > 0 Then
            StaEd = True
        End If
    Case "asw": 'Variablen und Listen
        Set RS31 = New ADODB.Recordset
        With RS31
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdTableDirect
        End With
        Set RS32 = New ADODB.Recordset
        With RS32
            .CursorLocation = adUseClient
            .Source = SQL2
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdTableDirect
        End With
        If RS31.RecordCount > 0 Then
            If RS32.RecordCount > 0 Then
                StaEd = True
            End If
        End If
    End Select
        
    Select Case ForNa
    Case "RechFa": RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", "01234", 0)
    Case "LabFax": RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", "01234", 0)
    Case "AnFxBo": RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", "01234", 0)
    End Select

    If StaEd = True Then
        If clFil.FilVor(FiNam) = True Then
            Select Case LCase(DaExt):
            Case "blg": 'Variablen und Listen
                    If DauDg <> vbNullString Then
                        If MxCod <> vbNullString Then
                            LLDeV RS31, DauDg, False, MxCod
                        Else
                            LLDeV RS31, DauDg
                        End If
                    Else
                        LLDeV RS31
                    End If
                    LLDeF RS32
                    RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
            Case "lbl": 'Nur Variablen
                    LLDeV RS31
                    RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LABEL, FiNam)
            Case "lst": 'Nur Listen
                    LLDeB
                    LLDeF RS32, DauDg
                    RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
            Case "crd": 'Nur Variablen
                    LLDeV RS31
                    RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_CARD, FiNam)
            Case "asw": 'Variablen und Listen
                    LLDeV RS31
                    LLDeF RS32
                    RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
            End Select
        Else
            Frage = WindowMess(Mld2, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                Select Case LCase(DaExt):
                Case "blg": 'Variablen und Listen
                        RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LIST Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                        If RetWe = 0 Then
                            LLDeV RS31
                            LLDeF RS32
                            RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
                        End If
                Case "lbl": 'Nur Variablen
                        RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LABEL Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                        If RetWe = 0 Then
                            LLDeV RS31
                            RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LABEL, FiNam)
                        End If
                Case "lst": 'Nur Listen
                        RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LIST Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                        If RetWe = 0 Then
                            LLDeF RS32
                            RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
                        End If
                Case "crd": 'Nur Variablen
                        RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_CARD Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                        If RetWe = 0 Then
                            LLDeV RS31
                            RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_CARD, FiNam)
                        End If
                Case "asw": 'Variablen und Listen
                        RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LIST Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                        If RetWe = 0 Then
                            LLDeV RS31
                            LLDeF RS32
                            RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
                        End If
                End Select
            End If
        End If
    Else
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
Else
    RS30.Close
    Set RS30 = Nothing
    DoEvents
    LlJobClose mHwnd
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
    Exit Sub
End If

Select Case LCase(DaExt):
Case "blg": 'Variablen und Listen
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
Case "lbl": 'Nur Variablen
    RS31.Close
    Set RS31 = Nothing
Case "lst": 'Nur Listen
    RS32.Close
    Set RS32 = Nothing
Case "crd": 'Nur Variablen
    RS31.Close
    Set RS31 = Nothing
Case "asw": 'Variablen und Listen
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
End Select

RS30.Close
Set RS30 = Nothing

DoEvents
LlJobClose mHwnd

Set clFil = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLEdit " & Err.Number
Resume Next

End Sub
Public Sub LLEdUn(ByVal Flag As Integer)
On Error GoTo DeErr
'Öffnet den Formulardesigner

Dim RetWe As Long
Dim GesZa As Long
Dim DaExt As String
Dim Mld1, Mld2, Mld3, Tit1 As String
Dim Frage As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmFormular
Set RpCo1 = frmMain.repCont1
Set RpCo2 = frmMain.repCont2
Set RpCo3 = frmMain.repCont3
Set RpCo4 = frmMain.repCont4
Set RpCo5 = frmMain.repCont5
Set RpCo6 = frmMain.repCont6
Set RpCo8 = frmMain.repCont8

Set clFil = New clsFile

Select Case Flag
Case 1:
    Set RpRws = RpCo1.Rows
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
Case 2:
    Set RpRws = RpCo2.Rows
    Set RpCls = RpCo2.Columns
    Set RpSel = RpCo2.SelectedRows
Case 3:
    Set RpRws = RpCo3.Rows
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case 4:
    Set RpRws = RpCo4.Rows
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
Case 5:
    Set RpRws = RpCo5.Rows
    Set RpCls = RpCo5.Columns
    Set RpSel = RpCo5.SelectedRows
Case 6:
    Set RpRws = RpCo6.Rows
    Set RpCls = RpCo6.Columns
    Set RpSel = RpCo6.SelectedRows
Case 8:
    Set RpRws = RpCo8.Rows
    Set RpCls = RpCo8.Columns
    Set RpSel = RpCo8.SelectedRows
End Select

GesZa = RpRws.Count
Mld1 = "Um den Formulardesigner nutzen zu können, müssen Beispieldaten vorhanden sein. Es sind noch keine entsprechenden Daten in Ihrem Programm vorhanden."
Mld2 = "Es ist keine oder eine falsche Formularvorlage zugeordnet. Möchten Sie eine neue Formularvorlage erstellen?"
Mld3 = "Bitte wählen Sie einen Namen für eine Formularvorlage"
Tit1 = "Formulardesigner"
DaExt = Right$(FiNam, 3)

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt

RetWe = LlDesignerProhibitAction(mHwnd, LL_SYSCOMMAND_MINIMIZE)
If GlRDP = True Then
    'RetWe = LlDesignerProhibitAction(mHwnd, 524) 'Datei > Neu
    'RetWe = LlDesignerProhibitAction(mHwnd, 726) 'Datei > Importieren
    'RetWe = LlDesignerProhibitAction(mHwnd, 509) 'Datei > Drucken/Exportieren
    'RetWe = LlDesignerProhibitAction(mHwnd, 519) 'Datei > Öffnen
    'RetWe = LlDesignerProhibitAction(mHwnd, 515) 'Datei > Speichern unter
    'RetWe = LlDesignerProhibitAction(mHwnd, 561) 'Objekte > Einfügen > Formularvorlage
    'RetWe = LlDesignerProhibitAction(mHwnd, 527) 'Objekte > Einfügen > Bild
    'RetWe = LlDesignerProhibitAction(mHwnd, 799)
End If

If GlLiz = False Then
    RetWe = LlDesignerProhibitEditingObject(mHwnd, "Summenzeile")
End If

If GesZa > 0 Then
    Set RpRow = RpRws(0)
    If RpRow.GroupRow = True Then
        Set RpRow = RpRws(1)
    End If

    If clFil.FilVor(FiNam) = True Then
        Select Case LCase(DaExt):
        Case "lbl": 'Nur Variablen
                LDeVar RpRow, Flag
                RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LABEL, FiNam)
        Case "lst": 'Nur Listen
                LDeVar RpRow, Flag
                LDeFel RpRow, Flag
                RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
        Case "crd": 'Nur Variablen
                LDeVar RpRow, Flag
                RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_CARD, FiNam)
        Case "blg": 'Beides
                LDeVar RpRow, Flag
                LDeFel RpRow, Flag
                RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
        End Select
    Else
        Frage = WindowMess(Mld2, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            Select Case LCase(DaExt):
            Case "lbl": 'Nur Variablen
                    RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LABEL Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                    If RetWe = 0 Then
                        LDeVar RpRow, Flag
                        RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LABEL, FiNam)
                    End If
            Case "lst": 'Nur Listen
                    RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LIST Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                    If RetWe = 0 Then
                        LDeFel RpRow, Flag
                        RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
                    End If
            Case "crd": 'Nur Variablen
                    RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_CARD Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                    If RetWe = 0 Then
                        LDeVar RpRow, Flag
                        RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_CARD, FiNam)
                    End If
            Case "blg": 'Beides
                    RetWe = LlSelectFileDlgTitleEx(mHwnd, FM.hwnd, Mld3, LL_PROJECT_LIST Or LL_FILE_ALSONEW, FiNam, 100, vbNullString)
                    If RetWe = 0 Then
                        LDeFel RpRow, Flag
                        RetWe = LlDefineLayout(mHwnd, FM.hwnd, "Formulardesigner", LL_PROJECT_LIST, FiNam)
                    End If
            End Select
        End If
    End If
Else
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

DoEvents
LlJobClose mHwnd

Set RpSel = Nothing
Set RpRws = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCo8 = Nothing

Set clFil = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLEdUn " & Err.Number
Resume Next

End Sub
Public Sub LLExBl()
On Error GoTo DeErr
'Exportiert einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim PatNr As Long
Dim StaPr As Boolean
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DaNeu As String 'neuer Dateiname
Dim DauDg As String
Dim MxCod As String
Dim ZipNa As String
Dim ExPfa As String 'Exportpfad
Dim ZpDir As String
Dim ZpPfa As String
Dim XReDa As String
Dim EReDa As String
Dim FiStr As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim ZuFer As Boolean 'ZUGFeRD
Dim ZipDo As Boolean
Dim ZipOk As Boolean
Dim DevMo As DEVMODEA
Dim DatNa() As String
Dim SQL1, SQL2 As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2

Mld1 = "Es existieren keine Daten, die exportiert werden können"
Tit1 = "Keine Daten vorhanden"
AnzPo = UBound(GloDr)
ZipDo = (AnzPo > 1)
If Right$(TmPfa, 1) <> "\" Then TmPfa = TmPfa & "\"
If ZipDo = True Then
    ZpDir = TmPfa & "LLZip_" & CreateID("Z")
    If Dir$(ZpDir, vbDirectory) = vbNullString Then MkDir ZpDir
    ZpDir = ZpDir & "\"
    ExPfa = ZpDir
Else
    ExPfa = StVer
    If ExPfa = vbNullString Then ExPfa = TmPfa
End If
If ExPfa <> vbNullString Then
    If Right$(ExPfa, 1) <> "\" Then ExPfa = ExPfa & "\"
End If

If ExFmt = "PDA" Then
    ExFmt = "PDF"
    ZuFer = True 'ZUGFeRD
End If

ReDim DatNa(AnzPo)

Select Case ForNa
Case "EMRAbr":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2a"
Case "Rechnu":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Auslan":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "RechLa":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "RechAl":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
Case "Besche":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
Case "RechSt":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Angebo":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Gutsch":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "LabKom":
        VerMo = True
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
Case "LabExp":
        VerMo = True
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
Case "LabAuf":
        VerMo = True
        SQL1 = "qryPrLaAu1"
        SQL2 = "qryPrLaAu2"
End Select

If SQL1 = vbNullString Then
    WindowMess Mld1, Dial4, Tit1, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", ExPfa)

For AktPo = 1 To AnzPo
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = vbNullString
    MxCod = vbNullString
    DauDg = S_DrDi(GloDr(AktPo))
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        MxCod = S_ReMx(GloDr(AktPo))
    End If

    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))

    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If

    If StaPr = True Then
        If VerMo = True Then
            If RS31.Fields("PatNr").Value > 0 Then
                PatNr = RS31.Fields("PatNr").Value
            Else
                PatNr = 0
            End If
            If RS31.Fields("Em_Aut").Value <> vbNullString Then
                VerDo = CBool(RS31.Fields("Em_Aut").Value)
            Else
                VerDo = False
            End If
            If VerDo = True Then 'Dokument verschlüsseln
                If RS31.Fields("Em_Pass").Value <> vbNullString Then
                    VerPa = RS31.Fields("Em_Pass").Value
                Else
                    VerDo = False
                End If
            Else
                VerDo = False
            End If
        End If

        If AnzPo > 1 Then
            If Left$(ForNa, 3) = "Lab" Then
                DaNeu = "Angebot_Beleg_" & RS31.Fields("Auftragsnummer").Value & "." & LCase(ExFmt)
            ElseIf Left$(ForNa, 3) = "Rec" Then
                DaNeu = "Rechnung_Beleg_" & RS31.Fields("Rechnungsnummer").Value & "." & LCase(ExFmt)
            End If
        Else
            If LCase(Right$(DaNam, 3)) = "zip" Then
                DaNam = Mid$(DaNam, 1, Len(DaNam) - 3) & LCase(ExFmt)
            End If
            DaNeu = DaNam
        End If

        DatNa(AktPo) = DaNeu
        '    If ZuFer = True Then 'ZUGFeRD
        '        EReDa = "E" & DaNeu
        '    Else
        
        AktZa = 1
        GesZa = RS32.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa

        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            LLDeV RS31, DauDg, False, MxCod, True
        Else
            LLDeV RS31, DauDg, , , True
        End If
        LLDeF RS32
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If

        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNeu)
        Select Case ExFmt
        Case "TXT":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
        Case "HTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "MHTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "XLS":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", ForNa)
        Case "RTF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
        Case "PDF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
            If VerDo = True Then 'Dokument verschlüsseln
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", VerPa)
            End If
        End Select

        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)

        If RetWe = 0 Then
            If AktPo = 1 Then
                FS.Show
                DoEvents
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                While Not RS32.EOF
                    LLDeF RS32
                    PrBr2.Value = AktZa
                    DoEvents
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
    
    DoEvents
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Unload FS
DoEvents

If DruOk = True Then
    If ZuFer = True Then 'ZUGFeRD
        S_ReExX
        DoEvents
        If AnzPo > 1 Then
            For AktPo = 1 To AnzPo
                XReDa = "X" & Left$(DatNa(AktPo), (Len(DatNa(AktPo)) - 3)) & "xml"
                SERech ExPfa & DatNa(AktPo), ExPfa & XReDa
            Next AktPo
        Else
            XReDa = "X" & Left$(DaNeu, (Len(DaNeu) - 3)) & "xml"
            SERech ExPfa & DaNeu, ExPfa & XReDa
        End If
    End If
    DoEvents
    If ZipDo = True Then
        ZpPfa = StVer
        If ZpPfa = vbNullString Then ZpPfa = TmPfa
        If Right$(ZpPfa, 1) <> "\" Then ZpPfa = ZpPfa & "\"
        ZipNa = ZpPfa & "Rechnungen_Belege_" & Format$(Now, "YYYYMMDD_HHMM") & ".zip"
        ZipOk = SZipp(ZipNa, ExPfa, True)
    End If

    If EmSen > 0 And EmSen < 3 Then 'Emailversand
        If ZipDo = True Then
            If ZipOk = True Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, ZipNa
            End If
        Else
            If ZuFer = True Then 'ZUGFeRD
                EReDa = "E" & DaNeu
                FiStr = ExPfa & EReDa
            Else
                FiStr = ExPfa & DaNeu
            End If
            SMaNe PatNr, EmAdr, , EmTex, EmBet, FiStr
        End If
    End If
Else
    SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
End If

Set clFil = Nothing

DoEvents
LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExBl " & Err.Number
Resume Next

End Sub
Public Sub LLExDv()
On Error GoTo DeErr
'Exportiert DATEV-Rechnungsbelege ohne ZIP/Email-Logik

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim StaPr As Boolean
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DaNeu As String 'neuer Dateiname
Dim DauDg As String
Dim MxCod As String
Dim ExPfa As String 'Exportpfad
Dim XReDa As String
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim DruOk As Boolean
Dim ZuFer As Boolean
Dim DevMo As DEVMODEA
Dim DatNa() As String
Dim SQL1, SQL2 As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2

Mld1 = "Es existieren keine Daten, die exportiert werden koennen"
Tit1 = "Keine Daten vorhanden"
AnzPo = UBound(GloDr)
If Right$(TmPfa, 1) <> "\" Then TmPfa = TmPfa & "\"
ExPfa = StVer
If ExPfa = vbNullString Then ExPfa = TmPfa
If ExPfa <> vbNullString Then
    If Right$(ExPfa, 1) <> "\" Then ExPfa = ExPfa & "\"
End If

If ExFmt = "PDA" Then
    ExFmt = "PDF"
    ZuFer = True 'ZUGFeRD
End If

ReDim DatNa(AnzPo)

SQL1 = "qryPrSimRe1"
SQL2 = "qryPrSimRe2"

If SQL1 = vbNullString Then
    WindowMess Mld1, Dial4, Tit1, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", ExPfa)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Quiet", 1)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.ShowResult", 0)

For AktPo = 1 To AnzPo
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = vbNullString
    MxCod = vbNullString
    DauDg = S_DrDi(GloDr(AktPo))
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        MxCod = S_ReMx(GloDr(AktPo))
    End If

    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))

    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If

    If StaPr = True Then
        DaNeu = "Rechnung_Beleg_" & RS31.Fields("Rechnungsnummer").Value & "." & LCase(ExFmt)
        DatNa(AktPo) = DaNeu

        AktZa = 1
        GesZa = RS32.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa

        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            LLDeV RS31, DauDg, False, MxCod, True
        Else
            LLDeV RS31, DauDg, , , True
        End If
        LLDeF RS32

        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNeu)
        Select Case ExFmt
        Case "TXT":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
        Case "HTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", "Rechnu")
        Case "MHTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", "Rechnu")
        Case "XLS":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", "Rechnu")
        Case "RTF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
        Case "PDF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", "Rechnu")
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
        End Select

        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)

        If RetWe = 0 Then
            If AktPo = 1 Then
                FS.Show
                DoEvents
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                While Not RS32.EOF
                    LLDeF RS32
                    PrBr2.Value = AktZa
                    DoEvents
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
    
    DoEvents
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Unload FS
DoEvents

If DruOk = True Then
    If ZuFer = True Then 'ZUGFeRD
        S_ReExX
        DoEvents
        For AktPo = 1 To AnzPo
            XReDa = "X" & Left$(DatNa(AktPo), (Len(DatNa(AktPo)) - 3)) & "xml"
            SERech ExPfa & DatNa(AktPo), ExPfa & XReDa
        Next AktPo
    End If
Else
    SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
End If

Set clFil = Nothing

DoEvents
LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExDv " & Err.Number
Resume Next

End Sub

Public Sub LLExFr()
On Error GoTo DeErr
'Exportiert den Fragebogen

Dim GesZa As Long
Dim AktZa As Long
Dim AktWe As Long
Dim RetWe As Long
Dim FraNr As Long
Dim SubNr As Long
Dim hwOrg As Long
Dim VaNum As Long
Dim VaStr As String
Dim VoStr As String
Dim ZipNa As String
Dim ExPfa As String 'Exportpfad
Dim ZpDir As String
Dim ZpPfa As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DaNeu As String 'Neuer Dateiname
Dim DauDg As String
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim VaPos As Integer
Dim TypNr As Integer
Dim AkPos As Integer
Dim DruOk As Boolean
Dim StaPr As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim TmpBo As Boolean
Dim ZipDo As Boolean
Dim ZipOk As Boolean
Dim VaAry() As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Mld1 = "Es existieren keine Daten, die exportiert werden können"
Tit1 = "Keine Daten vorhanden"
AnzPo = UBound(GloDr)
ZipDo = (AnzPo > 1)
If Right$(TmPfa, 1) <> "\" Then TmPfa = TmPfa & "\"
If ZipDo = True Then
    ZpDir = TmPfa & "LLZip_" & CreateID("Z")
    If Dir$(ZpDir, vbDirectory) = vbNullString Then MkDir ZpDir
    ZpDir = ZpDir & "\"
    ExPfa = ZpDir
Else
    ExPfa = StVer
    If ExPfa = vbNullString Then ExPfa = TmPfa
End If
If ExPfa <> vbNullString Then
    If Right$(ExPfa, 1) <> "\" Then ExPfa = ExPfa & "\"
End If

Select Case ForNa
Case "AnBeBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn2"
Case "AnFrBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
Case "AnFxBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
End Select

If SQL1 = vbNullString Then
    WindowMess Mld1, Dial4, Tit1, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", ExPfa)

For AktPo = 1 To AnzPo
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = vbNullString
    DauDg = S_AnTex(GloDr(AktPo))
    
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    
    If RS31.RecordCount > 0 Then
        If RS33.RecordCount > 0 Then
            StaPr = True
        End If
    End If

    If StaPr = True Then
        If VerMo = True Then
            If RS31.Fields("Em_Aut").Value <> vbNullString Then
                VerDo = CBool(RS31.Fields("Em_Aut").Value)
            Else
                VerDo = False
            End If
            If VerDo = True Then 'Dokument verschlüsseln
                If RS31.Fields("Em_Pass").Value <> vbNullString Then
                    VerPa = RS31.Fields("Em_Pass").Value
                Else
                    VerDo = False
                End If
            Else
                VerDo = False
            End If
        End If
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If
    
        If AnzPo > 1 Then
            DaNeu = "F" & RS31.Fields("ID5").Value & "." & LCase(ExFmt)
        Else
            If LCase(Right$(DaNam, 3)) = "zip" Then
                DaNam = Mid$(DaNam, 1, Len(DaNam) - 3) & LCase(ExFmt)
            End If
            DaNeu = DaNam
        End If

        AktZa = 1
        GesZa = RS33.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        
        Set RS32 = New ADODB.Recordset
        With RS32
            .CursorLocation = adUseClient
            .Fields.Append "Gruppe", adChar, 250
            .Fields.Append "IDKey", adChar, 20
            .Fields.Append "ID1", adBigInt
            .Fields.Append "ID6", adBigInt
            .Fields.Append "ID5", adBigInt
            .Fields.Append "ID0", adBigInt
            .Fields.Append "IDSub", adBigInt
            .Fields.Append "Wert", adChar, 250
            .Fields.Append "Kommentar", adVariant
            .Fields.Append "GuiID", adChar, 100
            .Fields.Append "VorText", adChar, 250
            .Fields.Append "Type", adInteger
            .Fields.Append "Sorter", adBigInt
            .Fields.Append "ID3", adBigInt
            .Fields.Append "Feldname", adChar, 100
            .Fields.Append "IDKurz", adChar, 250
            .Fields.Append "Bericht", adVariant
            .Fields.Append "Vorgaben", adVariant
            .Fields.Append "Zeilen", adBigInt
            .Fields.Append "Gross", adBigInt
            .Fields.Append "Lange", adBigInt
            .Fields.Append "Hohe", adBigInt
            .Fields.Append "Breite", adBigInt
            .Fields.Append "IDG", adBigInt
            If .State = adStateClosed Then .Open
        End With
        
        If RS32.Supports(adAddNew) Then
            Do
            If RS33.Fields("Type").Value <> vbNullString Then
                TypNr = CInt(RS33.Fields("Type").Value)
            Else
                TypNr = 1
            End If
            If RS33.Fields("ID6").Value > 0 Then
                FraNr = RS33.Fields("ID6").Value
            Else
                FraNr = 0
            End If
            If RS33.Fields("IDSub").Value > 0 Then
                SubNr = RS33.Fields("IDSub").Value
            Else
                SubNr = 0
            End If
            If RS33.Fields("Wert").Value <> vbNullString Then
                VaStr = RS33.Fields("Wert").Value
            Else
                VaStr = vbNullString
            End If
            If IsNumeric(VaStr) = True Then
                VaNum = CLng(VaStr)
            Else
                VaNum = 0
            End If
            VoStr = vbNullString
      
            RS32.AddNew
            RS32.Fields("IDSub").Value = SubNr
            If RS33.Fields("Gruppe").Value <> vbNullString Then RS32.Fields("Gruppe").Value = RS33.Fields("Gruppe").Value
            If RS33.Fields("IDKey").Value <> vbNullString Then RS32.Fields("IDKey").Value = RS33.Fields("IDKey").Value
            If RS33.Fields("ID1").Value <> vbNullString Then RS32.Fields("ID1").Value = RS33.Fields("ID1").Value
            If RS33.Fields("ID6").Value <> vbNullString Then RS32.Fields("ID6").Value = RS33.Fields("ID6").Value
            If RS33.Fields("ID5").Value <> vbNullString Then RS32.Fields("ID5").Value = RS33.Fields("ID5").Value
            If RS33.Fields("ID0").Value <> vbNullString Then RS32.Fields("ID0").Value = RS33.Fields("ID0").Value
            If RS33.Fields("GuiID").Value <> vbNullString Then RS32.Fields("GuiID").Value = RS33.Fields("GuiID").Value
            If RS33.Fields("VorText").Value <> vbNullString Then RS32.Fields("VorText").Value = RS33.Fields("VorText").Value
            If RS33.Fields("Type").Value <> vbNullString Then RS32.Fields("Type").Value = RS33.Fields("Type").Value
            If RS33.Fields("Sorter").Value <> vbNullString Then RS32.Fields("Sorter").Value = RS33.Fields("Sorter").Value
            If RS33.Fields("ID3").Value <> vbNullString Then RS32.Fields("ID3").Value = RS33.Fields("ID3").Value
            If RS33.Fields("Feldname").Value <> vbNullString Then RS32.Fields("Feldname").Value = RS33.Fields("Feldname").Value
            If RS33.Fields("IDKurz").Value <> vbNullString Then RS32.Fields("IDKurz").Value = RS33.Fields("IDKurz").Value
            If RS33.Fields("Bericht").Value <> vbNullString Then RS32.Fields("Bericht").Value = RS33.Fields("Bericht").Value
            If RS33.Fields("Zeilen").Value <> vbNullString Then RS32.Fields("Zeilen").Value = RS33.Fields("Zeilen").Value
            If RS33.Fields("Gross").Value <> vbNullString Then RS32.Fields("Gross").Value = RS33.Fields("Gross").Value
            If RS33.Fields("Lange").Value <> vbNullString Then RS32.Fields("Lange").Value = RS33.Fields("Lange").Value
            If RS33.Fields("Hohe").Value <> vbNullString Then RS32.Fields("Hohe").Value = RS33.Fields("Hohe").Value
            If RS33.Fields("Breite").Value <> vbNullString Then RS32.Fields("Breite").Value = RS33.Fields("Breite").Value
            If RS33.Fields("IDG").Value <> vbNullString Then RS32.Fields("IDG").Value = RS33.Fields("IDG").Value
    
            Select Case TypNr:
            Case 1: 'Textfeld
                    If VaStr <> vbNullString Then
                        RS32.Fields("Wert").Value = VaStr
                    End If
                    If RS33.Fields("Vorgaben").Value <> vbNullString Then
                        RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                    End If
            Case 2: 'Auswahlfeld
                    Set RS34 = New ADODB.Recordset
                    RS34.CursorLocation = adUseClient
                    Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                    If RS34.RecordCount > 0 Then
                        Do
                        If VoStr = vbNullString Then
                            VoStr = RS34.Fields("IDKurz").Value
                        Else
                            VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                        End If
                        RS34.MoveNext
                        Loop Until RS34.EOF
                    End If
                    RS34.Close
                    Set RS34 = Nothing
                    RS32.Fields("Wert").Value = VaStr
                    RS32.Fields("Vorgaben").Value = VoStr
            Case 3: 'Ankreuzfeld
                    If VaStr <> vbNullString Then
                        Select Case LCase(VaStr)
                        Case "ja": TmpBo = True
                        Case "nein": TmpBo = False
                        Case "yes": TmpBo = True
                        Case "no": TmpBo = False
                        Case "0": TmpBo = False
                        Case "1": TmpBo = True
                        Case "-1": TmpBo = True
                        Case "true": TmpBo = True
                        Case "false": TmpBo = False
                        Case Else: TmpBo = True
                        End Select
                        If TmpBo = True Then
                            RS32.Fields("Wert").Value = "Ja"
                        Else
                            RS32.Fields("Wert").Value = "Nein"
                        End If
                    End If
                    RS32.Fields("Vorgaben").Value = "Ja;Nein"
            Case 4: 'Einfachauswahl
                    If VaNum > 0 Then
                        Set RS34 = New ADODB.Recordset
                        RS34.CursorLocation = adUseClient
                        Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                        If RS34.RecordCount > 0 Then
                            AktWe = 1
                            Do
                            If AktWe = VaNum Then
                                VaStr = RS34.Fields("IDKurz").Value
                                Exit Do
                            End If
                            AktWe = AktWe + 1
                            RS34.MoveNext
                            Loop Until RS34.EOF
                            RS34.MoveFirst
                            Do
                            If VoStr = vbNullString Then
                                VoStr = RS34.Fields("IDKurz").Value
                            Else
                                VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                            End If
                            RS34.MoveNext
                            Loop Until RS34.EOF
                        End If
                        RS34.Close
                        Set RS34 = Nothing
                    End If
                    RS32.Fields("Wert").Value = VaStr
                    RS32.Fields("Vorgaben").Value = VoStr
            Case 5: 'Mehrfachauswahl
                    If VaNum > 0 Then
                        AktWe = 0
                        AkPos = 1
                        Do While VaNum <> 0
                            AktWe = VaNum And (VaNum - 1)
                            VaPos = (Math.Log(VaNum - AktWe) / Math.Log(2)) + 1
                            ReDim Preserve VaAry(AkPos)
                            VaAry(AkPos) = VaPos
                            VaNum = AktWe
                            AkPos = AkPos + 1
                        Loop
                        Set RS34 = New ADODB.Recordset
                        RS34.CursorLocation = adUseClient
                        Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                        If RS34.RecordCount > 0 Then
                            VaStr = vbNullString
                            AkPos = 1
                            Do
                            For VaPos = 1 To UBound(VaAry)
                                If VaAry(VaPos) = AkPos Then
                                    If VaStr = vbNullString Then
                                        VaStr = RS34.Fields("IDKurz").Value
                                    Else
                                        VaStr = VaStr & ";" & RS34.Fields("IDKurz").Value
                                    End If
                                End If
                            Next VaPos
                            AkPos = AkPos + 1
                            RS34.MoveNext
                            Loop Until RS34.EOF
                            RS34.MoveFirst
                            Do
                            If VoStr = vbNullString Then
                                VoStr = RS34.Fields("IDKurz").Value
                            Else
                                VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                            End If
                            RS34.MoveNext
                            Loop Until RS34.EOF
                        End If
                        RS34.Close
                        Set RS34 = Nothing
                        RS32.Fields("Wert").Value = VaStr
                        RS32.Fields("Vorgaben").Value = VoStr
                    End If
            Case 6: 'Zwischentext
                    RS32.Fields("Wert").Value = vbNullString
                    RS32.Fields("Vorgaben").Value = vbNullString
            Case 7: 'Datumsfeld
                    If IsDate(VaStr) = True Then
                        RS32.Fields("Wert").Value = Format$(VaStr, "dd.mm.yyyy")
                    End If
                    If RS33.Fields("Vorgaben").Value <> vbNullString Then
                        RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                    End If
            End Select
    
            RS33.MoveNext
            Loop Until RS33.EOF
            
            RS32.Update
            RS32.MoveFirst

            LLDeV RS31, DauDg
            LLDeF RS32

            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNeu)
            Select Case ExFmt
            Case "TXT":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
            Case "HTML":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
            Case "MHTML":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
            Case "XLS":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", ForNa)
            Case "RTF":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
            Case "PDF":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
                If VerDo = True Then 'Dokument verschlüsseln
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", VerPa)
                End If
            End Select

            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    
            If RetWe = 0 Then
                If AktPo = 1 Then
                    FS.Show
                    DoEvents
                End If
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        PrBr2.Value = AktZa
                        DoEvents
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
    End If
    
    RS31.Close
    RS32.Close
    RS33.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
    Set RS33 = Nothing
    
    DoEvents
    If TxDia.Text = "B" Then Exit For 'Abbrechen
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Unload FS
DoEvents
If DruOk = True Then
    If ZipDo = True Then
        ZpPfa = StVer
        If ZpPfa = vbNullString Then ZpPfa = TmPfa
        If Right$(ZpPfa, 1) <> "\" Then ZpPfa = ZpPfa & "\"
        ZipNa = ZpPfa & "Fragebogen_" & Format$(Now, "YYYYMMDD_HHMM") & ".zip"
        ZipOk = SZipp(ZipNa, ExPfa, True)
    End If
End If

LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExFr " & Err.Number
Resume Next

End Sub
Public Sub LLExKo()
On Error GoTo DeErr
'Exportiert mehrere kompakte Rechungen

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim PatNr As Long
Dim AkDat As Date
Dim TaSum As Single 'Tagessumme
Dim AnSum As Single 'Anzahlsungssumme
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim MeStr As String 'Memostring
Dim Warun As String 'Währung
Dim AkZif As String 'Ziffer
Dim DaNeu As String
Dim DauDg As String
Dim MxCod As String
Dim GuiNa As String
Dim ZipNa As String
Dim ExPfa As String 'Exportpfad
Dim ZpDir As String
Dim ZpPfa As String
Dim ReStr As String
Dim ReDat As String
Dim FiStr As String
Dim XReDa As String
Dim EReDa As String
Dim ExpOr As String  'Exportordner
Dim KriSt As String 'Kriterienstring
Dim VerPa As String 'Verschlüsselungspasswort
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim TagZa As Integer 'Tageszahl
Dim ReZal As Integer 'Rechungszahl
Dim LegZa As Integer 'Legendenzahl
Dim LeiZa As Integer 'Leiszungszahl
Dim GeZei As Integer 'Gesamttageszeit
Dim AkPos As Integer
Dim ZipDo As Boolean
Dim ZipOk As Boolean
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim ZuFer As Boolean 'ZUGFeRD
Dim DevMo As DEVMODEA
Dim DatNa() As String
Dim SQL1, SQL2, SQL3 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
AnzPo = UBound(GloDr)
ZipDo = (AnzPo > 1)
If Right$(TmPfa, 1) <> "\" Then TmPfa = TmPfa & "\"
If ZipDo = True Then
    ZpDir = TmPfa & "LLZip_" & CreateID("Z")
    If Dir$(ZpDir, vbDirectory) = vbNullString Then MkDir ZpDir
    ZpDir = ZpDir & "\"
    ExPfa = ZpDir
Else
    ExPfa = StVer
    If ExPfa = vbNullString Then ExPfa = TmPfa
End If
If ExPfa <> vbNullString Then
    If Right$(ExPfa, 1) <> "\" Then ExPfa = ExPfa & "\"
End If

If GlRDP = True Then
    If LCase(GlEPf) = LCase(GlIPf) Then
        ExpOr = GlDpf & "Import\"
    Else
        ExpOr = GlEPf
    End If
Else
    ExpOr = GlEPf
End If

If ExFmt = "PDA" Then
    ExFmt = "PDF"
    ZuFer = True 'ZUGFeRD
End If

ReDim DatNa(AnzPo)

Select Case ForNa
Case "RechKo":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechFa":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechKu":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechSk":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", ExPfa)

For AktPo = 1 To AnzPo
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = vbNullString
    MxCod = vbNullString
    DauDg = S_DrDi(GloDr(AktPo))
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        MxCod = S_ReMx(GloDr(AktPo))
    End If

    Set RS31 = New ADODB.Recordset 'Rechnungskopf
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    ReZal = RS31.RecordCount

    Set RS33 = New ADODB.Recordset 'Legende
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    LegZa = RS33.RecordCount
    
    Set RS34 = New ADODB.Recordset 'Abrechungstage
    RS34.CursorLocation = adUseClient
    Set RS34 = DBCmRe1(SQL3, "@IdStr", GloDr(AktPo))
    TagZa = RS34.RecordCount

    If ReZal > 0 Then
        If LegZa > 0 Then
            If TagZa > 0 Then
                StaPr = True
            End If
        End If
    End If

    If StaPr = True Then
        If RS31.Fields("Em_Aut").Value <> vbNullString Then
            VerDo = CBool(RS31.Fields("Em_Aut").Value)
        Else
            VerDo = False
        End If
        If VerDo = True Then
            If RS31.Fields("Em_Pass").Value <> vbNullString Then
                VerPa = RS31.Fields("Em_Pass").Value
            Else
                VerDo = False
            End If
        Else
            VerDo = False
        End If
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If
        
        If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
            ReStr = RS31.Fields("Rechnungsnummer").Value
        Else
            ReStr = Format$(AktPo, "0000")
        End If
        If RS31.Fields("Rechnungsdatum").Value <> vbNullString Then
            ReDat = Format$(RS31.Fields("Rechnungsdatum").Value, "dd.mm.yyyy")
        Else
            ReDat = Format$(Date, "dd.mm.yyyy")
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        Else
            PatNr = 0
        End If

        If AnzPo > 1 Then
            DaNeu = "Rechnung_Beleg_" & ReStr & "." & LCase(ExFmt)
        Else
            If LCase(Right$(DaNam, 3)) = "zip" Then
                DaNam = Mid$(DaNam, 1, Len(DaNam) - 3) & LCase(ExFmt)
            End If
            DaNeu = DaNam
        End If
                
        DatNa(AktPo) = DaNeu
        '    If ZuFer = True Then 'ZUGFeRD
        '        EReDa = "E" & DaNeu
        '    Else

        AktZa = 1
        GesZa = LegZa + TagZa
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        
        Set RS32 = New ADODB.Recordset
        With RS32
            .CursorLocation = adUseClient
            .Fields.Append "Datum", adDate
            .Fields.Append "Behandlungsdatum", adChar, 10
            .Fields.Append "Ziffer", adChar, 100
            .Fields.Append "Bezeichnung", adVariant
            .Fields.Append "Kommentar", adVariant
            .Fields.Append "Einzelbetrag", adDouble
            .Fields.Append "Gesamtbetrag", adDouble
            .Fields.Append "Anzahlungsbetrag", adDouble
            .Fields.Append "Währung", adChar, 10
            .Fields.Append "Anzahl", adInteger
            .Fields.Append "Multi", adDouble
            .Fields.Append "Typ", adInteger
            .Fields.Append "Zeit", adInteger
            .Fields.Append "ICD", adVariant
            .Fields.Append "Diagnose", adVariant
            .Fields.Append "Analog", adInteger
            .Open
        End With

        If RS32.Supports(adAddNew) Then
            Do
            If RS34.Fields("Datum").Value <> vbNullString Then
                AkDat = RS34.Fields("Datum").Value
                Set RS35 = New ADODB.Recordset
                RS35.CursorLocation = adUseClient
                If GlTyp < 2 Then
                    Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(AktPo), AkDat)
                Else
                    Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(AktPo), Format$(AkDat, "dd.mm.yyyy"))
                End If
                If RS35.RecordCount > 0 Then
                    MeStr = vbNullString
                    TaSum = 0
                    AnSum = 0
                    GeZei = 0
                    Do
                    If RS35.Fields("Ziffer").Value <> vbNullString Then
                        AkZif = RS35.Fields("Ziffer").Value
                        LeiZa = RS35.Fields("Anzahl").Value
        
                        If MeStr = vbNullString Then
                            If LeiZa = 1 Then
                                MeStr = AkZif
                            Else
                                MeStr = LeiZa & "x" & AkZif
                            End If
                        Else
                            If LeiZa = 1 Then
                                MeStr = MeStr & "; " & AkZif
                            Else
                                MeStr = MeStr & "; " & LeiZa & "x" & AkZif
                            End If
                        End If
                    End If
                    If RS35.Fields("Währung").Value <> vbNullString Then
                        Warun = RS35.Fields("Währung").Value
                    Else
                        Warun = "€"
                    End If
                    If RS35.Fields("Zeit").Value <> 0 Then
                        GeZei = GeZei + RS35.Fields("Zeit").Value
                    End If
                    If RS35.Fields("Gesamtbetrag").Value <> 0 Then
                        TaSum = TaSum + RS35.Fields("Gesamtbetrag").Value
                    End If
                    If RS35.Fields("Anzahlungsbetrag").Value <> 0 Then
                        AnSum = AnSum + RS35.Fields("Anzahlungsbetrag").Value
                    End If
                    RS35.MoveNext
                    Loop Until RS35.EOF
                    RS32.AddNew
                    RS32.Fields("Datum").Value = AkDat
                    RS32.Fields("Behandlungsdatum").Value = Format$(AkDat, "dd.mm.yyyy")
                    RS32.Fields("Bezeichnung").Value = MeStr
                    RS32.Fields("Gesamtbetrag").Value = TaSum
                    RS32.Fields("Anzahlungsbetrag").Value = AnSum
                    RS32.Fields("Währung").Value = Warun
                    RS32.Fields("Zeit").Value = GeZei
                    RS32.Fields("Typ").Value = 0
                End If
                RS35.Close
                Set RS35 = Nothing
            End If
            RS34.MoveNext
            Loop Until RS34.EOF
    
            Do
            RS32.AddNew
            If RS33.Fields("Ziffer").Value <> vbNullString Then RS32.Fields("Ziffer").Value = RS33.Fields("Ziffer").Value
            If RS33.Fields("Bezeichnung").Value <> vbNullString Then RS32.Fields("Bezeichnung").Value = RS33.Fields("Bezeichnung").Value
            If RS33.Fields("ICD").Value <> vbNullString Then RS32.Fields("ICD").Value = RS33.Fields("ICD").Value
            If RS33.Fields("Diagnose").Value <> vbNullString Then RS32.Fields("Diagnose").Value = RS33.Fields("Diagnose").Value
            If RS33.Fields("Multi").Value <> vbNullString Then RS32.Fields("Multi").Value = RS33.Fields("Multi").Value
            RS32.Fields("Typ").Value = RS33.Fields("Typ").Value
            If RS33.Fields("Einzelbetrag").Value <> 0 Then
                RS32.Fields("Einzelbetrag").Value = RS33.Fields("Einzelbetrag").Value
                RS32.Fields("Währung").Value = RS33.Fields("Währung").Value
            End If
            If GlAnl = True Then
                If RS33.Fields("Analog").Value <> vbNullString Then
                    RS32.Fields("Analog").Value = RS33.Fields("Analog").Value
                Else
                    RS32.Fields("Analog").Value = 0
                End If
            Else
                RS32.Fields("Analog").Value = 0
            End If
            If RS33.Fields("Kommentar").Value <> vbNullString Then RS32.Fields("Kommentar").Value = RS33.Fields("Kommentar").Value
            RS33.MoveNext
            Loop Until RS33.EOF
            
            RS32.Update
            RS32.MoveFirst
            
            If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
                LLDeV RS31, DauDg, False, MxCod, True
            Else
                LLDeV RS31, DauDg
            End If
            LLDeF RS32

            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNeu)
            Select Case ExFmt
            Case "TXT":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
            Case "HTML":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
            Case "MHTML":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
            Case "XLS":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", ForNa)
            Case "RTF":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
            Case "PDF":
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
                RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
                If VerDo = True Then 'Dokument verschlüsseln
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
                    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", VerPa)
                End If
            End Select

            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)

            If RetWe = 0 Then
                If AktPo = 1 Then
                    FS.Show
                    DoEvents
                End If
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        PrBr2.Value = AktZa
                        DoEvents
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        Else
            If AktPo = 1 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
                Exit For
            End If
        End If
        
        RS31.Close
        RS32.Close
        RS33.Close
        RS34.Close
        Set RS31 = Nothing
        Set RS32 = Nothing
        Set RS33 = Nothing
        Set RS34 = Nothing
    
        DoEvents
        PrBr1.Value = AktPo
        StaPr = False
    End If
Next AktPo

Unload FS
DoEvents

If DruOk = True Then
    If ZuFer = True Then 'ZUGFeRD
        S_ReExX
        DoEvents
        If AnzPo > 1 Then
            For AktPo = 1 To AnzPo
                XReDa = "X" & Left$(DatNa(AktPo), (Len(DatNa(AktPo)) - 3)) & "xml"
                SERech ExPfa & DatNa(AktPo), ExPfa & XReDa
            Next AktPo
        Else
            XReDa = "X" & Left$(DaNeu, (Len(DaNeu) - 3)) & "xml"
            SERech ExPfa & DaNeu, ExPfa & XReDa
        End If
    End If
    DoEvents
    If ZipDo = True Then
        ZpPfa = StVer
        If ZpPfa = vbNullString Then ZpPfa = TmPfa
        If Right$(ZpPfa, 1) <> "\" Then ZpPfa = ZpPfa & "\"
        ZipNa = ZpPfa & "Rechnungen_Belege_" & Format$(Now, "YYYYMMDD_HHMM") & ".zip"
        ZipOk = SZipp(ZipNa, ExPfa, True)
    End If

    If EmSen > 0 And EmSen < 3 Then 'Emailversand
        If ZipDo = True Then
            If ZipOk = True Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, ZipNa
            End If
        Else
            If ZuFer = True Then 'ZUGFeRD
                EReDa = "E" & DaNeu
                FiStr = ExPfa & EReDa
            Else
                FiStr = ExPfa & DaNeu
            End If
            SMaNe PatNr, EmAdr, , EmTex, EmBet, FiStr
        End If
    End If
Else
    SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
End If

Set clFil = Nothing

DoEvents
LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExKo " & Err.Number
Resume Next

End Sub
Public Sub LLExLi()
On Error GoTo DeErr
'Exportiert Listen

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim StaPr As Boolean
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim KriSt As String 'Kriterienstring
Dim DauDg As String
Dim ZipNa As String
Dim AnNam As String
Dim ForAn As String
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim AkPos As Integer
Dim DruOk As Boolean
Dim DevMo As DEVMODEA
Dim SQL1 As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

If GlFrn <> vbNullString Then 'Formulareordner
    AnNam = GlFrn & S_FoCh(ForNa)
Else
    AnNam = GlFrO & S_FoCh(ForNa)
End If
Mld1 = "Es existieren keine Daten, die exportiert werden können"
Tit1 = "Keine Daten vorhanden"
AkPos = InStrRev(AnNam, "\", Len(AnNam), 1)
DruDa = TmPfa & Mid$(AnNam, AkPos + 1, Len(AnNam) - (AkPos + 3)) & "LL"
AnzPo = UBound(GloDr)

If clFil.FilVor(AnNam) = False Then
    Exit Sub
End If

If GlTyp < 2 Then
    Select Case ForNa
    Case "RechLi": SQL1 = "SELECT * FROM dbo.qryPrSimRe7 ORDER BY ReNr DESC"
    Case "ReList": SQL1 = "SELECT * FROM dbo.qryPrSimRe7 ORDER BY ReNr DESC"
    Case "ResUbe": SQL1 = "SELECT * FROM dbo.qryPrSimRe8 ORDER BY Patient, Datum"
    Case "AbrLis": SQL1 = "SELECT * FROM dbo.qryPrAbrLi1 ORDER BY Beleg DESC, Datum, IDR"
    Case "BuList": SQL1 = "SELECT * FROM dbo.qryPrBuLis1 ORDER BY Datum DESC, Nr"
    Case "TerExp": SQL1 = "SELECT * FROM dbo.qryPrTerLis1 ORDER BY VonDat, ZeiVon"
    End Select
Else
    Select Case ForNa
    Case "RechLi": SQL1 = "SELECT * FROM qryPrSimRe7 ORDER BY [ReNr] DESC;"
    Case "ReList": SQL1 = "SELECT * FROM qryPrSimRe7 ORDER BY [ReNr] DESC;"
    Case "ResUbe": SQL1 = "SELECT * FROM qryPrSimRe8 ORDER BY [Patient], [Datum];"
    Case "AbrLis": SQL1 = "SELECT * FROM qryPrAbrLi1 ORDER BY [Beleg] DESC, [Datum], [IDR];"
    Case "BuList": SQL1 = "SELECT * FROM qryPrBuLis1 ORDER BY [Datum] DESC, [Nr];"
    Case "TerExp": SQL1 = "SELECT * FROM qryPrTerLis1 ORDER BY [VonDat], [ZeiVon];"
    End Select
End If

If SQL1 = vbNullString Then
    WindowMess Mld1, Dial4, Tit1, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

Set RS32 = New ADODB.Recordset
With RS32
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS32.RecordCount
If GesZa > 0 Then StaPr = True

If StaPr = True Then
    PrBr2.Min = 0
    PrBr2.Max = GesZa
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = vbNullString
    
    If ForNa = "ResUbe" Then
        DauDg = S_DrDi(RS32.Fields("ReNr").Value)
        If DauDg = vbNullString Then
            DauDg = "keine Diagnose"
        End If
    End If

    LLDeB
    LLDeF RS32, DauDg

    If DrVor = False Then
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNam)
        Select Case ExFmt
        Case "TXT":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
        Case "HTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "MHTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "XLS":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", ForNa)
        Case "RTF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
        Case "PDF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
        End Select
    End If

    If DrVor = False Then
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, AnNam, LL_PRINT_EXPORT, 0)
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    Else
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        If GlRDP = True Then
            RetWe = LlViewerProhibitAction(mHwnd, 116)
        End If
        RetWe = LlViewerProhibitAction(mHwnd, 115)
        If DrDia = True Then
            RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, AnNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
        Else
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, AnNam, LL_PRINT_EXPORT, 0)
        End If
        If DrSof = False Then
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
        End If
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    End If
    
    If RetWe = 0 Then
        If DrDia = True Then
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, AnNam)
            End If
        End If
        If RetWe = 0 Then
            If DrVor = False Then
                FS.Show
                DoEvents
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                While Not RS32.EOF
                    If ForNa = "ResUbe" Then
                        DauDg = S_DrDi(RS32.Fields("ReNr").Value)
                        If DauDg = vbNullString Then
                            DauDg = "keine Diagnose"
                        End If
                    End If
                    LLDeF RS32, DauDg
                    If DrVor = True Then 'Druckvorschau
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                    Else
                        PrBr2.Value = AktZa
                    End If
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
Else
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

RS32.Close
Set RS32 = Nothing

Select Case ForNa
Case "RechLi": DBCmEx0 "qrySimReDru"
Case "ReList": DBCmEx0 "qrySimReDru"
Case "ResUbe": DBCmEx0 "qrySimReDru"
Case "AbrLis": DBCmEx0 "qryPrOpDes"
Case "BuList": DBCmEx0 "qryPrBuDes"
Case "TerExp": DBCmEx0 "qryPrTerDes"
End Select

Unload FS
DoEvents

If DruOk = True Then
    If EmSen > 0 And EmSen < 3 Then 'Emailversand
        SMaNe 0, EmAdr, , EmTex, EmBet, StVer & DaNam
    End If
Else
    SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
End If

Set clFil = Nothing

If GesZa > 0 Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExLi " & Err.Number
Resume Next

End Sub
Public Sub LLExUn(ByVal TabRp As Integer)
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim AkWie As Integer
Dim DruOk As Boolean
Dim AkPos As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set FS = frmStatus
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo8 = FM.repCont8
Set RpCo3 = FM.repCont3
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Select Case TabRp
Case 1:
    Set RpRws = RpCo1.Rows
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
Case 2:
    Set RpRws = RpCo2.Rows
    Set RpCls = RpCo2.Columns
    Set RpSel = RpCo2.SelectedRows
Case 3:
    Set RpRws = RpCo3.Rows
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case 4:
    Set RpRws = RpCo4.Rows
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
Case 5:
    Set RpRws = RpCo5.Rows
    Set RpCls = RpCo5.Columns
    Set RpSel = RpCo5.SelectedRows
Case 6:
    Set RpRws = RpCo6.Rows
    Set RpCls = RpCo6.Columns
    Set RpSel = RpCo6.SelectedRows
Case 8:
    Set RpRws = RpCo8.Rows
    Set RpCls = RpCo8.Columns
    Set RpSel = RpCo8.SelectedRows
End Select

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
DaExt = Right$(FiNam, 3)

If Gesam = True Then
    GesZa = RpRws.Count
Else
    GesZa = RpSel.Count
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

If GesZa > 0 Then
    If Gesam = True Then
         Set RpRow = RpRws(0)
    Else
         Set RpRow = RpSel(0)
    End If
    If RpRow.GroupRow = True Then
        If Gesam = True Then
             Set RpRow = RpRws(1)
        Else
             Set RpRow = RpSel(1)
        End If
    End If
    
    PrBr2.Min = 0
    PrBr2.Max = GesZa
    FS.Caption = "Erzeugen der Druckdateien..."
        
    LDeVar RpRow, TabRp
    LDeFel RpRow, TabRp
    
    If DrVor = False Then
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNam)
        Select Case ExFmt
        Case "TXT":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "TXT.SeparatorChar", ";")
        Case "HTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "MHTML":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.AllInOneFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "HTML.Title", ForNa)
        Case "XLS":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.OnlyTableData", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "XLS.Title", ForNa)
        Case "RTF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Verbosity.Drawing", 0)
        Case "PDF":
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
        End Select
    End If

    RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
    RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
    RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
    RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)

    If RetWe = 0 Then
        If DrDia = True Then RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        If RetWe = 0 Then
            If DaExt = "lst" Then RetWe = LlPrint(mHwnd)
            RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
            If RetWe = 0 Then
                If Gesam = True Then
                    For Each RpRow In RpRws
                        If RpRow.GroupRow = False Then
                            LDeFel RpRow, TabRp
                            RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                        End If
                    Next RpRow
                Else
                    For Each RpRow In RpSel
                        If RpRow.GroupRow = False Then
                            LDeFel RpRow, TabRp
                            RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                        End If
                    Next RpRow
                End If
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

Unload FS
DoEvents

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If GesZa > 0 Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpRws = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCo8 = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLExUn " & Err.Number
Resume Next

End Sub
Public Sub LLFaLi()
On Error GoTo DeErr
'Faxt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim TeDru As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim StaPr As Boolean
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim AkPos As Integer
Dim DruOk As Boolean
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
AnzPo = UBound(GloDr)

SQL1 = "qryPrLaBe1"
SQL2 = "qryPrLaBe5"

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

Set RS31 = New ADODB.Recordset
RS31.CursorLocation = adUseClient
Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
Set RS32 = New ADODB.Recordset
RS32.CursorLocation = adUseClient
Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
If RS31.RecordCount > 0 Then
    If RS32.RecordCount > 0 Then
        StaPr = True
    End If
End If

If StaPr = True Then
    GesZa = RS32.RecordCount
    LLDeV RS31
    LLDeF RS32
    
    'RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, TeDru, 32, DevMo, 32)
    RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
    RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
    If GlRDP = True Then
        RetWe = LlViewerProhibitAction(mHwnd, 116)
    End If
    RetWe = LlViewerProhibitAction(mHwnd, 115)
    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_NORMAL, 0)
    RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    
    If RetWe = 0 Then
        RetWe = LlPrint(mHwnd)
        RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
        If RetWe = 0 Then
            If RetWe = 0 Then
                While Not RS32.EOF
                    LLDeF RS32
                    RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS31.Close
RS32.Close
Set RS31 = Nothing
Set RS32 = Nothing

LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLFaLi " & Err.Number
Resume Next

End Sub
Public Sub LLFaLs()
On Error GoTo DeErr
'Faxt mehrere Aufträge

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim TeDru As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim ZusDa As String 'Zsammengefasste Datei
Dim FilNa As String
Dim VorDa As String
Dim GuiNa As String
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim AkPos As Integer
Dim SeiZa As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
VorDa = GuiNa & "_" & "0001.LL"
ZusDa = TmPfa & VorDa
AnzPo = UBound(GloDr)

SQL1 = "qryPrLaBe1"
SQL2 = "qryPrLaBe5"

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

For AktPo = 1 To AnzPo
    FilNa = TmPfa & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & "_" & Format$(AktPo, "0000") & ".LL"
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."

    StaPr = False
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If

    If StaPr = True Then
        AktZa = 1
        GesZa = RS32.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        LLDeV RS31
        LLDeF RS32

        'RetWe = LlGetPrinterFromPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, TeDru, 32, DevMo, 32)
        RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        If GlRDP = True Then
            RetWe = LlViewerProhibitAction(mHwnd, 116)
        End If
        RetWe = LlViewerProhibitAction(mHwnd, 115)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)

        If RetWe = 0 Then
            If AktPo = 1 Then
                FS.Show
                DoEvents
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        PrBr2.Value = AktZa
                        DoEvents
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                                        
                    With clFil
                        .DaCop = DruDa & ";" & FilNa & vbNullChar
                        If .FilCop(3) = False Then
                            SPopu "Datei schreibgeschützt", "Die Datei kann nicht umbenant werden", IC48_Warning
                        Else
                            If RetWe = 0 Then DruOk = True
                        End If
                    End With

                End If
            End If
        End If
    Else
        If AktPo = 1 Then
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
            Exit For
        End If
    End If
    
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
    
    DoEvents
    If TxDia.Text = "B" Then Exit For 'Abbrechen
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Set clFil = Nothing

RetWe = LLStor(VorDa, GuiNa, AnzPo)

Unload FS
DoEvents

If DruOk = True Then
    hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
    SeiZa = LlStgsysGetPageCount(hwOrg)
    RetWe = LlStgsysPrint(hwOrg, DrNam, DrNam, 1, SeiZa, Kopie, LS_PRINTFLAG_FIT Or LS_PRINTFLAG_STACKEDCOPIES Or LS_PRINTFLAG_ABORTABLEMETER, "Formulardruck", FM.hwnd)
    LlStgsysStorageClose hwOrg
Else
    SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
End If

LlJobClose mHwnd

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLFaLs " & Err.Number
Resume Next

End Sub
Public Sub LLFrLi()
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim AktWe As Long
Dim RetWe As Long
Dim FraNr As Long
Dim SubNr As Long
Dim hwOrg As Long
Dim VaNum As Long
Dim PatNr As Long
Dim VaStr As String
Dim VoStr As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim DauDg As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim TmpBo As Boolean
Dim AkPos As Integer
Dim VaPos As Integer
Dim TypNr As Integer
Dim VaAry() As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"

Select Case ForNa
Case "AnBeBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn2"
        DauDg = S_AnTex(GloDr(1))
Case "AnFrBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
        DauDg = S_AnTex(GloDr(1))
Case "AnFxBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
        DauDg = S_AnTex(GloDr(1))
End Select

If SQL2 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

Set RS31 = New ADODB.Recordset
RS31.CursorLocation = adUseClient
Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))

Set RS33 = New ADODB.Recordset
RS33.CursorLocation = adUseClient
Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(1))

If RS31.RecordCount > 0 Then
    If RS33.RecordCount > 0 Then
        StaPr = True
    End If
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

If StaPr = True Then
    GesZa = RS33.RecordCount

    If VerMo = True Then 'Verschlüsselung möglich
        If RS31.Fields("ID0").Value > 0 Then
            PatNr = RS31.Fields("ID0").Value
        Else
            PatNr = 0
        End If
        If RS31.Fields("Em_Aut").Value <> vbNullString Then
            VerDo = CBool(RS31.Fields("Em_Aut").Value)
        Else
            VerDo = False
        End If
        If VerDo = True Then
            If RS31.Fields("Em_Pass").Value <> vbNullString Then
                VerPa = RS31.Fields("Em_Pass").Value
            Else
                VerDo = False
            End If
        Else
            VerDo = False
        End If
    End If
    
    If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
        VerDo = False
    End If

    Set RS32 = New ADODB.Recordset
    With RS32
        .CursorLocation = adUseClient
        .Fields.Append "Gruppe", adChar, 250
        .Fields.Append "IDKey", adChar, 20
        .Fields.Append "ID1", adBigInt
        .Fields.Append "ID6", adBigInt
        .Fields.Append "ID5", adBigInt
        .Fields.Append "ID0", adBigInt
        .Fields.Append "IDSub", adBigInt
        .Fields.Append "Wert", adChar, 250
        .Fields.Append "Kommentar", adVariant
        .Fields.Append "GuiID", adChar, 100
        .Fields.Append "VorText", adChar, 250
        .Fields.Append "Type", adInteger
        .Fields.Append "Sorter", adBigInt
        .Fields.Append "ID3", adBigInt
        .Fields.Append "Feldname", adChar, 100
        .Fields.Append "IDKurz", adChar, 250
        .Fields.Append "Bericht", adVariant
        .Fields.Append "Vorgaben", adVariant
        .Fields.Append "Zeilen", adBigInt
        .Fields.Append "Gross", adBigInt
        .Fields.Append "Lange", adBigInt
        .Fields.Append "Hohe", adBigInt
        .Fields.Append "Breite", adBigInt
        .Fields.Append "IDG", adBigInt
        If .State = adStateClosed Then .Open
    End With
    
    If RS32.Supports(adAddNew) Then
        Do
        If RS33.Fields("Type").Value <> vbNullString Then
            TypNr = CInt(RS33.Fields("Type").Value)
        Else
            TypNr = 1
        End If
        If RS33.Fields("ID6").Value > 0 Then
            FraNr = RS33.Fields("ID6").Value
        Else
            FraNr = 0
        End If
        If RS33.Fields("IDSub").Value > 0 Then
            SubNr = RS33.Fields("IDSub").Value
        Else
            SubNr = 0
        End If
        If RS33.Fields("Wert").Value <> vbNullString Then
            VaStr = RS33.Fields("Wert").Value
        Else
            VaStr = vbNullString
        End If
        If IsNumeric(VaStr) = True Then
            VaNum = CLng(VaStr)
        Else
            VaNum = 0
        End If
        VoStr = vbNullString

        RS32.AddNew
        RS32.Fields("IDSub").Value = SubNr
        If RS33.Fields("Gruppe").Value <> vbNullString Then RS32.Fields("Gruppe").Value = RS33.Fields("Gruppe").Value
        If RS33.Fields("IDKey").Value <> vbNullString Then RS32.Fields("IDKey").Value = RS33.Fields("IDKey").Value
        If RS33.Fields("ID1").Value <> vbNullString Then RS32.Fields("ID1").Value = RS33.Fields("ID1").Value
        If RS33.Fields("ID6").Value <> vbNullString Then RS32.Fields("ID6").Value = RS33.Fields("ID6").Value
        If RS33.Fields("ID5").Value <> vbNullString Then RS32.Fields("ID5").Value = RS33.Fields("ID5").Value
        If RS33.Fields("ID0").Value <> vbNullString Then RS32.Fields("ID0").Value = RS33.Fields("ID0").Value
        If RS33.Fields("GuiID").Value <> vbNullString Then RS32.Fields("GuiID").Value = RS33.Fields("GuiID").Value
        If RS33.Fields("VorText").Value <> vbNullString Then RS32.Fields("VorText").Value = RS33.Fields("VorText").Value
        If RS33.Fields("Type").Value <> vbNullString Then RS32.Fields("Type").Value = RS33.Fields("Type").Value
        If RS33.Fields("Sorter").Value <> vbNullString Then RS32.Fields("Sorter").Value = RS33.Fields("Sorter").Value
        If RS33.Fields("ID3").Value <> vbNullString Then RS32.Fields("ID3").Value = RS33.Fields("ID3").Value
        If RS33.Fields("Feldname").Value <> vbNullString Then RS32.Fields("Feldname").Value = RS33.Fields("Feldname").Value
        If RS33.Fields("IDKurz").Value <> vbNullString Then RS32.Fields("IDKurz").Value = RS33.Fields("IDKurz").Value
        If RS33.Fields("Bericht").Value <> vbNullString Then RS32.Fields("Bericht").Value = RS33.Fields("Bericht").Value
        If RS33.Fields("Zeilen").Value <> vbNullString Then RS32.Fields("Zeilen").Value = RS33.Fields("Zeilen").Value
        If RS33.Fields("Gross").Value <> vbNullString Then RS32.Fields("Gross").Value = RS33.Fields("Gross").Value
        If RS33.Fields("Lange").Value <> vbNullString Then RS32.Fields("Lange").Value = RS33.Fields("Lange").Value
        If RS33.Fields("Hohe").Value <> vbNullString Then RS32.Fields("Hohe").Value = RS33.Fields("Hohe").Value
        If RS33.Fields("Breite").Value <> vbNullString Then RS32.Fields("Breite").Value = RS33.Fields("Breite").Value
        If RS33.Fields("IDG").Value <> vbNullString Then RS32.Fields("IDG").Value = RS33.Fields("IDG").Value
        
        Select Case TypNr:
        Case 1: 'Textfeld
                If VaStr <> vbNullString Then
                    RS32.Fields("Wert").Value = VaStr
                End If
                If RS33.Fields("Vorgaben").Value <> vbNullString Then
                    RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                End If
        Case 2: 'Auswahlfeld
                Set RS34 = New ADODB.Recordset
                RS34.CursorLocation = adUseClient
                Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                If RS34.RecordCount > 0 Then
                    Do
                    If VoStr = vbNullString Then
                        VoStr = RS34.Fields("IDKurz").Value
                    Else
                        VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                    End If
                    RS34.MoveNext
                    Loop Until RS34.EOF
                End If
                RS34.Close
                Set RS34 = Nothing
                RS32.Fields("Wert").Value = VaStr
                RS32.Fields("Vorgaben").Value = VoStr
        Case 3: 'Ankreuzfeld
                If VaStr <> vbNullString Then
                    Select Case LCase(VaStr)
                    Case "ja": TmpBo = True
                    Case "nein": TmpBo = False
                    Case "yes": TmpBo = True
                    Case "no": TmpBo = False
                    Case "0": TmpBo = False
                    Case "1": TmpBo = True
                    Case "-1": TmpBo = True
                    Case "true": TmpBo = True
                    Case "false": TmpBo = False
                    Case Else: TmpBo = True
                    End Select
                    If TmpBo = True Then
                        RS32.Fields("Wert").Value = "Ja"
                    Else
                        RS32.Fields("Wert").Value = "Nein"
                    End If
                End If
                RS32.Fields("Vorgaben").Value = "Ja;Nein"
        Case 4: 'Einfachauswahl
                Set RS34 = New ADODB.Recordset
                RS34.CursorLocation = adUseClient
                Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                If RS34.RecordCount > 0 Then
                    If VaNum > 0 Then
                        AktWe = 1
                        Do
                        If AktWe = VaNum Then
                            VaStr = RS34.Fields("IDKurz").Value
                            Exit Do
                        End If
                        AktWe = AktWe + 1
                        RS34.MoveNext
                        Loop Until RS34.EOF
                        RS34.MoveFirst
                    End If
                    Do
                    If VoStr = vbNullString Then
                        VoStr = RS34.Fields("IDKurz").Value
                    Else
                        VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                    End If
                    RS34.MoveNext
                    Loop Until RS34.EOF
                End If
                RS34.Close
                Set RS34 = Nothing
                RS32.Fields("Wert").Value = VaStr
                RS32.Fields("Vorgaben").Value = VoStr
        Case 5: 'Mehrfachauswahl
                AktWe = 0
                AkPos = 1
                Do While VaNum <> 0
                    AktWe = VaNum And (VaNum - 1)
                    VaPos = (Math.Log(VaNum - AktWe) / Math.Log(2)) + 1
                    ReDim Preserve VaAry(AkPos)
                    VaAry(AkPos) = VaPos
                    VaNum = AktWe
                    AkPos = AkPos + 1
                Loop
                Set RS34 = New ADODB.Recordset
                RS34.CursorLocation = adUseClient
                Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                If RS34.RecordCount > 0 Then
                    If AkPos > 1 Then
                        VaStr = vbNullString
                        AkPos = 1
                        Do
                        For VaPos = 1 To UBound(VaAry)
                            If VaAry(VaPos) = AkPos Then
                                If VaStr = vbNullString Then
                                    VaStr = RS34.Fields("IDKurz").Value
                                Else
                                    VaStr = VaStr & ";" & RS34.Fields("IDKurz").Value
                                End If
                            End If
                        Next VaPos
                        AkPos = AkPos + 1
                        RS34.MoveNext
                        Loop Until RS34.EOF
                        RS34.MoveFirst
                    End If
                    Do
                    If VoStr = vbNullString Then
                        VoStr = RS34.Fields("IDKurz").Value
                    Else
                        VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                    End If
                    RS34.MoveNext
                    Loop Until RS34.EOF
                End If
                RS34.Close
                Set RS34 = Nothing
                RS32.Fields("Wert").Value = VaStr
                RS32.Fields("Vorgaben").Value = VoStr
        Case 6: 'Zwischentext
                RS32.Fields("Wert").Value = vbNullString
                RS32.Fields("Vorgaben").Value = vbNullString
        Case 7: 'Datumsfeld
                If IsDate(VaStr) = True Then
                    RS32.Fields("Wert").Value = Format$(VaStr, "dd.mm.yyyy")
                End If
                If RS33.Fields("Vorgaben").Value <> vbNullString Then
                    RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                End If
        End Select

        RS33.MoveNext
        Loop Until RS33.EOF
        
        RS32.Update
        RS32.MoveFirst

        LLDeV RS31, DauDg
        LLDeF RS32

        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
        End If
    
        If FaxNr <> vbNullString Then
            RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
        End If

        If RetWe = 0 Then
            If DrDia = True Then
                RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
            End If
            If RetWe = 0 Then
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                    If RetWe = 0 Then
                        While Not RS32.EOF
                            LLDeF RS32
                            RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                            RS32.MoveNext
                        Wend
                        Do
                        RetWe = LlPrintFieldsEnd(mHwnd)
                        Loop Until RetWe <> LL_WRN_REPEAT_DATA
                        RetWe = LlPrintEnd(mHwnd, 0)
                        If RetWe = 0 Then DruOk = True
                    End If
                End If
            End If
        End If
    End If
    RS32.Close
    Set RS32 = Nothing
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS31.Close
RS33.Close
Set RS31 = Nothing
Set RS33 = Nothing

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            If VerDo = True Then
                hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
               SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLFrLi " & Err.Number
Resume Next

End Sub
Public Sub LLFrLs()
On Error GoTo DeErr
'Druckt mehrere Aufträge

Dim GesZa As Long
Dim AktZa As Long
Dim AktWe As Long
Dim RetWe As Long
Dim FraNr As Long
Dim SubNr As Long
Dim hwOrg As Long
Dim VaNum As Long
Dim PatNr As Long
Dim VaStr As String
Dim VoStr As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim ZusDa As String 'Zsammengefasste Datei
Dim FilNa As String
Dim DauDg As String
Dim VorDa As String
Dim GuiNa As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim TmpBo As Boolean
Dim AkPos As Integer
Dim VaPos As Integer
Dim TypNr As Integer
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim SeiZa As Integer
Dim VaAry() As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain
Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
VorDa = GuiNa & "_" & "0001.LL"
ZusDa = TmPfa & VorDa
AnzPo = UBound(GloDr)

Select Case ForNa
Case "AnBeBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn2"
Case "AnFrBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
Case "AnFxBo":
        VerMo = True
        SQL1 = "qryPrAn1"
        SQL2 = "qryPrAn4"
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

For AktPo = 1 To AnzPo
    FilNa = TmPfa & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & "_" & Format$(AktPo, "0000") & ".LL"
    DauDg = vbNullString
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DauDg = S_AnTex(GloDr(AktPo))

    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    
    If RS31.RecordCount > 0 Then
        If RS33.RecordCount > 0 Then
            StaPr = True
        End If
    End If

    If StaPr = True Then
        AktZa = 1
        GesZa = RS33.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa

        If VerMo = True Then
            If RS31.Fields("ID0").Value > 0 Then
                PatNr = RS31.Fields("ID0").Value
            Else
                PatNr = 0
            End If
            If RS31.Fields("Em_Aut").Value <> vbNullString Then
                VerDo = CBool(RS31.Fields("Em_Aut").Value)
            Else
                VerDo = False
            End If
            If VerDo = True Then
                If RS31.Fields("Em_Pass").Value <> vbNullString Then
                    VerPa = RS31.Fields("Em_Pass").Value
                Else
                    VerDo = False
                End If
            Else
                VerDo = False
            End If
        End If
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If

        Set RS32 = New ADODB.Recordset
        With RS32
            .CursorLocation = adUseClient
            .Fields.Append "Gruppe", adChar, 250
            .Fields.Append "IDKey", adChar, 20
            .Fields.Append "ID1", adBigInt
            .Fields.Append "ID6", adBigInt
            .Fields.Append "ID5", adBigInt
            .Fields.Append "ID0", adBigInt
            .Fields.Append "IDSub", adBigInt
            .Fields.Append "Wert", adChar, 250
            .Fields.Append "Kommentar", adVariant
            .Fields.Append "GuiID", adChar, 100
            .Fields.Append "VorText", adChar, 250
            .Fields.Append "Type", adInteger
            .Fields.Append "Sorter", adBigInt
            .Fields.Append "ID3", adBigInt
            .Fields.Append "Feldname", adChar, 100
            .Fields.Append "IDKurz", adChar, 250
            .Fields.Append "Bericht", adVariant
            .Fields.Append "Vorgaben", adVariant
            .Fields.Append "Zeilen", adBigInt
            .Fields.Append "Gross", adBigInt
            .Fields.Append "Lange", adBigInt
            .Fields.Append "Hohe", adBigInt
            .Fields.Append "Breite", adBigInt
            .Fields.Append "IDG", adBigInt
            If .State = adStateClosed Then .Open
        End With
        
        If RS32.Supports(adAddNew) Then
            Do
            If RS33.Fields("Type").Value <> vbNullString Then
                TypNr = CInt(RS33.Fields("Type").Value)
            Else
                TypNr = 1
            End If
            If RS33.Fields("ID6").Value > 0 Then
                FraNr = RS33.Fields("ID6").Value
            Else
                FraNr = 0
            End If
            If RS33.Fields("IDSub").Value > 0 Then
                SubNr = RS33.Fields("IDSub").Value
            Else
                SubNr = 0
            End If
            If RS33.Fields("Wert").Value <> vbNullString Then
                VaStr = RS33.Fields("Wert").Value
            Else
                VaStr = vbNullString
            End If
            If IsNumeric(VaStr) = True Then
                VaNum = CLng(VaStr)
            Else
                VaNum = 0
            End If
            VoStr = vbNullString
      
            RS32.AddNew
            RS32.Fields("IDSub").Value = SubNr
            If RS33.Fields("Gruppe").Value <> vbNullString Then RS32.Fields("Gruppe").Value = RS33.Fields("Gruppe").Value
            If RS33.Fields("IDKey").Value <> vbNullString Then RS32.Fields("IDKey").Value = RS33.Fields("IDKey").Value
            If RS33.Fields("ID1").Value <> vbNullString Then RS32.Fields("ID1").Value = RS33.Fields("ID1").Value
            If RS33.Fields("ID6").Value <> vbNullString Then RS32.Fields("ID6").Value = RS33.Fields("ID6").Value
            If RS33.Fields("ID5").Value <> vbNullString Then RS32.Fields("ID5").Value = RS33.Fields("ID5").Value
            If RS33.Fields("ID0").Value <> vbNullString Then RS32.Fields("ID0").Value = RS33.Fields("ID0").Value
            If RS33.Fields("GuiID").Value <> vbNullString Then RS32.Fields("GuiID").Value = RS33.Fields("GuiID").Value
            If RS33.Fields("VorText").Value <> vbNullString Then RS32.Fields("VorText").Value = RS33.Fields("VorText").Value
            If RS33.Fields("Type").Value <> vbNullString Then RS32.Fields("Type").Value = RS33.Fields("Type").Value
            If RS33.Fields("Sorter").Value <> vbNullString Then RS32.Fields("Sorter").Value = RS33.Fields("Sorter").Value
            If RS33.Fields("ID3").Value <> vbNullString Then RS32.Fields("ID3").Value = RS33.Fields("ID3").Value
            If RS33.Fields("Feldname").Value <> vbNullString Then RS32.Fields("Feldname").Value = RS33.Fields("Feldname").Value
            If RS33.Fields("IDKurz").Value <> vbNullString Then RS32.Fields("IDKurz").Value = RS33.Fields("IDKurz").Value
            If RS33.Fields("Bericht").Value <> vbNullString Then RS32.Fields("Bericht").Value = RS33.Fields("Bericht").Value
            If RS33.Fields("Zeilen").Value <> vbNullString Then RS32.Fields("Zeilen").Value = RS33.Fields("Zeilen").Value
            If RS33.Fields("Gross").Value <> vbNullString Then RS32.Fields("Gross").Value = RS33.Fields("Gross").Value
            If RS33.Fields("Lange").Value <> vbNullString Then RS32.Fields("Lange").Value = RS33.Fields("Lange").Value
            If RS33.Fields("Hohe").Value <> vbNullString Then RS32.Fields("Hohe").Value = RS33.Fields("Hohe").Value
            If RS33.Fields("Breite").Value <> vbNullString Then RS32.Fields("Breite").Value = RS33.Fields("Breite").Value
            If RS33.Fields("IDG").Value <> vbNullString Then RS32.Fields("IDG").Value = RS33.Fields("IDG").Value
    
            Select Case TypNr:
            Case 1: 'Textfeld
                    If VaStr <> vbNullString Then
                        RS32.Fields("Wert").Value = VaStr
                    End If
                    If RS33.Fields("Vorgaben").Value <> vbNullString Then
                        RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                    End If
            Case 2: 'Auswahlfeld
                    Set RS34 = New ADODB.Recordset
                    RS34.CursorLocation = adUseClient
                    Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                    If RS34.RecordCount > 0 Then
                        Do
                        If VoStr = vbNullString Then
                            VoStr = RS34.Fields("IDKurz").Value
                        Else
                            VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                        End If
                        RS34.MoveNext
                        Loop Until RS34.EOF
                    End If
                    RS34.Close
                    Set RS34 = Nothing
                    RS32.Fields("Wert").Value = VaStr
                    RS32.Fields("Vorgaben").Value = VoStr
            Case 3: 'Ankreuzfeld
                    If VaStr <> vbNullString Then
                        Select Case LCase(VaStr)
                        Case "ja": TmpBo = True
                        Case "nein": TmpBo = False
                        Case "yes": TmpBo = True
                        Case "no": TmpBo = False
                        Case "0": TmpBo = False
                        Case "1": TmpBo = True
                        Case "-1": TmpBo = True
                        Case "true": TmpBo = True
                        Case "false": TmpBo = False
                        Case Else: TmpBo = True
                        End Select
                        If TmpBo = True Then
                            RS32.Fields("Wert").Value = "Ja"
                        Else
                            RS32.Fields("Wert").Value = "Nein"
                        End If
                    End If
                    RS32.Fields("Vorgaben").Value = "Ja;Nein"
            Case 4: 'Einfachauswahl
                    If VaNum > 0 Then
                        Set RS34 = New ADODB.Recordset
                        RS34.CursorLocation = adUseClient
                        Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                        If RS34.RecordCount > 0 Then
                            AktWe = 1
                            Do
                            If AktWe = VaNum Then
                                VaStr = RS34.Fields("IDKurz").Value
                                Exit Do
                            End If
                            AktWe = AktWe + 1
                            RS34.MoveNext
                            Loop Until RS34.EOF
                            RS34.MoveFirst
                            Do
                            If VoStr = vbNullString Then
                                VoStr = RS34.Fields("IDKurz").Value
                            Else
                                VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                            End If
                            RS34.MoveNext
                            Loop Until RS34.EOF
                        End If
                        RS34.Close
                        Set RS34 = Nothing
                    End If
                    RS32.Fields("Wert").Value = VaStr
                    RS32.Fields("Vorgaben").Value = VoStr
            Case 5: 'Mehrfachauswahl
                    If VaNum > 0 Then
                        AktWe = 0
                        AkPos = 1
                        Do While VaNum <> 0
                            AktWe = VaNum And (VaNum - 1)
                            VaPos = (Math.Log(VaNum - AktWe) / Math.Log(2)) + 1
                            ReDim Preserve VaAry(AkPos)
                            VaAry(AkPos) = VaPos
                            VaNum = AktWe
                            AkPos = AkPos + 1
                        Loop
                        Set RS34 = New ADODB.Recordset
                        RS34.CursorLocation = adUseClient
                        Set RS34 = DBCmRe1("qryKat08S", "@IdxNr", FraNr) 'Subfragen
                        If RS34.RecordCount > 0 Then
                            VaStr = vbNullString
                            AkPos = 1
                            Do
                            For VaPos = 1 To UBound(VaAry)
                                If VaAry(VaPos) = AkPos Then
                                    If VaStr = vbNullString Then
                                        VaStr = RS34.Fields("IDKurz").Value
                                    Else
                                        VaStr = VaStr & ";" & RS34.Fields("IDKurz").Value
                                    End If
                                End If
                            Next VaPos
                            AkPos = AkPos + 1
                            RS34.MoveNext
                            Loop Until RS34.EOF
                            RS34.MoveFirst
                            Do
                            If VoStr = vbNullString Then
                                VoStr = RS34.Fields("IDKurz").Value
                            Else
                                VoStr = VoStr & ";" & RS34.Fields("IDKurz").Value
                            End If
                            RS34.MoveNext
                            Loop Until RS34.EOF
                        End If
                        RS34.Close
                        Set RS34 = Nothing
                        RS32.Fields("Wert").Value = VaStr
                        RS32.Fields("Vorgaben").Value = VoStr
                    End If
            Case 6: 'Zwischentext
                    RS32.Fields("Wert").Value = vbNullString
                    RS32.Fields("Vorgaben").Value = vbNullString
            Case 7: 'Datumsfeld
                    If IsDate(VaStr) = True Then
                        RS32.Fields("Wert").Value = Format$(VaStr, "dd.mm.yyyy")
                    End If
                    If RS33.Fields("Vorgaben").Value <> vbNullString Then
                        RS32.Fields("Vorgaben").Value = RS33.Fields("Vorgaben").Value
                    End If
            End Select
    
            RS33.MoveNext
            Loop Until RS33.EOF
            
            RS32.Update
            RS32.MoveFirst
            
            LLDeV RS31, DauDg
            LLDeF RS32
    
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                If DrSof = False Then
                    RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
                End If
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
        
            If FaxNr <> vbNullString Then
                RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
            End If
    
            If RetWe = 0 Then
                If AktPo = 1 Then
                    FS.Show
                    DoEvents
                End If
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                    If RetWe = 0 Then
                        While Not RS32.EOF
                            LLDeF RS32
                            PrBr2.Value = AktZa
                            DoEvents
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                            RS32.MoveNext
                        Wend
                        Do
                        RetWe = LlPrintFieldsEnd(mHwnd)
                        Loop Until RetWe <> LL_WRN_REPEAT_DATA
                        RetWe = LlPrintEnd(mHwnd, 0)
                        If RetWe = 0 Then DruOk = True
                                                                        
                        If DrVor = True Then 'Druckvorschau
                            With clFil
                                .DaCop = DruDa & ";" & FilNa & vbNullChar
                                If .FilCop(3) = False Then
                                    SPopu "Datei schreibgeschützt", "Die Datei kann nicht umbenannt werden", IC48_Warning
                                Else
                                    If RetWe = 0 Then DruOk = True
                                End If
                            End With
                        Else
                            If RetWe = 0 Then DruOk = True
                        End If
                        
                    End If
                End If
            End If
        End If
    Else
        If AktPo = 1 Then
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
            Exit For
        End If
    End If
        
    RS31.Close
    RS32.Close
    RS33.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
    Set RS33 = Nothing

    DoEvents
    If TxDia.Text = "B" Then Exit For 'Abbrechen
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Set clFil = Nothing

RetWe = LLStor(VorDa, GuiNa, AnzPo)

Unload FS
DoEvents

If DruOk = True Then
    If DrSof = True Then
        hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
        SeiZa = LlStgsysGetPageCount(hwOrg)
        RetWe = LlStgsysPrint(hwOrg, DrNam, DrNam, 1, SeiZa, Kopie, LS_PRINTFLAG_FIT Or LS_PRINTFLAG_STACKEDCOPIES Or LS_PRINTFLAG_ABORTABLEMETER, "Formulardruck", FM.hwnd)
        LlStgsysStorageClose hwOrg
    Else
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            If VerDo = True Then
                hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, ZusDa, TmPfa, FM.hwnd)
            End If
        End If
    End If
Else
    SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If EmSen = 0 Then 'Emailversand
    If DrVor = True Then 'Druckvorschau
        If DruOk = True Then
            Select Case GlVTy
            Case 1: VoDrMa ZusDa
            Case 3: SDrVo ZusDa
            End Select
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLFrLs " & Err.Number
Resume Next

End Sub
Public Sub LLHeLs()
On Error GoTo DeErr
'Druckt mehrere Aufträge

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim MitNr As Long
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim ZusDa As String 'Zsammengefasste Datei
Dim DauDg As String
Dim MxCod As String
Dim EmEmp As String
Dim EmNam As String
Dim EmBrf As String
Dim ReStr As String
Dim ReDat As String
Dim EmSig As String
Dim FilNa As String
Dim VorDa As String
Dim GuiNa As String
Dim MaBet As String
Dim MaTex As String
Dim XReDa As String
Dim EReDa As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim EmVer As Boolean 'Email erfolgreich versandt
Dim StaPr As Boolean
Dim DoLnk As String
Dim PoLnk As String
Dim DoStr As String
Dim ExpOr As String  'Exportordner
Dim DruOk As Boolean
Dim CmTex As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim AkPos As Integer
Dim SeiZa As Integer
Dim Posit As Integer
Dim Lange As Integer
Dim MaIdx As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
VorDa = GuiNa & "_" & "0001.LL"
ZusDa = TmPfa & VorDa
AnzPo = UBound(GloDr)

If GlRDP = True Then
    If LCase(GlEPf) = LCase(GlIPf) Then
        ExpOr = GlDpf & "Import\"
    Else
        ExpOr = GlEPf
    End If
Else
    ExpOr = GlEPf
End If

Select Case ForNa
Case "EMRAbr":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2a"
Case "Rechnu":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Auslan":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "RechLa":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "RechAl":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
Case "Besche":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
Case "RechSt":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "StAnge":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "StGuts":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Gutsch":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "KraBla":
        VerMo = True
        SQL1 = "qryPrSimKr1"
        CmTex = True
Case "EiMahn": 'Einzelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
Case "SaMahn": 'Sammelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP4"
        SQL2 = "qryPrSimOP5"
Case "StMahn": 'Einzelmahnung (Alternativ)
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

If EmSen > 0 Then
    DrVor = True 'wenn Emailversand, dann Druckvorschau
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

For AktPo = 1 To AnzPo
    FilNa = TmPfa & GuiNa & "_" & Format$(AktPo, "0000") & ".LL"
    DauDg = vbNullString
    MxCod = vbNullString
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DrKop = GloKo(AktPo)
    If ForNa = "KraBla" Then
        DauDg = S_KrDp(GloDr(AktPo))
    Else
        DauDg = S_DrDi(GloDr(AktPo))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(AktPo))
        End If
    End If

    StaPr = False
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    If CmTex = True Then
        Select Case ForNa
        Case "KraBla":
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimKr5 WHERE (" & Krite & " AND (PatNr = " & GloDr(AktPo) & ")) ORDER BY Datum"
            Else
                SQL2 = "SELECT * FROM qryPrSimKr5 WHERE (" & Krite & " AND ([PatNr] = " & GloDr(AktPo) & ")) ORDER BY [Datum];"
            End If
        Case Else:
        End Select
        Set RS32 = DBCmRe0(SQL2, CmTex)
    Else
        Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    End If
    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If
    If StaPr = True Then
        If VerMo = True Then
            If RS31.Fields("Em_Aut").Value <> vbNullString Then
                VerDo = CBool(RS31.Fields("Em_Aut").Value)
            Else
                VerDo = False
            End If
            If VerDo = True Then
                If RS31.Fields("Em_Pass").Value <> vbNullString Then
                    VerPa = RS31.Fields("Em_Pass").Value
                Else
                    VerDo = False
                End If
            Else
                VerDo = False
            End If
        End If
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If
    
        If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
            ReStr = RS31.Fields("Rechnungsnummer").Value
        Else
            ReStr = Format$(AktPo, "0000")
        End If
        If RS31.Fields("Rechnungsdatum").Value <> vbNullString Then
            ReDat = Format$(RS31.Fields("Rechnungsdatum").Value, "dd.mm.yyyy")
        Else
            ReDat = Format$(Date, "dd.mm.yyyy")
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        Else
            PatNr = 0
        End If
        EmNam = RS31.Fields("R_Name").Value & ", " & RS31.Fields("R_Vorname").Value
        If RS31.Fields("Telefon5").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon5").Value
        ElseIf RS31.Fields("Telefon6").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon6").Value
        Else
            EmEmp = vbNullString
        End If
        If RS31.Fields("R_Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("R_Briefanrede").Value
        ElseIf RS31.Fields("Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("Briefanrede").Value
        Else
            EmBrf = vbNullString
        End If
        If RS31.Fields("IDP").Value <> vbNullString Then
            ManNr = RS31.Fields("IDP").Value
        Else
            ManNr = 0
        End If
        If RS31.Fields("IDM").Value <> vbNullString Then
            MitNr = RS31.Fields("IDM").Value
        Else
            MitNr = 0
        End If
        If GlMiA(GlSmI, 11) <> vbNullString Then
            EmSig = GlMiA(GlSmI, 11)
        Else
            EmSig = GlMiA(GlSmI, 1)
        End If

        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then
                MaIdx = AktZa
                Exit For
            End If
        Next AktZa

        AktZa = 1
        GesZa = RS32.RecordCount
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            LLDeV RS31, DauDg, False, MxCod
        Else
            LLDeV RS31, DauDg
        End If
        LLDeF RS32
        
        If GlDub = True Then
            DevMo.dmDuplex = DMDUP_VERTICAL
        End If

        Select Case GlSGr
        Case 2: DevMo.dmPaperSize = DMPAPER_A4
        Case 3: DevMo.dmPaperSize = DMPAPER_A5
        Case 4: DevMo.dmPaperSize = DMPAPER_A6
        End Select
        
        If GlDub = True Or GlSGr > 1 Then
            RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
        End If

        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        End If
        
        If RetWe = 0 Then
            If AktPo = 1 Then
                FS.Show
                DoEvents
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        PrBr2.Value = AktZa
                        DoEvents
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                    
                    If DrVor = True Then 'Druckvorschau
                        With clFil
                            .DaCop = DruDa & ";" & FilNa & vbNullChar
                            If .FilCop(3) = False Then
                                SPopu "Datei schreibgeschützt", "Die Datei kann nicht umbenannt werden", IC48_Warning
                            Else
                                If RetWe = 0 Then DruOk = True
                            End If
                        End With
                    Else
                        If RetWe = 0 Then
                            DruOk = True
                        End If
                    End If
                                        
                End If
            End If
        End If
        If DrKop = False Then 'Druckkopie
            If EmSen > 0 Then 'Emailversand
                If GlDru.XRech = True Then
                    If ReStr <> vbNullString Then
                        XReDa = GlTmp & "XRechnung_Beleg_" & ReStr & ".xml"
                    Else
                        XReDa = GlTmp & "XRechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".xml"
                    End If
                    RS31.MoveFirst
                    RS32.MoveFirst
                    DoEvents
                    S_XRech RS31, RS32, XReDa
                    DoEvents
                End If
            End If
        End If
    Else
        If AktPo = 1 Then
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
            Exit For
        End If
    End If
    
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing

    If DrKop = False Then 'Druckkopie
        If EmSep = True Then 'Emailempfängerseparation
            If EmSen > 0 Then 'Emailversand
                If ReStr <> vbNullString Then
                    DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                Else
                    DaNam = GlTmp & "Rechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                End If
                If VerDo = True Then 'Dokument verschlüsseln
                    hwOrg = LlStgsysStorageOpen(FilNa, TmPfa, False, True)
                    RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                    LlStgsysStorageClose hwOrg
                Else
                    RetWe = LlStgsysStorageConvert(FilNa, DaNam, "PDF")
                End If
                If GlDru.XRech = True Then
                    SERech DaNam, XReDa
                    DoEvents
                    If ReStr <> vbNullString Then
                        DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                    Else
                        DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    End If
                End If

                If EmSen = 3 Then 'Emailversand
                    If EmEmp <> vbNullString Then
                        With GlTxV
                            .TxStr = GlEmT(1, 3)
                            .DaStr = ReDat
                            .Datum = ReDat
                            .ReStr = ReStr
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = PatNr
                        End With
                        MaBet = SEmTx()
                        With GlTxV
                            .TxStr = GlEmT(1, 1)
                            .Datum = ReDat
                            .DaStr = ReDat
                            .ReStr = ReStr
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = PatNr
                        End With
                        MaTex = vbCrLf & EmBrf & vbCrLf & vbCrLf & SEmTx() & vbCrLf & vbCrLf & EmSig
                        EmVer = SEmSe(EmEmp, MaBet, MaTex, DaNam, , EmNam, False, True)
                    End If
                ElseIf EmSen = 6 Then
                    DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                    DoEvents
                    If DoStr <> vbNullString Then
                        Lange = Len(DoStr)
                        Posit = InStr(1, DoStr, ";", 1)
                        If Posit > 0 Then
                            DoLnk = Left$(DoStr, Posit - 1)
                            PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                            With GlTxV
                                .TxStr = GlEmT(12, 3)
                                .Datum = ReDat
                                .DaStr = ReDat
                                .ReStr = ReStr
                                .MitNr = MitNr
                                .ManNr = ManNr
                                .PatNr = PatNr
                            End With
                            MaBet = SEmTx()
                            With GlTxV
                                .TxStr = GlEmT(12, 1)
                                .Datum = ReDat
                                .DaStr = ReDat
                                .ReStr = ReStr
                                .MitNr = MitNr
                                .ManNr = ManNr
                                .PatNr = PatNr
                                .DoLnk = DoLnk
                                .PoLnk = PoLnk
                                .DoLan = GlVrw
                            End With
                            MaTex = vbCrLf & EmBrf & vbCrLf & vbCrLf & SEmTx() & vbCrLf & vbCrLf & EmSig
                            EmVer = SEmSe(EmEmp, MaBet, MaTex, , , EmNam, False, True)
                        End If
                    End If
                End If
            End If
        End If
    End If

    DoEvents
    If TxDia.Text = "B" Then Exit For 'Abbrechen
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Set clFil = Nothing

If EmSep = False Then 'Emailempfängerseparation
    RetWe = LLStor(VorDa, GuiNa, AnzPo)
End If

Unload FS
DoEvents

If DruOk = True Then
    If DrSof = True Then
        hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
        SeiZa = LlStgsysGetPageCount(hwOrg)
        RetWe = LlStgsysPrint(hwOrg, DrNam, DrNam, 1, SeiZa, Kopie, LS_PRINTFLAG_FIT Or LS_PRINTFLAG_STACKEDCOPIES Or LS_PRINTFLAG_ABORTABLEMETER, "Formulardruck", FM.hwnd)
        LlStgsysStorageClose hwOrg
    Else
        If EmSep = False Then 'Emailempfängerseparation
            If EmSen > 0 Then 'Emailversand
                If GlDru.GoBDk = True Then
                    If GlDru.ReAbs = True Then
                        If ReStr = vbNullString Then
                            DaNam = ExpOr & "Rechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        Else
                            If AnzPo > 2 Then
                                DaNam = ExpOr & "Rechnungen_Belege_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                            Else
                                DaNam = ExpOr & "Rechnung_Beleg_" & ReStr & ".pdf"
                            End If
                        End If
                        If VerDo = True Then 'Dokument verschlüsseln
                            hwOrg = LlStgsysStorageOpen(ZusDa, TmPfa, False, True)
                            RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                            LlStgsysStorageClose hwOrg
                        Else
                            RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
                        End If
                        If GlDru.XRech = True Then
                            SERech DaNam, XReDa
                            DoEvents
                            If ReStr <> vbNullString Then
                                DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                            Else
                                DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                            End If
                        End If
                        
                        If EmSen < 3 Then
                            SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                        ElseIf EmSen = 6 Then
                            DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                            DoEvents
                            If DoStr <> vbNullString Then
                                Lange = Len(DoStr)
                                Posit = InStr(1, DoStr, ";", 1)
                                If Posit > 0 Then
                                    DoLnk = Left$(DoStr, Posit - 1)
                                    PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                                    With GlTxV
                                        .TxStr = GlEmT(12, 3)
                                        .Datum = ReDat
                                        .DaStr = ReDat
                                        .ReStr = ReStr
                                        .MitNr = MitNr
                                        .ManNr = ManNr
                                        .PatNr = PatNr
                                    End With
                                    MaBet = SEmTx()
                                    With GlTxV
                                        .TxStr = GlEmT(12, 1)
                                        .Datum = ReDat
                                        .DaStr = ReDat
                                        .ReStr = ReStr
                                        .MitNr = MitNr
                                        .ManNr = ManNr
                                        .PatNr = PatNr
                                        .DoLnk = DoLnk
                                        .PoLnk = PoLnk
                                        .DoLan = GlVrw
                                    End With
                                    MaTex = EmTex & SEmTx()
                                    SMaNe PatNr, EmAdr, , MaTex, MaBet
                                End If
                            End If
                        End If
                    Else
                        SPopu "Rechnung abschließen", "Die Rechnungsdruckoptionen stehen nicht auf abschließen!", IC48_Forbidden
                    End If
                Else
                    DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
                    If EmSen < 3 Then
                        SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                    End If
                End If
            Else
                If GlVTy = 2 Then
                    RetWe = LlPreviewDisplay(mHwnd, ZusDa, TmPfa, FM.hwnd)
                End If
            End If
        End If
    End If
Else
    SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If EmSen = 0 Then 'Emailversand
    If DrVor = True Then 'Druckvorschau
        If DruOk = True Then
            Select Case GlVTy
            Case 1: VoDrMa ZusDa
            Case 3: SDrVo ZusDa
            End Select
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLHeLs " & Err.Number
Resume Next

End Sub
Public Sub LLKoLi()
On Error GoTo DeErr
'Druckt eine kompakte Einzelrechung

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim AkDat As Date
Dim TaSum As Double 'Tagessumme
Dim AnSum As Single 'Anzahlsungssumme
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim MeStr As String 'Memostring
Dim Warun As String 'Währung
Dim AkZif As String 'Ziffer
Dim DauDg As String
Dim MxCod As String
Dim ReStr As String
Dim ReDat As String
Dim MaBet As String
Dim MaTex As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim DoLnk As String
Dim PoLnk As String
Dim DoStr As String
Dim XReDa As String
Dim EReDa As String
Dim ExpOr As String 'Exportordner
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim AktPo As Integer 'Aktuelle Position
Dim KriSt As String  'Kriterienstring
Dim AnzPo As Integer 'Gesamtpositionen
Dim TagZa As Integer 'Tageszahl
Dim ReZal As Integer 'Rechungszahl
Dim LegZa As Integer 'Legendenzahl
Dim LeiZa As Integer 'Leiszungszahl
Dim GeZei As Integer 'Gesamttageszeit
Dim AkPos As Integer
Dim Posit As Integer
Dim Lange As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2, SQL3 As String

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
AnzPo = UBound(GloDr)

If GlRDP = True Then
    If LCase(GlEPf) = LCase(GlIPf) Then
        ExpOr = GlDpf & "Import\"
    Else
        ExpOr = GlEPf
    End If
Else
    ExpOr = GlEPf
End If

Select Case ForNa
Case "RechKo":
    VerMo = True
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechKu":
    VerMo = True
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechFa":
    VerMo = True
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechSk":
    VerMo = True
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
End Select

If SQL2 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

DauDg = S_DrDi(GloDr(1))
If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    MxCod = S_ReMx(GloDr(1))
End If

Set RS31 = New ADODB.Recordset 'Rechnungskopf
RS31.CursorLocation = adUseClient
Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
ReZal = RS31.RecordCount

Set RS33 = New ADODB.Recordset 'Legende
RS33.CursorLocation = adUseClient
Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
LegZa = RS33.RecordCount

Set RS34 = New ADODB.Recordset 'Abrechungstage
RS34.CursorLocation = adUseClient
Set RS34 = DBCmRe1(SQL3, "@IdStr", GloDr(1))
TagZa = RS34.RecordCount

If ReZal > 0 Then
    If LegZa > 0 Then
        If TagZa > 0 Then
            StaPr = True
        End If
    End If
End If

If EmSen > 0 Then
    DrVor = True 'wenn Emailversand, dann Druckvorschau
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

If StaPr = True Then
    GesZa = LegZa + TagZa
    
    If VerMo = True Then
        If RS31.Fields("Em_Aut").Value <> vbNullString Then
            VerDo = CBool(RS31.Fields("Em_Aut").Value)
        Else
            VerDo = False
        End If
        If VerDo = True Then 'Dokument verschlüsseln
            If RS31.Fields("Em_Pass").Value <> vbNullString Then
                VerPa = RS31.Fields("Em_Pass").Value
            Else
                VerDo = False
            End If
        Else
            VerDo = False
        End If
    End If
    
    If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
        VerDo = False
    End If

    If IsNull(RS31.Fields("Rechnungsnummer").Value) = False Then
        If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
            If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                ReStr = RS31.Fields("Rechnungsnummer").Value
            End If
        End If
    End If
    
    If ReStr <> vbNullString Then
        If RS31.Fields("Rechnungsdatum").Value <> vbNullString Then
            ReDat = Format$(RS31.Fields("Rechnungsdatum").Value, "dd.mm.yyyy")
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        End If
        If RS31.Fields("IDP").Value > 0 Then
            ManNr = RS31.Fields("IDP").Value
        End If
        If RS31.Fields("IDM").Value > 0 Then
            MitNr = RS31.Fields("IDM").Value
        End If
    End If

    Set RS32 = New ADODB.Recordset
    With RS32
        .CursorLocation = adUseClient
        .Fields.Append "Datum", adDate
        .Fields.Append "Behandlungsdatum", adChar, 10
        .Fields.Append "Ziffer", adChar, 100
        .Fields.Append "Bezeichnung", adVariant
        .Fields.Append "Kommentar", adVariant
        .Fields.Append "Einzelbetrag", adDouble
        .Fields.Append "Gesamtbetrag", adDouble
        .Fields.Append "Anzahlungsbetrag", adDouble
        .Fields.Append "Währung", adChar, 10
        .Fields.Append "Anzahl", adInteger
        .Fields.Append "Multi", adDouble
        .Fields.Append "Typ", adInteger
        .Fields.Append "Zeit", adInteger
        .Fields.Append "ICD", adVariant
        .Fields.Append "Diagnose", adVariant
        .Fields.Append "Analog", adInteger
        If .State = adStateClosed Then .Open
    End With

    If RS32.Supports(adAddNew) Then
        Do
        AkDat = RS34.Fields("Datum").Value
        Set RS35 = New ADODB.Recordset
        RS35.CursorLocation = adUseClient
        If GlTyp < 2 Then
            Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(1), AkDat)
        Else
            Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(1), Format$(AkDat, "dd.mm.yyyy"))
        End If
        If RS35.RecordCount > 0 Then
            MeStr = vbNullString
            TaSum = 0
            AnSum = 0
            GeZei = 0
            Do
            If RS35.Fields("Ziffer").Value <> vbNullString Then
                AkZif = RS35.Fields("Ziffer").Value
                LeiZa = RS35.Fields("Anzahl").Value

                If MeStr = vbNullString Then
                    If LeiZa = 1 Then
                        MeStr = AkZif
                    Else
                        MeStr = LeiZa & "x" & AkZif
                    End If
                Else
                    If LeiZa = 1 Then
                        MeStr = MeStr & "; " & AkZif
                    Else
                        MeStr = MeStr & "; " & LeiZa & "x" & AkZif
                    End If
                End If
            End If
            If RS35.Fields("Währung").Value <> vbNullString Then
                Warun = RS35.Fields("Währung").Value
            Else
                Warun = "€"
            End If
            If RS35.Fields("Zeit").Value <> 0 Then
                GeZei = GeZei + RS35.Fields("Zeit").Value
            End If
            If RS35.Fields("Gesamtbetrag").Value <> 0 Then
                TaSum = TaSum + RS35.Fields("Gesamtbetrag").Value
            End If
            If RS35.Fields("Anzahlungsbetrag").Value <> 0 Then
                AnSum = AnSum + RS35.Fields("Anzahlungsbetrag").Value
            End If
            RS35.MoveNext
            Loop Until RS35.EOF
            RS32.AddNew
            RS32.Fields("Datum").Value = AkDat
            RS32.Fields("Behandlungsdatum").Value = Format$(AkDat, "dd.mm.yyyy")
            RS32.Fields("Bezeichnung").Value = MeStr
            RS32.Fields("Gesamtbetrag").Value = TaSum
            RS32.Fields("Anzahlungsbetrag").Value = AnSum
            RS32.Fields("Währung").Value = Warun
            RS32.Fields("Zeit").Value = GeZei
            RS32.Fields("Typ").Value = 0
        End If
        RS35.Close
        Set RS35 = Nothing
        RS34.MoveNext
        Loop Until RS34.EOF

        Do
        RS32.AddNew
        If RS33.Fields("Ziffer").Value <> vbNullString Then RS32.Fields("Ziffer").Value = RS33.Fields("Ziffer").Value
        If RS33.Fields("Bezeichnung").Value <> vbNullString Then RS32.Fields("Bezeichnung").Value = RS33.Fields("Bezeichnung").Value
        If RS33.Fields("ICD").Value <> vbNullString Then RS32.Fields("ICD").Value = RS33.Fields("ICD").Value
        If RS33.Fields("Diagnose").Value <> vbNullString Then RS32.Fields("Diagnose").Value = RS33.Fields("Diagnose").Value
        If RS33.Fields("Multi").Value <> vbNullString Then RS32.Fields("Multi").Value = RS33.Fields("Multi").Value
        RS32.Fields("Typ").Value = RS33.Fields("Typ").Value
        If RS33.Fields("Einzelbetrag").Value <> 0 Then
            RS32.Fields("Einzelbetrag").Value = RS33.Fields("Einzelbetrag").Value
            RS32.Fields("Währung").Value = RS33.Fields("Währung").Value
        End If
        If GlAnl = True Then
            If RS33.Fields("Analog").Value <> vbNullString Then
                RS32.Fields("Analog").Value = RS33.Fields("Analog").Value
            Else
                RS32.Fields("Analog").Value = 0
            End If
        Else
            RS32.Fields("Analog").Value = 0
        End If
        If RS33.Fields("Kommentar").Value <> vbNullString Then RS32.Fields("Kommentar").Value = RS33.Fields("Kommentar").Value
        RS33.MoveNext
        Loop Until RS33.EOF
        
        RS32.Update
        RS32.MoveFirst
        
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            LLDeV RS31, DauDg, False, MxCod
        Else
            LLDeV RS31, DauDg
        End If
        LLDeF RS32
        
        If GlDub = True Then
            DevMo.dmDuplex = DMDUP_VERTICAL
        End If
        
        Select Case GlSGr
        Case 2: DevMo.dmPaperSize = DMPAPER_A4
        Case 3: DevMo.dmPaperSize = DMPAPER_A5
        Case 4: DevMo.dmPaperSize = DMPAPER_A6
        End Select
        
        If GlDub = True Or GlSGr > 1 Then
            RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
        End If
        
        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
        End If
        
        If FaxNr <> vbNullString Then
            RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
        End If

        If RetWe = 0 Then
            If DrDia = True Then
                RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
            End If
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
    End If
    RS32.Close
    Set RS32 = Nothing
    
    If DrKop = False Then 'Druckkopie
        If EmSen > 0 Then 'Emailversand
            If GlDru.XRech = True Then
                If ReStr <> vbNullString Then
                    XReDa = GlTmp & "XRechnung_Beleg_" & ReStr & ".xml"
                Else
                    XReDa = GlTmp & "XRechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".xml"
                End If
                RS31.MoveFirst
                RS32.MoveFirst
                DoEvents
                S_XRech RS31, RS32, XReDa
                DoEvents
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS31.Close
RS33.Close
RS34.Close
Set RS31 = Nothing
Set RS33 = Nothing
Set RS34 = Nothing

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If GlDru.GoBDk = True Then
                If GlDru.ReAbs = True Then
                    If ReStr = vbNullString Then
                        DaNam = ExpOr & "Rechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    Else
                        If AnzPo > 2 Then
                            DaNam = ExpOr & "Rechnungen_Belege_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        Else
                            DaNam = ExpOr & "Rechnung_Beleg_" & ReStr & ".pdf"
                        End If
                    End If
                    If VerDo = True Then 'Dokument verschlüsseln
                        hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                        RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                        LlStgsysStorageClose hwOrg
                    Else
                        RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                    End If
                    If GlDru.XRech = True Then
                        SERech DaNam, XReDa
                        DoEvents
                        If ReStr <> vbNullString Then
                            DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                        Else
                            DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        End If
                    End If
                    If GlDru.XRech = True Then
                        SERech DaNam, XReDa
                        DoEvents
                        If ReStr <> vbNullString Then
                            DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                        Else
                            DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        End If
                    End If
                    
                    If EmSen < 3 Then
                        SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                    ElseIf EmSen = 6 Then
                        DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                        DoEvents
                        If DoStr <> vbNullString Then
                            Lange = Len(DoStr)
                            Posit = InStr(1, DoStr, ";", 1)
                            If Posit > 0 Then
                                DoLnk = Left$(DoStr, Posit - 1)
                                PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                                With GlTxV
                                    .TxStr = GlEmT(12, 3)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                End With
                                MaBet = SEmTx()
                                With GlTxV
                                    .TxStr = GlEmT(12, 1)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                    .DoLnk = DoLnk
                                    .PoLnk = PoLnk
                                    .DoLan = GlVrw
                                End With
                                MaTex = EmTex & SEmTx()
                                SMaNe PatNr, EmAdr, , MaTex, MaBet
                            End If
                        End If
                    End If
                Else
                    SPopu "Rechnung abschließen", "Die Rechnungsdruckoptionen stehen nicht auf abschließen!", IC48_Forbidden
                End If
            Else
                If ReStr <> vbNullString Then
                    DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                Else
                    DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                End If
                If VerDo = True Then 'Dokument verschlüsseln
                    hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                    RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                    LlStgsysStorageClose hwOrg
                Else
                    RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                End If
                If EmSen < 3 Then
                    SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                End If
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If ReStr <> vbNullString Then
                DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
            Else
                DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            End If
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLKoLi " & Err.Number
Resume Next

End Sub
Public Sub LLKoLs()
    On Error GoTo DeErr
'Druckt mehrere kompakte Rechungen

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim AkDat As Date
Dim TaSum As Double 'Tagessumme
Dim AnSum As Double 'Anzahlsungssumme
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim MeStr As String 'Memostring
Dim Warun As String 'Währung
Dim AkZif As String 'Ziffer
Dim ZusDa As String 'Zsammengefasste Datei
Dim DauDg As String
Dim MxCod As String
Dim EmEmp As String
Dim EmNam As String
Dim EmBrf As String
Dim ReStr As String
Dim ReDat As String
Dim EmSig As String
Dim FilNa As String
Dim VorDa As String
Dim GuiNa As String
Dim MaBet As String
Dim MaTex As String
Dim KriSt As String 'Kriterienstring
Dim VerPa As String 'Verschlüsselungspasswort
Dim DoLnk As String
Dim PoLnk As String
Dim DoStr As String
Dim XReDa As String
Dim EReDa As String
Dim ExpOr As String  'Exportordner
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim TagZa As Integer 'Tageszahl
Dim ReZal As Integer 'Rechungszahl
Dim LegZa As Integer 'Legendenzahl
Dim LeiZa As Integer 'Leiszungszahl
Dim GeZei As Integer 'Gesamttageszeit
Dim AkPos As Integer
Dim SeiZa As Integer
Dim MaIdx As Integer
Dim Posit As Integer
Dim Lange As Integer
Dim EmVer As Boolean 'Email erfolgreich versandt
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim DevMo As DEVMODEA
Dim SQL1, SQL2, SQL3 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
VorDa = GuiNa & "_" & "0001.LL"
ZusDa = TmPfa & VorDa
AnzPo = UBound(GloDr)

If GlRDP = True Then
    If LCase(GlEPf) = LCase(GlIPf) Then
        ExpOr = GlDpf & "Import\"
    Else
        ExpOr = GlEPf
    End If
Else
    ExpOr = GlEPf
End If

Select Case ForNa
Case "RechKo":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechKu":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechFa":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
Case "RechSk":
    SQL1 = "qryPrSimRe1"
    SQL2 = "qryPrSimRe5"
    SQL3 = "qryPrSimRe4"
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

For AktPo = 1 To AnzPo
    FilNa = TmPfa & GuiNa & "_" & Format$(AktPo, "0000") & ".LL"
    DauDg = vbNullString
    MxCod = vbNullString
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    DrKop = GloKo(AktPo)
    DauDg = S_DrDi(GloDr(AktPo))
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        MxCod = S_ReMx(GloDr(AktPo))
    End If

    Set RS31 = New ADODB.Recordset 'Rechnungskopf
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
    ReZal = RS31.RecordCount

    Set RS33 = New ADODB.Recordset 'Legende
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
    LegZa = RS33.RecordCount
    
    Set RS34 = New ADODB.Recordset 'Abrechungstage
    RS34.CursorLocation = adUseClient
    Set RS34 = DBCmRe1(SQL3, "@IdStr", GloDr(AktPo))
    TagZa = RS34.RecordCount

    If ReZal > 0 Then
        If LegZa > 0 Then
            If TagZa > 0 Then
                StaPr = True
            End If
        End If
    End If

    If StaPr = True Then
        If RS31.Fields("Em_Aut").Value <> vbNullString Then
            VerDo = CBool(RS31.Fields("Em_Aut").Value)
        Else
            VerDo = False
        End If
        If VerDo = True Then 'Dokument verschlüsseln
            If RS31.Fields("Em_Pass").Value <> vbNullString Then
                VerPa = RS31.Fields("Em_Pass").Value
            Else
                VerDo = False
            End If
        Else
            VerDo = False
        End If
        
        If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
            VerDo = False
        End If

        If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
            ReStr = RS31.Fields("Rechnungsnummer").Value
        Else
            ReStr = Format$(AktPo, "0000")
        End If
        If RS31.Fields("Rechnungsdatum").Value <> vbNullString Then
            ReDat = Format$(RS31.Fields("Rechnungsdatum").Value, "dd.mm.yyyy")
        Else
            ReDat = Format$(Date, "dd.mm.yyyy")
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        Else
            PatNr = 0
        End If
        EmNam = RS31.Fields("R_Name").Value & ", " & RS31.Fields("R_Vorname").Value
        If RS31.Fields("Telefon5").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon5").Value
        ElseIf RS31.Fields("Telefon6").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon6").Value
        Else
            EmEmp = vbNullString
        End If
        If RS31.Fields("R_Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("R_Briefanrede").Value
        ElseIf RS31.Fields("Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("Briefanrede").Value
        Else
            EmBrf = vbNullString
        End If
        If RS31.Fields("IDP").Value <> vbNullString Then
            ManNr = RS31.Fields("IDP").Value
        Else
            ManNr = 0
        End If
        If RS31.Fields("IDM").Value <> vbNullString Then
            MitNr = RS31.Fields("IDM").Value
        Else
            MitNr = 0
        End If
        If GlMiA(GlSmI, 11) <> vbNullString Then
            EmSig = GlMiA(GlSmI, 11)
        Else
            EmSig = GlMiA(GlSmI, 1)
        End If
        
        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then
                MaIdx = AktZa
                Exit For
            End If
        Next AktZa
    
        AktZa = 1
        GesZa = LegZa + TagZa
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        
        Set RS32 = New ADODB.Recordset
        With RS32
            .CursorLocation = adUseClient
            .Fields.Append "Datum", adDate
            .Fields.Append "Behandlungsdatum", adChar, 10
            .Fields.Append "Ziffer", adChar, 100
            .Fields.Append "Bezeichnung", adVariant
            .Fields.Append "Kommentar", adVariant
            .Fields.Append "Einzelbetrag", adDouble
            .Fields.Append "Gesamtbetrag", adDouble
            .Fields.Append "Anzahlungsbetrag", adDouble
            .Fields.Append "Währung", adChar, 10
            .Fields.Append "Anzahl", adInteger
            .Fields.Append "Multi", adDouble
            .Fields.Append "Typ", adInteger
            .Fields.Append "Zeit", adInteger
            .Fields.Append "ICD", adVariant
            .Fields.Append "Diagnose", adVariant
            .Fields.Append "Analog", adInteger
            .Open
        End With

        If RS32.Supports(adAddNew) Then
            Do
            If RS34.Fields("Datum").Value <> vbNullString Then
                AkDat = RS34.Fields("Datum").Value
                Set RS35 = New ADODB.Recordset
                RS35.CursorLocation = adUseClient
                If GlTyp < 2 Then
                    Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(AktPo), AkDat)
                Else
                    Set RS35 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", GloDr(AktPo), Format$(AkDat, "dd.mm.yyyy"))
                End If
                If RS35.RecordCount > 0 Then
                    MeStr = vbNullString
                    TaSum = 0
                    AnSum = 0
                    GeZei = 0
                    Do
                    If RS35.Fields("Ziffer").Value <> vbNullString Then
                        AkZif = RS35.Fields("Ziffer").Value
                        LeiZa = RS35.Fields("Anzahl").Value
        
                        If MeStr = vbNullString Then
                            If LeiZa = 1 Then
                                MeStr = AkZif
                            Else
                                MeStr = LeiZa & "x" & AkZif
                            End If
                        Else
                            If LeiZa = 1 Then
                                MeStr = MeStr & "; " & AkZif
                            Else
                                MeStr = MeStr & "; " & LeiZa & "x" & AkZif
                            End If
                        End If
                    End If
                    If RS35.Fields("Währung").Value <> vbNullString Then
                        Warun = RS35.Fields("Währung").Value
                    Else
                        Warun = "€"
                    End If
                    If RS35.Fields("Zeit").Value <> 0 Then
                        GeZei = GeZei + RS35.Fields("Zeit").Value
                    End If
                    If RS35.Fields("Gesamtbetrag").Value <> 0 Then
                        TaSum = TaSum + RS35.Fields("Gesamtbetrag").Value
                    End If
                    If RS35.Fields("Anzahlungsbetrag").Value <> 0 Then
                        AnSum = AnSum + RS35.Fields("Anzahlungsbetrag").Value
                    End If
                    RS35.MoveNext
                    Loop Until RS35.EOF
                    RS32.AddNew
                    RS32.Fields("Datum").Value = AkDat
                    RS32.Fields("Behandlungsdatum").Value = Format$(AkDat, "dd.mm.yyyy")
                    RS32.Fields("Bezeichnung").Value = MeStr
                    RS32.Fields("Gesamtbetrag").Value = TaSum
                    RS32.Fields("Anzahlungsbetrag").Value = AnSum
                    RS32.Fields("Währung").Value = Warun
                    RS32.Fields("Zeit").Value = GeZei
                    RS32.Fields("Typ").Value = 0
                End If
                RS35.Close
                Set RS35 = Nothing
            End If
            RS34.MoveNext
            Loop Until RS34.EOF

            Do
            RS32.AddNew
            If RS33.Fields("Ziffer").Value <> vbNullString Then RS32.Fields("Ziffer").Value = RS33.Fields("Ziffer").Value
            If RS33.Fields("Bezeichnung").Value <> vbNullString Then RS32.Fields("Bezeichnung").Value = RS33.Fields("Bezeichnung").Value
            If RS33.Fields("ICD").Value <> vbNullString Then RS32.Fields("ICD").Value = RS33.Fields("ICD").Value
            If RS33.Fields("Diagnose").Value <> vbNullString Then RS32.Fields("Diagnose").Value = RS33.Fields("Diagnose").Value
            If RS33.Fields("Multi").Value <> vbNullString Then RS32.Fields("Multi").Value = RS33.Fields("Multi").Value
            RS32.Fields("Typ").Value = RS33.Fields("Typ").Value
            If RS33.Fields("Einzelbetrag").Value <> 0 Then
                RS32.Fields("Einzelbetrag").Value = RS33.Fields("Einzelbetrag").Value
                RS32.Fields("Währung").Value = RS33.Fields("Währung").Value
            End If
            If GlAnl = True Then
                If RS33.Fields("Analog").Value <> vbNullString Then
                    RS32.Fields("Analog").Value = RS33.Fields("Analog").Value
                Else
                    RS32.Fields("Analog").Value = 0
                End If
            Else
                RS32.Fields("Analog").Value = 0
            End If
            If RS33.Fields("Kommentar").Value <> vbNullString Then RS32.Fields("Kommentar").Value = RS33.Fields("Kommentar").Value
            RS33.MoveNext
            Loop Until RS33.EOF
            
            RS32.Update
            RS32.MoveFirst
            
            If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
                LLDeV RS31, DauDg, False, MxCod
            Else
                LLDeV RS31, DauDg
            End If
            LLDeF RS32
            
            If GlDub = True Then
                DevMo.dmDuplex = DMDUP_VERTICAL
            End If
            
            Select Case GlSGr
            Case 2: DevMo.dmPaperSize = DMPAPER_A4
            Case 3: DevMo.dmPaperSize = DMPAPER_A5
            Case 4: DevMo.dmPaperSize = DMPAPER_A6
            End Select
            
            If GlDub = True Or GlSGr > 1 Then
                RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
            End If
            
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                If DrSof = False Then
                    RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
                End If
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If

            If FaxNr <> vbNullString Then
                RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
            End If

            If RetWe = 0 Then
                If AktPo = 1 Then
                    FS.Show
                    DoEvents
                End If
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                    If RetWe = 0 Then
                        While Not RS32.EOF
                            LLDeF RS32
                            PrBr2.Value = AktZa
                            DoEvents
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                            RS32.MoveNext
                        Wend
                        Do
                        RetWe = LlPrintFieldsEnd(mHwnd)
                        Loop Until RetWe <> LL_WRN_REPEAT_DATA
                        RetWe = LlPrintEnd(mHwnd, 0)
                        If RetWe = 0 Then DruOk = True
                                                                        
                        If DrVor = True Then 'Druckvorschau
                            With clFil
                                .DaCop = DruDa & ";" & FilNa & vbNullChar
                                If .FilCop(3) = False Then
                                    SPopu "Datei schreibgeschützt", "Die Datei kann nicht umbenannt werden", IC48_Warning
                                Else
                                    If RetWe = 0 Then DruOk = True
                                End If
                            End With
                        Else
                            If RetWe = 0 Then DruOk = True
                        End If
                    
                    End If
                End If
            End If
            If DrKop = False Then 'Druckkopie
                If EmSen > 0 Then 'Emailversand
                    If GlDru.XRech = True Then
                        If ReStr <> vbNullString Then
                            XReDa = GlTmp & "XRechnung_Beleg_" & ReStr & ".xml"
                        Else
                            XReDa = GlTmp & "XRechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".xml"
                        End If
                        RS31.MoveFirst
                        RS32.MoveFirst
                        DoEvents
                        S_XRech RS31, RS32, XReDa
                        DoEvents
                    End If
                End If
            End If
        Else
            If AktPo = 1 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
                Exit For
            End If
        End If

        RS31.Close
        RS32.Close
        RS33.Close
        RS34.Close
        Set RS31 = Nothing
        Set RS32 = Nothing
        Set RS33 = Nothing
        Set RS34 = Nothing

        If DrKop = False Then 'Druckkopie
            If EmSep = True Then 'Emailempfängerseparation
                If EmSen > 0 Then 'Emailversand
                    If ReStr <> vbNullString Then
                        DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                    Else
                        DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    End If
                    If VerDo = True Then 'Dokument verschlüsseln
                        hwOrg = LlStgsysStorageOpen(FilNa, TmPfa, False, True)
                        RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                        LlStgsysStorageClose hwOrg
                    Else
                        RetWe = LlStgsysStorageConvert(FilNa, DaNam, "PDF")
                    End If
                    If GlDru.XRech = True Then
                        SERech DaNam, XReDa
                        DoEvents
                        If ReStr <> vbNullString Then
                            DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                        Else
                            DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        End If
                    End If

                    If EmSen = 3 Then 'Emailversand
                        If EmEmp <> vbNullString Then
                            With GlTxV
                                .TxStr = GlEmT(1, 3)
                                .Datum = ReDat
                                .DaStr = ReDat
                                .ReStr = ReStr
                                .MitNr = MitNr
                                .ManNr = ManNr
                                .PatNr = PatNr
                            End With
                            MaBet = SEmTx()
                            With GlTxV
                                .TxStr = GlEmT(1, 1)
                                .Datum = ReDat
                                .DaStr = ReDat
                                .ReStr = ReStr
                                .MitNr = MitNr
                                .ManNr = ManNr
                                .PatNr = PatNr
                                .DoLnk = DoLnk
                                .PoLnk = PoLnk
                                .DoLan = GlVrw
                            End With
                            MaTex = vbCrLf & EmBrf & vbCrLf & vbCrLf & SEmTx() & vbCrLf & vbCrLf & EmSig
                            EmVer = SEmSe(EmEmp, MaBet, MaTex, DaNam, , EmNam, False, True)
                        End If
                    ElseIf EmSen = 6 Then
                        DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                        DoEvents
                        If DoStr <> vbNullString Then
                            Lange = Len(DoStr)
                            Posit = InStr(1, DoStr, ";", 1)
                            If Posit > 0 Then
                                DoLnk = Left$(DoStr, Posit - 1)
                                PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                                With GlTxV
                                    .TxStr = GlEmT(12, 3)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                End With
                                MaBet = SEmTx()
                                With GlTxV
                                    .TxStr = GlEmT(12, 1)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                    .DoLnk = DoLnk
                                    .PoLnk = PoLnk
                                    .DoLan = GlVrw
                                End With
                                MaTex = vbCrLf & EmBrf & vbCrLf & vbCrLf & SEmTx() & vbCrLf & vbCrLf & EmSig
                                EmVer = SEmSe(EmEmp, MaBet, MaTex, , , EmNam, False, True)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
        DoEvents
        If TxDia.Text = "B" Then Exit For 'Abbrechen
        PrBr1.Value = AktPo
    End If
    StaPr = False
Next AktPo

Set clFil = Nothing

If EmSep = False Then 'Emailempfängerseparation
    RetWe = LLStor(VorDa, GuiNa, AnzPo)
End If

Unload FS
DoEvents

If DruOk = True Then
    If DrSof = True Then
        hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
        SeiZa = LlStgsysGetPageCount(hwOrg)
        RetWe = LlStgsysPrint(hwOrg, DrNam, DrNam, 1, SeiZa, Kopie, LS_PRINTFLAG_FIT Or LS_PRINTFLAG_STACKEDCOPIES Or LS_PRINTFLAG_ABORTABLEMETER, "Formulardruck", FM.hwnd)
        LlStgsysStorageClose hwOrg
    Else
        If EmSep = False Then 'Emailempfängerseparation
            If EmSen > 0 Then 'Emailversand
                If GlDru.GoBDk = True Then
                    If GlDru.ReAbs = True Then
                        If ReStr = vbNullString Then
                            DaNam = ExpOr & "Rechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        Else
                            If AnzPo > 2 Then
                                DaNam = ExpOr & "Rechnungen_Belege_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                            Else
                                DaNam = ExpOr & "Rechnung_Beleg_" & ReStr & ".pdf"
                            End If
                        End If
                        If VerDo = True Then 'Dokument verschlüsseln
                            hwOrg = LlStgsysStorageOpen(ZusDa, TmPfa, False, True)
                            RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                            LlStgsysStorageClose hwOrg
                        Else
                            RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
                        End If
                        If GlDru.XRech = True Then
                            SERech DaNam, XReDa
                            DoEvents
                            If ReStr <> vbNullString Then
                                DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                            Else
                                DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                            End If
                        End If
                        
                        If EmSen < 3 Then
                            SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                        ElseIf EmSen = 6 Then
                            DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                            DoEvents
                            If DoStr <> vbNullString Then
                                Lange = Len(DoStr)
                                Posit = InStr(1, DoStr, ";", 1)
                                If Posit > 0 Then
                                    DoLnk = Left$(DoStr, Posit - 1)
                                    PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                                    With GlTxV
                                        .TxStr = GlEmT(12, 3)
                                        .Datum = ReDat
                                        .DaStr = ReDat
                                        .ReStr = ReStr
                                        .MitNr = MitNr
                                        .ManNr = ManNr
                                        .PatNr = PatNr
                                    End With
                                    MaBet = SEmTx()
                                    With GlTxV
                                        .TxStr = GlEmT(12, 1)
                                        .Datum = ReDat
                                        .DaStr = ReDat
                                        .ReStr = ReStr
                                        .MitNr = MitNr
                                        .ManNr = ManNr
                                        .PatNr = PatNr
                                        .DoLnk = DoLnk
                                        .PoLnk = PoLnk
                                        .DoLan = GlVrw
                                    End With
                                    MaTex = EmTex & SEmTx()
                                    SMaNe PatNr, EmAdr, , MaTex, MaBet
                                End If
                            End If
                        End If
                    Else
                        SPopu "Rechnung abschließen", "Die Rechnungsdruckoptionen stehen nicht auf abschließen!", IC48_Forbidden
                    End If
                Else
                    DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".pdf"
                    If VerDo = True Then 'Dokument verschlüsseln
                        hwOrg = LlStgsysStorageOpen(ZusDa, TmPfa, False, True)
                        RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                        LlStgsysStorageClose hwOrg
                    Else
                        RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
                    End If
                    If EmSen < 3 Then
                        SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                    End If
                End If
            Else
                If GlVTy = 2 Then
                    RetWe = LlPreviewDisplay(mHwnd, ZusDa, TmPfa, FM.hwnd)
                End If
            End If
        End If
    End If
Else
    SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If EmSen = 0 Then 'Emailversand
    If DrVor = True Then 'Druckvorschau
        If DruOk = True Then
            Select Case GlVTy
            Case 1: VoDrMa ZusDa
            Case 3: SDrVo ZusDa
            End Select
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLKoLs " & Err.Number
Resume Next

End Sub
Private Sub LLMail(Optional ByVal VerDo As Boolean, Optional ByVal VerPa As String)
On Error GoTo DeErr

Dim RetWe As Long

If StVer <> vbNullString Then
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", StVer)
Else
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", TmPfa)
End If

RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FontMode", 5)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.CompressStreamMethod", 1)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnablePrinting", 1)
If VerDo = True Then 'Dokument verschlüsseln
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.Level", 1)
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EncryptFile", 1)
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableChanging", 0)
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Encryption.EnableCopying", 0)
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", VerPa)
End If
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Quiet", 1)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.ShowResult", 0)
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.Format", "PDF")
RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.Provider", "MSMAPI")
If EmAdr <> vbNullString Then RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.To", EmAdr)
If EmBet <> vbNullString Then RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.Subject", EmBet)
If EmTex <> vbNullString Then RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.Body", EmTex)

If EmDia = True Then
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.ShowDialog", 1)
Else
    RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Mail.ShowDialog", 0)
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLMail " & Err.Number
Resume Next

End Sub
Private Sub LLOpt()
On Error GoTo DeErr

Dim RetWe As Long

LlSetDlgboxMode LL_DLGBOXMODE_ALT9 Or LL_DLGBOXMODE_NOBITMAPS

RetWe = LlSetOption(mHwnd, LL_OPTION_ADDVARSTOFIELDS, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_ALLOW_LLX_EXPORTERS, -1)
RetWe = LlSetOptionString(mHwnd, LL_OPTIONSTR_DEFAULTIMAGEPATH, GlIPf) 'WICHTIG!
RetWe = LlSetOptionString(mHwnd, LL_OPTIONSTR_LICENSINGINFO, "55htDg") 'WICHTIG!
If GlRDP = True Then
    If DrVor = True Then 'Druckvorschau
        RetWe = LlSetOptionString(mHwnd, LL_OPTIONSTR_EXPORTS_ALLOWED, "PRN;PRV")
    End If
End If
RetWe = LlSetOption(mHwnd, LL_OPTION_DEBUGFLAG, 0) 'Debugmodus
RetWe = LlSetOption(mHwnd, LL_OPTION_INCREMENTAL_PREVIEW, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_AUTOMULTIPAGE, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_CALCSUMVARSONINVISIBLELINES, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_COMPRESSRTF, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_COMPRESSSTORAGE, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_CONVERTCRLF, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_DELAYTABLEHEADER, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_EMFRESOLUTION, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_EXTENDEDEVALUATION, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_HELPAVAILABLE, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_IMMEDIATELASTPAGE, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_INTERCHARSPACING, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_INCLUDEFONTDESCENT, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_MULTIPLETABLELINES, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_NEWEXPRESSIONS, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_NOFILEVERSIONUPGRADEWARNING, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_NOFOOTERPAGEWRAP, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_NONOTABLECHECK, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_NOPARAMETERCHECK, 1)
RetWe = LlSetOption(mHwnd, LL_OPTION_OFNDIALOG_EXPLORER, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_OFNDIALOG_NOPLACESBAR, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_ONLYONETABLE, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_PRVZOOM_PERC, 150)
RetWe = LlSetOption(mHwnd, LL_OPTION_REALTIME, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_RESOLUTIONCOMPATIBLETO9X, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_SCALABLEFONTSONLY, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_SETCREATIONINFO, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_SHOWPREDEFVARS, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_SKIPRETURNATENDOFRTF, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_SORTVARIABLES, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_SPACEOPTIMIZATION, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_STORAGESYSTEM, LL_STG_STORAGE)
RetWe = LlSetOption(mHwnd, LL_OPTION_SUPERVISOR, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_SUPPORTPAGEBREAK, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_SUPPORTS_PRNOPTSTR_EXPORT, 0) 'Testen
RetWe = LlSetOption(mHwnd, LL_OPTION_TABLE_COLORING, LL_COLORING_LL)
RetWe = LlSetOption(mHwnd, LL_OPTION_TABSTOPS, LL_TABS_DELETE)
RetWe = LlSetOption(mHwnd, LL_OPTION_UISTYLE, LL_OPTION_UISTYLE_OFFICE2003)
RetWe = LlSetOption(mHwnd, LL_OPTION_USEBARCODESIZES, -1) 'WICHTIG!
RetWe = LlSetOption(mHwnd, LL_OPTION_USECHARTFIELDS, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_USEHOSTPRINTER, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_VARSCASESENSITIVE, 0)
RetWe = LlSetOption(mHwnd, LL_OPTION_XLATVARNAMES, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_RIBBON_DEFAULT_ENABLEDSTATE, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_DESIGNER_RIBBONBACKGROUNDCOLOR, vbWhite)
RetWe = LlSetOption(mHwnd, LL_OPTION_INTERACTIVESORTING_REALDATAJOBPARAMETER, -1)
RetWe = LlSetOption(mHwnd, LL_OPTION_RESETPROJECTSTATE_FORCES_NEW_PRINTJOB, GlDrS)

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLOpt " & Err.Number
Resume Next

End Sub
Public Sub LLPrKa()
On Error GoTo DeErr
'Druckt einen Einzelauftrag für den Katalogexplorer

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim AkPos As Integer
Dim KriSt As String 'Kriterienstring
Dim SQL1 As String
Dim DevMo As DEVMODEA

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"

Select Case ForNa
Case "KatLis":
        Select Case Left$(GlNod, 1)
        Case "A": SQL1 = "SELECT * FROM qryKat01C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Gebührenkataloge
        Case "C": SQL1 = "SELECT * FROM qryKat03C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Diagnosekatalog
        Case "G": SQL1 = "SELECT * FROM qryKat02C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Laborparameter
        Case "I": SQL1 = "SELECT * FROM qryKat04C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Arzneikatalog
        Case "K": SQL1 = "SELECT * FROM qryKat02C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Begründungen
        Case "L": SQL1 = "SELECT * FROM qryKat02C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Anamnesetexte
        Case "M": Exit Sub                                                                          'Terminbetreffs
        Case "N": SQL1 = "SELECT * FROM qryKat08B WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Anamnesetebogen
        Case "O": SQL1 = "SELECT * FROM qryKat04C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Textphrasenkatalog
        Case "P": SQL1 = "SELECT * FROM qryKat12C WHERE [ID3] = " & IdxNr & " ORDER BY [Sorter];"   'Artikelkatalog
        End Select
Case "KetLis":
        Select Case Left$(GlNod, 1)
        Case "D": SQL1 = "SELECT * FROM qryKet01 WHERE [ID1] = " & IdxNr & ";" 'Gebührenketten
        Case "F": SQL1 = "SELECT * FROM qryKet03 WHERE [ID1] = " & IdxNr & ";" 'Diagnoseketten
        Case "H": SQL1 = "SELECT * FROM qryKet02 WHERE [ID1] = " & IdxNr & ";" 'Laborprofile
        Case "J": SQL1 = "SELECT * FROM qryKet04 WHERE [ID1] = " & IdxNr & ";" 'Arzneiketten
        Case "R": Exit Sub                                                     'Terminketten
        Case "Q": SQL1 = "SELECT * FROM qryKet12 WHERE [ID1] = " & IdxNr & ";" 'Artikelketten
        End Select
Case "InvLis":
        Select Case Left$(GlNod, 1)
        Case "P": SQL1 = "SELECT * FROM qryKat12C WHERE [ID3]=" & IdxNr & ";"   'Artikelkatalog
        End Select
Case "BesLis":
        Select Case Left$(GlNod, 1)
        Case "P": SQL1 = "SELECT * FROM qryKat12C WHERE [ID3]=" & IdxNr & ";"   'Artikelkatalog
        End Select
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

Set RS32 = New ADODB.Recordset
With RS32
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
If RS32.RecordCount > 0 Then
    StaPr = True
End If

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

If StaPr = True Then
    GesZa = RS32.RecordCount
    LLDeF RS32
    If DrVor = True Then 'Druckvorschau
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        If GlRDP = True Then
            RetWe = LlViewerProhibitAction(mHwnd, 116)
        End If
        RetWe = LlViewerProhibitAction(mHwnd, 115)
        If DrDia = True Then
            RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
        Else
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        End If
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
    Else
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        If DrDia = True Then
            RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
        Else
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        End If
    End If
    If RetWe = 0 Then
        If DrDia = True Then
            RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        End If
        If RetWe = 0 Then
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS32.Close
Set RS32 = Nothing

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then 'Druckvorschautyp
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then 'Druckvorschautyp
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrKa " & Err.Number
Resume Next

End Sub
Public Sub LLPrLa()
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamtpositionen
Dim AkWie As Integer
Dim AkPos As Integer
Dim SQL1 As String

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
DaExt = Right$(FiNam, 3)

If GlTyp < 2 Then
    Select Case ForNa
    Case "AdrEti": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "DiaEti": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "KarEt1": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "KarEt2": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "PhoEti": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "StaEt2": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "VerEti": SQL1 = "SELECT * FROM dbo.qryPrSimAd1 ORDER BY Kurzbezeichnung"
    Case "LabBef": SQL1 = "SELECT * FROM dbo.qryPrLaBef"
    End Select
Else
    Select Case ForNa
    Case "AdrEti": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "DiaEti": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "KarEt1": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "KarEt2": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "PhoEti": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "StaEt2": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "VerEti": SQL1 = "SELECT * FROM qryPrSimAd1 ORDER BY [Kurzbezeichnung];"
    Case "LabBef": SQL1 = "SELECT * FROM qryPrLaBef;"
    End Select
End If

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

Set RS31 = New ADODB.Recordset
With RS31
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With

GesZa = RS31.RecordCount

If GesZa > 0 Then
    StaPr = True
End If

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

If StaPr = True Then
    LLDeV RS31
    Select Case LCase(DaExt):
    Case "lbl": 'Nur Variablen
            LLDeV RS31
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_NORMAL, 0)
                End If
            End If
    Case "crd": 'Nur Variablen
            LLDeV RS31
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
                End If
            End If
    End Select
    If RetWe = 0 Then
        If DrDia = True Then
            RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        End If
        If RetWe = 0 Then
            RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
            If RetWe = 0 Then
                If Wiede = True Then
                    For AkWie = 1 To WieAn
                        LLDeV RS31
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AkWie / WieAn)))
                        While LlPrint(mHwnd) = LL_WRN_REPEAT_DATA
                        Wend
                    Next AkWie
                Else
                    While Not RS31.EOF
                        LLDeV RS31
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                        While LlPrint(mHwnd) = LL_WRN_REPEAT_DATA
                        Wend
                        AktZa = AktZa + 1
                        RS31.MoveNext
                    Wend
                End If
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS31.Close
Set RS31 = Nothing

Select Case ForNa
Case "LabBef": DBCmEx0 "qryLabBeSe"
Case "UberEU": DBCmEx0 "qrySimReDru"
End Select

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then 'Druckvorschautyp
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrLa " & Err.Number
Resume Next

End Sub
Public Sub LLPrLi()
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim GuiNa As String
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim DauDg As String
Dim MxCod As String
Dim EmEmp As String
Dim EmNam As String
Dim EmBrf As String
Dim ReStr As String
Dim ReDat As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim DoLnk As String
Dim PoLnk As String
Dim DoStr As String
Dim XReDa As String
Dim EReDa As String
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim CmTex As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim AkPos As Integer
Dim Gefun As Integer
Dim Posit As Integer
Dim Lange As Integer
Dim KriSt As String  'Kriterienstring
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain

AktZa = 1
GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
DaExt = Right$(FiNam, 3)
ExFmt = "PDF"

Select Case ForNa
Case "EMRAbr":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2a"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Rechnu":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Auslan":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "RechLa":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "RechAl":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Besche":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2b"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "RechSt":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "StAnge":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "StGuts":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Uberwe":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe1"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "UberEU":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe1"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Angebo":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Gutsch":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "Liefer":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
        DauDg = S_DrDi(GloDr(1))
        If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
            MxCod = S_ReMx(GloDr(1))
        End If
Case "TagPro":
        SQL1 = "qrySimAbKra07"
        SQL2 = Krite 'Tagesprotokoll
        CmTex = True
Case "KraBla":
        VerMo = True
        SQL1 = "qryPrSimKr1"
        SQL2 = Krite
        CmTex = True
        DauDg = S_KrDp(GloDr(1))
Case "EiMahn": 'Einzelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
Case "SaMahn": 'Sammelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP4"
        SQL2 = "qryPrSimOP5"
Case "StMahn": 'Einzelmahnung (Alternativ)
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
Case "LabKom":
        VerMo = True
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
Case "LabAuf":
        VerMo = True
        SQL1 = "qryPrLaAu1"
        SQL2 = "qryPrLaAu2"
Case "PatUms":
        SQL1 = "qryPrSimTaPa3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimTaPa1 WHERE " & Krite & " ORDER BY Patient, Sorter"
        Else
            SQL2 = "SELECT * FROM qryPrSimTaPa1 WHERE " & Krite & " ORDER BY [Patient], [Datum];"
        End If
        CmTex = True
Case "TagUms":
        SQL1 = "qryPrSimTag1"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimTag1 WHERE " & Krite & " ORDER BY Sorter, Patient"
        Else
            SQL2 = "SELECT * FROM qryPrSimTag1 WHERE " & Krite & " ORDER BY [Sortierung], [Patient];"
        End If
        CmTex = True
Case "LabUms":
        SQL1 = "qryPrSimTag5"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimTag5 WHERE " & Krite & " ORDER BY Sorter, Patient"
        Else
            SQL2 = "SELECT * FROM qryPrSimTag5 WHERE " & Krite & " ORDER BY [Sortierung], [Patient];"
        End If
        CmTex = True
Case "StaDia": 'Diagnosestatistik
        SQL1 = "qryPrSimDia1"
        Select Case ZeRau
        Case 1: 'Monat
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimDia1a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimDia1a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        Case 2: 'Jahr
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimDia2a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimDia2a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        Case 3: 'Qurtal
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimDia3a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimDia3a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        End Select
        CmTex = True
Case "StaGeb":
        SQL1 = "qryPrSimGeb1"
        Select Case ZeRau
        Case 1: 'Monat
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimGeb1a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimGeb1a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        Case 2: 'Jahr
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimGeb2a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimGeb2a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        Case 3: 'Qurtal
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrSimGeb3a WHERE " & Krite & " ORDER BY Anzahl DESC"
            Else
                SQL2 = "SELECT * FROM qryPrSimGeb3a WHERE " & Krite & " ORDER BY [Anzahl] DESC;"
            End If
        End Select
        CmTex = True
Case "TagSum": 'Tagessummen
        SQL1 = "qryPrSimTag2"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimTag2 WHERE " & Krite & " ORDER BY Sorter, Mandant"
        Else
            SQL2 = "SELECT * FROM qryPrSimTag2 WHERE " & Krite & " ORDER BY [Sortierung], [Mandant];"
        End If
        CmTex = True
Case "RecSum":
        SQL1 = "qryPrSimTag3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimTag3 WHERE " & Krite & " ORDER BY Sorter, IDR, Gedruckt"
        Else
            SQL2 = "SELECT * FROM qryPrSimTag3 WHERE " & Krite & " ORDER BY [Sortierung], [IDR], [Gedruckt];"
        End If
        CmTex = True
Case "EiUbMo": 'Einnahmenüberschussrechnung
        SQL1 = "qryPrBuEr3"
        Select Case ZeRau
        Case 1: 'Monat
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrBuEr1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto]"
            Else
                SQL2 = "SELECT * FROM qryPrBuEr1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto];"
            End If
        Case 2: 'Jahr
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrBuEr2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Konto]"
            Else
                SQL2 = "SELECT * FROM qryPrBuEr2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Konto];"
            End If
        Case 3: 'Qurtal
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrBuEr4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Quartal], [Konto]"
            Else
                SQL2 = "SELECT * FROM qryPrBuEr4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Quartal], [Konto];"
            End If
        End Select
        CmTex = True
Case "EiUbGr":
        SQL1 = "qryPrBuEr3"
        If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryPrBuEr1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto]"
            Else
                SQL2 = "SELECT * FROM qryPrBuEr1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto];"
            End If
        CmTex = True
Case "SaKoMo": 'Saldenliste der Sachkonten"
        SQL1 = "qryPrBuKo3"
        Gefun = InStr(1, Krite, "IDT", 1)
        Select Case ZeRau
        Case 1: 'Monat
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK], [Monat]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK], [Monat];"
                End If
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo1 WHERE " & Krite & " ORDER BY [Jahr], [IDK], [Monat]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo1 WHERE " & Krite & " ORDER BY [Jahr], [IDK], [Monat];"
                End If
            End If
        Case 2: 'Jahr
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK];"
                End If
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo2 WHERE " & Krite & " ORDER BY [Jahr], [IDK]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo2 WHERE " & Krite & " ORDER BY [Jahr], [IDK];"
                End If
            End If
        Case 3: 'Qurtal
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK], [Quartal]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDK], [Quartal];"
                End If
            
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuKo4 WHERE " & Krite & " ORDER BY [Jahr], [IDK], [Quartal]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuKo4 WHERE " & Krite & " ORDER BY [Jahr], [IDK], [Quartal];"
                End If
            End If
        End Select
        CmTex = True
Case "SaGeMo": 'Saldenliste der Geldkonten
        SQL1 = "qryPrBuBa3"
        Select Case ZeRau
        Case 1: 'Monat
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB], [Monat]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB], [Monat];"
                End If
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa1 WHERE " & Krite & " ORDER BY [Jahr], [IDB], [Monat]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa1 WHERE " & Krite & " ORDER BY [Jahr], [IDB], [Monat];"
                End If
            End If
        Case 2: 'Jahr
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB];"
                End If
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa2 WHERE " & Krite & " ORDER BY [Jahr], [IDB]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa2 WHERE " & Krite & " ORDER BY [Jahr], [IDB];"
                End If
            End If
        Case 3: 'Qurtal
            If Gefun > 0 Then
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB], [Quartal]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [IDB], [Quartal];"
                End If
            Else
                If GlTyp < 2 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuBa4 WHERE " & Krite & " ORDER BY [Jahr], [IDB], [Quartal]"
                Else
                    SQL2 = "SELECT * FROM qryPrBuBa4 WHERE " & Krite & " ORDER BY [Jahr], [IDB], [Quartal];"
                End If
            End If
        End Select
        CmTex = True
Case "KonAus": 'Einzelkontenauswertung
        SQL1 = "qryPrBuJo3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrBuJo1a WHERE " & Krite & " ORDER BY Datum"
        Else
            SQL2 = "SELECT * FROM qryPrBuJo1a WHERE " & Krite & " ORDER BY [Datum];"
        End If
        CmTex = True
Case "KonGru": 'Einzelkontenauswertung (Gruppiert Sachkonten)
        SQL1 = "qryPrBuJo3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrBuJo2a WHERE " & Krite & " ORDER BY Konto, Datum"
        Else
            SQL2 = "SELECT * FROM qryPrBuJo2a WHERE " & Krite & " ORDER BY [Konto], [Datum];"
        End If
        CmTex = True
Case "KonGel": 'Einzelkontenauswertung (Gruppiert Geldkonten)
        SQL1 = "qryPrBuJo3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrBuJo4a WHERE " & Krite & " ORDER BY IDB, Datum"
        Else
            SQL2 = "SELECT * FROM qryPrBuJo4a WHERE " & Krite & " ORDER BY [IDB], [Datum];"
        End If
        CmTex = True
Case "ReUmsa":
        SQL1 = "qryPrReUm3"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrReUm1a WHERE " & Krite & " ORDER BY Sort"
        Else
            SQL2 = "SELECT * FROM qryPrReUm1a WHERE " & Krite & " ORDER BY [Sort];"
        End If
        CmTex = True
Case "KomAus": 'Kompaktauswertung
        SQL1 = "qryPrBuKom"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrBuKom2 WHERE " & Krite & " ORDER BY Jahr, Sort"
        Else
            SQL2 = "SELECT * FROM qryPrBuKom2 WHERE " & Krite & " ORDER BY [Jahr], [Sort];"
        End If
        CmTex = True
Case "UmStMo": 'Umsatzsteuerübersicht
        SQL1 = "qryPrBuSt3"
        Select Case ZeRau
        Case 1: 'Monat
            If GlTyp < 2 Then
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto]"
                Else
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt1 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Monat], [Konto]"
                End If
            Else
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM qryPrBuSt1a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Monat], [Konto];"
                Else
                    SQL2 = "SELECT * FROM qryPrBuSt1 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Monat], [Konto];"
                End If
            End If
        Case 2: 'Jahr
            If GlTyp < 2 Then
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Konto]"
                Else
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt2 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Konto]"
                End If
            Else
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM qryPrBuSt2a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Konto];"
                Else
                    SQL2 = "SELECT * FROM qryPrBuSt2 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Konto];"
                End If
            End If
        Case 3: 'Qurtal
            If GlTyp < 2 Then
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Quartal], [Konto]"
                Else
                    SQL2 = "SELECT * FROM dbo.qryPrBuSt4 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Quartal], [Konto]"
                End If
            Else
                If InStr(1, Krite, "IDT", 1) > 0 Then
                    SQL2 = "SELECT * FROM qryPrBuSt4a WHERE " & Krite & " ORDER BY [IDT], [Jahr], [EA], [Quartal], [Konto];"
                Else
                    SQL2 = "SELECT * FROM qryPrBuSt4 WHERE " & Krite & " ORDER BY [Jahr], [EA], [Quartal], [Konto];"
                End If
            End If
        End Select
        CmTex = True
Case "BuKont": 'Kontenplan
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM qryPrKon1 WHERE " & Krite & " ORDER BY Kontonummer"
        Else
            SQL2 = "SELECT * FROM qryPrKon1 WHERE " & Krite & " ORDER BY [Kontonummer];"
        End If
        CmTex = True
Case "ReUber":
        SQL1 = "qryPrSimReL2"
        If GlTyp < 2 Then
            SQL2 = "SELECT * FROM dbo.qryPrSimReL1 WHERE " & Krite & " ORDER BY Datum"
        Else
            SQL2 = "SELECT * FROM qryPrSimReL1 WHERE " & Krite & " ORDER BY [Datum];"
        End If
        CmTex = True
Case "TerPat": 'Terminzettel
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
Case "TerSer":
        VerMo = True
        SQL1 = "qryPrTerm2"
        SQL2 = "qryPrTerm2"
Case "TeUbNo":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbEu":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "QuiTer":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "QuiTem":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbAN":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbAE":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeEiMa":
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
Case "TeEiSa":
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
Case "KasPro":
        VerMo = False
        SQL1 = "qryPrBuZal1"
        SQL2 = "qryPrBuZal1"
End Select

If SQL2 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

Select Case LCase(DaExt):
Case "blg": 'Variablen und Listen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    If CmTex = True Then
        Set RS32 = DBCmRe0(SQL2, CmTex)
    Else
        Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
    End If
    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If
Case "lst": 'Nur Listen
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    If CmTex = True Then
        Set RS32 = DBCmRe0(SQL2, CmTex)
    Else
        Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
    End If
    If RS32.RecordCount > 0 Then
        StaPr = True
    End If
Case "crd": 'Nur Variablen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
    If RS31.RecordCount > 0 Then
        StaPr = True
    End If
Case "asw": 'Variablen und Listen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe0(SQL1)
    Set RS32 = New ADODB.Recordset
    RS32.CursorLocation = adUseClient
    Set RS32 = DBCmRe0(SQL2)
    If RS31.RecordCount > 0 Then
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    End If
End Select

If EmSen > 0 Then
    DrVor = True 'wenn Emailversand, dann Druckvorschau
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

If StaPr = True Then
    If VerMo = True Then
        If RS31.Fields("Em_Aut").Value <> vbNullString Then
            VerDo = CBool(RS31.Fields("Em_Aut").Value)
        Else
            VerDo = False
        End If
        If VerDo = True Then 'Dokument verschlüsseln
            If RS31.Fields("Em_Pass").Value <> vbNullString Then
                VerPa = RS31.Fields("Em_Pass").Value
            Else
                VerDo = False
            End If
        Else
            VerDo = False
        End If
    End If
    
    If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
        VerDo = False
    End If

    Select Case ForNa
    Case "EMRAbr":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Rechnu":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Auslan":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "RechLa":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "RechAl":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "RechSt":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "StAnge":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "StGuts":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Uberwe":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "UberEU":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Angebo":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Gutsch":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "Liefer":
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "EiMahn": 'Einzelmahnung
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "SaMahn": 'Sammelmahnung
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    Case "StMahn": 'Einzelmahnung (Alternativ)
            If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
                If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
                    ReStr = RS31.Fields("Rechnungsnummer").Value
                End If
            End If
    End Select
    
    If ReStr <> vbNullString Then
        If RS31.Fields("Rechnungsdatum").Value <> vbNullString Then
            ReDat = Format$(RS31.Fields("Rechnungsdatum").Value, "dd.mm.yyyy")
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        End If
        If RS31.Fields("IDP").Value > 0 Then
            ManNr = RS31.Fields("IDP").Value
        End If
        If RS31.Fields("IDM").Value > 0 Then
            MitNr = RS31.Fields("IDM").Value
        End If
        EmNam = RS31.Fields("R_Name").Value & ", " & RS31.Fields("R_Vorname").Value
        If RS31.Fields("Telefon5").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon5").Value
        ElseIf RS31.Fields("Telefon6").Value <> vbNullString Then
            EmEmp = RS31.Fields("Telefon6").Value
        Else
            EmEmp = vbNullString
        End If
        If RS31.Fields("R_Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("R_Briefanrede").Value
        ElseIf RS31.Fields("Briefanrede").Value <> vbNullString Then
            EmBrf = RS31.Fields("Briefanrede").Value
        Else
            EmBrf = vbNullString
        End If
    End If

    Select Case LCase(DaExt):
    Case "blg": 'Variablen und Listen
            If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
                LLDeV RS31, DauDg, False, MxCod
            Else
                LLDeV RS31, DauDg
            End If
            LLDeF RS32
            GesZa = RS32.RecordCount
    Case "lst": 'Nur Listen
            LLDeB
            LLDeF RS32
            GesZa = RS32.RecordCount
    Case "crd": 'Variablen
            LLDeV RS31
            GesZa = RS31.RecordCount
    Case "asw": 'Variablen und Listen
            LLDeV RS31
            LLDeF RS32
            GesZa = RS32.RecordCount
    End Select
    
    If GlDub = True Then 'Dublexdruck aktivieren
        DevMo.dmDuplex = DMDUP_VERTICAL
    End If
    
    Select Case GlSGr
    Case 2: DevMo.dmPaperSize = DMPAPER_A4
    Case 3: DevMo.dmPaperSize = DMPAPER_A5
    Case 4: DevMo.dmPaperSize = DMPAPER_A6
    End Select
    
    If GlDub = True Or GlSGr > 1 Then
        RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
    End If
    
    If DaExt = "crd" Then
        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
            End If
        End If
    Else
        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
        End If
    End If

    If FaxNr <> vbNullString Then
        RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
    End If

    If RetWe = 0 Then
        If DrDia = True Then
            RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        End If
        If RetWe = 0 Then
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    If DaExt <> "crd" Then
                        While Not RS32.EOF
                            LLDeF RS32
                            RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                            While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                RetWe = LlPrint(mHwnd)
                            Wend
                            AktZa = AktZa + 1
                            RS32.MoveNext
                        Wend
                    End If
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
    End If

    If Proto = True Then 'Soll Protokolleintrag als PDF erzeugt werden?
        Select Case LCase(DaExt):
        Case "blg": 'Variablen und Listen
                RS31.MoveFirst
                RS32.MoveFirst
                LLDeV RS31, DauDg, True
                LLDeF RS32
                GesZa = RS32.RecordCount
        Case "lst": 'Nur Listen
                RS32.MoveFirst
                LLDeB
                LLDeF RS32
                GesZa = RS32.RecordCount
        Case "crd": 'Variablen
                RS31.MoveFirst
                LLDeV RS31, vbNullString, True
                GesZa = RS31.RecordCount
        Case "asw": 'Variablen und Listen
                RS31.MoveFirst
                RS32.MoveFirst
                LLDeV RS31, vbNullString, True
                LLDeF RS32
                GesZa = RS32.RecordCount
        End Select

        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNam)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", GlBPf)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Quiet", 1)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.ShowResult", 0)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", vbNullString)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
        
        If DaExt = "crd" Then
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, 0)
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        End If
        If RetWe = 0 Then
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                If DaExt <> "crd" Then
                    While Not RS32.EOF
                        LLDeF RS32
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                End If
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
            End If
        End If
    End If
    If EmSen > 0 Then 'Emailversand
        If GlDru.XRech = True Then
            If ReStr <> vbNullString Then
                XReDa = GlTmp & "XRechnung_Beleg_" & ReStr & ".xml"
            Else
                XReDa = GlTmp & "XRechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".xml"
            End If
            RS31.MoveFirst
            RS32.MoveFirst
            DoEvents
            S_XRech RS31, RS32, XReDa
            DoEvents
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

Select Case LCase(DaExt):
Case "blg": 'Variablen und Listen
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
Case "lst": 'Nur Listen
    RS32.Close
    Set RS32 = Nothing
Case "crd": 'Nur Listen
    RS31.Close
    Set RS31 = Nothing
Case "asw": 'Variablen und Listen
    RS31.Close
    RS32.Close
    Set RS31 = Nothing
    Set RS32 = Nothing
End Select

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If GlDru.GoBDk = True Then
                If GlDru.ReAbs = True Then
                    If ReStr <> vbNullString Then
                        DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                    Else
                        DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    End If
                    If VerDo = True Then 'Dokument verschlüsseln
                        hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                        RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                        LlStgsysStorageClose hwOrg
                    Else
                        RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                    End If
                    If GlDru.XRech = True Then
                        SERech DaNam, XReDa
                        DoEvents
                        If ReStr <> vbNullString Then
                            DaNam = GlTmp & "ERechnung_Beleg_" & ReStr & ".pdf"
                        Else
                            DaNam = GlTmp & "ERechnung_Beleg_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                        End If
                    End If
                    
                    If EmSen < 3 Then
                        SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                    ElseIf EmSen = 6 Then
                        DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                        DoEvents
                        If DoStr <> vbNullString Then
                            Lange = Len(DoStr)
                            Posit = InStr(1, DoStr, ";", 1)
                            If Posit > 0 Then
                                DoLnk = Left$(DoStr, Posit - 1)
                                PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                                With GlTxV
                                    .TxStr = GlEmT(12, 3)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                End With
                                EmBet = SEmTx()
                                With GlTxV
                                    .TxStr = GlEmT(12, 1)
                                    .Datum = ReDat
                                    .DaStr = ReDat
                                    .ReStr = ReStr
                                    .MitNr = MitNr
                                    .ManNr = ManNr
                                    .PatNr = PatNr
                                    .DoLnk = DoLnk
                                    .PoLnk = PoLnk
                                    .DoLan = GlVrw
                                End With
                                EmTex = EmTex & SEmTx()
                                SMaNe PatNr, EmAdr, , EmTex, EmBet
                            End If
                        End If
                    End If
                Else
                    SPopu "Rechnung abschließen", "Die Rechnungsdruckoptionen stehen nicht auf abschließen!", IC48_Forbidden
                End If
            Else
                If ReStr <> vbNullString Then
                    DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                Else
                    DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                End If
                If VerDo = True Then 'Dokument verschlüsseln
                    hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                    RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                    LlStgsysStorageClose hwOrg
                Else
                    RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                End If
                If EmSen < 3 Then
                    SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                End If
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If ReStr <> vbNullString Then
                DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
            Else
                DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            End If
            If VerDo = True Then 'Dokument verschlüsseln
                hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrLi " & Err.Number
Resume Next

End Sub
Public Sub LLPrLm()
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim hwAnf As Long
Dim PatNr As Long
Dim AkDat As Date
Dim TaSum As Double 'Tagessumme
Dim AnSum As Single 'Anzahlsungssumme
Dim GuiNa As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim DruAn As String 'Vorschaudatei
Dim DauDg As String
Dim MxCod As String
Dim ReStr As String
Dim ReDat As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim DoLnk As String
Dim FoLnk As String
Dim DoStr As String
Dim AnNam As String
Dim ForAn As String
Dim ForTm As String
Dim KriSt As String  'Kriterienstring
Dim Warun As String 'Währung
Dim AkZif As String 'Ziffer
Dim MeStr As String 'Memostring
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim KmpRe As Boolean
Dim AkPos As Integer
Dim Gefun As Integer
Dim Posit As Integer
Dim Lange As Integer
Dim TagZa As Integer 'Tageszahl
Dim ReZal As Integer 'Rechungszahl
Dim LegZa As Integer 'Legendenzahl
Dim LeiZa As Integer 'Leiszungszahl
Dim GeZei As Integer 'Gesamttageszeit
Dim DevMo As DEVMODEA
Dim SQL1, SQL2, SQL3, SQL4, SQL5 As String

Set FM = frmMain

ForTm = ForNa 'WICHTIG!
ForAn = GlFrm(2, FoIdx)
If GlFrn <> vbNullString Then 'Formulareordner
    AnNam = GlFrn & S_FoCh(ForAn)
Else
    AnNam = GlFrO & S_FoCh(ForAn)
End If

Select Case ForAn
Case "RechKo": KmpRe = True
Case "RechKu": KmpRe = True
Case "RechFa": KmpRe = True
Case "RechSk": KmpRe = True
End Select

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
AkPos = InStrRev(AnNam, "\", Len(AnNam), 1)
DruAn = TmPfa & Mid$(AnNam, AkPos + 1, Len(AnNam) - (AkPos + 3)) & "LL"
ExFmt = "PDF"

SQL1 = "qryPrSimOP1"
SQL2 = "qryPrSimOP2"
If KmpRe = True Then
    SQL3 = "qryPrSimRe1"
    SQL4 = "qryPrSimRe5"
    SQL5 = "qryPrSimRe4"
Else
    SQL3 = "qryPrSimRe1"
    SQL4 = "qryPrSimRe2"
End If

If MaRec > 0 Then
    DauDg = S_DrDi(MaRec)
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        MxCod = S_ReMx(MaRec)
    End If
End If

Set RS31 = New ADODB.Recordset
RS31.CursorLocation = adUseClient
Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
Set RS32 = New ADODB.Recordset
RS32.CursorLocation = adUseClient
Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
If RS31.RecordCount > 0 Then
    If RS32.RecordCount > 0 Then
        StaPr = True
    End If
End If

If MaRec > 0 Then
    If KmpRe = True Then
    Set RS33 = New ADODB.Recordset
        RS33.CursorLocation = adUseClient 'Rechnungskopf
        Set RS33 = DBCmRe1(SQL3, "@IdStr", MaRec)
        ReZal = RS33.RecordCount
        Set RS35 = New ADODB.Recordset 'Legende
        RS35.CursorLocation = adUseClient
        Set RS35 = DBCmRe1(SQL4, "@IdStr", MaRec)
        LegZa = RS35.RecordCount
        Set RS36 = New ADODB.Recordset 'Abrechungstage
        RS36.CursorLocation = adUseClient
        Set RS36 = DBCmRe1(SQL5, "@IdStr", MaRec)
        TagZa = RS35.RecordCount
        If ReZal = 0 Then
            StaPr = False
        ElseIf LegZa = 0 Then
            StaPr = False
        ElseIf TagZa = 0 Then
            StaPr = False
        End If
    Else
        Set RS33 = New ADODB.Recordset
        RS33.CursorLocation = adUseClient
        Set RS33 = DBCmRe1(SQL3, "@IdStr", MaRec)
        Set RS34 = New ADODB.Recordset
        RS34.CursorLocation = adUseClient
        Set RS34 = DBCmRe1(SQL4, "@IdStr", MaRec)
        If RS33.RecordCount = 0 Then
            StaPr = False
        ElseIf RS34.RecordCount = 0 Then
            StaPr = False
        End If
    End If
Else
    StaPr = False
End If

If EmSen > 0 Then
    DrVor = True 'wenn Emailversand, dann Druckvorschau
End If

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

If StaPr = True Then
    If RS31.Fields("Em_Aut").Value <> vbNullString Then
        VerDo = CBool(RS31.Fields("Em_Aut").Value)
    Else
        VerDo = False
    End If
    If VerDo = True Then 'Dokument verschlüsseln
        If RS31.Fields("Em_Pass").Value <> vbNullString Then
            VerPa = RS31.Fields("Em_Pass").Value
        Else
            VerDo = False
        End If
    Else
        VerDo = False
    End If
    
    If EmSen = 0 Then 'kein Emailversand, keine PDF Verschlüsselung
        VerDo = False
    End If
    
    If RS31.Fields("Rechnungsnummer").Value <> vbNullString Then
        If Len(RS31.Fields("Rechnungsnummer").Value) > 1 Then
            ReStr = RS31.Fields("Rechnungsnummer").Value
        End If
    End If
    If RS31.Fields("PatNr").Value > 0 Then
        PatNr = RS31.Fields("PatNr").Value
    Else
        PatNr = 0
    End If
        
    Select Case GlSGr
    Case 2: DevMo.dmPaperSize = DMPAPER_A4
    Case 3: DevMo.dmPaperSize = DMPAPER_A5
    Case 4: DevMo.dmPaperSize = DMPAPER_A6
    End Select

    If GlDub = True Then 'Dublexdruck aktivieren
        DevMo.dmDuplex = DMDUP_VERTICAL
    End If

    '---------- Mahnung ----------

    AktZa = 1
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        LLDeV RS31, DauDg, False, MxCod
    Else
        LLDeV RS31, DauDg
    End If
    LLDeF RS32
    GesZa = RS32.RecordCount

    If GlDub = True Or GlSGr > 1 Then
        RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
    End If

    RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
    If DrVor = True Then 'Druckvorschau
        If GlRDP = True Then
            RetWe = LlViewerProhibitAction(mHwnd, 116)
        End If
        RetWe = LlViewerProhibitAction(mHwnd, 115)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        If DrSof = False Then
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
        End If
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    Else
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
    End If

    If FaxNr <> vbNullString Then
        RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
    End If

    If RetWe = 0 Then
        RetWe = LlPrint(mHwnd)
        If RetWe = 0 Then
            RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
            If RetWe = 0 Then
                While Not RS32.EOF
                    LLDeF RS32
                    RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then
                    DruOk = True
                End If
            End If
        End If
    End If

    '---------- Rechnung ----------
    
    AktZa = 1
    ForNa = ForAn 'WICHTIG!

    If KmpRe = True Then
        GesZa = LegZa + TagZa

        Set RS34 = New ADODB.Recordset
        With RS34
            .CursorLocation = adUseClient
            .Fields.Append "Datum", adDate
            .Fields.Append "Behandlungsdatum", adChar, 10
            .Fields.Append "Ziffer", adChar, 100
            .Fields.Append "Bezeichnung", adVariant
            .Fields.Append "Kommentar", adVariant
            .Fields.Append "Einzelbetrag", adDouble
            .Fields.Append "Gesamtbetrag", adDouble
            .Fields.Append "Anzahlungsbetrag", adDouble
            .Fields.Append "Währung", adChar, 10
            .Fields.Append "Anzahl", adInteger
            .Fields.Append "Multi", adDouble
            .Fields.Append "Typ", adInteger
            .Fields.Append "Zeit", adInteger
            .Fields.Append "ICD", adVariant
            .Fields.Append "Diagnose", adVariant
            .Fields.Append "Analog", adInteger
            If .State = adStateClosed Then .Open
        End With
        
        If RS34.Supports(adAddNew) Then
            Do
            AkDat = RS36.Fields("Datum").Value
            Set RS37 = New ADODB.Recordset
            RS37.CursorLocation = adUseClient
            Set RS37 = DBCmRe2("qryPrSimRe6", "@IdStr", "@IdDat", MaRec, AkDat)
            If RS37.RecordCount > 0 Then
                MeStr = vbNullString
                TaSum = 0
                AnSum = 0
                GeZei = 0
                Do
                If RS37.Fields("Ziffer").Value <> vbNullString Then
                    AkZif = RS37.Fields("Ziffer").Value
                    LeiZa = RS37.Fields("Anzahl").Value
    
                    If MeStr = vbNullString Then
                        If LeiZa = 1 Then
                            MeStr = AkZif
                        Else
                            MeStr = LeiZa & "x" & AkZif
                        End If
                    Else
                        If LeiZa = 1 Then
                            MeStr = MeStr & "; " & AkZif
                        Else
                            MeStr = MeStr & "; " & LeiZa & "x" & AkZif
                        End If
                    End If
                End If
                If RS37.Fields("Währung").Value <> vbNullString Then
                    Warun = RS37.Fields("Währung").Value
                Else
                    Warun = "€"
                End If
                If RS37.Fields("Zeit").Value <> 0 Then
                    GeZei = GeZei + RS37.Fields("Zeit").Value
                End If
                If RS37.Fields("Gesamtbetrag").Value <> 0 Then
                    TaSum = TaSum + RS37.Fields("Gesamtbetrag").Value
                End If
                If RS37.Fields("Anzahlungsbetrag").Value <> 0 Then
                    AnSum = AnSum + RS37.Fields("Anzahlungsbetrag").Value
                End If
                RS37.MoveNext
                Loop Until RS37.EOF
                RS34.AddNew
                RS34.Fields("Datum").Value = AkDat
                RS34.Fields("Behandlungsdatum").Value = Format$(AkDat, "dd.mm.yyyy")
                RS34.Fields("Bezeichnung").Value = MeStr
                RS34.Fields("Gesamtbetrag").Value = TaSum
                RS34.Fields("Anzahlungsbetrag").Value = AnSum
                RS34.Fields("Währung").Value = Warun
                RS34.Fields("Zeit").Value = GeZei
                RS34.Fields("Typ").Value = 0
            End If
            RS37.Close
            Set RS37 = Nothing
            RS36.MoveNext
            Loop Until RS36.EOF

            Do
            RS34.AddNew
            If RS35.Fields("Ziffer").Value <> vbNullString Then RS34.Fields("Ziffer").Value = RS35.Fields("Ziffer").Value
            If RS35.Fields("Bezeichnung").Value <> vbNullString Then RS34.Fields("Bezeichnung").Value = RS35.Fields("Bezeichnung").Value
            If RS35.Fields("ICD").Value <> vbNullString Then RS34.Fields("ICD").Value = RS35.Fields("ICD").Value
            If RS35.Fields("Diagnose").Value <> vbNullString Then RS34.Fields("Diagnose").Value = RS35.Fields("Diagnose").Value
            If RS35.Fields("Multi").Value <> vbNullString Then RS34.Fields("Multi").Value = RS35.Fields("Multi").Value
            RS34.Fields("Typ").Value = RS35.Fields("Typ").Value
            If RS35.Fields("Einzelbetrag").Value <> 0 Then
                RS34.Fields("Einzelbetrag").Value = RS35.Fields("Einzelbetrag").Value
                RS34.Fields("Währung").Value = RS35.Fields("Währung").Value
            End If
            If GlAnl = True Then
                If RS35.Fields("Analog").Value <> vbNullString Then
                    RS34.Fields("Analog").Value = RS35.Fields("Analog").Value
                Else
                    RS34.Fields("Analog").Value = 0
                End If
            Else
                RS34.Fields("Analog").Value = 0
            End If
            If RS35.Fields("Kommentar").Value <> vbNullString Then RS34.Fields("Kommentar").Value = RS35.Fields("Kommentar").Value
            RS35.MoveNext
            Loop Until RS35.EOF
            
            RS34.Update
            RS34.MoveFirst
        End If
    
    Else
    
        GesZa = RS34.RecordCount
        
    End If
    
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        LLDeV RS33, DauDg, False, MxCod
    Else
        LLDeV RS33, DauDg
    End If
    LLDeF RS34

    If GlDub = True Or GlSGr > 1 Then
        RetWe = LlSetPrinterInPrinterFile(mHwnd, LL_PROJECT_LIST, FiNam, -1, DrNam, DevMo)
    End If

    RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
    If DrVor = True Then 'Druckvorschau
        If GlRDP = True Then
            RetWe = LlViewerProhibitAction(mHwnd, 116)
        End If
        RetWe = LlViewerProhibitAction(mHwnd, 115)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, AnNam, LL_PRINT_EXPORT, 0)
        If DrSof = False Then
            RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
        End If
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
    Else
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, AnNam, LL_PRINT_EXPORT, 0)
    End If

    If FaxNr <> vbNullString Then
        RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
    End If

    If RetWe = 0 Then
        RetWe = LlPrint(mHwnd)
        If RetWe = 0 Then
            RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
            If RetWe = 0 Then
                While Not RS34.EOF
                    LLDeF RS34
                    RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS34.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then
                    DruOk = True
                End If
            End If
        End If
    End If
    RS34.Close
    Set RS34 = Nothing
        
    '---------- Zusammenfassen ----------
    
    If RetWe = 0 Then
        hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
        hwAnf = LlStgsysStorageOpen(DruAn, TmPfa, False, True)
        RetWe = LlStgsysAppend(hwOrg, hwAnf)
        LlStgsysStorageClose hwAnf
        LlStgsysStorageClose hwOrg
    End If

    ForNa = ForTm 'WICHTIG!

    '---------- Protokolleintrag ----------

    If Proto = True Then 'Soll Protokolleintrag als PDF erzeugt werden?
        RS31.MoveFirst
        RS32.MoveFirst
        LLDeV RS31, DauDg, True
        LLDeF RS32
        GesZa = RS32.RecordCount

        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNam)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", GlBPf)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Quiet", 1)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.ShowResult", 0)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", vbNullString)
        RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
        
        RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
        RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
        RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
        RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        
        If RetWe = 0 Then
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                While Not RS32.EOF
                    LLDeF RS32
                    While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                        RetWe = LlPrint(mHwnd)
                    Wend
                    AktZa = AktZa + 1
                    RS32.MoveNext
                Wend
                Do
                RetWe = LlPrintFieldsEnd(mHwnd)
                Loop Until RetWe <> LL_WRN_REPEAT_DATA
                RetWe = LlPrintEnd(mHwnd, 0)
            End If
        End If
    End If
    
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

RS31.Close
RS32.Close
RS33.Close
Set RS31 = Nothing
Set RS32 = Nothing
Set RS33 = Nothing

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If GlDru.GoBDk = True Then
                If GlDru.ReAbs = True Then
                    If ReStr <> vbNullString Then
                        DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                    Else
                        DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                    End If
                    If VerDo = True Then 'Dokument verschlüsseln
                        hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                        RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                        LlStgsysStorageClose hwOrg
                    Else
                        RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                    End If
                    If EmSen < 3 Then
                        SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                    End If
                Else
                    SPopu "Rechnung abschließen", "Die Rechnungsdruckoptionen stehen nicht auf abschließen!", IC48_Forbidden
                End If
            Else
                If ReStr <> vbNullString Then
                    DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
                Else
                    DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
                End If
                If VerDo = True Then 'Dokument verschlüsseln
                    hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                    RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                    LlStgsysStorageClose hwOrg
                Else
                    RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
                End If
                If EmSen < 3 Then
                    SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
                End If
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            If ReStr <> vbNullString Then
                DaNam = GlTmp & "Rechnung_Beleg_" & ReStr & ".pdf"
            Else
                DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            End If
            If VerDo = True Then 'Dokument verschlüsseln
                hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
        If .FilVor(DruAn) = True Then
            .DaLoe = DruAn & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy 'Druckvorschautyp
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrLm " & Err.Number
Resume Next

End Sub

Public Sub LLPrLs()
On Error GoTo DeErr
'Druckt mehrere Aufträge

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim ZusDa As String 'Zsammengefasste Datei
Dim FilNa As String
Dim VorDa As String
Dim GuiNa As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim CmTex As Boolean
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim VerMo As Boolean 'Verschlüsselung möglich
Dim AktPo As Integer 'Aktuelle Position
Dim AnzPo As Integer 'Gesamteinträge
Dim SeiZa As Integer
Dim AkPos As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

GuiNa = CreateID("P")
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
VorDa = GuiNa & "_" & "0001.LL"
DaExt = Right$(FiNam, 3)
ZusDa = TmPfa & VorDa
AnzPo = UBound(GloDr)

Select Case ForNa
Case "Uberwe":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe1"
Case "UberEU":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe1"
Case "Angebo":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Gutsch":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "Liefer":
        VerMo = True
        SQL1 = "qryPrSimRe1"
        SQL2 = "qryPrSimRe2"
Case "EiMahn": 'Einzelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
Case "SaMahn": 'Sammelmahnung
        VerMo = True
        SQL1 = "qryPrSimOP4"
        SQL2 = "qryPrSimOP5"
Case "StMahn": 'Einzelmahnung (Alternativ)
        VerMo = True
        SQL1 = "qryPrSimOP1"
        SQL2 = "qryPrSimOP2"
Case "LabKom":
        VerMo = True
        SQL1 = "qryPrLaBe1"
        SQL2 = "qryPrLaBe5"
Case "LabAuf":
        VerMo = True
        SQL1 = "qryPrLaAu1"
        SQL2 = "qryPrLaAu2"
Case "TerPat": 'Terminzettel
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
Case "TerSer":
        VerMo = True
        SQL1 = "qryPrTerm2"
        SQL2 = "qryPrTerm2"
Case "TeUbNo":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbEu":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "QuiTer":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "QuiTem":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbAN":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeUbAE":
        VerMo = True
        SQL1 = "qryPrTerm5"
        SQL2 = "qryPrTerm5"
Case "TeEiMa":
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
Case "TeEiSa":
        VerMo = True
        SQL1 = "qryPrTerm1"
        SQL2 = "qryPrTerm4"
End Select

If SQL1 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail VerDo, VerPa

For AktPo = 1 To AnzPo
    FilNa = TmPfa & GuiNa & "_" & Format$(AktPo, "0000") & ".LL"
    PrBr1.Min = 0
    PrBr1.Max = AnzPo
    FS.Caption = "Erzeugen der Druckdateien..."
    
    StaPr = False
    Select Case LCase(DaExt):
    Case "blg": 'Variablen und Listen
        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
        Set RS32 = New ADODB.Recordset
        RS32.CursorLocation = adUseClient
        Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
        If RS31.RecordCount > 0 Then
            If RS32.RecordCount > 0 Then
                StaPr = True
            End If
        End If
    Case "lst": 'Nur Listen
        Set RS32 = New ADODB.Recordset
        RS32.CursorLocation = adUseClient
        Set RS32 = DBCmRe1(SQL2, "@IdStr", GloDr(AktPo))
        If RS32.RecordCount > 0 Then
            StaPr = True
        End If
    Case "crd": 'Nur Variablen
        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(AktPo))
        If RS31.RecordCount > 0 Then
            StaPr = True
        End If
    Case "asw": 'Variablen und Listen
        Set RS31 = New ADODB.Recordset
        RS31.CursorLocation = adUseClient
        Set RS31 = DBCmRe0(SQL1)
        Set RS32 = New ADODB.Recordset
        RS32.CursorLocation = adUseClient
        Set RS32 = DBCmRe0(SQL2)
        If RS31.RecordCount > 0 Then
            If RS32.RecordCount > 0 Then
                StaPr = True
            End If
        End If
    End Select

    If StaPr = True Then
        If VerMo = True Then
            If RS31.Fields("Em_Aut").Value <> vbNullString Then
                VerDo = CBool(RS31.Fields("Em_Aut").Value)
            Else
                VerDo = False
            End If
            If VerDo = True Then 'Dokument verschlüsseln
                If RS31.Fields("Em_Pass").Value <> vbNullString Then
                    VerPa = RS31.Fields("Em_Pass").Value
                Else
                    VerDo = False
                End If
            Else
                VerDo = False
            End If
        End If
        If RS31.Fields("PatNr").Value > 0 Then
            PatNr = RS31.Fields("PatNr").Value
        Else
            PatNr = 0
        End If

        AktZa = 1
        Select Case LCase(DaExt):
        Case "blg": 'Variablen und Listen
                LLDeV RS31
                LLDeF RS32
                GesZa = RS32.RecordCount
        Case "lst": 'Nur Listen
                LLDeB
                LLDeF RS32
                GesZa = RS32.RecordCount
        Case "crd": 'Variablen
                LLDeV RS31
                GesZa = RS31.RecordCount
        Case "asw": 'Variablen und Listen
                LLDeV RS31
                LLDeF RS32
                GesZa = RS32.RecordCount
        End Select
        PrBr2.Min = 0
        PrBr2.Max = GesZa
        If DaExt = "crd" Then
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
                End If
                If DrSof = False Then
                    RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
                End If
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
                End If
            End If
        Else
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
                If DrSof = False Then
                    RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
                End If
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
            End If
        End If
        
        If FaxNr <> vbNullString Then
            RetWe = LlPrintSetProjectParameter(mHwnd, "LL.FAX.RecipNumber", FaxNr, 0)
        End If
        
        If RetWe = 0 Then
            If AktPo = 1 Then
                If DrDia = True Then
                    RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
                End If
            End If
            If RetWe = 0 Then
                If AktPo = 1 Then
                    FS.Show
                    DoEvents
                End If
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                    If RetWe = 0 Then
                        If DaExt <> "crd" Then
                            While Not RS32.EOF
                                LLDeF RS32
                                PrBr2.Value = AktZa
                                DoEvents
                                While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                                RS32.MoveNext
                            Wend
                        End If
                        Do
                        RetWe = LlPrintFieldsEnd(mHwnd)
                        Loop Until RetWe <> LL_WRN_REPEAT_DATA
                        RetWe = LlPrintEnd(mHwnd, 0)
                        If RetWe = 0 Then DruOk = True
                        
                        If DrVor = True Then 'Druckvorschau
                            With clFil
                                .DaCop = DruDa & ";" & FilNa & vbNullChar
                                If .FilCop(3) = False Then
                                    SPopu "Datei schreibgeschützt", "Die Datei kann nicht umbenannt werden", IC48_Warning
                                Else
                                    If RetWe = 0 Then
                                        DruOk = True
                                    End If
                                End If
                            End With
                        Else
                            If RetWe = 0 Then
                                DruOk = True
                            End If
                        End If
                        
                    End If
                End If
            End If
        End If
        DoEvents
        
        If Proto = True Then 'Soll Protokolleintrag als PDF erzeugt werden?
            ExFmt = "PDF"
            RS31.MoveFirst
            RS32.MoveFirst
            LLDeV RS31, vbNullString, True
            LLDeF RS32
            
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.File", DaNam)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Path", GlBPf)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.Quiet", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "Export.ShowResult", 0)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.FitText", 1)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.Title", ForNa)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.OwnerPassword", PasOw)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.UserPassword", vbNullString)
            RetWe = LlXSetParameter(mHwnd, LL_LLX_EXTENSIONTYPE_EXPORT, ExFmt, "PDF.DontStackWorldModifications", 0)
        
            If DaExt = "crd" Then
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, 0)
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, ExFmt)
                RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
            End If
            If RetWe = 0 Then
                RetWe = LlPrint(mHwnd)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
        
    Else
        If AktPo = 1 Then
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
            Exit For
        End If
    End If
    
    Select Case LCase(DaExt):
    Case "blg": 'Variablen und Listen
        RS31.Close
        RS32.Close
        Set RS31 = Nothing
        Set RS32 = Nothing
    Case "lst": 'Nur Listen
        RS32.Close
        Set RS32 = Nothing
    Case "crd": 'Nur Listen
        RS31.Close
        Set RS31 = Nothing
    Case "asw": 'Variablen und Listen
        RS31.Close
        RS32.Close
        Set RS31 = Nothing
        Set RS32 = Nothing
    End Select
    
    DoEvents
    If TxDia.Text = "B" Then Exit For 'Abbrechen
    PrBr1.Value = AktPo
    StaPr = False
Next AktPo

Set clFil = Nothing

RetWe = LLStor(VorDa, GuiNa, AnzPo)

Unload FS
DoEvents

If DruOk = True Then
    If DrSof = True Then
        hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
        SeiZa = LlStgsysGetPageCount(hwOrg)
        RetWe = LlStgsysPrint(hwOrg, DrNam, DrNam, 1, SeiZa, Kopie, LS_PRINTFLAG_FIT Or LS_PRINTFLAG_STACKEDCOPIES Or LS_PRINTFLAG_ABORTABLEMETER, "Formulardruck", FM.hwnd)
        LlStgsysStorageClose hwOrg
    Else
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            If VerDo = True Then 'Dokument verschlüsseln
                hwOrg = LlStgsysStorageOpen(ZusDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(ZusDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, ZusDa, TmPfa, FM.hwnd)
            End If
        End If
    End If
Else
    SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Or EmSen = 6 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If EmSen = 0 Then 'Emailversand
    If DrVor = True Then 'Druckvorschau
        If DruOk = True Then
            Select Case GlVTy
            Case 1: VoDrMa ZusDa
            Case 3: SDrVo ZusDa
            End Select
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrLs " & Err.Number
Resume Next

End Sub
Public Sub LLPrPa()
On Error GoTo DeErr
'Druckt einen Einzelauftrag mit Parametern

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim AkSal As Single 'Saldo
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim StaPr As Boolean
Dim DruOk As Boolean
Dim CmTex As Boolean
Dim AktPo As Integer 'Aktuelle Position
Dim KriSt As String  'Kriterienstring
Dim AkPos As Integer
Dim DevMo As DEVMODEA
Dim SQL1, SQL2 As String

Set FM = frmMain

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
DaExt = Right$(FiNam, 3)

Select Case ForNa
Case "BuJour":
        If GlTyp > 1 Then
            SQL2 = "SELECT * FROM qryPrBuJo1 WHERE " & Krite & " ORDER BY [Datum], [Beleg];"
        Else
            SQL2 = "SELECT * FROM dbo.qryPrBuJo1 WHERE " & Krite & " ORDER BY Datum, Beleg"
        End If
        AkSal = AnSal
        CmTex = True
Case "BuKass":
        If GlTyp > 1 Then
            SQL2 = "SELECT * FROM qryPrBuKas WHERE " & Krite & " ORDER BY [Datum], [Beleg];"
        Else
            SQL2 = "SELECT * FROM dbo.qryPrBuKas WHERE " & Krite & " ORDER BY Datum, Beleg"
        End If
        AkSal = AnSal
        CmTex = True
End Select

If SQL2 = vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Sub
End If

Select Case LCase(DaExt):
Case "blg": 'Variablen und Listen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
    If RS31.RecordCount > 0 Then
        If RS33.RecordCount > 0 Then
            StaPr = True
        End If
    End If
Case "lst": 'Nur Listen
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe0(SQL2, CmTex)
    If RS33.RecordCount > 0 Then
        StaPr = True
    End If
Case "crd": 'Nur Variablen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe1(SQL1, "@IdStr", GloDr(1))
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe1(SQL2, "@IdStr", GloDr(1))
    If RS31.RecordCount > 0 Then
        If RS33.RecordCount > 0 Then
            StaPr = True
        End If
    End If
Case "asw": 'Variablen und Listen
    Set RS31 = New ADODB.Recordset
    RS31.CursorLocation = adUseClient
    Set RS31 = DBCmRe0(SQL1)
    Set RS33 = New ADODB.Recordset
    RS33.CursorLocation = adUseClient
    Set RS33 = DBCmRe0(SQL2)
    If RS31.RecordCount > 0 Then
        If RS33.RecordCount > 0 Then
            StaPr = True
        End If
    End If
End Select

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
LLOpt
LLMail

If StaPr = True Then
    GesZa = RS33.RecordCount

    Set RS32 = New ADODB.Recordset
    With RS32 'Felder Anlegen
        .CursorLocation = adUseClient
        For Each FL1 In RS33.Fields
            If FL1.Type = adChar Then
                .Fields.Append FL1.Name, FL1.Type, 250
            Else
                If FL1.Type > 200 Then
                    .Fields.Append FL1.Name, FL1.Type, 250
                Else
                    Select Case ForNa
                    Case "BuJour":
                        If FL1.Name = "Saldo" Then
                            .Fields.Append "Saldo", adDouble
                        Else
                            .Fields.Append FL1.Name, FL1.Type
                        End If
                    Case "BuKass":
                        If FL1.Name = "Saldo" Then
                            .Fields.Append "Saldo", adDouble
                        Else
                            .Fields.Append FL1.Name, FL1.Type
                        End If
                    Case Else:
                        .Fields.Append FL1.Name, FL1.Type
                    End Select
                End If
            End If
            
        Next FL1
        If .State = adStateClosed Then .Open
    End With

    If RS32.Supports(adAddNew) Then
        RS33.MoveFirst 'Felder Füllen
        Do
        RS32.AddNew
        For Each FL1 In RS33.Fields
            Select Case ForNa
            Case "BuJour":
                If FL1.Name = "Saldo" Then
                    AkSal = AkSal + (RS32.Fields("Einnahme").Value - RS32.Fields("Ausgabe").Value)
                    RS32.Fields("Saldo").Value = AkSal
                Else
                    If Not IsNull(RS33.Fields(FL1.Name).Value) Then RS32.Fields(FL1.Name).Value = RS33.Fields(FL1.Name).Value
                End If
            Case "BuKass":
                If FL1.Name = "Saldo" Then
                    AkSal = AkSal + (RS32.Fields("Einnahme").Value - RS32.Fields("Ausgabe").Value)
                    RS32.Fields("Saldo").Value = AkSal
                Else
                    If Not IsNull(RS33.Fields(FL1.Name).Value) Then RS32.Fields(FL1.Name).Value = RS33.Fields(FL1.Name).Value
                End If
            End Select
        Next FL1
        RS33.MoveNext
        Loop Until RS33.EOF
        RS32.Update
        RS32.MoveFirst
    End If

    Select Case LCase(DaExt):
    Case "blg": 'Variablen und Listen
            LLDeV RS31
            LLDeF RS32
    Case "lst": 'Nur Listen
            LLDeB
            LLDeF RS32
    Case "crd": 'Variablen
            LLDeV RS31
    Case "asw": 'Variablen und Listen
            LLDeV RS31
            LLDeF RS32
    End Select
    If DaExt = "crd" Then
        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_NORMAL, 0)
            End If
        End If
    Else
        If DrVor = True Then 'Druckvorschau
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If GlRDP = True Then
                RetWe = LlViewerProhibitAction(mHwnd, 116)
            End If
            RetWe = LlViewerProhibitAction(mHwnd, 115)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
            If DrSof = False Then
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            End If
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, Kopie)
            RetWe = LlPrintSetOption(mHwnd, LL_PRNOPT_COPIES, LL_COPIES_HIDE)
        Else
            RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
            If DrDia = True Then
                RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
            Else
                RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
            End If
        End If
    End If
    If RetWe = 0 Then
        If DrDia = True Then
            RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        End If
        If RetWe = 0 Then
            RetWe = LlPrint(mHwnd)
            If RetWe = 0 Then
                RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
                If RetWe = 0 Then
                    While Not RS32.EOF
                        LLDeF RS32
                        RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                        While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                            RetWe = LlPrint(mHwnd)
                        Wend
                        AktZa = AktZa + 1
                        RS32.MoveNext
                    Wend
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                    RetWe = LlPrintEnd(mHwnd, 0)
                    If RetWe = 0 Then DruOk = True
                End If
            End If
        End If
    End If
    RS32.Close
    Set RS32 = Nothing
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

Select Case LCase(DaExt):
Case "blg": 'Variablen und Listen
    RS31.Close
    RS33.Close
    Set RS31 = Nothing
    Set RS33 = Nothing
Case "lst": 'Nur Listen
    RS33.Close
    Set RS33 = Nothing
Case "crd": 'Nur Listen
    RS31.Close
    Set RS31 = Nothing
Case "asw": 'Variablen und Listen
    RS31.Close
    RS33.Close
    Set RS31 = Nothing
    Set RS33 = Nothing
End Select

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then 'Emailversand
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then 'Emailversand
                SMaNe 0, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If StaPr = True Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrPa " & Err.Number
Resume Next

End Sub
Public Sub LLPrUn(ByVal TabRp As Integer)
On Error GoTo DeErr
'Druckt einen Einzelauftrag

Dim GesZa As Long
Dim AktZa As Long
Dim RetWe As Long
Dim hwOrg As Long
Dim PatNr As Long
Dim ManNr As Long
Dim MitNr As Long
Dim DaExt As String
Dim DruNa As String * 100 'Druckername
Dim PorNa As String * 100 'Portname
Dim DruDa As String 'Vorschaudatei
Dim DoStr As String
Dim DoLnk As String
Dim PoLnk As String
Dim RzStr As String
Dim RzDat As String
Dim VerPa As String 'Verschlüsselungspasswort
Dim AkWie As Integer
Dim VerDo As Boolean 'Dokument verschlüsseln
Dim DruOk As Boolean
Dim AkPos As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo8 = FM.repCont8
Set RpCo3 = FM.repCont3
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6

Select Case TabRp
Case 1:
    Set RpRws = RpCo1.Rows
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
Case 2:
    Set RpRws = RpCo2.Rows
    Set RpCls = RpCo2.Columns
    Set RpSel = RpCo2.SelectedRows
Case 3:
    Set RpRws = RpCo3.Rows
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case 4:
    Set RpRws = RpCo4.Rows
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
Case 5:
    Set RpRws = RpCo5.Rows
    Set RpCls = RpCo5.Columns
    Set RpSel = RpCo5.SelectedRows
Case 6:
    Set RpRws = RpCo6.Rows
    Set RpCls = RpCo6.Columns
    Set RpSel = RpCo6.SelectedRows
Case 8:
    Set RpRws = RpCo8.Rows
    Set RpCls = RpCo8.Columns
    Set RpSel = RpCo8.SelectedRows
End Select

AktZa = 1
AkPos = InStrRev(FiNam, "\", Len(FiNam), 1)
DruDa = TmPfa & Mid$(FiNam, AkPos + 1, Len(FiNam) - (AkPos + 3)) & "LL"
DaExt = Right$(FiNam, 3)

If Gesam = True Then
    GesZa = RpRws.Count
Else
    GesZa = RpSel.Count
End If

If EmSen > 0 Then DrVor = True 'wenn Emailversand, dann Druckvorschau

If GesZa > 0 Then
    If Gesam = True Then
         Set RpRow = RpRws(0)
    Else
         Set RpRow = RpSel(0)
    End If
    If RpRow.GroupRow = True Then
        If Gesam = True Then
             Set RpRow = RpRws(1)
        Else
             Set RpRow = RpSel(1)
        End If
    End If
    
    Select Case GlBut
    Case RibTab_Rezeptmodul:
                PatNr = RpRow.Record(0).Value
                VerPa = S_AdIdi(PatNr, "Em_Pass")
                VerDo = S_AdIdi(PatNr, "Em_Aut")
    Case RibTab_Belegmodul:
                PatNr = RpRow.Record(0).Value
                VerPa = S_AdIdi(PatNr, "Em_Pass")
                VerDo = S_AdIdi(PatNr, "Em_Aut")
    End Select

    mHwnd = LlJobOpen(CMBTLANG_DEFAULT)
    LLOpt
    LLMail VerDo, VerPa

    Select Case LCase(DaExt):
    Case "lbl": 'Nur Variablen
            LDeVar RpRow, TabRp
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LABEL, FiNam, LL_PRINT_EXPORT, 0)
                End If
            End If
    Case "lst": 'Nur Listen
            LDeVar RpRow, TabRp
            LDeFel RpRow, TabRp
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
            End If
    Case "crd": 'Nur Variablen
            LDeVar RpRow, TabRp
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_CARD, FiNam, LL_PRINT_EXPORT, 0)
                End If
            End If
    Case "blg": 'Beides
            LDeVar RpRow, TabRp
            LDeFel RpRow, TabRp
            If DrVor = True Then 'Druckvorschau
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If GlRDP = True Then
                    RetWe = LlViewerProhibitAction(mHwnd, 116)
                End If
                RetWe = LlViewerProhibitAction(mHwnd, 115)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
                RetWe = LlPrintSetOptionString(mHwnd, LL_PRNOPTSTR_EXPORT, "PRV")
            Else
                RetWe = LlPreviewSetTempPath(mHwnd, TmPfa)
                If DrDia = True Then
                    RetWe = LlPrintWithBoxStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, LL_BOXTYPE_NORMALMETER, FM.hwnd, "Formulardruck")
                Else
                    RetWe = LlPrintStart(mHwnd, LL_PROJECT_LIST, FiNam, LL_PRINT_EXPORT, 0)
                End If
            End If
    End Select

    If RetWe = 0 Then
        If DrDia = True Then
            RetWe = LlPrintOptionsDialog(mHwnd, FM.hwnd, FiNam)
        End If
        If RetWe = 0 Then
            If DaExt = "lst" Then
                RetWe = LlPrint(mHwnd)
            End If
            RetWe = LlPrintGetPrinterInfo(mHwnd, DruNa, 100, PorNa, 100)
            If RetWe = 0 Then

                If DaExt = "lst" Then
                    If Gesam = True Then
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                LDeFel RpRow, TabRp
                                RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                            End If
                        Next RpRow
                    Else
                        For Each RpRow In RpSel
                            If RpRow.GroupRow = False Then
                                LDeFel RpRow, TabRp
                                RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                            End If
                        Next RpRow
                    End If
                ElseIf DaExt = "blg" Then
                    If Gesam = True Then
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                LDeFel RpRow, TabRp
                                RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                            End If
                        Next RpRow
                    Else
                        For Each RpRow In RpSel
                            If RpRow.GroupRow = False Then
                                LDeFel RpRow, TabRp
                                RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                While LlPrintFields(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                            End If
                        Next RpRow
                    End If
                Else
                    If Gesam = True Then
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                PatNr = RpRow.Record(0).Value
                                RzStr = Format$(RpRow.Record(1).Value, "00000")
                                RzDat = Format$(RpRow.Record(4).Value, "dd.mm.yyyy")
                                ManNr = RpRow.Record(11).Value
                                MitNr = RpRow.Record(71).Value
                                LDeVar RpRow, TabRp
                                RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                While LlPrint(mHwnd) = LL_WRN_REPEAT_DATA
                                    RetWe = LlPrint(mHwnd)
                                Wend
                                AktZa = AktZa + 1
                            End If
                        Next RpRow
                    Else
                        If Wiede = True Then
                            For AkWie = 1 To WieAn
                                LDeVar RpRow, TabRp
                                    RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AkWie / WieAn)))
                                    While LlPrint(mHwnd) = LL_WRN_REPEAT_DATA
                                        RetWe = LlPrint(mHwnd)
                                    Wend
                            Next AkWie
                        Else
                            For Each RpRow In RpSel
                                If RpRow.GroupRow = False Then
                                    PatNr = RpRow.Record(0).Value
                                    RzStr = Format$(RpRow.Record(1).Value, "00000")
                                    RzDat = Format$(RpRow.Record(4).Value, "dd.mm.yyyy")
                                    ManNr = RpRow.Record(11).Value
                                    MitNr = RpRow.Record(71).Value
                                    LDeVar RpRow, TabRp
                                    RetWe = LlPrintSetBoxText(mHwnd, "Ausgabe an: " & DruNa, (100 * (AktZa / GesZa)))
                                    While LlPrint(mHwnd) = LL_WRN_REPEAT_DATA
                                        RetWe = LlPrint(mHwnd)
                                    Wend
                                    AktZa = AktZa + 1
                                End If
                            Next RpRow
                        End If
                    End If
                End If
                
                If DaExt = "lst" Then
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                ElseIf DaExt = "blg" Then
                    Do
                    RetWe = LlPrintFieldsEnd(mHwnd)
                    Loop Until RetWe <> LL_WRN_REPEAT_DATA
                End If
                RetWe = LlPrintEnd(mHwnd, 0)
                If RetWe = 0 Then DruOk = True
            End If
        End If
    End If
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

If DrVor = True Then 'Druckvorschau
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            If VerDo = True Then
                hwOrg = LlStgsysStorageOpen(DruDa, TmPfa, False, True)
                RetWe = LlStgsysConvert(hwOrg, DaNam, "PDF;PDF.Encryption.EncryptFile=1;PDF.Encryption.Level=1;PDF.Encryption.EnableChanging=0;PDF.Encryption.EnablePrinting=1;PDF.OwnerPassword=" & PasOw & ";PDF.UserPassword=" & VerPa)
                LlStgsysStorageClose hwOrg
            Else
                RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            End If
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            ElseIf EmSen = 6 Then
                DoStr = SMCUp(DaNam, PatNr) 'Downloadlink
                DoEvents
                If DoStr <> vbNullString Then
                    Lange = Len(DoStr)
                    Posit = InStr(1, DoStr, ";", 1)
                    If Posit > 0 Then
                        DoLnk = Left$(DoStr, Posit - 1)
                        PoLnk = Mid$(DoStr, Posit + 1, Lange - Posit)
                        With GlTxV
                            Select Case GlBut
                            Case RibTab_Rezeptmodul: .TxStr = GlEmT(15, 3)
                            Case RibTab_Belegmodul: .TxStr = GlEmT(16, 3)
                            End Select
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = PatNr
                            .Datum = RzDat
                            .ReStr = RzStr
                        End With
                        EmBet = SEmTx()
                        With GlTxV
                            Select Case GlBut
                            Case RibTab_Rezeptmodul: .TxStr = GlEmT(15, 1)
                            Case RibTab_Belegmodul: .TxStr = GlEmT(16, 1)
                            End Select
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = PatNr
                            .Datum = RzDat
                            .ReStr = RzStr
                            .DoLnk = DoLnk
                            .PoLnk = PoLnk
                            .DoLan = GlVrw
                        End With
                        EmTex = EmTex & SEmTx()
                        SMaNe PatNr, EmAdr, , EmTex, EmBet
                    End If
                End If
            End If
        Else
            If GlVTy = 2 Then
                RetWe = LlPreviewDisplay(mHwnd, DruDa, TmPfa, FM.hwnd)
            End If
        End If
    Else
        SPopu "Fehler beim Drucken", "Beim Drucken ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
Else
    If DruOk = True Then
        If EmSen > 0 Then 'Emailversand
            DaNam = GlTmp & ForNa & "_" & Format$(Now, "YYYYMMDD_HHMM") & ".pdf"
            RetWe = LlStgsysStorageConvert(DruDa, DaNam, "PDF")
            If EmSen < 3 Then
                SMaNe PatNr, EmAdr, , EmTex, EmBet, DaNam
            End If
        End If
    Else
        SPopu "Fehler beim Exportieren", "Beim Exportieren ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
End If

DoEvents
LlJobClose mHwnd

If GlVTy = 2 Then
    Set clFil = New clsFile
    With clFil
        If .FilVor(DruDa) = True Then
            .DaLoe = DruDa & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
End If

If GesZa > 0 Then
    If EmSen = 0 Then 'Emailversand
        If DrVor = True Then 'Druckvorschau
            If DruOk = True Then
                Select Case GlVTy
                Case 1: VoDrMa DruDa
                Case 3: SDrVo DruDa
                End Select
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpRws = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCo8 = Nothing

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLPrUn " & Err.Number
Resume Next

End Sub
Private Sub LLReg()
On Error GoTo DeErr
'Liest Benutzerdaten aus der Registry

Dim RetWe As Long

RetWe = LlDefineVariableExt(mHwnd, "AnwVorname", IniGetVal("Adress", "AVoNa"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwName", IniGetVal("Adress", "AName"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwStraße", IniGetVal("Adress", "AStra"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwPLZ", IniGetVal("Adress", "APLZ"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwOrt", IniGetVal("Adress", "AOrt"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwTelefon", IniGetVal("Adress", "ATele"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwTelefax", IniGetVal("Adress", "AFax"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBank", IniGetVal("Adress", "ABank"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ", IniGetVal("Adress", "AnBLZ"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwKonto", IniGetVal("Adress", "AnKto"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwSteuer", IniGetVal("Adress", "ASteu"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBeruf", IniGetVal("Adress", "ABeru"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwTitel", IniGetVal("Adress", "ATite"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwIKNr", IniGetVal("Adress", "AIKNr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwLANR", IniGetVal("Adress", "AIKNr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwEmail", IniGetVal("Adress", "AEmail"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwInternet", IniGetVal("Adress", "AInter"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN", IniGetVal("Adress", "AIBANr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwPraxis", IniGetVal("Adress", "APraxis"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBank2", IniGetVal("Adress", "ABan2"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBLZ2", IniGetVal("Adress", "AnBL2"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwKonto2", IniGetVal("Adress", "AnKt2"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwSprechzeiten", IniGetVal("Adress", "ASpr1"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwIBAN2", IniGetVal("Adress", "AIBAN2"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBIC", IniGetVal("Adress", "ABICNr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBIC2", IniGetVal("Adress", "ABICN2"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwGID", IniGetVal("Adress", "AGlIDr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwAbsender", IniGetVal("Adress", "AGlIDr"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwGirocode", IniGetVal("Adress", "AnKto"), LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwGLN", "GLN", LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwZSR", "ZSR", LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwBunLan", "Bundesland", LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwKVBezirk", "KVBezirk", LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "AnwKanton", "Kanton", LL_TEXT, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Lizenziert", CInt(GlLiz), LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Druckkopie", DrKop, LL_BOOLEAN, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "Katalog", GlFri, LL_NUMERIC, vbNullString)
RetWe = LlDefineVariableExt(mHwnd, "TSESerien", GlTSN, LL_TEXT, vbNullString)

Exit Sub

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLReg " & Err.Number
Resume Next

End Sub
Private Function LLStor(ByVal VorDa As String, ByVal GuiNa As String, ByVal AnzPo As Integer) As Boolean
On Error GoTo DeErr
'Fasst einzelne Druckjobdateien zusammen

Dim RetWe As Long
Dim hwOrg As Long
Dim hwAnf As Long
Dim AnfDa As String
Dim AktPo As Integer

Set FM = frmMain
Set FS = frmStatus
Set PrBr1 = FS.prbStat1
Set PrBr2 = FS.prbStat2
Set TxDia = FS.txtDummy

Set clFil = New clsFile

If AnzPo > 1 Then
    PrBr2.Min = 0
    PrBr2.Max = AnzPo
    hwOrg = LlStgsysStorageOpen(VorDa, TmPfa, False, True)
    For AktPo = 1 To AnzPo
        AnfDa = GuiNa & "_" & Format$(AktPo, "0000") & ".LL"
        hwAnf = LlStgsysStorageOpen(AnfDa, TmPfa, False, True)
        RetWe = LlStgsysAppend(hwOrg, hwAnf)
        LlStgsysStorageClose hwAnf
        With clFil
            If GlVTy = 2 Then 'Druckvorschautyp
                If .FilVor(TmPfa & AnfDa) = True Then
                    .DaLoe = TmPfa & AnfDa & vbNullChar
                    .FilLoe
                End If
            End If
        End With
        PrBr2.Value = AktPo
        DoEvents
    Next AktPo
    LlStgsysStorageClose hwOrg
    LLStor = True
End If

Set clFil = Nothing

Exit Function

DeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "LLStor " & Err.Number
Resume Next

End Function
Public Property Let PfaTmp(Flag As String)
    TmPfa = Flag
End Property

Public Property Get PfaTmp() As String
    PfaTmp = TmPfa
End Property
Public Property Let Protok(Flag As Boolean)
    Proto = Flag
End Property
Public Property Get Protok() As Boolean
    Protok = Proto
End Property
Public Property Let StaVer(Flag As String)
    StVer = Flag
End Property

Public Property Get StaVer() As String
    StaVer = StVer
End Property
Public Property Let Wieder(Flag As Boolean)
    Wiede = Flag
End Property
Public Property Get Wieder() As Boolean
    Wieder = Wiede
End Property
Public Property Let WieAnz(Flag As Integer)
    WieAn = Flag
End Property
Public Property Get WieAnz() As Integer
    WieAnz = WieAn
End Property
Public Property Let ZeiRau(Flag As Integer)
    ZeRau = Flag
End Property
Public Property Get ZeiRau() As Integer
    ZeiRau = ZeRau
End Property

Private Sub Class_Initialize()
    Kopie = 1
End Sub
