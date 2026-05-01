VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private mHwnd As Long    'Fensterhandle
Private mName As String  'Fontname
Private mGros As Long    'Fontgröße
Private mFarb As Long    'Fontfarbe
Private mFett As Boolean 'Fett
Private mKurs As Boolean 'Kursiv
Private mUnte As Boolean 'Unterstichen
Private mStrk As Boolean 'Durchgestrichen
Private mDrNa As String  'Druckername
Private mDhDc As String  'Drucker hDc
Private mDrPo As String  'Portname
Private mDrGe As String  'Gesamtname
Private mDrVo As Integer 'Drucke von
Private mDrBi As Integer 'Drucke bis
Private mDrKo As Integer 'Druckkopien
Private mDrQu As Integer 'Druckqualität
Private mDrSo As Boolean 'Drucksortierung

Private Const HWND_BROADCAST = &HFFFF
Private Const WM_SETTINGCHANGE = &H1A
Private Const WM_WININICHANGE = &H1A
Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8

' Printer attribute flags used with PRINTER_INFORMATION_2
Private Const PRINTER_ATTRIBUTE_QUEUED            As Long = &H1
Private Const PRINTER_ATTRIBUTE_DIRECT            As Long = &H2 'Der Drucker hat eine direkte
Private Const PRINTER_ATTRIBUTE_DEFAULT           As Long = &H4 'Der Drucker ist der 'Standarddrucker
Private Const PRINTER_ATTRIBUTE_SHARED            As Long = &H8 'Der Drucker ist ein Netzwerkdrucker
Private Const PRINTER_ATTRIBUTE_NETWORK           As Long = &H10
Private Const PRINTER_ATTRIBUTE_HIDDEN            As Long = &H20
Private Const PRINTER_ATTRIBUTE_LOCAL             As Long = &H40
Private Const PRINTER_ATTRIBUTE_ENABLE_DEVQ       As Long = &H80
Private Const PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS   As Long = &H100
Private Const PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST As Long = &H200 'Erledigt die Ausdrucke (Jobs)
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE      As Long = &H400
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI       As Long = &H800
Private Const PRINTER_ATTRIBUTE_RAW_ONLY          As Long = &H1000
Private Const PRINTER_ATTRIBUTE_PUBLISHED         As Long = &H2000

Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

'Flags für Fontdialog
Private Const CF_ANSIONLY = &H400           'Zeigt nur Windows- oder Unicode-Fonts
Private Const CF_APPLY = &H200              'Zeigt einen "Übernehmen" Button an
Private Const CF_BOTH = &H3                 'Listet Drucker- und Bildschirm-Fonts
Private Const CF_EFFECTS = &H100            'Erlaubt Fontbesonderheiten
Private Const CF_ENABLEHOOK = &H8           'Aktiviert die Callback-Funktion
Private Const CF_ENABLETEMPLATE = &H10      'Der Dialog benutzt Template's
Private Const CF_ENABLETEMPLATEHANDLE = &H20 'Verwendet den durch hInstance festgelegten Dialog
Private Const CF_FIXEDPITCHONLY = &H4000    'Listet nur Fixed-Pitch Fonts
Private Const CF_FORCEFONTEXIST = &H10000   'Verweigert die Eingabe nicht aufgeführter Fonts
Private Const CF_INITTOLOGFONTSTRUCT = &H40 'Setzt die Startwerte, welche über die LOGFONT-Struktur angegeben wurden
Private Const CF_LIMITSIZE = &H2000         'Erlaubt nur Schriftgrößen im Bereich "nSizeMin" und "nSizeMax"
Private Const CF_NOOEMFONTS = &H800         'Zeigt keine OEM Fonts
Private Const CF_NOFACESEL = &H80000        'Kein Standard Facenamen selektieren
Private Const CF_NOSCRIPTSEL = &H800000     'Kein Standard Script selektieren
Private Const CF_NOSIZESEL = &H200000       'keine Standardgröße setzen
Private Const CF_NOSIMULATIONS = &H1000     'Kein Beispiel (Vorschau) anzeigen
Private Const CF_NOSTYLESEL = &H100000      'kein Standard-Stil setzen
Private Const CF_NOVECTORFONTS = &H800      'keine Vector-Fonts anzeigen
Private Const CF_NOVERTFONTS = &H1000000    'keine vertikal ausgerichtete Fonts anzeigen
Private Const CF_PRINTERFONTS = &H2         'Listet Drucker-Fonts
Private Const CF_SCALABLEONLY = &H20000     'Listet nur skalierbare Fonts
Private Const CF_SCREENFONTS = &H1          'Listet Bildschirm-Fonts
Private Const CF_SCRIPTSONLY = &H400        'Listet nur Windows- oder Unicode-Fonts
Private Const CF_SELECTSCRIPT = &H400000    'Listet nur Script-Fonts
Private Const CF_SHOWHELP = &H4             'Zeigt den Hilfe-Button an
Private Const CF_TTONLY = &H40000           'Listet nur TrueType-Schriftarten
Private Const CF_USESTYLE = &H80            'Verwendet die in "lpStyle" angegebenen Werte
Private Const CF_WYSIWYG = &H8000           'Listet nur Fonts, die Drucker- und Bildschirm-Fonts gleichzeitig sind

'nFontType Konstanten
Private Const BOLD_FONTTYPE = &H100     'fett
Private Const ITALIC_FONTTYPE = &H200   'kursiv
Private Const PRINTER_FONTTYPE = &H4000 'Drucker-Fonts
Private Const REGULAR_FONTTYPE = &H400  'reguläre Fonts
Private Const SCREEN_FONTTYPE = &H2000  'Bildschirm-Fonts
Private Const SIMULATED_FONTTYPE = &H8000

'lfWeight Konstanten
Private Const FW_DONTCARE = 0      'Standard
Private Const FW_THIN = 100        'super dünn
Private Const FW_EXTRALIGHT = 200  'extra dünn
Private Const FW_LIGHT = 300       'dünn
Private Const FW_NORMAL = 400      'normal
Private Const FW_MEDIUM = 500      'mittel
Private Const FW_SEMIBOLD = 600    'etwas dicker
Private Const FW_BOLD = 700        'fett
Private Const FW_EXTRABOLD = 800   'extra fett
Private Const FW_HEAVY = 900       'super fett

'lfCharSet Konstanten
Private Const ANSI_CHARSET = 0          'Ansi Zeichensatz
Private Const ARABIC_CHARSET = 178      'Arabisch (NT/2000)
Private Const BALTIC_CHARSET = 186      'Baltisch (Win 9x)
Private Const CHINESEBIG5_CHARSET = 136 'Chinesisch
Private Const DEFAULT_CHARSET = 1       'Standard
Private Const EASTEUROPE_CHARSET = 238  'Osteuropäisch (Win 9x)
Private Const GB2312_CHARSET = 134      'Englisch
Private Const GREEK_CHARSET = 161       'Griechisch (Win 9x)
Private Const HANGEUL_CHARSET = 129     'Handgeul
Private Const HEBREW_CHARSET = 177      'Hebräisch (NT/2000)
Private Const JOHAB_CHARSET = 130       'Johab (Win 9x)
Private Const MAC_CHARSET = 77          'Mac (Win 9x)
Private Const OEM_CHARSET = 255         'OEM
Private Const RUSSIAN_CHARSET = 204     'Russisch (Win 9x)
Private Const SHIFTJIS_CHARSET = 128    'ShiftJis
Private Const SYMBOL_CHARSET = 2        'Symbolisch
Private Const THAI_CHARSET = 222        'Thailändisch (NT/2000)
Private Const TURKISH_CHARSET = 162     'Türkisch (Win 9x)

'lfOutPrecision Konstanten
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_OUTLINE_PRECIS = 8
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_TT_PRECIS = 4

'lfClipPrecision Konstanten
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_EMBEDDED = 128
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_STROKE_PRECIS = 2

'lfQuality Konstanten
Private Const ANTIALIASED_QUALITY = 4
Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const NONANTIALIASED_QUALITY = 3
Private Const PROOF_QUALITY = 2

'lfPitchAndFamily Konstanten
Private Const DEFAULT_PITCH = 0     'Standard
Private Const FIXED_PITCH = 1       'Fest
Private Const VARIABLE_PITCH = 2    'Variabel
Private Const FF_DECORATIVE = 80    'Dekoriert
Private Const FF_DONTCARE = 0       'Egal
Private Const FF_MODERN = 48        'Modern
Private Const FF_ROMAN = 16         'Roman
Private Const FF_SCRIPT = 64        'Script
Private Const FF_SWISS = 32         'Swiss

'PRINTDLG_TYPE flags-Konstanten
Private Const PD_ALLPAGES = &H0 ' selektiert den "Alle Seiten"-OptionButton
Private Const PD_COLLATE = &H10 ' selektiert die "Sortieren"-Checkbox
Private Const PD_DISABLEPRINTTOFILE = &H80000 '  "In Datei Drucken" deaktivieren
Private Const PD_ENABLEPRINTHOOK = &H1000 ' Hook-Prozedur soll benutzt werden
Private Const PD_ENABLEPRINTTEMPLATE = &H4000 ' benutzt Template-Namen
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000 ' benutzt Template-Handles
Private Const PD_ENABLESETUPHOOK = &H2000 ' benutzt eine Hook-Prozedur für den Setup-Dialog
Private Const PD_ENABLESETUPTEMPLATE = &H8000 ' benutzt einen Template-Namen für den
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000 ' benutzt ein Template-Handle
Private Const PD_HIDEPRINTTOFILE = &H100000 '  "In Datei Drucken"-Button verstecken
Private Const PD_NONETWORKBUTTON = &H200000 ' Netzwerk -Button verstecken
Private Const PD_NOPAGENUMS = &H8 '  "Von Bis" ist deaktiviert
Private Const PD_NOSELECTION = &H4 '  "Auswahl" ist deaktiviert
Private Const PD_NOWARNING = &H80 ' keine Warnungen wenn der gewählte Drucker nicht
Private Const PD_PAGENUMS = &H2 ' selektiert den "Von Bis"-OptionButton
Private Const PD_PRINTSETUP = &H40 ' öffnet den Drucker-Setup-Dialog
Private Const PD_PRINTTOFILE = &H20 ' selektiert die "In Datei Drucken"-Checkbox
Private Const PD_RETURNDC = &H100 ' gibt einen Device-Kontext des Druckers zurück
Private Const PD_RETURNDEFAULT = &H400 ' lädt den Standard-Drucker, dazu müssen bei
Private Const PD_RETURNIC = &H200 ' gibt einen Informations-Kontext zu dem gewählten
Private Const PD_SELECTION = &H1 ' selektiert die "Auswahl"-Option
Private Const PD_SHOWHELP = &H800 ' Zeigt den Hilfe-Button an
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000 ' berücksichtigt

'DEVMODE dmFields-Konstanten
Private Const DM_ORIENTATION = &H1 ' dmOrientation wird gefüllt/gelesen
Private Const DM_PAPERSIZE = &H2 ' dmPaperSize wird gefüllt/gelesen
Private Const DM_PAPERLENGTH = &H4 ' dmPaperLenght wird gefüllt/gelesen
Private Const DM_PAPERWIDTH = &H8 ' dmPaperWidth wird gefüllt/gelesen
Private Const DM_SCALE = &H10 ' dmScale wird gefüllt/gelesen
Private Const DM_COPIES = &H100 ' dmCopies wird gefüllt/gelesen
Private Const DM_DEFAULTSOURCE = &H200 ' dmDefaultSource wird gefüllt/gelesen
Private Const DM_PRINTQUALITY = &H400 ' dmPrintQuality wird gefüllt/gelesen
Private Const DM_COLOR = &H800 ' dmColor wird gefüllt/gelesen
Private Const DM_DUPLEX = &H1000 ' dmDuplex wird gefüllt/gelesen
Private Const DM_YRESOLUTION = &H2000 ' dmYResolution wird gefüllt/gelesen
Private Const DM_TTOPTION = &H4000 ' dmTTOption wird gefüllt/gelesen
Private Const DM_COLLATE = &H8000 ' dmCollate wird gefüllt/gelesen
Private Const DM_FORMNAME = &H10000 ' dmFormName wird gefüllt/gelesen
Private Const DM_LOGPIXELS = &H20000 ' dmLogPixels wird gefüllt/gelesen
Private Const DM_BITSPERPEL = &H40000 ' dmBitsPerPixel wird gefüllt/gelesen
Private Const DM_PELSWIDTH = &H80000 ' dmPelsWidth wird gefüllt/gelesen
Private Const DM_PELSHEIGHT = &H100000 ' dmPelsHeight wird gefüllt/gelesen
Private Const DM_DISPLAYFLAGS = &H200000 ' dmDisplayFlags wird gefüllt/gelesen
Private Const DM_DISPLAYFREQUENCY = &H400000 ' dmDisplayFrequency wird gefüllt/gelesen
Private Const DM_ICMMETHOD = &H800000 ' dmICMMethod wird gefüllt/gelesen
Private Const DM_ICMINTENT = &H1000000 ' dmICMIntent wird gefüllt/gelesen
Private Const DM_MEDIATYPE = &H2000000 ' dmMediaType wird gefüllt/gelesen
Private Const DM_DITHERTYPE = &H4000000 ' dmDitherType wird gefüllt/gelesen
Private Const DM_PANNINGWIDTH = &H20000000 ' dmPanningWidth wird gefüllt/gelesen
Private Const DM_PANNINGHEIGHT = &H40000000 ' dmPanningHeight wird gefüllt/gelesen

'DEVMODE dmOrientation-Konstanten
Private Const DMORIENT_PORTRAIT = 1 ' Portrait wurde gewählt (vertikal Drucken)
Private Const DMORIENT_LANDSCAPE = 2 ' Landscape wurde gewählt (horizontal Drucken)

'DEVMODE dmPaperSize-Konstanten
Private Const DMPAPER_LETTER = 1 ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_LEGAL = 5 ' Blatt 8,5 x 14 Inch
Private Const DMPAPER_10X11 = 45 ' 10 x 11 Inch
Private Const DMPAPER_10X14 = 16 ' 10 x 14 Inch
Private Const DMPAPER_11X17 = 17 ' 11 x 17 Inch
Private Const DMPAPER_15X11 = 46 ' 15 x 11 Inch
Private Const DMPAPER_9X11 = 44 ' 9 x 11 Inch
Private Const DMPAPER_A_PLUS = 57 ' A Plus Blatt
Private Const DMPAPER_A2 = 66 ' DIN A2 Blatt
Private Const DMPAPER_A3 = 8 ' DIN A3 Blatt
Private Const DMPAPER_A3_EXTRA = 63 ' DIN A3 Extra Blatt
Private Const DMPAPER_A3_EXTRA_TRANSVERSE = 68 ' DIN A3 Extra Blatt querlaufend
Private Const DMPAPER_A3_TRANSVERSE = 67 ' DIN A3 Blatt querlaufend
Private Const DMPAPER_A4 = 9 ' DIN A4 Blatt
Private Const DMPAPER_A4_EXTRA = 53 ' DIN A4 Extra Blatt
Private Const DMPAPER_A4_PLUS = 60 ' DIN A4 Plus Blatt
Private Const DMPAPER_A4_TRANSVERSE = 55 ' DIN A4 querlaufend
Private Const DMPAPER_A4SMALL = 10 ' DIN A4 klein (210 x 297 Millimeter) Blatt
Private Const DMPAPER_A5 = 11 ' DIN A5 Blatt
Private Const DMPAPER_A5_EXTRA = 64 ' DIN A5 Extra Blatt
Private Const DMPAPER_A5_TRANSVERSE = 61 ' DIN A5 querlaufend
Private Const DMPAPER_B_PLUS = 58 ' B Plus Blatt
Private Const DMPAPER_B4 = 12 ' B4 Blatt
Private Const DMPAPER_B5 = 13 ' B5 Blatt
Private Const DMPAPER_B5_EXTRA = 65 ' B5 Extra Blatt
Private Const DMPAPER_B5_TRANSVERSE = 62 ' B5 Blatt querlaufend
Private Const DMPAPER_CSHEET = 24 ' C Blatt (17 x 22 Inch)
Private Const DMPAPER_DSHEET = 25 ' D Blatt (22x34 Inch)
Private Const DMPAPER_ENV_10 = 20 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_11 = 21 ' Briefumschlag 11 (4,5 x 10,375 Inch)
Private Const DMPAPER_ENV_12 = 22 ' Briefumschlag 12 (4,75 x 11 Inch)
Private Const DMPAPER_ENV_14 = 23 ' Briefumschlag 14 (5 x 11,5 Inch)
Private Const DMPAPER_ENV_9 = 19 ' Briefumschlag 9 (3,875 x 8,875 Inch)
Private Const DMPAPER_ENV_B4 = 33 ' Briefumschlag B4 (250 x 353 Millimeter)
Private Const DMPAPER_ENV_B5 = 34 ' Briefumschlag B5 (176 x 250 Millimeter)
Private Const DMPAPER_ENV_B6 = 35 ' Briefumschlag B6 (176 x 125 Millimeter)
Private Const DMPAPER_ENV_C3 = 29 ' Briefumschlag C3 (324 x 458 Millimeter)
Private Const DMPAPER_ENV_C4 = 30 ' Briefumschlag C4 (229 x 324 Millimeter)
Private Const DMPAPER_ENV_C5 = 28 ' Briefumschlag C5 (162 x 229 Millimeter)
Private Const DMPAPER_ENV_C6 = 31 ' Briefumschlag C6 (114 x 162 Millimeter)
Private Const DMPAPER_ENV_C65 = 32 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_DL = 27 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_INVITE = 47 ' Einladungs Briefumschlag
Private Const DMPAPER_ENV_ITALY = 36 ' Italienischer Briefumschlag (110 x 230 Millimeter)
Private Const DMPAPER_ENV_MONARCH = 37 ' Monarchischer Briefumschlag (3,875 x 7,5 ' Inch)
Private Const DMPAPER_ENV_PERSONAL = 38 ' Persönlicher Briefumschlag (3,625 x 6,5 ' Inch)
Private Const DMPAPER_ESHEET = 26 ' E Blatt (34 x 44 Inch)
Private Const DMPAPER_EXECUTIVE = 7 ' Executive Blatt (7.25 x 10,5 Inch)
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40 ' Deutscher Standard Fanfold (8,5 x 12 ' Inch)
Private Const DMPAPER_FANFOLD_US = 39 ' US Standard Fanfold (14,875 + 11 Inch)
Private Const DMPAPER_FIRST = 1 ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_FOLIO = 14 ' Folin 8,5 + 13 Inch
Private Const DMPAPER_ISO_B4 = 42 ' ISO B4 Blatt
Private Const DMPAPER_JAPANESE_POSTCARD = 43 ' Japanische Postkarte
Private Const DMPAPER_LAST = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_LEDGER = 4 ' Ledger (17x11 Inch)
Private Const DMPAPER_LEGAL_EXTRA = 51 ' Rechtlich Extra
Private Const DMPAPER_LETTER_EXTRA = 50 ' Blatt Extra
Private Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56 ' Blatt Extra querlaufend
Private Const DMPAPER_LETTER_PLUS = 59 ' Blatt Plus
Private Const DMPAPER_LETTER_TRANSVERSE = 54 ' Blatt querlaufend
Private Const DMPAPER_LETTERSMALL = 2 ' Blatt klein (8,5 x 11 Inch)
Private Const DMPAPER_NOTE = 18 ' Note Blatt (8,5 x 11 Inch)
Private Const DMPAPER_QUARTO = 15 ' Quarto Blatt (215 x 275 Millimeter)
Private Const DMPAPER_STATEMENT = 6 ' Statement Blatt (5,5 x 8,5 Inch)
Private Const DMPAPER_TABLOID = 3 ' Tabloid Blatt (11 x 17 Inch)
Private Const DMPAPER_TABLOID_EXTRA = 52 ' Tabloid Extra Blatt
Private Const DMPAPER_USER = 256 ' benutzerdefinierte Größe

'DEVMODE dmDefaultSource-Konstanten
Private Const DMBIN_ONLYONE = 1 ' nur ein Blatt
Private Const DMBIN_UPPER = 1 ' Behälter Oben
Private Const DMBIN_LOWER = 2 ' Behälter Unten
Private Const DMBIN_MIDDLE = 3 ' Behälter Mitte
Private Const DMBIN_MANUAL = 4 ' Behälter mit manueller Füllung
Private Const DMBIN_ENVELOPE = 5 ' Briefumschlag Behälter
Private Const DMBIN_ENVMANUAL = 6 ' Briefumschlag Behälter mit Manueller Füllung
Private Const DMBIN_AUTO = 7 ' automatisches füllen
Private Const DMBIN_TRACTOR = 8 ' Blatt Einzug
Private Const DMBIN_SMALLFMT = 9 ' kleines Format laden
Private Const DMBIN_LARGEFMT = 10 ' großes Format laden
Private Const DMBIN_LARGECAPACITY = 11 ' große Kapazität
Private Const DMBIN_CASSETTE = 14 ' Blatt Kassette
Private Const DMBIN_FORMSOURCE = 15 ' Form Papier Quelle

'DEVMODE dmPrintQuality-Konstanten
Private Const DMRES_DRAFT = -1 ' Entwurf
Private Const DMRES_LOW = -2 '  niedrig
Private Const DMRES_MEDIUM = -3 ' mittel
Private Const DMRES_HIGH = -4 ' hoch

'DEVMODE dmColor-Konstanten
Private Const DMCOLOR_MONOCHROME = 1 ' Gerät unterstützt keine Schwarzweißausgabe
Private Const DMCOLOR_COLOR = 2 ' Gerät unterstützt keine Farbausgabe

'DEVMODE dmDuplex-Konstanten
Private Const DMDUP_SIMPLEX = 1 ' einseitiges Drucken
Private Const DMDUP_VERTICAL = 2 ' beidseitig Drucken bei vertikalem Seitenwechsel
Private Const DMDUP_HORIZONTAL = 3 ' beidseitig Drucken bei horizontalem Seitenwechsel

'DEVMODE dmTTOption-Konstanten
Private Const DMTT_BITMAP = 1 ' druckt TrueType-Fonts als Grafiken
Private Const DMTT_DOWNLOAD = 2 ' druckt Downloaded TrueType-Fonts als Soft-Fonts
Private Const DMTT_SUBDEV = 4 ' druckt Geräte-Fonts statt TrueType-Fonts

'DEVMODE dmCollate-Konstanten
Private Const DMCOLLATE_FALSE = 0 ' Drucker kann keine Seiten

Private Const DMCOLLATE_TRUE = 1 ' Drucker kann Seiten sortieren

'DEVMODE dmDisplayFlags-Konstanten
Private Const DM_GRAYSCALE = 1 ' Gerät unterstützt keine Farben, Grautöne werden unterstützt
Private Const DM_INTERLACED = 2 ' Gerät unterstützt Farben

'DEVMODE dmICMMethod-Konstanten
Private Const DMICMMETHOD_NONE = 1 ' ICM ist abgeschaltet
Private Const DMICMMETHOD_SYSTEM = 2 ' ICM wird von Windows gesteuert
Private Const DMICMMETHOD_DRIVER = 3 ' ICM wird vom Treiber gesteuert
Private Const DMICMMETHOD_DEVICE = 4 ' ICM wird vom Gerät gesteuert

'DEVMODE dmICMIntent-Konstanten
Private Const DMICM_SATURATE = 1 ' Sättigung wird optimiert
Private Const DMICM_CONTRAST = 2 ' Kontrast wird optimiert
Private Const DMICM_COLORMETRIC = 3 ' exakte Farben

'DEVMODE dmMediaType-Konstanten
Private Const DMMEDIA_STANDARD = 1 ' Standardpapier
Private Const DMMEDIA_GLOSSY = 2 ' Fotopapier
Private Const DMMEDIA_TRANSPARECNY = 3 ' Folie

'DEVMODE dmDitherType-Konstanten
Private Const DMDITHER_NONE = 1 ' kein Dithering
Private Const DMDITHER_COARSE = 2 ' Dithering mit grobem Brush
Private Const DMDITHER_FINE = 3 ' Dithering mit feinem Brush
Private Const DMDITHER_LINEART = 4 ' Linien mit schwarz, weiß und grau
Private Const DMDITHER_GRAYSCALE = 5 ' Grauskala

'GlobalAlloc wFlags-Konstanten
Private Const GHND = &H40 ' Kombination von GMEM_MOVEABLE mit GMEM_ZEROINIT
Private Const GMEM_DDESHARE = &H2000 ' optimiert den Speicher mit der DDE
Private Const GMEM_DISCARDABLE = &H100 ' Speicher kann überschrieben werden, kann nicht mit GMEM_FIXED kombiniert werden
Private Const GMEM_FIXED = &H0 ' fixiert den Speicher und die Funktion gibt einen Pointer zurück
Private Const GMEM_MOVEABLE = &H2 ' Der Speicher ist beweglich
Private Const GMEM_NOCOMPACT = &H10 ' Speicher wird nicht minimiert oder andere Blöcke für ihn überschrieben
Private Const GMEM_NODISCARD = &H20 ' überschreibt keine überschreibbaren Blöcke um  den Speicher zu reservieren
Private Const GMEM_ZEROINIT = &H40 ' initialisiert den Inhalt des Speicherblocks bei 0
Private Const GPTR = &H42 ' kombiniert GMEM_FIXED und GEMEM_ZEROINIT

Private Const DEVICE_FONTTYPE = &H2 'Der Font ist ein Device-Font
Private Const RASTER_FONTTYPE = &H1 'Der Font ist ein Raster-Font
Private Const TRUETYPE_FONTTYPE = &H4 'Der Font ist ein TrueType-Font

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Const DT_CALCRECT = &H400
Private Const DT_EDITCONTROL = &H2000
Private Const DT_NOCLIP = &H100

Private Const PRINTER_ENUM_CONNECTIONS = &H4 ' (Win NT) PRINTER_INFO_4-Konstante für Informationen über Netzwerkdrucker
Private Const PRINTER_ENUM_DEFAULT = &H1 ' (Win 9x) Ermittelt Informationen über den Standarddrucker des Systems
Private Const PRINTER_ENUM_LOCAL = &H2 ' Ermittelt alle Lokalen Drucker
Private Const PRINTER_ENUM_NAME = &H8 ' Ermittelt alle Drucker anhand des Netzwerknamens (Name-Variable der Funktion)
Private Const PRINTER_ENUM_NETWORK = &H40 ' (Win NT) PRINTER_INFO_4-Konstante. Ermittelt alle Netzwerkdrucker einer Domain
Private Const PRINTER_ENUM_REMOTE = &H10 ' (Win NT) PRINTER_INFO_4-Konstante. Ermittelt alle Netzwerkdrucker einer Domain
Private Const PRINTER_ENUM_SHARED = &H20 ' Ermittelt alle Drucker mit eingeschränktem Zugriff

Private Const ACL_REVISION = (2) 'AclRevision-Konstante
Private Const CCHFORMNAME = 32 'Pufferlänge-Konstante
Private Const CCHDEVICENAME = 32 'Pufferlänge-Konstante
Private Const PRINTER_STATUS_BUSY = &H200 ' Der Drucker ist nicht bereit
Private Const PRINTER_STATUS_DOOR_OPEN = &H400000 ' Die Klappe des Druckers ist geöffnet
Private Const PRINTER_STATUS_ERROR = &H2 ' Es liegt ein Druckerfehler vor
Private Const PRINTER_STATUS_INITIALIZING = &H8000 ' Der Drucker wird initialisiert
Private Const PRINTER_STATUS_IO_ACTIVE = &H100 ' Eine Datenkommunikation mit dem Drucker findet statt
Private Const PRINTER_STATUS_MANUAL_FEED = &H20 ' Der Drucker ladet Papier mit manuellen Papiereinzug nach
Private Const PRINTER_STATUS_NO_TONER = &H40000 ' Der Drucker hat keine Farbe mehr
Private Const PRINTER_STATUS_NOT_AVAILABLE = &H1000 ' Es kann nicht auf den Drucker zugegriffen werden
Private Const PRINTER_STATUS_OFFLINE = &H80 ' Der Drucker ist offline
Private Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000 ' Der Drucker hat nicht genügend Speicher zum Fortsetzen der letzten Aktion
Private Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800 ' Das Papierausgabefach ist voll
Private Const PRINTER_STATUS_PAGE_PUNT = &H80000 ' Der Druck wurde wegen zu hoher Komplexität abgebrochen
Private Const PRINTER_STATUS_PAPER_JAM = &H8 ' Das Druckerpapier kann nicht nachgezogen werden
Private Const PRINTER_STATUS_PAPER_OUT = &H10 ' Das Druckerpapier ist leer
Private Const PRINTER_STATUS_PAPER_PROBLEM = &H40 ' Es liegt ein Problem mit der Papierzufuhr vor
Private Const PRINTER_STATUS_PAUSED = &H1 ' (Win NT) Der Drucker befindet sich in einer Pause-Phase
Private Const PRINTER_STATUS_PENDING_DELETION = &H4 ' (Win NT) Die Ausdrucke in der Warteschleife werden gelöscht
Private Const PRINTER_STATUS_PRINTING = &H400 ' Der Drucker druckt
Private Const PRINTER_STATUS_PROCESSING = &H4000 ' Der Drucker erhält Informationen
Private Const PRINTER_STATUS_TONER_LOW = &H20000 ' Die Farbpatronen des Druckers gehen zur Neige
Private Const PRINTER_STATUS_USER_INTERVENTION = &H100000 ' Der Benutzer hat in die Druckeraktion eingegriffen
Private Const PRINTER_STATUS_WAITING = &H2000 ' Der Drucker wartet
Private Const PRINTER_STATUS_WARMING_UP = &H10000 ' Der Drucker wärmt sich auf EnumPrinters flags-Konstanten

Private m_iMinPage As Integer  ' Local copy of Min
Private m_iMaxPage As Integer  ' Local copy of Max
Private m_iFromPage As Integer ' Local copy of FromPage
Private m_iToPage As Integer   ' Local copy of ToPage

Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type
 
Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
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
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
 
Private Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As DEVMODE_TYPE
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

Private Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As DEVMODE_TYPE
    DesiredAccess As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Private Type CHOOSEFONT_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum PrinterConstants
    cdlPDAllPages = &H0
    cdlPDCollate = &H10
    cdlPDDisablePrintToFile = &H80000
    cdlPDHelpButton = &H800
    cdlPDHidePrintToFile = &H100000
    cdlPDNoPageNums = &H8
    cdlPDNoSelection = &H4
    cdlPDNoWarning = &H80
    cdlPDPageNums = &H2
    cdlPDPrintSetup = &H40
    cdlPDPrintToFile = &H20
    cdlPDReturnDC = &H100
    cdlPDReturnDefault = &H400
    cdlPDReturnIC = &H200
    cdlPDSelection = &H1
    cdlPDUseDevModeCopies = &H40000
End Enum

Private Enum ErrorConstants
    cdlCancel = 32755
End Enum

Public Event DruGef(ByVal DruNam As String, ByVal PorNam As String, ByVal DrvInf As String, ByVal Kommen As String, ByVal DruZa As Integer) 'WICHTIG!

Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (lpcf As CHOOSEFONT_TYPE) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Property Let DruBis(ByVal Flag As Integer)
Attribute DruBis.VB_Description = "Drucke bis Seite"
    mDrBi = Flag
End Property

Public Property Get DruBis() As Integer
    DruBis = mDrBi
End Property

Private Function DruCnv(Add As Long) As String
On Error Resume Next

Dim sTemp As String * 512
Dim RetWe As Long

RetWe = lstrcpy(sTemp, Add)
If (InStr(1, sTemp, Chr(0)) = 0) Then
     DruCnv = vbNullString
Else
     DruCnv = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
End If

End Function

Public Function DruDia() As Long
Attribute DruDia.VB_Description = "Zeigt den Druckdialog (benötigt: DruVon; DruBis)"
On Error Resume Next
'Ruft den Druckdialog auf

Dim RetWe As Long
Dim hMeMo As Long
Dim pMeMo As Long
Dim hMeNa As Long
Dim pMeNa As Long
Dim DevMo As DEVMODE_TYPE
Dim DevNa As DEVNAMES_TYPE
Dim PrDia As PRINTDLG_TYPE

With DevMo
    .dmSize = Len(DevMo)
    .dmDeviceName = Printer.DeviceName
    .dmFields = DM_PRINTQUALITY
End With

hMeMo = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMo))
pMeMo = GlobalLock(hMeMo)
MoveMemory ByVal pMeMo, DevMo, Len(DevMo)
GlobalUnlock hMeMo

With DevNa
    .wDriverOffset = Len(Printer.DriverName)
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .extra = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.Port & vbNullChar
End With

hMeNa = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevNa))
pMeNa = GlobalLock(hMeNa)
MoveMemory ByVal pMeNa, DevNa, Len(DevNa)
GlobalUnlock hMeNa

With PrDia
    .hInstance = App.hInstance
    .hwndOwner = mHwnd
    .lStructSize = Len(PrDia)
    .hDevMode = hMeMo
    .hDevNames = hMeNa
    .flags = PD_DISABLEPRINTTOFILE Or PD_RETURNDC Or PD_ALLPAGES Or PD_PAGENUMS Or PD_USEDEVMODECOPIESANDCOLLATE
    .nFromPage = mDrVo
    .nToPage = mDrBi
    .nMinPage = mDrVo
    .nMaxPage = mDrBi
End With

RetWe = PrintDlg(PrDia)
If RetWe = 0 Then
    GlobalFree PrDia.hDevMode
    GlobalFree PrDia.hDevNames
    Exit Function
End If

pMeMo = GlobalLock(PrDia.hDevMode)
MoveMemory DevMo, ByVal pMeMo, Len(DevMo)
GlobalUnlock PrDia.hDevMode
GlobalFree PrDia.hDevMode

pMeNa = GlobalLock(PrDia.hDevNames)
MoveMemory DevNa, ByVal pMeNa, Len(DevNa)
GlobalUnlock PrDia.hDevNames
GlobalFree PrDia.hDevNames

mDrNa = DevMo.dmDeviceName
mDrVo = PrDia.nFromPage
mDrBi = PrDia.nToPage
mDhDc = PrDia.hDC
mDrKo = DevMo.dmCopies
mDrSo = CBool(DevMo.dmCollate)

If CBool(DevMo.dmPrintQuality And DMRES_HIGH) = True Then
  mDrQu = 1
ElseIf CBool(DevMo.dmPrintQuality And DMRES_MEDIUM) = True Then
  mDrQu = 2
ElseIf CBool(DevMo.dmPrintQuality And DMRES_LOW) = True Then
  mDrQu = 3
ElseIf CBool(DevMo.dmPrintQuality And DMRES_DRAFT) = True Then
  mDrQu = 4
End If

DeleteObject PrDia.hDC

DruDia = RetWe

End Function

Public Sub DruErm(ByVal DruTyp As Integer)
On Error Resume Next

Dim AktZa As Long
Dim retval As Long
Dim Printers() As PRINTER_INFO_2
Dim TmpBuffer() As Long
Dim BuffSizeNeed As Long
Dim NumPrinters As Long

Select Case DruTyp
Case 1: 'Ermittelt alle Lokalen Drucker
    ReDim TmpBuffer(0) ' Alle Druckerinformationen ermitteln
    retval = EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, 2, TmpBuffer(0), (UBound(TmpBuffer) + 1) * 4, BuffSizeNeed, NumPrinters)
    ReDim TmpBuffer(BuffSizeNeed / 4)
    retval = EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, 2, TmpBuffer(0), (UBound(TmpBuffer) + 1) * 4, BuffSizeNeed, NumPrinters)
    ReDim Printers(NumPrinters - 1) 'Die Druckerinformationen manuell in die Strukturen kopieren
Case 2: 'Ermittelt alle Drucker anhand des Netzwerknamens (Name-Variable der Funktion)
    ReDim TmpBuffer(0) ' Alle Druckerinformationen ermitteln
    retval = EnumPrinters(PRINTER_ENUM_NAME, vbNullString, 2, TmpBuffer(0), (UBound(TmpBuffer) + 1) * 4, BuffSizeNeed, NumPrinters)
    ReDim TmpBuffer(BuffSizeNeed / 4)
    retval = EnumPrinters(PRINTER_ENUM_NAME, vbNullString, 2, TmpBuffer(0), (UBound(TmpBuffer) + 1) * 4, BuffSizeNeed, NumPrinters)
    ReDim Printers(NumPrinters - 1) 'Die Druckerinformationen manuell in die Strukturen kopieren
End Select

For AktZa = 0 To NumPrinters - 1
    With Printers(AktZa)
        DruVar .pServerName, TmpBuffer(AktZa * 21)
        DruVar .pPrinterName, TmpBuffer(AktZa * 21 + 1)
        DruVar .pShareName, TmpBuffer(AktZa * 21 + 2)
        DruVar .pPortName, TmpBuffer(AktZa * 21 + 3)
        DruVar .pDriverName, TmpBuffer(AktZa * 21 + 4)
        DruVar .pComment, TmpBuffer(AktZa * 21 + 5)
        DruVar .pLocation, TmpBuffer(AktZa * 21 + 6)
        MoveMemory .pDevMode, TmpBuffer(AktZa * 21 + 7), Len(.pDevMode)
        DruVar .pSepFile, TmpBuffer(AktZa * 21 + 8)
        DruVar .pPrintProcessor, TmpBuffer(AktZa * 21 + 9)
        DruVar .pDatatype, TmpBuffer(AktZa * 21 + 10)
        DruVar .pParameters, TmpBuffer(AktZa * 21 + 11)
        If TmpBuffer(AktZa * 21 + 12) < 0 Then
            MoveMemory .pSecurityDescriptor, ByVal TmpBuffer(AktZa * 21 + 12), Len(.pSecurityDescriptor)
        End If
        .Attributes = TmpBuffer(AktZa * 21 + 13)
        .Priority = TmpBuffer(AktZa * 21 + 14)
        .DefaultPriority = TmpBuffer(AktZa * 21 + 15)
        .StartTime = TmpBuffer(AktZa * 21 + 16)
        .UntilTime = TmpBuffer(AktZa * 21 + 17)
        .Status = TmpBuffer(AktZa * 21 + 18)
        .cJobs = TmpBuffer(AktZa * 21 + 19)
        .AveragePPM = TmpBuffer(AktZa * 21 + 20)
        RaiseEvent DruGef(.pPrinterName, .pPortName, .pDriverName, .pComment, AktZa)
    End With
Next AktZa

End Sub
Public Property Let DruGes(ByVal Flag As String)
Attribute DruGes.VB_Description = "Gesamtname des Druckers"
    mDrGe = Flag
End Property

Public Property Get DruGes() As String
    DruGes = mDrGe
End Property

Public Property Let DruhDc(ByVal Flag As Long)
    mDhDc = Flag
End Property


Public Property Get DruhDc() As Long
    DruhDc = mDhDc
End Property
Private Sub DruGet(ByVal Buffr As String, DrNam As String, PrtPo As String)
On Error Resume Next

Dim DrvNr As Integer
Dim DrPrt As Integer

DrNam = vbNullString
PrtPo = vbNullString

DrvNr = InStr(Buffr, ",")
If DrvNr > 0 Then
    DrNam = Left(Buffr, DrvNr - 1)
    DrPrt = InStr(DrvNr + 1, Buffr, ",")
    If DrPrt > 0 Then
        PrtPo = Mid(Buffr, DrvNr + 1, DrPrt - DrvNr - 1)
    End If
End If

End Sub

Public Sub DruInf()
Attribute DruInf.VB_Description = "Zeigt Informationen über den Standarddrucker"
On Error Resume Next
'Liefert Informationen über den Standarddrucker

Dim RetWe As Long
Dim BufWe As Long
Dim BufNe As Long
Dim Buffr As String

Buffr = Space(8192)
RetWe = GetProfileString("windows", "Device", "", Buffr, Len(Buffr))

If RetWe Then
    Buffr = Mid(Buffr, 1, RetWe)
    BufWe = InStr(Buffr, ",")
    mDrNa = Mid(Buffr, 1, BufWe - 1)
    BufNe = InStr(BufWe + 1, Buffr, ",")
    mDrGe = Mid(Buffr, BufWe + 1, BufNe - BufWe - 1)
    mDrPo = Mid(Buffr, BufNe + 1)
Else
    mDrNa = vbNullString
    mDrPo = vbNullString
    mDrGe = vbNullString
End If

End Sub

Public Property Let DruKop(ByVal Flag As Integer)
Attribute DruKop.VB_Description = "Druckkopien"
    mDrKo = Flag
End Property

Public Property Get DruKop() As Integer
    DruKop = mDrKo
End Property

Public Sub DruLst()
Attribute DruLst.VB_Description = "Listet alle Drucker auf"
On Error Resume Next
'Liefert Informationen über alle Drucker

Dim RetWe As Long
Dim ReStr As String
Dim Buffr As String
Dim DruNr As Integer
Dim AktZa As Integer

Dim BufWe As Long
Dim BufNe As Long
Dim PrNam As String
Dim DrNam As String
Dim PorNa As String

AktZa = 1
Buffr = Space(8192)
RetWe = GetProfileString("PrinterPorts", vbNullString, "", Buffr, Len(Buffr))

Do
DruNr = InStr(Buffr, Chr(0))
If DruNr > 0 Then
    ReStr = Left(Buffr, DruNr - 1)
    If Len(Trim(ReStr)) Then
        RaiseEvent DruGef(ReStr, vbNullString, vbNullString, vbNullString, AktZa)
    End If
    Buffr = Mid(Buffr, DruNr + 1)
Else
    If Len(Trim(Buffr)) Then
        RaiseEvent DruGef(Buffr, vbNullString, vbNullString, vbNullString, AktZa)
    End If
    Buffr = vbNullString
End If

AktZa = AktZa + 1
Loop While DruNr > 0

End Sub

Public Property Let DruNam(ByVal Flag As String)
Attribute DruNam.VB_Description = "Druckername"
    mDrNa = Flag
End Property

Public Property Get DruNam() As String
    DruNam = mDrNa
End Property

Public Property Let DruPor(ByVal Flag As String)
Attribute DruPor.VB_Description = "Portname"
    mDrPo = Flag
End Property

Public Property Get DruPor() As String
    DruPor = mDrPo
End Property

Public Property Let DruQua(ByVal Flag As Integer)
Attribute DruQua.VB_Description = "Druckqualität: 1=Hoch 2=Mittel 3=Niedrig 4=Entwurf"
    mDrQu = Flag
End Property

Public Property Get DruQua() As Integer
    DruQua = mDrQu
End Property

Public Sub DruSet()
Attribute DruSet.VB_Description = "Setzt den Standarddrucker"
On Error Resume Next
'Setzt den Standarddrucker

Dim RetWe As Long
Dim OsInf As OSVERSIONINFO

OsInf.dwOSVersionInfoSize = Len(OsInf)
OsInf.szCSDVersion = Space$(128)
RetWe = GetVersionEx(OsInf)

Select Case OsInf.dwPlatformId
Case VER_PLATFORM_WIN32_WINDOWS:
    DruW95
Case VER_PLATFORM_WIN32_NT:
    DruWNT OsInf.dwMajorVersion
End Select

End Sub

Public Property Let DruSor(ByVal Flag As Boolean)
Attribute DruSor.VB_Description = "Drucksortierung"
    mDrSo = Flag
End Property

Public Property Get DruSor() As Boolean
    DruSor = mDrSo
End Property

Private Sub DruStd(ByVal PrtNa As String, ByVal DrNam As String, ByVal PrtPo As String)
On Error Resume Next

Dim DevLn As String
Dim RetWe As Long
Dim SetWe As Long

DevLn = PrtNa & "," & DrNam & "," & PrtPo
RetWe = WriteProfileString("windows", "Device", DevLn)
SetWe = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")

End Sub
Public Sub DruTex(ByVal TxStr As String)
On Error Resume Next

Dim RetWe As Long
Dim MyRct As RECT
 
Printer.Font.SIZE = GlTFt.SIZE
Printer.ScaleMode = vbPixels

'Achtung! Die Angaben müssen in Pixeln gemacht werden!
MyRct.Left = 400
MyRct.Top = 200
MyRct.Right = Printer.ScaleWidth - 400
MyRct.Bottom = Printer.ScaleHeight - 400

Printer.Print ""

RetWe = DrawText(Printer.hDC, TxStr, Len(TxStr), MyRct, DT_LEFT Or DT_WORDBREAK)

Printer.EndDoc

End Sub
Private Function DruVar(ByRef StrDest As String, ByRef VarSrc As Long)
On Error Resume Next
'Kopiert einen String anhand eines Pointers in eine String-Variable

StrDest = Space(lstrlen(VarSrc))

MoveMemory ByVal StrDest, ByVal VarSrc, Len(StrDest)

End Function
Public Property Get DruVon() As Integer
Attribute DruVon.VB_Description = "Drucke von Seite"
    DruVon = mDrVo
End Property

Private Sub DruW95()
On Error Resume Next

Dim PrHan As Long
Dim RetWe As Long
Dim ZwWrt As Long
Dim PrtNa As String
Dim PrInf As PRINTER_INFO_5
Dim PrDev As PRINTER_DEFAULTS

PrtNa = mDrNa
If PrtNa = vbNullString Then
    Exit Sub
End If

PrDev.pDatatype = 0&
PrDev.DesiredAccess = PRINTER_ALL_ACCESS Or PrDev.DesiredAccess
RetWe = OpenPrinter(PrtNa, PrHan, PrDev)
If RetWe = False Then
    Exit Sub
End If

RetWe = GetPrinter(PrHan, 5, ByVal 0&, 0, ZwWrt)
ReDim T((ZwWrt \ 4)) As Long

RetWe = GetPrinter(PrHan, 5, T(0), ZwWrt, ZwWrt)
If RetWe = False Then
    Exit Sub
End If

PrInf.pPrinterName = DruCnv(T(0))
PrInf.pPortName = DruCnv(T(1))
PrInf.Attributes = T(2)
PrInf.DeviceNotSelectedTimeout = T(3)
PrInf.TransmissionRetryTimeout = T(4)

PrInf.Attributes = PRINTER_ATTRIBUTE_DEFAULT

RetWe = SetPrinter(PrHan, 5, PrInf, 0)
If RetWe = False Then
    Exit Sub
End If

ClosePrinter PrHan

End Sub

Private Sub DruWNT(ByVal HauVer As Long)
On Error Resume Next

Dim RetWe As Long
Dim Buffr As Long
Dim Komma As Long
Dim Ergeb As String

If HauVer >= 5 Then
   RetWe = SetDefaultPrinter(mDrNa)
Else
    Buffr = 1024
    Ergeb = Space$(Buffr)
    If GetProfileString("PrinterPorts", ByVal mDrNa, "", Ergeb, Buffr) Then
        Komma = InStr(Ergeb, ",")
        Komma = InStr(Komma + 1, Ergeb, ",")
        If Komma Then
            Ergeb = Left$(Ergeb, Komma - 1)
            Ergeb = mDrNa & "," & Ergeb
            RetWe = WriteProfileString("Windows", "device", Ergeb)
            RetWe = SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0, SMTO_NORMAL, 500, ByVal 0&)
        End If
    End If
End If

End Sub

Public Function FntDia() As Long
Attribute FntDia.VB_Description = "Zeigt den Fontauswahldialog an"
On Error Resume Next
'Zeigt den Fontdialog

Dim RetWe As Long
Dim TmpSt As String
Dim Auswa As CHOOSEFONT_TYPE
Static LogFn As LOGFONT

With Auswa
    .flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    .hwndOwner = mHwnd
    .lStructSize = Len(Auswa)
    .lpLogFont = VarPtr(LogFn)
    .hInstance = App.hInstance
    .hDC = Printer.hDC
    .nFontType = SCREEN_FONTTYPE
    .rgbColors = 0
End With

If mName = vbNullString Then
    LogFn.lfFaceName = Space(31) & vbNullChar
Else
    LogFn.lfFaceName = mName & vbNullChar
End If

If mGros <> 0 Then
    LogFn.lfHeight = mGros
Else
    LogFn.lfHeight = 10
End If

RetWe = ChooseFont(Auswa)
TmpSt = StrConv(LogFn.lfFaceName, vbUnicode)

mName = Left$(TmpSt, InStr(1, TmpSt, vbNullChar) - 1)
mGros = Auswa.iPointSize / 10
mFarb = Auswa.rgbColors
mFett = CBool(LogFn.lfWeight >= FW_BOLD)
mKurs = CBool(LogFn.lfItalic)
mUnte = CBool(LogFn.lfUnderline)
mStrk = CBool(LogFn.lfStrikeOut)

FntDia = RetWe

End Function
Public Property Let FoBol(ByVal Flag As Boolean)
Attribute FoBol.VB_Description = "Font Bold"
    mFett = Flag
End Property

Public Property Get FoBol() As Boolean
    FoBol = mFett
End Property

Public Property Let FoFar(ByVal Flag As Long)
Attribute FoFar.VB_Description = "Font Farbe"
    mFarb = Flag
End Property

Public Property Get FoFar() As Long
    FoFar = mFarb
End Property

Public Property Let FoGro(ByVal Flag As Long)
Attribute FoGro.VB_Description = "Font Größe"
    mGros = Flag
End Property

Public Property Get FoGro() As Long
    FoGro = mGros
End Property

Public Property Let FoKur(ByVal Flag As Boolean)
Attribute FoKur.VB_Description = "Font Kursiv"
    mKurs = Flag
End Property

Public Property Get FoKur() As Boolean
    FoKur = mKurs
End Property

Public Property Let FoNam(ByVal Flag As String)
Attribute FoNam.VB_Description = "Font Name"
    mName = Flag
End Property

Public Property Get FoNam() As String
    FoNam = mName
End Property

Public Property Let FoStk(ByVal Flag As Boolean)
Attribute FoStk.VB_Description = "Font Strike"
    mStrk = Flag
End Property

Public Property Get FoStk() As Boolean
    FoStk = mStrk
End Property

Public Property Let FoUnt(ByVal Flag As Boolean)
Attribute FoUnt.VB_Description = "Font Unterstrichen"
    mUnte = Flag
End Property

Public Property Get FoUnt() As Boolean
    FoUnt = mUnte
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Fensterhandle"
    hwnd = mHwnd
End Property

Public Property Let hwnd(ByVal Flag As Long)
    mHwnd = Flag
End Property
