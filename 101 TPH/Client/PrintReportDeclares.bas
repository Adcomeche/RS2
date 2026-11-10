Attribute VB_Name = "PrintReportDeclares"
Option Explicit

' DevMode structure and constants

Type DevMode
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
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'/* field selection bits */
Global Const DM_ORIENTATION = &H1
Global Const DM_PAPERSIZE = &H2
Global Const DM_PAPERLENGTH = &H4
Global Const DM_PAPERWIDTH = &H8
Global Const DM_SCALE = &H10
Global Const DM_COPIES = &H100
Global Const DM_DEFAULTSOURCE = &H200
Global Const DM_PRINTQUALITY = &H400
Global Const DM_COLOR = &H800
Global Const DM_DUPLEX = &H1000
Global Const DM_YRESOLUTION = &H2000
Global Const DM_TTOPTION = &H4000

'/* orientation selections */
Global Const DMORIENT_PORTRAIT = 1
Global Const DMORIENT_LANDSCAPE = 2

'/* paper selections */
'/*  Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' *  50 and 56.  Don't use this range when defining new paper sizes.
' */
Global Const DMPAPER_FIRST = 1
Global Const DMPAPER_LETTER = 1               '/* Letter 8 1/2 x 11 in               */
Global Const DMPAPER_LETTERSMALL = 2          '/* Letter Small 8 1/2 x 11 in         */
Global Const DMPAPER_TABLOID = 3              '/* Tabloid 11 x 17 in                 */
Global Const DMPAPER_LEDGER = 4               '/* Ledger 17 x 11 in                  */
Global Const DMPAPER_LEGAL = 5                '/* Legal 8 1/2 x 14 in                */
Global Const DMPAPER_STATEMENT = 6            '/* Statement 5 1/2 x 8 1/2 in         */
Global Const DMPAPER_EXECUTIVE = 7            '/* Executive 7 1/4 x 10 1/2 in        */
Global Const DMPAPER_A3 = 8                   '/* A3 297 x 420 mm                    */
Global Const DMPAPER_A4 = 9                   '/* A4 210 x 297 mm                    */
Global Const DMPAPER_A4SMALL = 10             '/* A4 Small 210 x 297 mm              */
Global Const DMPAPER_A5 = 11                  '/* A5 148 x 210 mm                    */
Global Const DMPAPER_B4 = 12                  '/* B4 250 x 354                       */
Global Const DMPAPER_B5 = 13                  '/* B5 182 x 257 mm                    */
Global Const DMPAPER_FOLIO = 14               '/* Folio 8 1/2 x 13 in                */
Global Const DMPAPER_QUARTO = 15              '/* Quarto 215 x 275 mm                */
Global Const DMPAPER_10X14 = 16               '/* 10x14 in                           */
Global Const DMPAPER_11X17 = 17               '/* 11x17 in                           */
Global Const DMPAPER_NOTE = 18                '/* Note 8 1/2 x 11 in                 */
Global Const DMPAPER_ENV_9 = 19               '/* Envelope #9 3 7/8 x 8 7/8          */
Global Const DMPAPER_ENV_10 = 20              '/* Envelope #10 4 1/8 x 9 1/2         */
Global Const DMPAPER_ENV_11 = 21              '/* Envelope #11 4 1/2 x 10 3/8        */
Global Const DMPAPER_ENV_12 = 22              '/* Envelope #12 4 \276 x 11           */
Global Const DMPAPER_ENV_14 = 23              '/* Envelope #14 5 x 11 1/2            */
Global Const DMPAPER_CSHEET = 24              '/* C size sheet                       */
Global Const DMPAPER_DSHEET = 25              '/* D size sheet                       */
Global Const DMPAPER_ESHEET = 26              '/* E size sheet                       */
Global Const DMPAPER_ENV_DL = 27              '/* Envelope DL 110 x 220mm            */
Global Const DMPAPER_ENV_C5 = 28              '/* Envelope C5 162 x 229 mm           */
Global Const DMPAPER_ENV_C3 = 29              '/* Envelope C3  324 x 458 mm          */
Global Const DMPAPER_ENV_C4 = 30              '/* Envelope C4  229 x 324 mm          */
Global Const DMPAPER_ENV_C6 = 31              '/* Envelope C6  114 x 162 mm          */
Global Const DMPAPER_ENV_C65 = 32             '/* Envelope C65 114 x 229 mm          */
Global Const DMPAPER_ENV_B4 = 33              '/* Envelope B4  250 x 353 mm          */
Global Const DMPAPER_ENV_B5 = 34              '/* Envelope B5  176 x 250 mm          */
Global Const DMPAPER_ENV_B6 = 35              '/* Envelope B6  176 x 125 mm          */
Global Const DMPAPER_ENV_ITALY = 36           '/* Envelope 110 x 230 mm              */
Global Const DMPAPER_ENV_MONARCH = 37         '/* Envelope Monarch 3.875 x 7.5 in    */
Global Const DMPAPER_ENV_PERSONAL = 38        '/* 6 3/4 Envelope 3 5/8 x 6 1/2 in    */
Global Const DMPAPER_FANFOLD_US = 39          '/* US Std Fanfold 14 7/8 x 11 in      */
Global Const DMPAPER_FANFOLD_STD_GERMAN = 40  '/* German Std Fanfold 8 1/2 x 12 in   */
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41  '/* German Legal Fanfold 8 1/2 x 13 in */

Global Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

Global Const DMPAPER_USER = 256

'/* bin selections */
Global Const DMBIN_FIRST = 1
Global Const DMBIN_UPPER = 1
Global Const DMBIN_ONLYONE = 1
Global Const DMBIN_LOWER = 2
Global Const DMBIN_MIDDLE = 3
Global Const DMBIN_MANUAL = 4
Global Const DMBIN_ENVELOPE = 5
Global Const DMBIN_ENVMANUAL = 6
Global Const DMBIN_AUTO = 7
Global Const DMBIN_TRACTOR = 8
Global Const DMBIN_SMALLFMT = 9
Global Const DMBIN_LARGEFMT = 10
Global Const DMBIN_LARGECAPACITY = 11
Global Const DMBIN_CASSETTE = 14
Global Const DMBIN_LAST = DMBIN_CASSETTE

Global Const DMBIN_USER = 256             '/* device specific bins start here */

'/* print qualities */
Global Const DMRES_DRAFT = -1
Global Const DMRES_LOW = -2
Global Const DMRES_MEDIUM = -3
Global Const DMRES_HIGH = -4

'/* color enable/disable for color printers */
Global Const DMCOLOR_MONOCHROME = 1
Global Const DMCOLOR_COLOR = 2

'/* duplex enable */
Global Const DMDUP_SIMPLEX = 1
Global Const DMDUP_VERTICAL = 2
Global Const DMDUP_HORIZONTAL = 3

'/* TrueType options */
Global Const DMTT_BITMAP = 1          '/* print TT fonts as graphics */
Global Const DMTT_DOWNLOAD = 2        '/* download TT fonts as soft fonts */
Global Const DMTT_SUBDEV = 3          '/* substitute device fonts for TT fonts */


