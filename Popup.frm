VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Popup 
   Caption         =   "Form2"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2715
   LinkTopic       =   "Form2"
   Picture         =   "Popup.frx":0000
   ScaleHeight     =   1230
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1560
      Top             =   1560
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3000
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Established a new connection"
      Height          =   255
      Left            =   -240
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)
Private Sub Form_Load()
Dim t As Single
Dim rtn As Long
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 10
t = Timer

pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY

Label1.Caption = Right(PopupReason, Len(PopupReason) - InStrRev(PopupReason, "\"))
iml32.ListImages.Clear
GetIcon PopupReason

If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If
Timer1.Enabled = True
Me.Visible = True

 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Unload Me

'If Button = vbLeftButton Then
'  ReleaseCapture
'  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Unload Me

'If Button = vbLeftButton Then
'  ReleaseCapture
'  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End If
End Sub

Private Sub Timer1_Timer()
If Me.Visible = True Then
Me.Visible = False
Timer1.Enabled = False
Unload Me
End If
End Sub

Private Function GetIcon(FileName As String) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
'Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  'With pic16
  '  Set .Picture = LoadPicture("")
  '  .AutoRedraw = True
  '  r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
  '  .Refresh
  'End With
  'Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  'Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function


