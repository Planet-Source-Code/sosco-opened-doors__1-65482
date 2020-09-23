VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Opened Doors"
   ClientHeight    =   7800
   ClientLeft      =   450
   ClientTop       =   1140
   ClientWidth     =   12675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   12675
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   10320
      Top             =   7080
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   3225
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   360
         Top             =   120
      End
      Begin VB.ListBox lstConnection 
         Appearance      =   0  'Flat
         Height          =   6465
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "local"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   5760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer tmrUpdate 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   4560
         Top             =   5280
      End
      Begin MSWinsockLib.Winsock wsMain 
         Index           =   0
         Left            =   4080
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Image Image3 
         Height          =   195
         Left            =   80
         Picture         =   "Form1.frx":2D16
         Top             =   460
         Width           =   195
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   3045
         TabIndex        =   17
         Top             =   840
         Width           =   75
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   3120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label cmdScan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start Scaning"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   7800
         Width           =   3015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ports found:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   2400
         X2              =   2400
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP you are looking for:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Left            =   -240
         TabIndex        =   11
         Top             =   2880
         Width           =   240
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11760
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   8040
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock wins 
      Index           =   0
      Left            =   11280
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   7575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7545
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "Monitor state: "
            TextSave        =   "Monitor state: "
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5760
      Top             =   3360
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2820
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   4440
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   5040
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5040
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "iml32"
      SmallIcons      =   "iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Direction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Host"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "File Path"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView picture2 
      Height          =   4260
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   7514
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   11040
      Picture         =   "Form1.frx":2F60
      Top             =   8160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   10800
      Picture         =   "Form1.frx":3062
      Top             =   8160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu amr 
         Caption         =   "Automatic Refresh"
      End
      Begin VB.Menu Refreshlist 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ViewCon 
         Caption         =   "View Connections"
      End
      Begin VB.Menu seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu ShowPop 
         Caption         =   "Show Popup"
         Checked         =   -1  'True
      End
      Begin VB.Menu pcnn 
         Caption         =   "Popup Connection"
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu stm 
         Caption         =   "Start Monitor"
      End
      Begin VB.Menu shl 
         Caption         =   "Show Log"
      End
      Begin VB.Menu seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu snifer 
         Caption         =   "Snifer"
         Shortcut        =   {F3}
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu MinSystray 
         Caption         =   "Minimise to System tray"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu viewt 
      Caption         =   "View"
      Begin VB.Menu lup 
         Caption         =   "Local used ports"
         Checked         =   -1  'True
      End
      Begin VB.Menu lpc 
         Caption         =   "Local processes"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu uup 
         Caption         =   "Used Ports"
      End
      Begin VB.Menu l7 
         Caption         =   "-"
      End
      Begin VB.Menu ht 
         Caption         =   "How To.."
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O
Dim askk, ast, asf
Dim catTop
'Icon Sizes in pixels




Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
    End Type




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
    ByVal X&, ByVal y&, ByVal Flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO
Public Tablenum As Long
Private pTcpTable As MIB_TCPTABLE

Public Sub RefreshView()
  Dim i As Integer, o As Integer
  Dim fileNum As String
  Dim Item As ListItem
  
On Error Resume Next
  ListView1.ListItems.Clear
   
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    iml32.ListImages.Clear
    iml16.ListImages.Clear
    
    DoEvents
  LoadProcesses
  
  StatusBar1.Panels(1).Text = Winsock1.LocalHostName & " : " & Winsock1.LocalIP
  StatusBar1.Panels(2).Text = "Last Refresh - " & Time
  
  For i = 0 To StatsLen - 1
  
  If Connection(i).FileName <> "" Then Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))) Else Set Item = ListView1.ListItems.Add(, , "Unknown")
    
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(1) = "Incomming" Else Item.SubItems(1) = "Outgoing"
    Item.SubItems(2) = Connection(i).LocalPort
    Item.SubItems(3) = Connection(i).RemoteHost
    Item.SubItems(4) = Connection(i).RemotePort
    Item.SubItems(5) = Connection(i).State
    Item.SubItems(6) = Connection(i).FileName
    
    'Item.EnsureVisible
    
  Next
  
  DoEvents
  GetAllIcons
  ShowIcons

  DoEvents
    Me.MousePointer = vbNormal
    'Label2.Caption = "Netstat status as of: " & Date & " " & Time
End Sub



Private Sub amr_Click()


If amr.Checked = False Then
    SaveSetting "@Revive", "Settings", "ref", "0"
    amr.Checked = True
    Timer3.Enabled = True
Else
        SaveSetting "@Revive", "Settings", "ref", "1"
    Timer3.Enabled = False
    amr.Checked = False
End If
End Sub

Private Sub cmdCheck_Click()
lstConnection.Clear
Dim i
For i = txtFrom.Text To txtTo.Text
If wsMain(i).State <> sckConnected Then
'lstConnection.AddItem i & " ---- Could not connect"

lstConnection.AddItem i & " :: false"
ElseIf wsMain(i).State = sckConnected Then
'lstConnection.AddItem i & " :: << TRUE >>"

    If i = "7" Then lstConnection.AddItem i & " :: [TRUE] -->ECHO Port"
  If i = "19" Then lstConnection.AddItem i & " :: [TRUE] -->CHARGE Port"
  If i = "20" Then lstConnection.AddItem i & " :: [TRUE] -->FTP Port"
  If i = "21" Then lstConnection.AddItem i & " :: [TRUE] -->FTP Port"
  If i = "22" Then lstConnection.AddItem i & " :: [TRUE] -->SSH Port"
  If i = "23" Then lstConnection.AddItem i & " :: [TRUE] -->TELNET Port"
  If i = "25" Then lstConnection.AddItem i & " :: [TRUE] -->SMTP Port"
  If i = "53" Then lstConnection.AddItem i & " :: [TRUE] -->DNS Port"
  If i = "79" Then lstConnection.AddItem i & " :: [TRUE] -->FINGER Port"
  If i = "80" Then lstConnection.AddItem i & " :: [TRUE] -->HTTP Port"
If i = "110" Then lstConnection.AddItem i & " :: [TRUE] -->POP3"
If i = "137" Then lstConnection.AddItem i & " :: [TRUE] -->NBIOS"
If i = "138" Then lstConnection.AddItem i & " :: [TRUE] -->NBIOS"
If i = "139" Then lstConnection.AddItem i & " :: [TRUE] -->NBIOS"
If i = "220" Then lstConnection.AddItem i & " :: [TRUE] -->IMAP"
If i = "443" Then lstConnection.AddItem i & " :: [TRUE] -->HTTPS(SSL)"
If i = "1433" Then lstConnection.AddItem i & " :: [TRUE] -->SQL Server"
If i = "1434" Then lstConnection.AddItem i & " :: [TRUE] -->SQL Monitor"
If i = "1080" Then lstConnection.AddItem i & " :: [TRUE] -->Internet Proxy"
If i = "1863" Then lstConnection.AddItem i & " :: [TRUE] -->MSN"
If i = "5050" Then lstConnection.AddItem i & " :: [TRUE] -->Yahoo Chat"
If i = "5100" Then lstConnection.AddItem i & " :: [TRUE] -->Yahoo WebCam"
If i = "5000" Then lstConnection.AddItem i & " :: [TRUE] -->yahoo Voice Chat"
If i = "5001" Then lstConnection.AddItem i & " :: [TRUE] -->Yahoo Voice Chat"
If i = "5101" Then lstConnection.AddItem i & " :: [TRUE] -->Yahoo P2P Messages"



End If
Next
Label6.Caption = "0"
End Sub

Private Sub cmdScan_Click()
On Error Resume Next
Release
lstConnection.Clear
Dim Port As Long
For Port = txtFrom.Text To txtTo.Text
Load wsMain(Port)
wsMain(Port).Connect txtIP.Text, Port
Label6.Caption = Port
If Port >= txtTo.Text Then
Label6.Caption = "Waiting..."
tmrUpdate.Enabled = True
End If
Next
End Sub

Private Sub Exit_Click()
TrayDelete
Unload Me
End Sub



Private Sub Form_Load()
Me.Visible = False
catTop = 0
Dim abba
askk = GetSetting("@Revive", "Settings", "ask", "0")
ast = GetSetting("@Revive", "Settings", "ast", "0")
abba = GetSetting("@Revive", "Settings", "ref", "0")

StatusBar1.Panels(3).Text = "Monitor state is Off"
StatusBar1.Panels(3).Picture = Image1.Picture
If askk = 0 Then
    pcnn.Checked = False
    Else: pcnn.Checked = True
End If
If ast = 0 Then
    ShowPop.Checked = False
    Else: ShowPop.Checked = True
End If
If abba = 0 Then
    amr.Checked = False
    Else: amr.Checked = True
End If

Open App.Path & "\ping.htm" For Output As #1
Text6.Text = "<font size='+3' color = '#B21804'><u><b>" & Winsock1.LocalHostName & " 's Log </u></b></font><br><br>"
Print #1, Text6.Text
Close #1

pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY


 picture2.View = lvwReport

    'Add the Column Headers for your ListView Control
    picture2.ColumnHeaders.Add 1, "SessionID", "Session ID"
    picture2.ColumnHeaders.Add 2, "ProcessID", "Process ID"
    picture2.ColumnHeaders.Add 3, "ProcessName", "Process Name"
    picture2.ColumnHeaders(3).Width = 3500
    picture2.ColumnHeaders.Add 4, "UserID", "User ID"
    picture2.ColumnHeaders(4).Width = picture2.Width - (picture2.ColumnHeaders(1).Width * 3) - 300

    GetWTSProcesses




DoEvents
RefreshView





End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim cEvent As Single
cEvent = X / Screen.TwipsPerPixelX
Select Case cEvent
    Case RightUp
        PopupMenu Menu
End Select
End Sub

Private Sub Form_Resize()

On Error Resume Next
ListView1.Width = Me.Width - 130
ListView1.Left = 0
ListView1.Height = Me.Height - 1050
picture2.Left = ListView1.Left
picture2.Width = ListView1.Width
picture2.Height = ListView1.Height

Picture1.Top = ListView1.Top + 250
Picture1.Left = ListView1.Left
Picture1.Height = ListView1.Height - 400
cmdScan.Top = Picture1.Height - 400
lstConnection.Height = cmdScan.Top - 1200
ListView1.ColumnHeaders(1).Width = 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = 1300
ListView1.ColumnHeaders(7).Width = ListView1.Width \ 2 + 1000


End Sub

Private Sub Image3_Click()
lstConnection.Clear
Timer2.Enabled = True
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image3.ToolTipText = "ping -t " & Winsock1.LocalIP
End Sub

Private Sub ListView1_DblClick()
    snifer_Click
    txtIP.Text = ListView1.SelectedItem.SubItems(3)

End Sub

Private Sub lpc_Click()

Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In picture2.ListItems
  FileName = Item.SubItems(Item.ListSubItems.Count) ' & Item.Text
  GetIcon FileName, Item.Index
Next




lpc.Checked = True
lup.Checked = False
picture2.Visible = True
ListView1.Visible = False

End Sub

Private Sub lup_Click()
lpc.Checked = False
lup.Checked = True
picture2.Visible = False
ListView1.Visible = True
End Sub

Private Sub MinSystray_Click()
If MinSystray.Checked = False Then
    TrayAdd hwnd, Me.Icon, "System Tray", MouseMove
    MinSystray.Caption = "Restore from System tray"
    MinSystray.Checked = True
    Me.WindowState = 1
    Me.Hide
    Exit Sub
Else
    Me.WindowState = 0
    Me.Show
    TrayDelete
    MinSystray.Caption = "Minimize to System tray"
    MinSystray.Checked = False
    Exit Sub
End If
End Sub

Private Sub pcnn_Click()
If pcnn.Checked = False Then
    pcnn.Checked = True
    SaveSetting "@Revive", "Settings", "ask", "1"
    askk = 0
Else
    pcnn.Checked = False
    SaveSetting "@Revive", "Settings", "ask", "0"
    askk = 1
End If

End Sub

Private Sub Refreshlist_Click()
  RefreshView
  GetWTSProcesses
End Sub

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Item.SubItems(Item.ListSubItems.Count) ' & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
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
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub shl_Click()
Form3.web.Navigate App.Path & "\ping.htm"
Form3.Show
End Sub

Private Sub ShowPop_Click()
If ShowPop.Checked = True Then
SaveSetting "@Revive", "Settings", "ast", "0"
ShowPopup = False
ShowPop.Checked = False
Else
SaveSetting "@Revive", "Settings", "ast", "1"
ShowPopup = True
ShowPop.Checked = True
End If
End Sub

Private Sub snifer_Click()
lstConnection.Clear


If snifer.Checked = False Then
    Picture1.Visible = True
    snifer.Checked = True
    Exit Sub
Else
    Picture1.Visible = False
    snifer.Checked = False
    Exit Sub
End If
End Sub

Private Sub stm_Click()
If stm.Checked = False Then
    Dim kk
    List1.ListIndex = kk
    For kk = 0 To 3
        List1.ListIndex = kk
        Load wins(kk + 1)
        wins(kk + 1).LocalPort = List1.Text
        wins(kk + 1).Listen
 
    Next kk
    stm.Checked = True
    Me.StatusBar1.Panels(3).Text = "Monitor state is On"
    Me.StatusBar1.Panels(3).Picture = Image2.Picture
    stm.Caption = "Stop Monitor"
    Exit Sub
Else
    Dim kki
    List1.ListIndex = kki
    For kki = 0 To 3
        Unload wins(kki + 1)

    Next kki

    stm.Checked = False
     Me.StatusBar1.Panels(3).Text = "Monitor state is Off"
     Me.StatusBar1.Panels(3).Picture = Image1.Picture
     stm.Caption = "Start Monitor"
    Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
'RefreshView
 
Dim pdwSize As Long
Dim bOrder As Long
Dim nRet As Long
Dim TableLen As Long

nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)

TableLen = pTcpTable.dwNumEntries
If Tablenum <> TableLen Then RefreshView
Tablenum = TableLen

End Sub

Private Sub Timer2_Timer()
lstConnection.Clear
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   
  'ping an ip address, passing the
  'address and the ECHO structure
   Call Ping(txtIP.Text, ECHO)
   
  'display the results from the ECHO structure
   lstConnection.AddItem GetStatusCode(ECHO.status)
   lstConnection.AddItem ECHO.Address
   lstConnection.AddItem ECHO.RoundTripTime & " ms"
   lstConnection.AddItem ECHO.DataSize & " bytes"
   
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      pos = InStr(ECHO.Data, Chr$(0))
      lstConnection.AddItem Left$(ECHO.Data, pos - 1)
   End If

   lstConnection.AddItem ECHO.DataPointer

End Sub

Private Sub Timer3_Timer()
Refreshlist_Click
End Sub

Private Sub tmrUpdate_Timer()
cmdCheck_Click
tmrUpdate.Enabled = False
End Sub

Private Sub txtIP_Change()
If txtIP.Text = "local" Or txtIP.Text = "locall" Or txtIP.Text = "home" Then
txtIP.Text = wsMain(0).LocalIP
End If
End Sub

Private Sub txtIP_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If txtIP.Text = "local" Or txtIP.Text = "locall" Or txtIP.Text = "home" Then
txtIP.Text = wsMain(0).LocalIP
End If
End Sub

Private Sub uup_Click()
Form3.web.Navigate App.Path & "\ports.htk"
Form3.Show
End Sub

Private Sub ViewCon_Click()

ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add 1, , "File", 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders.Add 2, , "Direction", 1000, lvwColumnCenter
ListView1.ColumnHeaders.Add 3, , "Local Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 4, , "Remote Host", ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders.Add 5, , "Remote Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 6, , "Status", 1300, lvwColumnCenter
ListView1.ColumnHeaders.Add 7, , "File Path", ListView1.Width \ 2 + 1000

RefreshView
End Sub

Private Sub citeste_file()



End Sub
Private Sub wins_ConnectionRequest(Index As Integer, ByVal requestID As Long)
List1.ListIndex = Index - 1
'MsgBox "Someone is trying to connect (ip is under log) Port , " & wins(Index).RemoteHostIP '& List1.Text
If askk = "0" Then
     Exit Sub
Else
    Dim abba As New Form2
    Load abba
    abba.Label1.Caption = wins(Index).RemoteHostIP
    abba.Label4.Caption = Index
    catTop = catTop + abba.Height
    abba.Top = catTop: abba.Left = Screen.Width - abba.Width - 10
    
    abba.Show




  If List1.Text = "7" Then abba.Label7.Caption = "ECHO Port"
  If List1.Text = "19" Then abba.Label7.Caption = "CHARGE Port"
  If List1.Text = "20" Then abba.Label7.Caption = "FTP Port"
  If List1.Text = "21" Then abba.Label7.Caption = "FTP Port"
  If List1.Text = "22" Then abba.Label7.Caption = "SSH Port"
  If List1.Text = "23" Then abba.Label7.Caption = "TELNET Port"
  If List1.Text = "25" Then abba.Label7.Caption = "SMTP Port"
  If List1.Text = "53" Then abba.Label7.Caption = "DNS Port"
  If List1.Text = "79" Then abba.Label7.Caption = "FINGER Port"
  If List1.Text = "80" Then abba.Label7.Caption = "HTTP Port"
If List1.Text = "110" Then abba.Label7.Caption = "POP3"
If List1.Text = "137" Then abba.Label7.Caption = "NBIOS"
If List1.Text = "138" Then abba.Label7.Caption = "NBIOS"
If List1.Text = "139" Then abba.Label7.Caption = "NBIOS"
If List1.Text = "220" Then abba.Label7.Caption = "IMAP"
If List1.Text = "443" Then abba.Label7.Caption = "HTTPS(SSL)"
If List1.Text = "1433" Then abba.Label7.Caption = "SQL Server"
If List1.Text = "1434" Then abba.Label7.Caption = "SQL Monitor"
If List1.Text = "1080" Then abba.Label7.Caption = "Internet Proxy"
If List1.Text = "1863" Then abba.Label7.Caption = "MSN"
If List1.Text = "5050" Then abba.Label7.Caption = "Yahoo Chat"
If List1.Text = "5100" Then abba.Label7.Caption = "Yahoo WebCam"
If List1.Text = "5000" Then abba.Label7.Caption = "yahoo Voice Chat"
If List1.Text = "5001" Then abba.Label7.Caption = "Yahoo Voice Chat"
If List1.Text = "5101" Then abba.Label7.Caption = "Yahoo P2P Messages"
  


End If



Open App.Path & "\ping.htm" For Output As #1
Text6.Text = Text6.Text + "<br><font color='#0456B2'>[" & Date & "] [" & Time & "]</font>" & "<font color='#B2041D'> -- On <b>" & List1.Text & "</b> from: <b><a href='http://www.whois.sc/" & wins(Index).RemoteHostIP & "'>" & wins(Index).RemoteHostIP & "</a></b></font>"
Print #1, Text6.Text
Close #1

End Sub


Function Release()
Dim i

For i = txtFrom.Text To txtTo.Text
wsMain(i).Close
Next
End Function

Private Sub picture2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    picture2.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    picture2.Sorted = True
    End Sub

Private Sub GetWTSProcesses()
   Dim retVal As Long
   Dim Count As Long
   Dim i As Integer
   Dim lpBuffer As Long
   Dim p As Long
   Dim udtProcessInfo As WTS_PROCESS_INFO
   Dim itmAdd As ListItem

   picture2.ListItems.Clear
   retVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
   If retVal Then ' WTSEnumerateProcesses was successful
      p = lpBuffer
        For i = 1 To Count
            ' Count is the number of Structures in the buffer
            ' WTSEnumerateProcesses returns a pointer, so copy it to a
            ' WTS_PROCESS_INO UDT so you can access its members
            CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
            ' Add items to the ListView control
            Set itmAdd = picture2.ListItems.Add(i, , CStr(udtProcessInfo.SessionID))
                itmAdd.SubItems(1) = CStr(udtProcessInfo.ProcessID)
                ' Since pProcessName contains a pointer, call GetStringFromLP to get the
                ' variable length string it points to
                If udtProcessInfo.ProcessID = 0 Then
                    itmAdd.SubItems(2) = "System Idle Process"
                Else
                    itmAdd.SubItems(2) = GetStringFromLP(udtProcessInfo.pProcessName)
                End If
                
                'itmAdd.SubItems(3) = CStr(udtProcessInfo.pUserSid)
                itmAdd.SubItems(3) = GetUserName(udtProcessInfo.pUserSid)

                ' Increment to next WTS_PROCESS_INO structure in the buffer
                p = p + LenB(udtProcessInfo)
        Next i

        Set itmAdd = Nothing
        WTSFreeMemory lpBuffer   'Free your memory buffer
    Else
        ' Error occurred calling WTSEnumerateProcesses
        ' Check Err.LastDllError for error code
        MsgBox "Error occurred calling WTSEnumerateProcesses.  " & "Check the Platform SDK error codes in the MSDN Documentation" & " for more information.", vbCritical, "Error " & Err.LastDllError
    End If
    End Sub

Function GetUserName(sID As Long) As String
    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    GetUserName = Left$(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left$(retname, InStr(retname, vbNullChar) - 1)
    End Function

Private Function GetStringFromLP(ByVal StrPtr As Long) As String
    Dim b As Byte
    Dim tempStr As String
    Dim bufferStr As String
    Dim Done As Boolean

    Done = False
    Do
        ' Get the byte/character that StrPtr is pointing to.
        CopyMemory b, ByVal StrPtr, 1
        If b = 0 Then  ' If you've found a null character, then you're done.
            Done = True
        Else
            tempStr = Chr$(b)  ' Get the character for the byte's value
            bufferStr = bufferStr & tempStr 'Add it to the string
                
            StrPtr = StrPtr + 1  ' Increment the pointer to next byte/char
        End If
    Loop Until Done
    GetStringFromLP = bufferStr
    End Function
