VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Used Ports"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6765
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.BoxList LstFont 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9128
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LstFont.Config Me, "Arial", 7, 45, RGB(229, 229, 229), vbWhite, vbBlack

'config columns
LstFont.Title "Port", 5
LstFont.Title "Description", 50
LstFont.Activate True
'---------------------

LstFont.Add " 999": LstFont.Add " Test"

LstFont.BoxNew  'new line
LstFont.Selected LstFont.ListCount - 1

LstFont.Add "1": LstFont.Add " tcpmux"

LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1


LstFont.Add "5": LstFont.Add " rje": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "7": LstFont.Add " echo": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "9": LstFont.Add " discard": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "11": LstFont.Add " systat": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "13": LstFont.Add " daytime": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "15": LstFont.Add " netstat": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "17": LstFont.Add " qotd": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "18": LstFont.Add " send/rwp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "19": LstFont.Add " chargen": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "20": LstFont.Add " ftp-data": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "21": LstFont.Add " ftp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "22": LstFont.Add " ssh, pcAnywhere": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "23": LstFont.Add " Telnet": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "25": LstFont.Add " SMTP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "27": LstFont.Add " ETRN": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "29": LstFont.Add " msg-icp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "31": LstFont.Add " msg-auth": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "33": LstFont.Add " dsp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "37": LstFont.Add " time": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "38": LstFont.Add " RAP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "39": LstFont.Add " rlp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "42": LstFont.Add " nameserv, WINS": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "43": LstFont.Add " whois, nickname": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "49": LstFont.Add " TACACS, Login Host Protocol": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "50": LstFont.Add " RMCP, re-mail-ck": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "53": LstFont.Add " DNS": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "57": LstFont.Add " MTP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "59": LstFont.Add " NFILE": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "63": LstFont.Add " whois++": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "66": LstFont.Add " sql*net": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "67": LstFont.Add " bootps": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "68": LstFont.Add " bootpd/dhcp": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "69": LstFont.Add " Trivial File Transfer Protocol (tftp)": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "70": LstFont.Add " Gopher": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "79": LstFont.Add " finger": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "80": LstFont.Add " www-http": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "88": LstFont.Add " Kerberos, WWW": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "95": LstFont.Add " supdup": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "96": LstFont.Add " DIXIE": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "139": LstFont.Add " NetBIOS": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "143": LstFont.Add " IMAP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "210": LstFont.Add " Z39.50": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "218": LstFont.Add " MPP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "220": LstFont.Add " IMAP3": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "259": LstFont.Add " ESRO": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "264": LstFont.Add " FW1_topo": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "311": LstFont.Add " Apple WebAdmin": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "521": LstFont.Add " RIPng": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "522": LstFont.Add " ULS": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "531": LstFont.Add " IRC": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "543": LstFont.Add " KLogin, AppleShare over IP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "545": LstFont.Add " QuickTime": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "548": LstFont.Add " AFP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "554": LstFont.Add " Real Time Streaming Protocol": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "555": LstFont.Add " phAse Zero": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "563": LstFont.Add " NNTP over SSL": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "575": LstFont.Add " VEMMI": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "581": LstFont.Add " Bundle Discovery Protocol": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "593": LstFont.Add " MS-RPC": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "608": LstFont.Add " SIFT/UFT": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "626": LstFont.Add " Apple ASIA": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "631": LstFont.Add " IPP (Internet Printing Protocol)": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "635": LstFont.Add " mountd": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "636": LstFont.Add " sldap": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "642": LstFont.Add " EMSD": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "648": LstFont.Add " RRP (NSI Registry Registrar Protocol)": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "655": LstFont.Add " tinc": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "660": LstFont.Add " Apple MacOS Server Admin": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "666": LstFont.Add " Doom": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "674": LstFont.Add " ACAP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "687": LstFont.Add " AppleShare IP Registry": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "700": LstFont.Add " buddyphone": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "705": LstFont.Add " AgentX for SNMP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "901": LstFont.Add " swat, realsecure": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "993": LstFont.Add " s-imap": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "995": LstFont.Add " s-pop": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1723": LstFont.Add " PPTP control port": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1755": LstFont.Add " Windows Media .asf": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1758": LstFont.Add " TFTP multicast": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1812": LstFont.Add " RADIUS server": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1813": LstFont.Add " RADIUS accounting": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1818": LstFont.Add " ETFTP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1973": LstFont.Add " DLSw DCAP/DRAP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1985": LstFont.Add " HSRP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "1999": LstFont.Add " Cisco AUTH": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2001": LstFont.Add " glimpse": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2049": LstFont.Add " NFS": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2064": LstFont.Add " distributed.net": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2065": LstFont.Add " DLSw": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2066": LstFont.Add " DLSw": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2106": LstFont.Add " MZAP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2140": LstFont.Add " DeepThroat": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2301": LstFont.Add " Compaq Insight Management Web Agents": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "2327": LstFont.Add " Netscape Conference": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "3305": LstFont.Add " ODETTE": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "3306": LstFont.Add " mySQL": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "3389": LstFont.Add " RDP Protocol (Terminal Server)": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "3521": LstFont.Add " netrek": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "4000": LstFont.Add " icq, command-n-conquer": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "4321": LstFont.Add " rwhois": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "4333": LstFont.Add " mSQL": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "4827": LstFont.Add " HTCP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "5010": LstFont.Add " Yahoo! Messenger": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "5060": LstFont.Add " SIP": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "5190": LstFont.Add " AIM": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "6667": LstFont.Add " IRC": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "6776": LstFont.Add " Sub7": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "7007": LstFont.Add " MSBD, Windows Media encoder": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "30029": LstFont.Add " AOL Admin": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1
LstFont.Add "31337": LstFont.Add " Back Orifice   ": LstFont.BoxNew: LstFont.Selected LstFont.ListCount - 1














End Sub
