VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form3 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Show Log"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6435
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      ExtentX         =   11245
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
web.Width = Me.Width - 100
web.Height = Me.Height - 350
End Sub
