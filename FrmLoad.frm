VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmLoad 
   BorderStyle     =   0  'None
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "FrmLoad.frx":0000
   ScaleHeight     =   2280
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Progbar 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © SOSCO, 2005"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pleas wait!        Loading..."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Greenlight As Boolean
Private Sub Form_Load()





Me.Visible = True
Progbar.Max = 10
Progbar.Value = 0
Greenlight = False

Progbar.Value = 5
loadit
Do Until Greenlight = True
DoEvents
Loop

Progbar.Value = 10

Form1.Visible = True
Unload Me
Dim kkt
On Error GoTo a
            Open App.Path & "\ports.stf" For Input As #1
            Do
                Line Input #1, kkt
                If kkt = "#end" Then
                    Close #1
                    Exit Sub
                Else
                    Form1.List1.AddItem kkt
                End If
            Loop
            Close #1
a: MsgBox "'ports.stf' can not found in program's directory!", vbCritical, "Opened Doors"
End Sub
Public Sub loadit()
DoEvents
Load Form1
Greenlight = True
End Sub
