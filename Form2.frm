VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1200
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "unknown"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Close connection for:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "555.555.555.555"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   260
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "yes"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "no"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
Form1.wins(Label4.Caption).Close
Me.Hide
Unload Me
End Sub

Private Sub Label3_Click()
Me.Hide
Unload Me

End Sub

