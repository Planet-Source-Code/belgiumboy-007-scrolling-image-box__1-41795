VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ScrollingImage ScrollingImage1 
      Height          =   4575
      Left            =   105
      TabIndex        =   2
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8070
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   285
      Left            =   6945
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    ScrollingImage1.SetImage Text1.Text
End Sub

Private Sub Form_Load()
    Text1.Text = App.Path & "\59.jpg"
End Sub
