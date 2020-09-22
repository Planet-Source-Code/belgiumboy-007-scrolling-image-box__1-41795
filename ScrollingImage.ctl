VERSION 5.00
Begin VB.UserControl ScrollingImage 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   ScaleHeight     =   4305
   ScaleWidth      =   5520
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   3840
      Width           =   255
   End
   Begin VB.VScrollBar vs 
      Height          =   3855
      LargeChange     =   20
      Left            =   5160
      SmallChange     =   5
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hs 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      SmallChange     =   5
      TabIndex        =   0
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   240
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "ScrollingImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub hs_Change()
    Image1.Left = -(hs.Value * 4)
End Sub

Private Sub UserControl_Initialize()
    vs.Enabled = False
    hs.Enabled = False
    Picture1.ZOrder vbBringToFront
    Image1.ZOrder vbSendToBack
End Sub

Private Sub UserControl_Resize()
    vs.Left = UserControl.Width - vs.Width
    hs.Top = UserControl.Height - hs.Height
    vs.Height = hs.Top
    hs.Width = vs.Left
    Image1.Height = hs.Top
    Image1.Width = vs.Left
    Picture1.Left = vs.Left
    Picture1.Top = hs.Top
End Sub

Public Sub SetImage(ByVal Path As String)
    Set Image1.Picture = LoadPicture(Path)
    
    Image1.Top = 0
    Image1.Left = 0
    
    If Image.Height <= UserControl.Height Then
        vs.Enabled = False
    Else
        vs.Enabled = True
        vs.Min = 0
        vs.Max = (Image1.Height - vs.Height) / 4
        vs.Value = 0
    End If
    
    If Image.Width <= UserControl.Width Then
        hs.Enabled = False
    Else
        hs.Enabled = True
        hs.Min = 0
        hs.Max = (Image1.Width - hs.Width) / 4
        hs.Value = 0
    End If
End Sub

Private Sub vs_Change()
    Image1.Top = -(vs.Value * 4)
End Sub
