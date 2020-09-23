VERSION 5.00
Object = "*\AtpnBevelMe.vbp"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   4035
   ClientTop       =   1785
   ClientWidth     =   6330
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin tpnBevelMe.BevelMe BevelMe1 
      Left            =   1620
      Top             =   3465
      _ExtentX        =   820
      _ExtentY        =   873
      BevelWidth      =   5
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4140
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   375
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   930
      TabIndex        =   0
      Top             =   735
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
BevelMe1.formcolor = &HFF&

End Sub

Private Sub Form_Load()
'UserControl11.init

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   BevelMe1.bevelwidth = Val(Text1)
End If

End Sub
