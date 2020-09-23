VERSION 5.00
Begin VB.UserControl BevelMe 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BevelMe.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "BevelMe.ctx":0CA4
End
Attribute VB_Name = "BevelMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'  Very simple user control concept:
'  Place the control on any form
'  Set the control's FormColor property
'  Set the BevelWidth property (0 to 5)
'  When the Form is run, it is made borderless
'  you can size it and drag it
'  AND - a nice bevel is added to it for you
'
'  any problems, email me
'  afontana@bigpond.net.au
'
'**********************************************




Option Explicit

Private WithEvents objParent As Form
Attribute objParent.VB_VarHelpID = -1
Dim SizeDir As String
'Default Property Values:
Const m_def_FormColor = &HC0C0C0
Const m_def_BackColor = &HC0C0C0
Const m_def_BevelWidth = 4
Const m_def_GrabHandleWidth = 105
'Property Variables:
Dim m_FormColor As OLE_COLOR
Dim m_BackColor As Long
Dim m_BevelWidth As Integer
Dim m_GrabHandleWidth As Integer





Public Sub Init()
On Error GoTo ErrH
Set objParent = UserControl.Parent


Exit Sub
ErrH:

End Sub
Private Sub UserControl_InitProperties()

m_GrabHandleWidth = m_def_GrabHandleWidth
m_FormColor = m_def_FormColor
m_BevelWidth = m_def_BevelWidth
m_FormColor = m_def_FormColor
UserControl.BackColor = m_FormColor

UserControl.Parent.BorderStyle = 0
UserControl.Parent.ControlBox = False
UserControl.Parent.Appearance = 0
UserControl.Parent.MinButton = False
UserControl.Parent.MaxButton = False
UserControl.Parent.BackColor = m_FormColor
UserControl.Parent.AutoRedraw = True



End Sub


Private Sub objParent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   If SizeDir = "" Then
      DragOBJ objParent
   Else
      SizeOBJ objParent, SizeDir
   End If
End If

End Sub


Private Sub objParent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GrabWid As Integer
GrabWid = m_GrabHandleWidth

SizeDir = ""
   Select Case X
   Case 0 To GrabWid
      Select Case Y
         Case 0 To GrabWid
            objParent.MousePointer = 8
            SizeDir = "TL"
         Case objParent.Height - GrabWid To objParent.Height
            objParent.MousePointer = 6
            SizeDir = "BL"
            
      Case Else
         objParent.MousePointer = 9
         SizeDir = "L"
      
      End Select
      
   Case objParent.Width - GrabWid To objParent.Width
      Select Case Y
      Case 0 To GrabWid
         objParent.MousePointer = 6
         SizeDir = "TR"
      Case objParent.Height - GrabWid To objParent.Height
         objParent.MousePointer = 8
         SizeDir = "BR"
      Case Else
         objParent.MousePointer = 9
         SizeDir = "R"
      End Select
   
   Case Else
      
      Select Case Y
      Case 0 To GrabWid
         Select Case X
            Case 0 To GrabWid
               objParent.MousePointer = 8
               SizeDir = "TL"

            Case objParent.Width - GrabWid To objParent.Width
               objParent.MousePointer = 6
               SizeDir = "TR"

         Case Else
            objParent.MousePointer = 7
            SizeDir = "T"

         End Select
      
      Case objParent.Height - GrabWid To objParent.Height
         Select Case X
            Case 0 To GrabWid
               objParent.MousePointer = 6
               SizeDir = "BL"

            Case objParent.Width - GrabWid To objParent.Width
               objParent.MousePointer = 8
               SizeDir = "BR"

         Case Else
            objParent.MousePointer = 7
            SizeDir = "B"

         End Select
      Case Else
         objParent.MousePointer = 0
         SizeDir = ""

      End Select
   
   End Select

End Sub


Private Sub objParent_Resize()

Call BevelMaster(objParent, m_FormColor, m_BevelWidth)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,105
Public Property Get GrabHandleWidth() As Integer
   GrabHandleWidth = m_GrabHandleWidth
End Property

Public Property Let GrabHandleWidth(ByVal New_GrabHandleWidth As Integer)
   m_GrabHandleWidth = New_GrabHandleWidth
   PropertyChanged "GrabHandleWidth"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Init
   m_GrabHandleWidth = PropBag.ReadProperty("GrabHandleWidth", m_def_GrabHandleWidth)
   m_BevelWidth = PropBag.ReadProperty("BevelWidth", m_def_BevelWidth)

   m_FormColor = PropBag.ReadProperty("FormColor", m_def_FormColor)
      UserControl.BackColor = m_FormColor

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 31 * 15
UserControl.Height = 33 * 15

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)


   Call PropBag.WriteProperty("GrabHandleWidth", m_GrabHandleWidth, m_def_GrabHandleWidth)
   Call PropBag.WriteProperty("BevelWidth", m_BevelWidth, m_def_BevelWidth)
   

   Call PropBag.WriteProperty("FormColor", m_FormColor, m_def_FormColor)
   UserControl.BackColor = m_FormColor
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,&Hc0c0c0

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,4
Public Property Get BevelWidth() As Integer
   BevelWidth = m_BevelWidth
   If BevelWidth < 0 Then
      BevelWidth = 0
   ElseIf BevelWidth > 5 Then
      BevelWidth = 5
   End If
   
End Property

Public Property Let BevelWidth(ByVal New_BevelWidth As Integer)

   m_BevelWidth = New_BevelWidth
   PropertyChanged "BevelWidth"
   If m_BevelWidth < 0 Then
      m_BevelWidth = 0
   ElseIf m_BevelWidth > 5 Then
      m_BevelWidth = 5
   End If

   Call DoTheBevel

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&Hc0c0c0
Public Property Get FormColor() As OLE_COLOR
   FormColor = m_FormColor
End Property

Public Property Let FormColor(ByVal New_FormColor As OLE_COLOR)
   m_FormColor = New_FormColor
   PropertyChanged "FormColor"
   Call DoTheBevel
   UserControl.BackColor = m_FormColor
End Property


Sub DoTheBevel()
On Error GoTo ErrH
Call BevelMaster(objParent, m_FormColor, m_BevelWidth)
Exit Sub
ErrH:


End Sub
