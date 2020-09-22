VERSION 5.00
Begin VB.UserControl XPFlatButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   ScaleHeight     =   1140
   ScaleWidth      =   1620
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   -3530
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   480
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XP Flat Button"
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
   Begin VB.Image imgButton 
      Height          =   720
      Left            =   0
      Picture         =   "XPFlatButton.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1290
   End
   Begin VB.Image imgDown 
      Height          =   720
      Left            =   1200
      Picture         =   "XPFlatButton.ctx":3102
      Top             =   2760
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image imgActive 
      Height          =   720
      Left            =   1200
      Picture         =   "XPFlatButton.ctx":6204
      Top             =   2040
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image imgIdle 
      Height          =   720
      Left            =   1200
      Picture         =   "XPFlatButton.ctx":9306
      Top             =   1320
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "XPFlatbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim Cursor As POINTAPI
'Default Property Values:
'Const m_def_Caption = "XP Flat Button"
'Property Variables:
'Dim m_Caption As Single
Dim m_Image As Picture
'Event Declarations:
Event Click() 'MappingInfo=imgButton,imgButton,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."



Private Sub Image1_Click()
RaiseEvent Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgDown.Picture
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
Dim tmpcursor As POINTAPI
GetCursorPos tmpcursor
Cursor.X = tmpcursor.X
Cursor.Y = tmpcursor.Y
Timer1.Enabled = True
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
End Sub

Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgDown.Picture
End Sub


Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
Dim tmpcursor As POINTAPI
GetCursorPos tmpcursor
Cursor.X = tmpcursor.X
Cursor.Y = tmpcursor.Y
Timer1.Enabled = True
End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
End Sub


Private Sub lblCaption_Click()
RaiseEvent Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgDown.Picture
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
Dim tmpcursor As POINTAPI
GetCursorPos tmpcursor
Cursor.X = tmpcursor.X
Cursor.Y = tmpcursor.Y
Timer1.Enabled = True
End Sub


Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton.Picture = imgActive.Picture
End Sub

Private Sub Timer1_Timer()
Dim cursorTMP As POINTAPI
GetCursorPos cursorTMP
If cursorTMP.X = Cursor.X And cursorTMP.Y = Cursor.Y Then

Else
imgButton.Picture = imgIdle
Timer1.Enabled = False
End If

End Sub

Private Sub UserControl_Resize()
imgButton.Width = ScaleWidth
imgButton.Height = ScaleHeight
lblCaption.Width = ScaleWidth
lblCaption.Top = (ScaleHeight - lblCaption.Height) / 2
Image1.Left = (ScaleWidth - Image1.Width) / 2
Image1.Top = (ScaleHeight - Image1.Height) / 2
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=12,0,0,XP Flat Button
'Public Property Get Caption() As Single
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As Single)
'    m_Caption = New_Caption
'    PropertyChanged "Caption"
'End Property




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgButton,imgButton,-1,ToolTipText
Public Property Get Tooltip() As String
Attribute Tooltip.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    Tooltip = imgButton.ToolTipText
End Property

Public Property Let Tooltip(ByVal New_Tooltip As String)
    imgButton.ToolTipText() = New_Tooltip
    PropertyChanged "Tooltip"
End Property

Private Sub imgButton_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgButton,imgButton,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = imgButton.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    imgButton.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Caption = m_def_Caption
    'Set m_Image = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    'Set m_Image = PropBag.ReadProperty("Image", Nothing)
    imgButton.ToolTipText = PropBag.ReadProperty("Tooltip", "")
    imgButton.Enabled = PropBag.ReadProperty("Enabled", True)
    lblCaption.Caption = PropBag.ReadProperty("Caption", lblCaption)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "XP Flat Button")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("Tooltip", imgButton.ToolTipText, "")
    Call PropBag.WriteProperty("Enabled", imgButton.Enabled, True)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, Label1)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "XP Flat Button")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label1,Label1,-1,Caption
'Public Property Get Caption() As String
'    Caption = lblCaption.Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    lblCaption.Caption() = New_Caption
'    PropertyChanged "Caption"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
End Property

