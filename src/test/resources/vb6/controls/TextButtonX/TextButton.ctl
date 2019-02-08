VERSION 5.00
Begin VB.UserControl TextButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   DefaultCancel   =   -1  'True
   ForeColor       =   &H00000000&
   ScaleHeight     =   3600
   ScaleWidth      =   9165
   ToolboxBitmap   =   "TextButton.ctx":0000
   Begin VB.PictureBox picMiddle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   135
      Picture         =   "TextButton.ctx":0312
      ScaleHeight     =   420
      ScaleWidth      =   3405
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   3405
      Begin VB.Image imgFig 
         Height          =   255
         Left            =   570
         Top             =   75
         Width           =   255
      End
      Begin VB.Label lbCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1455
         TabIndex        =   9
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      Picture         =   "TextButton.ctx":6B44
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3540
      Picture         =   "TextButton.ctx":C199
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picRightDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3570
      Picture         =   "TextButton.ctx":117C1
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picMiddleDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   150
      Picture         =   "TextButton.ctx":16DDB
      ScaleHeight     =   420
      ScaleWidth      =   3405
      TabIndex        =   4
      Top             =   1500
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.PictureBox picLeftDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      Picture         =   "TextButton.ctx":1D409
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picRightUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3570
      Picture         =   "TextButton.ctx":22A41
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   1065
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picMiddleUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   150
      Picture         =   "TextButton.ctx":28069
      ScaleHeight     =   420
      ScaleWidth      =   3405
      TabIndex        =   1
      Top             =   1065
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.PictureBox picLeftUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      Picture         =   "TextButton.ctx":2E89B
      ScaleHeight     =   420
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   1065
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "TextButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim b_MoveDown As Boolean
'Default Property Values:
Const m_def_Enabled = True

'Property Variables:
Dim m_Enabled As Boolean
Dim m_Picture As Picture

Dim RightButtonPressed As Boolean

'Event Declarations:
Event Click() 'MappingInfo=lbCaption,lbCaption,-1,Click
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lbCaption,lbCaption,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lbCaption,lbCaption,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub imgFig_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub imgFig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub imgFig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub lbCaption_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub lbCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub lbCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub picLeft_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub picLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub picLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub picRight_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub picRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub picMiddle_DblClick()
   Call UserControl_DblClick
End Sub

Private Sub picMiddle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub picMiddle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If m_Enabled = False Then
        Exit Sub
    End If
    
    If RightButtonPressed = True Then
        Exit Sub
    End If
    
    Call MoveDown
End Sub
'
'Rotina para pintar a borda de foco
Private Sub DrawFocus()
Dim y As Long
Dim x As Long
Dim x2 As Long
Dim y2 As Long

    For x = 1 To (picMiddle.ScaleWidth / Screen.TwipsPerPixelX) - 2 Step 2
        x2 = x * Screen.TwipsPerPixelX
        picMiddle.PSet (x2, 65), QBColor(0)
        picMiddle.PSet (x2, picMiddle.ScaleHeight - 65), QBColor(0)
    Next x
    
    For y = 6 To (picMiddle.ScaleHeight / Screen.TwipsPerPixelY) - 6 Step 2
        y2 = y * Screen.TwipsPerPixelY
        picMiddle.PSet (0, y2), QBColor(0)
        picMiddle.PSet (picMiddle.Width - 20, y2), QBColor(0)
    Next y
End Sub

Private Sub UserControl_EnterFocus()
    DrawFocus
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13) Then
    MoveUp
    UserControl_EnterFocus
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If m_Enabled = False Then
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        RightButtonPressed = True
        Exit Sub
    Else
        RightButtonPressed = False
    End If
    
    Call MoveDown
    RaiseEvent MouseDown(Button, Shift, x, y)
    UserControl_EnterFocus
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled = False Then
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        RightButtonPressed = True
        Exit Sub
    Else
        RightButtonPressed = False
    End If
    
    Call MoveUp
    RaiseEvent MouseUp(Button, Shift, x, y)
    RaiseEvent Click
    UserControl_EnterFocus
End Sub

Private Sub UserControl_ExitFocus()
    picMiddle.Cls
    Call MoveUp
End Sub

Private Sub UserControl_Initialize()
    lbCaption.ForeColor = RGB(50, 50, 50)
End Sub

Private Sub MoveDown()
    If b_MoveDown = False Then
        b_MoveDown = True
        picLeft.Picture = picLeftDown.Picture
        picMiddle.Picture = picMiddleDown.Picture
        picRight.Picture = picRightDown.Picture
        lbCaption.Top = lbCaption.Top + 15
        imgFig.Top = imgFig.Top + 15
        lbCaption.ForeColor = RGB(40, 40, 40)
        SetCapture (UserControl.hWnd)
    End If
End Sub

Private Sub MoveUp()
    If b_MoveDown = True Then
        b_MoveDown = False
        picLeft.Picture = picLeftUp.Picture
        picMiddle.Picture = picMiddleUp.Picture
        picRight.Picture = picRightUp.Picture
        lbCaption.Top = lbCaption.Top - 15
        imgFig.Top = imgFig.Top - 15
        lbCaption.ForeColor = RGB(50, 50, 50)
        Call ReleaseCapture
    End If
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    'UserControl.Enabled = New_Enabled
    lbCaption.ForeColor = IIf(New_Enabled = True, RGB(50, 50, 50), RGB(95, 95, 95))
    PropertyChanged "Enabled"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    
    lbCaption.Caption = PropBag.ReadProperty("Caption", "Texto")
    lbCaption.Left = (picMiddle.Width - TextWidth(lbCaption.Caption)) / 2
    lbCaption.ForeColor = IIf(m_Enabled = True, RGB(50, 50, 50), RGB(95, 95, 95))
    
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", Caption, "Texto")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Private Sub UserControl_Resize()

On Local Error GoTo ErrLine

    'Fixa largura do botão = largura da picture
    UserControl.Height = picMiddle.Height
    
    'Limite inferior
    If UserControl.Width < 500 Then
        UserControl.Width = 500
    End If
    
    'Limite superior
    If UserControl.Width > 6945 Then
        UserControl.Width = 6945
    End If
    
    picMiddle.Width = UserControl.Width - picRight.Width - picLeft.Width
    picRight.Left = picMiddle.Width + picLeft.Width
    lbCaption.Left = (picMiddle.Width - TextWidth(lbCaption.Caption)) / 2
    imgFig.Left = Abs((picMiddle.Width - imgFig.Width)) / 2
    imgFig.Top = Abs(picMiddle.Height - imgFig.Height) / 2
    
    Exit Sub
ErrLine:
    Err.Clear
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lbCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lbCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    lbCaption.Left = (picMiddle.Width - TextWidth(lbCaption.Caption)) / 2
End Property

Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set imgFig.Picture = New_Picture
    If (Not (New_Picture Is Nothing)) Then
        imgFig.Left = Abs((picMiddle.Width / 2 - imgFig.Width / 2))
        imgFig.Top = Abs((picMiddle.Height - imgFig.Height) / 2)
    End If

    PropertyChanged "Picture"
End Property














