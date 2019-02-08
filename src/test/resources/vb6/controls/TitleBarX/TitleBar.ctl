VERSION 5.00
Begin VB.UserControl TitleBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ScaleHeight     =   3600
   ScaleWidth      =   19200
   ToolboxBitmap   =   "TitleBar.ctx":0000
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6945
      ScaleHeight     =   420
      ScaleWidth      =   2550
      TabIndex        =   20
      Top             =   0
      Width           =   2550
   End
   Begin VB.PictureBox picB_Dir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6270
      Picture         =   "TitleBar.ctx":0312
      ScaleHeight     =   210
      ScaleWidth      =   105
      TabIndex        =   17
      Top             =   90
      Width           =   105
   End
   Begin VB.PictureBox picClipControlsX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9885
      Picture         =   "TitleBar.ctx":169A
      ScaleHeight     =   285
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   60
      Width           =   915
      Begin VB.PictureBox picCloseButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   615
         Picture         =   "TitleBar.ctx":7004
         ScaleHeight     =   285
         ScaleWidth      =   300
         TabIndex        =   7
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picMaxButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   315
         Picture         =   "TitleBar.ctx":C869
         ScaleHeight     =   285
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picMinButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15
         Picture         =   "TitleBar.ctx":12099
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox picFundo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1830
      Picture         =   "TitleBar.ctx":17879
      ScaleHeight     =   420
      ScaleWidth      =   975
      TabIndex        =   19
      Top             =   570
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   780
      Picture         =   "TitleBar.ctx":18E2B
      ScaleHeight     =   210
      ScaleWidth      =   5490
      TabIndex        =   16
      Top             =   90
      Width           =   5490
   End
   Begin VB.PictureBox picB_Esq 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   570
      Picture         =   "TitleBar.ctx":1FFB5
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   15
      Top             =   90
      Width           =   210
   End
   Begin VB.PictureBox picButtonX2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   45
      Picture         =   "TitleBar.ctx":2136E
      ScaleHeight     =   420
      ScaleWidth      =   390
      TabIndex        =   18
      Top             =   0
      Width           =   390
   End
   Begin VB.PictureBox picLogoOld 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6960
      ScaleHeight     =   420
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   1485
      Width           =   1635
   End
   Begin VB.PictureBox picMinBtDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      Picture         =   "TitleBar.ctx":26DA3
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   13
      Top             =   1860
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMaxBtDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   390
      Picture         =   "TitleBar.ctx":2C55B
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   12
      Top             =   1860
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picCloseBtDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      Picture         =   "TitleBar.ctx":31D35
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picCloseBtUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   735
      Picture         =   "TitleBar.ctx":3754F
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMaxBtUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   390
      Picture         =   "TitleBar.ctx":3CDB4
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picMinBtUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      Picture         =   "TitleBar.ctx":425E4
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picButtonXUp2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   930
      Picture         =   "TitleBar.ctx":47DC4
      ScaleHeight     =   420
      ScaleWidth      =   390
      TabIndex        =   3
      Top             =   585
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picButtonXDown2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1335
      Picture         =   "TitleBar.ctx":4D7F9
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   585
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picButtonXDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   435
      Picture         =   "TitleBar.ctx":53246
      ScaleHeight     =   420
      ScaleWidth      =   390
      TabIndex        =   1
      Top             =   585
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picButtonXUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   15
      Picture         =   "TitleBar.ctx":58CA7
      ScaleHeight     =   420
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   585
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long 'hide or show de form

'Event Declarations:
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MinButtonClick() 'MappingInfo=picMinButton,picMinButton,-1,Click
Event MaxButtonClick() 'MappingInfo=picMaxButton,picMaxButton,-1,Click
Event CloseButtonClick() 'MappingInfo=picCloseButton,picCloseButton,-1,Click
Event MinButtonMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picMinButton,picMinButton,-1,MouseUp
Event MaxButtonMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picMaxButton,picMaxButton,-1,MouseUp
Event CloseButtonMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picCloseButton,picCloseButton,-1,MouseUp
Event CloseButtonMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picCloseButton,picCloseButton,-1,MouseDown
Event MaxButtonMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picMaxButton,picMaxButton,-1,MouseDown
Event MinButtonMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picMinButton,picMinButton,-1,MouseDown


Dim PtX As Single
Dim PtY As Single
'Default Property Values:
Const m_def_ControlBox = 0
Const m_def_FormMovable = 0
Const m_def_DisableFunctionMinButton = 0
Const m_def_DisableFunctionMaxButton = 0
Const m_def_DisableFunctionCloseButton = 0
Const m_def_Logo = True
Const m_def_TitleName = "Coloque a seguite instrução no Form_Resize(): TitleBar1.Width = TitleBar1.Parent.Width"
Const m_def_MenuName = ""
'Property Variables:
Dim m_ControlBox As Boolean
Dim m_FormMovable As Boolean
Dim m_DisableFunctionMinButton As Boolean
Dim m_DisableFunctionMaxButton As Boolean
Dim m_DisableFunctionCloseButton As Boolean
Dim m_Logo As Boolean
Dim m_TitleName As String
Dim m_MenuName As String

Dim m_MinButton As Boolean
Dim m_MaxButton As Boolean
Dim m_CloseButton As Boolean


Dim FormHandle As Long

Public Property Get Enabled() As Boolean
    Enabled = picClipControlsX.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    picClipControlsX.Enabled = New_Enabled
End Property

'Simula double click na barra de título do windows
Private Sub picB_Dir_DblClick()
    Call MaxButton_Click
End Sub

Private Sub picB_Dir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Simula double click na barra de título do windows
Private Sub picB_Esq_DblClick()
    Call MaxButton_Click
End Sub

Private Sub picB_Esq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picMinButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    picMinButton.Picture = picMinBtUp.Picture
    RaiseEvent MinButtonMouseUp(Button, Shift, X, Y)
    RaiseEvent MinButtonClick

    If m_DisableFunctionMinButton Then
        Exit Sub
    End If

    UserControl.Parent.WindowState = vbMinimized
End Sub

Private Sub picMaxButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    picMaxButton.Picture = picMaxBtUp.Picture
    RaiseEvent MaxButtonMouseUp(Button, Shift, X, Y)
    RaiseEvent MaxButtonClick

    'Necessário testar m_ControlBox e m_MaxButton devido a simulação desta operação por um Double_Click
    'na barra de titulo
    If m_DisableFunctionMaxButton = True Or m_ControlBox = False Or m_MaxButton = False Then
        Exit Sub
    End If

    If UserControl.Parent.WindowState = vbMaximized Then
        UserControl.Parent.WindowState = vbNormal
    Else
        UserControl.Parent.WindowState = vbMaximized
    End If
End Sub

Private Sub picCloseButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    picCloseButton.Picture = picCloseBtUp.Picture
    RaiseEvent CloseButtonMouseUp(Button, Shift, X, Y)
    RaiseEvent CloseButtonClick

    If m_DisableFunctionCloseButton Then
        Exit Sub
    End If

    Unload UserControl.Parent
End Sub

Private Sub picCloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picCloseButton.Enabled = False Or picClipControlsX.Enabled = False Then
        Exit Sub
    End If
    
    picCloseButton.Picture = picCloseBtDown.Picture
    RaiseEvent CloseButtonMouseDown(Button, Shift, X, Y)
End Sub

Private Sub picMaxButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picMaxButton.Enabled = False Or picClipControlsX.Enabled = False Then
        Exit Sub
    End If
    
    picMaxButton.Picture = picMaxBtDown.Picture
    RaiseEvent MaxButtonMouseDown(Button, Shift, X, Y)
End Sub

Private Sub picMinButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picMinButton.Enabled = False Or picClipControlsX.Enabled = False Then
        Exit Sub
    End If
    
    picMinButton.Picture = picMinBtDown.Picture
    RaiseEvent MinButtonMouseDown(Button, Shift, X, Y)
End Sub

Private Sub MaxButton_Click()
On Local Error Resume Next
    RaiseEvent MaxButtonClick

    'Necessário testar m_ControlBox e m_MaxButton devido a simulação desta operação por um Double_Click
    'na barra de titulo
    If m_DisableFunctionMaxButton = True Or m_ControlBox = False Or m_MaxButton = False Then
        Exit Sub
    End If

    If UserControl.Parent.WindowState = vbMaximized Then
        UserControl.Parent.WindowState = vbNormal
    Else
        UserControl.Parent.WindowState = vbMaximized
    End If
End Sub

Public Property Get MinButton() As Boolean
    MinButton = m_MinButton
End Property

Public Property Let MinButton(ByVal New_MinButton As Boolean)
    m_MinButton = New_MinButton
    picMinButton.Visible = m_MinButton
    PropertyChanged "MinButton"
End Property

Public Property Get MaxButton() As Boolean
    MaxButton = m_MaxButton
End Property

Public Property Let MaxButton(ByVal New_MaxButton As Boolean)
    m_MaxButton = New_MaxButton
    picMaxButton.Visible = m_MaxButton
    PropertyChanged "MaxButton"
End Property

Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    m_CloseButton = New_CloseButton
    picCloseButton.Visible = m_CloseButton
    PropertyChanged "CloseButton"
End Property

Private Sub ControlResize()
On Local Error Resume Next
    
    Dim aux As String

    If UserControl.Parent.WindowState = 1 Then
        Exit Sub
    End If
    
    'Acerta tamanho da picture de fundo
    'picFundo.Width = UserControl.Width
    
    'Posiciona botões
    picClipControlsX.Left = UserControl.Parent.Width - picClipControlsX.Width - 150
    
    'Necessário pois ao redimencionar a tela em runtime as figuras estão sendo apagadas
    picClipControlsX.Refresh
    picMaxButton.Refresh
    picMinButton.Refresh
    picCloseButton.Refresh
    
    'Posiciona Logo
    picLogo.Left = UserControl.Parent.Width / 2 - picLogo.Width / 2
    
    aux = m_TitleName
    If picName.TextWidth(aux) <= 13300 Then 'picName.width > 13400
        If m_ControlBox = True Then
            'Veririca se barra com nome esta batendo no clipcontrols e redimensiona barra de nome para cortar
            'informação
            If picName.Left + picName.TextWidth(aux) + 100 + picB_Dir.Width > picClipControlsX.Left Then
                picName.Width = picClipControlsX.Left - (picName.Left + 120)
            Else
                picName.Width = picName.TextWidth(aux) + 100
            End If
        Else
            'Veririca se barra com nome esta batendo no final do form e redimensiona barra de nome para cortar
            'informação
            If picName.Left + picName.TextWidth(aux) + 100 + picB_Dir.Width > UserControl.Width Then
                picName.Width = UserControl.Width - (picName.Left + 300)
            Else
                picName.Width = picName.TextWidth(aux) + 100
            End If
        End If
        picB_Dir.Left = picName.Left + picName.Width
    End If
    
    picName.Cls
    picName.Print aux
    
    Call AtualizaFundo
    
'    'Verifica se barra com nome não batendo nos botões e
'    'redimenciona ela
'    If picB_Dir.Left + picB_Dir.Width > picClipControlsX.Left Then
'        picClipControlsX.Left = picB_Dir.Left + picB_Dir.Width
'        UserControl.Parent.Width = picClipControlsX.Left + picClipControlsX.Width + 150
'    End If
End Sub



'Atualiza a figura de fundo do controle
Private Sub AtualizaFundo()

'Não funciona neste controle... Não sei porque... Deve ser loop infinito de repaint
'    Dim X As Long, Y As Long
'
'    X = 0
'    Y = 0
'    Do
'        X = X + picFundo.Width
'        picTitle.PaintPicture picFundo.Picture, X, Y
'    Loop Until X >= picTitle.Width


 Dim X As Long, Y As Long
    
    X = 0
    Y = 0
    While X < UserControl.Width
        UserControl.PaintPicture picFundo.Picture, X, Y
        X = X + picFundo.Width
    Wend



'    Dim largura As Double
'    Dim n As Double
'    Dim i As Long
'    Dim n_controls As Long
'
'    largura = picTitle.Width - picFundo(0).Width * picFundo.Count
'    'Verifica se deve acrescentar figuras no fundo do controle para compor o desenho
'    If (largura > 0) Then
'        n = largura / picFundo(0).Width
'        If (n < 1) Then 'Se for fracao menor que 1 arredenda para 1 bloco a ser acrescentado
'            n = 1
'        Else
'            If (n / Int(n)) > 1 Then 'Se der fracao maior que 1 pega a parte intereira e soma +1
'                n = Int(n) + 1
'            End If
'        End If
'    End If
'
'    If n > 0 Then
'        For i = 1 To n
'            n_controls = picFundo.Count
'            Load picFundo(n_controls)
'            Set picFundo(n_controls).Container = picTitle
'            picFundo(n_controls).ZOrder 1
'            picFundo(n_controls).Top = picFundo(0).Top
'            picFundo(n_controls).Left = picFundo(n_controls - 1).Left + picFundo(0).Width
'            picFundo(n_controls).Visible = True
'        Next i
'    End If
End Sub

Private Sub picB_Dir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PtY = Y
        PtX = X
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picB_Dir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    If m_FormMovable = False Then
        Exit Sub
    End If
    
    If Button = vbLeftButton Then
        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
        Call ControlResize
    End If
End Sub

Private Sub picB_Esq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PtY = Y
        PtX = X
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picB_Esq_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    If m_FormMovable = False Then
        Exit Sub
    End If
    
    If Button = vbLeftButton Then
        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
        Call ControlResize
    End If
End Sub

'Simula double click na barra de título do windows
Private Sub picName_DblClick()
    Call MaxButton_Click
End Sub

Private Sub picName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PtY = Y
        PtX = X
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    If m_FormMovable = False Then
        Exit Sub
    End If
    
    If Button = vbLeftButton Then
        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
        Call ControlResize
    End If
End Sub

Private Sub picName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picName_Paint()
    Call ControlResize
End Sub

'Simula double click na barra de título do windows
Private Sub picLogo_DblClick()
    Call MaxButton_Click
End Sub

Private Sub picLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PtY = Y
        PtX = X
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    If m_FormMovable = False Then
        Exit Sub
    End If
    
    If Button = vbLeftButton Then
        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
        Call ControlResize
    End If
End Sub

Private Sub picLogo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Private Sub picFundo_Click(Index As Integer)
'    Call MaxButton_Click
'End Sub
'
'Private Sub picFundo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbLeftButton Then
'        PtY = Y
'        PtX = X
'    End If
'
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub picFundo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
'    If m_FormMovable = False Then
'        Exit Sub
'    End If
'
'    If Button = vbLeftButton Then
'        If UserControl.Parent.WindowState = 2 Then
'            Exit Sub
'        End If
'        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
'        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
'        Call ControlResize
'    End If
'End Sub
'
'Private Sub picFundo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub

'Simula double click na barra de título do windows
Private Sub UserControl_DblClick()
    Call MaxButton_Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PtY = Y
        PtX = X
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    If m_FormMovable = False Then
        Exit Sub
    End If

    If Button = vbLeftButton Then
        UserControl.Parent.Top = UserControl.Parent.Top + (Y - PtY)
        UserControl.Parent.Left = UserControl.Parent.Left + (X - PtX)
        Call ControlResize
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub picButtonX2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picButtonX2.Picture = picButtonXDown2.Picture
End Sub

Private Sub picButtonX2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
    Dim i As Integer
    Dim aux As Integer
    
    picButtonX2.Picture = picButtonXUp2.Picture
    
    aux = IIf(UserControl.Parent.ScaleMode = 1, 1, 15)

    For i = 0 To UserControl.ParentControls.Count - 1
        If TypeOf UserControl.ParentControls(i) Is Menu Then
            If UserControl.ParentControls(i).Name = m_MenuName Then
                UserControl.Parent.PopupMenu UserControl.ParentControls(i), , picButtonX2.Left \ aux, (picButtonX2.Top + picButtonX2.Height) \ aux
                Exit Sub
            End If
        End If
    Next i
End Sub

Public Property Get TitleName() As String
    TitleName = m_TitleName
End Property

Public Property Let TitleName(ByVal New_TitleName As String)
    m_TitleName = New_TitleName
    PropertyChanged "TitleName"
    Call ControlResize
End Property

Public Property Get MenuName() As String
    MenuName = m_MenuName
End Property

Public Property Let MenuName(ByVal New_MenuName As String)
    m_MenuName = New_MenuName
    PropertyChanged "MenuName"
End Property

Private Sub UserControl_Initialize()
     Call AtualizaFundo
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Local Error Resume Next
    m_TitleName = m_def_TitleName
    m_MenuName = m_def_MenuName
    m_MinButton = True
    m_MaxButton = True
    m_CloseButton = True
    m_Logo = m_def_Logo
    m_DisableFunctionMinButton = m_def_DisableFunctionMinButton
    m_DisableFunctionMaxButton = m_def_DisableFunctionMaxButton
    m_DisableFunctionCloseButton = m_def_DisableFunctionCloseButton
    m_FormMovable = m_def_FormMovable
    m_ControlBox = m_def_ControlBox
    
    Call AtualizaFundo
End Sub

Private Sub UserControl_Paint()
On Local Error Resume Next
    Call ControlResize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Local Error Resume Next
    m_TitleName = PropBag.ReadProperty("TitleName", m_def_TitleName)
    m_MenuName = PropBag.ReadProperty("MenuName", m_def_MenuName)
    m_MinButton = PropBag.ReadProperty("MinButton", True)
    m_MaxButton = PropBag.ReadProperty("MaxButton", True)
    m_CloseButton = PropBag.ReadProperty("CloseButton", True)
    m_ControlBox = PropBag.ReadProperty("ControlBox", m_def_ControlBox)
    m_Logo = PropBag.ReadProperty("Logo", m_def_Logo)
    
    picMinButton.Visible = m_MinButton
    picMaxButton.Visible = m_MaxButton
    picCloseButton.Visible = m_CloseButton
    picClipControlsX.Visible = m_ControlBox
    picLogo.Visible = m_Logo
    
    picButtonX2.Left = 45
    picB_Esq.Top = 90
    picB_Esq.Left = 570
    picB_Dir.Top = 90
    picName.Top = 90
    picClipControlsX.Top = 60
    
    If m_MenuName = "" Then
        If picB_Esq.Left = 570 Then
            picB_Esq.Left = picB_Esq.Left - 500
            picName.Left = picName.Left - 500
            picB_Dir.Left = picB_Dir.Left - 500
        End If
        picButtonX2.Visible = False
    Else
        If picB_Esq.Left = 70 Then
            picB_Esq.Left = picB_Esq.Left + 500
            picName.Left = picName.Left + 500
            picB_Dir.Left = picB_Dir.Left + 500
        End If
        picButtonX2.Visible = True
        
    End If
    m_DisableFunctionMinButton = PropBag.ReadProperty("DisableFunctionMinButton", m_def_DisableFunctionMinButton)
    m_DisableFunctionMaxButton = PropBag.ReadProperty("DisableFunctionMaxButton", m_def_DisableFunctionMaxButton)
    m_DisableFunctionCloseButton = PropBag.ReadProperty("DisableFunctionCloseButton", m_def_DisableFunctionCloseButton)
    m_FormMovable = PropBag.ReadProperty("FormMovable", m_def_FormMovable)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Local Error Resume Next
    Call PropBag.WriteProperty("TitleName", m_TitleName, m_def_TitleName)
    Call PropBag.WriteProperty("MenuName", m_MenuName, m_def_MenuName)
    Call PropBag.WriteProperty("MinButton", m_MinButton, True)
    Call PropBag.WriteProperty("MaxButton", m_MaxButton, True)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, True)
    Call PropBag.WriteProperty("Logo", m_Logo, m_def_Logo)
    Call PropBag.WriteProperty("DisableFunctionMinButton", m_DisableFunctionMinButton, m_def_DisableFunctionMinButton)
    Call PropBag.WriteProperty("DisableFunctionMaxButton", m_DisableFunctionMaxButton, m_def_DisableFunctionMaxButton)
    Call PropBag.WriteProperty("DisableFunctionCloseButton", m_DisableFunctionCloseButton, m_def_DisableFunctionCloseButton)
    Call PropBag.WriteProperty("FormMovable", m_FormMovable, m_def_FormMovable)
    Call PropBag.WriteProperty("ControlBox", m_ControlBox, m_def_ControlBox)
End Sub

Private Sub UserControl_Resize()
On Local Error Resume Next
    Call ControlResize
End Sub

Private Sub UserControl_Show()
On Local Error Resume Next
    
    Dim i As Integer
    
    'UserControl.Parent.Caption = ""
    
    For i = 0 To UserControl.ParentControls.Count - 1
        If TypeOf UserControl.ParentControls(i) Is TitleBar Then
            UserControl.ParentControls.Item(i).Top = -1
            UserControl.ParentControls.Item(i).Left = -1
        End If
    Next i
    
    picName.ForeColor = RGB(240, 240, 240)
    UserControl.Parent.BackColor = RGB(77, 77, 77)
    UserControl.Width = UserControl.Parent.Width
    UserControl.Height = picFundo.Height
    picClipControlsX.Left = UserControl.Parent.Width - picClipControlsX.Width - 150
    picLogo.Left = UserControl.Parent.Width / 2 - picLogo.Width / 2
    FormHandle = UserControl.Parent.hwnd
    
    Call AtualizaFundo
End Sub

Public Property Get Logo() As Boolean
    Logo = m_Logo
End Property

Public Property Let Logo(ByVal New_Logo As Boolean)
    m_Logo = New_Logo
    picLogo.Visible = m_Logo
    PropertyChanged "Logo"
End Property

Public Property Get DisableFunctionMinButton() As Boolean
    DisableFunctionMinButton = m_DisableFunctionMinButton
End Property

Public Property Let DisableFunctionMinButton(ByVal New_DisableFunctionMinButton As Boolean)
    m_DisableFunctionMinButton = New_DisableFunctionMinButton
    PropertyChanged "DisableFunctionMinButton"
End Property

Public Property Get DisableFunctionMaxButton() As Boolean
    DisableFunctionMaxButton = m_DisableFunctionMaxButton
End Property

Public Property Let DisableFunctionMaxButton(ByVal New_DisableFunctionMaxButton As Boolean)
    m_DisableFunctionMaxButton = New_DisableFunctionMaxButton
    PropertyChanged "DisableFunctionMaxButton"
End Property

Public Property Get DisableFunctionCloseButton() As Boolean
    DisableFunctionCloseButton = m_DisableFunctionCloseButton
End Property

Public Property Let DisableFunctionCloseButton(ByVal New_DisableFunctionCloseButton As Boolean)
    m_DisableFunctionCloseButton = New_DisableFunctionCloseButton
    PropertyChanged "DisableFunctionCloseButton"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FormMovable() As Boolean
    FormMovable = m_FormMovable
End Property

Public Property Let FormMovable(ByVal New_FormMovable As Boolean)
    m_FormMovable = New_FormMovable
    PropertyChanged "FormMovable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ControlBox() As Boolean
    ControlBox = m_ControlBox
End Property

Public Property Let ControlBox(ByVal New_ControlBox As Boolean)
    m_ControlBox = New_ControlBox
    picClipControlsX.Visible = m_ControlBox
    PropertyChanged "ControlBox"
End Property



