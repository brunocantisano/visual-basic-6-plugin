VERSION 5.00
Object = "{81F858A6-BFC7-46DB-800E-FC4BA3108A1E}#1.0#0"; "TextButtonX.ocx"
Object = "{A18AD222-B4B8-41B1-949D-B22E707EF1A6}#1.0#0"; "TitleBarX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMsgBox 
   BackColor       =   &H004D4D4D&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgList 
      Left            =   3300
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":080F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":1331
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":1DC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCano 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      Picture         =   "frmMsgBox.frx":287A
      ScaleHeight     =   420
      ScaleWidth      =   6135
      TabIndex        =   1
      Top             =   1485
      Width           =   6135
      Begin TextButtonX.TextButton btResp 
         Height          =   420
         Index           =   0
         Left            =   1478
         TabIndex        =   2
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   741
      End
   End
   Begin TitleBarX.TitleBar tlbTitle 
      Height          =   420
      Left            =   -1
      TabIndex        =   0
      Top             =   -1
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   741
      TitleName       =   ""
      Logo            =   0   'False
      FormMovable     =   -1  'True
   End
   Begin VB.Image picMessage 
      Height          =   615
      Left            =   165
      Top             =   585
      Width           =   675
   End
   Begin VB.Label lbPrompt 
      AutoSize        =   -1  'True
      BackColor       =   &H004D4D4D&
      Caption         =   "Prompt"
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Index           =   0
      Left            =   1785
      TabIndex        =   3
      Top             =   780
      Width           =   495
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Incluido On local Error
Private Sub btResp_Click(Index As Integer)
On Local Error Resume Next
    Resp = CLng(btResp(Index).Tag)
    Unload Me
End Sub

'Incluido On local Error
Private Sub Form_Activate()
On Local Error GoTo erro
    Static aux As Boolean
    
    If Not aux Then
        aux = True
        Select Case Me.Tag
            Case vbCritical:
                Call PlaySound(vbCritical)
            Case vbQuestion:
                Call PlaySound(vbQuestion)
            Case vbExclamation:
                Call PlaySound(vbExclamation)
            Case vbInformation:
                Call PlaySound(vbInformation)
            Case Else
                Call PlaySound(0)
        End Select
    End If
sai:
Exit Sub
erro:
Resume sai
End Sub

'Incluido On local Error
Private Sub Form_Resize()
On Local Error Resume Next
    tlbTitle.Width = tlbTitle.Parent.Width
End Sub




