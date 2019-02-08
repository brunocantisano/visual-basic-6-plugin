VERSION 5.00
Object = "{24509B9F-DD6F-451F-91E1-A5FBD88B90F2}#1.0#0"; "TitleBarX.ocx"
Object = "{57BA5BD9-8862-49FC-9CB5-457926B9D10E}#1.0#0"; "TextButtonX.ocx"
Begin VB.Form Form1 
   BackColor       =   &H004D4D4D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   -60
   ClientTop       =   -120
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TitleBarX.TitleBar TitleBar1 
      Height          =   420
      Left            =   -1
      TabIndex        =   5
      Top             =   -1
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      TitleName       =   ""
      Logo            =   0   'False
      ControlBox      =   -1  'True
   End
   Begin TextButtonX.TextButton TextButton1 
      Height          =   420
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   741
      Caption         =   "Copia"
   End
   Begin VB.TextBox Txtdestino 
      Height          =   345
      Left            =   1140
      TabIndex        =   3
      Top             =   750
      Width           =   2070
   End
   Begin VB.TextBox Txtorigem 
      Height          =   345
      Left            =   1140
      TabIndex        =   2
      Top             =   390
      Width           =   2070
   End
   Begin VB.Label Label2 
      BackColor       =   &H004D4D4D&
      Caption         =   "destino"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H004D4D4D&
      Caption         =   "origem"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   435
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lret As Long
Dim fileop As SHFILEOPSTRUCT

Private Sub Form_Load()
    TitleBar1.TitleName = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Sub

Private Sub TextButton1_Click()

    With fileop
      .hwnd = 0
      
      .wFunc = FO_COPY
      
      .pFrom = Txtorigem & vbNullChar & vbNullChar
      
      .pTo = Txtdestino.Text & vbNullChar & vbNullChar
      
      .lpszProgressTitle = "Aguarde, realizando copia..."
      
      .fFlags = FOF_SIMPLEPROGRESS Or FOF_RENAMEONCOLLISION
    
    End With
    
    lret = SHFileOP(fileop)
    
    If result <> 0 Then 'a operaçao falhou
       'MessageBox.CustomMessage Err.LastDllError   'exibe o erro retornado pela API
       MsgBox Err.LastDllError    'exibe o erro retornado pela API
    Else
      If fileop.fAnyOperationsAborted <> 0 Then
         'MessageBox.CustomMessage "Operação falhou !!!"
         MsgBox "Operação falhou !!!"
      Else
        MsgBox "Cópia efetuada com sucesso!!!"
        'MessageBox.CustomMessage "Cópia efetuada com sucesso!!!"
      End If
    End If
End Sub
