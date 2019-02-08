VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
'Subistitui a MsgBox padrão.<br>
'Vantagens:<br>
'   Estar totalmente compatível com a anterior MsgBox;<br>
'   Esta em Portuques;<br>
'   Esta no padrão de interface;<br>
'   Não fica por baixo de janelas "always on the top";<br>





Dim test As VbMsgBoxResult


'Constant             Value       Description
'vbOKOnly               0           Display OK button only.
'vbOKCancel             1           Display OK and Cancel buttons.
'vbAbortRetryIgnore     2           Display Abort, Retry, and Ignore buttons.
'vbYesNoCancel          3           Display Yes, No, and Cancel buttons.
'vbYesNo                4           Display Yes and No buttons.
'vbRetryCancel          5           Display Retry and Cancel buttons.
'vbCritical             16          Display Critical Message icon.
'vbQuestion             32          Display Warning Query icon.
'vbExclamation          48          Display Warning Message icon.
'vbInformation          64          Display Information Message icon.
'vbDefaultButton1       0           First button is default.
'vbDefaultButton2       256         Second button is default.
'vbDefaultButton3       512         Third button is default.
'vbDefaultButton4       768         Fourth button is default.


'Return Values

'Constant     Value     Description
'vbOK           1           OK
'vbCancel       2           Cancel
'vbAbort        3           Abort
'vbRetry        4           Retry
'vbIgnore       5           Ignore
'vbYes          6           Yes
'vbNo           7           No






''
'Chamada para criação da mensagem<br>
' @param Prompt Define Texto livre usado para descrever a mensagem.<br>
'        Para pular linhas usar vbNewLine ou vbCrLf
' @param Buttons Define tipo de botões que irão ser visualizados, vide MsgBox
' @Param Title Define Título que será apresentado na janela.
' @return Retorna valores padão. Vide valores de retorno da MsgBox.

'Incluido On local Error
Public Function MsgBox(Optional ByVal Prompt As String = "", Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "CustomMessageBox") As VbMsgBoxResult
    Dim TipoBotoes As Integer
    Dim TipoMensagem As Integer
    Dim DefaultButton As Integer
    Dim i As Integer
    Dim MsgForm As frmMsgBox
    
On Local Error GoTo erro

    Set MsgForm = New frmMsgBox
    
    With MsgForm
    
        Resp = 0
    
        'Verifica que tipo de butões irão ser mostrados
        TipoBotoes = 7 And Buttons 'Macara com 111 para pegar apenas as informações referentes a que tipo de botões serão usados
        
        If TipoBotoes > 5 Then
            TipoBotoes = vbOKOnly
        End If
        
        Select Case TipoBotoes
            Case vbOKOnly:
                .btResp(0).Caption = "Ok"
                .btResp(0).Tag = 1
            Case vbOKCancel:
                Load .btResp(.btResp.Count)
                .btResp(0).Caption = "Ok"
                .btResp(0).Tag = 1
                .btResp(1).Caption = "Cancelar"
                .btResp(1).Tag = 2
            Case vbAbortRetryIgnore:
                Load .btResp(.btResp.Count)
                Load .btResp(.btResp.Count)
                .btResp(0).Caption = "Abortar"
                .btResp(0).Tag = 3
                .btResp(1).Caption = "Repetir"
                .btResp(1).Tag = 4
                .btResp(2).Caption = "Ignorar"
                .btResp(2).Tag = 5
            Case vbYesNoCancel:
                Load .btResp(.btResp.Count)
                Load .btResp(.btResp.Count)
                .btResp(0).Caption = "Sim"
                .btResp(0).Tag = 6
                .btResp(1).Caption = "Não"
                .btResp(1).Tag = 7
                .btResp(2).Caption = "Cancelar"
                .btResp(2).Tag = 2
            Case vbYesNo:
                Load .btResp(.btResp.Count)
                .btResp(0).Caption = "Sim"
                .btResp(0).Tag = 6
                .btResp(1).Caption = "Não"
                .btResp(1).Tag = 7
            Case vbRetryCancel:
                Load .btResp(.btResp.Count)
                .btResp(0).Caption = "Repetir"
                .btResp(0).Tag = 4
                .btResp(1).Caption = "Cancelar"
                .btResp(1).Tag = 2
            Case Else
                .btResp(0).Caption = "Ok" 'Segurança
                .btResp(0).Tag = 1
        End Select
        
        'Verifica qual o tipo de mensagem
        TipoMensagem = Buttons And 120   'Mascara com 1110000 para pegar apenas o tipo de mensagem associada
        
        Select Case TipoMensagem
            Case vbCritical:
                .picMessage.Picture = .imgList.ListImages(2).Picture
                .picMessage.Tag = 1
                .Tag = vbCritical
            Case vbQuestion:
                .picMessage.Picture = .imgList.ListImages(4).Picture
                .picMessage.Tag = 1
                .Tag = vbQuestion
            Case vbExclamation:
                .picMessage.Picture = .imgList.ListImages(1).Picture
                .picMessage.Tag = 1
                .Tag = vbExclamation
            Case vbInformation:
                .picMessage.Picture = .imgList.ListImages(3).Picture
                .picMessage.Tag = 1
                .Tag = vbInformation
            Case Else
                .picMessage.Tag = 0
                .picMessage.Visible = False
        End Select
        
        DefaultButton = Buttons And 768 'Mascara com 1100000000 para pegar apenas qual botão será o padrão
        
        Select Case DefaultButton
            Case vbDefaultButton1:
                .btResp(0).Default = True
            Case vbDefaultButton2:
                If .btResp.Count >= 1 Then
                    .btResp(1).Default = True
                Else
                    .btResp(0).Default = True
                End If
            Case vbDefaultButton3:
                If .btResp.Count >= 2 Then
                    .btResp(2).Default = True
                Else
                    .btResp(0).Default = True
                End If
            Case vbDefaultButton4:
                '.btResp(3).Default = True   'Não sei qual é o quarto botão pois so instacio no máximo 3!
            Case Else
                .btResp(0).Default = True
        End Select
        
        'Insere Titulo
        .tlbTitle.TitleName = Title
        
        Dim aux As Variant
        Dim Pos As Integer, X As Integer, Xl As Integer
        
        'Reposiciona controle caso não tenha figura ilustrativa
        If .picMessage.Tag = 1 Then
            .lbPrompt(0).Left = .picMessage.Left + .picMessage.Width + 300
        Else
            .lbPrompt(0).Left = 300
        End If
        
        'Troca eventual CrLf por NewLine para termos um parametro seguro para fazer o split
        Prompt = Replace(Prompt, vbCrLf, vbNewLine)
        
        aux = Split(Prompt, vbNewLine)
        
        For i = 0 To UBound(aux)
            
            'Não tenta instanciar o controle que já existe
            If i > 0 Then
                Load .lbPrompt(.lbPrompt.Count)
            End If
            
            .lbPrompt(i).Caption = aux(i)
            .lbPrompt(i).Visible = True
            If i > 0 Then
                .lbPrompt(i).Top = .TextHeight(.lbPrompt(i).Caption) + .lbPrompt(i - 1).Top
                If .TextWidth(.lbPrompt(i).Caption) > X Then
                    X = .TextWidth(.lbPrompt(i).Caption)
                    Xl = .lbPrompt(i).Left
                End If
            Else
                X = .TextWidth(.lbPrompt(i).Caption)
                Xl = .lbPrompt(i).Left
            End If
        Next i
        
        'Redimensiona janela
        If Xl + X + 300 > 4200 Then
            .Width = Xl + X + 300
            .picCano.Width = .Width
        End If
        
        If .lbPrompt(.lbPrompt.Count - 1).Top + .lbPrompt(.lbPrompt.Count - 1).Height > .picCano.Top Then
            .picCano.Top = .lbPrompt(.lbPrompt.Count - 1).Top + .lbPrompt(.lbPrompt.Count - 1).Height + 200
            .Height = .picCano.Top + .picCano.Height + 100
        End If
        
        'Reposiciona botões
        .btResp(0).Left = ((.Width - (.btResp(0).Width * .btResp.Count)) / 2)
        For i = 1 To .btResp.Count - 1
            .btResp(i).Left = .btResp(i - 1).Left + .btResp(i - 1).Width - 15
            .btResp(i).Visible = True
        Next i

        Call Ontop(MsgForm, True)

        MsgForm.Show vbModal
            
        If Resp <> 0 Then
            MsgBox = Resp
        End If
        
        Set MsgForm = Nothing
    End With
sai:
Exit Function

erro:
Resume sai
End Function

