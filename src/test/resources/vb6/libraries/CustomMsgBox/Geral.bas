Attribute VB_Name = "Geral"
Option Explicit


Global CustomMsgBox As New cCustomMsgBox
Global Resp As Long







Enum soundConstants
   SND_SYNC = &H0 ' play synchronously (default)
   SND_ASYNC = &H1 ' play asynchronously
   SND_NODEFAULT = &H2 ' silence (!default) if sound not found
   SND_MEMORY = &H4 ' pszSound points to a memory file
   SND_LOOP = &H8 ' loop the sound until next sndPlaySound
   SND_NOSTOP = &H10 ' don't stop any currently playing sound
   SND_NOWAIT = &H2000 ' don't wait if the driver is busy
   SND_ALIAS = &H10000 ' name is a registry alias
   SND_ALIAS_ID = &H110000 ' alias is a predefined ID
   SND_FILENAME = &H20000 ' name is file name
   SND_RESOURCE = &H40004 ' name is resource name or atom
   SND_PURGE = &H40 ' purge non-static events for task
   SND_APPLICATION = &H80 ' look for application specific association
End Enum 'soundConstants


' **************************************************************************************************************************************
' * PUBLIC INTERFACE- WIN32 API CONSTANTS
' *
' *

    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_SHOWWINDOW = &H40



' **************************************************************************************************************************************
' * PUBLIC INTERFACE- WIN32 API DECLARATIONS
' *
' *

    'Usado para setar Um Form no topo da exibição
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
        ByVal cy As Long, ByVal wFlags As Long) As Long
        
    'Usado para tocar um audio específico
    Public Declare Function sndPlaySound Lib "winmm.dll" _
    Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long




'**********************************************************************************
'*  DESCRIÇÃO:
'*   Coloca uma Form no no topo da exibição
'*  AUTOR:  Adriano
'*  DATA DE CRIAÇÃO:    03/03/2005
'*----------------------------------------------------------------------------------
'*  MODIFICAÇÃO:
'*  AUTOR:
'*  REVISÃO 0002 (data) :
'***********************************************************************************
Public Sub Ontop(FRM As Form, style As Boolean)

On Local Error GoTo OnTop_Error

    Dim lWinPos As Long

    'Style = True Window Is onTop
    If style Then lWinPos = -1 Else lWinPos = -2
        SetWindowPos FRM.hWnd, lWinPos, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Exit Sub
OnTop_Error:
    Err.Clear
End Sub

'Incluido On local Error
Public Sub PlaySound(ByVal Tipo As Long)
On Local Error GoTo erro
    Select Case Tipo
        Case vbCritical:
            sndPlaySound "SystemHand", SND_ALIAS + SND_ASYNC
        Case vbQuestion:
            sndPlaySound "SystemQuestion", SND_ALIAS + SND_ASYNC
        Case vbExclamation:
            sndPlaySound "SystemExclamation", SND_ALIAS + SND_ASYNC
        Case vbInformation:
            sndPlaySound "SystemAsterisk", SND_ALIAS + SND_ASYNC
        Case Else
            sndPlaySound ".Default", SND_ALIAS + SND_ASYNC
    End Select
sai:
Exit Sub
erro:
Resume sai
End Sub






