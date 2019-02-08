Attribute VB_Name = "Module1"
Public Declare Function SHFileOperation Lib _
"shell32.dll" Alias "SHFileOperationA" _
(lpFileOp As Any) As Long

Public Declare Sub SHFreeNameMappings Lib _
"shell32.dll" (ByVal hNameMappings As Long)

Public Declare Sub CopyMemory Lib "KERNEL32" _
Alias "RtlMoveMemory" (hpvDest As Any, hpvSource _
As Any, ByVal cbCopy As Long)

Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As FO_Functions
pFrom As String
pTo As String
fFlags As FOF_Flags
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String 'only used if FOF_SIMPLEPROGRESS
End Type

Public Enum FO_Functions
FO_MOVE = &H1
FO_COPY = &H2
FO_DELETE = &H3
FO_RENAME = &H4
End Enum

Public Enum FOF_Flags
FOF_MULTIDESTFILES = &H1
FOF_CONFIRMMOUSE = &H2
FOF_SILENT = &H4
FOF_RENAMEONCOLLISION = &H8
FOF_NOCONFIRMATION = &H10
FOF_WANTMAPPINGHANDLE = &H20
FOF_ALLOWUNDO = &H40
FOF_FILESONLY = &H80
FOF_SIMPLEPROGRESS = &H100
FOF_NOCONFIRMMKDIR = &H200
FOF_NOERRORUI = &H400
FOF_NOCOPYSECURITYATTRIBS = &H800
FOF_NORECURSION = &H1000
FOF_NO_CONNECTED_ELEMENTS = &H2000
FOF_WANTNUKEWARNING = &H4000
End Enum

Public Type SHNAMEMAPPING
pszOldPath As String
pszNewPath As String
cchOldPath As Long
cchNewPath As Long
End Type

Public Function SHFileOP(ByRef lpFileOp As SHFILEOPSTRUCT) As Long
    Dim result As Long
    Dim lenFileop As Long
    Dim foBuf() As Byte
    
    lenFileop = LenB(lpFileOp)
    ReDim foBuf(1 To lenFileop) 'the size of the structure.
    
    Call CopyMemory(foBuf(1), lpFileOp, lenFileop)
    
    Call CopyMemory(foBuf(19), foBuf(21), 12)
    result = SHFileOperation(foBuf(1))
    
    SHFileOP = result
End Function

