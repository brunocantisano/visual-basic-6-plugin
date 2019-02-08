Attribute VB_Name = "Geral"
Option Explicit

Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Private Declare Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function GetMenu Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, _
ByVal bRevert As Long) As Long

Private Declare Function SetMenuInfo Lib "user32" _
(ByVal hmenu As Long, _
mi As MENUINFO) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
(ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" _
(ByVal OLE_COLOR As Long, _
ByVal HPALETTE As Long, _
pccolorref As Long) As Long





Public Sub SetMenuColour(ByVal hwndfrm As Long)
    'set application menu colour
    Dim mi As MENUINFO
    Dim flags As Long
    Dim clrref As Long

    clrref = RGB(36, 36, 36)
    flags = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
    
    With mi
    .cbSize = Len(mi)
    .fMask = flags
    .hbrBack = CreateSolidBrush(clrref)
    End With
    
    SetMenuInfo GetMenu(hwndfrm), mi
    DrawMenuBar hwndfrm
End Sub

Public Sub SetSysMenuColour(ByVal hwndfrm As Long)
    
    'set system menu colour
    Dim mi As MENUINFO
    Dim hSysMenu As Long
    Dim clrref As Long
    
    'convert a Windows colour (OLE colour)
    'to a valid RGB colour if required
    clrref = clrref = RGB(36, 36, 36)
    
    'get handle to the system menu,
    'fill in struct, assign to menu,
    'and force a redraw with the
    'new attributes
    hSysMenu = GetSystemMenu(hwndfrm, False)
    
    With mi
    .cbSize = Len(mi)
    .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
    .hbrBack = CreateSolidBrush(clrref)
    End With
    
    SetMenuInfo hSysMenu, mi
    DrawMenuBar hSysMenu

End Sub











