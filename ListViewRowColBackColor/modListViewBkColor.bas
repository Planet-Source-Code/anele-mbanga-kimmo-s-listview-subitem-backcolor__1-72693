Attribute VB_Name = "Module1"
Option Explicit

Public Const GWL_EXSTYLE = -20
Public Const GWL_HINSTANCE = -6
Public Const GWL_HWNDPARENT = -8
Public Const GWL_ID = -12
Public Const GWL_STYLE = -16
Public Const GWL_USERDATA = -21
Public Const GWL_WNDPROC = -4
Public Const DWL_DLGPROC = 4
Public Const DWL_MSGRESULT = 0
Public Const DWL_USER = 8

Public Const NM_CUSTOMDRAW = (-12&)
Public Const WM_NOTIFY As Long = &H4E&
Public Const CDDS_PREPAINT As Long = &H1&
Public Const CDRF_NOTIFYITEMDRAW As Long = &H20&
Public Const CDDS_ITEM As Long = &H10000
Public Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Public Const CDRF_NEWFONT As Long = &H2&
Public Const CDDS_SUBITEM  As Long = &H20000
Public Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20&

Public Type NMHDR
    hWndFrom As Long   ' Window handle of control sending message
    idFrom As Long        ' Identifier of control sending message
    code  As Long          ' Specifies the notification code
End Type

' sub struct of the NMCUSTOMDRAW struct
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' generic customdraw struct
Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type

' listview specific customdraw struct
Public Type NMLVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
    ' if IE >= 4.0 this member of the struct can be used
    iSubItem As Integer
End Type

Public g_addProcOld As Long
Public g_MaxItems As Long
Public g_MaxColumns As Long
Public clr() As Long

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case iMsg
    Case WM_NOTIFY
        Dim udtNMHDR As NMHDR
        CopyMemory udtNMHDR, ByVal lParam, 12&
        
        With udtNMHDR
            If .code = NM_CUSTOMDRAW Then
                Dim udtNMLVCUSTOMDRAW As NMLVCUSTOMDRAW
                CopyMemory udtNMLVCUSTOMDRAW, ByVal lParam, Len(udtNMLVCUSTOMDRAW)
                With udtNMLVCUSTOMDRAW.nmcd
                    Select Case .dwDrawStage
                    Case CDDS_PREPAINT
                        WindowProc = CDRF_NOTIFYITEMDRAW
                        Exit Function
                    Case CDDS_ITEMPREPAINT
                        WindowProc = CDRF_NOTIFYSUBITEMDRAW
                        Exit Function
                    Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
                        If clr(.dwItemSpec, udtNMLVCUSTOMDRAW.iSubItem) <> 0 Then
                            ' a color has been specified, then write row, column
                            udtNMLVCUSTOMDRAW.clrTextBk = clr(.dwItemSpec, udtNMLVCUSTOMDRAW.iSubItem)
                        Else
                            'there is no color, then revert to white background
                            udtNMLVCUSTOMDRAW.clrTextBk = RGB(255, 255, 255)
                        End If
                        CopyMemory ByVal lParam, udtNMLVCUSTOMDRAW, Len(udtNMLVCUSTOMDRAW)
                        WindowProc = CDRF_NEWFONT
                        Exit Function
                    End Select
                End With
            End If
        End With
    End Select
    WindowProc = CallWindowProc(g_addProcOld, hWnd, iMsg, wParam, lParam)
End Function

Public Sub SetLIBackColor(lv As ListView, Row As Integer, Col As Integer, BkColor As Long)
    ' the first column cannot be changed yet
    If Col <= 1 Then Col = 2
    clr(Row - 1, Col - 1) = BkColor
    ' a refresh will repaint the listview thus trapping the events
    lv.Refresh
End Sub

