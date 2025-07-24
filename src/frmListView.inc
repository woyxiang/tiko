'    tiko editor - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
    

''
''  frmListView.inc
''   

#include once "frmListView.bi"
#include once "modCompileErrors.bi"


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmListView_PositionWindows( byval hLV as HWND ) as LRESULT
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(hLV)
    if pWindow = 0 then exit function

    ' Position the Header and the ListBox
    dim as HWND hHeader = GetDlgItem( hLV, IDC_FRMLISTVIEW_HEADER )
    dim as HWND hListBox = GetDlgItem( hLV, IDC_FRMLISTVIEW_LISTBOX )
    
    dim as RECT rc = AfxGetWindowClientRect( hLV )
    dim as long nClientWidth = rc.right - rc.left
    dim as long nClientHeight = rc.bottom - rc.top
    dim as long nHeaderHeight = AfxGetWindowHeight(hHeader)
    
    SetWindowPos( hHeader, 0, _
        0, 0, nClientWidth, nHeaderHeight, SWP_NOZORDER or SWP_SHOWWINDOW )
    
    SetWindowPos( hListBox, 0, _
        0, nHeaderHeight, nClientWidth, nClientHeight - nHeaderHeight, SWP_NOZORDER or SWP_SHOWWINDOW )
    
    AfxRedrawWindow(hHeader)
    AfxRedrawWindow(hListBox)

    function = 0
end function


' ========================================================================================
' frmListViewHeader_SubclassProc 
' ========================================================================================
function frmListViewHeader_SubclassProc ( _
            byval hWin   as HWND, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
            ) as LRESULT

    select case uMsg

        case WM_MOUSEMOVE
            ' Track that we are over the control in order to catch the 
            ' eventual WM_MOUSEHOVER and WM_MOUSELEAVE events
            dim tme as TrackMouseEvent
            tme.cbSize = sizeof(TrackMouseEvent)
            tme.dwFlags = TME_HOVER or TME_LEAVE
            tme.hwndTrack = hWin
            tme.dwHoverTime = 1    ' HOVER_DEFAULT is 500
            TrackMouseEvent(@tme) 
            
        case WM_MOUSEHOVER
            
        case WM_MOUSELEAVE
            
        case WM_RBUTTONDOWN
            
        case WM_LBUTTONUP
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
            ' if the specified point is in the client area of the list box, or one if it is outside the 
            ' client area.
            if hiword(idx) <> 1 then
            end if
                    
        case WM_ERASEBKGND
            return true
        
        case WM_PAINT
            dim pWindow as CWindow ptr = AfxCWindowPtr(GetParent(hWin))

            dim as clsDoubleBuffer b
            
            b.BeginDoubleBuffer(hWin)

            dim as LISTVIEW_DATA_TYPE ptr pData = frmListView_GetPointer( GetParent(hWin) )
            if pData then 
                ' Set default fore and back colors
                b.SetForeColors( pData->ForeColorHot, pData->ForeColorHot )
                b.SetBackColors( pData->BackColor, pData->BackColorHot )

                ' Clear the entire back client area
                b.PaintClientRect()
              
                b.SetFont( GUIFONT_9 )

                dim as RECT rc = b.rcClient
                for i as long = lbound(pData->ColData) to ubound(pData->ColData)
                    rc.right = rc.left + pData->ColData(i).nWidth
                    b.PaintRect( @rc )
                    b.PaintText( pData->ColData(i).wszText, @rc, DT_LEFT )
                    rc.left = rc.right
                next
            end if
            
            b.EndDoubleBuffer()

            return 0

        case WM_NCDESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hWin, @frmListViewHeader_SubclassProc, uIdSubclass )
    end select
    
    ' For messages that we don't deal with
    function = DefSubclassProc( hWin, uMsg, _wParam, _lParam )

end function


' ========================================================================================
' frmListViewListBox_SubclassProc 
' ========================================================================================
function frmListViewListBox_SubclassProc ( _
            byval hWin   as HWND, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
            ) as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(GetParent(hWin))
    static as long accumDelta

    ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
    dim as boolean bEnterPressed = false
    if (uMsg = WM_KEYUP) andalso (loword(_wParam) = VK_RETURN) then 
        uMsg = WM_LBUTTONDBLCLK
        bEnterPressed = true
    end if
    
    
    select case uMsg
        
        case WM_MOUSEWHEEL
            ' accumulate delta until scroll one line (up +120, down -120). 
            ' 120 is the Microsoft default delta
            dim as long zDelta = GET_WHEEL_DELTA_WPARAM( _wParam )
            dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
            accumDelta = accumDelta + zDelta
            if accumDelta >= 120 then       ' scroll up 3 lines
                nTopIndex = nTopIndex - 3
                nTopIndex = max( 0, nTopIndex )
                SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
                accumDelta = 0
                frmListView_PositionWindows(hWin)
            elseif accumDelta <= -120 then  ' scroll down 3 lines
                nTopIndex = nTopIndex + 3
                SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
                accumDelta = 0
                frmListView_PositionWindows(hWin)
            end if

'        Case WM_MOUSEMOVE
'            ' Track that we are over the control in order to catch the 
'            ' eventual WM_MOUSEHOVER and WM_MOUSELEAVE events
'            dim tme as TrackMouseEvent
'            tme.cbSize = sizeof(TrackMouseEvent)
'            tme.dwFlags = TME_HOVER or TME_LEAVE
'            tme.hwndTrack = hWin
'            tme.dwHoverTime = 1    ' HOVER_DEFAULT is 500
'            TrackMouseEvent(@tme) 
            
'        case WM_MOUSEHOVER
            
'        case WM_MOUSELEAVE
            
'        case WM_RBUTTONDOWN
            
        case WM_LBUTTONUP
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
            ' if the specified point is in the client area of the list box, or one if it is outside the 
            ' client area.
            if hiword(idx) <> 1 then
                ListBox_SetCurSel( hWin, idx )
            end if
                    
        case WM_LBUTTONDBLCLK 
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            if (hiword(idx) <> 1) orelse (bEnterPressed = true) then
                SetDocumentErrorPosition( GetParent(hWin), gCompile.CompileID )    
            end if
            exit function
        
        case WM_KEYUP
            select case loword(_wParam)
                case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
            end select
            exit function

        case WM_CHAR   ' prevent the annoying beep!
            if _wParam = VK_RETURN then return 0
            if _wParam = VK_ESCAPE then return 0
                            
        case WM_ERASEBKGND
            ' Get item height and client size
            dim as RECT rcClient: GetClientRect(hWin, @rcClient)
            dim as long itemHeight = SendMessage(hWin, LB_GETITEMHEIGHT, 0, 0)
            dim as long itemCount  = SendMessage(hWin, LB_GETCOUNT, 0, 0)
            dim as long paintedHeight = itemCount * itemHeight

            ' Fill remaining area
            if paintedHeight < rcClient.bottom then
                dim as RECT rcFill
                SetRect( @rcFill, 0, paintedHeight, rcClient.right, rcClient.bottom )
                dim as HBRUSH hbr = CreateSolidBrush( ghOutput.BackColor )
                dim as HDC hdc = cast(HDC, _wParam)
                FillRect( hdc, @rcFill, hbr )
            end if
            
            return true

        case WM_NCDESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hWin, @frmListViewListBox_SubclassProc, uIdSubclass )
    end select
    
    ' For messages that we don't deal with
    function = DefSubclassProc( hWin, uMsg, _wParam, _lParam )

end function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnMeasureItem( _
            byval hwnd as HWND, _
            byval lpmis as MEASUREITEMSTRUCT ptr _
            ) as long

    ' Set the height of the list box items. 
    dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
    lpmis->itemHeight = pWindow->ScaleY(EXPLORERITEM_HEIGHT_COMPACT)

    function = 0
end function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnDrawItem( _
            byval hwnd as HWND, _
            byval lpdis as const DRAWITEMSTRUCT ptr _
            ) as long

    dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
    if pWindow = 0 then exit function
    if lpdis = 0 then exit function
    
    if ( lpdis->itemAction = ODA_DRAWENTIRE ) orelse _
        ( lpdis->itemAction = ODA_SELECT ) orelse _
        ( lpdis->itemAction = ODA_FOCUS ) then

        dim as clsDoubleBuffer b

        b.BeginDoubleBuffer(lpdis->hwndItem, lpdis->hDC, lpdis->rcItem)

        dim as boolean IsHot = false
        if ListBox_GetCurSel(lpdis->hwndItem) = lpdis->itemID then IsHot = true

        dim as LISTVIEW_DATA_TYPE ptr pData = frmListView_GetPointer(hwnd)

        if pData then 
            dim as COLORREF backclr = iif( IsHot, pData->BackColorHot, pData->BackColor)
            dim as COLORREF foreclr = iif( IsHot, pData->ForeColorHot, pData->ForeColorHot)
            b.SetForeColors(foreclr, foreclr)
            b.SetBackColors(backclr, backclr)

            ' Create our rect that works with the entire line
            dim as RECT rc
            SetRect( @rc, 0, 0, b.rcClientWidth, b.rcClientHeight )
            dim as RECT rcText = rc 
            
            ' Paint the entire background
            b.PaintRect( @rc )

            b.SetFont( GUIFONT_9 )

            dim as CWSTR wszCaption = AfxGetListBoxText(lpdis->hwndItem, lpdis->ItemID)

            dim as long nCount = AfxStrParseCount(wszCaption, chr(9))
            for i as long = 0 to nCount - 1
                rcText.right = rcText.left + pData->ColData(i).nWidth
                ' offset the left margin ever so slightly
                rcText.left += pWindow->ScaleX(4)
                dim as CWSTR wszText = AfxStrParse(wszCaption, i + 1, chr(9)) 
                b.PaintText( wszText, @rcText, DT_LEFT )
                rcText.left = rcText.right
            next

        end if
        
        b.EndDoubleBuffer()

    end if   

    function = true
    
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnCommand( _
            byval hwnd as HWND, _
            byval id as long, _
            byval hwndCtl as HWND, _
            byval codeNotify as UINT _
            ) as LRESULT

    select case codeNotify
        case LBN_SELCHANGE
            ' update the highlighting of the current line
            AfxRedrawWindow(hwndCtl)
            ' update the scrollbar position if necessary
            frmListView_PositionWindows(hwnd)
    end select      

    function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnDestroy( byval hwnd as HWND ) as LRESULT
    dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
    if pWindow = 0 then exit function
    
    ' Destroy the manually allocated LISTVIEW_DATA_TYPE pointer
    if pWindow->UserData(0) > 0 then delete( cast(LISTVIEW_DATA_TYPE ptr, pWindow->UserData(0)) )

    ' Delete all the panel CWindow pointers because these child windows
    ' are about to get destroyed.
'    pWindow = AfxCWindowPtr( HWND_FRMPANEL_MENU ): if pWindow then delete(pWindow)

    function = 0
end Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmListView_PositionWindows( hwnd )
    end if
    function = 0
end function
        
        
' ========================================================================================
' Process WM_PAINT message for window/dialog: frmListView
' ========================================================================================
function frmListView_OnPaint( byval hwnd as HWND ) as LRESULT
            
    dim as clsDoubleBuffer b
    
    b.BeginDoubleBuffer(hwnd)

    ' Set default fore and back colors
    b.SetForeColors( ghOutput.ForeColor, ghOutput.ForeColor )
    b.SetBackColors( ghOutput.BackColor, ghOutput.BackColor )

    ' Clear the entire back client area
    b.PaintClientRect()
    
    b.EndDoubleBuffer()

    function = 0
end function


' ========================================================================================
' frmListView Window procedure
' ========================================================================================
function frmListView_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    select case uMsg
        HANDLE_MSG (hwnd, WM_SIZE,        frmListView_OnSize)
        HANDLE_MSG (hwnd, WM_PAINT,       frmListView_OnPaint)
        HANDLE_MSG (hwnd, WM_DESTROY,     frmListView_OnDestroy)
        HANDLE_MSG (hwnd, WM_COMMAND,     frmListView_OnCommand)
        HANDLE_MSG (hwnd, WM_MEASUREITEM, frmListView_OnMeasureItem)
        HANDLE_MSG (hwnd, WM_DRAWITEM,    frmListView_OnDrawItem)

    case WM_ERASEBKGND
        return true

    end select

    ' for messages that we don't deal with
    function = DefWindowProc( hwnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmListView_Show
' ========================================================================================
function frmListView_Show( byval hwndParent as hwnd ) as HWND

    '  Create the main window and child controls
    dim pWindow as CPanelWindow ptr = new CPanelWindow

    dim as HWND hLV = pWindow->Create( hwndParent, @frmListView_WndProc, 0, 0, 0, 0 )

    dim as LISTVIEW_DATA_TYPE ptr pData = new LISTVIEW_DATA_TYPE
    pWindow->UserData(0) = cast(LONG_PTR, pData)
    
    ' Must use GetDlgItem to retrieve the child window handles because there
    ' could be multiple instances of this ListView control.

    dim as HWND hHeader = _
        pWindow->AddControl( "LABEL", , IDC_FRMLISTVIEW_HEADER, "", 0, 0, 0, EXPLORERITEM_HEIGHT_COMPACT, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or SS_NOTIFY or SS_LEFT, _
        WS_EX_LEFT or WS_EX_LTRREADING, , _
        cast(SUBCLASSPROC, @frmListViewHeader_SubclassProc), _
        IDC_FRMLISTVIEW_HEADER, cast(DWORD_PTR, @pWindow))

    dim as HWND hListBox = _
        pWindow->AddControl( "LISTBOX", , IDC_FRMLISTVIEW_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or WS_TABSTOP or _
        LBS_NOINTEGRALHEIGHT or LBS_OWNERDRAWFIXED or LBS_HASSTRINGS or LBS_NOTIFY, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmListViewListBox_SubclassProc), _
        IDC_FRMLISTVIEW_LISTBOX, cast(DWORD_PTR, @pWindow))

'    dim pWindowPanel as CPanelWindow ptr = New CPanelWindow
'    HWND_FRMPANEL_VSCROLLBAR = _
'        pWindowPanel->Create( HWND_FRMPANEL, @frmPanelVScroll_WndProc, 0, 0, SCROLLBAR_WIDTH_PANEL, 0 )
'    pWindow->ClassStyle = CS_DBLCLKS

    frmListView_PositionWindows(hLV)
    
    function = hLV
    
end function
