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

#include once "frmHelpViewer.bi"


#define IDC_LEFTPANEL   1000
#define IDC_RIGHTPANEL   1001

dim shared as HWND HWND_LEFTPANEL
dim shared as HWND HWND_RIGHTPANEL

dim shared as POINT ptSplitPrev
dim shared as long xDeltaSplitter

' ========================================================================================
' Calculate RECT area of the splitter bar
' ========================================================================================
function frmHelpViewer_CalcSplitRect() as RECT
    dim pWindow as CWindow ptr = AfxCWindowPtr( HWND_FRMMAIN )
    if pWindow = 0 then exit function

    dim as long nSplitterWidth = pWindow->ScaleX(4)
    dim as RECT rc
    GetClientRect( HWND_FRMHELPVIEWER, @rc )
    rc.left = AfxGetWindowWidth( HWND_LEFTPANEL )
    rc.right = rc.left + nSplitterWidth

    function = rc
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmHelpViewer_PositionWindows() as LRESULT
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMHELPVIEWER)
    if pWindow = 0 then exit function

    dim as long nLeftPanelWidth = AfxGetWindowWidth(HWND_LEFTPANEL)
    if nLeftPanelWidth = 0 then nLeftPanelWidth = pWindow->ScaleX(200)
    
    dim as long nSplitterWidth = pWindow->ScaleX(4)

    dim as RECT rcClient: GetClientRect(HWND_FRMHELPVIEWER, @rcClient)
    dim as long nClientHeight = rcClient.bottom - rcClient.top
    dim as long nClientWidth = rcClient.right - rcClient.left
    
    dim as long xDelta = 0
    if gApp.bDragActive = true then xDelta = xDeltaSplitter
        
    nLeftPanelWidth += xDelta
    if nLeftPanelWidth < nSplitterWidth then nLeftPanelWidth = nSplitterWidth
    if nLeftPanelWidth >= (nClientWidth - nSplitterWidth * 2) then _
        nLeftPanelWidth = (nClientWidth - nSplitterWidth * 2)

    SetWindowPos( HWND_LEFTPANEL, 0, _
        0, 0, nLeftPanelWidth, nClientHeight, _
        SWP_NOZORDER or SWP_SHOWWINDOW )
    
    SetWindowPos( HWND_RIGHTPANEL, 0, _
        nLeftPanelWidth + nSplitterWidth, 0, _
        nClientWidth - nSplitterWidth - nLeftPanelWidth, nClientHeight, _
        SWP_NOZORDER or SWP_SHOWWINDOW )
    
    function = 0
end function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnMouseMove( _
            byval hwnd as HWND, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

    ' HITTEST (SPLITTER)
    dim as RECT rcSplitter = frmHelpViewer_CalcSplitRect()
    if (isMouseOverRECT(hwnd, rcSplitter) = true) orelse (gApp.bDragActive = true) then
        SetCursor( ghCursorSizeWE )
    else    
        SetCursor( LoadCursor( null, IDC_ARROW ))
    end if

    if (gApp.bDragActive = true) andalso (x <> ptSplitPrev.x) then
        xDeltaSplitter = (x - ptSplitPrev.x)
        frmHelpViewer_PositionWindows
    end if

    ptSplitPrev = type(x, y)
            
    function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnLButtonDown( _
            byval hwnd as HWND, _
            byval fDoubleClick as boolean, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long
                                        
    dim as RECT rcSplitter = frmHelpViewer_CalcSplitRect()
    if isMouseOverRECT( hwnd, rcSplitter ) then
        gApp.bDragActive = true
        ptSplitPrev = type(x, y)
        SetCursor(ghCursorSizeWE)
        SetCapture( HWND_FRMHELPVIEWER )
    end if

    function = 0
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnLButtonUp( _
            byval hwnd as HWND, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

    if gApp.bDragActive then
       gApp.bDragActive = false
       frmHelpViewer_PositionWindows
       ReleaseCapture
    end if   
    SetCursor( LoadCursor( null, IDC_ARROW ))

    function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmHelpViewer_PositionWindows
    end if
    function = 0
end Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnCreate( _
            byval hwnd as HWND, _
            byval lpCreateStructPtr as LPCREATESTRUCT _
            ) as boolean

    '  Message cracker macro expects a True to be returned for a successful
    '  OnCreate handler even though returning -1 from a standard WM_CREATE
    '  call would stop creating the window. This is just one of those Windows
    '  inconsistencies.
    return true
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnCommand( _
            byval hwnd as HWND, _
            byval id as long, _
            byval hwndCtl as HWND, _
            byval codeNotify as UINT _
            ) as LRESULT

    select case id
        
'        case IDC_FRMOPTIONS_CMDOK
'            if codeNotify = BN_CLICKED then
'                frmOptions_SaveEditorOptions()
'                ' Reload the theme file in case it has changed
'                LoadThemeFile()
'                ' Refresh the frmOutput controls (listviews and RichEdits)
'                AfxRedrawWindow( HWND_FRMOUTPUT )
'                frmOutput_SetControlColors()
'                frmOutput_PositionWindows()
'                ' Refresh the menubar. The rest of the GUI will refresh via frmMain_PositionWindows
'                AfxRedrawWindow( HWND_FRMMAIN_MENUBAR )
'                frmMenuBar_PositionWindows()
                
'                ' Reposition main windows in case hide toolbar/statusbar selected
'                frmMain_PositionWindows
                
'                ' Apply the newly saved options to all open Scintilla windows
'                dim pFrame as CWindow ptr
'                dim pDoc as clsDocument Ptr = gApp.pDocList
'                do until pDoc = 0
'                    pDoc->ApplyProperties
'                    pDoc = pDoc->pDocNext
'                loop   
'                SendMessage( hwnd, WM_CLOSE, 0, 0 )
'                exit function
'            end if
        
        case IDCANCEL
            if codeNotify = BN_CLICKED then
                SendMessage( hwnd, WM_CLOSE, 0, 0 )
                exit function
            end if
    
    end select

    function = 0
end function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnNotify( _
            byval hwnd as HWND, _
            byval id as long, _
            byval pNMHDR as NMHDR ptr _
            ) as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
    
    function = 0
end function


' ========================================================================================
' Process WM_CTLCOLORSTATIC message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnCtlColorStatic( _
            byval hwnd as HWND, _
            byval hdc as HDC, _
            byval hWndChild as HWND, _
            byval nType as long _
            ) as HBRUSH

    if hWndChild = HWND_LEFTPANEL then
        return GetSysColorBrush( COLOR_BTNFACE )
    elseif hWndChild = HWND_RIGHTPANEL then
        return GetSysColorBrush( COLOR_HIGHLIGHT )
    end if
    
    function = 0      
end function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnPaint( byval hwnd as HWND ) as LRESULT
    dim as clsDoubleBuffer b

    b.BeginDoubleBuffer(hwnd)
    b.SetBackColors( ghMain.BackColor, ghMain.BackColor )
    b.PaintClientRect()

    ' Draw any splitter between the two panels
    dim as RECT rcSplitter = frmHelpViewer_CalcSplitRect()
    'b.SetBackColors( ghEditor.Divider, ghEditor.Divider )
    b.SetBackColors( bgr(255,0,0), bgr(255,0,0) )
    b.PaintRect( @rcSplitter )
    
    b.EndDoubleBuffer()
    
    function = 0
end function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnClose( byval hwnd as HWND ) as LRESULT
    ' Enables parent window keeping parent's zorder
'    EnableAllModeless()
    DestroyWindow hwnd
    function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHelpViewer
' ========================================================================================
function frmHelpViewer_OnDestroy( byval hwnd as HWND ) as LRESULT
    dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
    if pWindow then delete(pWindow)
    
    HWND_FRMHELPVIEWER = 0

    function = 0
end function


' ========================================================================================
' frmOptions Window procedure
' ========================================================================================
function frmHelpViewer_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    select case uMsg
        HANDLE_MSG (hwnd, WM_SIZE,           frmHelpViewer_OnSize)
        HANDLE_MSG (hwnd, WM_PAINT,          frmHelpViewer_OnPaint)
        HANDLE_MSG (hwnd, WM_CREATE,         frmHelpViewer_OnCreate)
        HANDLE_MSG (hwnd, WM_CLOSE,          frmHelpViewer_OnClose)
        HANDLE_MSG (hwnd, WM_DESTROY,        frmHelpViewer_OnDestroy)
        HANDLE_MSG (hwnd, WM_COMMAND,        frmHelpViewer_OnCommand)
        HANDLE_MSG (hwnd, WM_NOTIFY,         frmHelpViewer_OnNotify)
        HANDLE_MSG (hwnd, WM_CTLCOLORSTATIC, frmHelpViewer_OnCtlColorStatic)
        HANDLE_MSG (hwnd, WM_LBUTTONDOWN,    frmHelpViewer_OnLButtonDown)
        HANDLE_MSG (hwnd, WM_LBUTTONUP,      frmHelpViewer_OnLButtonUp)
        HANDLE_MSG (hwnd, WM_MOUSEMOVE,      frmHelpViewer_OnMouseMove)
    end select

    function = DefWindowProc(hwnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' frmHelpViewerListBox_SubclassProc 
' ========================================================================================
function frmHelpViewerListBox_SubclassProc ( _
            byval hWin   as HWND, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
            ) as LRESULT

    static as long accumDelta
    
    ' keep track of last index we were over so that we only issue a 
    ' repaint if the cursor has moved off of the line
    static as long nLastIdx = -1
        
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
'                frmPanelVScroll_PositionWindows( SW_SHOWNA )
            elseif accumDelta <= -120 then  ' scroll down 3 lines
                nTopIndex = nTopIndex + 3
                SendMessage( hWin, LB_SETTOPINDEX, nTopIndex, 0 ) 
                accumDelta = 0
'                frmPanelVScroll_PositionWindows( SW_SHOWNA )
            end if

        case WM_MOUSEMOVE
            ' Track that we are over the control in order to catch the 
            ' eventual WM_MOUSEHOVER and WM_MOUSELEAVE events
            dim tme as TrackMouseEvent
            tme.cbSize = sizeof(TrackMouseEvent)
            tme.dwFlags = TME_HOVER or TME_LEAVE
            tme.hwndTrack = hWin
            tme.dwHoverTime = 1
            TrackMouseEvent(@tme) 

            ' get the item rect that the mouse is over and only invalidate
            ' that instead of the entire listbox
            dim as RECT rc
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
            ' if the specified point is in the client area of the list box, or one if it is outside the 
            ' client area.
            if hiword(idx) <> 1 then
                if idx <> nLastIdx then
                    ListBox_GetItemRect( hWin, idx, @rc )
                    InvalidateRect( hWin, @rc, true )
                    ListBox_GetItemRect( hWin, nLastIdx, @rc )
                    InvalidateRect( hWin, @rc, true )
                    nLastIdx = idx
                end if
            end if
                            
        case WM_MOUSELEAVE
            nLastIdx = -1
            AfxRedrawWindow(hWin)   
            
        case WM_LBUTTONUP
            ' determine if we clicked on a function name or a node header
            dim as RECT rc
            dim as long idx = Listbox_ItemFromPoint( hWin, GET_X_LPARAM(_lParam), GET_Y_LPARAM(_lParam)) 
            ' The return value contains the index of the nearest item in the LOWORD. The HIWORD is zero 
            ' if the specified point is in the client area of the list box, or one if it is outside the 
            ' client area.
            if hiword(idx) <> 1 then
'                dim as clsDocument ptr pDoc = cast(clsDocument ptr, ListBox_GetItemData( hWin, idx ))
'                dim as CWSTR wszCaption = AfxGetListBoxText( hWin, idx )
'                if (left(wszCaption, 4) = "true") orelse (left(wszCaption, 5) = "false") then
'                    OpenSelectedDocument( pDoc->DiskFilename )
'                else
'                    ' Attempt to show the function name
'                    dim as long nLineNum = getFunctionsLinenumber( wszCaption )
'                    dim as CWSTR wszFunctionName = getFunctionsFunctionName( wszCaption )
'                    dim as CWSTR wszDiskFilename 
'                    if pDoc then wszDiskFilename = pDoc->DiskFilename
'                    OpenSelectedDocument( wszDiskFilename, wszFunctionName, nLineNum )
'                end if
            end if

                                    
        case WM_ERASEBKGND
            ' if the number of lines in the listbox maybe less than the number per page then 
            ' calculate from last item to bottom of listbox, otherwise calculate based on
            ' the mod of the lineheight to listbox height so we can color the partial line
            ' that won't be displayed at the bottom of the list.
            dim as RECT rc: GetClientRect( hWin, @rc )

            dim as RECT rcItem  
            SendMessage( hWin, LB_GETITEMRECT, 0, cast(LPARAM, @rcItem) )
            dim as long itemHeight = rcItem.bottom - rcItem.top
            dim as long NumItems = ListBox_GetCount(hWin)
            dim as long nTopIndex = SendMessage( hWin, LB_GETTOPINDEX, 0, 0 ) 
            dim as long visible_rows = 0
            dim as long ItemsPerPage = 0
            dim as long bottom_index = 0
                        
            if NumItems > 0 then
                ItemsPerPage = (rc.bottom - rc.top) / itemHeight
                bottom_index = (nTopIndex + ItemsPerPage) 
                if bottom_index >= NumItems then bottom_index = NumItems - 1
                visible_rows = (bottom_index - nTopIndex) + 1
                rc.top = visible_rows * itemHeight 
            end if

            if rc.top < rc.bottom then
                dim as HDC _hDC = cast(HDC, _wParam)
                PaintRect( _hDC, @rc, ghPanel.BackColor )
            end if

            ValidateRect( hWin, @rc )
            return true
        
        Case WM_DESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hWin, @frmHelpViewerListBox_SubclassProc, uIdSubclass )
    end select
    
    ' For messages that we don't deal with
    function = DefSubclassProc( hWin, uMsg, _wParam, _lParam )

end function


' ========================================================================================
' frmHelpViewer_Show
' ========================================================================================
function frmHelpViewer_Show( byval hWndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow

    dim as CWSTR wszText = APPNAME & " " & L(73,"Help")
    HWND_FRMHELPVIEWER = pWindow->Create( _
        hWndParent, wszText, @frmHelpViewer_WndProc, 0, 0, 0, 0, _
        WS_POPUP or WS_OVERLAPPEDWINDOW or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT )
    pWindow->ClassStyle = CS_DBLCLKS

    ' Set the small and large icon for the main window (must be set after main window is created)
    pWindow->BigIcon   = LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
    pWindow->SmallIcon = LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

    pWindow->SetClientSize(625, 460)
    pWindow->Center(pWindow->hWindow, hWndParent)

'    HWND_FRMHELPVIEWER_LISTBOX = _ 
'        pWindow->AddControl("LISTBOX", , IDC_FRMHELPVIEWER_LISTBOX, "", 0, 0, 200, 460, _
'        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or WS_TABSTOP or _
'        LBS_NOINTEGRALHEIGHT or LBS_OWNERDRAWFIXED or LBS_HASSTRINGS or LBS_NOTIFY, _
'        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
'        cast(SUBCLASSPROC, @frmHelpViewerListBox_SubclassProc), _
'        IDC_FRMHELPVIEWER_LISTBOX, cast(DWORD_PTR, @pWindow))

    HWND_LEFTPANEL = pWindow->AddControl("LABEL", , IDC_LEFTPANEL, "", 0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or SS_LEFT or SS_NOTIFY, _
        WS_EX_LEFT or WS_EX_LTRREADING)

    HWND_RIGHTPANEL = pWindow->AddControl("LABEL", , IDC_RIGHTPANEL, "", 0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or SS_LEFT or SS_NOTIFY, _
        WS_EX_LEFT or WS_EX_LTRREADING)

    frmHelpViewer_PositionWindows()
    
    ShowWindow( HWND_FRMHELPVIEWER, SW_SHOW )
        
    return 0
end function

