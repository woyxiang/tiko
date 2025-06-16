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
''  frmPanel.inc
''   

#include once "frmPanel.bi"
#include once "frmExplorer.bi"
#include once "frmBookmarks.bi"
#include once "frmFunctions.bi"

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmPanel_PositionWindows() as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
    if pWindow = 0 then exit function
    
    ' Get the entire client area
    dim as Rect rc
    GetClientRect( HWND_FRMPANEL, @rc )

    redim gPanelButton(1) as PANEL_BUTTON_TYPE
    
    gPanelButton(1).rc.left = rc.left
    gPanelButton(1).rc.right = rc.right
    gPanelButton(1).rc.bottom = rc.bottom
    gPanelButton(1).rc.top = rc.bottom - pWindow->ScaleY(TOPTABS_HEIGHT)
    
    gPanelButton(0).rc.left = rc.left
    gPanelButton(0).rc.right = rc.right
    gPanelButton(0).rc.bottom = gPanelButton(1).rc.top 
    gPanelButton(0).rc.top = gPanelButton(1).rc.top - pWindow->ScaleY(TOPTABS_HEIGHT)

    if gApp.GetDocumentCount = 0 then
        SetRectEmpty( @gPanelButton(0).rc )
        SetRectEmpty( @gPanelButton(1).rc )
    end if

    ' show/hide of the child window is done separately in order to prevent blank client area from 
    ' painting before the lists are loaded. Allow a 2 pixel right border between listbox and
    ' form edge in order to allow for panel resizing. 
    dim as long nLeft = 0
    if gConfig.ExplorerPositionRight then nLeft = pWindow->ScaleX(3)  ' allow for grab
    
    dim as long nTop = rc.top  + pWindow->ScaleY(46)
    dim as long nWidth = rc.right  - rc.Left - pWindow->ScaleX(3)   
    dim as long nHeight = gPanelButton(0).rc.top - rc.top - nTop

    SetWindowPos( HWND_FRMEXPLORER, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER )
    SetWindowPos( HWND_FRMBOOKMARKS, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER )
    SetWindowPos( HWND_FRMFUNCTIONS, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER )

    select case gPanel.hActiveChild
        case HWND_FRMEXPLORER
            ShowWindow( HWND_FRMBOOKMARKS, SW_HIDE )
            ShowWindow( HWND_FRMFUNCTIONS, SW_HIDE )

            gPanel.wszLabel = L(247,"EXPLORER")
            ShowWindow( HWND_FRMEXPLORER, SW_SHOW )
            gPanelVScroll.hListBox = HWND_FRMEXPLORER_LISTBOX

            gPanelButton(0).wszCaption = ucase( L(223,"Function List") )
            gPanelButton(0).hActionChild = HWND_FRMFUNCTIONS
            gPanelButton(1).wszCaption = ucase( L(188,"Bookmarks") )
            gPanelButton(1).hActionChild = HWND_FRMBOOKMARKS
        
        case HWND_FRMBOOKMARKS
            ShowWindow( HWND_FRMEXPLORER, SW_HIDE )
            ShowWindow( HWND_FRMFUNCTIONS, SW_HIDE )

            gPanel.wszLabel = ucase( L(188,"Bookmarks") )
            ShowWindow( HWND_FRMBOOKMARKS, SW_SHOW )
            gPanelVScroll.hListBox = HWND_FRMBOOKMARKS_LISTBOX
        
            gPanelButton(0).wszCaption = L(247,"EXPLORER")
            gPanelButton(0).hActionChild = HWND_FRMEXPLORER
            gPanelButton(1).wszCaption = ucase( L(223,"Function List") )
            gPanelButton(1).hActionChild = HWND_FRMFUNCTIONS

        case HWND_FRMFUNCTIONS
            ShowWindow( HWND_FRMEXPLORER, SW_HIDE )
            ShowWindow( HWND_FRMBOOKMARKS, SW_HIDE )

            gPanel.wszLabel = ucase( L(223,"Function List") )
            ShowWindow( HWND_FRMFUNCTIONS, SW_SHOW )
            gPanelVScroll.hListBox = HWND_FRMFUNCTIONS_LISTBOX

            gPanelButton(0).wszCaption = L(247,"EXPLORER")
            gPanelButton(0).hActionChild = HWND_FRMEXPLORER
            gPanelButton(1).wszCaption = ucase( L(188,"Bookmarks") )
            gPanelButton(1).hActionChild = HWND_FRMBOOKMARKS
    end select

    ' calculate the rect for the panel label (EXPLORER, FUNCTION LIST, BOOKMARKS)
    with gPanel.rcLabel 
        .left = pWindow->ScaleX( 10 )
        .top = pWindow->ScaleY( 10 )
        .bottom = .top + pWindow->ScaleY( 20 )
        .right = .left + _
                pWindow->ScaleX( _
                    getTextWidth(HWND_FRMPANEL, gPanel.wszLabel, ghStatusBar.hFontStatusBar, 10) _
                )
    end with

    ' calculate the actual more actions "..." button itself
    dim as long ActionButtonWidth = 24
    dim as long ActionButtonHeight = 22
    with gPanel.rcActionMenu
        .left = rc.right - pWindow->ScaleX(ActionButtonWidth + 10)
        .top = pWindow->ScaleY(7)
        .right = .left + pWindow->ScaleX(ActionButtonWidth)
        .bottom = .top + pWindow->ScaleY(ActionButtonHeight)
    end with
    
    ' calculate the build/execute button
    dim as long BuildExecuteButtonWidth = 22
    dim as long BuildExecuteButtonHeight = 22
    with gPanel.rcBuildExecute
        .left = gPanel.rcActionMenu.left - pWindow->ScaleX(BuildExecuteButtonWidth)
        .top = pWindow->ScaleY(7)
        .right = .left + pWindow->ScaleX(BuildExecuteButtonWidth)
        .bottom = .top + pWindow->ScaleY(BuildExecuteButtonHeight)
    end with

    RedrawWindow( gPanelVScroll.hListBox, null, null, RDW_INVALIDATE or RDW_UPDATENOW )                
    RedrawWindow( HWND_FRMPANEL, null, null, RDW_INVALIDATE or RDW_UPDATENOW )                

    function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmPanel
' ========================================================================================
function frmPanel_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmPanel_PositionWindows
    end if
    function = 0
end function
        
' ========================================================================================
' Do hit test to determine if BuildExecute button in action Area was clicked
' ========================================================================================
function isPanelBuildExecuteButtonHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gPanel.rcBuildExecute, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test to determine if "..." button in action Area was clicked
' ========================================================================================
function isPanelActionButtonHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gPanel.rcActionMenu, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test on Button0 
' ========================================================================================
function isPanelButton0HitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gPanelButton(0).rc, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test on Button1
' ========================================================================================
function isPanelButton1HitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gPanelButton(1).rc, pt ) then return true
    function = false
end function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmPanel
' ========================================================================================
function frmPanel_OnPaint( byval hwnd as HWND ) as LRESULT
            
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
    if pWindow = 0 then exit function

    dim as PAINTSTRUCT ps
    dim as HDC hDc
    dim as HPEN hPenNull
    dim as HPEN oldPen
    dim as HFONT oldFont
    dim as HBRUSH oldBrush
    dim as RECT rc
    
    hDC = BeginPaint(hwnd, @ps)

    SaveDC(hDC)
    dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
    dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

    dim memDC as HDC      ' Double buffering
    dim hbit as HBITMAP   ' Double buffering

    memDC = CreateCompatibleDC( hDC )
    hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
    if hbit then hbit = SelectObject( memDC, hbit )

    FillRect( memDC, @ps.rcPaint, ghPanel.hPanelBrush )
    
    ' Draw the panel Label
    SetTextColor( memDC, ghPanel.forecolor )
    SetBkColor( memDC, ghPanel.backcolor )
    oldFont = SelectObject( memDC, ghStatusBar.hFontStatusBar )

    dim as long wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
    DrawText( memDC, gPanel.wszLabel.sptr, -1, cast(lpRect, @gPanel.rcLabel), wsStyle )
    SelectObject( memDC, oldFont )
    
    ' Draw the BuildExecute menu item only if there is currently an open file or
    ' project is active.
    dim as boolean bShowButton = false
    if gApp.IsProjectActive then
        if gApp.GetDocumentCount > 0 then bShowButton = true
    else
        if gTTabCtl.GetItemCount > 0 then bShowButton = true        
    end if
    if bShowButton then
        SetTextColor( memDC, ghPanel.ForeColorHot )
        if isPanelBuildExecuteButtonHitTest( hwnd ) then
            oldBrush = SelectObject( memDC, ghPanel.hBackBrushHot )
            SetBkColor( memDC, ghPanel.BackColorHot )
        else
            oldBrush = SelectObject( memDC, ghPanel.hPanelBrush )
            SetBkColor( memDC, ghPanel.BackColor )
        end if
        rc = gPanel.rcBuildExecute

        hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
        oldPen = SelectObject( memDC, hPenNull )
        RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
        DeleteObject SelectObject( memDC, oldPen )
        
        wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER
        oldFont = SelectObject( memDC, ghMenuBar.hFontSymbol )
        rc.top += pWindow->ScaleY(2): rc.bottom -= pWindow->ScaleY(2): 
        DrawText( memDC, wszBuildExecute, -1, cast(lpRect, @rc), wsStyle )
        SelectObject( memDC, oldFont )
        SelectObject( memDC, oldBrush )
    end if

    ' Draw the "..." menu item
    SetTextColor( memDC, ghPanel.ForeColorHot )
    if isPanelActionButtonHitTest( hwnd ) then
        oldBrush = SelectObject( memDC, ghPanel.hBackBrushHot )
        SetBkColor( memDC, ghPanel.BackColorHot )
    else
        oldBrush = SelectObject( memDC, ghPanel.hPanelBrush )
        SetBkColor( memDC, ghPanel.BackColor )
    end if
    rc = gPanel.rcActionMenu

    hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
    oldPen = SelectObject( memDC, hPenNull )
    RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
    DeleteObject SelectObject( memDC, oldPen )
    
    wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER
    oldFont = SelectObject( memDC, ghMenuBar.hFontSymbolLargeBold )
    DrawText( memDC, wszMoreActions, -1, cast(lpRect, @rc), wsStyle )
    SelectObject( memDC, oldFont )
    SelectObject( memDC, oldBrush )

    ' Draw the bottom two Panel buttons
    dim as long penWidth = pWindow->ScaleX(1)
    dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghTopTabs.Divider )
    oldPen = SelectObject( memDC, hPenSolid)

    dim as HBRUSH hBrush
    dim as COLORREF foreclr, backclr
    dim as boolean isHot 
    
    for i as long = lbound(gPanelButton) to ubound(gPanelButton)
        dim as RECT rc = gPanelButton(i).rc

        isHot = isMouseOverRECT( hwnd, rc )
        hBrush = iif( isHot, ghPanel.hBackBrushButtonHot, ghPanel.hBackBrushButton)
        backclr = iif( isHot, ghPanel.BackColorButtonHot, ghPanel.BackColorButton)
        foreclr = iif( isHot, ghPanel.ForeColorButtonHot, ghPanel.ForeColorButton)
        
        oldBrush = SelectObject( memDC, hBrush )
        SetBkColor( memDC, backclr )
        SetTextColor( memDC, foreclr )

        SelectPen( memDC, hPenSolid )
        Rectangle( memDC, rc.left, rc.top, rc.right, rc.bottom )
        wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
        oldFont = SelectObject( memDC, ghStatusBar.hFontStatusBar )
        DrawText( memDC, gPanelButton(i).wszCaption.sptr, -1, Cast(lpRect, @rc), wsStyle )

        SelectObject( memDC, oldFont )
        SelectObject( memDC, oldBrush )
    next   

    BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

    ' Cleanup
    DeleteObject SelectObject( memDC, oldPen )
    if hPenSolid then DeleteObject( hPenSolid )
    if hPenNull  then DeleteObject( hPenNull )
    if hbit  then DeleteObject( SelectObject(memDC, hbit) )
    if memDC then DeleteDC( memDC )

    RestoreDC( hDC, -1 )

    EndPaint( hwnd, @ps )

    function = 0
end function



' ========================================================================================
' frmPanel Window procedure
' ========================================================================================
function frmPanel_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    static as boolean isPrevHotBuildExecute, isPrevHotAction, isPrevHotButton0, isPrevHotButton1
    static hTooltip as hwnd

    select case uMsg
        HANDLE_MSG (hwnd, WM_SIZE,  frmPanel_OnSize)
        HANDLE_MSG (hwnd, WM_PAINT, frmPanel_OnPaint)
    
    case WM_ERASEBKGND
        return true

    case WM_MOUSEMOVE
        dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
        if pWindow = 0 then exit function

        dim tme as TrackMouseEvent
        tme.cbSize = sizeof(TrackMouseEvent)
        tme.dwFlags = TME_HOVER or TME_LEAVE
        tme.hwndTrack = hwnd
        TrackMouseEvent(@tme) 

        ' Repaint the gPanelMenuRect to ensure Hot highlighting
        dim as boolean isHotBuildExecute, isHotAction, isHotButton0, isHotButton1 
        isHotBuildExecute = isPanelBuildExecuteButtonHitTest( hwnd )
        isHotAction = isPanelActionButtonHitTest( hwnd )
        isHotButton0 = isPanelButton0HitTest( hwnd )
        isHotButton1 = isPanelButton1HitTest( hwnd )

        if isHotBuildExecute <> isPrevHotBuildExecute then   
            AfxRedrawWindow( hwnd )
            isPrevHotBuildExecute = isHotBuildExecute
        end if   

        if isHotAction <> isPrevHotAction then   
            AfxRedrawWindow( hwnd )
            isPrevHotAction = isHotAction
        end if   
            
        if isHotButton0 <> isPrevHotButton0 then   
            AfxRedrawWindow( hwnd )
            isPrevHotButton0 = isHotButton0
        end if   

        if isHotButton1 <> isPrevHotButton1 then   
            AfxRedrawWindow( hwnd )
            isPrevHotButton1 = isHotButton1
        end if



        ' PANELS LEFT/RIGHT SPLITTER
        dim as point pt
        dim as Rect rc
        GetWindowRect HWND_FRMPANEL, @rc

        if gConfig.ExplorerPositionRight then
            rc.left = rc.left - pWindow->ScaleX(3)
        else
            rc.Left = rc.Right - pWindow->ScaleX(3)
        end if

        GetCursorPos(@pt)
        if PtInRect( @rc, pt ) then
            if WindowFrompoint(pt) = HWND_FRMPANEL then
            SetCursor( ghCursorSizeWE )
            end if
        end if

        if gApp.bDragActive then
            if gApp.hwndPanel = HWND_FRMPANEL then
                ' Get the current rect of the panel and compare right side to the current cursor position
                ' and then move the right side to equal the cursor position.
                dim as long nDiff = pt.x - rc.right
                if gConfig.ExplorerPositionRight then nDiff = rc.left - pt.x
                GetClientRect HWND_FRMPANEL, @rc
                dim as long nWidth = (rc.right - rc.left + nDiff)
                ' make sure width does not go below a specific size otherwise the user will not
                ' be able to grab the panel edge to resize
                nWidth = max( nWidth, pWindow->ScaleX(80) )
                SetWindowPos( HWND_FRMPANEL, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE or SWP_NOZORDER )
                RedrawWindow( HWND_FRMPANEL, null, null, RDW_INVALIDATE or RDW_UPDATENOW or RDW_ALLCHILDREN)                
                frmMain_PositionWindows()
            end if
        end if
        
    case WM_LBUTTONDOWN
        dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
        if pWindow = 0 then exit function

        ' PANELS LEFT/RIGHT SPLITTER
        dim as Rect rc
        dim as point pt 
        gApp.bDragActive = false 
        GetWindowRect HWND_FRMPANEL, @rc

        if gConfig.ExplorerPositionRight then
            rc.left = rc.left - pWindow->ScaleX(3)
        else
            rc.Left = rc.Right - pWindow->ScaleX(3)
        end if

        GetCursorPos(@pt)
        if PtInRect( @rc, pt ) then
            if WindowFrompoint(pt) = HWND_FRMPANEL then
            SetCursor( ghCursorSizeWE )
            gApp.bDragActive = true 
            gApp.hwndPanel   = HWND_FRMPANEL
            SetCapture( HWND_FRMPANEL )
            end if
        end if

        
    case WM_LBUTTONUP
        if gApp.bDragActive then
            gApp.bDragActive = false 
            gApp.hwndPanel = 0
            ReleaseCapture()
        end if
        
        if isPanelBuildExecuteButtonHitTest( hwnd ) then
            OnCommand_CompileCommon(IDM_BUILDEXECUTE)
            AfxRedrawWindow(hwnd)
        end if
        
        if isPanelActionButtonHitTest( hwnd ) then
            dim as HMENU hPopupMenu 
            select case gPanel.hActiveChild
            case HWND_FRMEXPLORER
                hPopupMenu = CreateExplorerActionButtonContextMenu()
            case HWND_FRMFUNCTIONS
                hPopupMenu = CreateFunctionsActionButtonContextMenu()
            case HWND_FRMBOOKMARKS
                hPopupMenu = CreateBookmarksActionButtonContextMenu()
            end select
            ' Popup the menu to the bottom of the Action Button (right aligned)
            dim as RECT rc = gPanel.rcActionMenu   ' work with a copy
            MapWindowpoints( HWND_FRMPANEL, HWND_DESKTOP, cast(point ptr, @rc), 2 )
            TrackPopupMenu( hPopUpMenu, TPM_RIGHTALIGN, _
                                rc.right, rc.bottom, 0, HWND_FRMMAIN, byval null)
            DestroyMenu( hPopUpMenu )
            Return true   ' prevent further processing that leads to WM_CONTEXTMENU
        end if
        
        if isPanelButton0HitTest( hwnd ) then
            gPanel.hActiveChild = gPanelButton(0).hActionChild
            frmPanel_PositionWindows()
        end if
        
        if isPanelButton1HitTest( hwnd ) then
            gPanel.hActiveChild = gPanelButton(1).hActionChild
            frmPanel_PositionWindows()
        end if


    case WM_MOUSELEAVE
        isPrevHotBuildExecute = false
        isPrevHotAction = false
        isPrevHotButton0 = false
        isPrevHotButton1 = false
        AfxDeleteTooltip( hTooltip, hwnd )
        hTooltip = 0
        AfxRedrawWindow(hwnd)


    case WM_MOUSEHOVER
        dim as CWSTR wszTooltip 
        if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( hwnd, "", false, false )
        if isPanelBuildExecuteButtonHitTest( hwnd ) then
            wszTooltip  = L(65,"Build And Execute")
        end if
        if isPanelActionButtonHitTest( hwnd ) then
            wszTooltip  = L(440, "More Actions") & "..."
        end if
        ' Display the tooltip
        AfxSetTooltipText( hTooltip, hwnd, wszTooltip )
        AfxRedrawWindow( hwnd )
        
    end select

    ' for messages that we don't deal with
    function = DefWindowProc( hwnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmPanel_Show
' ========================================================================================
function frmPanel_Show( byval hwndParent as hwnd ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow
    pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

    ' Only make the Panel initially visible if it was already visible
    ' when the most previous instance of the program closed. Also, set the width of
    ' the window to the last used visible width.
    dim as long nWidth = iif(gConfig.ShowPanel, gConfig.ShowPanelWidth, 250)
    
    HWND_FRMPANEL = pWindow->Create( hwndParent, "Panel Window", @frmPanel_WndProc, _
        0, 0, nWidth, 0, _
        WS_CHILD or iif(gConfig.ShowPanel, WS_VISIBLE, 0) or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
    ' Disable background erasing by only assigning the one style
    pWindow->ClassStyle = CS_DBLCLKS

    ' Create the Explorer child window
    frmExplorer_Show( HWND_FRMPANEL )
    frmBookmarks_Show( HWND_FRMPANEL )
    frmFunctions_Show( HWND_FRMPANEL )
    gPanel.hActiveChild = HWND_FRMEXPLORER
    
    function = 0
    
end function
