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
''  frmPanelMenu.inc
''   

#include once "frmPanelMenu.bi"


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmPanelMenu_PositionWindows() as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_MENU)
    if pWindow = 0 then exit function
    
    ' Get the entire client area
    dim as Rect rc
    GetClientRect( HWND_FRMPANEL_MENU, @rc )


    function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmPanel
' ========================================================================================
function frmPanelMenu_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmPanelMenu_PositionWindows
    end if
    function = 0
end function
        

' ========================================================================================
' Process WM_PAINT message for window/dialog: frmPanelMenu
' ========================================================================================
function frmPanelMenu_OnPaint( byval hwnd as HWND ) as LRESULT
            
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_MENU)
    if pWindow = 0 then exit function

    dim as PAINTSTRUCT ps
    dim as HDC hDc
    dim as HPEN hPenNull
    dim as HPEN oldPen
    dim as HFONT oldFont
    dim as HBRUSH oldBrush
    dim as RECT rc
    
    dim as long wsStyle 
    
    hDC = BeginPaint(hwnd, @ps)

    SaveDC(hDC)
    dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
    dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

    dim memDC as HDC      ' Double buffering
    dim hbit as HBITMAP   ' Double buffering

    memDC = CreateCompatibleDC( hDC )
    hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
    if hbit then hbit = SelectObject( memDC, hbit )

    FillRect( memDC, @ps.rcPaint, GetSysColorBrush(COLOR_HOTLIGHT))    'ghPanel.hPanelBrush )

'    ' calculate the build/execute button
'    dim as long BuildExecuteButtonWidth = 22
'    dim as long BuildExecuteButtonHeight = 22
'    with gPanel.rcBuildExecute
'        .left = 0
'        .top = pWindow->ScaleY(7)
'        .right = .left + pWindow->ScaleX(BuildExecuteButtonWidth)
'        .bottom = .top + pWindow->ScaleY(BuildExecuteButtonHeight)
'    end with

    
'    ' Draw the BuildExecute menu item only if there is currently an open file or
'    ' project is active.
'    dim as boolean bShowButton = false
'    if gApp.IsProjectActive then
'        if gApp.GetDocumentCount > 0 then bShowButton = true
'    else
'        if gTTabCtl.GetItemCount > 0 then bShowButton = true        
'    end if
'    if bShowButton then
'        SetTextColor( memDC, ghPanel.ForeColorHot )
'        if isMouseOverRECT( hwnd, gPanel.rcBuildExecute ) then
'            oldBrush = SelectObject( memDC, ghPanel.hBackBrushHot )
'            SetBkColor( memDC, ghPanel.BackColorHot )
'        else
'            oldBrush = SelectObject( memDC, ghPanel.hPanelBrush )
'            SetBkColor( memDC, ghPanel.BackColor )
'        end if
'        rc = gPanel.rcBuildExecute

'        hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
'        oldPen = SelectObject( memDC, hPenNull )
'        RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
'        DeleteObject SelectObject( memDC, oldPen )
        
'        wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
'        oldFont = SelectObject( memDC, ghMenuBar.hFontSymbol )
'        DrawText( memDC, wszBuildExecute, -1, cast(lpRect, @rc), wsStyle )
'        SelectObject( memDC, oldFont )
'        SelectObject( memDC, oldBrush )
'    end if


    BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

    ' Cleanup
    DeleteObject SelectObject( memDC, oldPen )
    if hPenNull  then DeleteObject( hPenNull )
    if hbit  then DeleteObject( SelectObject(memDC, hbit) )
    if memDC then DeleteDC( memDC )

    RestoreDC( hDC, -1 )

    EndPaint( hwnd, @ps )

    function = 0
end function



' ========================================================================================
' frmPanelMenu Window procedure
' ========================================================================================
function frmPanelMenu_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    static as boolean isPrevHotBuildExecute
    static hTooltip as hwnd

    select case uMsg
        HANDLE_MSG (hwnd, WM_SIZE,  frmPanelMenu_OnSize)
        HANDLE_MSG (hwnd, WM_PAINT, frmPanelMenu_OnPaint)
    
    case WM_ERASEBKGND
        return true

    case WM_MOUSEMOVE
        dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_MENU)
        if pWindow = 0 then exit function

        dim tme as TrackMouseEvent
        tme.cbSize = sizeof(TrackMouseEvent)
        tme.dwFlags = TME_HOVER or TME_LEAVE
        tme.hwndTrack = hwnd
        TrackMouseEvent(@tme) 

'        ' Repaint the gPanelMenuRect to ensure Hot highlighting
'        dim as boolean isHotBuildExecute, isHotAction, isHotButton0, isHotButton1 
'        isHotBuildExecute = isPanelBuildExecuteButtonHitTest( hwnd )
'        isHotAction = isPanelActionButtonHitTest( hwnd )
'        isHotButton0 = isPanelButton0HitTest( hwnd )
'        isHotButton1 = isPanelButton1HitTest( hwnd )

'        if isHotBuildExecute <> isPrevHotBuildExecute then   
'            AfxRedrawWindow( hwnd )
'            isPrevHotBuildExecute = isHotBuildExecute
'        end if   

'        if isHotAction <> isPrevHotAction then   
'            AfxRedrawWindow( hwnd )
'            isPrevHotAction = isHotAction
'        end if   
            
'        if isHotButton0 <> isPrevHotButton0 then   
'            AfxRedrawWindow( hwnd )
'            isPrevHotButton0 = isHotButton0
'        end if   

'        if isHotButton1 <> isPrevHotButton1 then   
'            AfxRedrawWindow( hwnd )
'            isPrevHotButton1 = isHotButton1
'        end if


    case WM_LBUTTONDOWN
        dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_MENU)
        if pWindow = 0 then exit function

        
    case WM_LBUTTONUP
'        if isPanelBuildExecuteButtonHitTest( hwnd ) then
'            OnCommand_CompileCommon(IDM_BUILDEXECUTE)
'            AfxRedrawWindow(hwnd)
'        end if
        

    case WM_MOUSELEAVE
'        isPrevHotBuildExecute = false
'        isPrevHotAction = false
'        isPrevHotButton0 = false
'        isPrevHotButton1 = false
        AfxDeleteTooltip( hTooltip, hwnd )
        hTooltip = 0
        AfxRedrawWindow(hwnd)


    case WM_MOUSEHOVER
        dim as CWSTR wszTooltip 
        if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( hwnd, "", false, false )
'        if isPanelBuildExecuteButtonHitTest( hwnd ) then
'            wszTooltip  = L(65,"Build And Execute")
'        end if
'        if isPanelActionButtonHitTest( hwnd ) then
'            wszTooltip  = L(440, "More Actions") & "..."
'        end if
        ' Display the tooltip
        AfxSetTooltipText( hTooltip, hwnd, wszTooltip )
        AfxRedrawWindow( hwnd )
        
    end select

    ' for messages that we don't deal with
    function = DefWindowProc( hwnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmPanelMenu_Show
' ========================================================================================
function frmPanelMenu_Show( byval hwndParent as hwnd ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow

    HWND_FRMPANEL_MENU = pWindow->Create( hwndParent, "Panel Menu", @frmPanelMenu_WndProc, _
        0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
    ' Disable background erasing by only assigning the one style
    pWindow->ClassStyle = CS_DBLCLKS

    function = 0
    
end function
