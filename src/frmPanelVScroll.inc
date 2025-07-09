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
''  frmPanelVScroll.inc
''   

#include once "frmPanel.bi"


' ========================================================================================
' Calculate the RECT that holds the client coordinates of the scrollbar's vertical thumb
' Returns True if RECT is not empty
' ========================================================================================
function frmPanelVScroll_calcVThumbRect() as boolean
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
    if pWindow = 0 then exit function

    ' calculate the vertical scrollbar in client coordinates
    SetRectEmpty( @gPanelVScroll.rc )
    
    dim as long nTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 

    dim as RECT rc 
    with gPanelVScroll
        GetClientRect( gPanelVScroll.hListBox, @rc )
        .listBoxHeight = rc.bottom - rc.top
        .itemHeight = SendMessage( gPanelVScroll.hListBox, LB_GETITEMHEIGHT, 0, 0 )
        .numItems = ListBox_GetCount( gPanelVScroll.hListBox ) 
        .itemsPerPage = .listBoxHeight / .itemHeight
        .thumbHeight = (.itemsPerPage / .numItems) * .listBoxHeight
        GetClientRect( HWND_FRMPANEL_VSCROLLBAR, @rc )
        .rc.Left = rc.Left
        .rc.top = rc.top + ((nTopIndex / .numItems) * .listBoxHeight)
        .rc.Right = rc.right
        .rc.bottom = .rc.top + .thumbHeight
        if .numItems < .itemsPerPage then return true
    end with

    function = 0
end function

    
' ========================================================================================
' Position the VScrollBar over the Explorer listbox
' ========================================================================================
function frmPanelVScroll_PositionWindows( byval nShowState as long ) as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_VSCROLLBAR)
    if pWindow = 0 then exit function
    
    ' Position the VScrollBar relative to the Explorer Listbox using screen coordinates
    dim as Rect rc = AfxGetWindowRect( gPanelVScroll.hListBox )
    dim as long nScrollWidth = pWindow->ScaleX(SCROLLBAR_WIDTH_PANEL)
    
    SetWindowPos( HWND_FRMPANEL_VSCROLLBAR, HWND_TOP, _
                    rc.right - nScrollWidth, _
                    rc.top, _
                    nScrollWidth, _
                    rc.bottom - rc.top, _
                    SWP_NOACTIVATE )

    ' returns true if RECT is empty
    if frmPanelVScroll_calcVThumbRect() then nShowState = SW_HIDE 
    ShowWindow( HWND_FRMPANEL_VSCROLLBAR, nShowState )
    if nShowState = SW_HIDE then exit function

    AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )

    function = 0
end function


' ========================================================================================
' frmPanelVScroll Window procedure
' ========================================================================================
function frmPanelVScroll_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    static as point prev_pt   ' screen pt.y cursor position

    select case uMsg
        case WM_LBUTTONDOWN
            dim as point pt: GetCursorPos( @pt )
            frmPanelVScroll_calcVThumbRect()   ' in client coordinates
            dim as RECT rc = gPanelVScroll.rc ' covert copy to screen coordinates
            MapWindowpoints( HWND_FRMPANEL_VSCROLLBAR, HWND_DESKTOP, cast(point ptr, @rc), 2)
            if PtInRect( @rc, pt ) then
                prev_pt = pt
                gApp.bDragActive = true
                SetCapture( HWnd )
            else
                ' we have clicked on a PageUp or PageDn
                dim as long nTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 
                if pt.y < rc.top then
                    nTopIndex = max( nTopIndex - gPanelVScroll.itemsPerPage, 0 )
                    SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
                    frmPanelVScroll_calcVThumbRect()   ' in client coordinates
                    AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
                    AfxRedrawWindow( gPanelVScroll.hListBox )
                elseif pt.y > rc.bottom then
                    dim as long nMaxTopIndex = gPanelVScroll.numItems - gPanelVScroll.itemsPerPage
                    nTopIndex = min( nTopIndex + gPanelVScroll.itemsPerPage, nMaxTopIndex )
                    SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
                    frmPanelVScroll_calcVThumbRect()   ' in client coordinates
                    AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
                    AfxRedrawWindow( gPanelVScroll.hListBox )
                end if
            end if
            
        case WM_MOUSEMOVE
            if gApp.bDragActive then
                dim as point pt: GetCursorPos( @pt )
                if pt.y <> prev_pt.y then 
                    dim as long delta = pt.y - prev_pt.y 
                    ' convert to client coordinates for ease of use
                    dim as RECT rc: GetClientRect( HWND_FRMPANEL_VSCROLLBAR, @rc )

                    gPanelVScroll.rc.top = max(0, gPanelVScroll.rc.top + delta)
                    gPanelVScroll.rc.top = min(gPanelVScroll.rc.top, rc.bottom-gPanelVScroll.thumbHeight)
                    gPanelVScroll.rc.bottom = gPanelVScroll.rc.top + gPanelVScroll.thumbHeight

                    prev_pt = pt
                    
                    dim as long nPrevTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 
                    dim as long nLastIndex = (gPanelVScroll.rc.bottom / rc.bottom) * gPanelVScroll.numItems
                    dim as long nTopIndex = nLastIndex - gPanelVScroll.itemsPerPage
                    if nTopIndex <> nPrevTopIndex then
                        SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
                        AfxRedrawWindow( gPanelVScroll.hListBox )
                    end if   
                    AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
                end if
            end if
            
        case WM_LBUTTONUP   
            gApp.bDragActive = false
            prev_pt.x = 0
            prev_pt.y = 0
            ReleaseCapture
        
        case WM_ERASEBKGND
            return true
            
        case WM_PAINT
            dim as clsDoubleBuffer b

            b.BeginDoubleBuffer(hwnd)
            b.SetBackColors( ghPanel.BackColorScrollBar, ghPanel.BackColorScrollBar )
            b.PaintClientRect()
            b.SetBackColors( ghPanel.ForeColorScrollBar, ghPanel.ForeColorScrollBar )
            b.PaintRect( @gPanelVScroll.rc )

            b.EndDoubleBuffer()

            return 0

    end select

    ' for messages that we don't deal with
    function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmPanelVScroll_Show
' ========================================================================================
function frmPanelVScroll_Show( byval hWndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    ' Pointer destroyed at end of frmMain prior to application end
    dim pWindow as CPanelWindow ptr = new CPanelWindow

    HWND_FRMPANEL_VSCROLLBAR = pWindow->Create( hWndParent, _
        @frmPanelVScroll_WndProc, 0, 0, SCROLLBAR_WIDTH_PANEL, 0, _
        WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_NOACTIVATE )

    function = 0
    
end function
