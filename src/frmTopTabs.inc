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


' ========================================================================================
' Show or hide the drop down shadow that will display below the top tabs control. The
' width of the shadow is also influenced by whether the Find/Replace dialog is active.
' ========================================================================================
function frmTopTabs_ShowShadow() as long
    
    if IsWindowVisible( HWND_FRMMAIN_TOPTABS ) = 0 then exit function
    dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
    
    if pDoc then
        dim as RECT rcFindReplace = AfxGetWindowRect( HWND_FRMFINDREPLACE )
        dim as RECT rcTopTabs = AfxGetWindowRect( HWND_FRMMAIN_TOPTABS )
        dim as long nLeftWidth = rcTopTabs.right - rcTopTabs.left

        dim as long nLine = SciExec( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0)
        dim as long nShow = iif( nLine > 0, SW_SHOWNA, SW_HIDE )
        
        ' if the Find/Replace popup window is active then we always show the show
        ' the shadow window is a popup style so need to use screen coordinates
        if IsWindowVisible( HWND_FRMFINDREPLACE ) then 
            nShow = SW_SHOWNA
            nLeftWidth = rcFindReplace.left - rcTopTabs.left 
        end if
            
        SetWindowPos( HWND_FRMMAIN_TOPTABS_SHADOW, 0, _
                    rcTopTabs.left, rcTopTabs.bottom, _
                    nLeftWidth, AfxGetWindowHeight( HWND_FRMMAIN_TOPTABS_SHADOW ), _
                    SWP_NOZORDER or SWP_NOACTIVATE )

        ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, nShow ) 

    else
        ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, SW_HIDE )
    end if   

    function = 0
end function


' ========================================================================================
' Adjust all of the Tab rects left and right dimensions by the incoming amount
' ========================================================================================
function frmTopTabs_AdjustTabRects( byval nAdjAmount as long ) as long
    ' nAdjAmount can be +/-
    dim pDoc as clsDocument ptr
    
    for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs) 
        pDoc = gTTabCtl.tabs(i).pDoc
        if pDoc then
            gTTabCtl.tabs(i).rcTab.left += nAdjAmount
            gTTabCtl.tabs(i).rcTab.right += nAdjAmount
            gTTabCtl.tabs(i).rcIcon.left += nAdjAmount
            gTTabCtl.tabs(i).rcIcon.right += nAdjAmount
            gTTabCtl.tabs(i).rcText.left += nAdjAmount
            gTTabCtl.tabs(i).rcText.right += nAdjAmount
            gTTabCtl.tabs(i).rcClose.left += nAdjAmount
            gTTabCtl.tabs(i).rcClose.right += nAdjAmount
        end if
    next   
        
    function = 0
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmTopTabs_PositionWindows() as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN_TOPTABS)
    if pWindow = 0 then exit function
    
    ' Get the entire client area
    dim as Rect rc: GetClientRect( HWND_FRMMAIN_TOPTABS, @rc )
    
    ' Calculate the RECT positions for all of the top tabs
    dim pDoc as clsDocument ptr 
    dim as long nTextWidth = 0
    dim as long nLeft = rc.left 
    dim as long TabHeight = TOPTABS_HEIGHT
    dim as long rcTextWidth = 70
    dim as long rcCloseWidth = 20
    dim as long rcIconWidth = 20 
    dim as long rightBorderWidth = 2
    dim as long hmargin = 3
    dim as long ActionPanelWidth = 116
    dim as long ActionButtonWidth = 24
    dim as long ActionButtonHeight = 22
    dim as long PrevTabsButtonWidth = 20
    dim as long PrevTabsButtonHeight = 20
    dim as long SplitButtonWidth = 20
    dim as long SplitButtonHeight = 20
    dim as long AddFileButtonWidth = 24
    
    ' calculate the right hand side Action area panel that holds the
    ' more actions "..." button.
    gTTabCtl.rcActionPanel = rc
    gTTabCtl.rcActionPanel.left = gTTabCtl.rcActionPanel.right - pWindow->ScaleX( ActionPanelWidth )
    
    ' calculate the Prev and Next Tabs buttons
    dim as long vmargin = (TabHeight - PrevTabsButtonHeight) / 2
    gTTabCtl.rcPrevTabs.left = gTTabCtl.rcActionPanel.left + pWindow->ScaleX( hmargin * 2)
    gTTabCtl.rcPrevTabs.top = rc.top + pWindow->ScaleY(vmargin)
    gTTabCtl.rcPrevTabs.bottom = gTTabCtl.rcPrevTabs.top + pWindow->ScaleY( PrevTabsButtonHeight )
    gTTabCtl.rcPrevTabs.right = gTTabCtl.rcPrevTabs.left + pWindow->ScaleX( PrevTabsButtonWidth )

    gTTabCtl.rcNextTabs.left = gTTabCtl.rcPrevTabs.right + pWindow->ScaleX( hmargin )
    gTTabCtl.rcNextTabs.top = rc.top + pWindow->ScaleY(vmargin)
    gTTabCtl.rcNextTabs.bottom = gTTabCtl.rcNextTabs.top + pWindow->ScaleY( PrevTabsButtonHeight )
    gTTabCtl.rcNextTabs.right = gTTabCtl.rcNextTabs.left + pWindow->ScaleX( PrevTabsButtonWidth )

    ' calculate the split editor button
    vmargin = (TabHeight - SplitButtonHeight) / 2
    gTTabCtl.rcSplitEditor.left = gTTabCtl.rcNextTabs.right + pWindow->ScaleX( hmargin )
    gTTabCtl.rcSplitEditor.top = rc.top + pWindow->ScaleY(vmargin)
    gTTabCtl.rcSplitEditor.bottom = gTTabCtl.rcSplitEditor.top + pWindow->ScaleY( SplitButtonHeight )
    gTTabCtl.rcSplitEditor.right = gTTabCtl.rcSplitEditor.left + pWindow->ScaleY( SplitButtonWidth )

    ' calculate the actual more actions "..." button itself
    vmargin = (TabHeight - ActionButtonHeight) / 2
    gTTabCtl.rcActionButton.left = gTTabCtl.rcSplitEditor.right + pWindow->ScaleX( hmargin )
    gTTabCtl.rcActionButton.top = rc.top + pWindow->ScaleY(vmargin)
    gTTabCtl.rcActionButton.bottom = gTTabCtl.rcActionButton.top + pWindow->ScaleY( ActionButtonHeight )
    gTTabCtl.rcActionButton.right = gTTabCtl.rcActionButton.left + pWindow->ScaleY( ActionButtonWidth )

    ' calculate the + add new file button. 
    ' Left position is calculated below once we know hpw many tabs will be drawn in the tab control.
    gTTabCtl.rcAddFileButton = gTTabCtl.rcActionButton 
    gTTabCtl.rcAddFileButton.left = 0

    for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs) 
        pDoc = gTTabCtl.tabs(i).pDoc
        if pDoc then
            ' Determine the length of the text
            gTTabCtl.tabs(i).wszText = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
            nTextWidth = getTextWidth( HWND_FRMMAIN_TOPTABS, gTTabCtl.tabs(i).wszText, ghMenuBar.hFontMenuBar, 0 )
            if nTextWidth < 70 then nTextWidth = 70

            ' calculate the various tab dimensions
            gTTabCtl.tabs(i).rcTab = rc
            gTTabCtl.tabs(i).rcTab.left = nLeft
            gTTabCtl.tabs(i).rcTab.right = nLeft + pWindow->ScaleX(rcIconWidth + nTextWidth + _
               hmargin + rcCloseWidth + hmargin + rightBorderWidth)
            
            gTTabCtl.tabs(i).rcIcon = gTTabCtl.tabs(i).rcTab
            gTTabCtl.tabs(i).rcIcon.top = gTTabCtl.tabs(i).rcIcon.top + pWindow->ScaleY(vmargin)
            gTTabCtl.tabs(i).rcIcon.bottom = gTTabCtl.tabs(i).rcIcon.top + pWindow->ScaleY(rcCloseWidth)
            gTTabCtl.tabs(i).rcIcon.left = gTTabCtl.tabs(i).rcTab.left 
            gTTabCtl.tabs(i).rcIcon.right = gTTabCtl.tabs(i).rcIcon.left + pWindow->ScaleX(rcIconWidth)
            
            gTTabCtl.tabs(i).rcText = gTTabCtl.tabs(i).rcTab
            gTTabCtl.tabs(i).rcText.left = gTTabCtl.tabs(i).rcIcon.right
            gTTabCtl.tabs(i).rcText.right = gTTabCtl.tabs(i).rcText.left + pWindow->ScaleX(nTextWidth)
            
            gTTabCtl.tabs(i).rcClose = gTTabCtl.tabs(i).rcTab
            gTTabCtl.tabs(i).rcClose.top = gTTabCtl.tabs(i).rcClose.top + pWindow->ScaleY(vmargin)
            gTTabCtl.tabs(i).rcClose.bottom = gTTabCtl.tabs(i).rcClose.top + pWindow->ScaleY(rcCloseWidth)
            gTTabCtl.tabs(i).rcClose.left = gTTabCtl.tabs(i).rcText.right + hmargin + hmargin
            gTTabCtl.tabs(i).rcClose.right = gTTabCtl.tabs(i).rcClose.left + pWindow->ScaleX(rcCloseWidth)

            gTTabCtl.tabs(i).pDoc = pDoc
            gTTabCtl.tabs(i).isHot = false
            nLeft = gTTabCtl.tabs(i).rcTab.right
            
        end if
    next
    
    ' The + Add New file button
    gTTabCtl.rcAddFileButton.left = nLeft + pWindow->ScaleX( hmargin )
    ' Ensure that the button stays within the tab area and always gets drawn whether
    ' it is immediately after the last tab or at the very end next top the right panel.
    dim as long MaxLeftAddFileButton = gTTabCtl.rcActionPanel.left - pWindow->ScaleY( AddFileButtonWidth )
    gTTabCtl.rcAddFileButton.left = min(gTTabCtl.rcAddFileButton.left, MaxLeftAddFileButton)
    gTTabCtl.rcAddFileButton.right = gTTabCtl.rcAddFileButton.left + pWindow->ScaleY( AddFileButtonWidth )
    

    ' We now know the exact width and postions of all tabs so now we need to calculate
    ' and adjust the gTTabCtl.FirstDisplayTab based on where the gTTabCtl.CurSel is located.
    if gTTabCtl.GetItemCount = 0 then
        ShowWindow( HWND_FRMMAIN_TOPTABS_SHADOW, SW_HIDE )
    else
        ' Adjust all of the Rects based on the current starting tab position
        dim as long nAdjAmount = gTTabCtl.tabs(gTTabCtl.FirstDisplayTab).rcTab.left
        frmTopTabs_AdjustTabRects( -(nAdjAmount) )
        
        gTTabCtl.ClientRightEdge = gTTabCtl.rcActionPanel.left
        
        if gTTabCtl.CurSel = -1 then
            ' Put if check for CurSel being invalid otherwise other tab calculations
            ' Scenarios will fail via GPF on invalid array access.
            
        ' SCENARIO #1
        ' New gTTabCtl.CurSel is located before gTTabCtl.FirstDisplayTab so it is visually
        ' not on screen. We simply need to move the gTTabCtl.CurSel rect into the
        ' the first position (gTTabCtl.FirstDisplayTab) and adjust all rects accordingly.
        elseif gTTabCtl.CurSel < gTTabCtl.FirstDisplayTab then
            ' The amount of the adjustment is equal to the difference between the
            ' current first tab's left edge and the new cursel's left edge.
            nAdjAmount = gTTabCtl.tabs(gTTabCtl.FirstDisplayTab).rcTab.left - gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.left
            frmTopTabs_AdjustTabRects( nAdjAmount )
            gTTabCtl.FirstDisplayTab = gTTabCtl.CurSel
        
        ' SCENARIO #2
        ' The gTTabCtl.CurSel is already located within view but the right client edge
        ' for the tabcontrol may have grown wider so if the last tab's (ubound(gTTabCtl.tabs)
        ' right edge falls before the right edge then there may be an opportunity to shift 
        ' all of the tabs right to help fill the space.
        elseif gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.right < gTTabCtl.ClientRightEdge then
            if gTTabCtl.tabs(ubound(gTTabCtl.tabs)).rcTab.right < gTTabCtl.ClientRightEdge then
            nAdjAmount = 0
                if gTTabCtl.FirstDisplayTab - 1 >= 0 then
                    dim as long nTabWidth = _
                            gTTabCtl.tabs(gTTabCtl.FirstDisplayTab - 1).rcTab.right - gTTabCtl.tabs(gTTabCtl.FirstDisplayTab - 1).rcTab.left
                    if gTTabCtl.tabs(ubound(gTTabCtl.tabs)).rcTab.right + nTabWidth <= gTTabCtl.ClientRightEdge then
                        frmTopTabs_AdjustTabRects( nTabWidth )
                        gTTabCtl.FirstDisplayTab = gTTabCtl.FirstDisplayTab - 1
                    end if   
                end if
            end if
        
        ' SCENARIO #3
        ' gTTabCtl.CurSel is located past the tab control's right edge. We can begin
        ' "removing" tabs at the start of the view until the gTTabCtl.CurSel tab comes
        ' completely within view.
        elseif gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.right > gTTabCtl.ClientRightEdge then
            ' need to determine how far to move the gTTabCtl.FirstDisplayTab
            nAdjAmount = 0
            for i as long = gTTabCtl.FirstDisplayTab to gTTabCtl.CurSel 
                ' remove the width of each successive tab starting with the
                ' first tab and work our way until we reach the target tab.
                dim as long nTabWidth = gTTabCtl.tabs(i).rcTab.right - gTTabCtl.tabs(i).rcTab.left
                nAdjAmount += nTabWidth
                if gTTabCtl.tabs(gTTabCtl.CurSel).rcTab.right - nAdjAmount <= gTTabCtl.ClientRightEdge then
                    gTTabCtl.FirstDisplayTab = i + 1
                    frmTopTabs_AdjustTabRects( -(nAdjAmount) )
                    exit for
                end if   
            next
        end if

        ' sanity checks
        if gTTabCtl.FirstDisplayTab > ubound(gTTabCtl.tabs) then 
            gTTabCtl.FirstDisplayTab = ubound(gTTabCtl.tabs)
        elseif gTTabCtl.FirstDisplayTab < 0 then 
            gTTabCtl.FirstDisplayTab = 0 
        end if   
    
    end if

    AfxRedrawWindow( HWND_FRMMAIN_TOPTABS )

    function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmTopTabs
' ========================================================================================
function frmTopTabs_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT

    if state <> SIZE_MINIMIZED then
        frmTopTabs_PositionWindows()
    end If

    function = 0
end function


' ========================================================================================
' Do hit test to determine what tab is currently under the mouse cursor
' ========================================================================================
function getHotTabHitTest( byval hWin as HWND ) as long
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    dim as long hotTab = -1
    for i as long = lbound(gTTabCtl.tabs) to ubound(gTTabCtl.tabs)
        if PtInRect( @gTTabCtl.tabs(i).rcTab, pt ) then
            hotTab = i
            gTTabCtl.tabs(i).isHot = true
        else   
            gTTabCtl.tabs(i).isHot = false
        end if
    next
    function = hotTab
end function

' ========================================================================================
' Do hit test to determine if "X" close button on tab was clicked
' ========================================================================================
function isTabCloseHitTest( byval hWin as HWND, byval idx as long ) as long
    if gTTabCtl.IsSafeIndex(idx) = false then exit function
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.tabs(idx).rcClose, pt ) then return idx
    function = -1
end function

' ========================================================================================
' Do hit test to determine if "+" add new file button was clicked
' ========================================================================================
function isAddFileButtonHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcAddFileButton, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test to determine if "..." button in action Area was clicked
' ========================================================================================
function isActionButtonHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcActionButton, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test to determine if Prev Tabs button in action Area was clicked
' ========================================================================================
function isPrevTabsHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcPrevTabs, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test to determine if Next Tabs button in action Area was clicked
' ========================================================================================
function isNextTabsHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcNextTabs, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test to determine if SplitEditor button in action Area was clicked
' ========================================================================================
function isSplitEditorHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcSplitEditor, pt ) then return true
    function = false
end function

' ========================================================================================
' Do hit test if over the Action Panel (no action but prevents other tooltips)
' ========================================================================================
function isActionPanelHitTest( byval hWin as HWND ) as boolean
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    if PtInRect( @gTTabCtl.rcActionPanel, pt ) then return true
    function = false
end function

' ========================================================================================
' frmTopTabsShadow_WndProc Window procedure
' ========================================================================================
function frmTopTabsShadow_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    select Case uMsg
        case WM_DESTROY
            dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
            if pWindow = 0 then delete pWindow
    end select      

    ' for messages that we don't deal with
    function = DefWindowProc(hwnd, uMsg, wParam, lParam)

end function

' ========================================================================================
' frmTopTabs Window procedure
' ========================================================================================
function frmTopTabs_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    static as long curr_idxHot = -1
    static hTooltip as HWND

    
    select Case uMsg
        HANDLE_MSG (hwnd, WM_SIZE, frmTopTabs_OnSize)
    
    case WM_DESTROY
        dim pWindow as CWindow ptr = AfxCWindowPtr(hwnd)
        if pWindow = 0 then delete pWindow
        DestroyWindow( HWND_FRMMAIN_TOPTABS_SHADOW )
        
    case WM_ERASEBKGND
        return true

    case WM_MOUSEMOVE
        dim tme as TrackMouseEvent
        tme.cbSize = sizeof(TrackMouseEvent)
        tme.dwFlags = TME_HOVER or TME_LEAVE
        tme.hwndTrack = hwnd
        TrackMouseEvent(@tme) 

        if IsWindow(hTooltip) = 0 then 
            hTooltip = AfxAddTooltip( hwnd, "", false, false )
            ' Set the tooltip max width in order to handle the multiline
            ' text needed for the split editor button.
            SendMessage( hTooltip, TTM_SETMAXTIPWIDTH, 0, 200 )
        end if
        
        curr_idxHot = getHotTabHitTest( hwnd )
        AfxRedrawWindow( hwnd )
        
        if gApp.bDragTabActive then
            if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
            ' Only activate tab drag if we have moved delta 5px xpos from original position
            dim as point pt = (GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
            if abs(pt.x - gApp.ptDragTabPrev.x) >= 5 then
                SetCursor( LoadCursor(0, MAKEINTRESOURCE(OCR_SIZEALL)) )
                if gTTabCtl.CurSel <> curr_idxHot then
                    ' Swap the two elements in the array
                    swap gTTabCtl.tabs(gTTabCtl.CurSel), gTTabCtl.tabs(curr_idxHot)
                    gTTabCtl.SetFocusTab( curr_idxHot )
                    frmTopTabs_PositionWindows()
                end if
            end if    
        end if
        
        
    case WM_MOUSELEAVE
        ' this will reset all tabs isHot to -1 and curr_idxHot to -1
        curr_idxHot = getHotTabHitTest( hwnd )
        AfxDeleteTooltip( hTooltip, hwnd )
        hTooltip = 0
        AfxRedrawWindow( hwnd )


    case WM_MOUSEHOVER
        dim as CWSTR wszTooltip 
        ' test for Action Button, Filename caption or Close caption
        curr_idxHot = getHotTabHitTest( hwnd )
        if isPrevTabsHitTest( hwnd ) then
            wszTooltip  = L(431, "Scroll Tabs Left")
        elseif isNextTabsHitTest( hwnd ) then
            wszTooltip  = L(432, "Scroll Tabs Right")
        elseif isActionButtonHitTest( hwnd ) then
            wszTooltip  = L(430, "Tab List")
        elseif isAddFileButtonHitTest( hwnd ) then
            wszTooltip  = L(003, "New File")
        elseif isSplitEditorHitTest( hwnd ) then
            ' If CTRL is pressed then use the top/bottom split, otherwise
            ' use the left/right split icon.
            if (wszIconSplitEditor = wszIconSplitTopBottom) then
                wszTooltip  = L(070,"Toggle Split Editor Top/Bottom")
            else    
                wszTooltip  = L(069,"Toggle Split Editor Left/Right") & _
                                    chr(13,10) & "[Alt]: " & L(071,"Split Top/Bottom")
            end if

        elseif isActionPanelHitTest( hwnd ) then
            ' do nothing. no tooltip.
        elseif isTabCloseHitTest( hwnd, curr_idxHot ) = curr_idxHot then
            wszTooltip  = L(86, "Close Tab")
        else
            if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
            wszTooltip = gTTabCtl.tabs(curr_idxHot).pDoc->DiskFilename   
        end if
        ' Display the tooltip
        AfxSetTooltipText( hTooltip, hwnd, wszTooltip )
        AfxRedrawWindow( hwnd )


    case WM_RBUTTONDOWN
        if isPrevTabsHitTest( hwnd ) then
        elseif isNextTabsHitTest( hwnd ) then
        elseif isActionButtonHitTest( hwnd ) then
        elseif isActionPanelHitTest( hwnd ) then
        elseif isAddFileButtonHitTest( hwnd ) then
        elseif isSplitEditorHitTest( hwnd ) then
        else
            ' Create the popup menu
            curr_idxHot = getHotTabHitTest( hwnd )
            if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
            gTTabCtl.SetFocusTab(curr_idxHot)
            dim as point pt
            dim as HMENU hPopupMenu = CreateTopTabCtlContextMenu(curr_idxHot)
            GetCursorPos @pt
            TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, byval Null)
            DestroyMenu hPopUpMenu
            Return True   ' prevent further processing that leads to WM_CONTEXTMENU
        end if

    case WM_LBUTTONDOWN
        if isPrevTabsHitTest( hwnd ) then
        elseif isNextTabsHitTest( hwnd ) then
        elseif isActionButtonHitTest( hwnd ) then
        elseif isActionPanelHitTest( hwnd ) then
        elseif isAddFileButtonHitTest( hwnd ) then
        elseif isSplitEditorHitTest( hwnd ) then
        else ' regular tab
            curr_idxHot = getHotTabHitTest( hwnd )
            if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
            if (isTabCloseHitTest( hwnd, curr_idxHot ) = curr_idxHot) andalso (curr_idxHot <> gTTabCtl.CurSel) then
                ' we are mouse down on an "X" close on a tab that is not currently the active tab. We
                ' will simply close that tab rather than bring it into focus and then delete.
                ' allow the mouse message to eventually bubble up WM_LBUTTONUP where we do the close.
            else
                ' we have clicked on tab so bring it into focus
                gTTabCtl.SetFocusTab( curr_idxHot )
                dim as point pt = (GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))
                gApp.bDragTabActive = true
                gApp.ptDragTabPrev = pt
                AfxRedrawWindow( hwnd )
            end if
            SetCapture(hwnd)
        end if

    case WM_LBUTTONUP
        if isPrevTabsHitTest( hwnd ) then
            ' if the first displayed tab is not index 0 then there must be tabs before it
            ' that can be displayed. if no previous tabs then set focus to the first tab.
            if gTTabCtl.FirstDisplayTab > 0 then
                gTTabCtl.SetFocusTab( gTTabCtl.FirstDisplayTab - 1 )
            else   
                gTTabCtl.SetFocusTab(0)
            end if
        elseif isNextTabsHitTest( hwnd ) then
            ' determine if there are any tabs to the right of the current display view. Bring
            ' the first one found (if any) into view. Keep track of the found tab because if
            ' none found then simply set the CurSel to the last tab.
            dim as long nFound = -1
            for i as long = gTTabCtl.FirstDisplayTab to gTTabCtl.GetItemCount - 1
                if gTTabCtl.tabs(i).rcTab.left > gTTabCtl.ClientRightEdge then
                    nFound = i: exit for
                    exit for
                end if
            next
            if nFound = -1 then nFound = gTTabCtl.GetItemCount - 1
            gTTabCtl.SetFocusTab(nFound)

        elseif isActionButtonHitTest( hwnd ) then
            dim as HMENU hPopupMenu = CreateTopTabsActionButtonContextMenu()
            ' Popup the menu to the bottom of the Action Button (right aligned)
            dim as RECT rc = gTTabCtl.rcActionButton   ' work with a copy
            MapWindowpoints( HWND_FRMMAIN_TOPTABS, HWND_DESKTOP, cast(point ptr, @rc), 2 )
            dim as long id = TrackPopupMenu(hPopUpMenu, _
                                TPM_RIGHTALIGN or TPM_RETURNCMD, _
                                rc.right, rc.bottom, 0, HWND_FRMMAIN, byval null)
            ' Return value is 1 based because 0 indicates that menu was cancelled
            if (id > 0) andalso (id <= gTTabCtl.GetItemCount) then
                gTTabCtl.SetFocusTab(id-1)
            end if
            DestroyMenu( hPopUpMenu )
            Return true   ' prevent further processing that leads to WM_CONTEXTMENU

        elseif isAddFileButtonHitTest( hwnd ) then
            ' Add a new file tab
            PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_FILENEWLASTTAB, 0), 0 )

        elseif isSplitEditorHitTest( hwnd ) then
            ' Split editor view
            ' IDM_SPLITLEFTRIGHT, IDM_SPLITTOPBOTTOM
            if (wszIconSplitEditor = wszIconSplitTopBottom) then
                PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_SPLITTOPBOTTOM, 0), 0 )
            else    
                PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(IDM_SPLITLEFTRIGHT, 0), 0 )
            end if
            
        elseif isActionPanelHitTest( hwnd ) then
        else
            ' Reset the mouse pointer
            SetCursor( LoadCursor( null, IDC_ARROW ))
            ReleaseCapture
            gApp.bDragTabActive = false
            ' if we are still over an "X" close during buttonup then close the tab
            curr_idxHot = getHotTabHitTest( hwnd )
            if gTTabCtl.IsSafeIndex(curr_idxHot) = false then exit function
            if isTabCloseHitTest( hwnd, curr_idxHot ) = curr_idxHot then
                gTTabCtl.CloseTab( curr_idxHot )
                frmTopTabs_PositionWindows()
            end if
        end if

    case WM_PAINT
        dim as PAINTSTRUCT ps
        dim as HDC hDc

        dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN_TOPTABS)
        if pWindow = 0 then exit function

        hDC = BeginPaint( hwnd, @ps )

        SaveDC(hDC)
        dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
        dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

        dim memDC as HDC      ' Double buffering
        dim hbit as HBITMAP   ' Double buffering

        memDC = CreateCompatibleDC( hDC )
        hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
        if hbit then hbit = SelectObject( memDC, hbit )

        ' Fill in the entire back panel width across the top of the screen
        FillRect( memDC, @ps.rcPaint, ghTopTabs.hPanelBrush )

        ' Create a black pen that acts as the divider for each tab
        dim as long penWidth = pWindow->ScaleX(1)
        dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghTopTabs.Divider )
        SelectObject( memDC, hPenSolid)
        
        dim as HPEN hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen

        dim as point pt: GetCursorPos( @pt )
        dim as RECT rc
        dim as long wsStyle
        
        ' All of the rc calculations have already been done in frmTopTabs_PostionWindows
        if gTTabCtl.FirstDisplayTab > ubound(gTTabCtl.tabs) then gTTabCtl.FirstDisplayTab = lbound(gTTabCtl.tabs)
        for i as long = gTTabCtl.FirstDisplayTab to ubound(gTTabCtl.tabs)
            ' paint this tab based on active/inactive status
            if i = gTTabCtl.CurSel then
                SetBkColor( memDC, ghTopTabs.BackColorHot )
                SetTextColor( memDC, ghTopTabs.ForeColorHot )
                FillRect( memDC, @gTTabCtl.tabs(i).rcTab, ghTopTabs.hBackBrushHot )
            else
                SetBkColor( memDC, ghTopTabs.BackColor )
                SetTextColor( memDC, ghTopTabs.ForeColor )
                FillRect( memDC, @gTTabCtl.tabs(i).rcTab, ghTopTabs.hBackBrush )
            end if
            
            ' display the tab text
            wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            SelectObject(memDC, ghMenuBar.hFontMenuBar)
            DrawText( memDC, gTTabCtl.tabs(i).wszText.sptr, -1, cast(lpRect, @gTTabCtl.tabs(i).rcText), wsStyle )
            
            ' if this document is dirty then the dirty circle icon overrides the display of the close icon
            ' unless the tab isHot. The dirty icon displays always for active and inactive tabs.
            dim as boolean isDirty 
            if cbool(SciExec( gTTabCtl.tabs(i).pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) then isDirty = true
            if gTTabCtl.tabs(i).pDoc->UserModified then isDirty = true
            
            if isDirty then
                ' if mouse is over the rcClose then draw the close icon, otherwise draw the dirty circle
                SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
                DrawText( memDC, wszIconCompileResult, -1, cast(lpRect, @gTTabCtl.tabs(i).rcClose), wsStyle )
            end if
            
            if (i = gTTabCtl.CurSel) or (gTTabCtl.tabs(i).isHot = true) then
                ' if this is the active tab or it is mouse hot then display the close icon only 
                ' if the dirty icon is not already being displayed
                rc = gTTabCtl.tabs(i).rcClose
                if isDirty = false then
                    SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
                    DrawText( memDC, wszIconClose, -1, cast(lpRect, @rc), wsStyle )
                end if
                if (isTabCloseHitTest( hwnd, i ) = i) andalso (gApp.bDragTabActive = false) then
                    ' if we are hovered over the "X" close icon rect then highlight it
                    selectPen( memDC, hPenNull )
                    SelectObject( memDC, ghTopTabs.hCloseBrushHot )
                    RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
                    SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
                    SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
                    DrawText( memDC, wszIconClose, -1, cast(lpRect, @rc), wsStyle )
                end if   
            end if
            
            ' Draw the righthand side black divider
            selectPen( memDC, hPenSolid )
            MoveToEx( memDC, gTTabCtl.tabs(i).rcTab.right - penWidth, gTTabCtl.tabs(i).rcTab.top, null )
            LineTo( memDC, gTTabCtl.tabs(i).rcTab.right - penWidth, gTTabCtl.tabs(i).rcTab.bottom )   
            
            ' Do the calculation of the LastDisplayTab here because we need to
            ' do testing of when the tab rect goes beyond the visual area.
            if gTTabCtl.tabs(i).rcTab.right < gTTabCtl.rcActionPanel.left then
                gTTabCtl.LastDisplayTab = i
            end if

        next
        
        
        ' Draw the + new file button
        SelectObject( memDC, ghTopTabs.hPanelBrush )
        SetBkColor( memDC, ghTopTabs.BackColor )
        SetTextColor( memDC, ghTopTabs.ForeColor )
        if isAddFileButtonHitTest( hwnd ) then
            SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            SelectObject( memDC, ghTopTabs.hCloseBrushHot )
        end if
        rc = gTTabCtl.rcAddFileButton
        FillRect( memDC, @rc, ghTopTabs.hPanelBrush )
        selectPen( memDC, hPenNull )
        RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
        wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
        SelectObject(memDC, ghMenuBar.hFontSymbol)
        DrawText( memDC, wszIconAddFileButton, -1, cast(lpRect, @rc), wsStyle )
        
        
        ' Lastly, draw the TopTabs Action Area (basically, the little panel to the
        ' right of all of the top tabs that has the "..." icon.
        if gTTabCtl.GetItemCount then
            ' paint the full Action Panel
            FillRect( memDC, @gTTabCtl.rcActionPanel, ghTopTabs.hPanelBrush )
            
            ' Draw the PrevTabs button
            SelectObject( memDC, ghTopTabs.hPanelBrush )
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            if isPrevTabsHitTest( hwnd ) then
                SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
                SelectObject( memDC, ghTopTabs.hCloseBrushHot )
            end if
            rc = gTTabCtl.rcPrevTabs
            selectPen( memDC, hPenNull )
            RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
            wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            SelectObject(memDC, ghMenuBar.hFontSymbol)
            DrawText( memDC, wszIconChevronLeft, -1, cast(lpRect, @rc), wsStyle )

            ' Draw the NextTabs button
            SelectObject( memDC, ghTopTabs.hPanelBrush )
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            if isNextTabsHitTest( hwnd ) then
                SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
                SelectObject( memDC, ghTopTabs.hCloseBrushHot )
            end if
            rc = gTTabCtl.rcNextTabs
            selectPen( memDC, hPenNull )
            RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
            wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            SelectObject(memDC, ghMenuBar.hFontSymbol)
            DrawText( memDC, wszIconChevronRight, -1, cast(lpRect, @rc), wsStyle )

            ' Draw the SplitEditor button
            SelectObject( memDC, ghTopTabs.hPanelBrush )
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            if isSplitEditorHitTest( hwnd ) then
                SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
                SelectObject( memDC, ghTopTabs.hCloseBrushHot )
            end if
            rc = gTTabCtl.rcSplitEditor
            selectPen( memDC, hPenNull )
            RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
            wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            SelectObject(memDC, ghMenuBar.hFontSymbol)
            DrawText( memDC, wszIconSplitEditor, -1, cast(lpRect, @rc), wsStyle )
            
            ' Draw the Action buttom "..."
            SelectObject( memDC, ghTopTabs.hPanelBrush )
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            if isActionButtonHitTest( hwnd ) then
                SetBkColor( memDC, ghTopTabs.CloseBackColorHot )
                SelectObject( memDC, ghTopTabs.hCloseBrushHot )
            end if
            rc = gTTabCtl.rcActionButton
            selectPen( memDC, hPenNull )
            RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
            wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            SelectObject(memDC, ghMenuBar.hFontSymbolLarge)
            DrawText( memDC, wszIconMoreActions, -1, cast(lpRect, @rc), wsStyle )
        end if
    
        BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

        ' Cleanup
        if hPenSolid then deleteObject( hPenSolid )
        if hPenNull  then deleteObject( hPenNull )
        if hbit  then deleteObject( SelectObject(memDC, hbit) )
        if memDC then deleteDC( memDC )
        RestoreDC( hDC, -1 )

        EndPaint( hwnd, @ps )
    
    end select
    
    ' for messages that we don't deal with
    function = DefWindowProc(hwnd, uMsg, wParam, lParam)

end function



' ========================================================================================
' frmTopTabs_Show
' ========================================================================================
function frmTopTabs_Show( byval hwndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow

    HWND_FRMMAIN_TOPTABS = pWindow->Create( hwndParent, "", @frmTopTabs_WndProc, _
        0, 0, 0, TOPTABS_HEIGHT, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)

    ' Disable background erasing by only assigning the one style
    pWindow->ClassStyle = CS_DBLCLKS

    ' create semi-transparent window offset under our tabcontrol in order to simulate a shadow.
    pWindow = New CWindow
    HWND_FRMMAIN_TOPTABS_SHADOW = pWindow->Create( HWND_FRMMAIN, "", _
        @frmTopTabsShadow_WndProc, 0, 0, 0, pWindow->ScaleY(1), _
        WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, WS_EX_LAYERED or WS_EX_NOACTIVATE )
    pWindow->ClassStyle = CS_DBLCLKS
    pWindow->Brush = GetSysColorBrush(COLOR_WINDOWTEXT + 1)   ' black background
    ' 0 totally transparent, 255 totally opaque
    SetLayeredWindowAttributes( HWND_FRMMAIN_TOPTABS_SHADOW, GetSysColor(COLOR_WINDOWTEXT + 1) , 100, LWA_ALPHA ) 
    
    frmTopTabs_PositionWindows()
    function = 0
    
end function

