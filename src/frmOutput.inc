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

#include once "frmOutput.bi"


' ========================================================================================
' Clear data from all of the controls in the frmOutput windows. This is needed
' when Projects are loaded and closed.
' ========================================================================================
function frmOutput_ResetAllControls() as long 
    ListView_DeleteAllItems( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVRESULTS) )
    AfxSetWindowText( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), "" )
    ListView_DeleteAllItems( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVSEARCH) )
    ListView_DeleteAllItems( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO) )
    AfxSetWindowText( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), "" )
    function = 0
end function


' ========================================================================================
' Ensure that the correct notes are shown
' ========================================================================================
function frmOutput_ShowNotes() as long 

    dim HCTL as HWND = GETDLGITEM(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)

    if gApp.IsProjectActive then
        AfxSetWindowText(hCtl, gApp.ProjectNotes)
    else   
        AfxSetWindowText(hCtl, gApp.NonProjectNotes)
    end if

    function = 0
end function


' ========================================================================================
' Update the TODO listview
' ========================================================================================
function frmOutput_UpdateToDoListview() as long 
    dim as hwnd hLV = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO)
    ListView_DeleteAllItems( hLV )
    dim as long n = 0

    dim pData as DB2_DATA ptr
    
    gdb2.dbRewind()
    do 
        pData = gdb2.dbGetNext
        if pData = 0 then exit do
        if pData->id <> DB2_TODO then continue do
        FF_ListView_InsertItem( hLV, n, 0, "" ) 
        FF_ListView_InsertItem( hLV, n, 1, ltrim(wstr(pData->nLineStart))) 
        FF_ListView_InsertItem( hLV, n, 2, ltrim(wstr(pData->fileName)))
        FF_ListView_InsertItem( hLV, n, 3, ltrim(wstr(pData->ElementData))) 
        n = n + 1
    loop

    function = 0
end function

' ========================================================================================
' Update the SEARCH listview
' ========================================================================================
function frmOutput_UpdateSearchListview( byref wszResultFile as wstring ) as long 
    dim hLV as HWND = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVSEARCH )

    ListView_DeleteAllItems( hLV )
    dim as long n = 0

    if AfxFileExists(wszResultFile) = 0 then exit function
    
    dim as CWSTR wst

    dim pStream as CTextStream
    if pStream.OpenUnicode(wszResultFile) <> S_OK then exit function
    
    do until pStream.EOS
        wst = pStream.ReadLine
        
        wst = trim(wst)
        if len(wst) = 0 then continue do

        dim as CWSTR wszFilename, wszLineNum, wszDescription
        dim as long f1, f2
        
        ' Search for the 2nd semicolon
        f1 = instr( 3, wst, ":" )
        if f1 then f2 = instr( f1 + 1, wst, ":" )

        if f1 then wszFilename = left(wst, f1 - 1)
        if f2 > f1 then wszLinenum = mid(wst, f1 + 1, f2 - f1 - 1)
        if f2 then wszDescription = rtrim(mid(wst, f2 + 1))

        FF_ListView_InsertItem( hLV, n, 0, "" ) 
        FF_ListView_InsertItem( hLV, n, 1, wszLineNum ) 
        FF_ListView_InsertItem( hLV, n, 2, wszFilename )
        FF_ListView_InsertItem( hLV, n, 3, wszDescription ) 
        n = n + 1

    loop
    pStream.Close

    AfxDeleteFile( wszResultFile )
    
    ' Show the search results
    gOutputTabsCurSel = 2 
    ShowWindow( HWND_FRMOUTPUT, SW_SHOW )
    frmMain_PositionWindows
    frmOutput_PositionWindows

    function = 0
end function
            

' ========================================================================================
' Show/Hide correct child controls
' ========================================================================================
function frmOutput_ShowHideOutputControls( byval hwnd as HWND ) as LRESULT
    dim as HWND hCtrl
    dim as RECT rc
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMOUTPUT)
    if pWindow = 0 then exit function

    ' By default, hide all controls
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_TABS), SW_SHOW
    AfxRedrawWindow( GetDlgItem(hwnd, IDC_FRMOUTPUT_TABS) )
    
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_LVRESULTS), SW_HIDE
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_TXTLOGFILE), SW_HIDE
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_LVSEARCH), SW_HIDE
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_LVTODO), SW_HIDE
    ShowWindow GetDlgItem(hwnd, IDC_FRMOUTPUT_TXTNOTES), SW_HIDE
    
    select case gOutputTabsCurSel
        case 0    ' compiler results
            hCtrl = GetDlgItem(hwnd, IDC_FRMOUTPUT_LVRESULTS)
            ListView_SetColumnWidth( hCtrl, 3, LVSCW_AUTOSIZE_USEHEADER ) 

        case 1    ' compiler log file
            hCtrl = GetDlgItem(hwnd, IDC_FRMOUTPUT_TXTLOGFILE)
            GetClientRect( hCtrl, @rc )
            rc.left = rc.left + pWindow->ScaleX(20)
            SendMessage( hCtrl, EM_SETRECT, 0, cast(LPARAM, @rc) )

        case 2    ' search results
            hCtrl = GetDlgItem(hwnd, IDC_FRMOUTPUT_LVSEARCH)
            ListView_SetColumnWidth( hCtrl, 3, LVSCW_AUTOSIZE_USEHEADER ) 

        case 3    ' ToDo list
            ' ensure last column is sized to fit to end of client area.
            hCtrl = GetDlgItem(hwnd, IDC_FRMOUTPUT_LVTODO)
            ListView_SetColumnWidth( hCtrl, 3, LVSCW_AUTOSIZE_USEHEADER ) 

        case 4    ' Notes
            hCtrl = GetDlgItem(hwnd, IDC_FRMOUTPUT_TXTNOTES)
            GetClientRect( hCtrl, @rc )
            rc.left = rc.left + pWindow->ScaleX(20)
            SendMessage( hCtrl, EM_SETRECT, 0, cast(LPARAM, @rc) )
    
    end select
    
    SetWindowPos( hCtrl, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE or SWP_SHOWWINDOW )    
    SetFocus( hCtrl )
            
    function = 0
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmOutput_PositionWindows() as LRESULT
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMOUTPUT)
    if pWindow = 0 then exit function
    
    dim as long nTop  = pWindow->ScaleY(2)
    dim as long nLeft = 0 
    dim as long nTabsHeight
    dim as Rect rc: GetClientRect( HWND_FRMOUTPUT, @rc )
    
    dim as HWND hTabs = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABS)

    nTabsHeight = AfxGetWindowHeight(hTabs)
    SetWindowPos( hTabs, 0, nLeft, nTop, rc.right - rc.left, nTabsHeight, SWP_NOZORDER )

    ' Position the child controls
    nTop = nTop + nTabsHeight + pWindow->ScaleY(8)
    SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
    SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVRESULTS), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
    SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVSEARCH), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
    SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
    SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER

    ' Calculate the tabs rects
    GetClientRect( hTabs, @rc )

    dim wszText as wstring * 100
    dim as long nTextLen 
    dim as HFONT hFontText = ghStatusBar.hFontStatusBar
    dim as long hmargin = pWindow->ScaleX(10) 
    rc.top = rc.top + pWindow->ScaleY(2)
    rc.left = rc.left + pWindow->ScaleX(10)

    wszText = ucase(L(191, "Compiler Results"))
    nTextLen = pWindow->ScaleX(getTextWidth( HWND_FRMOUTPUT, wszText, hFontText, 0 ))
    gOutputTabs(0).wszText = wszText
    gOutputTabs(0).rcTab = rc: gOutputTabs(0).rcText = rc
    gOutputTabs(0).rcText.left = gOutputTabs(0).rcTab.left + hmargin 
    gOutputTabs(0).rcText.right = gOutputTabs(0).rcText.left + nTextLen 
    gOutputTabs(0).rcTab.right = gOutputTabs(0).rcText.right + hmargin

    wszText = ucase(L(252, "Compiler Log File"))
    nTextLen = pWindow->ScaleX(getTextWidth( HWND_FRMOUTPUT, wszText, hFontText, 0 ))
    gOutputTabs(1).wszText = wszText
    gOutputTabs(1).rcTab = rc: gOutputTabs(1).rcText = rc
    gOutputTabs(1).rcTab.left = gOutputTabs(0).rcTab.right
    gOutputTabs(1).rcText.left = gOutputTabs(1).rcTab.left + hmargin 
    gOutputTabs(1).rcText.right = gOutputTabs(1).rcText.left + nTextLen 
    gOutputTabs(1).rcTab.right = gOutputTabs(1).rcText.right + hmargin

    wszText = ucase(L(262, "Search Results"))
    nTextLen = pWindow->ScaleX(getTextWidth( HWND_FRMOUTPUT, wszText, hFontText, 0 ))
    gOutputTabs(2).wszText = wszText
    gOutputTabs(2).rcTab = rc: gOutputTabs(2).rcText = rc
    gOutputTabs(2).rcTab.left = gOutputTabs(1).rcTab.right
    gOutputTabs(2).rcText.left = gOutputTabs(2).rcTab.left + hmargin 
    gOutputTabs(2).rcText.right = gOutputTabs(2).rcText.left + nTextLen 
    gOutputTabs(2).rcTab.right = gOutputTabs(2).rcText.right + hmargin

    wszText = ucase(L(263, "TODO"))
    nTextLen = pWindow->ScaleX(getTextWidth( HWND_FRMOUTPUT, wszText, hFontText, 0 ))
    gOutputTabs(3).wszText = wszText
    gOutputTabs(3).rcTab = rc: gOutputTabs(3).rcText = rc
    gOutputTabs(3).rcTab.left = gOutputTabs(2).rcTab.right
    gOutputTabs(3).rcText.left = gOutputTabs(3).rcTab.left + hmargin 
    gOutputTabs(3).rcText.right = gOutputTabs(3).rcText.left + nTextLen 
    gOutputTabs(3).rcTab.right = gOutputTabs(3).rcText.right + hmargin

    wszText = ucase(L(264, "Notes"))
    nTextLen = pWindow->ScaleX(getTextWidth( HWND_FRMOUTPUT, wszText, hFontText, 0 ))
    gOutputTabs(4).wszText = wszText
    gOutputTabs(4).rcTab = rc: gOutputTabs(4).rcText = rc
    gOutputTabs(4).rcTab.left = gOutputTabs(3).rcTab.right
    gOutputTabs(4).rcText.left = gOutputTabs(4).rcTab.left + hmargin 
    gOutputTabs(4).rcText.right = gOutputTabs(4).rcText.left + nTextLen 
    gOutputTabs(4).rcTab.right = gOutputTabs(4).rcText.right + hmargin

    dim as long rcCloseWidth = 20
    dim as long rcCloseHeight = 20
    dim as long vmargin = pWindow->ScaleY( (OUTPUT_TABS_HEIGHT - rcCloseHeight) / 2 )
    GetClientRect( hTabs, @rc )
    gOutputCloseRect.top = gOutputTabs(4).rcTab.top + vmargin
    gOutputCloseRect.bottom = gOutputTabs(4).rcTab.bottom - vmargin
    gOutputCloseRect.right = rc.right - hmargin 
    gOutputCloseRect.left = gOutputCloseRect.right - pWindow->ScaleX(rcCloseWidth)

    ' Determine which child controls should be shown or hidden
    frmOutput_ShowHideOutputControls(HWND_FRMOUTPUT)
    
    function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmOutput
' ========================================================================================
function frmOutput_OnSize( _
            byval hwnd as HWND, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT

    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmOutput_PositionWindows
    end if

    function = 0
end function
    

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOutput
' ========================================================================================
function frmOutput_OnCommand( _
            byval hwnd as HWND, _
            byval id as long, _
            byval hwndCtl as HWND, _
            byval codeNotify as UINT _
            ) as LRESULT

    select case codeNotify
        
        case EN_CHANGE
            ' Notes have been modified. Save them to the correct global variable
            ' to ensure that the changes are not lost when documents or projects
            ' are switched.
            if id = IDC_FRMOUTPUT_TXTNOTES then
                if gApp.IsProjectActive then
                    gApp.ProjectNotes = AfxGetWindowText(hwndCtl)
                else
                    gApp.NonProjectNotes = AfxGetWindowText(hwndCtl)
                end if   
                exit function
            end if
            
    end select
                        
    function = 0
end function
        
    
' ========================================================================================
' Processes messages for the subclassed frmOutput Compile Results and TODO listviews .
' ========================================================================================
function frmOutput_Listview_SubclassProc ( _
                    byval hwnd   as HWND, _                 ' // Control window handle
                    byval uMsg   as UINT, _                 ' // Type of message
                    byval wParam as WPARAM, _               ' // First message parameter
                    byval lParam as LPARAM, _               ' // Second message parameter
                    byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
                    byval dwRefData as DWORD_PTR _          ' // pointer to reference data
                    ) as LRESULT

    ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
    if (uMsg = WM_KEYUP) andalso (loword(wParam) = VK_RETURN) then uMsg = WM_LBUTTONDBLCLK
        
    select case uMsg

        case WM_GETDLGCODE
            ' All keyboard input
            function = DLGC_WANTALLKEYS
            exit function
            
        case WM_NOTIFY
            ' Handle custom draw of the listview header   
            dim ptnmhdr as NMHDR ptr             
            dim ptnmcd  as NMCUSTOMDRAW ptr    
                    
            ptnmhdr = cast(NMHDR ptr, lParam)

            ' need to prevent re-entry into HDN_ITEMCHANGING because
            ' calling ListView_SetColumnWidth triggers another HDN_ITEMCHANGING
            static as boolean inChanging = false
            
            if ptnmhdr->code = HDN_ITEMCHANGING then
                ' notification from the ListView header control that the user is
                ' resizing a header item via the mouse. Update the last column width
                ' to ensure it covers the non-client area (because difficult painting
                ' this area).
                if inChanging then exit function
                inChanging = true
                ListView_SetColumnWidth( hwnd, 3, LVSCW_AUTOSIZE_USEHEADER )
                inChanging = false
            end if
            
            if ptnmhdr->code = NM_CUSTOMDRAW then      
                ptnmcd = cast(NMCUSTOMDRAW ptr, lParam)
                            
                ' Determine the stage of the paint cycle
                select case ptnmcd->dwDrawStage
                        
                    case CDDS_PREPAINT
                        ' Control is to notify parent about each item being drawn
                        return CDRF_NOTIFYITEMDRAW 
                                
                    ' Items are being painted
                    case CDDS_ITEMPREPAINT     
                        select case ptnmcd->dwItemSpec
                            case 0, 1, 2, 3   ' columns
                            ' Paint the whole cell ourselves
                            FillRect( ptnmcd->hdc, @ptnmcd->rc, ghOutput.hPanelBrush )
                            SetTextColor( ptnmcd->hdc, ghOutput.ForeColorHot )
                            SetBkColor( ptnmcd->hdc, ghOutput.BackColor )
                            dim wszText as wstring * MAX_PATH
                            'Header_GetItemText is currently bugged. Report submitted to Jose to correct.
                            'Header_GetItemText( ptnmhdr->hwndFrom, ptnmcd->dwItemSpec, @wszText, MAX_PATH )
                            ListView_GetHeaderText( hwnd, ptnmcd->dwItemSpec, @wszText, MAX_PATH )
                            dim as HFONT oldFont
                            oldFont = selectObject( ptnmcd->hdc, ghStatusBar.hFontStatusBar )
                            dim as long wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE
                            DrawText( ptnmcd->hdc, wszText, -1, Cast(lpRect, @ptnmcd->rc), wsStyle )
                            selectObject( ptnmcd->hdc, oldFont )
                            return CDRF_SKIPDEFAULT
                        end select
                            
                end select
                        
                return true
            end if
        
        case WM_LBUTTONDBLCLK
            SetDocumentErrorPosition( hwnd, gCompile.CompileID )    
            exit function
        
        case WM_KEYUP
            select case Loword(wParam)
                case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
            End select
            exit function

        case WM_CHAR   ' prevent the annoying beep!
            if wParam = VK_RETURN then return 0
            if wParam = VK_ESCAPE then return 0

        case WM_DESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hwnd, @frmOutput_Listview_SubclassProc, uIdSubclass )

    End select

    ' Default processing of Windows messages
    function = DefSubclassProc( hwnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' Process WM_MOUSEMOVE message for window/dialog: frmOutput
' ========================================================================================
function frmOutput_OnMouseMove( _
            byval hwnd as HWND, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
    if pWindow = 0 then exit function

    ' HITTEST (PANELS SPLITTER)
    dim as point pt 
    dim as RECT rc
    GetWindowRect( HWND_FRMOUTPUT, @rc )
    rc.Bottom = rc.Top + pWindow->ScaleY(3)
    GetCursorPos(@pt)
    if PtInRect( @rc, pt ) then
        if WindowFrompoint(pt) = HWND_FRMOUTPUT then
            SetCursor( ghCursorSizeNS )
        end if
    end if
    
    if gApp.bDragActive then
        if gApp.hwndPanel = HWND_FRMOUTPUT then
            GetCursorPos(@pt)
            GetWindowRect( HWND_FRMOUTPUT, @rc )
            dim as long nHeight 
            dim as long nDiff = pt.y - rc.top
            ' Adjust the height. The positioning will be taken care of in PositionMainWindows().
            rc.top = rc.top + nDiff
    
            ' Don't move the Output pane if the top is less than the bottom of the TopTabs
            dim as RECT rc2
            dim as long nTopLimit 
            GetWindowRect( HWND_FRMMAIN_MENUBAR, @rc2 )
            nTopLimit = rc2.bottom
            if gTTabCtl.GetItemCount then 
                GetWindowRect( HWND_FRMMAIN_TOPTABS, @rc2 )
                nTopLimit = rc2.bottom
            end if
            rc.top = max(rc.top, nTopLimit)
            nHeight = (rc.bottom-rc.top) - pWindow->ScaleY(4) ' allow room to grab the top
            
            ' The minimum height of the Output window when visible is the height of the tabs
            nHeight = max( nHeight, pWindow->ScaleY(OUTPUT_TABS_HEIGHT) )

            SetWindowPos( HWND_FRMOUTPUT, 0, 0, 0, rc.Right - rc.Left, nHeight, SWP_NOMOVE or SWP_NOZORDER )
            frmMain_PositionWindows
            exit function
        end if
    end if

    function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmOutput
' ========================================================================================
function frmOutput_OnLButtonDown( _
            byval hwnd as HWND, _
            byval fDoubleClick as Boolean, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
    if pWindow = 0 then exit function

    ' HITTEST (PANELS TOP/BOTTOM SPLITTER)
    dim as Rect rc
    dim as point pt 

    gApp.bDragActive = False 
    
    GetWindowRect HWND_FRMOUTPUT, @rc
    rc.Bottom = rc.Top + pWindow->ScaleY(3)
    GetCursorPos(@pt)
    if PtInRect( @rc, pt ) then
        if WindowFrompoint(pt) = HWND_FRMOUTPUT then
            SetCursor( ghCursorSizeNS )
            gApp.bDragActive = True 
            gApp.hwndPanel   = HWND_FRMOUTPUT
            SetCapture( HWND_FRMOUTPUT )
        end if
        exit function
    end if

    function = 0
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmOutput
' ========================================================================================
function frmOutput_OnLButtonUp( _
            byval hwnd as HWND, _
            byval x as long, _
            byval y as long, _
            byval keyflags as UINT _
            ) as long

    ' HITTEST (PANELS TOP/BOTTOM SPLITTER)
    if gApp.bDragActive then
        gApp.bDragActive = False 
        gApp.hwndPanel = 0
        ReleaseCapture()
    end if
    SetCursor( LoadCursor( null, IDC_ARROW ))
    
    function = 0
end function


' ========================================================================================
' Do hit test to determine what tab is currently under the mouse cursor
' ========================================================================================
function frmOutputTabs_getHotTabHitTest( byval hWin as HWND ) as long
    dim as point pt: GetCursorPos( @pt )
    MapWindowpoints( HWND_DESKTOP, hWin, cast( point ptr, @pt ), 1 )
    dim as long hotTab = -1
    for i as long = lbound(gOutputTabs) to ubound(gOutputTabs)
        if PtInRect( @gOutputTabs(i).rcTab, pt ) then
            hotTab = i
            gOutputTabs(i).isHot = true
        else   
            gOutputTabs(i).isHot = false
        end if
    next
    function = hotTab
end function


' ========================================================================================
' frmOutputTabs_SubclassProc 
' ========================================================================================
function frmOutputTabs_SubclassProc ( _
            byval hWin   as HWND, _                 ' // Control window handle
            byval uMsg   as UINT, _                 ' // Type of message
            byval _wParam as WPARAM, _               ' // First message parameter
            byval _lParam as LPARAM, _               ' // Second message parameter
            byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
            byval dwRefData as DWORD_PTR _          ' // pointer to reference data
            ) as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMOUTPUT)
    static as long accumDelta
    
    ' keep track of last index we were over so that we only issue a 
    ' repaint if the cursor has moved off of the tab
    static as long nLastIdx = -1
    static as boolean isLastClose = false
    static hTooltip as hwnd
        
    select case uMsg
            
        case WM_MOUSEMOVE
            ' Track that we are over the control in order to catch the 
            ' eventual WM_MOUSELEAVE event
            dim tme as TrackMouseEvent
            tme.cbSize = sizeof(TrackMouseEvent)
            tme.dwFlags = TME_HOVER or TME_LEAVE
            tme.hwndTrack = hWin
            TrackMouseEvent(@tme) 

            if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( hWin, "", false, false )

            dim as long idx = frmOutputTabs_getHotTabHitTest( hWin )
            if idx <> nLastIdx then
                nLastIdx = idx
                AfxRedrawWindow(hWin)   
            end if
            
            dim as boolean isClose = isMouseOverRECT( hWin, gOutputCloseRect )
            if isClose <> isLastClose then
                isLastClose = isClose
                AfxRedrawWindow(hWin)   
            end if
            return 0
                        
        case WM_MOUSEHOVER
            dim as CWSTR wszTooltip = ""
            if isMouseOverRECT( hWin, gOutputCloseRect ) = true then
                ' Display the tooltip
                wszTooltip  = L(161, "Close")
            end if
            AfxSetTooltipText( hTooltip, hWin, wszTooltip )
            return 0

        case WM_MOUSELEAVE
            ' reset the hot tab index
            frmOutputTabs_getHotTabHitTest( hWin )
            nLastIdx = -1
            AfxDeleteTooltip( hTooltip, hWin )
            hTooltip = 0
            AfxRedrawWindow(hWin)   
            return 0
        
        case WM_LBUTTONUP
            if isMouseOverRECT( hWin, gOutputCloseRect ) = true then
                OnCommand_ViewOutput()   ' toggle the Output window off
            else
                dim as long idx = frmOutputTabs_getHotTabHitTest( hWin )
                if idx = -1 then exit function
                gOutputTabsCurSel = idx
                AfxRedrawWindow(hWin)   
                frmOutput_ShowHideOutputControls( HWND_FRMOUTPUT )
            end if
            return 0
            
        case WM_ERASEBKGND
            return true
        
        case WM_PAINT
            dim as PAINTSTRUCT ps
            dim as HDC hDC
            dim as RECT rc
            hDC = BeginPaint( hWin, @ps )
            
            SaveDC(hDC)
            dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
            dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

            dim memDC as HDC      ' Double buffering
            dim hbit as HBITMAP   ' Double buffering
            dim oldFont as HFONT
            dim oldPen as HPEN
            dim oldBrush as HBRUSH
            dim oldBmp as HBITMAP
            
            memDC = CreateCompatibleDC( hDC )
            hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
            
            SaveDC(memDC)
            oldBmp = selectObject( memDC, hbit )

            FillRect( memDC, @ps.rcPaint, ghOutput.hPanelBrush )
    
            for i as long = lbound(gOutputTabs) to ubound(gOutputTabs)
            rc = gOutputTabs(i).rcTab
            if (i = gOutputTabsCurSel) or (gOutputTabs(i).isHot = true) then
                SetBkColor( memDC, ghOutput.BackColorHot )
                SetTextColor( memDC, ghOutput.ForeColorHot )
                FillRect( memDC, @rc, ghOutput.hBackBrushHot )
            else
                SetBkColor( memDC, ghOutput.BackColor )
                SetTextColor( memDC, ghOutput.ForeColor )
                FillRect( memDC, @rc, ghOutput.hBackBrush )
            end if
            dim as long wsStyle = DT_NOPREFIX or DT_CENTER or DT_VCENTER or DT_SINGLELINE
            oldFont = selectObject( memDC, ghStatusBar.hFontStatusBar )
            DrawText( memDC, gOutputTabs(i).wszText.sptr, -1, Cast(lpRect, @rc), wsStyle )
            selectObject( memDC, oldFont )
            next

            dim as HPEN hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
            dim as long wsStyle = DT_NOPREFIX or DT_CENTER or DT_TOP 
            rc = gOutputCloseRect
            oldFont = selectObject( memDC, ghMenuBar.hFontSymbolSmall )
            SetTextColor( memDC, ghOutput.ForeColorHot )
            DrawText( memDC, wszClose, -1, Cast(lpRect, @rc), wsStyle )
            selectObject( memDC, oldFont )
            
            if isMouseOverRECT( hWin, rc ) then
                ' if we are hovered over the "X" close icon rect then highlight it
                oldPen = selectPen( memDC, hPenNull )
                oldBrush = selectObject( memDC, ghOutput.hCloseBrushHot )
                RoundRect( memDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
                SetBkColor( memDC, ghOutput.CloseBackColorHot )
                SetTextColor( memDC, ghOutput.ForeColorHot )
                oldFont = selectObject( memDC, ghMenuBar.hFontSymbolSmall )
                DrawText( memDC, wszClose, -1, cast(lpRect, @rc), wsStyle )
                selectObject( memDC, oldPen )
                selectObject( memDC, oldFont )
                selectObject( memDC, oldBrush )
            end if   

            ' Paint a simple line under the currently active tab
            dim as HPEN hPenSolid = CreatePen( PS_SOLID, pWindow->ScaleY(2), ghOutput.ForeColor )
            if gOutputTabsCurSel <> -1 then
                rc = gOutputTabs(gOutputTabsCurSel).rcText
                SetBkColor( memDC, ghOutput.ForeColor )
                oldPen = selectPen( memDC, hPenSolid )
                MoveToEx( memDC, rc.left, rc.bottom - pWindow->ScaleY(4), Null )
                LineTo( memDC, rc.right, rc.bottom - pWindow->ScaleY(4) )
                selectObject( memDC, oldPen )
            end if
            
            ' Paint a simple line at the top of the window that will act as a 
            ' visual separator between the Output window and the Scintilla window.
            if hPenSolid then DeleteObject( hPenSolid )
            hPenSolid = CreatePen( PS_SOLID, pWindow->ScaleY(1), ghOutput.Divider )
            SetBkColor( memDC, ghOutput.BackColor )
            oldPen = selectPen( memDC, hPenSolid )
            MoveToEx( memDC, ps.rcPaint.left, ps.rcPaint.top, null )
            LineTo( memDC, ps.rcPaint.right, ps.rcPaint.top )
            selectObject( memDC, oldPen )
            
            BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

            selectObject( memDC, oldBmp )

            ' Cleanup
            RestoreDC( memDC, -1 )
            DeleteObject( hbit )
            if memDC then DeleteDC( memDC )
            
            if hPenSolid then DeleteObject( hPenSolid )
            if hPenNull then DeleteObject( hPenNull )
            RestoreDC( hDC, -1 )
            
            EndPaint( hWin, @ps )
            exit function

        
        case WM_DESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hWin, @frmOutputTabs_SubclassProc, uIdSubclass )
            
    end select
    
    ' For messages that we don't deal with
    function = DefSubclassProc( hWin, uMsg, _wParam, _lParam )

end function


' ========================================================================================
' frmOutput_RichEdit_SubclassProc Window procedure
' ========================================================================================
function frmOutput_RichEdit_SubclassProc ( _
        byval hWin   as HWND, _                 ' // Control window handle
        byval uMsg   as UINT, _                 ' // Type of message
        byval _wParam as WPARAM, _               ' // First message parameter
        byval _lParam as LPARAM, _               ' // Second message parameter
        byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
        byval dwRefData as DWORD_PTR _          ' // pointer to reference data
        ) as LRESULT

    dim pWindow as CWindow ptr = AfxCWindowPtr(hWin)

    select case uMsg
        
        case WM_CONTEXTMENU 
            ' Create the right click popup menu
            dim as CWSTR wszText = RichEdit_GetSelText( hWin )
            dim hPopUpMenu as HMENU = CreatePopupMenu()
            if len(wszText) then 
                AppendMenu( hPopUpMenu, MF_ENABLED, IDM_CUT, wstr("Cut") )
                AppendMenu( hPopUpMenu, MF_ENABLED, IDM_COPY, wstr("Copy") )
            end if
            if RichEdit_CanPaste( hWin, 0 ) then
                if len(wszText) then 
                    AppendMenu( hPopUpMenu, MF_SEPARATOR, 0, "" )
                end if   
                AppendMenu( hPopUpMenu, MF_ENABLED, IDM_PASTE, wstr("Paste") )
            end if   

            dim as long nResult
            nResult = TrackPopupMenu( hPopUpMenu, TPM_RETURNCMD or TPM_NONOTIFY, _
                            loword(_lParam), hiword(_lParam), 0, HWND_FRMOUTPUT, 0 ) 
            select case nResult
                case IDM_CUT:   SendMessage( hWin, WM_CUT, 0, 0 )
                case IDM_COPY:  SendMessage( hWin, WM_COPY, 0, 0 )
                case IDM_PASTE: SendMessage( hWin, WM_PASTE, 0, 0 )
            end select
            DestroyMenu hPopUpMenu
            return 0

        case WM_DESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass( hWin, @frmOutput_RichEdit_SubclassProc, uIdSubclass )
    end select
    
    ' For messages that we don't deal with
    function = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

end function

' ========================================================================================
' frmOutput Window procedure
' ========================================================================================
function frmOutput_WndProc( _
            byval hwnd   as HWND, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

    select case uMsg
        HANDLE_MSG (hwnd, WM_COMMAND,     frmOutput_OnCommand)
        HANDLE_MSG (hwnd, WM_SIZE,        frmOutput_OnSize)
        HANDLE_MSG (hwnd, WM_LBUTTONUP,   frmOutput_OnLButtonUp)
        HANDLE_MSG (hwnd, WM_LBUTTONDOWN, frmOutput_OnLButtonDown)
        HANDLE_MSG (hwnd, WM_MOUSEMOVE,   frmOutput_OnMouseMove)
    
    case WM_ERASEBKGND
        return true

    case WM_PAINT            
        dim as PAINTSTRUCT ps
        dim as HDC hDC
        
        hDC = BeginPaint( hwnd, @ps )
        SaveDC( hDC )
        FillRect( hDC, @ps.rcPaint, ghOutput.hPanelBrush )
        RestoreDC( hDC, -1 )
        EndPaint( hwnd, @ps )
        exit function

    end select

    function = DefWindowProc( hwnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' Set the colors for the frmOutput controls. This is also called when the 
' user changes the theme (dark/light)
' ========================================================================================
function frmOutput_SetControlColors() as long
    dim cf as CHARFORMATW 
    cf.cbSize = sizeof(cf)
    cf.dwMask = CFM_COLOR
    cf.crTextColor = ghOutput.forecolorhot

    dim as HWND hCtl
    hCtl = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE )
    SendMessage( hCtl, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @cf) ) 
    SendMessage( hCtl, EM_SETBKGNDCOLOR , 0, cast(LPARAM, ghOutput.backcolor) )
    hCtl = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES )
    SendMessage( hCtl, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @cf) ) 
    SendMessage( hCtl, EM_SETBKGNDCOLOR , 0, cast(LPARAM, ghOutput.backcolor) )
    hCtl = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVRESULTS )
    ListView_SetBkColor( hCtl, ghOutput.BackColor )
    ListView_SetTextColor( hCtl, ghOutput.ForeColorHot )
    ListView_SetTextBkColor( hCtl, ghOutput.BackColor )
    hCtl = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVSEARCH )
    ListView_SetBkColor( hCtl, ghOutput.BackColor )
    ListView_SetTextColor( hCtl, ghOutput.ForeColorHot )
    ListView_SetTextBkColor( hCtl, ghOutput.BackColor )
    hCtl = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO )
    ListView_SetBkColor( hCtl, ghOutput.BackColor )
    ListView_SetTextColor( hCtl, ghOutput.ForeColorHot )
    ListView_SetTextBkColor( hCtl, ghOutput.BackColor )
    function = 0
end function

' ========================================================================================
' frmOutput_Show
' ========================================================================================
function frmOutput_Show( byval hwndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow
    pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

    HWND_FRMOUTPUT = pWindow->Create( hwndParent, "", @frmOutput_WndProc, 0, 0, 0, 180, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
    pWindow->ClassStyle = CS_DBLCLKS
    
    dim as HWND hCtl, hLV, hLB
    
    ' custom tab control (we paint our Tabs on this control)
    hCtl = _
    pWindow->AddControl("LABEL", , IDC_FRMOUTPUT_TABS, , 0, 0, 0, OUTPUT_TABS_HEIGHT, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or SS_NOTIFY or SS_LEFT, _
        WS_EX_LEFT or WS_EX_LTRREADING, , _
        cast(SUBCLASSPROC, @frmOutputTabs_SubclassProc), _
        IDC_FRMOUTPUT_TABS, cast(DWORD_PTR, @pWindow))

    hCtl = pWindow->AddControl("RICHEDIT", , IDC_FRMOUTPUT_TXTLOGFILE, "", _
                    0, 0, 0, 0, _
                    WS_CHILD or WS_TABSTOP or WS_VSCROLL or _
                    ES_MULTILINE or ES_LEFT or ES_AUTOVSCROLL, _
                    WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR, _
                    0, @frmOutput_RichEdit_SubclassProc, IDC_FRMOUTPUT_TXTLOGFILE, null )
        AfxSetWindowFont( hCtl, ghStatusBar.hFontStatusBar )
        SendMessage( hCtl, EM_SETEVENTMASK, 0, cast(LPARAM, ENM_SELCHANGE or ENM_CHANGE) )

    hCtl = pWindow->AddControl("RICHEDIT", , IDC_FRMOUTPUT_TXTNOTES, "", _
                    0, 0, 0, 0, _
                    WS_CHILD or WS_TABSTOP or WS_VSCROLL or _
                    ES_MULTILINE or ES_LEFT or ES_AUTOVSCROLL or ES_WANTRETURN, _
                    WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR, _
                    0, @frmOutput_RichEdit_SubclassProc, IDC_FRMOUTPUT_TXTNOTES, null )
        AfxSetWindowFont( hCtl, ghStatusBar.hFontStatusBar )
        SendMessage( hCtl, EM_SETEVENTMASK, 0, cast(LPARAM, ENM_SELCHANGE or ENM_CHANGE) )
        frmOutput_ShowNotes()   

    hLV = _
        pWindow->AddControl("LISTVIEW", , IDC_FRMOUTPUT_LVRESULTS, "", 0, 0, 0, 0, _
        WS_CHILD or WS_TABSTOP or LVS_REPORT or LVS_SINGLESEL, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LVRESULTS, Cast(DWORD_PTR, @pWindow))

        ' Configure the ListView
        AfxSetWindowFont( hLV, ghStatusBar.hFontStatusBar )
        dim as long dwExStyle = ListView_GetExtendedListViewStyle(hLV)
        dwExStyle = dwExStyle or LVS_EX_FULLROWselect or LVS_EX_DOUBLEBUFFER or LVS_EX_FLATSB
        ListView_SetExtendedListViewStyle(hLV, dwExStyle)
        ListView_MakeHeaderFlat(hLV)
        ListView_AddColumn( hLV, 0, "", pWindow->ScaleX(20) )
        ListView_AddColumn( hLV, 1, L(253, "Line"), pWindow->ScaleX(75) )
        ListView_AddColumn( hLV, 2, L(254, "File"), pWindow->ScaleX(250) )
        ListView_AddColumn( hLV, 3, L(255, "Description"), pWindow->ScaleX(480) )
        
    ' Search results Listview
    hLV = _
        pWindow->AddControl("LISTVIEW", , IDC_FRMOUTPUT_LVSEARCH, "", 0, 0, 0, 0, _
        WS_CHILD or WS_TABSTOP or LVS_REPORT or LVS_SINGLESEL, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LVSEARCH, cast(DWORD_PTR, @pWindow))

        ' Configure the ListView
        AfxSetWindowFont( hLV, ghStatusBar.hFontStatusBar )
        dwExStyle = ListView_GetExtendedListViewStyle(hLV)
        dwExStyle = dwExStyle or LVS_EX_FULLROWselect or LVS_EX_DOUBLEBUFFER or LVS_EX_FLATSB
        ListView_SetExtendedListViewStyle(hLV, dwExStyle)
        ListView_MakeHeaderFlat(hLV)
        ListView_AddColumn( hLV, 0, "", pWindow->ScaleX(20) )
        ListView_AddColumn( hLV, 1, L(253, "Line"), pWindow->ScaleX(75) )
        ListView_AddColumn( hLV, 2, L(254, "File"), pWindow->ScaleX(250) )
        ListView_AddColumn( hLV, 3, L(255, "Description"), pWindow->ScaleX(480) )

    ' TODO listview
    hLV = _
        pWindow->AddControl("LISTVIEW", , IDC_FRMOUTPUT_LVTODO, "", 0, 0, 0, 0, _
        WS_CHILD or WS_TABSTOP or LVS_REPORT or LVS_SINGLESEL, _
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LVTODO, cast(DWORD_PTR, @pWindow))
                                        
        ' Configure the ListView
        AfxSetWindowFont( hLV, ghStatusBar.hFontStatusBar )
        dwExStyle = ListView_GetExtendedListViewStyle(hLV)
        dwExStyle = dwExStyle or LVS_EX_FULLROWselect or LVS_EX_DOUBLEBUFFER or LVS_EX_FLATSB
        ListView_SetExtendedListViewStyle(hLV, dwExStyle)
        ListView_MakeHeaderFlat(hLV)
        ListView_AddColumn( hLV, 0, "", pWindow->ScaleX(20) )
        ListView_AddColumn( hLV, 1, L(253, "Line"), pWindow->ScaleX(75) )
        ListView_AddColumn( hLV, 2, L(254, "File"), pWindow->ScaleX(250) )
        ListView_AddColumn( hLV, 3, L(255, "Description"), pWindow->ScaleX(480) )

    frmOutput_SetControlColors
    frmOutput_PositionWindows
    
    function = 0

end function

