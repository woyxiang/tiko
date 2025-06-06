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

#include once "frmMain.bi"
#include once "clsDocument.bi"


' ========================================================================================
' View fold actions
' ========================================================================================
function OnCommand_ViewFold( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
    
    select case id      
        case IDM_FOLDTOGGLE: pDoc->FoldToggle( pDoc->GetCurrentLineNumber() )
        case IDM_FOLDBELOW:  pDoc->FoldToggleOnwards( pDoc->GetCurrentLineNumber() )
        case IDM_FOLDALL:    pDoc->FoldAll()
        case IDM_UNFOLDALL:  pDoc->UnFoldAll()
    end select
    
    function = 0
end function

' ========================================================================================
' View zoom actions
' ========================================================================================
function OnCommand_ViewZoom( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
    
    dim as HWND hEdit = pDoc->hWndActiveScintilla
    select case id      
        Case IDM_ZOOMIN:  SciExec(hEdit, SCI_ZOOMIN, 0, 0)
        Case IDM_ZOOMOUT: SciExec(hEdit, SCI_ZOOMOUT, 0, 0)
    end select
    
    function = 0
end function

' ========================================================================================
' View toggle split editor
' ========================================================================================
function OnCommand_ViewSplit( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

    select case id
        case IDM_SPLITLEFTRIGHT
            ' If already left/right then toggle it closed

        case IDM_SPLITTOPBOTTOM
            ' If already top/bottom then toggle it closed
            if pDoc->EditorSplitMode = SplitTopBottom then
                pDoc->SplitY = 0
                pDoc->EditorSplitMode = SplitNone
                ' Get the number of visible lines in the Top split and then
                ' set the MAIN document's first top line to be the current
                ' MAIN editing line numebr less thhe top split number. This 
                ' makes it less visibly jarring. Switch the active scintilla
                ' to ensure that GetCurrentLineNumber gets the value from the
                ' correct document.
                pDoc->hWndActiveScintilla = pDoc->hWindow(0)
                dim as long nLines =  pDoc->LinesPerPage(1)
                dim as long nCurLine =  pDoc->GetCurrentLineNumber()
                SciExec( pDoc->hWindow(0), SCI_SETFIRSTVISIBLELINE, nCurLine - nLines, 0) 
            else    
                dim as RECT rc = AfxGetWindowRect(pDoc->hWindow(0))
                MapWindowpoints( HWND_DESKTOP, HWND_FRMMAIN, cast(point ptr, @rc), 2 )
                pDoc->SplitY = (rc.bottom - rc.top) / 2
                pDoc->EditorSplitMode = SplitTopBottom 
                dim as long nLines =  pDoc->LinesPerPage(0) / 2 - 3
                dim as long nCurLine =  pDoc->GetCurrentLineNumber()
                SciExec( pDoc->hWindow(1), SCI_SETFIRSTVISIBLELINE, nCurLine - nLines, 0) 
                SciExec( pDoc->hWindow(0), SCI_SETFIRSTVISIBLELINE, nCurLine, 0) 
            end if
    end select
    
    frmMain_PositionWindows()
    frmMain_SetFocusToCurrentCodeWindow()
            
    function = 0
end function


' ========================================================================================
' Toggle editing focus amongst open windows depending whether spilt in effect
' ========================================================================================
function OnCommand_ViewToggleFocus( byval pDoc as clsDocument ptr ) as LRESULT
    if pDoc = 0 then exit function
    
    dim hCurrent as HWND = pDoc->hWndActiveScintilla
    dim hFindCtl as HWND = GetDlgItem( HWND_FRMFINDREPLACE, IDC_FRMFINDREPLACE_TXTFIND )
    if IsWindowVisible(HWND_FRMFINDREPLACE) = false then hFindCtl = 0
        
    dim idx as long
    
    dim hWin(2) as HWND
    hWin(0) = pDoc->hWindow(0): if GetFocus() = hWin(0) then idx = 0
    hWin(1) = pDoc->hWindow(1): if GetFocus() = hWin(1) then idx = 1
    hWin(2) = hFindCtl: if GetFocus() = hWin(2) then idx = 2
    
    do
        idx += 1
        if idx > ubound(hWin) then idx = lbound(hWin)
        hCurrent = hWin(idx)
    loop until IsWindowVisible(hCurrent)

    SetFocus(hCurrent)
    frmMain_SetStatusbar
    
    function = 0
end function


' ========================================================================================
' View FunctionList, Explorer, Output, Bookmarks windows
' ========================================================================================
function OnCommand_ViewFunctionList() as LRESULT
    ShowWindow( HWND_FRMPANEL, SW_SHOW )
    gPanel.hActiveChild = HWND_FRMFUNCTIONS
    frmPanel_PositionWindows
    frmMain_PositionWindows
    function = 0
end function

function OnCommand_ViewExplorer() as LRESULT
    ShowWindow( HWND_FRMPANEL, SW_SHOW )
    gPanel.hActiveChild = HWND_FRMEXPLORER
    frmPanel_PositionWindows
    frmMain_PositionWindows
    function = 0
end function

function OnCommand_ViewBookmarksList() as LRESULT
    ShowWindow( HWND_FRMPANEL, SW_SHOW )
    gPanel.hActiveChild = HWND_FRMBOOKMARKS
    frmPanel_PositionWindows
    frmMain_PositionWindows
    function = 0
end function

function OnCommand_ViewOutput() as LRESULT
    ShowWindow( HWND_FRMOUTPUT, iif(IsWindowVisible(HWND_FRMOUTPUT), SW_HIDE, SW_SHOW) )
    frmMain_PositionWindows
    function = 0
end function

' ========================================================================================
' View Restore Main window size
' ========================================================================================
function OnCommand_ViewRestoreMain() as LRESULT
    dim rc as Rect
    SystemParametersInfo( SPI_GETWORKAREA, 0, @rc, 0 )
    MoveWindow( HWND_FRMMAIN, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE )
    function = 0
end function

' ========================================================================================
' Move Explorer Window Left or Right
' ========================================================================================
function OnCommand_ViewExplorerPosition() as LRESULT
    gConfig.ExplorerPositionRight = not gConfig.ExplorerPositionRight 
    frmMain_PositionWindows()
    function = 0
end function
                
' ========================================================================================
' View Todo, Notes
' ========================================================================================
function OnCommand_ViewToDo() as LRESULT
    ' toggle close the Output window if it is open and already at ToDo 
    if IsWindowVisible(HWND_FRMOUTPUT) then
        if gOutputTabsCurSel = 3 then
            ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
            frmMain_PositionWindows
            exit function
        end if
    else      
        ShowWindow( HWND_FRMOUTPUT, SW_SHOW )
    end if
    frmMain_PositionWindows
    gOutputTabsCurSel = 3
    frmOutput_PositionWindows
    AfxRedrawWindow( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABS) )
    function = 0
end function

function OnCommand_ViewNotes() as LRESULT
    ' toggle close the Output window if it is open and already at Notes
    if IsWindowVisible(HWND_FRMOUTPUT) then
        if gOutputTabsCurSel = 4 then
            ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
            frmMain_PositionWindows
            exit function
        end if
    else      
        ShowWindow( HWND_FRMOUTPUT, SW_SHOW )
    end if
    frmMain_PositionWindows
    gOutputTabsCurSel = 4
    frmOutput_PositionWindows
    AfxRedrawWindow( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABS) )
    function = 0
end function
