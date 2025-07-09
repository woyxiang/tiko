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

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmHelpViewer_PositionWindows() as LRESULT
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMHELPVIEWER)
    if pWindow = 0 then exit function
    
   
    function = 0
end function


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

'    if hWndChild = GetDlgItem( hwnd, IDC_FRMOPTIONS_LBLCATEGORY ) then
'        ' Set the category label to blue text 
'        SetTextColor hdc, 8388608
'        SetBkColor   hdc, GetSysColor( COLOR_BTNFACE )
'        return GetSysColorBrush( COLOR_BTNFACE )
'    end if
    
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
        HANDLE_MSG (hwnd, WM_CREATE,   frmHelpViewer_OnCreate)
        HANDLE_MSG (hwnd, WM_CLOSE,    frmHelpViewer_OnClose)
        HANDLE_MSG (hwnd, WM_DESTROY,  frmHelpViewer_OnDestroy)
        HANDLE_MSG (hwnd, WM_COMMAND,  frmHelpViewer_OnCommand)
        HANDLE_MSG (hwnd, WM_NOTIFY,   frmHelpViewer_OnNotify)
        HANDLE_MSG (hwnd, WM_CTLCOLORSTATIC, frmHelpViewer_OnCtlColorStatic)
    end select

    function = DefWindowProc(hwnd, uMsg, wParam, lParam)

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
        WS_POPUP or WS_CAPTION or WS_SYSMENU or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME or WS_EX_CONTROLPARENT or WS_EX_LEFT )
    pWindow->ClassStyle = CS_DBLCLKS

    pWindow->SetClientSize(625, 460)
    pWindow->Center(pWindow->hWindow, hWndParent)

    frmHelpViewer_PositionWindows()
    
    ShowWindow( HWND_FRMHELPVIEWER, SW_SHOW )
        
    return 0
end function

