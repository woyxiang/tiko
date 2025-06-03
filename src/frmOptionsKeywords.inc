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

#include once "frmOptionsKeywords.bi"
#include once "frmOptionsKeywordsWinApi.bi"


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsKeywords
' ========================================================================================
private function frmOptionsKeywords_OnCommand( byval hwnd as HWND, _
                                                byval id as long, _
                                                byval hwndCtl as HWND, _
                                                byval codeNotify as UINT _
                                                ) as LRESULT
    select case id
        case IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS
            if codeNotify = EN_CHANGE then 
                ' The EN_CHANGE notification code is not sent when the ES_MULTILINE style is used and the text is sent through WM_SETTEXT. 
                ' Set the flag in gConfig to indicate that the keywords need to be saved to disk should
                ' the user press OK to close the Environment Options dialog.
                gConfig.bKeywordsDirty = true
            end if
    end select

    function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmOptionsKeywords
' ========================================================================================
private function frmOptionsKeywords_OnDestroy( byval hwnd as HWND ) as LRESULT
    Dim as HFONT hFont = AfxGetWindowFont(GetDlgItem(hwnd, IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS))
    DeleteFont(hFont)
    function = 0
end function


' ========================================================================================
' frmOptionsKeywords Window procedure
' ========================================================================================
private function frmOptionsKeywords_WndProc( byval hwnd   as HWND, _
                                                byval uMsg   as UINT, _
                                                byval wParam as WPARAM, _
                                                byval lParam as LPARAM _
                                                ) as LRESULT

    select case uMsg
        HANDLE_MSG (hwnd, WM_COMMAND, frmOptionsKeywords_OnCommand)
        HANDLE_MSG (hwnd, WM_DESTROY, frmOptionsKeywords_OnDestroy)
    end select

    function = DefWindowProc( hwnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmOptionsKeywords_Show
' ========================================================================================
public function frmOptionsKeywords_Show( byval hwndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow
    pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

    HWND_FRMOPTIONSKEYWORDS = pWindow->Create( hwndParent, "", @frmOptionsKeywords_WndProc, 0, 0, 0, 0, _
        WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING )
    ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
    
    dim as HWND hTextBox = _
    pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSKEYWORDS_TXTKEYWORDS, "", 0, 0, 386, 345, _
        WS_CHILD or WS_VISIBLE or WS_VSCROLL or WS_TABSTOP or ES_LEFT or _
        ES_MULTILINE or ES_AUTOVSCROLL or ES_WANTRETURN, _
        WS_EX_CLIENTEDGE or WS_EX_LEFT or WS_EX_LTRREADING )
    SendMessage( hTextBox, EM_SETLIMITTEXT, (1024 * 100), 0 )    ' max 100K text   
    Dim as HFONT hFont = pWindow->CreateFont("Courier New", 9)
    AfxSetWindowFont hTextBox, hFont, True

    gConfig.bKeywordsDirty = false
    SetWindowTextW hTextBox, wstr(gConfig.FBKeywords)
    
    function = 0
    
end function

