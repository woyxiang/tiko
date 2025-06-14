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
    
#include once "frmRecent.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Load the Recent Files/Projects Treeview
' ========================================================================================
public function frmRecent_LoadTreeview() as LRESULT

    dim as CWSTR wzStr 
    dim as HTREEITEM hNode, hFirstNode
    dim as boolean fEmpty = true 

    dim as HWND hTree = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEVIEW) 

    Treeview_DeleteAllItems(hTree) 

    ''  RECENT FILES
    wzStr = ucase(L(10, "Recent Files")) 
    hNode = TreeView_AppendItem( hTree, TVI_ROOT, wzStr )
    hFirstNode = hNode
    gApp.hRecentFilesRootNode = hNode
    Treeview_SetBold( hTree, hNode, true) 
    for i as long = 0 to 9 
        ' if the file no longer exists then remove it from the MRU list 
        gConfig.MRU(i) = ProcessFromCurdriveApp(gConfig.MRU(i)) 
        if AfxFileExists(gConfig.MRU(i)) then 
            TreeView_AppendItem( hTree, hNode, gConfig.MRU(i) )
            fEmpty = false 
        end if 
    next 
    if fEmpty then 
        wzStr = @L(11,"(Empty)") 
        TreeView_AppendItem( hTree, hNode, wzStr, 0 ) 
    end if 
    Treeview_Expand(hTree, hNode, TVE_EXPAND)  


    ''  RECENT PROJECTS
    wzStr = ucase(L(219, "Recent Projects")) 
    hNode = TreeView_AppendItem( hTree, TVI_ROOT, wzStr ) 
    gApp.hRecentProjectsRootNode = hNode
    Treeview_SetBold( hTree, hNode, true) 
    for i as long = 0 To 9 
        ' if the file no longer exists then remove it from the MRU list 
        gConfig.MRUPROJECT(i) = ProcessFromCurdriveApp(gConfig.MRUPROJECT(i)) 
        if AfxFileExists(gConfig.MRUPROJECT(i)) then 
            TreeView_AppendItem( hTree, hNode, gConfig.MRUPROJECT(i) )
            fEmpty = false 
        end if 
    next 
    if fEmpty then 
        wzStr = @L(11,"(Empty)") 
        TreeView_AppendItem( hTree, hNode, wzStr )
    end if 
    Treeview_Expand(hTree, hNode, TVE_EXPAND)  
    
    ' Set the view to the first line
    TreeView_EnsureVisible(hTree, hFirstNode)
    
    function = 0
        
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
public function frmRecent_PositionWindows() as LRESULT
    
    dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMRECENT)
    if pWindow = 0 then exit function
    
    ' Get the entire client area
    dim as Rect rc
    GetClientRect(HWND_FRMRECENT, @rc)

    dim as HWND hBtn
    dim as long nLeft, nTop, nWidth, nHeight
    
    nTop = pWindow->ScaleY(12)
    hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNNEWFILE)
    SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                        rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                        SWP_SHOWWINDOW or SWP_NOZORDER )
    AfxRedrawWindow(hBtn)

    nTop = nTop + pWindow->ScaleY(40)
    hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENFILE)
    SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                        rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                        SWP_SHOWWINDOW or SWP_NOZORDER )
    AfxRedrawWindow(hBtn)
    
    nTop = nTop + pWindow->ScaleY(40)
    hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNNEWPROJECT)
    SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                        rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                        SWP_SHOWWINDOW or SWP_NOZORDER )
    AfxRedrawWindow(hBtn)

    nTop = nTop + pWindow->ScaleY(40)
    hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENPROJECT)
    SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                        rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                        SWP_SHOWWINDOW or SWP_NOZORDER )
    AfxRedrawWindow(hBtn)

    nTop = nTop + pWindow->ScaleY(48)
    hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEVIEW)
    SetWindowPos( hBtn, 0, pWindow->ScaleX(4), nTop, _
                    rc.Right-rc.Left-pWindow->ScaleX(8), rc.bottom - nTop, _
                    SWP_SHOWWINDOW or SWP_NOZORDER )

    function = 0
end function
    

' ========================================================================================
' Process WM_SIZE message for window/dialog: frmMain
' ========================================================================================
private function frmRecent_OnSize( byval hwnd as HWND, _
                                    byval state as UINT, _
                                    byval cx as long, _
                                    byval cy as long _
                                    ) as LRESULT
    if state <> SIZE_MINIMIZED then
        ' Position all of the child windows
        frmRecent_PositionWindows
    end if
    function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmRecent
' ========================================================================================
private function frmRecent_OnCommand( byval hwnd as HWND, _
                                        byval id as long, _
                                        byval hwndCtl as HWND, _
                                        byval codeNotify as UINT _
                                        ) as LRESULT
    Select case id
        case IDC_FRMRECENT_BTNNEWFILE
            if codeNotify = BN_CLICKED then
                OnCommand_FileNew(HWND_FRMMAIN) 
            end if
        case IDC_FRMRECENT_BTNOPENFILE
            if codeNotify = BN_CLICKED then
                OnCommand_FileOpen(HWND_FRMMAIN) 
            end if
        case IDC_FRMRECENT_BTNNEWPROJECT
            if codeNotify = BN_CLICKED then
                OnCommand_ProjectNew(HWND_FRMMAIN) 
            end if
        case IDC_FRMRECENT_BTNOPENPROJECT
            if codeNotify = BN_CLICKED then
                OnCommand_ProjectOpen(HWND_FRMMAIN) 
            end if
    end Select
        
    function = 0
end function

        
' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmRecent
' ========================================================================================
private function frmRecent_OnDestroy( byval hwnd as HWND ) as LRESULT
    dim pWindow as CWindow Ptr = AfxCWindowPtr(hwnd)

    dim pButton as CXpButton ptr 
    pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNNEWFILE)
    if pButton then delete pButton
    pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNOPENFILE)
    if pButton then delete pButton
    pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNNEWPROJECT)
    if pButton then delete pButton
    pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNOPENPROJECT)
    if pButton then delete pButton
    
    dim as HFONT hFont
    dim as hwnd hTree
    hTree = GetDlgItem( hwnd, IDC_FRMRECENT_TREEVIEW)
    
    function = 0
end function


' ========================================================================================
' Processes messages for the subclassed frmRecent Treeview controls.
' ========================================================================================
private function frmRecent_Tree_SubclassProc ( _
                    byval hwnd   as HWND, _                 ' // Control window handle
                    byval uMsg   as UINT, _                 ' // Type of message
                    byval wParam as WPARAM, _               ' // First message parameter
                    byval lParam as LPARAM, _               ' // Second message parameter
                    byval uIdSubclass as UINT_PTR, _        ' // The subclass ID
                    byval dwRefData as DWORD_PTR _          ' // Pointer to reference data
                    ) as LRESULT

    ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
    if (uMsg = WM_KEYUP) andalso (loword(wParam) = VK_RETURN) then uMsg = WM_LBUTTONDBLCLK
        
    select case uMsg

        case WM_GETDLGCODE
            ' All keyboard input
            function = DLGC_WANTALLKEYS
            Exit function

        case WM_LBUTTONDBLCLK
            dim as HWND hTree = hwnd
            dim as HTREEITEM hItem = Treeview_GetSelection(hTree)
            if hItem then   
                ' Get the text for the line that was double clicked on.
                dim as CWSTR wzStr = String(MAX_PATH, 0)
                TreeView_GetItemText( hTree, hItem, wzStr, MAX_PATH )
                if AfxFileExists(wzStr) then
                    if Treeview_GetParent(hTree, hItem) = gApp.hRecentProjectsRootNode then 
                        OpenProjectSafely(HWND_FRMMAIN, wzStr) 
                    Elseif Treeview_GetParent(hTree, hItem) = gApp.hRecentFilesRootNode then 
                        frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                                false, _    ' bIsNewFile
                                                false, _    ' bIsTemplate
                                                true, _     ' bShowInTab
                                                false, _    ' bIsInclude
                                                wzStr, _  ' pwszName
                                                0 )  ' pDocIn
                    end if
                end if
                frmMain_PositionWindows
                SetFocusScintilla
            end if   
            
        case WM_KEYUP
            select case loword(wParam)
            case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
            end Select
            return 0

        case WM_CHAR   ' prevent the annoying beep!
            if wParam = VK_RETURN then return 0

        case WM_DESTROY
            ' REQUIRED: Remove control subclassing
            RemoveWindowSubclass hwnd, @frmRecent_Tree_SubclassProc, uIdSubclass

    end Select

    ' Default processing of Windows messages
    function = DefSubclassProc(hwnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' frmRecent Window procedure
' ========================================================================================
private function frmRecent_WndProc( byval hwnd   as HWND, _
                                    byval uMsg   as UINT, _
                                    byval wParam as WPARAM, _
                                    byval lParam as LPARAM _
                                    ) as LRESULT

    select case uMsg
        HANDLE_MSG (hwnd, WM_SIZE,    frmRecent_OnSize)
        HANDLE_MSG (hwnd, WM_COMMAND, frmRecent_OnCommand)
        HANDLE_MSG (hwnd, WM_DESTROY, frmRecent_OnDestroy)
    end Select

    function = DefWindowProc(hwnd, uMsg, wParam, lParam)

end function


' ========================================================================================
' frmRecent_Show
' ========================================================================================
public function frmRecent_Show( byval hwndParent as HWND ) as LRESULT

    '  Create the main window and child controls
    dim pWindow as CWindow ptr = new CWindow
    pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

    HWND_FRMRECENT = pWindow->Create( hwndParent, "Recent Window", @frmRecent_WndProc, 0, 0, 250, 0, _
        WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)

    ' Disable background erasing by only assigning the one style
    pWindow->ClassStyle = CS_DBLCLKS
    pWindow->SetFont( "Segoe UI Semibold", 9)
    pWindow->Brush = GetSysColorBrush(COLOR_BTNFACE)

    dim pButton as CXpButton ptr 
    pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNNEWFILE, L(411, "New File"), 0, 0, 0, 0)
    pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
    pButton->DisableTheming
    pButton->SetTextForeColor(BGR(255,255,250))
    pButton->SetButtonBkColor(BGR(9,71,113))
    pButton->SetButtonBkColorHot(BGR(63,63,70))

    pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNOPENFILE, L(248, "Open File"), 0, 0, 0, 0)
    pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
    pButton->DisableTheming
    pButton->SetTextForeColor(BGR(255,255,250))
    pButton->SetButtonBkColor(BGR(9,71,113))
    pButton->SetButtonBkColorHot(BGR(63,63,70))
    
    pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNOPENPROJECT, L(249, "Open Project"), 0, 0, 0, 0)
    pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
    pButton->DisableTheming
    pButton->SetTextForeColor(BGR(255,255,250))
    pButton->SetButtonBkColor(BGR(9,71,113))
    pButton->SetButtonBkColorHot(BGR(63,63,70))

    pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNNEWPROJECT, L(180, "New Project"), 0, 0, 0, 0)
    pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
    pButton->DisableTheming
    pButton->SetTextForeColor(BGR(255,255,250))
    pButton->SetButtonBkColor(BGR(9,71,113))
    pButton->SetButtonBkColorHot(BGR(63,63,70))

    dim as HWND hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMRECENT_TREEVIEW, "", 0, 0, 0, 0, _
        WS_CHILD or WS_VISIBLE or WS_TABSTOP or _
        TVS_SHOWSELALWAYS or TVS_FULLROWSELECT or TVS_TRACKSELECT, _   
        WS_EX_LEFT or WS_EX_RIGHTSCROLLBAR, , _
        cast(SUBCLASSPROC, @frmRecent_Tree_SubclassProc), IDC_FRMRECENT_TREEVIEW, cast(DWORD_PTR, @pWindow))

    dim as HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
    SendMessage( hTree, WM_SETFONT, cast(WPARAM, hFont), Ctrue )
    SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(24), 0 )
    SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   
    SendMessage( hTree, TVM_SETBKCOLOR, 0, cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )

    frmRecent_LoadTreeview
    frmRecent_PositionWindows

    function = 0
    
end function



