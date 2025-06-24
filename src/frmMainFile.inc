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
' Create a new editing window
' ========================================================================================
function OnCommand_FileNew( _
            byval hwnd as HWND, _
            byval bShowAfterLastDisplayTab as boolean _
            ) as clsDocument ptr
    
    dim pDocNew as clsDocument ptr
    
    dim as boolean bShowInTab = true
    if bShowAfterLastDisplayTab then bShowInTab = false
    
    pDocNew = frmMain_OpenFileSafely(hwnd, _
                                        true,  _   ' bIsNewFile
                                        false, _   ' bIsTemplate
                                        bShowInTab,  _   ' bShowInTab
                                        false, _   ' bIsInclude
                                        "", _      ' wszName
                                        0 )        ' pDocIn
    if bShowAfterLastDisplayTab then
        ' Move the new file to the end of the tabs currently displayed. Basically,
        ' right before the "+" Add New file button. 
        LoadExplorerFiles()
        ' Add the new document to the top tabcontrol
        dim as long idx = gTTabCtl.InsertTab( pDocNew, gTTabCtl.LastDisplayTab )  
        frmTopTabs_PositionWindows()
        gTTabCtl.SetFocusTab(idx)
        frmMain_SetFocusToCurrentCodeWindow
    end if
    
    function = pDocNew
end function


' ========================================================================================
' Open one or more files for editing
' ========================================================================================
function OnCommand_FileOpen( _
            byval hwnd as HWND, _
            byval bShowInTab as boolean = true, _
            byval bIsTemplate as boolean = false _
            ) as LRESULT

    dim pDoc as clsDocument ptr
    dim pDocIn as clsDocument ptr

    ' Display the Open File Dialog
    dim pItems as IShellItemArray ptr = AfxIFileOpenDialogMultiple(hwnd, IDM_FILEOPEN)
    if pItems = null then exit function
    dim dwItemCount as long, i as long, pItem as IShellItem ptr, pwszName as wstring ptr
    pItems->lpVtbl->GetCount(pItems, @dwItemCount)
    ' Variable length array to hold sequence of TabCtrl tabs to open. We save the sequence
    ' here and open them only after all over documents have been loaded. This will look
    ' visually more appealing for those files that will display in the top Tab Control.
    dim pDocTabs(dwItemCount - 1) as clsDocument ptr

    SetCursor( LoadCursor(0, IDC_WAIT) )
    pDoc = gTTabCtl.GetActiveDocumentPtr()
    if pDoc then SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORWAIT, 0 )

    ' Prevent the constant updating of the app caption bar with filename
    gApp.IsFileLoading = true  
    gApp.FileLoadingCount = 0
    
    for i = 0 To dwItemCount - 1
        pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
        if pItem then
            pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
            if pwszName then 
                ' Save the folder where this file was opened from into the gApp.wszLastOpenFolder
                ' variable so that subsequent uses of the File Open dialog will default to 
                ' this folder.
                gApp.wszLastOpenFolder = AfxStrPathName( "PATH", *pwszName )

                ' Test to see if the file is already loaded in the editor. if it is, then
                ' bypass loading it again thereby creating multiple ghost instances.
                pDoc = gApp.GetDocumentPtrByFilename(*pwszName) 
                if pDoc then
                    if pDoc->GetActiveScintillaPtr = 0 then pDocIn = pDoc
                end if   
                if (pDoc = 0) orelse (pDocIn <> 0) then
                    dim as boolean bIsNewFile = bIsTemplate
                    pDoc = frmMain_OpenFileSafely( _
                                hwnd, _
                                bIsNewFile, _  ' bIsNewFile
                                bIsTemplate, _ ' bIsTemplate
                                false, _       ' bShowInTab
                                false, _       ' bIsInclude
                                *pwszName, _   ' pwszName
                                pDocIn _       ' pDocIn
                                )     
                end if

                ' Give this document a default project type depending on its file extension
                if (pDoc->IsNewFlag = false) andalso (pDoc->ProjectFileType = FILETYPE_UNDEFINED) then
                    if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
                        gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )
                    end if   
                end if
                
                if bShowInTab then pDocTabs(i) = pDoc
                CoTaskMemFree(pwszName)
                pwszName = null
            end if
            pItem->lpVtbl->Release(pItem)
            pItem = null
        end if
    next
    pItems->lpVtbl->Release(pItems)
    
    LoadExplorerFiles()
    LoadFunctionsFiles()
    
    ' Display all of the tabs
    dim as long iTab = -1
    for i = lbound(pDocTabs) to ubound(pDocTabs)
        if pDocTabs(i) = 0 then continue for
        iTab = gTTabCtl.GetTabIndexByDocumentPtr(pDocTabs(i))
        if iTab = -1 then iTab = gTTabCtl.AddTab(pDocTabs(i))  ' Add the new document to the top tabcontrol
    next   
    if iTab > -1 then gTTabCtl.SetFocusTab(iTab)

    gApp.IsFileLoading = false
    
    gApp.wszPanelText = ""   ' reset filename text that displays in StatusBar panel
    frmMain_SetStatusbar

    ' Update the MRU list. Do this after the tabs are loaded because IsProjectLoading
    ' will prevent the list from updating.
    if gApp.IsProjectActive = false then
        for i = lbound(pDocTabs) to ubound(pDocTabs)
            if pDocTabs(i) then UpdateMRUList(pDocTabs(i)->DiskFilename) 
        next     
    end if
    
    SetCursor( LoadCursor(0, IDC_ARROW) )
    ResetScintillaCursors
    frmMain_PositionWindows
    frmMain_SetFocusToCurrentCodeWindow
    
    function = 0

end function


' ========================================================================================
' Save the incoming pDoc file to disk
' ========================================================================================
function OnCommand_FileSave( _
            byval hwnd as HWND, _
            byval pDoc as clsDocument ptr, _
            byval bSaveAs as boolean, _
            byval bSaveAll as boolean _
            ) as boolean

    if gApp.GetDocumentCount = 0 then exit function
    
    dim as long i, iStart, iEnd
    
    if bSaveAll then
        ' Save All
        iStart = 0
        iEnd = gTTabCtl.GetItemCount - 1
    else
        ' Save
        iStart = gTTabCtl.GetTabIndexByDocumentPtr(pDoc)
        iEnd = iStart
    end if 

    for i = iStart to iEnd
        ' Get the document pointer and then save file to disk
        if gTTabCtl.IsSafeIndex(i) then
            pDoc = gTTabCtl.tabs(i).pDoc
            if pDoc = 0 then continue for
        end if   

        If (cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 ) <> 0)) orelse _
           (pDoc->UserModified = true) orelse (bSaveAs = true) then   
            ' Do the actual saving to disk
            dim wszFilename as CWSTR = pDoc->DiskFilename
            if pDoc->SaveFile(bSaveAs) = false then
                pDoc->ApplyProperties
                pDoc->ParseDocument()
                ' Apply document properties to this file because the file extension may have
                ' changed. for example from Untitled to *.bas
                if pDoc->DiskFilename <> wszFilename then
                    if gApp.IsProjectActive then
                        gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )    
                    end if
                end if
            else
                'return true   
                exit for
            end if
        end if    

        ' Ensure that the Tab displays the correct filename in case it was changed
        gTTabCtl.SetTabText(-1)
    next
    
    LoadExplorerFiles()
    LoadFunctionsFiles()

    frmMain_SetFocusToCurrentCodeWindow

    function = false
end function


' ========================================================================================
' Rename the incoming pDoc file to disk
' ========================================================================================
function OnCommand_FileRename( _
            byval hwnd as HWND, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

    dim wszFilename_original  as wstring * MAX_PATH
    dim wszFilename_AutoSave  as wstring * MAX_PATH
    dim wszFilename  as wstring * MAX_PATH
    dim wszExtension as wstring * MAX_PATH

    ' Display the Rename Dialog (Save File Dialog)
    wszFilename_original = pDoc->DiskFilename
    dim pwszName as wstring ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wszFilename_original, @wszExtension, IDM_FILERENAME)
    if pwszName then
        wszFilename = *pwszName
        ' Apply document properties to this file because the file extension may have
        ' changed. for example from Untitled to *.bas
        if pDoc->DiskFilename <> wszFilename then
            if gApp.IsProjectActive then
                gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )    
            end if
        end if
        if AfxRenameFile( wszFilename_original, wszFilename ) then
            pDoc->DiskFilename = wszFilename
            ' Rename any existing autosave file
            wszFilename_AutoSave = OnCommand_FileAutoSaveGenerateFilename(wszFilename)
            AfxRenameFile( pDoc->AutoSaveFilename, wszFilename_AutoSave) 
            pDoc->AutoSaveFilename = wszFilename_AutoSave
        end if    

        CoTaskMemFree(pwszName)

    else
        return false
    end if

    LoadExplorerFiles()
    LoadFunctionsFiles()

    frmMain_SetFocusToCurrentCodeWindow

    function = 0
end function


' ========================================================================================
' Save all open files to disk
' ========================================================================================
function OnCommand_FileSaveAll( byval hwnd as HWND ) as boolean

    if gTTabCtl.GetItemCount = 0 then exit function
                
    return OnCommand_FileSave( hwnd, 0, false, true )
end function


' ========================================================================================
' Close current (or all) open file(s)
' ========================================================================================
function OnCommand_FileClose( _
            byval hwnd as HWND, _
            byval veFileClose as eFileClose, _
            byval nTabNum as long = -1 _
            ) as LRESULT

    dim wText as wstring * MAX_PATH  
    dim pDoc as clsDocument ptr
    dim as long r
    
    ' if a Project is not active then we need to save the current non-project notes
    ' when this file is closed. It is possible that this file is being closed and
    ' a project is being opened.
    if gApp.IsProjectActive = false then
        gApp.NonProjectNotes = AfxGetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES))
    end if
    
    ' Build a list of files to check
    dim as long nCurSel 
    ' if the incoming nTabNum is -1 then we want to close the current
    ' active document, otherwise close the specified tab.
    nCurSel = iif(nTabNum = -1, gTTabCtl.CurSel, nTabNum)

    dim as long nCount = gTTabCtl.GetItemCount

    if nCount = 0 then return true
    
    redim bCloseIndex(nCount - 1) as boolean

    for i as long = 0 to nCount - 1 
        SELECT CASE veFileClose
            case EFC_CLOSEALL:         bCloseIndex(i) = true
            case EFC_CLOSECURRENT:     bCloseIndex(i) = iif(i = nCurSel, true, false)
            case EFC_CLOSEALLOTHERS:   bCloseIndex(i) = iif(i <> nCurSel, true, false)
            case EFC_CLOSEALLFORWARD:  bCloseIndex(i) = iif(i > nCurSel, true, false) 
            case EFC_CLOSEALLBACKWARD: bCloseIndex(i) = iif(i < nCurSel, true, false)
        end select
    next
    
    ' Save the pDoc of the current active document in case we need to us it after
    ' the deletes are done to restore the correct tab index.
    dim as clsDocument ptr pDocActive = gTTabCtl.tabs(gTTabCtl.CurSel).pDoc
    
    ' Must do everything in reverse order because when Tabs are removed
    ' then the indexes would be out of sync if done in ascending order.
    for i as long = nCount - 1 to 0 step -1
        if bCloseIndex(i) = false then continue for

        pDoc = gTTabCtl.GetDocumentPtr(i)
        if pDoc = 0 then return true
    
        if cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified then
            gTTabCtl.SetFocusTab(i)
            r = MessageBox( hwnd, L(76,"Save current changes?") & wstr(" ") & wText, @wstr(APPNAME), _
                            MB_YESNOCANCEL Or MB_ICONQUESTION)
            if r = IDCANCEL then exit function
            if r = IDYES then 
                r = pDoc->SaveFile()
                if r = true then exit function   ' save was cancelled
                elseif r = IDNO then
                ' Delete any existing AutoSave file because at this point the user doesn't
                ' care about any unsaved changes.
                if AfxFileExists(pDoc->AutoSaveFilename) then AfxDeleteFile(pDoc->AutoSaveFilename)
            end if
        end if
        
        ' Hide the Scintilla editing windows
        for ii as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
            ShowWindow( pDoc->hWindow(ii), SW_HIDE)
        next
        SetRectEmpty( @pDoc->rcSplitButton )
        
        ' remove tab from gTTabCtl.tabs() array
        gTTabCtl.RemoveElement(i)
            
        ' Only remove this document from the global collection if it is not
        ' part of any active Project or if it is new/unsaved file. Files that are part of a project are
        ' all closed (tabs/nodes) automatically in one shot rather than individually.
        dim as boolean bRemoveFile 
        if pDoc->IsNewFlag = true then 
            bRemoveFile = true
        else
            bRemoveFile = iif( gApp.IsProjectActive, false, true )
        end if

        if bRemoveFile then
            ' Remove all references from the gdb2 database. Only remove the reference
            ' if the file is not part of an active project, otherwise, we will lose
            ' references to function names, etc.
            gdb2.dbDelete( pDoc->DiskFilename )
            gApp.RemoveDocument( pDoc )
            LoadExplorerFiles()
        end if
    next
        
    ' Set the active tab to the closest tab to the one just removed if it was the 
    ' current active tab that was removed.
    if (veFileClose = EFC_CLOSECURRENT) andalso (nCurSel = gTTabCtl.CurSel) then
        dim as long idx = gTTabCtl.CurSel - 1
        gTTabCtl.CurSel = -1  ' it is no longer valid
        if idx < lbound(gTTabCtl.tabs) then idx = lbound(gTTabCtl.tabs)
        if gTTabCtl.IsSafeIndex( idx ) then
            gTTabCtl.CurSel = gTTabCtl.SetFocusTab( idx )
        end if
    else
        ' we need to set our gTTabCtl.CurSel to the correct array index because
        ' the array has been resized and the cursel would no longer be valid.
        gTTabCtl.CurSel = -1  ' it is no longer valid
        if pDoc then gTTabCtl.CurSel = gTTabCtl.SetTabIndexByDocumentPtr(pDocActive)
    end if

    frmMain_PositionWindows
    frmExplorer_PositionWindows
    frmTopTabs_PositionWindows
    
    ' important to repaint the workspace to remove any separator bar and splitter
    ' rectangle should all documents be closed. Removes screen artifacts.
    AfxRedrawWindow( HWND_FRMMAIN )
    
    function = true
end function


' ========================================================================================
' Message received from Explorer popup menu
' ========================================================================================
function OnCommand_FileExplorerMessage( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
    
    if id = IDM_FILEOPEN_EXPLORERLISTBOX then
        OpenSelectedDocument( pDoc->DiskFilename, "" )

    elseif id = IDM_FILESAVE_EXPLORERLISTBOX then
        OnCommand_FileSave( HWND_FRMMAIN, pDoc, false )

    elseif id = IDM_FILESAVEAS_EXPLORERLISTBOX then
        OnCommand_FileSave( HWND_FRMMAIN, pDoc, true )

    elseif IDM_FILECLOSE_EXPLORERLISTBOX then
        dim as long nCurSel = gTTabCtl.GetTabIndexByDocumentPtr( pDoc )
        if nCurSel <> -1 then
            OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT, nCurSel )
        end if   
    end if   

    function = 0
end function


' ========================================================================================
' Generate an AutoSave filename based on the incoming full path DiskFilename
' ========================================================================================
function OnCommand_FileAutoSaveGenerateFilename( byval wszInFilename as CWSTR ) as CWSTR
    dim as CWSTR wszFilePath = AfxStrPathname( "PATH", wszInFilename )
    dim as CWSTR wszFilename = AfxStrPathname( "NAMEX", wszInFilename )
    if len(wszFilename) = 0 then return ""
    dim as CWSTR wszAutoSaveFilename = wszFilePath & "#" & wszFilename & "#"
    return wszAutoSaveFilename
end function


' ========================================================================================
' Check the incoming filename to determine if a newer AutoSave file already exists. If
' yes, then ask the user if he wants to use the AutoSave file. if yes, then copy then
' delete the original file and rename the AutoSave file to the original returning the
' original filename.
' ========================================================================================
function OnCommand_FileAutoSaveFileCheck( byval wszInFilename as CWSTR ) as CWSTR

#if ENABLE_AUTOSAVE
    dim as FILETIME ftOriginal
    dim as FILETIME ftAutoSave
    dim as CWSTR wszAutoSaveFilename = OnCommand_FileAutoSaveGenerateFilename(wszInFilename)
    dim as long nResult
    
    ftOriginal = AfxGetFileLastWriteTime( wszInFilename )
    ftAutoSave = AfxGetFileLastWriteTime( wszAutoSaveFilename )
    
    if AfxFileTimeToVariantTime(ftAutoSave) > AfxFileTimeToVariantTime(ftOriginal) then
        nResult = MessageBox( HWND_FRMMAIN, _
                L(428,"An automatically saved version the following file is more recent:") & vbcrlf & _
                chr(34) & wszInFilename & chr(34) & vbcrlf & vbcrlf & _
                L(429,"Would you like to load the auto save version instead?"), _
                L(427, "Auto Save"), MB_ICONQUESTION or MB_YESNO) 

        if nResult = IDYES then
            ' User wants to re-use the previously saved AutoSave file. Delete the
            ' incoming filename and rename the AutoSave to the original name.
            AfxDeleteFile( wszInFilename )
            AfxRenameFile( wszAutoSaveFilename, wszInFilename )
        elseif nResult = IDNO then
            ' User does not want to use the existing AutoSave file so simply delete it.
            AfxDeleteFile( wszAutoSaveFilename )
        end if   
        
    end if
#endif

    ' No matter if the original or autosave file is selected by the
    ' user, the filename being returned will inevitably be the same
    ' as the incoming filename.
    function = wszInFilename
end function


' ========================================================================================
' Enable/Disable File auto save features
' ========================================================================================
function OnCommand_FileAutoSaveTimerProc( _
            byval hWndTimer as HWND, _
            byval uMsg as long, _
            byval idEvent as long, _
            byval dwTime as long _
            ) as long
                    
    ' Iterate all of the documents and determine which ones need to have
    ' the special #filename# temporary file created.
    if gConfig.AutoSaveFiles = false then exit function

#if ENABLE_AUTOSAVE    
    dim pDoc as clsDocument ptr
    pDoc = gApp.pDocList
    do until pDoc = 0
        if pDoc->AutoSaveRequired then
            if len(pDoc->AutoSaveFilename) then
            pDoc->SaveFile( false, true )
            pDoc->AutoSaveRequired = false
            end if   
        end if
        pDoc = pDoc->pDocnext
    loop
#endif

    function = 0
end function

' ========================================================================================
' Kill AutoSave timer that checks for file changes between saves. This function
' is called whenever the AutoSave menu item is clicked or when editor closes.
' ========================================================================================
function OnCommand_FileAutoSaveKillTimer() as LRESULT
    KillTimer( HWND_FRMMAIN, gConfig.idAutoSaveTimer )
    function = 0
end function

' ========================================================================================
' Start AutoSave timer that checks for file changes between saves. This function
' is called whenever the AutoSave menu item is clicked or when the application
' starts and gConfig.AutoSaveFiles = true
' ========================================================================================
function OnCommand_FileAutoSaveStartTimer() as LRESULT
    if gConfig.AutoSaveFiles = false then exit function
    OnCommand_FileAutoSaveKillTimer()
    SetTimer( HWND_FRMMAIN, gConfig.idAutoSaveTimer, _
                gConfig.AutoSaveInterval * 1000, _
                cast(TimerProc, @OnCommand_FileAutoSaveTimerProc) )       
    function = 0
end function

' ========================================================================================
' Enable/Disable File auto save features
' ========================================================================================
function OnCommand_FileAutoSave( byval hwnd as HWND ) as LRESULT
    ' if AutoSave is enabled then start timer that checks for AutoSave conditions
    if gConfig.AutoSaveFiles then 
        ' any files that were dirty prior to Autosave being enabled should
        ' now get processed by the new timer when it fires.
        OnCommand_FileAutoSaveStartTimer()
            
    elseif gConfig.AutoSaveFiles = false then
        ' if AutoSave is disabled then delete any existing #filename# temporary files
        OnCommand_FileAutoSaveKillTimer()
        ' Delete any temporary ## files
        dim pDoc as clsDocument ptr = gApp.pDocList
        do until pDoc = 0
            AfxDeleteFile( pDoc->AutoSaveFilename )
            pDoc = pDoc->pDocnext
        loop
    
    end if
    
    function = 0
end function


' ========================================================================================
' Load a previously saved session
' ========================================================================================
function OnCommand_FileLoadSession( byval hwnd as HWND ) as LRESULT
    ' Display the Open File Dialog
    dim pwszName as wstring ptr = AfxIFileOpenDialogW(hwnd, IDM_LOADSESSION)
    if pwszName then 
        if gApp.IsProjectActive then 
            if OnCommand_ProjectClose(hwnd) = false then exit function
        else   
            if OnCommand_FileClose(hwnd, EFC_CLOSEALL) = false then exit function
        end if
        ' Clear any previous info from the Output windows
        frmOutput_ResetAllControls()
        gConfig.LoadSessionFile(*pwszName) 
        CoTaskMemFree(pwszName)
    end if
    function = 0
end function


' ========================================================================================
' Save a session
' ========================================================================================
function OnCommand_FileSaveSession( byval hwnd as HWND ) as LRESULT
    dim wszFilename  as wstring * MAX_PATH = ""
    dim wszExtension as wstring * MAX_PATH = "session"
    dim pwszName as wstring ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wszFilename, @wszExtension, IDM_SAVESESSION)
    if pwszName then
        gConfig.SaveSessionFile(*pwszName)
        CoTaskMemFree(pwszName)
    end if
    function = 0
end function


' ========================================================================================
' Close a session
' ========================================================================================
function OnCommand_FileCloseSession( byval hwnd as HWND ) as LRESULT
    gConfig.wszLastActiveSession = ""
    gConfig.SaveConfigFile()
    gConfig.CloseSessionFile(gApp.SessionFilename)
    function = 0
end function
        
