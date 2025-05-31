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
' Add one or more files to the project (but don't open them to Tabs)
' ========================================================================================
Function OnCommand_ProjectFilesAdd( byval HWnd As HWnd ) as LRESULT
   OnCommand_FileOpen( HWnd, false )
   LoadExplorerFiles()
   function = 0
end function


' ========================================================================================
' Save project file to disk
' ========================================================================================
Function OnCommand_ProjectSave( _
            byval HWnd As HWnd, _
            ByVal bSaveAs As BOOLEAN = False _
            ) As LRESULT
            
   ' Save all dirty files
   OnCommand_FileSaveAll( HWnd )
   
   ' Do the actual saving to disk of the Project file
   gApp.SaveProject(bSaveAs)

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function 


' ========================================================================================
' Close currently active project
' ========================================================================================
Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT

   ' Save the project configuration to disk and any dirty files
   ' Do the actual saving to disk of the Project file
   gApp.SaveProject(false)
   if OnCommand_FileClose( HWnd, EFC_CLOSEALL ) = false then exit function
   
   ' If the entire app is closing down then we do not have to process the
   ' remaining project cleanup commands. 
   if gApp.IsShutdown then return true
      
   ' If we get this far then we can safely remove all of the open pDocs
   gApp.RemoveAllDocuments
   
   ' If the Find/Replace window is open then close it.
   DestroyWindow(HWND_FRMFINDREPLACE)

   ' Reset the Project variables
   gApp.IsProjectActive      = false
   gApp.ProjectName          = ""
   gApp.ProjectFilename      = ""
   gApp.ProjectBuild         = ""
   gApp.ProjectOther32       = ""
   gApp.ProjectOther64       = ""
   gApp.ProjectManifest      = false
   gApp.ProjectNotes         = ""
   gApp.ProjectCommandLine   = ""
   gApp.wszLastOpenFolder    = ""
   gApp.ProjectDefaultFont   = "Segoe UI,9,400,0,0,0,1"
   
   gApp.wszPanelText = ""
   gApp.hIconPanel = 0
       
   ' Clear any previous info from the Output windows
   frmOutput_ResetAllControls
   
   ' Reset the ProjectNotes text box with the generic non-project related notes
   AfxSetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), gApp.NonProjectNotes)

   ' If the compiler Output window is open then ensure that it is now hidden
   ShowWindow( HWND_FRMOUTPUT, SW_HIDE )   
   
   frmExplorer_PositionWindows
   frmMain_PositionWindows
   LoadExplorerFiles()
    
   frmMain_SetStatusbar
   frmMain_SetFocusToCurrentCodeWindow

   Function = true
End Function


' ========================================================================================
' Create a new project
' ========================================================================================
Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT

   
   ' Ensure that the Project Manager window has been created (show it now as well)
   frmProjectOptions_Show( HWnd, true ) 
   
   if gApp.IsNewProjectFlag then 
      ' Need to remove resource file (if created) from the document list otherwise it
      ' will be loaded again and it will display twice in the Explorer.
      gApp.RemoveAllDocuments()
      
      ' Open the new project so it displays in the Explorer treeview. Pass the 
      ' variable wText because passing gApp.ProjectFilename will cause that variable
      ' to get reset because the parameter is byref and eventually that variable
      ' gets assigned a null.
      dim as CWSTR wszText = gApp.ProjectFilename
      frmMain_OpenProjectSafely(HWND_FRMMAIN, wszText) 
   end if
      
   gApp.IsNewProjectFlag = false
   gApp.IsProjectLoading = FALSE

   frmExplorer_PositionWindows
   frmMain_PositionWindows
   
   ' Refresh the Project Explorer file name list
   LoadExplorerFiles()
   LoadFunctionsFiles()

   ' This will update the main window to show the project name in the window caption
   frmMain_SetFocusToCurrentCodeWindow
   
   Function = 0
End Function


' ========================================================================================
' Open a Project
' ========================================================================================
Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT

   ' Display the Open File Dialog
   Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, IDM_PROJECTOPEN)
   If pwszName Then 
      ' Pass the info to our generic project open function to handle everything.
      frmMain_OpenProjectSafely(HWND_FRMMAIN, *pwszName) 
      CoTaskMemFree(pwszName)
   End If
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Set the file type based on file extension
' ========================================================================================
function OnCommand_ProjectSetFileType( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

   ' This code also handles right click messages from the Explorer Treeview. 
   dim as CWSTR wszFileType

   select case id
      case IDM_SETFILENORMAL, IDM_SETFILENORMAL_EXPLORERTREEVIEW 
         wszFileType = FILETYPE_NORMAL
      case IDM_SETFILEMODULE, IDM_SETFILEMODULE_EXPLORERTREEVIEW
         wszFileType = FILETYPE_MODULE
      case IDM_SETFILEMAIN, IDM_SETFILEMAIN_EXPLORERTREEVIEW
         wszFileType = FILETYPE_MAIN
      case IDM_SETFILERESOURCE, IDM_SETFILERESOURCE_EXPLORERTREEVIEW
         wszFileType = FILETYPE_RESOURCE
      case IDM_SETFILEHEADER, IDM_SETFILEHEADER_EXPLORERTREEVIEW
         wszFileType = FILETYPE_HEADER
      case is > IDM_SETCATEGORY
         wszFileType = gConfig.Cat(id-IDM_SETCATEGORY).idFileType
   end select
   
   if pDoc = 0 then exit function

   if pDoc->ProjectFileType <> wszFileType then
      ' If attempting to set MAIN or RESOURCE then we need to move any existing
      ' pDoc that have that filetype over to the NORMAL branch.
      if (wszFileType = FILETYPE_MAIN) or (wszFileType = FILETYPE_RESOURCE) then
         dim pDoc2 as clsDocument ptr = iif(wszFileType = FILETYPE_MAIN, gApp.GetMainDocumentPtr, gApp.GetResourceDocumentPtr)
         if pDoc2 then
            pDoc2->ProjectFileType = FILETYPE_NORMAL
         end if   
      end if
      ' Set the new FileType for the currently selected document
      gApp.ProjectSetFileType( pDoc, wszFileType )
      LoadExplorerFiles()
      ' Highlight the selected tab file in the Explorer listbox
      frmExplorer_SelectItemData( pDoc )
      frmMain_SetStatusbar
   end if

   function = 0
end function


' ========================================================================================
' Remove file from the project
' ========================================================================================
function OnCommand_ProjectRemove( _
            byval id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT

   if pDoc = 0 then exit function
   if gTTabCtl.SetTabIndexByDocumentPtr( pDoc ) <> -1 then
     if OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT ) = false then exit function
   end if  

   ' File close succeeded. Remove this document from the project collection
   gdb2.dbDelete( pDoc->DiskFilename )
   gApp.RemoveDocument( pDoc )

   LoadExplorerFiles()

   if gTTabCtl.IsSafeIndex(gTTabCtl.CurSel) then
      frmExplorer_SelectItemData( gTTabCtl.tabs(gTTabCtl.CurSel).pDoc )
   end if

   function = 0
end function
