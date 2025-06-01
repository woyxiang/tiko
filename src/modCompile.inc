''    tiko editor - Programmer's Code Editor for the FreeBASIC Compiler
''    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
''
''    This program is free software: you can redistribute it and/or modify
''    it under the terms of the GNU General Public License as published by
''    the Free Software Foundation, either version 3 of the License, or
''    (at your option) any later version.
''
''    This program is distributed in the hope that it will be useful,
''    but WITHOUT any WARRANTY; without even the implied warranty of
''    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
''    GNU General Public License for more details.

''
''
''   modCompile.inc
''
''   Code here handles the entire compiling process. It prepares the necessary
''   command line data to pass to the applicable 32 or 64 version of the FBC
''   compiler. Compile time errors are parsed through analyis of the log file
''   and presented to the user via the Compiler Results listview and the Compiler
''   Log File text box.
''
''

#include once "modCompile.bi"
#include once "modCompileErrors.bi"
#include once "frmUserTools.bi"
#include once "frmOutput.bi"

Dim Shared gCompile As COMPILE_TYPE

''
'' Main module that handles the entire compile process
''
function code_Compile( ByVal wID As Long ) As boolean

   Dim pDoc As clsDocument Ptr
   
   Dim wszTemp      As WString * MAX_PATH 
   Dim wszFileExe   As WString * MAX_PATH
   dim wszCompile   as CWSTR
   dim wszCommand   as CWSTR 
   dim wszParams    as CWSTR
   dim sConsoleText as string 
    
   Dim As CWSTR wExeCmd, wDQ, wsLogSt, wst
         
   Dim As BOOLEAN fCompile32, fCompile64
   
   dim as long idxBuild, f, i, n, nCount
   
   wDQ = wchr(34) ' unicode double quotes
                             
   gCompile.wszFullCommandLine = ""
   gCompile.wszFullLogFile = ""
   gCompile.wszOutputMsg = ""
   gCompile.ResourceFile = ""
   gCompile.LinkModules = ""
   gCompile.bInvalidImagePath = false
   
   
   ' Indicate in the StatusBar that we have started to compile
   SetCompileStatusBarMessage( L(449, "Generating compiler files") & "...", 0 )
   

   ' Can only continue to compile if the primary source code file
   ' is not dirty and needs to be saved.
   if wID <> IDM_QUICKRUN THEN
      If gConfig.CompileAutosave Then
         ' If there any "New" files that have not been edited yet causing them to show 
         ' as dirty then we will mark them now so that the FileSaveAll command will
         ' ask to save it (otherwise a compile error will show).
         dim pDoc as clsDocument ptr = gApp.pDocList
         do until pDoc = 0
            if pDoc->IsNewFlag then pDoc->UserModified = true
            pDoc = pDoc->pDocNext
         loop
         If OnCommand_FileSaveAll(HWND_FRMMAIN) Then 
            SetCompileStatusBarMessage( "", 0 )
            Return True  ' user cancelled save
         end if   
      End If
   end if
   
   ' Get the array index for the current selected build (project or document)
   idxBuild = frmBuildConfig_getActiveBuildIndex()
   if idxBuild = -1 then idxBuild = 0
   
   ' Perform some pre-compile checks to see if we should continue.
   gConfig.FBWINCompiler32 = ProcessFromCurdriveApp(gConfig.FBWINCompiler32)
   gConfig.FBWINCompiler64 = ProcessFromCurdriveApp(gConfig.FBWINCompiler64)
   if gConfig.Builds(idxBuild).Is32bit then
      fCompile32 = True: fCompile64 = False
      gCompile.CompilerPath = gConfig.FBWINcompiler32
   end if
   if gConfig.Builds(idxBuild).Is64bit then
      gCompile.CompilerPath = gConfig.FBWINcompiler64
      fCompile32 = False: fCompile64 = True
   End if

   ' Convert relative path to absolute path if needed.
   if AfxPathIsRelative(gCompile.CompilerPath) then
      gCompile.CompilerPath = AfxPathCombine(AfxGetExePathName, gCompile.CompilerPath)
   END IF


   ' Check to see if the compiler exists
   If AfxFileExists(gCompile.CompilerPath) = 0 Then 
      wszTemp = L(202,"Invalid defined compiler path.")
      wszTemp = wszTemp & vbcrlf & "Path: " & gCompile.CompilerPath
      SetCompileStatusBarMessage( "", 0 )
      MessageBox( HWND_FRMMAIN, wszTemp, L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
   
   
   ' Set some compile flags depending on the type of compile requested.
   gCompile.CompileID = wID
   Select Case wID
      Case IDM_BUILDEXECUTE: gCompile.RunAfterCompile = True
      Case IDM_QUICKRUN:     gCompile.RunAfterCompile = True
      Case IDM_COMPILE:      gCompile.RunAfterCompile = False
   End Select
   
   
   Dim pDocMain As clsDocument Ptr 
   If gApp.IsProjectActive Then
      pDocMain = gApp.GetMainDocumentPtr()
      If pDocMain = 0 Then
         SetCompileStatusBarMessage( "", 0 )
         MessageBox( HWND_FRMMAIN, L(208,"No Main file specified for the project."), L(201,"Error"), _
                           MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
         Function = False: Exit Function
      End If
   Else
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
   End If
   If pDocMain = 0 Then 
      SetCompileStatusBarMessage( "", 0 )
      return 0
   end if
   
    ' Run any precompile tools
   for y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
      if gConfig.Tools(y).Action = USERTOOL_ACTION_PRECOMPILE then
         frmUserTools_ExecuteUserTool(y)
      end if   
   next

   ' If this is a QuickRun then we need to copy the current source code to a temporary bas file
   ' and compile that bas file instead. The resulting bas file will be deleted after the compiling
   ' has completed and exe will be deleted when the exe is terminated. 
   dim as CWSTR wszTempMainFilename
   if wID = IDM_QUICKRUN then
      if pDocMain->IsNewFlag then 
         gCompile.MainFolder = AfxGetExePathName
      else
         gCompile.MainFolder = AfxStrPathname("PATH", pDocMain->DiskFilename)
      end if
      wszTempMainFilename   = GetTemporaryFilename(gCompile.MainFolder, "bas")
      gCompile.MainFilename = wszTempMainFilename
      gCompile.MainName     = AfxStrPathname("NAME", gCompile.MainFilename)
      
      dim as CWSTR wszText
      wszText = pDocMain->GetText()

      dim as long f = freefile
      if open (gCompile.MainFilename for output as #f) = 0 then
         print #f, UnicodeToUtf8(wszText)
         close #f
      end if

   else
      gCompile.MainFilename = pDocMain->DiskFilename
      gCompile.MainName     = AfxStrPathname("NAME", pDocMain->DiskFilename)
      gCompile.MainFolder   = AfxStrPathname("PATH", pDocMain->DiskFilename)
   end if
   
   gCompile.ResourceFile = ""    ' default that there is no resource
   gCompile.CompileFlags = gConfig.Builds(idxBuild).wszOptions

   ' Iterate the optional semi colon separated list of compiler include paths
   gCompile.CompileIncludes = ""
   dim as CWSTR includePath
   dim as long numpaths = AfxStrParseCount(gConfig.CompilerIncludes, ";")
   for i as long = 1 to numpaths
      includePath = rtrim(AfxStrParse(gConfig.CompilerIncludes, i, ";"))
      if len(includePath) then 
         gCompile.CompileIncludes = gCompile.CompileIncludes + " -i " + chr(34) + includePath + chr(34)
      end if
   next

   If gApp.IsProjectActive Then
      Dim pDocResource As clsDocument Ptr = gApp.GetResourceDocumentPtr()
      if pDocResource then gCompile.ResourceFile = pDocResource->DiskFilename

      gCompile.ObjFolder = gCompile.MainFolder & APPEXTENSION & "\"
      gCompile.ObjFolderShort = ".\" & APPEXTENSION & "\"

      If fCompile32 Then 
         gCompile.ObjID = "32.o"
         gCompile.CompileFlags = gCompile.CompileFlags + " " + gApp.ProjectOther32 +" " + gCompile.CompileIncludes          
      End If
      If fCompile64 Then 
         gCompile.ObjID = "64.o"
         gCompile.CompileFlags = gCompile.CompileFlags + " " + gApp.ProjectOther64 + " " + gCompile.CompileIncludes          
      End If   
      
      ' Make sure the folders exist
      SHCreateDirectory( 0, gCompile.ObjFolder )

   Else
      ' No active project
      gCompile.CompileFlags = gCompile.CompileFlags + " " + gConfig.CompilerSwitches + " " + gCompile.CompileIncludes  
   End If


   ' Search main source code for any user embedded compile directives. These will override
   ' anything that was set at the default or project level.
   redim directives(any) as COMPILE_DIRECTIVES
   pDocMain->CompileDirectives(directives())
   
   for i as long = lbound(directives) to ubound(directives)
      select case directives(i).DirectiveFlag 
         case IDM_CONSOLE
            n = instr(gCompile.CompileFlags, " -S GUI ")
            if n then gCompile.CompileFlags = AfxStrDelete(gCompile.CompileFlags, n, 8)
            gCompile.CompileFlags = gCompile.CompileFlags + " -s console "
         case IDM_GUI
            n = instr(gCompile.CompileFlags, " -S CONSOLE ")
            if n then gCompile.CompileFlags = AfxStrDelete(gCompile.CompileFlags, n, 12)
            gCompile.CompileFlags = gCompile.CompileFlags + " -s gui "
         case IDM_RESOURCE
            gCompile.ResourceFile = directives(i).DirectiveText
         case IDM_LINKMODULES
            gCompile.LinkModules = gCompile.LinkModules & " " & directives(i).DirectiveText
      end select
   next
print "before CompileDirectives: "; gCompile.LinkModules
   gCompile.LinkModules = AddPathsToLinkModules(gCompile.LinkModules)
print "after AddPathsToLinkModules: "; gCompile.LinkModules


   ' Need to determine the output filename in order to search for any existing
   ' running process.
   gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".exe"
   If Instr(" " & Ucase(gCompile.CompileFlags), " -DLL") Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), " -DYLIB") Then
      gCompile.OutputFilename = gCompile.MainFolder & gCompile.MainName & ".dll"
   End If
   If Instr(" " & Ucase(gCompile.CompileFlags), " -LIB") Then
      gCompile.OutputFilename = gCompile.MainFolder & "lib" & gCompile.MainName & ".a"
   End If


   ' Need to test if the resulting application to be compiled is actually running
   ' in memory. This would cause the compile to fail.
   If IsProcessRunning(@gCompile.OutputFilename) Then
      SetCompileStatusBarMessage( "", 0 )
      MessageBox( HWND_FRMMAIN, L(200,"Program running") & "...", L(201,"Error"), _
                        MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 )
      Function = False: Exit Function
   End If
      

   gApp.IsCompiling = true    

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORWAIT, 0 )


   ' Change to the output code directory
   ChDir gCompile.MainFolder
   
   gCompile.StartTime = Timer
   

   ' If this is a project then we need to compile all of the modules first
   ' in order to create the necessary *.o object files
   If gApp.IsProjectActive Then
       gCompile.CompileFlags = " " & gCompile.CompileFlags & " "
      dim as string g_flag = iif(instr(gCompile.CompileFlags, " -g "), " -g -nostrip ", "")
      dim as string e_flag = iif(instr(gCompile.CompileFlags, " -e "), " -e ", "")
      dim as string ex_flag = iif(instr(gCompile.CompileFlags, " -ex "), " -ex ", "")
      dim as string exx_flag = iif(instr(gCompile.CompileFlags, " -exx "), " -exx ", "")
      
      dim as long nFileCount   ' used to calculate when to fire an AfxDoEvents
      pDoc = gApp.pDocList
      do until pDoc = 0
         If pDoc->ProjectFileType = FILETYPE_MODULE Then
            ' Get the base name of the file for constructing the object filename
            wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
            wszTemp = gCompile.ObjFolderShort & wszTemp & gCompile.ObjID
            
            ' Compare the source code file datetime to the object. If the source code
            ' date time is greater then we need to recompile it, otherwise we will simply
            ' link to the existing object file.
            Dim As FILETIME ft1 = AfxGetFileLastWriteTime(pDoc->DiskFilename) ' source file
            Dim As FILETIME ft2 = AfxGetFileLastWriteTime(wszTemp)  ' object file

            If (AfxFileTimeToVariantTime(ft1) > AfxFileTimeToVariantTime(ft2)) or _
               (wID = IDM_REBUILDALL)  Then

               nFileCount = nFileCount + 1
               if (nFileCount mod 5) = 0 then AfxDoEvents
               
               wst = L(203, "Compiling") + "... " + AfxStrPathname("NAMEX", pDoc->DiskFilename) 
               SetCompileStatusBarMessage( wst, 0 )
 
               wszCommand = gCompile.CompilerPath
               wszParams = g_flag & e_flag & ex_flag & exx_flag &  _
                           gCompile.CompileIncludes + " " + _
                           " -c -b " + wDQ + pDoc->DiskFilename + wDQ + _
                           " -v -o " + wDQ + wszTemp + wDQ 
               RedirConsoleToFile( wszCommand, wszParams, sConsoleText )
               gCompile.wszFullCommandLine = wszCommand & " " & wszParams
               gCompile.wszFullLogFile = sConsoleText

               ' Need to check to see if any errors occurred while compiling this
               ' object file. If yes, then stop processing object files and display
               ' the error.
               if ParseLogForError(sConsoleText, false, wID, fCompile64, true) then   ' error detected
                  gApp.IsCompiling = false
                  SetCursor( hCurSave )
                  ResetScintillaCursors()
                  return true
               end if
               sConsoleText = ""
            End If
            
         End If
         pDoc = pDoc->pDocNext
      loop
   End If

   
   ' Compile the Main file. If this is a project then we also need to link
   ' in all of the *.o object files.
   ' Per fxm post https://www.freebasic.net/forum/viewtopic.php?f=8&p=266126#p266126
   ' and https://www.freebasic.net/forum/viewtopic.php?f=3&t=28723&p=274431#p274431
   ' Do not specify file extension with -m parameter.
   ' So the main module file must be called twice in the command line:
   ' - after compile option '-m', but without specified extension,
   ' - and also like any module to compile, but there with its specified extension.
   'dim as CWSTR wszMfile = wDQ + AfxStrPathName("PATH", gCompile.MainFilename) + AfxStrPathName("NAME", gCompile.MainFilename) + wDQ
   'dim as CWSTR wszMfile = AfxStrPathName("NAME", gCompile.MainFilename) 
   'wszParams = "-m " + wszMfile + " " + wDQ + gCompile.MainFilename + wDQ 
   wszParams = "-m " + wDQ + gCompile.MainFilename + wDQ 


   if len(gCompile.LinkModules) then
      wszParams = wszParams + " " + gCompile.LinkModules
   end if

   if len(gCompile.ResourceFile) then  
      wszParams = wszParams + " " + wDQ + gCompile.ResourceFile + wDQ
   end if  
   
   
   ' Ensure verbose compiler message output
   wszParams = wszParams + " -v " + gCompile.CompileFlags 
   
   f = instr(wszParams, " -x ") 
   if f = 0 then
      wszParams = wszParams + " -x " + wDQ + gCompile.OutputFilename + wDQ
   else
      ' The ParseLogForError() function below will identify the correct
      ' filename for gCompile.OutputFilename based on the man file being linked.
   end if
   
   If gApp.IsProjectActive Then
      pDoc = gApp.pDocList
      do until pDoc = 0
         If pDoc->ProjectFileType = FILETYPE_MODULE Then
            ' Get the base name of the file for constructing the object filename
            wszTemp = AfxStrPathname("NAME", pDoc->DiskFilename)
            wszParams = wszParams + " " + gCompile.ObjFolderShort + wszTemp + gCompile.ObjID
         End If
         pDoc = pDoc->pDocNext
      loop
   End If

   ' Attempt to delete any existing output file so it will not exist
   ' should the compile fail.
   if AfxFileExists( gCompile.OutputFilename ) then
      AfxDeleteFile( gCompile.OutputFilename )
   end if

   wst = L(203, "Compiling") + "... " + AfxStrPathname("NAMEX", gCompile.MainFilename) 
   SetCompileStatusBarMessage( wst, 0 )
   SetCursor( LoadCursor(0, IDC_WAIT) )


   sConsoleText = ""
   wszCommand = gCompile.CompilerPath
   RedirConsoleToFile( wszCommand, wszParams, sConsoleText )
   gCompile.wszFullCommandLine = wszCommand & " " & wszParams
   gCompile.wszFullLogFile = sConsoleText
               

   ' Attempt to fix the random problem whereby the spinning mouse wait cursor is not reset
   ' to an arrow. It seems to happen in random cases whereby an error is thrown and a new 
   ' document is loaded to position to the error line. Very hard to reproduce. We will manually
   ' reset all Scintilla cursors here for all open documents. Maybe this will help.
   SetCursor( hCurSave )
   ResetScintillaCursors()
   
   gCompile.EndTime = Timer
   gCompile.SystemTime = AfxSystemSystemTime 

   ' If this was a QuickRun then delete the temporary code file that was created.
   ' Both the temporary disk file and (if applicable) the VD MAIN CODEGEN file.
   if wID = IDM_QUICKRUN THEN 
      AfxDeleteFile( gCompile.MainFilename )
      AfxDeleteFile( wszTempMainFilename )
   end if    
   
   ' Add any temp resource file to the list of temp files to later be deleted.
   gApp.AddQuickRunEXE(gCompile.TempResourceFile)
      
   if ParseLogForError(sConsoleText, true, wID, fCompile64, false) = true then 
      ' Error was found. Editor has now been set to the error position. Nothing
      ' more we can do now but exit out. 
      ' Make sure gApp.IsCompiling = false before returning
   else
      ' Run any postcompile tools
      for y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
         if gConfig.Tools(y).Action = USERTOOL_ACTION_POSTCOMPILE then
            frmUserTools_ExecuteUserTool(y)
         end if   
      next
      
      If gCompile.RunAfterCompile Then 
         ' If gCompile.OutputFilename is a relative filename then convert it to full filename.
         gCompile.OutputFilename = AfxStrReplace( gCompile.OutputFilename, "/", "\" )
         if AfxPathIsRelative( gCompile.OutputFilename ) then 
            gCompile.OutputFilename = AfxPathCombine( gCompile.MainFolder, gCompile.OutputFilename )
         end if

         RunEXE( _
            gCompile.OutputFilename,    _
            iif( gApp.IsProjectActive, gApp.ProjectCommandLine, gApp.wszCommandLine ) _
            )
      End If
   
   end if
   
   gApp.IsCompiling = false

   Function = True   ' successful
End Function


''
''
''
function AddPathsToLinkModules( byval modules as CWSTR ) as CWSTR
   ' SEARCH ORDER
   ' (1) Current folder
   ' (2) User defined Include paths
   ' (3) \inc folder for the selected compiler toolchain

   dim as CWSTR wszResult  ' the final fully path qualified list of space separated module filenames
   dim as CWSTR wszTemp
   
   dim as CWSTR wszCurrentFolder = AfxStrPathname("PATH", gCompile.MainFilename)
   dim as CWSTR wszCompilerIncFolder = AfxStrPathname("PATH", gCompile.CompilerPath) & "inc\"
   
   dim as CWSTR wszLinkModules = AfxStrShrink(gCompile.LinkModules)  
   
   dim as long nummodules = AfxStrParseCount(wszLinkModules, " ")
   for i as long = 1 to nummodules
      dim as CWSTR modulename = trim(AfxStrParse(wszLinkModules, i, " "))
      ' Remove any enclosing double quotes from the filename
      modulename = AfxStrRemove(modulename, chr(34))

      if len(modulename) = 0 then continue for

      ' (1) Current folder
      wszTemp = wszCurrentFolder & modulename
      if AfxFileExists(wszTemp) then
         wszResult = wszResult & chr(34) & wszTemp & chr(34) & " "
         continue for
      end if
   
      ' (2) User defined Include paths
      ' Iterate the optional semi colon separated list of compiler include paths
      gCompile.CompileIncludes = ""
      dim as CWSTR includePath
      dim as long numpaths = AfxStrParseCount(gConfig.CompilerIncludes, ";")
      for i as long = 1 to numpaths
         includePath = rtrim(AfxStrParse(gConfig.CompilerIncludes, i, ";"))
         ' Remove any enclosing double quotes from the includepath
         includePath = AfxStrRemove(includePath, chr(34)) & "\"   ' ensure trailing backslash
         if len(includePath) then 
            wszTemp = includePath & modulename
            if AfxFileExists(wszTemp) then
               wszResult = wszResult & chr(34) & wszTemp & chr(34) & " "
               continue for: continue for
            end if
         end if
      next
   
      ' (3) \inc folder for the selected compiler toolchain
      wszTemp = wszCompilerIncFolder & modulename
      if AfxFileExists(wszTemp) then
         wszResult = wszResult & chr(34) & wszTemp & chr(34) & " "
         continue for
      end if
   next
   
   function = trim(wszResult)
end function

