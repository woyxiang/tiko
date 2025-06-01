'    tiko editor - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General public License for more details.

#pragma once

'  Scintilla Control identifiers
#define IDC_SCINTILLA              100

' File encodings
#define FILE_ENCODING_ANSI         0
#define FILE_ENCODING_UTF8_BOM     1
#define FILE_ENCODING_UTF16_BOM    2

#define FILETYPE_UNDEFINED        "0"
#define FILETYPE_MAIN             "1"
#define FILETYPE_MODULE           "2"
#define FILETYPE_NORMAL           "3"
#define FILETYPE_RESOURCE         "4"
#define FILETYPE_HEADER           "5"


' Structure that holds all of the user embedded compiler directives
' in the source code. Currently, only the main source file is searched
' for the '#CONSOLE ON|OFF directive but others can be added as needed.
type COMPILE_DIRECTIVES
   DirectiveFlag as long              ' IDM_GUI, IDM_CONSOLE, IDM_RESOURCE
   DirectiveText as CWSTR             ' resource filename, link modules
end type

' Forward references
type clsDocument_ as clsDocument

' Enum used to distinguish bewteen a sub/function and Property Get/Set
enum ClassProperty
   None   = 0   ' sub/function
   Getter = 1
   Setter = 2
   ctor   = 3   ' constructor
   dtor   = 4   ' destructor
end enum

' type that holds data for all project files as it they are loaded from
' the project file.
type PROJECT_FILELOAD_DATA
   wszFilename    as CWSTR        ' full path and filename
   wszFiletype    as CWSTR        ' pDoc->ProjectFileType
   bLoadInTab     as boolean
   wszBookmarks   as CWSTR        ' pDoc->GetBookmarks()
   wszFoldPoints  as CWSTR        ' pDoc->GetFoldPoints()
   IsDesigner     as boolean
   IsDesignerView as long         
   nFirstLine     as long         ' first line of main view 
   nPosition      as long         ' current position of main view
   nFirstLine1    as long         ' first line of second view 
   nPosition1     as long         ' current position of second view
   nSplitPosition as long         ' pDoc->SplitY
   nFocusEdit     as long         ' View 0 or View 1
end type


type clsDocument
   private:
      ' 2 Scintilla direct pointers to accommodate split editing
      m_pSci(1)             as any ptr      
      m_hWndActiveScintilla as HWND
      m_UserModified        as boolean               ' user manually changes state outside of Scintilla changes
      
   public:
      pDocNext              as clsDocument_ ptr = 0  ' pointer to next document in linked list 
      IsNewFlag             as boolean
      LoadingFromFile       as boolean
      
      docData               as PROJECT_FILELOAD_DATA    ' loaded from project files
      
      ' 2 Scintilla controls to accommodate split editing
      ' hWindow(0) is our MAIN control (bottom)
      ' hWindow(1) is our split control (top)
      hWindow(1)            as HWND   ' Scintilla split edit windows 
      
      ' Code document related
      ProjectFiletype       as CWSTR = FILETYPE_UNDEFINED
      DiskFilename          as wstring * MAX_PATH
      AutoSaveFilename      as wstring * MAX_PATH    '#filename#
      AutoSaveRequired      as boolean
      DateFileTime          as FILETIME  
      bBookmarkExpanded     as boolean = true     ' Bookmarks list expand/collapse state
      bFunctionsExpanded    as boolean = true     ' Functions list expand/collapse state
      bHasFunctions         as boolean = false    ' FunctionList to determine if click will display the File
      FileEncoding          as long = FILE_ENCODING_ANSI       
      bNeedsParsing         as boolean            ' Document requires to be parsed due to changes.
      DeletedButKeep        as boolean            ' file no longer exists but keep open anyway
      DocumentBuild         as string             ' specific build configuration to use for this document
      sMatchWord            as string             ' for the incremental autocomplete search
      AutoCompletetype      as long               ' AUTOC_DIMAS, AUTOC_TYPE
      AutoCStartPos         as long
      AutoCTriggerStartPos  as long
      lastCaretPos          as long               ' used for checking in SCN_UPDATEUI
      lastXOffsetPos        as long               ' used for checking in SCN_UPDATEUI (horizontal offset)
      LastCharTyped         as long               ' used to test for BACKSPACE resetting the autocomplete popup.

      ' Following used for split edit views
      rcSplitButton         as RECT               ' Split gripper vertical for Scintilla window
      SplitY                as long               ' Y coordinate of vertical splitter
      bEditorIsSplit        as boolean
      bSizing               as boolean
      ptPrev                as POINT
      
      static NextFileNum as long
      
      declare property hWndActiveScintilla() as hwnd
      declare property hWndActiveScintilla(byval hWindow as hwnd)
      
      declare property UserModified() as boolean
      declare property UserModified( byval nModified as boolean )
      
      declare function ParseDocument() as boolean
      declare function IsValidScintillaID( byval idScintilla as long ) as boolean
      declare function GetActiveScintillaPtr() as any ptr
      declare function CreateCodeWindow( byval hWndParent as HWnd, byval IsNewFile as boolean, byval IsTemplate as boolean = False, byref wszFileName as wstring = "") as HWnd
      declare function FindReplace( byval strFindText as string, byval strReplaceText as string ) as long
      declare function InsertFile() as boolean
      declare function SaveFile(byval bSaveAs as boolean = False, byval bAutoSaveOnly as boolean = false) as boolean
      declare function ApplyProperties() as long
      declare function GetTextRange( byval cpMin as long, byval cpMax as long) as string
      declare function ChangeSelectionCase( byval fCase as long) as long 
      declare function GetCurrentLineNumber() as long
      declare function SelectLine( byval nLineNum as long ) as long
      declare function GetLine( byval nLine as long) as string
      declare function IsFunctionLine( byval lineNum as long ) as long
      declare function GotoNextFunction() as long
      declare function GotoPrevFunction() as long
      declare function GetLineCount() as long
      declare function GetSelText() as string
      declare function GetText() as string
      declare function SetText( byref sText as const string ) as long 
      declare function SetLine( byval nLineNum as long, byval sNewText as string) as long
      declare function AppendText( byref sText as const string ) as long 
      declare function CenterCurrentLine() as long 
      declare function GetSelectedLineRange( byref startLine as long, byref endLine as long, byref startPos as long, byref endPos as long ) as long 
      declare function BlockComment( byval flagBlock as boolean ) as long
      declare function CurrentLineUp() as long
      declare function CurrentLineDown() as long
      declare function MoveCurrentLines( byval flagMoveDown as boolean ) as long
      declare function NewLineBelowCurrent() as long
      declare function ToggleBookmark( byval nLine as long ) as long
      declare function NextBookmark() as long 
      declare function PrevBookmark() as long 
      declare function FoldToggle( byval nLine as long ) as long
      declare function FoldAll() as long
      declare function UnFoldAll() as long
      declare function FoldToggleOnwards( byval nLine as long) as long
      declare function ConvertEOL( byval nMode as long) as long
      declare function TabsToSpaces() as long
      declare function GetWord( byval curPos as long = -1 ) as string
      declare function GetBookmarks() as string
      declare function SetBookmarks( byval sBookmarks as string ) as long
      declare function GetFoldPoints() as string
      declare function SetFoldPoints( byval sFoldPoints as string ) as long
      declare function GetCurrentFunctionName( byref sFunctionName as string, byref nGetSet as ClassProperty ) as long
      declare function LineDuplicate() as long
      declare function SetMarkerHighlight() as long
      declare function RemoveMarkerHighlight() as long
      declare function IsMultiLineSelection() as boolean
      declare function HasMarkerHighlight() as boolean
      declare function FirstMarkerHighlight() as long
      declare function LastMarkerHighlight() as long
      declare function LinesPerPage( byval idxWindow as long ) as long
      declare function CompileDirectives( Directives() as COMPILE_DIRECTIVES) as long
      declare destructor
end type
dim clsDocument.NextFileNum as long = 0
