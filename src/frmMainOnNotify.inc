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

#include once "frmMainOnNotify.bi"

#include once "clsApp.bi"
#include once "clsDocument.bi"
#include once "frmMain.bi"


' ========================================================================================
' Handle Character Autocompletion if that editor option is active
' ========================================================================================
function frmMain_InsertMatchedChars( _
            byval pDoc as clsDocument ptr, _
            byval ch as long _
            ) as boolean
    
    if gConfig.CharacterAutoComplete = 0 then exit function
    
    if pDoc = 0 then exit function
    
    dim as any ptr pSci = pDoc->GetActiveScintillaPtr()
    if pSci = 0 then exit function
    
    dim as long caretPos = SciMsg( pSci, SCI_GETCURRENTPOS, 0, 0)
    dim as long docStart = caretPos = 1
    dim as long docend   = caretPos = SciMsg( pSci, SCI_GETTEXTLENGTH, 0, 0)

    ' Get the styling of the current line to determine if we are in a 
    ' multiline or single line comment block then abort the autoinsert.
    select case SciMsg( pSci, SCI_GETSTYLEAT, caretPos, 0)
        case SCE_B_MULTILINECOMMENT, SCE_B_COMMENT
            exit function
    end select 

    dim as long nCharPrev = iif( docStart, _
                        SciMsg( pSci, SCI_GETCHARAT, caretPos, 0), _
                        SciMsg( pSci, SCI_GETCHARAT, caretPos-2, 0) )
    
    dim as long nCharNext = SciMsg( pSci, SCI_GETCHARAT, caretPos, 0)

    dim as boolean isCharPrevBlank = iif( instr(chr(nCharPrev), any chr(32, 9, 10, 13)), true, false )

    dim as boolean isCharNextBlank = iif( instr(chr(nCharNext), any chr(32, 9, 10, 13)), true, false ) 
    if nCharNext = docend then isCharNextBlank = true

    dim as boolean isEnclosed = false
    if (nCharPrev = asc("(") and nCharNext = asc(")") ) or _
        (nCharPrev = asc("{") and nCharNext = asc("}") ) or _
        (nCharPrev = asc("[") and nCharNext = asc("]") ) then
        isEnclosed = true
    end if                      
                            
    dim as boolean isSpaceEnclosed = false
    if (nCharPrev = asc("(") and isCharNextBlank ) or _
        (isCharPrevBlank and nCharNext = asc(")") ) or _
        (nCharPrev = asc("{") and isCharNextBlank)  or _
        (isCharPrevBlank and nCharNext = asc("}") ) or _
        (nCharPrev = asc("[") and isCharNextBlank)  or _
        (isCharPrevBlank and nCharNext = asc("]") ) then
        isSpaceEnclosed = true
    end if   

    dim as boolean isCharOrString = false
    if (isCharPrevBlank and isCharNextBlank) or isEnclosed or isSpaceEnclosed then
        isCharOrString = true
    end if

    dim as boolean charNextIsCharOrString = iif( instr(chr(nCharNext), any chr(34,92)), true, false )

    dim as wstring * 10 wszText
        
    select case chr(ch)
        case "("
            if (charNextIsCharOrString) then return false
            wszText = ")"
            SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
            
        case "{"
            if (charNextIsCharOrString) then return false
            wszText = "}"
            SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
            
        case "["
            if (charNextIsCharOrString) then return false
            wszText = "]"
            SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
            
        case chr(34)
            '// 0x22 = "
            if (nCharPrev = 34) and (nCharNext = 34) then
            SciMsg( pSci, SCI_DELETERANGE, caretPos, 1)
            SciMsg( pSci, SCI_GOTOPOS, caretPos, 0)
            return false
            end if

            if (isCharOrString) then
            wszText = chr(34)
            SciMsg( pSci, SCI_INSERTTEXT, caretPos, cast(LPARAM, @wszText) )
            end if    
        
    end select 

    function = true
end function


' ========================================================================================
' This sub finds the position of a brace at the current caret position
' and the position of its matching pair.
' ========================================================================================
sub frmMain_findMatchingBracePos( _
            byval pDoc as clsDocument ptr, _
            byref braceAtCaret as integer, _
            byref braceOpposite as integer _
            )
    
    dim as integer caretPos = SciExec( pDoc->hWndActiveScintilla, SCI_GETCURRENTPOS, 0, 0)

    braceAtCaret = -1
    braceOpposite = -1
    dim as integer charBefore = 0

    dim as integer lengthDoc = SciExec( pDoc->hWndActiveScintilla, SCI_GETLENGTH, 0, 0)

    if (lengthDoc > 0) andalso (caretPos > 0) then
        charBefore = SciExec( pDoc->hWndActiveScintilla, SCI_GETCHARAT, caretPos - 1, 0)
    end if

    ' Priority goes to the character before the caret
    if (charBefore <> 0) andalso instr("()[]{}", chr(charBefore)) then
        braceAtCaret = caretPos - 1
    end if

    if (lengthDoc > 0) andalso (braceAtCaret < 0) then
        ' No brace found, so check the character after the caret
        dim as integer charAfter = SciExec( pDoc->hWndActiveScintilla, SCI_GETCHARAT, caretPos, 0)
        if (charAfter <> 0) andalso instr("()[]{}", chr(charAfter)) then
            braceAtCaret = caretPos
        end if
    end if

    if braceAtCaret >= 0 then
        braceOpposite = SciExec( pDoc->hWndActiveScintilla, SCI_BRACEMATCH, braceAtCaret, 0)
    end if

end sub


' ========================================================================================
' This function performs brace matching and highlighting.
' It returns true if a brace (matched or unmatched) is found near the caret.
' ========================================================================================
function frmMain_DoBraceMatching( byval pDoc as clsDocument ptr ) as boolean

    if gConfig.BraceHighlight = false then
        SciExec( pDoc->hWndActiveScintilla, SCI_BRACEHIGHLIGHT, -1, -1 )
        return false
    end if         

    dim as integer braceAtCaret = -1
    dim as integer braceOpposite = -1

    frmMain_findMatchingBracePos( pDoc, braceAtCaret, braceOpposite )

    if (braceAtCaret <> -1) andalso (braceOpposite = -1) then
        ' Found a brace, but it's unmatched
        SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETFORE, 9, ghEditor.ForeColorBracebad )
        SciExec( pDoc->hWndActiveScintilla, SCI_BRACEBADLIGHT, braceAtCaret, 0 )
        '_pEditView->execute(SCI_SETHIGHLIGHTGUIDE, 0)
    else
        ' Found a matched pair or no brace at all
        SciExec( pDoc->hWndActiveScintilla, SCI_INDICSETFORE, 9, ghEditor.ForeColorBracegood )
        SciExec( pDoc->hWndActiveScintilla, SCI_BRACEHIGHLIGHT, braceAtCaret, braceOpposite )

'        if _pEditView->isShownIndentGuide() then
'            dim columnAtCaret as integer ptr = _pEditView->execute(SCI_GETCOLUMN, braceAtCaret)
'            dim columnOpposite as integer ptr = _pEditView->execute(SCI_GETCOLUMN, braceOpposite)
'            _pEditView->execute(SCI_SETHIGHLIGHTGUIDE, iif(columnAtCaret < columnOpposite, columnAtCaret, columnOpposite))
'        end if
    end if

'    const enable as boolean = (braceAtCaret <> -1) andalso (braceOpposite <> -1)
'    enableCommand(IDM_SEARCH_GOTOMATCHINGBRACE, enable, MENU or TOOLBAR)
'    enableCommand(IDM_SEARCH_SELECTMATCHINGBRACES, enable, MENU)

    return (braceAtCaret <> -1)

end function


' ========================================================================================
' Process Scintilla Notifications
' ========================================================================================
function Scintilla_OnNotify( _
            byval hwnd as HWND, _
            byval pNSC as SCNOTIFICATION ptr _
            ) as long

    if pNSC = 0 then exit function
    
    dim pDoc as clsDocument ptr 

    dim as HWND hEdit
    dim as long nLine, nFoldLevel

    select case pNSC->hdr.code
            
        case SCN_UPDATEUI    
            if gApp.SuppressNotify then return true
            pDoc = gApp.GetDocumentPtrByWindow( pNSC->hdr.hwndFrom )
            if pDoc then 
                if pDoc->IsValidScintillaID( pNSC->hdr.idFrom ) then
                    frmMain_SetStatusbar
                                    
                    ' Update the vertical scrollbars because mousewheel may have been
                    ' used to scroll a split window without keyboard focus.
                    if pDoc->EditorSplitMode <> SplitNone then
                        frmEditorVScroll_calcVThumbRect(pDoc)
                        AfxRedrawWindow( HWND_FRMEDITOR_VSCROLLBAR(0) )
                        AfxRedrawWindow( HWND_FRMEDITOR_VSCROLLBAR(1) )
                    end if    

                    '// Has the caret changed position
                    var caretPos = SciExec( pDoc->hWndActiveScintilla, SCI_GETCURRENTPOS, 0, 0)

                    if pDoc->lastCaretPos <> caretPos then
                        pDoc->lastCaretPos = caretPos

                        '// If the xOffset has changed then we need to update the horizontal scrollbar
                        dim as long GetXOffset = SciExec( pDoc->hWndActiveScintilla, SCI_GETXOFFSET, 0, 0 )
                        if pDoc->lastXOffsetPos <> GetXOffset then
                            pDoc->lastXOffsetPos = GetXOffset
                            AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(0) )
                            AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(1) )
                        end if

                        '// Update any occurrence highlights
                        if gConfig.OccurrenceHighlight then
                            static as string sCurWord, sPrevWord
                            sCurWord = pDoc->GetWord()
                            if sCurWord <> sPrevWord then
                                frmMain_HighlightWord( pDoc, sCurWord )
                                sPrevWord = sCurWord
                            end if
                        end if
                        
                        '// Update any brace highlighting
                        frmMain_DoBraceMatching( pDoc )

                    end if
                    
                end if
            end if
            return true
            
        
        case SCN_MODIFYATTEMPTRO
            ' Attempting to modify a read-only document.
            MessageBeep(MB_ICONWARNING)
                    
                    
        case SCN_MODIFIED
            if gApp.SuppressNotify then return true

            ' Show line and column. Only do on modification of text otherwise we will have
            ' a huge slowdown when notifications sent for UI updates.
            
            if (pNSC->modificationType and SC_MOD_INSERTTEXT) orelse _
            (pNSC->modificationType and SC_MOD_DELETETEXT) then

                ' Set parsing flag to true. We use a parse flag rather than depending on the dirty
                ' flags because the file could stay dirty for a long time before it gets saved and 
                ' every call to the parser would happen even in cases where only directional 
                ' movement has occured with no text modifications at all.
                
                ' Get the pDoc from the Scintilla HWND because this is the most reliable
                ' method of knowing which scintilla window sent the notification (rather than
                ' using gTTabCtl.GetActiveDocumentPtr. Possible that a control is sending the message
                ' as it is loading and is not yet the active document ptr.
                pDoc = gApp.GetDocumentPtrByWindow( pNSC->hdr.hwndFrom )
                if pDoc then 
                    if pDoc->LoadingFromFile = false then
                        pDoc->AutoSaveRequired = true
                    end if   
                    frmMain_SetStatusbar

                    pDoc->bNeedsParsing = true
                    if (pNSC->modificationType and SC_MOD_DELETETEXT) then
                        ' Do a check to see if we have backspaced to a point where there is a
                        ' period to try to popup an autocomplete or codetip.
                        dim as long nPos, nChar
                        hEdit = pDoc->hWndActiveScintilla
                        nPos  = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0 ) - 1
                        nChar = SciExec( hEdit, SCI_GETCHARAT, nPos, 0 )
                        select case chr(nChar)
                            case ".", ">"    '  Show autocomplete list for TYPE variables
                                ' dot "." or ">"  part of a pointer
                                ' Need to PostMessage in order to give time for notification to complete. The "."
                                ' is also used as a character that selects and closes the popup list. If no
                                ' PostMessage then the "." will automatically select the first entry in the list
                                ' and close.
                        pDoc->AutoCTriggerStartPos = nPos
                                pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                                PostMessage( HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, nChar, 0 )
                        end select 
                    end if
                end if
            end if
            return true
            
        case SCN_MARGINCLICK
            ' Folder margin
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            if pDoc then 
                hEdit      = pDoc->hWndActiveScintilla
                nLine      = SciExec(hEdit, SCI_LINEFROMPOSITION, pNSC->position, 0)
                nFoldLevel = SciExec(hEdit, SCI_GETFOLDLEVEL, nLine, 0)
                select case pNSC->margin 
                    case 1   ' left margin (bookmarks)
                        pDoc->ToggleBookmark(nLine)
                        LoadBookmarksFiles
                        AfxRedrawWindow( HWND_FRMBOOKMARKS )
                    case 2   ' fold margin
                        ' if is the head line...
                        if (nFoldLevel and SC_FOLDLEVELHEADERFLAG) <> 0 then
                            pDoc->FoldToggle(nLine) 
                        end if   
                end select 
            end if           
                        
        
        case SCN_AUTOCCANCELLED
            ' Destroy the popup information window
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            if pDoc then 
                ' If the last typed character was a BACKSPACE then do NOT reset the 
                ' autocomplete type otherwise the autocomplete popup will not continue
                ' to popup on subsequent characters being added.
                if pDoc->LastCharTyped = VK_BACK then
                    ' backspace pressed... do nothing, unless we are backspacing past where
                    ' the popup trigger occurred.
                    dim as long curPos = SciExec(pDoc->hWndActiveScintilla, SCI_GETCURRENTPOS, 0, 0) 
                    if curPos <= pDoc->AutoCTriggerStartPos then
                        pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    end if
                else
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                end if   
            end if

            
        case SCN_AUTOCSELECTION
            ' A selection occured from the AutoComplete listbox. We do the insertion
            ' ourselves because the startpos of the word is not the same as the start position
            ' used by the autocomplete listbox (because we are doing incremental searches).
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            
            if pDoc then 
                hEdit = pDoc->hWndActiveScintilla
                
                ' Get the position where the listbox was opened.
                dim as long nStartPos = pDoc->AutoCStartPos 
                dim as long nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
                
                ' Get the match word that existed when the listbox was displayed
                dim as long nLenMatchWord = len(pDoc->sMatchWord)
                
                ' Set the word that was selected in the listbox
                dim as string sText = *cast(zstring ptr, pNSC->lpText)
                
                if ucase(sText) <> ucase(pDoc->sMatchWord) then
                    SciExec( hEdit, SCI_SETSEL, nStartPos-nLenMatchWord, nPos)
                    SciExec( hEdit, SCI_REPLACESEL, 0, cast(LPARAM, strptr(sText)))
                    nPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
                end if
                SciExec( hEdit, SCI_SETSEL, nPos, nPos)
                    
                ' Now that we have inserted our own text, cancel the autoinsertion by
                ' the autocomplete listbox.
                SciExec( hEdit, SCI_AUTOCCANCEL, 0, 0)
                pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                                
                if chr(pNSC->ch) = "=" then
                    ' Add a space before and after Equals sign whenever the user presses
                    ' the equals sign to terminate ab autocomplete popup list.
                    PostMessage(hwnd, MSG_USER_APPENDEQUALSSIGN, 0, 0)
                end if
                
                ' if a CodeTip was displayed prior to the autocomplete popup list then
                ' redisplay that CodeTip now.
                ShowCodetip( pDoc )
                
                return true   
            end if
    

        case SCN_AUTOCCHARDELETED    
            if gConfig.CodeTips then
                ' User deleted a character while autocompletion list was active. Display
                ' new contents of list because the underlying word being typed as changed.
                ShowAutocompleteList(SCN_AUTOCCHARDELETED)
            end if
                                    
        
        case SCN_CHARADDED
            if gApp.SuppressNotify then exit function
            pDoc = gTTabCtl.GetActiveDocumentPtr()
            if pDoc = 0 then exit function
            hEdit = pDoc->hWndActiveScintilla
        
			pDoc->bNeedsParsing = true

            ' Attempt to do a Character Autocompletion if that editor option is active.
            frmMain_InsertMatchedChars( pDoc, pNSC->ch )
            
            select case chr(pNSC->ch)
                case ".", ">"    '  Show autocomplete list for TYPE variables
                    ' dot "." or ">"  part of a pointer
                    ' Need to PostMessage in order to give time for notification to complete. The "."
                    ' is also used as a character that selects and closes the popup list. if no
                    ' PostMessage then the "." will automatically select the first entry in the list
                    ' and close.
                    pDoc->AutoCTriggerStartPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    if PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function
            
                case ")"      ' Close and active code tip 
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    SciExec( hEdit, SCI_CALLTIPCANCEL, 0, 0 )
                
                case "(", ","      ' Show code tip 
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    if PostMessage(HWND_FRMMAIN, MSG_USER_SHOWAUTOCOMPLETE, pNSC->ch, 0) then exit function

                case chr(13)  ' ENTER KEY PRESSED
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    AttemptAutoInsert()
        
                case chr(32)  ' Space key pressed (Insert AutoComplete constructs if applicable)
                    pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
                    if ShowAutocompleteList() then exit function

                case else   ' all other letters
                    if ShowAutocompleteList() then exit function

            end select 
            
    end select 

    function = 0
end Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
function frmMain_OnNotify( _
            byval hwnd as HWND, _
            byval id as long, _
            byval pNMHDR as NMHDR ptr _
            ) as LRESULT

    dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
    if pDoc then 
        if pDoc->IsValidScintillaID(id) then
            ' Process Scintilla control notification messages. First, check to see if the notifications
            ' have been suppressed by a bulk modification activity that should only update the screen
            ' at the end of its process (eg. moving large numbers of selected lines up or down).
            if gApp.SuppressNotify then exit function
            Scintilla_OnNotify hwnd, cast(SCNOTIFICATION ptr, pNMHDR)
            exit function
        end if
    end if

    function = 0
end Function


'