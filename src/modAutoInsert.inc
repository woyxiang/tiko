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

''
''  Attempt to insert the second half of a syntax construct. for example,
''  The END FUNCTION, LOOP, next, ...  This is activated when ENTER is pressed
''  in the code editor and gConfig. AutoComplete is TRUE.
''

#include once "modAutoInsert.bi"
#include once "frmFunctions.bi"


' ========================================================================================
' Determine if end of the block statement exists.
' ========================================================================================
private function CanCompleteBlockStatement( byval pDoc as clsDocument ptr, _
                                            byval idBlockType as long, _
                                            byval nStartLine as long _
                                            ) as boolean
    
    if gConfig.AutoComplete = false then return false
    
    '// January 15, 2023. Default to return True always thereby bypassing the logic
    '// of this function. Rather than "smartly" trying to determine if a matching end
    '// keyword exists, we will simply honor the user's wishes and insert the completion
    '// keywords.
    return true
    
    
    dim as long NumLines, nLenMatch
    dim as string sStartMatch, sEndMatch, sLine
    
    dim as hwnd hEdit = pDoc->hWndActiveScintilla

    ' Add a unique terminating delimiter to the phrase in order to make it
    ' easier to match lines later that only contain the matching end phrase
    ' and no following space character (ie. they exist alone on the line).
    ' The algorithm needs to search for a matching end phrase but if it first
    ' finds a matching start phrase for the block type 
    select case idBlockType
        case BLOCK_STATEMENT_IF:          sStartMatch = "if ":     sEndMatch = "end if|"
        case BLOCK_STATEMENT_FOR:         sStartMatch = "for ":    sEndMatch = "next|"
        case BLOCK_STATEMENT_SELECT:      sStartMatch = "SELECT ": sEndMatch = "END SELECT|"
        case BLOCK_STATEMENT_WHILE:       sStartMatch = "WHILE ":  sEndMatch = "WEND|"
        case BLOCK_STATEMENT_DO:          sStartMatch = "DO ":     sEndMatch = "LOOP|"
        case BLOCK_STATEMENT_FUNCTION:    sStartMatch = "":        sEndMatch = "END FUNCTION|"
        case BLOCK_STATEMENT_SUB:         sStartMatch = "":        sEndMatch = "END SUB|"
        case BLOCK_STATEMENT_PROPERTY:    sStartMatch = "":        sEndMatch = "END PROPERTY|"
        case BLOCK_STATEMENT_OPERATOR:    sStartMatch = "":        sEndMatch = "END OPERATOR|"
        case BLOCK_STATEMENT_TYPE:        sStartMatch = "TYPE ":   sEndMatch = "END TYPE|"
        case BLOCK_STATEMENT_WITH:        sStartMatch = "WITH ":   sEndMatch = "END WITH|"
        case BLOCK_STATEMENT_ENUM:        sStartMatch = "ENUM ":   sEndMatch = "END ENUM|"
        case BLOCK_STATEMENT_UNION:       sStartMatch = "UNION ":  sEndMatch = "END UNION|"
        case BLOCK_STATEMENT_CONSTRUCTOR: sStartMatch = "":        sEndMatch = "END CONSTRUCTOR|"
        case BLOCK_STATEMENT_DESTRUCTOR:  sStartMatch = "":        sEndMatch = "END DESTRUCTOR|"
    end select
    
    NumLines = SciExec( hEdit, SCI_GETLINECOUNT, 0, 0) 
    nStartLine = nStartLine + 1
    
    for i as long = nStartLine to NumLines - 1
        ' Does this line start with the matching phrase we are looking for?
        sLine = ucase( ltrim(pDoc->GetLine(i), any chr(32,9)) )
        
        nLenMatch = len(sStartMatch)
        if nLenMatch then
            if sStartMatch = mid(sLine, 1, nLenMatch) then
            return true    ' a start block was found before a matching end was found.
            end if
        end if
        
        nLenMatch = len(sEndMatch)
        if nLenMatch then
            if sEndMatch = mid(sLine, 1, nLenMatch-1) & "|" then
            return false    ' a matching end to the block statement already exists
            end if
        end if
        
        ' Stop search if we find the start of another block statement structure and
        ' allow the insert.
        if (idBlockType <> BLOCK_STATEMENT_FUNCTION) and (idBlockType <> BLOCK_STATEMENT_SUB) then
            if left(sLine, 3)  = "if "        then return true
            if left(sLine, 4)  = "for "       then return true
            if left(sLine, 3)  = "DO "        then return true
            if left(sLine, 7)  = "SELECT "    then return true
            if left(sLine, 6)  = "WHILE "     then return true
        end if
        if left(sLine, 9) = "FUNCTION " then
            if instr(10, sLine, "=") = 0 then return true
        end if
        if left(sLine, 4)  = "SUB "       then return true
        if left(sLine, 17) = "PRIVATE FUNCTION " then return true
        if left(sLine, 16) = "PUBLIC FUNCTION "  then return true
        if left(sLine, 12) = "PRIVATE SUB "  then return true
        if left(sLine, 11) = "PUBLIC SUB "   then return true
        if left(sLine, 9)  = "PROPERTY "     then return true
        if left(sLine, 9)  = "OPERATOR "     then return true
        if left(sLine, 6)  = "UNION "        then return true
        if left(sLine, 5)  = "ENUM "         then return true
    next
        
    ' Default that we can allow the insert
    function = true
end function


' ========================================================================================
' Return the number of tabs that can fill a string of incoming size
' ========================================================================================
function NumTabsFromSpaces( byval numSpaces as long ) as long
    if gConfig.TabIndentSpaces then return numSpaces
    Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
    if pDoc = 0 then return 0
    dim as long IndentSize = SciExec(pDoc->hWndActiveScintilla, SCI_GETINDENT, 0, 0)
    if IndentSize = 0 then return 0
    return (numSpaces / IndentSize)
end function


' ========================================================================================
' Return a fill string with all spaces or TABs depending on the config setting
' ========================================================================================
function FillString( byval sSpaces as string ) as string
    ' if convert tabs to spaces then nothing needs to done because
    ' the incoming sSpaces is alreayd in the correct format.
    if gConfig.TabIndentSpaces then return sSpaces
    ' Try to convert the incoming sSpaces to TABs.
    dim as long numTabs = NumTabsFromSpaces( len(sSpaces) )
    return string(numTabs, chr(9))
end function


' ========================================================================================
' Attempt to autocomplete a statement block
' ========================================================================================
public function AttemptAutoInsert() as Long
    ' Attempts to Autocomplete a FOR/DO/SELECT, etc block. This function also deals
    ' with AutoIndentation so need to handle both separately depending on the user
    ' chosen settings.
    
    Dim as HWND hEdit
    Dim as long nLine, nCurLine, curPos, LineLen, nFoldLevel, nSpaces, IndentSize, i
    Dim as string strFill, strTemp, strCurLine, strPrevLine, strPrevLineOrig, strRightText
        

    Dim pDoc As clsDocument Ptr 

    pDoc = gTTabCtl.GetActiveDocumentPtr()
    if pDoc = 0 then exit function

    hEdit  = pDoc->hWndActiveScintilla

    ' Current and previous lines
    curPos      = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
    nLine       = pDoc->GetCurrentLineNumber
    strCurLine  = pDoc->GetLine(nLine)
    strPrevLine = pDoc->GetLine(nLine-1)
    if len(strPrevLine) = 0 then exit function

    ' Get the styling of the current line to determine if we are in a 
    ' multiline or single line comment block then abort the autoinsert.
    select case SciExec(hEdit, SCI_GETSTYLEAT, curPos, 0)
        case SCE_B_MULTILINECOMMENT
            exit function
        case SCE_B_COMMENT
            ' Allow to continue for single line comments because we want the ENTER
            ' key to position our cursor under the preceeding ' mark.
    end select
    
    
    ' Save a non-uppercased version of the strPrevLine for For/next insert
    strPrevLineOrig = strPrevLine
    
    ' Get the tab width and indent size (these are actually both the same 
    ' size as set in the pDoc->ApplyProperties code).
    'TabSize     = SciExec(hEdit, SCI_GETTABWIDTH, 0, 0) '<-- not needed
    IndentSize  = SciExec(hEdit, SCI_GETINDENT, 0, 0)
    
    ' Calculate the number of spaces to fill on the left
    for i = 1 to len(strPrevLine)
        if mid(strPrevLine, i, 1) <> " " then
            if mid(strPrevLine, i, 1) = chr(9) then
                nSpaces = nSpaces + IndentSize
            else
                Exit For
            end if
        else
            nSpaces = nSpaces + 1
        end if
    next
    strPrevLine = trim(ucase(strPrevLine), any chr(32, 9))
    
    if gConfig.AutoIndentation = 0 then
        nSpaces = 0: IndentSize = 0
    end if


    ''''''''''
    ' IF/THEN
    '   Before autoindenting an if statement make sure that this
    '   is in fact a multiline if statement.
    if (left(strPrevLine, 3) = "if " andalso right(strPrevLine, 5) = " then") then
        ' Remove the current line because we will add it again below
        ' once we have wrapped it in an end if. It will include any chunk of
        ' text that was after the then in a single line IF/THEN
        SciExec(hEdit, SCI_LINEDELETE, 0, 0)
        strFill = FillString(space(nSpaces + IndentSize)) & ltrim(strCurline, any chr(32,9)) & vbcrlf
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_IF, nLine) then
            strFill = strFill & FillString(space(nSpaces)) & "end if" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize)
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if
    
    
    ''''''''''
    ' else / elseIF
    '   Search backwards in the most recent IF/THEN line and use the 
    '   indentation of that line for the else or elseif statement.
    if gConfig.AutoIndentation then
        if (left(strPrevLine, 4) = "else") orelse (left(strPrevLine, 7) = "elseif ") then
            nCurLine = pDoc->GetCurrentLineNumber
            for i = nCurLine to 0 step -1
                strTemp = AfxStrReplace(pDoc->GetLine(i), chr(9), space(IndentSize))
                if left(ltrim(ucase(strTemp)), 3) = "if " then
                    ' Reposition the else/elseif line to line up with the if line
                    nSpaces = instr(ucase(strTemp), "if ") - 1
                    strPrevLine = FillString(space(nSpaces)) & Trim(pDoc->GetLine(nLine-1), Any Chr(32, 9))  ' need it to be original case
                    pDoc->SetLine(nLine-1, strPrevLine)
                    strFill = FillString(space(nSpaces + IndentSize)) & strCurLine  
                    pDoc->SetLine(nLine, strFill)
                    ' Add the current editing position
                    curPos = SciExec(hEdit, SCI_POSITIONFROMLINE, nLine, 0) + NumTabsFromSpaces(nSpaces + IndentSize)
                    SciExec(hEdit, SCI_SETSEL, curPos, curPos)
                    exit function
                end if
            next
        end if
    end if


    '''''''''''''
    ' SELECT CASE
    if left(strPrevLine, 12) = "SELECT CASE " then
        SciExec(hEdit, SCI_LINEDELETE, 0, 0)
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_SELECT, nLine) then
            strFill = strFill & "case " & strCurLine & vbcrlf & FillString(space(nSpaces)) & "end select" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) + 5
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    ''''''''''
    ' CASE
    '   Search backwards in the most recent CASE line and use the 
    '   indentation of that line for the CASE or CASE else statement.
    if gConfig.AutoIndentation then
        if left(strPrevLine, 5) = "CASE " then  ' this will catch CASE else as well
            strPrevLine = AfxStrReplace( pDoc->GetLine(nLine-1), chr(9), space(IndentSize))
            nSpaces = instr(ucase(strPrevLine), "CASE ") - 1
            strPrevLine = Trim(strPrevLine)  ' need it to be original case
            nCurLine = pDoc->GetCurrentLineNumber - 2  ' skip the current CASE and look for previous
            for i = nCurLine to 0 step -1
            strTemp = AfxStrReplace( pDoc->GetLine(i), chr(9), space(IndentSize) )
                dim as string strTempTrim = ltrim(ucase(strTemp))
                if left(strTempTrim, 12) = "SELECT CASE " then
                    exit for   ' no need to keep searching
                elseif left(strTempTrim, 5) = "CASE " then
                    nSpaces = instr(ucase(strTemp), "CASE ") - 1
                    exit for
                end if
            next
            ' Reposition the CASE line to line up with the previous CASE line 
            strPrevLine = FillString(space(nSpaces)) & strPrevLine
            pDoc->SetLine(nLine-1, strPrevLine)
            strFill = FillString(space(nSpaces + IndentSize)) & strCurLine  
            pDoc->SetLine(nLine, strFill)
            ' Add the current editing position
            curPos = SciExec(hEdit, SCI_POSITIONFROMLINE, nLine, 0) + NumTabsFromSpaces(nSpaces + IndentSize)
            SciExec(hEdit, SCI_SETSEL, curPos, curPos)
            exit function
        end if
    end if

            
    '''''''''''''
    ' FOR/next
    if left(strPrevLine, 4) = "for " then
        strFill = FillString(space(nSpaces + IndentSize))
        dim as CWSTR wszLoopVar
        if gConfig.FornextVariable then
            ' Determine the loop variable and append it to the "next" statement
            strPrevLine = ltrim(AfxStrShrink(strPrevLineOrig))
            wszLoopVar = " " & AfxStrParseAny(strPrevLine, 2, " =")
        end if   
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_FOR, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "next" & wszLoopVar & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if


    '''''''''''''
    ' WHILE/WEND
    if (left(strPrevLine, 6) = "WHILE ") or (strPrevLine = "WHILE") then
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_WHILE, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "wend" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize )
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if
    

    '''''''''''''
    ' DO/LOOP
    if (left(strPrevLine, 3) = "DO ") or (strPrevLine = "DO") then
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_DO, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "loop" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize)
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' FUNCTION/END FUNCTION
    if (left(strPrevLine, 9)  = "FUNCTION ") orelse _
        (left(strPrevLine, 17) = "PRIVATE FUNCTION ") orelse _
        (left(strPrevLine, 16) = "PUBLIC FUNCTION ") then
        ' Determine if this is a FUNCTION = statement rather than a true function.
        strTemp = strPrevLine
        i = instr(strTemp, "(")
        if i then strTemp = left(strTemp, i-1)
        if instr(strTemp, "=") then exit function
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_FUNCTION, nLine) then
            strFill = strFill  & vbcrlf & FillString(space(nSpaces)) & "end function" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize)
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        frmFunctions_ReparseFiles()
        exit function
    end if

    '''''''''''''
    ' SUB/END SUB
    if (left(strPrevLine, 4)  = "SUB ") orelse _
        (left(strPrevLine, 12) = "PRIVATE SUB ") orelse _
        (left(strPrevLine, 11) = "PUBLIC SUB") then
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_SUB, nLine) then
            strFill = strFill  & vbcrlf & FillString(space(nSpaces)) & "end sub" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize)
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        frmFunctions_ReparseFiles()
        exit function
    end if

    '''''''''''''
    ' PROPERTY/END PROPERTY
    if (left(strPrevLine, 9)  = "PROPERTY ") then
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_PROPERTY, nLine) then
            strFill = strFill  & vbcrlf & FillString(space(nSpaces)) & "end property" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' OPERATOR/END OPERATOR
    if (left(strPrevLine, 9)  = "OPERATOR ") then
        strFill = FillString(space(nSpaces + IndentSize))
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_OPERATOR, nLine) then
            strFill = strFill  & vbcrlf & FillString(space(nSpaces)) & "end operator" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' CONSTRUCTOR/END CONSTRUCTOR
    if (left(strPrevLine, 12)  = "CONSTRUCTOR ") then
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_CONSTRUCTOR, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end constructor" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' DESTRUCTOR/END DESTRUCTOR
    if (left(strPrevLine, 11)  = "DESTRUCTOR ") then
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_DESTRUCTOR, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end destructor" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' TYPE/END TYPE
    if (left(strPrevLine, 5)  = "TYPE ") then
        ' Determine if this is a TYPE = statement rather than a true TYPE structure
        if instr(ucase(strPrevLine), " AS ") then exit function
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_TYPE, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end type" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' WITH/END WITH
    if (left(strPrevLine, 4)  = "WITH") then
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_WITH, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end with" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' ENUM/END ENUM
    if (left(strPrevLine, 4)  = "ENUM") then
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_ENUM, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end enum" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' UNION/END UNION
    if (left(strPrevLine, 6)  = "UNION ") then
        strFill = FillString(space(nSpaces + IndentSize)) 
        if CanCompleteBlockStatement(pDoc, BLOCK_STATEMENT_UNION, nLine) then
            strFill = strFill & vbcrlf & FillString(space(nSpaces)) & "end union" & vbcrlf
        end if          
        SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))
        curPos = curPos + NumTabsFromSpaces(nSpaces + IndentSize) 
        SciExec(hEdit, SCI_SETSEL, curPos, curPos)
        exit function
    end if

    '''''''''''''
    ' Add the same spaces on the left that the line above
    strFill = FillString(space(nSpaces))
    SciExec(hEdit, SCI_ADDTEXT, len(strFill), strptr(strFill))

    function = 0
end function

