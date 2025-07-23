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
    
#pragma once 

#define IDC_FRMLISTVIEW_HEADER               1000
#define IDC_FRMLISTVIEW_LISTBOX              1001
#define IDC_FRMLISTVIEW_VSCROLL              1002

type LISTVIEW_COLUMN_TYPE
    wszText as CWSTR      ' row 0 is the header data
    nWidth as long
end type

type LISTVIEW_DATA_TYPE
    ColData(any) as LISTVIEW_COLUMN_TYPE
    ForeColor as COLORREF
    ForeColorHot as COLORREF
    BackColor as COLORREF
    BackColorHot as COLORREF
end type


declare function frmListView_PositionWindows( byval hLV as HWND ) as LRESULT
declare function frmListView_SelectItem( byval hLV as HWND, byval nLineNum as long ) as long
declare function frmListView_GetCount( byval hLV as HWND ) as long
declare function frmListView_AddString( byval hLV as HWND, byval wszText as CWSTR ) as long
declare function frmListView_DeleteAllItems( byval hLV as HWND ) as long
declare function frmListView_GetItemText( byval hLV as HWND, byval nLineNum as long, byval nColNum as long ) as CWSTR
declare function frmListView_GetSelection( byval hLV as HWND ) as long
declare function frmListView_AddColumn( _
    byval hLV as HWND, byval nColNum as long, byval wszText as CWSTR, byval nWidth as long ) as long
declare function frmListView_SetForeColors( _
    byval hLV as HWND, byval ForeColor as COLORREF, byval ForeColorHot as COLORREF ) as long
declare function frmListView_SetBackColors( _
    byval hLV as HWND, byval BackColor as COLORREF, byval BackColorHot as COLORREF ) as long
declare function frmListView_Show( byval hwndParent as hwnd ) as HWND

