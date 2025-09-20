' ########################################################################################
' Microsoft Windows
' File: CW_ListView_01.bas
' Contents: ListView control
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

#define IDC_LISTVIEW 1001

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Create the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - ListView control", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
'   pWindow.ClassStyle = CS_DBLCLKS   ' // Change the window style to avoid flicker
   pWindow.SetClientSize(770, 365)
   ' // Center the window
   pWindow.Center

   ' // Adds a listview
   DIM hListView AS HWND = pWindow.AddControl("ListView", hWin, IDC_LISTVIEW)
   pWindow.SetWindowPos hListView, NULL, 0, 0, 770, 365, SWP_NOZORDER
   ' // Anchor the ListView
   pWindow.AnchorControl(IDC_LISTVIEW, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Add some extended styles
   DIM dwExStyle AS DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hListView)
   dwExStyle = dwExStyle OR LVS_EX_FULLROWSELECT OR LVS_EX_GRIDLINES
   ListView_SetExtendedListViewStyle(hListView, dwExStyle)

   ' // Set the text color of the ListView
   ListView_SetTextColor(hListView, RGB_BLUE)

   ' // Add the header's column names
   DIM dwsText AS DWSTRING
   FOR i AS LONG = 0 TO 9
      dwsText = "Column " & WSTR(i)
      ListView_AddColumn(hListView, i, dwsText, pWindow.ScaleX(110))
   NEXT

   ' // Populate the ListView with some data
   FOR i AS LONG = 0 to 29
      dwsText = "Column 1 Row " + WSTR(i)
      ListView_AddItem(hListView, i, 0, dwsText)
      FOR x AS LONG = 0 TO 9
         dwsText = "Column " & WSTR(x) & " Row " + WSTR(i)
         ListView_SetItemText(hListView, i, x, dwsText)
      NEXT
   NEXT

   ' // Select the fist item (ListView items are zero based)
   ListView_SelectItem(hListView, 0)
   ' // Set the focus in the ListView
   SetFocus hListView

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns �1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
