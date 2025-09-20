' ########################################################################################
' Microsoft Windows
' File: CW_Header_01.bas
' Contents: CWindow - Header
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_HEADER = 1001

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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Header control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Add the header control.
   DIM hHeader AS HWND = pWindow.AddControl("Header", hWin, IDC_HEADER)
   ' // Insert items
   DIM thdi AS HDITEMW
   DIM wszItem AS WSTRING * 260
   wszItem   = "Item 1"
   thdi.Mask = HDI_WIDTH OR HDI_FORMAT OR HDI_TEXT
   thdi.fmt  = HDF_LEFT OR HDF_STRING
   thdi.cxy  = pWindow.ScaleX(80)
   thdi.pszText  = @wszItem
   Header_InsertItem(hHeader, 1, @thdi)
   wszItem   = "Item 2"
   Header_InsertItem(hHeader, 2, @thdi)
   wszItem   = "Item 3"
   Header_InsertItem(hHeader, 3, @thdi)
   wszItem   = "Item 4"
   Header_InsertItem(hHeader, 4, @thdi)
   wszItem   = "Item 5"
   Header_InsertItem(hHeader, 5, @thdi)

   ' // Adds a button
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 270, 155, 75, 30)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

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

      CASE WM_NOTIFY
         ' // Retrieve the header item clicked
         DIM hd AS NMHEADERW
         CBNMTYPESET(hd, wParam, lParam)
         IF hd.hdr.idFrom = IDC_HEADER THEN
            SELECT CASE hd.hdr.code
               CASE HDN_ITEMCLICKW
                  MessageBox hwnd, "You have clicked item " & WSTR(hd.iItem + 1), "Message", MB_OK
            END SELECT
         END IF

      CASE WM_SIZE
         ' // Resize the header control
         DIM hHeader AS HWND = GetDLgItem(hwnd, IDC_HEADER)
         DIM thdl AS HDLAYOUT, twp AS WINDOWPOS, trc AS RECT
         GetClientRect hwnd, @trc
         thdl.prc = @trc
         thdl.pwpos = @twp
         Header_Layout(hHeader, @thdl)
         SetWindowPos hHeader, NULL, twp.x, twp.y, twp.cx, twp.cy, SWP_NOZORDER OR SWP_NOACTIVATE
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
