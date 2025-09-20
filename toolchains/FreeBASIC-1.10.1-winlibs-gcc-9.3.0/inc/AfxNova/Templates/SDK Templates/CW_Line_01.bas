' ########################################################################################
' Microsoft Windows
' File: CW_Line_01.bas
' Contents: CWindow - Line control
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/RGB_Colors.bi"
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_LINE1 = 1001
CONST IDC_LINE2 = 1002
CONST IDC_LINE3 = 1003

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Line control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Set the main window background color
   pWindow.SetBackColor(RGB_OLDLACE)

   ' // Adds the controls
   pWindow.AddControl("Line", hWin, IDC_LINE1, "", 20, 20, 360, 1, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY)
   pWindow.AddControl("Line", hWin, IDC_LINE2, "", 20, 40, 360, 15, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY)
   pWindow.AddControl("Line", hWin, IDC_LINE3, "", 20, 80, 360, 30, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY)

   ' // Anchor the controls
   pWindow.AnchorControl(IDC_LINE1, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(IDC_LINE2, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(IDC_LINE3, AFX_ANCHOR_WIDTH)

   ' // Adds a button
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 300, 160, 75, 30)
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
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
            CASE IDC_LINE1
               IF CBCTLMSG(wParam, lParam) = STN_CLICKED THEN
                  MessageBoxW hwnd, "Line 1 clicked", "Line example", MB_OK
               END IF
            CASE IDC_LINE2
               IF CBCTLMSG(wParam, lParam) = STN_CLICKED THEN
                  MessageBoxW hwnd, "Line 2 clicked", "Line example", MB_OK
               END IF
            CASE IDC_LINE3
               IF CBCTLMSG(wParam, lParam) = STN_CLICKED THEN
                  MessageBoxW hwnd, "Line 3 clicked", "Line example", MB_OK
               END IF
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
