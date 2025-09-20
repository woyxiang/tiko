' ########################################################################################
' Microsoft Windows
' File: CW_GDI_DrawShadedRectangle_01.bas
' Contents: Drawing a Shaded Rectangle
' To draw a shaded rectangle, define a TRIVERTEX array with two elements and a single
' GRADIENT_RECT structure. The following code example shows how to draw a shaded rectangle
' using the GradientFill function with the GRADIENT_FILL_RECT mode defined.
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GDI - Drawing a shaded rectangle", @WndProc)
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   pWindow.WindowStyle = WS_POPUPWINDOW OR WS_CAPTION   ' // make the window not resizable
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 300)
   ' // Centers the window
   pWindow.Center

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a shaded rectangle using the GradientFill function.
' ========================================================================================
SUB DrawShadedRectangle (BYVAL hdc AS HDC, BYVAL rxRatio AS SINGLE, BYVAL ryRatio AS SINGLE)

   ' // Create an array of TRIVERTEX structures that describe
   ' // positional and color values for each vertex. For a rectangle,
   ' // only two vertices need to be defined: upper-left and lower-right.

   DIM vertex(1) AS TRIVERTEX

   vertex(0).x     = 50 * rxRatio
   vertex(0).y     = 100 * ryRatio
   vertex(0).Red   = &h0000
   vertex(0).Green = &h8000
   vertex(0).Blue  = &h8000
   vertex(0).Alpha = &h0000

   vertex(1).x     = 450 * rxRatio
   vertex(1).y     = 200 * ryRatio
   vertex(1).Red   = &h0000
   vertex(1).Green = &hD000
   vertex(1).Blue  = &hD000
   vertex(1).Alpha = &h0000

   ' // Create a GRADIENT_TRIANGLE structure that
   ' // references the TRIVERTEX vertices.

   DIM gRect AS GRADIENT_RECT

   gRect.UpperLeft  = 0
   gRect.LowerRight = 1

   ' // Draw a shaded rectangle.
   GradientFill hdc, @vertex(0), 2, @gRect, 1, GRADIENT_FILL_RECT_H

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC rxRatio AS SINGLE     ' // Horizontal scaling ratio
   STATIC ryRatio AS SINGLE     ' // Vertical scaling ratio

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns �1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         ' // Get a pointer to the CWindow class from the CREATESTRUCT structure
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' // Get the scaling ratios
         rxRatio = pWindow->rxRatio
         ryRatio = pWindow->ryRatio
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

      CASE WM_PAINT
         DIM ps AS PAINTSTRUCT
         DIM hdc AS HDC = BeginPaint(hwnd, @ps)
         DrawShadedRectangle hdc, rxRatio, ryRatio
         EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
