# CMaskedEdit Class
The `CMaskedEdit` class supports a masked edit control, which validates user input against a mask and displays the validated results according to a template.

**Include file**: AfxNova/CMaskedEdit.inc
  
### Constructors

| Name | Description |  
| ---------- | ----------------- |  
|[Constructors](#constructors1)| Creates an instance of the class and the control. |  

### Functions

| Name | Description |  
| ---------- | ----------------- |  
|[AfxCMaskedEditPtr](#afxcmaskededitptr)| Returns a pointer to the **CMaskedEdit** class. |  

### Public Methods  
  
| Name | Description |  
| ---------- | ----------------- |  
|[Create](#create)|Creates an instance of the control|
|[EnableMask](#enablemask)|Initializes the masked edit control.|
|[GetText](#gettext)|Retrieves validated text from the masked edit control.|
|[GetTextLength](#gettextlength)|Retrieves the length of the text.|
|[hWindow](#hWindow)|Gets the control window handle. |
|[SetPos](#setpos)|Sets the position of the cursor.|
|[SetValidChars](#setvalidchars)|Specifies a string of valid characters that the user can enter.|
|[SetText](#settext)|Displays a prompt in the masked edit control.|

## Constructors

```
CONSTRUCTOR CMaskedEdit
CONSTRUCTOR CMaskedEdit (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the parent `CWindow`class. |
| *cID* | The control identifier, an integer value used to notify its parent about events. The application determines the control identifier; it must be unique for all controls with the same parent window. |
| *X* | The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* |  The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the control. |
| *nHeight* | The height of the control. |
| *dwStyle* | The window styles of the control being created.<br>Default styles: WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL. |
| *dwExStyle* | The extended window styles of the control being created.<br> Default extended style: WS_EX_CLIENTEDGE |

### Remarks  
 Perform the following steps to use the `CMaskedEdit` control in your application:  
  
 1. Embed a `CMaskedEdit` object into your window class.  
  
 2. Call the [EnableMask](#enablemask) method to specify the mask.  
  
 3. Call the [SetValidChars](#setvalidchars) method to specify the list of valid characters.  
  
 4. Call the [SetText](#settext) method to specify the default text for the masked edit control.  
  
 5. Call the [GetText](#gettext) method to retrieve the validated text.  
  
 
### Example  
The following example demonstrates how to set up a mask (for example a phone number) by using the **EnableMask** method to create the mask for the masked edit control, and **SetText**method to display a prompt in the masked edit control.

```
DIM pMaskedEdit AS CMaskedEdit = CMaskedEdit(@pWindow, IDC_MASKED, 10, 30, 280, 23)
pMaskedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMaskedEdit.SetText("(123) 123-1212"), TRUE
```

### Full example
```
' ########################################################################################
' Microsoft Windows
' Contents: CWindow - Masked Edit Control
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CMaskedEdit.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, WCOMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_MASKED = 1001

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "Masked edit control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(300, 150)
   ' // Centers the window
   pWindow.Center

   ' // Add a masked edit control
   DIM pMaskedEdit AS CMaskedEdit = CMaskedEdit(@pWindow, IDC_MASKED, 50, 30, 200, 23)
   SetFocus pMaskedEdit.hWindow
   pMaskedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
   pMaskedEdit.SetText("(123) 123-1212"), TRUE

   ' // Adds a button
   pWindow.AddControl("Button", , IDCANCEL, "&Close", 190, 100, 75, 23)

   ' // Anchor controls
   pWindow.AnchorControl(IDC_MASKED, AFX_ANCHOR_WIDTH)
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

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_MASKED
               IF CBCTLMSG(wPAram, lParam) = EN_CHANGE THEN
                  DIM hEdit AS HWND = GetDlgItem(hwnd, LOWORD(wParam))
                  DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(hEdit)
                  IF pMaskedEdit THEN
                     DIM dwsText AS DWSTRING
                     ' // Get the text including the mask
                     dwsText = pMaskedEdit->GetText(FALSE)
                     ' // Get the text excluding the mask
                     dwsText = pMaskedEdit->GetText(TRUE)
                  END IF
               END IF
         END SELECT


    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
---

## Create

```
FUNCTION CMaskedEdit.Create (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1) AS HWND
```

### Parameters  

Same parameters that the `Constructor`.

```
DIM pMakedEdit AS CMaskedEdit
pMaskEdit.Create(CMaskedEdit(@pWindow, IDC_MASKED, 10, 30, 280, 23)
pMakedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMakedEdit.SetText("(123) 123-1212"), TRUE
```

## EnableMask  

Initializes the masked edit control.  
  
```  
SUB EnableMask (BYVAL lpszMask AS WSTRING PTR, BYVAL lpszInputTemplate AS WSTRING PTR, _
   BYREF chMaskInputTemplate AS DWSTRING = "_", BYVAL lpszValid AS WSTRING PTR = NULL)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *lpszMask* | A mask string that specifies the type of character that can appear at each position in the user input. The length of the *lpszInputTemplate* and *lpszMask* parameter strings must be the same. See the Remarks section for more detail about mask characters. |
| *lpszInputTemplate* | A mask template string that specifies the literal characters that can appear at each position in the user input. Use the underscore character ("\_") as a character placeholder. The length of the *lpszInputTemplate* and *lpszMask* parameter strings must be the same.  |
| *chMaskInputTemplate* | A default character that the framework substitutes for each invalid character in the user input. The default value of this parameter is underscore ("\_"). |
| *lpszValid* | A string that contains a set of valid characters. NULL indicates that all characters are valid. The default value of this parameter is NULL.  |
  
### Remarks  
 
Use this method to create the mask for the masked edit control.
  
The following table list the default mask characters:  
  
| Mask Character | Definition |  
|----------------|------------|  
| D | Digit.|  
| d | Digit or space.|  
| + | Plus ('+'), minus ('-'), or space.|  
| C | Alphabetic character.|  
| c | Alphabetic character or space.|  
| A | Alphanumeric character.|  
| a | Alphanumeric character or space.|  
| * | A printable character.|  

---

## GetText  

Retrieves validated text from the masked edit control.  
  
```  
FUNCTION GetText () AS DWSTRING
FUNCTION GetText (BYVAL bGetMaskedCharsOnly AS BOOLEAN) AS DWSTRING
```  
| Parameter  | Description |
| ---------- | ----------- |
| *bGetMaskedCharsOnly* | IF TRUE, returns the text with the input mask removed; if FALSE, returns the text with the mask, as displayed in the control. |

### Return Value  

The text from the masked edit control.

#### Example

```
pMskEd.EnableMask("       cc       ddddd-dddd", "State: __, Zip: _____-____", "_")
pMskEd.SetText "State: NY, Zip: 12345-6789", TRUE
print pMskEd.GetText(FALSE)   ' // Returns "State: NY, Zip: 12345-6789"
print pMskEd.GetText(TRUE)   ' // Returns NY123456789
```

## GetTextLength

Retrieves the length of the text.

```  
FUNCTION GetTextLength (BYVAL bGetMaskedCharsOnly AS BOOLEAN = FALSE) AS LONG
```  
| Parameter  | Description |
| ---------- | ----------- |
| *bGetMaskedCharsOnly* | IF TRUE, returns the length of text with the input mask removed; if FALSE, returns the length text with the mask, as displayed in the control. |

### Return Value  

The length text from the masked edit control.

#### Example

```
pMskEd.EnableMask("       cc       ddddd-dddd", "State: __, Zip: _____-____", "_")
pMskEd.SetText "State: NY, Zip: 12345-6789", TRUE
print pMskEd.GetTextLength   ' // Returns 14
print pMskEd.GetTextLength(TRUE)   ' // Returns 10
```
---

## hWindow

Gets the control window handle.

```
PROPERTY hWindow () AS HWND
```

#### Return value

The window handle.

---

## SetPos

Sets the position of the cursor.

```
SUB SetPos (BYVAL nPos AS LONG = -1)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | 0-based index. |

---

## SetText  

Displays a prompt in the masked edit control.  
  
```  
FUNCTION SetText (BYREF dwsText AS DWSTRING) AS BOOLEAN
FUNCTION SetText (BYREF dwsText AS DWSTRING, BYVAL bSetMaskedCharsOnly AS BOOLEAN) AS BOOLEAN
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *dwsText* | Points to a string that will be used as a prompt. |
| *bSetMaskedCharsOnly* | Set the masked characters only, without the input mask. |

#### Example

```
pMskEd.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMskEd.SetText("1231231212")
```
```
pMskEd.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
pMskEd.SetText("(123) 123-1212"), TRUE
```
---

## SetValidChars  

Specifies a string of valid characters that the user can enter.  
  
```  
SUB SetValidChars (BYVAL lpszValid AS WSTRING PTR)
```  
  
| Parameter  | Description |
| ---------- | ----------- |
| *lpszValid* | A string that contains the set of valid input characters. NULL means that all characters are valid. The default value of this parameter is NULL. |
  
### Remarks  
Use this method to define a list of valid characters. If an input character is not in this list, masked edit control will not accept it. 
  
The following code example accepts only hexadecimal numbers.  
 
```
pMskEd.EnableMask("  AAAA"), _   ' // Mask string
("0x____"), _   ' // Template string
("_")   ' // The default character that replaces the backspace character
pMskEd.SetValidChars("1234567890ABCDEFabcdef")   ' // Valid string characters
pMskEd.SetText("0x01AF"), TRUE
```
---
## AfxCMaskedEditPtr

Returns a pointer to the **CMaskedEdit** class given the handle of the masked edit control or the handle of its parent window and its control identifier.

```
FUNCTION AfxCMaskedEditPtr (BYVAL hMaskedEdit AS HWND) AS CMaskedEdit PTR
FUNCTION AfxCMaskedEditPtr (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CMaskedEdit PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMaskedEdit* | Handle of the masked edit control. |
| *hParent* | Handle of the parent window. |
| *cID* | Identifier of the control. |

#### Return value

A pointer to the **CMaskedEdit** class or NULL.

#### Usage examples

```
DIM hMasKedEdit AS HWND = GetDlgItem(hParent, cID)
DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(hMasKedEdit))
```
```
DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(hParent, cID)
```
---
