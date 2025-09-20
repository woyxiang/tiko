# CDispInvoke Class

`CDispInvoke` allows calling COM methods and properties with Free Basic.

**Include file**: AfxNova/CDispInvoke.inc

# Constructors

| Name       | Description |
| ---------- | ----------- |
| [Constructor(PROGID)](#constructor1) | Creates a single uninitialized object of the class associated with a specified ProgID or CLSID. |
| [Constructor(CLSID)](#constructor2) | Creates a single uninitialized object of the class associated with a specified CLSID. |
| [Constructor(IDispatch)](#constructor3) | Creates a single uninitialized object of the class associated with a pointer to a Dispatch interface. |
| [Constructor(VARIANT)](#constructor4) | Creates a single uninitialized object of the class associated with a variant of the type VT_DISPATCH. |
| [Constructor(LibName)](#constructor5) | Loads the specified library from file and creates an instance of an object. |

---

# Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator \*](#operator1) | Returns the underlying IDispatch pointer. |
| [Operator LET](#operator2) | Assigns an IDispatch pointer and increases the reference count. |
| [vptr](#vptr) | Clears the contents of the class and returns the address of the underlying IDispatch pointer. |

---

# Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Attach](#attach) | Attaches a Dispatch pointer. |
| [Detach](#detach) | Detaches a Dispatch pointer. |
| [DispInvoke](#dispInvoke) | Calls a method or a get property. |
| [DispObj](#dispObj) | Returns a counted reference of the underlying dispatch pointer. |
| [DispPtr](#dispPtr) | Returns a pointer to the dispatch interface. |
| [Get](#get) | Calls the specified property of an interface and gets the value returned. |
| [GetArgErr](#getargerr) | Returns the index within rgvarg of the first argument that has an error. |
| [GetVarResult](#getvarresult) | Returns the last result code returned by a call to the Invoke method. |
| [GetLcid](#getlcid) | Retrieves the locale identifier used by the class. |
| [Invoke](#invoke) | Calls a method or a get property. |
| [SetLcid](#setlcid) | Sets the locale identifier used by the class. |
| [Put](#put) | Calls the specified property of an interface and sets the passed value. |
| [PutRef](#putref) | Assigns a value to an interface member property that contains a reference to an object. |
| [Set](#set) | Calls the specified property of an interface and sets the passed value. |
| [SetRef](#setref) | Assigns a value to an interface member property that contains a reference to an object. |

---

## Error and result codes

| Name       | Description |
| ---------- | ----------- |
| [GetErrorInfo](#geterrorinfo) | Returns a description of the specified error code. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |

---
## <a name="Constructor1"></a>Constructor(ProgID)

Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.

```
CONSTRUCTOR CDispInvoke (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszProgID* | The ProgID or the CLSID of the object to create.<br>A ProgID such as "MSCAL.Calendar.7"<br>A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}" |
| *wszLicKey* | Optional. The license key as a unicode string. |

#### Usage example

```
DIM pDispInvoke AS CDispInvoke = "Scripting.Dictionary"
' -or-
pDispInvoke = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
```

#### Example

The following example demonstrates how to use it to create an instance of the VBScript RegExp object and execute a search.

```
#include once "AfxNova/CDispInvoke.inc"
USING AfxNova

' // Create an instance of the RegExp object
DIM pDisp AS CDispInvoke = "VBScript.RegExp"
' // To check for success, see if the value returned by the DispPtr method is not null
IF pDisp.DispPtr = NULL THEN END

' // Set some properties
' // Use VARIANT_TRUE or CTRUE, not TRUE, because Free Basic TRUE is a BOOLEAN data type, not a LONG
pDisp.Put("Pattern") = ".is"
pDisp.Put("IgnoreCase") = VARIANT_TRUE
pDisp.Put("Global") = VARIANT_TRUE

' // Execute a search
DIM pMatches AS CDispInvoke = pDisp.Invoke("Execute", "IS1 is2 IS3 is4")
' // Parse the collection of matches
IF pMatches.DispPtr THEN
   ' // Get the number of matches
   DIM nCount AS LONG = VAL(pMatches.Get("Count"))
   ' // This is equivalent to:
   ' DIM dvRes AS DVARIANT = pMatches.Get("Count")
   ' DIM nCount AS LONG = (VAL(dvRes)
   FOR i AS LONG = 0 TO nCount -1
      ' // Get a pointer to the Match object
      ' // When using COM Automation, it's not always necessary to make sure that the
      ' // passed variant with a numeric value is of the exact type, since the standard
      ' // implementation of DispInvoke tries to coerce parameters. However, it is always
      ' // safer to use a syntax like DVARIANT(i, "LONG")) than DVARIANT(i)
'      DIM pMatch AS CDIspInvoke = pMatches.Get("Item", DVARIANT(i, "LONG"))
      DIM pMatch AS CDIspInvoke = pMatches.Get("Item", i)
      IF pMatch.DispPtr THEN
         ' // Get the value of the match and convert it to a string
         print pMatch.Get("Value")
      END IF
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```
---

# <a name="Constructor2"></a>Constructor(CLSID)

Creates a single uninitialized object of the class associated with a specified CLSID.

```
CONSTRUCTOR CDispInvoke (BYREF classID AS CONST CLSID)
CONSTRUCTOR CDispInvoke (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *wszClsID* | A CLSID in string format. |
| *wszIID* | The identifier of the interface to be used to communicate with the object. |

#### Usage examples

```
DIM pDispInvoke AS CDispInvoke = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
```

```
DIM pDispInvoke AS CDispInvoke = (CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'    CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
```

## <a name="Constructor3"></a>Constructor(IDispatch)

Creates a single uninitialized object of the class associated with a pointer to a Dispatch interface.

```
CONSTRUCTOR CDispInvoke (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | Pointer to a Dispatch interface. |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling AddRef. This is the usual case when we assign directly an already AddRefed pointer returned by a COM method. If it is true, then **AddRef is called**. This is needed when we pass a raw pointer. |

#### Example

The following example combines CDispInvoke and CWmiDisp to set the specified printer as the default one.

```
#include once "AfxNova/CDispInvoke.inc"
#include once "AfxNova/CWmiDisp.inc"
USING AfxNova

' // Connect with WMI in the local computer and get the properties of the specified printer
DIM pDisp AS CDispInvoke = CWmiServices( _
   $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:" & _
   "Win32_Printer.DeviceID='OKI B410'").ServicesObj

' // Set the printer as the default printer
pDisp.Invoke("SetDefaultPrinter")

PRINT
PRINT "Press any key..."
SLEEP
```
---

## <a name="Constructor4"></a>Constructor(VARIANT)

Creates a single uninitialized object of the class associated with a variant of the type VT_DISPATCH.

```
CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT PTR, BYVAL fAddRef AS BOOLEAN = TRUE)
CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT, BYVAL fAddRef AS BOOLEAN = TRUE)
CONSTRUCTOR CDispInvoke (BYREF cvDisp AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *vDisp* | A VT_DISPATCH variant. |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling AddRef. This is the usual case when we assign directly an already AddRefed pointer returned by a COM method. If it is true, then **AddRef** is called. This is needed when we pass a raw pointer. |
| *cvDisp* | A VT_DISPATCH DVARIANT. |

---

## <a name="Constructor5"></a>Constructor(LibName)

Loads the specified library from file and creates an instance of an object.

```
CONSTRUCTOR CDispInvoke (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszLibName* | Full path where the library is located. |
| *rclsid* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *riid* | The identifier of the interface to be used to communicate with the object. |
| *wszLicKey* | Optional. The license key. |

#### Remarks

Not every component is a suitable candidate for use under this overloaded method.
* Only in-process servers (DLLs) are supported.
* Components that are system components or part of the operating system, such as XML, Data Access, Internet Explorer, or DirectX, aren't supported
* Components that are part of an application, such Microsoft Office, aren't supported.
* Components intended for use as an add-in or a snap-in, such as an Office add-in or a control in a Web browser, aren't supported.
* Components that manage a shared physical or virtual system resource aren't supported.
* Visual ActiveX controls aren't supported because they need to be initialized and activated by the OLE container.

#### Usage example

```
DIM pDispInvoke AS CDispInvoke = (CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'    CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
```
---

## <a name="Operator1"></a>Operator *

Returns the underlying IDispatch pointer.

```
OPERATOR * (BYREF cdi AS CDispInvoke) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cdi* | An instance of the CDispInvoke class. |

---

## <a name="Operator2"></a>Operator LET (=)

Assigns an IDispatch pointer and increases the reference count.

```
OPERATOR Let (BYVAL pdisp AS IDispatch PTR)
OPERATOR Let (BYVAL vDisp AS VARIANT PTR)
OPERATOR Let (BYVAL vDisp AS VARIANT)
OPERATOR Let (BYREF cvDisp AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | An IDispatch pointer. |
| *vDisp* | A VT_DISPATCH variant. |
| *cvDisp* | A VT_DISPATCH DVARIANT. |

---

## vptr

Clears the contents of the class and returns the address of the underlying IDispatch pointer.

```
FUNCTION vptr () AS IDispatch PTR PTR
```

#### Remarks

To pass the class to an OUT BYVAL IDispatch PTR parameter.

If we pass a **CDispInvoke** variable to a function with an OUT IDispatch parameter without first clearing the contents of the class, we may have a memory leak.

---

## Attach

Attaches a Dispatch pointer to the class.

```
FUNCTION Attach (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | Pointer to a Dispatch interface, |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling **AddRef**.<br>If it is true, then **AddRef** is called. This is needed when we pass a raw pointer. |

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| E_INVALIDARG | The passed pointer is null. |

---

## Detach

Detaches the Dispatch pointer from the class.

```
FUNCTION Detach () AS IDispatch PTR
```

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |

#### Remarks

Extracts and returns the encapsulated interface pointer, and then clears the encapsulated pointer storage to NULL. This removes the interface pointer from encapsulation. It is up  to the caller to call **Release** on the returned interface pointer.

---

## DispInvoke

Calls a method or a get property.

```
FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL dispIdMember AS DISPID, BYVAL prgArgs AS VARIANT PTR = NULL, _
   BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL pwszName AS WSTRING PTR, BYVAL prgArgs AS VARIANT PTR = NULL, _
   BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wFlags* | Flags describing the context of the Invoke call. |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the method or property to call. |
| *prgArgs* | Array of variant parameters in reversed order. |
| *cArgs* | The number of arguments in the *prgArgs* array. |
| *lcid* | The locale context in which to interpret arguments. The lcid is used by the **GetIDsOfNames** function, and is also passed to Invoke to allow the object to interpret its arguments specific to a locale. Applications that do not support multiple national languages can ignore this parameter. |

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |

#### Remarks

Extracts and returns the encapsulated interface pointer, and then clears the encapsulated pointer storage to NULL. This removes the interface pointer from encapsulation. It is up  to the caller to call **Release** on the returned interface pointer.

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| DISP_E_BADPARAMCOUNT | The number of elements provided to **DISPPARAMS** is different from the number of arguments accepted by the method or property. |
| DISP_E_BADVARTYPE | One of the arguments in **DISPPARAMS** is not a valid variant type. |
| DISP_E_EXCEPTION | The application needs to raise an exception. |
| DISP_E_MEMBERNOTFOUND | The requested member does not exist. |
| DISP_E_NONAMEDARGS | This implementation of IDispatch does not support named arguments. |
| DISP_E_OVERFLOW | One of the arguments in **DISPPARAMS** could not be coerced to the specified type. |
| DISP_E_PARAMNOTFOUND | One of the parameter IDs does not correspond to a parameter on the method. In this case, puArgErr is set to the first argument that contains the error. |
| DISP_E_TYPEMISMATCH | One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within rgvarg is returned in *puArgErr*. |
| DISP_E_UNKNOWNINTERFACE | The interface identifier passed in riid is not IID_NULL. |
| DISP_E_UNKNOWNLCID | The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned. |
| DISP_E_PARAMNOTOPTIONAL | A required parameter was omitted. |

#### Remarks

This method is called internally by the **Get**, **Put**, **PutRef** and **Invoke** methods of the **CDispInvoke class**. You don't need to call it directly.

---

## DispObj

Returns a counted reference of the underlying dispatch pointer. You must release it, e.g. calling call **IUnknown_Release** or the function **AfxSafeRelease** when no longer need it.

```
FUNCTION DispObj () AS IDispatch PTR
```
---

## DispPtr

Returns a pointer to the dispatch interface. Don't call **IUnknown_Release** on it.

```
FUNCTION DispPtr () AS IDispatch PTR
```
---

## Get

Calls the specified property of an interface and gets the value returned. The overloaded methods allow to use the **Get** method passing one argument, two arguments or none.

```
FUNCTION Get (BYVAL dispID AS DISPID) AS DVARIANT
FUNCTION Get (BYVAL dispID AS DISPID, BYREF dvArg AS DVARIANT) AS DVARIANT
FUNCTION Get (BYVAL dispID AS DISPID, BYREF dvArg1 AS DVARIANT, BYREF dvArg2 AS DVARIANT) AS DVARIANT
FUNCTION Get (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF dvArg AS DVARIANT) AS DVARIANT
FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF dvArg1 AS DVARIANT, BYREF dvArg2 AS DVARIANT) AS DVARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *dvArg1* | DVARIANT. First argument. |
| *dvArg2* | DVARIANT. Second argument. |

#### Return value

The property value.

#### Example

```
#include "windows.bi"
#include "AfxNova/CWmiDisp.inc"
USING AfxNova

' // Connect to WMI using a moniker
' // Note: $ is used to avoid the pedantic warning of the compiler about escape characters
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
IF pServices.ServicesPtr = NULL THEN END
' // Execute a query
DIM hr AS HRESULT = pServices.ExecQuery("SELECT Caption, SerialNumber FROM Win32_BIOS")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END
' // Get the number of objects retrieved
DIM nCount AS LONG = pServices.ObjectsCount
print "Count: ", nCount
IF nCount = 0 THEN PRINT "No objects found" : SLEEP : END
' // Enumerate the objects using the standard IEnumVARIANT enumerator (NextObject method)
' // and retrieve the properties using the CDispInvoke class.
DIM pDispServ AS CDispInvoke = pServices.NextObject
IF pDispServ.DispPtr THEN
   PRINT "Caption: "; pDispServ.Get("Caption")
   PRINT "Serial number: "; pDispServ.Get("SerialNumber")
END IF
PRINT
PRINT "Press any key..."
SLEEP
```
---

## GetArgErr

Returns the index within rgvarg of the first argument that has an error. Arguments are stored in pDispParams->rgvarg in reverse order, so the first argument is the one with the highest index in the array. This parameter is returned only when the resulting return value is DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND.

```
FUNCTION GetArgErr () AS UINT
```
---

## GetLastResult

Returns the result code returned by the last executed method..

```
FUNCTION GetLastResult () AS HRESULT
```
---

## <a name="setresult"></a>SetResult

Sets the result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The error code returned by the methods. |

---

### <a name="geterrorinfo"></a>GetErrorInfo

Returns a description of the last error code.
```
PRIVATE GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```
---

## GetVarResult

The result of a call to the Invoke method. Not usually needed because Invoke already returns it as the result of the function.

```
FUNCTION GetVarResult () AS DVARIANT
```
---

## GetLcid

Returns the locale identifier used by the class.

```
FUNCTION GetLcid () AS LCID
```
---

## Invoke

Calls a method or a get property.

```
FUNCTION Invoke (BYVAL dispID AS DISPID) AS DVARIANT
FUNCTION Invoke (BYVAL dispID AS DISPID, dvArg1...dvArg16) AS DVARIANT
FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, dvArg1...dvArg16) AS DVARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the method or property to call. |
| *dvArg1...dvArg16* | DVARIANT. Parameters that will be passed to **IDIspatch.Invoke** as an array of variants in reversed order. |

#### Return value

Returns a variant with the result of the call when the **Invoke** method is used instead of **Get** to retrieve the value of a property. If it is not a get property, it returns a null variant.

#### Remarks

For optional parameters, we must use a VT_ERROR VARIANT with a value of DISP_E_PARAMNOTFOUND. You can use the function **AfxDVarOptPrm** or the macro **DVAR_OPTPRM** for this purpose.

To check for success or failure, call the **GetLastResult** method. It will return S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| DISP_E_BADPARAMCOUNT | The number of elements provided to **DISPPARAMS** is different from the number of arguments accepted by the method or property. |
| DISP_E_BADVARTYPE | One of the arguments in **DISPPARAMS** is not a valid variant type. |
| DISP_E_EXCEPTION | The application needs to raise an exception. |
| DISP_E_MEMBERNOTFOUND | The requested member does not exist. |
| DISP_E_NONAMEDARGS | This implementation of IDispatch does not support named arguments. |
| DISP_E_OVERFLOW | One of the arguments in **DISPPARAMS*** could not be coerced to the specified type. |
| DISP_E_PARAMNOTFOUND | One of the parameter IDs does not correspond to a parameter on the method. In this case, *puArgErr* is set to the first argument that contains the error. |
| DISP_E_TYPEMISMATCH | One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within *rgvarg* is returned in *puArgErr*. |
| DISP_E_UNKNOWNINTERFACE | The interface identifier passed in *riid* is not IID_NULL. |
| DISP_E_UNKNOWNLCID | The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned. |
| DISP_E_PARAMNOTOPTIONAL | A required parameter was omitted. |

#### Usage example

```
#include once "AfxNova/CDIspInvoke.inc"
USING AfxNova

' // Create an instance of the Msxml2 object
DIM pDisp AS CDispInvoke = "Msxml2.XMLHTTP.6.0"
pDisp.Invoke("open", "GET", "https://sourceforge.net/", FALSE)
pDisp.Invoke("Send")
DIM dwsResponse AS DWSTRING = pDisp.Get("ResponseText")
print dwsResponse
```
---

## SetLcid

Sets de locale identifier used by the class.

```
FUNCTION SetLcid (BYVAL lcid AS LCID)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lcid* | The locale context in which to interpret arguments. The *lcid* is used by the **GetIDsOfNames** function, and is also passed to **Invoke** to allow the object to interpret its arguments specific to a locale. Applications that do not support multiple national languages can ignore this parameter. |

---

## Put

Calls the specified property of an interface and sets the passed value.

```
PROPERTY Put (BYVAL dispID AS DISPID, BYREF dvArg AS DVARIANT)
PROPERTY Put (BYVAL pwszName AS WSTRING PTR, BYREF dvArg AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *dvArg* | DVARIANT. The value to set. |

---

## PutRef

Assigns a value to an interface member property that contains a reference to an object. For example, a reference to another interface.

```
PROPERTY PutRef (BYVAL dispID AS DISPID, BYREF dvArg AS DVARIANT)
PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYREF dvArg AS DVARIANT)
PROPERTY PutRef (BYVAL dispID AS DISPID, BYVAL pv AS ANY PTR)
PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYVAL pv AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *dvArg* | DVARIANT. The value to set. |
| *pv* | Pointer to an interface. |

---

## Set

Calls the specified property of an interface and sets the passed value.

```
FUNCTION Set (BYVAL dispID AS DISPID, BYREF dvArg AS DVARIANT) AS HRESULT
FUNCTION Set (BYVAL dispID AS DISPID, BYREF dvArg1 ... dvArg2 AS DVARIANT) AS HRESULT
FUNCTION Set (BYVAL dispID AS DISPID, BYREF dvArg1 ... dvArg3 AS DVARIANT) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dvArg AS DVARIANT) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dvArg1 ... dvArg2 AS DVARIANT) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dvArg1 ... dvArg3 AS DVARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *dvArg* | DVARIANT. The value to set. |
| *dvArg1...dvArg3* | DVARIANT. The last parameter is the value to set. dvArg1 and dvArg2 are the index values. |

#### Return value

S_OK (0) on success or an HRESULT code.

---

## SetRef

Assigns a value to an interface member property that contains a reference to an object. For example, a reference to another interface.

```
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dvArg AS DVARIANT) AS HRESULT
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dvArg1 ... dvArg2 AS DVARIANT) AS HRESULT
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dvArg1 ... dvArg3 AS DVARIANT) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dvArg AS DVARIANT) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dvArg1 ... dvArg2 AS DVARIANT) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dvArg1 ... dvArg3 AS DVARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *dvArg1...dvArg3* | DVARIANT. The last parameter is the value to set. dvArg1 and dvArg2 are the index values. |

#### Return value

S_OK (0) on success or an HRESULT code.

---
