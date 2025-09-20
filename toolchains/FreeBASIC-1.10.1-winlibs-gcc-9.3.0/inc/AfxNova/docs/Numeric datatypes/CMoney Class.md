# CMoney Class

`CMoney` is a wrapper class for the `CURRENCY` data type. `CURRENCY` is implemented as an 8-byte two's-complement integer value scaled by 10,000. This gives a fixed-point number with 15 digits to the left of the decimal point and 4 digits to the right. The `CURRENCY` data type is extremely useful for calculations involving money, or for any fixed-point calculations where accuracy is important.

The `CMoney` wrapper implements arithmetic, assignment, and comparison operations for this fixed-point type, and provides access to the numbers on either side of the decimal point in the form of two components: an integer component which stores the value to the left of the decimal point, and a fractional component which stores the value to the right of the decimal point. The fractional component is stored internally as an integer value between -9999 (CY_MIN_FRACTION) and +9999 (CY_MAX_FRACTION). The function `GetFraction` returns a value scaled by a factor of 10000 (CY_SCALE).

When specifying the integer and fractional components of a `CMoney` object, remember that the fractional component is a number in the range 0 to 9999. This is important when dealing with a currency such as the US dollar that expresses amounts using only two significant digits after the decimal point. Even though the last two digits are not displayed, they must be taken into account.

---

### Constructors

Create a new instance of the `CMoney` class and assigns the values passed to it.

```
CONSTRUCTOR CMoney
CONSTRUCTOR CMoney (BYREF cSrc AS CMoney)
CONSTRUCTOR CMoney (BYVAL cySrc AS CURRENCY)
CONSTRUCTOR CMoney (BYVAL nInteger AS LONGLONG)
CONSTRUCTOR CMoney (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
CONSTRUCTOR CMoney (BYVAL dSrc AS DOUBLE)
CONSTRUCTOR CMoney (BYVAL dSrc AS DECIMAL)
CONSTRUCTOR CMoney (BYREF szSrc AS STRING)
CONSTRUCTOR CMoney (BYVAL varSrc AS VARIANT)
```

#### Examples

```
DIM c AS CMoney = 123
DIM c AS CMoney = CMoney(123, 45)
DIM c AS CMoney = 12345.1234
DIM c AS CMoney = "77777.999"
```
#### Remark

The constructor that accepts a DOUBLE value is particularly useful, because it allows to set the integer and fractional parts at the same time with just a number, e.g.

```
DIM c AS CMoney = 12345.1234
```

#### Testing code:

```
#INCLUDE ONCE "AfxNova/CMoney.inc"
USING AfxNova

DIM c AS CMoney = 12345.1234
print c
c = c + 111.11
print c
c = c - 111.11
print c
c = c * 2
print c
c = c / 2
print c
c += 123
print c
c -= 123
print c
c *= 2.3
print c
c /= 2.3
print c
c = c ^ 2
print c
c = SQR(c)
print c
DIM c2 AS CMoney = c
print c2
DIM c3 AS CMoney = c * 2
print c3
DIM c4 AS CMoney = c3 / 2
print c4
DIM c5 AS CMoney = "1234.789"
print c5
DIM c6 AS CMoney
c6 = "77777.999"
print c6
DIM c7 AS CMoney
c7 = c6
print c7
DIM cl AS CMoney = 3
cl = LOG(cl)
print cl
DIM v AS VARIANT = cl
dim cv AS CMoney = v
print cv
print "--------------"
DIM cx AS CMoney
FOR i AS LONG = 1 TO 1000000
   cx += 0.0001
NEXT
PRINT "0.0001 added 1,000,000 times = "; cx
```
---

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#operator1) | Assigns a value to a **CMoney** variable. |
| [CAST operators](#operator2) | Converts a **CMoney** into another data type. |
| [Operator \*](#operator3) | Returns the address of the underlying **CURRENCY** structure. |
| [Comparison operators](#operator4) | Compares currency numbers. |
| [Math operators](#operator5) | Add, subtract, multiply or divide currency numbers. |

---

### Methods

| Name       | Description |
| ---------- | ----------- |
| [FormatCurrency](#formatcurrency) | Formats a currency into a string form. |
| [FormatNumber](#formatnumber) | Formats a currency into a string form. |
| [GetFraction](#getfraction) | Returns the fractional component of a currency value. |
| [GetInteger](#getinteger) | Returns the integer component of a currency value. |
| [Round](#round) | Rounds the currency to a specified number of decimal places. |
| [SetValues](#setvalues) | Sets the integer and fractional components. |
| [ToVar](#tovar) | Returns the currency as a VT_CY variant. |

---

## <a name="operator1"></a>Operator LET (=)

Assigns a value to a `CMoney` variable.

#### Examples

```
DIM c AS CMoney
c = 123
c = 12345.1234
c = "77777.999"
```

Passing a DOUBLE value allows to set the integer and fractionary parts at the same time with just a number, e.g.

```
DIM c AS CMoney = 12345.1234
```
---

## <a name="operator2"></a>CAST Operators

Converts a `CMoney` into another data type.

```
OPERATOR CAST () AS DOUBLE
OPERATOR CAST () AS STRING
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS VARIANT
```

These operators aren't called directly, They perform the conversion when the target is not a CMoney variable. For example, when assigning a `CMoney` variable to an standard numeric variable, the CAST() AS DOUBLE operator is implicity called.

---

## <a name="operator3"></a>Operator *

Returns the address of the underlying `CURRENCY` structure.

```
OPERATOR * (BYREF cur AS CMoney) AS CURRENCY PTR
```
---

## <a name="operator4"></a>Comparison operators

```
OPERATOR = (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
OPERATOR <> (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
OPERATOR < (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
OPERATOR > (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
OPERATOR <= (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
OPERATOR >= (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
```

#### Examples

```
DIM c AS CMoney = 1.23
IF c = 1.23 THEN PRINT "equal" ELSE PRINT "different"
DIM c2 AS CMoney = 1.23
IF c = c2 THEN PRINT "equal" ELSE PRINT "different"
DIM c3 AS SINGLE = 1.23
IF c = c3 THEN PRINT "equal" ELSE PRINT "different"
```
---

## <a name="operator5"></a>Math operators

```
OPERATOR + (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
OPERATOR += (BYREF cur AS CMoney)
OPERATOR - (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
OPERATOR -= (BYREF cur AS CMoney)
OPERATOR * (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
OPERATOR *= (BYREF cur AS CMoney)
OPERATOR / (BYREF cur AS CMoney, BYVAL cOperand AS CMoney) AS CMoney
OPERATOR /= (BYREF cOperand AS CMoney)
```

#### Examples

```
DIM c AS CMoney = 12345.1234
c = c + 111.11
c = c - 111.11
c = c * 2
c = c / 2
c += 123
c -= 123
c *= 2.3
c /= 2.3
```

#### Remarks

The version OPERATOR - (BYREF cur AS CMoney) AS CMoney changes the sign of a currency number.

```
DIM c AS CMoney = 123
c = -c
```

Or you can use the `NOT` operator:

```
DIM c AS CMoney = 123
c = NOT c
```

Other FreeBasic operators such `AND`, `MOD`, `OR`, `SHL` and `SHR` can also be used with `CMoney` variables, e.g. c = c AND 1, c = c MOD 5, etc.

---

## <a name="formatcurrency"></a>FormatCurrency

Formats a currency into a string form.

```
FUNCTION FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br> 0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set.  |

#### Return value

A string containing the formatted value.

#### Remarks

Same as **FormatNumber** but adding the currency symbol.

#### Example

```
DIM cur AS CMoney = 12345.1234
PRINT cur.FormatCurrency   --> 12.345,12 € (Spain)
```
---

## <a name="formatnumber"></a>FormatNumber

Formats a currency into a string form.
```
FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br> 0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set.  |

#### Return value

A string containing the formatted value.

#### Remarks

Same as **FormatCurrency** but without adding the currency symbol.

#### Example

```
DIM cur AS CMoney = 12345.1234
PRINT cur.FormatNumber   --> 12.345,12 (Spain)
```
---

## <a name="getfraction"></a>GetFraction

Returns the fractional component of a currency value.

```
FUNCTION GetFraction () AS SHORT
```
---

## <a name="getinteger"></a>GetInteger

Returns the integer component of a currency value.

```
FUNCTION GetInteger () AS LONGLONG
```
---

## <a name="round"></a>Round

Rounds the currency to a specified number of decimal places.

```
FUNCTION Round (BYVAL nDecimals AS LONG) AS CMoney
```
---

## <a name="setvalues"></a>SetValues

Sets the integer and fractional components.

```
FUNCTION SetValues (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nInteger* | The value to be assigned to the integer component of the m_currency data member. The sign of the integer component must match the sign of the existing fractional component. *nInteger* must be in the range CY_MIN_INTEGER to CY_MAX_INTEGER inclusive. |
| *nFraction* | The value to be assigned to the fractional component of the m_currency data member.The sign of the fractional component must the same as the integer component, and the value must be in range -9999 (CY_MIN_FRACTION) to +9999 (CY_MAX_FRACTION). |

#### Remarks

Based on 4 digits. To set .2, pass 2000, to set .0002, pass a 2.

---

## <a name="tovar"></a>ToVar

Returns the currency as a VT_CY variant.

```
FUNCTION ToVar () AS VARIANT
```

#### Remarks

Can be used to assign a currency directly to a VT_CY `VARIANT`

```
DIM c AS CMoney = 12345.1234
DIM v AS VARIANT = c.ToVar
```
or to a `DVARIANT`
```
DIM c AS CMoney = 12345.1234
DIM dv AS DVARIANT = c.ToVar
```
---
