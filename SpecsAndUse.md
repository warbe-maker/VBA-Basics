
# Specifics and usage examples for *mBasic* VBA services
provided in the [mBasic][1] Standard module, a component hosted [^1] in the [Basic.xlsb][2] Workbook which provided an elaborated regression test environment (both available for download on GitHub. See the [README][2] for a [summary of all services][3].  
This document supplements the [README Summary of services][4] 

## Specifics and usage of the *Align* services
#### Syntax: `Align(string[, align][, length][, fill][, margin][, col_arranged])`

| Argument    | Description                   | Default  |
|-------------|-------------------------------|----------|
|*string*     | Obligatory, string expression |
|*align*      | Enumerated alignment          |`enAlignLeft`|
|*length*     | Optional, integer valuse, final length of the result |`0`          |
|*fill*       | Optional, string expression   | `" "`         |
|*margin*     | Optional, string expression   |`vbNullString`|
|*colarranged*| Optional, boolean expression  |`False`       |

### Alignment with margins
When a margin is provided, the final length will be the specified length plus the length of a left and a right margin. A margin is typically used when the string is aligned as an Item of several items arranged in columns when the column delimiter is a vbNullString. When the column delimiter is a vertical bar a margin of a single space is the default *).
### Leading/trailing spaces and fills
The provided string may contain leading or trailing spaces. Leading spaces are preserved when the string is left aligned, trailing spaces are preserved when the string is aligned right. In any other case leading and trailing spaces are un-stripped.
### Examples of "simple" alignments
|Call                     |Returns     |
|:------------------------|------------|
|`Align("Abcde", enAlignLeft, 8, " -")`    |`"Abcde --"` |
|`Align("Abcde", enAlignRight, 8, "- ")`   |`"-- Abcde"` |
|`Align("Abcde", enAlignCentered, 8, " -")`|`" Abcde -"` |
|`Align("Abcde", enAlignCentered, 9, " -")`|`"- Abcde -"`|
|`Align("Abcde", enAlignRight, 7, "- ")`   |`"- Abcde"`  |        
|`Align("Abcde", enAlignCentered, 7, "-")` |`"-Abcde-"`  |
|`Align("Abcde", enAlignLeft, 7, "-")`     |`"Abcde--"`  |
|`Align("Abcde", enAlignRight, 6, "-")`    |`"-Abcde"`   |
|`Align("Abcde", enAlignCentered, 6, "-")` |`"Abcde-"`   |
|`Align("Abcde", enAlignLeft, 4, "-")`     |`"Abcd"`     |
|`Align("Abcde", enAlignRight, 4, "-")`    |`"Abcd"`     |
|`Align("Abcde", enAlignLeft, 4, "-", " ")`|`"Abcd"`     |

### Column arranged alignment
The function is also used to align items arranged in columns with the following specifics:  
- The provided length is regarded the maximum (when the provided string is longer it is truncated to the right).
- The final result string has any specified margin (left and right) added.
- When a fill is specified the final string has at least one added. As an example,  when the fill string is " -", the margin is a single space and the alignment is left, a string "xxx" is returned as " xxx -------- " a string "xxxxxxxxxx" is returned as " xxxxxxxxxx - "
- The provided length is the final length returned.
- Any specified margin is ignored.
- A specified fill is added only to end up with the specified length.

### Examples of column arranged alignments
|Call                     |Returns     |
|:------------------------|------------|
|`Align("Abcde", enAlignLeft, 8, " -", " ", True)`    |`" Abcde ---- "`  |
|`Align("Abcde", enAlignRight, 8, "- ", " ", True)`   |`" ---- Abcde "`  |
|`Align("Abcde", enAlignCentered, 8, " -", " ", True)`|`" -- Abcde --- "`|
|`Align("Abcde", enAlignCentered, 7, " -", " ", True)`|`" -- Abcde -- "` |
|`Align("Abcde", enAlignRight, 7, "- ", " ", True)`   |`" --- Abcde "`   |
|`Align("Abcde", enAlignCentered, 7, "-", " ", True)` |`" --Abcde-- "`   |
|`Align("Abcde", enAlignLeft, 7, "-", " ", True)`     |`" Abcde--- "`    |
|`Align("Abcde", enAlignRight, 6, "-", " ", True)`    |`" --Abcde "`     |
|`Align("Abcde", enAlignCentered, 6, "-", " ", True)` |`" -Abcde-- "`    |
|`Align("Abcde", enAlignLeft, 4, "-", " ", True)`     |`" Abcd- "`       |
|`Align("Abcde", enAlignRight, 4, "-", " ", True)`    |`" -Abcd "`       |
|`Align("Abcde", enAlignCentered, 4, "-", " ", True)` |`" -Abcd- "`      |

## Specifics and usage of the *Arry* service
*Arry*, implemented as Property Get/Let provides a universal array read/write service.  
- **WRITE** Returns the provided *array* with the provided item either simply added, when no *indices* are provided or having an item added (or replaced) at a given *index/indices*. The returned array may have expanded any dimension's **upper bound** (not only the last one!). The **lower bound** of the dimensions remain the same however.
- **READ** Returns from a provided *array* the item addressed by *indices*, with a default (defaults to `Empty`) for any not existing (i.e. out of bounds) item.

#### Syntax: `Arry(array[, indices][, default])`

| Argument   | Description |
|------------|-------------|
|*array*     | Obligatory, an existing, redim-ed or not, allocated or not, array.|
|*indices*   | Optional, a single integer, a string of indices delimited by a comma, or an Array or Collection of *indices*.|
|*default*   | Optional, defaults to Empty, returned for an not existing index/indices.|

### 1-dimensional array service
Write with the *indices* argument may be omitted. A yet un-dimension-ed and/or un-allocated *array* is returned with the first item added, an allocated *array* is returned with the new item added or added at given index (expanded on the fly).  
### Multi-dimensional arrays
- To write the first or any subsequent item to an allocated or un-allocated multi-dimensional *array* the provision of *indices* (one for each dimension) is obligatory.
- In contrast to VBA's `ReDim` statement this service is able to extend any dimension's **upper bound** while adding, writing, or updating an item. For re-specifying the **lower bound** and or the **upper bound** of <u>any</u> dimension see the *ArryReDim* service which also may  add dimensions (up to max 8).

### Usage example of the *Arry* service
Note: The examples use: ***ArryItems*** for the verification of the number of empty and non-empty items, and ***ArryDims*** to verify the array's number of dimensions.

```vb
Option Explicit

Private arrMy1DimArray As Variant
Private arrMy3DimArray As Variant

' ---------------------------------------------------------------------------
' Encapsulation of read from, write to arrMy1DimArray, Empty as explicit
' default for non-active items.
' ---------------------------------------------------------------------------
Private Property Get My1DimArray(ByVal m_indices As Variant) As Variant
    My1DimArray = Arry(arrMy1DimArray, m_indices, Empty)
End Property

Private Property Let My1DimArray(Optional ByVal m_indices As Variant, ByVal m_item As Variant)
    Arry(arrMy1DimArray, m_indices, Empty) = m_item
End Property

Private Sub AnyProc1()
' ---------------------------------------------------------------------------
' Usage example 1-dim array
' ---------------------------------------------------------------------------
    
    Set arrMy1DimArray = Nothing
    My1DimArray(10) = "any"          ' writes item number 10, items 1 to 9 remain Empty,
                                 ' establishes the array arrMy1DimArray with its first item
    Debug.Assert My1DimArray(10) = "any"
    My1DimArray(100) = "another"     ' writes item number 100, items 11 to 99 remain Empty
    Debug.Assert My1DimArray(100) = "another"
    My1DimArray(10) = "any-updated"  ' updates item number 10
    Debug.Assert My1DimArray(10) = "any-updated"
    Debug.Assert My1DimArray(2) = Empty
    
End Sub

Private Sub Example3DimArray1()
' ---------------------------------------------------------------------------
' Usage example 3-dim array, direct use of Arry service, un-Redim-ed
' ---------------------------------------------------------------------------
    
    Dim s As String
    Dim l As Long
    
    Set arrMy3DimArray = Nothing
    Debug.Assert ArryDims(arrMy3DimArray) = 0
    
    '~~ Write first item thereby establishing the array
    Arry(arrMy3DimArray, "5,5,2") = "Item(5,5,2)"
    s = Arry(arrMy3DimArray, "5,5,2")
    Debug.Assert s = "Item(5,5,2)"
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray, True) = 1
    
    Arry(arrMy3DimArray, "6,1,3") = "Item(6,1,3)"        ' writes subsequent item, expand upper bound of dim 1 and 3
    s = Arry(arrMy3DimArray, "6,1,3")
    Debug.Assert s = "Item(6,1,3)"
    Debug.Assert ArryItems(arrMy3DimArray, True) = 2
    l = ArryItems(arrMy3DimArray) ' 7 x 6 x 4 = 168 (Base 0)
    Debug.Assert l = 168
      
End Sub

Private Sub Example3DimArray2()
' ---------------------------------------------------------------------------
' Usage example 3-dim array, direct use of Arry service, pre-Redim-ed
' ---------------------------------------------------------------------------
    
    Dim s As String
    Dim l As Long
    
    Set arrMy3DimArray = Nothing
    ReDim arrMy3DimArray(2 To 3, 1 To 4, 1 To 3)
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray) = 24 ' 2 x 4 x 3
    
    '~~ Write first item thereby establishing the array
    Arry(arrMy3DimArray, "5,5,2") = "Item(5,5,2)"
    s = Arry(arrMy3DimArray, "5,5,2")
    Debug.Assert s = "Item(5,5,2)"
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray, True) = 1    ' count non-empty only
    l = ArryItems(arrMy3DimArray) ' 4 x 5 x 3 = 60
    Debug.Assert l = 60
    
    Arry(arrMy3DimArray, "6,1,3") = "Item(6,1,3)"        ' writes subsequent item, expand upper bound of dim 1 and 3
    s = Arry(arrMy3DimArray, "6,1,3")
    Debug.Assert s = "Item(6,1,3)"
    Debug.Assert ArryItems(arrMy3DimArray, True) = 2
    l = ArryItems(arrMy3DimArray) ' 5 x 5 x 3 = 75
    Debug.Assert l = 75
      
End Sub
```
## Specifics and usage of the *ArryDims* service
The service returns for a provided array the number of dimensions and optionally each dimension's bounds. For a yet not specified or Redim-ed array the number of dimensions returned = 0.

#### Syntax of *ArryDims*: `ArryDims(array[, dimspecs][, dimensions])`

| Argument   | Description |
|------------|-------------|
|*array*     | Variant expression, representing an allocated or not allocated array. When yet not Redim-ed the returned number of dimensions is 0  |
|*dimspecs*  | optional, a Collection holding the dimensions specifics|
|*dimensions*| optional, the number of dimensions returned |

### Usage example of the *ArryDims* service
```vb
    Dim arrMy    As Variant
    Dim cllSpecs As Collection
    Dim i        As Long
    Dim iDims    As Long
    Dim iLbnd    As Long
    Dim iUbnd    As Long

    ' Example 1: Makes use of the fact that the function
    '            returns 0 for a not allocated array
    '            (the loop will do nothing)
    For i = 1 To ArryDims(arrMy, cllSpecs)
        ' any code
    Next i
    
    ' Example 2: Obtains the bound for each processed dimension which
    '            instead of writing LBound(arrMy, i), UBound(arrMy, i)
    ArryDims arrMy, cllSpecs, iDims
    For i = 1 To iDims         ' iDims is same as cllSpecs.Count 
        iLbnd = cllSpecs(i)(1) ' same as iLbnd = LBound(arrMy, i)
        iUbnd = cllSpecs(i)(2) ' same as iUbnd = UBound(arrMy, i)
        ' any code
    Next i
```
## Specifics and use of the *Coll* service
*Coll*, implemented as a Property Get/Let provides a universal read/write service for Collections. It supports read/write with any index and automatically updates existing items. The recommended use is a Collection-specific Get/Let Property.

### Example for the use of the *Coll* service
```vb
Option Explicit
Private cllMyColl As Collection

' Encapsulation of Coll, Empty as default for non-active items, 
' reads from, writes to cllMyColl
Private Property Get MyColl(ByVal m_arg As Variant) As Variant
    MyCollection = Coll(cllMyColl, m_arg, Empty)
End Property

Private Property Let MyColl(Optional ByVal m_arg As Variant, ByVal m_item As Variant)
    Coll(cllMyColl, m_arg, Empty) = m_item
End Property

' Usage
Private Any Sub
    MyColl(10) = "any"          ' writes item number 10, items 1 to 9 with Empty,                                     ' establishes the Collection cllMyColl when yet not done 
    MyColl(100) = "another"     ' writes item number 100, items 11 to 99 with Empty
    MyColl(10) = "any-updated"  ' updates item number 10
End Sub
```

## Error handling, Application errors 
### Scheme
```vb
Private Sub MyService()
    Const PROC = "MyService"
    On Error Goto eh
    
    ' code
    
xt: Exit Sub    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
```

### Application error (AppErr)
```vb
Private Function AppErr(ByVal a_no As Long) As Long
    If a_no >= 0 _
    Then AppErr = a_no + vbObjectError _
    Else AppErr = Abs(a_no - vbObjectError)
End Function
```
### The source of an error (ErrSrc)
```vb
Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "<component-name>." & e_proc
End Function
```
### The common error message (ErrMsg)
```vb
Private Function ErrMsg(ByVal e_source As String, _
               Optional ByVal e_no As Long = 0, _
               Optional ByVal e_dscrptn As String = vbNullString, _
               Optional ByVal e_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service displaying:
' - the source (procedure) of the error
' - the kind of error (RunTime or Application)
' - the path to the error (when available)
' - a debugging option button
' - an "About:" section when the e_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' The display may optionally use the "Common VBA Message Service (fMsg/mMsg)
' when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by means of
' the VBA.MsgBox in case "mMsg" is not availablenot.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(e_source, e_no, e_dscrptn, e_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(e_source, e_no, e_dscrptn, e_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If e_no = 0 Then e_no = Err.Number
    If e_line = 0 Then ErrLine = Erl
    If e_source = vbNullString Then e_source = Err.Source
    If e_dscrptn = vbNullString Then e_dscrptn = Err.Description
    If e_dscrptn = vbNullString Then e_dscrptn = "--- No error description available ---"
    '~~ About
    ErrDesc = e_dscrptn
    If InStr(e_dscrptn, "||") <> 0 Then
        ErrDesc = Split(e_dscrptn, "||")(0)
        ErrAbout = Split(e_dscrptn, "||")(1)
    End If
    '~~ Type of error
    If e_no < 0 Then
        ErrType = "Application Error ": ErrNo = AppErr(e_no)
    Else
        ErrType = "VB Runtime Error ":  ErrNo = e_no
        If e_dscrptn Like "*DAO*" _
        Or e_dscrptn Like "*ODBC*" _
        Or e_dscrptn Like "*Oracle*" _
        Then ErrType = "Database Error "
    End If
    
    '~~ Title
    If e_source <> vbNullString Then ErrSrc = " in: """ & e_source & """"
    If e_line <> 0 Then ErrAtLine = " at line " & e_line
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")
    '~~ Description
    ErrText = "Error: " & vbLf & ErrDesc
    '~~ About
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

```
[^1]: The term *hosted* I use for all my Common Components to indicate the Workbook which is "the home" of the component, which means where the development and maintenance mainly happens and where the means for testing including regression testing are located. 

[1]: https://github.com/warbe-maker/VBA-Basics/blob/master/CompMan/source/mBasic.bas
[2]: https://github.com/warbe-maker/VBA-Basics#basic-vba-services
[3]: https://github.com/warbe-maker/VBA-Basics#summary-of-services
[4]: https://github.com/warbe-maker/VBA-Basics/blob/master/README.md#summary-of-services