<style>
.syntax {
    font-size: 18px;
    font-family: monospace;
}
.syntax i {
    font-style: italic;
}
</style>
# Specifications and usage examples of basic services
All services are provided in the [mBasic][1] Standard module. The component is hosted in the [Basic.xlsb][2] Workbook which provided an elaborated regression test environment (bot available for download on GitHub.

## The *Arry* service
Universal array read/write service.  
- **WRITE** Returns the provided *array* with the provided item either simply added, when no *indices* are provided or having an item added (or replaced) at a given *index/indices*. The returned array may have expanded any dimension's **upper bound** (not only the last one!). The **lower bound** of the dimensions remain the same however.
- **READ** Returns from a provided *array* the item addressed by *indices*, with a default (defaults to `Empty`) for any not existing (i.e. out of bounds) item.

### Syntax of the *Arry* service
**Syntax**: `Arry(array[, indices][, default])`<br>
<pre class="syntax">
<b>Arr</b>(<i>array</i>[, <i>indices</i>][, <i>default</i>])
</pre>

| Argument   | Description |
|------------|-------------|
|*array*     | An existing, Redim-ed or not, allocated or not, array.|
|*indices*   | A single integer, a string of indices delimited by a comma, or an Array or Collection of *indices*.|
|*default*   | Optional, defaults to Empty, returned for an not existing index/indices.|

## Specifics of the *Arry* service
### 1-dimensional arrays
Write with the *indices* argument may be omitted. A yet un-dimension-ed and/or un-allocated *array* is returned with the first item added, an allocated *array* is returned with the new item added or added at given index (expanded on the fly).  
### Multi-dimensional arrays
- To write the first or any subsequent item to an allocated or un-allocated multi-dimensional *array* the provision of *indices* (one for each dimension) is obligatory.
- In contrast to VBA's `ReDim` statement this service is able to extend any dimension's **upper bound** while adding, writing, or updating an item. For re-specifying the **lower bound** and or the **upper bound** of <u>any</u> dimension see the *ArryReDim* service which also may  add dimensions (up to max 8).

### Example Arry

## The *ArryDims* service
Returns the number of dimensions and each dimension's bounds for a provided array - not necessarily allocated.

### Syntax of ArryDims
<pre class="syntax">
<b>ArryDims</b>(<i>array</i>[, <i>dimspecs</i>][, <i>dimensions</i>])
</pre>

| Argument   | Description |
|------------|-------------|
|*array*     | a Variant representing an allocated or not allocated array. When yet not Redim-ed the returned number of dimensions is 0  |
|*dimspecs*  | optional, a Collection holding the dimensions specifics|
|*dimensions*| optional, the number of dimensions returned |

### Example ArryDims
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
## Coll (Property Get/Let)
Universal read/write service for Collections. Supports read/write with any index and automatically updates existing items. The recommended use is a Collection-specific Get/Let Property.

### Coll usage (example)
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

[1]: https://github.com/warbe-maker/VBA-Basics/blob/master/CompMan/source/mBasic.bas
[2]: https://github.com/warbe-maker/VBA-Basics