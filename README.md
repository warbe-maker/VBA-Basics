# Basic VBA Services
A (personal) collection of basic services, collected over a long time,  used in many VB-projects. Some of them [^1] serve my personal common Excel VB-Project development needs.  
>Many of the services may raise an [application error](#application-error-apperr) when inadequately used. See also [Common error handling](#common-error-handling) together with a comprehensive error message `ErrMsg` service is used.   


# Services
|Service             |Kind&nbsp;[^2]|Description |
|---------------------|:------------:|-------------|
|***Align***          |F|Returns a provided string in a specified length, with optional margins (left and right) which defaults to none, aligned left, right or centered, with an optional fill string which defaults to spaces.<br>Specifics:<br>- When a margin is provided, the final length will be the specified length plus the length of a left and a right margin. A margin is typically used when the string is aligned as an Item of several items arranged in columns when the column delimiter is a vbNullString. When the column delimiter is a vertical bar a margin of a single space is the default *).<br>- The provided string may contain leading or trailing spaces. Leading spaces are preserved when the string is left aligned, trailing spaces are preserved when the string is aligned right. In any other case leading and trailing spaces are un-stripped.<br>- The function is also used to align items arranged in columns.<br> *) Column arranged option = TRUE (defaults to False):<br>- The provided length is regarded the maximum. I.e. when the provided string is longer it is truncated to the right.<br>- The final Result string has any specified margin (left and right) added.<br>- When a fill is specified the final string has at this one added. For example when the fill string is " -", the margin is a single space and the alignment is left, a string "xxx" is returned as " xxx -------- " a string "xxxxxxxxxx" is returned as: " xxxxxxxxxx - " Column arranged option = FALSE (the default):<br>- The provided length is the final length returned.<br>- Any specified margin is ignored.<br>- A specified fill is added only to end up with the specified length.|
|***AlignCntr***       |S|Called by ***Align*** or directly.|
|***AlignLeft***       |S|Called by ***Align*** or directly.|
|***AlignRght***       |S|Called by ***Align*** or directly.|
|***AppErr***          |S|Ensures that the number of a programmed ***Application-Error*** never conflicts with a _VB-Runtime-Error_ by adding the constant _vbObjectError_ which turns it into a negative value. In return, in an error message, it turns the negative _Application-Error_ number back into its original positive number and identifies the error as an application error.[^3]|
|***AppIsInstalled***  |F|Returns TRUE when an application identified by its .exe is installed, i.e. available in the systems path.|
|***Arry***            |P|**Syntax**: `Arry(*array*[, *indices*][, *default*])`<br>Provides a common, universal array READ and array WRITE service.<br>**The WRITE service**: Returns the provided *array* with the provided item either:<br>- Simply added, when no *indices* are provided<br>- Having added or replaced the item at given *indices* by considering that the returned array has new from/to specifics for any of its dimensions at any level whereby the from specific for any dimension remains the same<br>- Having created a 1 to 8 dimensions *array* depending on the provided *indices* with the provided item added or replaced<br>- Having re-dimension-ed the provided/input *array* to the provided *indices* with the provided item added or updated.<br>The *indices* may be provided as a single integer, indicating that the *array* is or will be 1-dimensional<br>- a string of indices delimited by a comma, indicating that the provided or returned *array* is multi-dimensional<br>- an Array or Collection of indices.<br>Note: In contrast to VBA's ReDim statement this service is able to extend any "to" specification of any dimension (not only the last one) with the only constraint that the "from" specification of any dimension will remain the same (see ***ArryReDim*** for re-specifying any of the dimensions ans even adding new dimensions).<br>Constraints: - For a yet not dimension-ed and/or not allocated *array* items may be added by simply not specifying an index<br>- For an already dimension-ed and/or allocated *array* the provision of an index for each of its dimensions is obligatory.<br>**The READ service**:<br>Returns from a provided *array* the item addressed by *indices*, with a default (defaults to Empty) for any Empty item.|
|***ArryAsDict***      |F|Returns a Dictionary with all items of a provided array, each with a key compiled from the indices delimited by a comma.|
|***ArryAsRng***       |S|**Syntax:** `ArryAsRng *array*, *range*, *transposed*`<br>Copies the content of a one or two dimensional ***array*** to a provided ***range***, optionally ***transposed*** (defaults to False).|
|***ArryCompare***     |F  |Returns a Dictionary with the provided number of lines/items (defaults to all) which differ between two one dimensional arrays. When no difference is encountered the returned Dictionary is empty (Count = 0).|
|***ArryDiffers***     |F|Returns TRUE when two arrays are different (stops with the first difference).|
|***ArryDims***        |F|**Syntax**: `ArryDims(*array*[, *specifics*][, *dimensions*])`<br>Returns the number of ***dimensions*** and each dimension's bound ***specifics***. An un-allocated or empty ***array*** returns 0 dimensions and empty ***specifics***.<br>See [usage examples](#examples-arrydims).|
|***ArryErase***       |S|**Syntax**: `ArryErase *array*, *array* , ...`<br>Sets any provided *array* to Empty. The *array's* specifics regarding dimensions and bounds remain intact.|
|***ArryIndices***     |F|Returns provided *indices* as Collection whereby the *indices* may be provided as integers (one for each dimension), as an array of integers, or as a string of integers delimited by a , (comma). This is a "helper-function" applicable wherever array indices cannot be provided as ParamArray, when indices are a Property argument for instance.|
|***ArryItems***       |F|**Syntax**: `ArryItems(*array*, *include_empty*)`<br>Returns the number of items/elements in a multi-dimensional array or a nested array, by default (*include_empty* = False) excluding Empty items. Note: Nested arrays are items in an array which are an array (also called a jagged array). An un-allocated array returns 0.|
|***ArryIsAllocated*** |F|Returns TRUE when the provided array has at least one item.|
|***ArryNextIndex***   |F|**Syntax**: `ArryNextIndex(*array*, *indices*)`<br>Returns for a provided *array* and given *indices* the logically next and TRUE when there is a next one, i.e. the provided *indices* are not indicating each dimensions upper bound.|
|***ArryReDim***       |S|Returns a provided multi-dimensional *array* with new *dimension_specs* whereby any - not only the last dimension) may be redim-ed. The new *dimension_specs* are provided as strings following the format: "<dimension>:<from>,<to>" whereby <dimension> is either addressing a dimension in the current *array*, i.e. before the redim has taken place, or a + for a new dimension. Since only new or dimensions with changed from/to specs are provided the information will be used to compile the final redim-ed *array*'s dimensions which must not exceed 8.<br>Requires References to: "Microsoft Scripting Runtime" and "Microsoft VBScript Regular Expressions 5.5".|
|***ArryRemoveItems*** |S|**Syntax**: `ArryRemoveItems *array*[, *item*][, *index*][, *number*]`<br>Returns a 1-dimensional *array* with a given *number* of items removed, beginning with a starting *item* number or an *index*. Any inappropriate provision of arguments raises an [application error](##application-error-apperr). When the last Item in an array is removed the returned array is returned Empty (same as erased) which means it is no longer allocated.|
|***ArryTrim***        |S|Returns a provided array with items leading and trailing spaces removed, Empty items removed and items with a vbNullString removed. Any items with `vbCr`, `vbCrLf`, `vbLf` are ignored/kept.|
|***ArryUnload***      |S|Returns a 2-dim array with all items of a multi-dimensional (max 8 dimensions) array unloaded, whereby the first item is the indices delimited by a comma and the second item is the array's item. I.e. a multi-dimendional array is unloaded in a "flat" 2-dim array.|
|***ArryUnloadToDict***|S|**Syntax**: `ArryUnloadToDict *array*, *dictionary*[, *indices*]`<br>Returns all items in a multi-dimensional *array* as *dictionary* with the items as the array's item and the array item's indices delimited by a comma as key. The procedure is called recursively for nested arrays collecting the *indices*.|
|***BoP/EoP***         |S|Common **B**egin-**o**f-**P**rocedure/**E**nd-**o**f-**P**rocedure  interface for the optional [Common VBA Execution Trace Service][3]. Obligatory copy Private for any VB-Component using the ***[Common VBA Execution Trace Service][3] but not having the mBasic common component installed. 1), 2).|
|***BoC/EoC***         |S|Analog to the above for to-be-traced code sequences  1), 2) |
|***Center***          |F|Returns a string centered within a string with of a given length.|
|***CleanTrim***       |F|Returns a provide string cleaned from any non-printable characters.|
|***Coll***            |P|**Syntax**: `Coll(*collection*[, *argument][, *default*]`<br>Universal **READ** from and **WRITE** to a *collection*.<br>- **READ**: When the provided *argument* is an integer which addresses a not existing index or an index of which the element = Empty, Empty is returned, else the element's content. When the provided *argument* is not an integer the index of the first element which is identical with the provided argument is returned. When no item is identical, a *default* - which defaults to Empty - is returned.<br>- **WRITE**/Let: Writes items with any index by filling/adding the gap with Empty items.|
|***ErrMsg***          |F|Universal error message display service. Obligatory Private copy for any VB-Component using the common error service but not having this component installed. The function displays:<br>- a debugging option allowing to resume the error raising code line<br>- an optional additional "About:" section when the provided error description has an additional string concatenated by two vertical bars (\|\|)<br>- the error message by means of the [Common VBA Message Service (fMsg/mMsg)][1] when installed, and active (Cond. Comp. Arg. `mMsg = 1`) else by means of `VBA.MsgBox`. The function uses ***AppErr*** to identify/distinguish programmed application errors (`Err.Raise AppErr(n), ....`) and turn them back into their positive number. The function is obligatory as `Private` copy for any VB-Component using the ***[Common VBA Execution Trace Service][3]*** and/or the ***[Common VBA Error Services][2]*** but not having the mBasic common component installed. 1), 2).|
|***Max***              |F|Returns the maximum of provided arguments which may be a numeric value, a string, an Arrays, or a Collections. Of strings items the length is considered the numeric value. When an argument is an Array or a Collection the items again may be a numeric value, a string, an Arrays, or a Collection.<br>**Constraint:**  When an argument is an Array or a Collection and it contains a single item which again is an Array or a Collection, an [application error](#application-error-apperr) is raised. Nested Arrays and/or Collections may again have items which are an Array or a Collection but only among other numeric or string items. This constraint is caused by the fact that function calls with a single argument which is an Array or a Collection are considered recursive calls and thus will result in a loop.|
|***Min***              |F|Returns the minimum value of any number of values provided.|
|***KeySort***          |F|Returns a provided Dictionary sorted by key. |
|***PointsPerPixel***   |F|Return the DPI of the current display/monitor.|
|***README***           |S|Displays a given url with a given bookmark in the computer's default browser. The url defaults to this 'component's README url in the public GitHub repo, the bookmark defaults to vbNullString.|
|***SelectFolder***     |||
|***ShellRun***         |S|Opens a folder, an email-app, a url, an Access instance, etc. all by means of the default application. Courtesy of: Dev Ashish.<br>Example:<br>`ShellRun "https://github.com/warbe-maker/VBA-Basics"`<br> displays this README in this _Common Component's_ public GiHub repo.|
|***Spaced***           ||Returns a provided string 'spaced' with non-braking spaces. Spaces already in the provided string are doubled.|
|***SysFrequency***     |F|Returns the current number of ticks per seconds which is the precision for the ***TimeBegin*** end ***TimerEnd*** function.|
|***TimedDoEvents***    |S|Debug.Print the time delay of a performed `DoEvents`.<br>**Background:** `DoEvents` are often recommended to solve un-explainable problems. However, according to the `DoEvents` documentation the way how and why it may solve a problem is pretty miraculous. Anyhow, even when it helped it should be taken into consideration that `DoEvents`  enable keyboard interactions while a process executes. In case of an endless loop with embedded `DoEvents` this may be a godsend. But it as well may cause unpredictable timing  results. When, instead of performing `DoEvents` directly `TimedDoEvents` is used at least the resulting performance delay in milliseconds is printed to the _VBE Immediate Window_.|
|***TimerBegin***       |F|Returns the current system ticks as the start of a timer. |
|***TimerEnd***         |F|Returns, based on a provided ***TimerBegin***:<br>'- the end-ticks<br>- the elapsed ticks<br>- the elapsed time in the provided format which defaults to `"hh:mm:ss.0000"`.)|

# Examples
## Examples ArryDims
```vb
    Dim cllSpecs As Collection
    Dim lDims    As Long
    Dim i        As Long
    
    ' Example 1
    For i = 1 To ArryDims(arr)
        ' any processing
    Next i
    ' Example 2
    ArryDims arr, cllSpecs, lDims
    For i = 1 To lDims
        ' cllSpecs(i)(1) = LBound of the dimension
        ' cllSpecs(i)(2) = Ubound of the dimension
    Next i
```

## Example ...

## The error handling of services
### Scheme
```vb
Private Sub MyService()
    Const PROC = "MyService"
    On Error Goto eh
    
xt: Exit Sub    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
```

### Used procedures 
#### Application error (AppErr)
```vb
Private Function AppErr(ByVal a_no As Long) As Long
    If a_no >= 0 _
    Then AppErr = a_no + vbObjectError _
    Else AppErr = Abs(a_no - vbObjectError)
End Function
```

#### The source of an error (ErrSrc)
```vb
Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "<component-name>." & e_proc
End Function
```

#### The common error message (ErrMsg)
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


[^1]: It goes without saying that my VB-Projects use this _mBasic_ component. However, all my _Common Components_ use some services as `Private Sub` copy. This keeps them 100% autonomous, i.e. independent from this and other components but still serve my personal use of them. The service i am talking about are:  
-&nbsp;***BoP/EoP***, and ***ErrMsg*** to keep them independent from the ***[Common VBA Error Services][2]***  
-&nbsp;***BoP/EoP***, ***BoC/EoC*** to keep the use of the ***[Common VBA Execution Trace Service][3]*** optional

[^2]: S=Sub, F=Function, P=Property (r=read/Get, w=write/Let)

[^3]: ***AppErr*** is used with the ***ErrMsg*** service (and the ***[Common VBA Error Services][2]***). The unique identification of the error causing procedure allows _Application Error Numbers_ starting from 1 to n in each procedure! No need for a VB-Project global list of the used _Application Error Numbers_. The number only has a meaning within the procedure it is used.

[1]: https://github.com/warbe-maker/VBA-Message
[2]: https://github.com/warbe-maker/VBA-Error
[3]: https://github.com/warbe-maker/VBA-Trace
[4]: https://github.com/warbe-maker/VBA-Basics