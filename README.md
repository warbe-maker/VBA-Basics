# Basic VBA Services
A collection of personnel/basic VBA services, collected over time and used in many VB-projects. Some of them [^1] serve my personal common Excel VB-Project development needs.  
>Many of the services may raise an [application error][8] when inadequately used which displays a comprehensive error message (see [Common error handling][7]). The used procedures are part of these basic VBA services. This document is supplemented by [Specifics and usage examples for *mBasic* VBA services][]  

# Installation and use
All services are provided in the Standard module [mBasic][9] which may be copied into an VB-Project. The component is hosted [^1] in the [Basic.xlsb][4] Workbook which provided an elaborated regression test environment (bot available for download on GitHub.

# Summary of services
## Array services
|Service             |Kind&nbsp;[^2]|Description |
|---------------------|:------------:|-------------|
|***Arry***            |P|Common, universal array READ and WRITE service supporting automated upper bound expansion for multi-dimensional array. See    |
|***ArryAsDict***      |F|Returns a Dictionary with all items of a provided array, each with a key compiled from the indices delimited by a comma.|
|***ArryAsRng***       |S|**Syntax:** `ArryAsRng *array*, *range*, *transposed*`<br>Copies the content of a one or two dimensional ***array*** to a provided ***range***, optionally ***transposed*** (defaults to False).|
|***ArryCompare***     |F  |Returns a Dictionary with the provided number of lines/items (defaults to all) which differ between two one dimensional arrays. When no difference is encountered the returned Dictionary is empty (Count = 0).|
|***ArryDiffers***     |F|Returns TRUE when two arrays are different (stops with the first difference).|
|***ArryDims***        |F|Returns the number of *dimensions* and optionally each dimension's bounds for a provided allocated or un-allocated array. For a yet not Redim-ed array 0 dimensions is returned. See [usage examples][5].|
|***ArryErase***       |S|**Syntax**: `ArryErase *array*, *array* , ...`<br>Sets any provided *array* to Empty. The *array's* specifics regarding dimensions and bounds remain intact.|
|***ArryIndices***     |F|Returns provided *indices* as Collection whereby the *indices* may be provided as integers (one for each dimension), as an array of integers, or as a string of integers delimited by a , (comma). This is a "helper-function" applicable wherever array indices cannot be provided as ParamArray, when indices are a Property argument for instance.|
|***ArryItems***       |F|**Syntax**: `ArryItems(array, defaultsexcluded)`<br>Returns the number of items in a multi-dimensional array or a nested array. The latter is an array of which one or more items are again possibly multi-dimensional arrays. An un-allocated array returns 0. When *defaultsexcluded* = True, only "active= items/elements are counted.|
|***ArryIsAllocated*** |F|Returns TRUE when the provided array has at least one item.|
|***ArryNextIndex***   |F|**Syntax**: `ArryNextIndex(*array*, *indices*)`<br>Returns for a provided *array* and given *indices* the logically next and TRUE when there is a next one, i.e. the provided *indices* are not indicating each dimensions upper bound.|
|***ArryReDim***       |S|Returns a provided multi-dimensional *array* with new *dimension_specs* whereby any - not only the last dimension) may be redim-ed. The new *dimension_specs* are provided as strings following the format: "<dimension>:<from>,<to>" whereby <dimension> is either addressing a dimension in the current *array*, i.e. before the redim has taken place, or a + for a new dimension. Since only new or dimensions with changed from/to specs are provided the information will be used to compile the final redim-ed *array*'s dimensions which must not exceed 8.<br>Requires References to: "Microsoft Scripting Runtime" and "Microsoft VBScript Regular Expressions 5.5".|
|***ArryRemoveItems*** |S|**Syntax**: `ArryRemoveItems *array*[, *item*][, *index*][, *number*]`<br>Returns a 1-dimensional *array* with a given *number* of items removed, beginning with a starting *item* number or an *index*. Any inappropriate provision of arguments raises an [application error][8]. When the last Item in an array is removed the returned array is returned Empty (same as erased) which means it is no longer allocated.|
|***ArryTrim***        |S|Returns a provided array with items leading and trailing spaces removed, Empty items removed and items with a vbNullString removed. Any items with `vbCr`, `vbCrLf`, `vbLf` are ignored/kept.|
|***ArryUnload***      |S|Returns a 2-dim array with all items of a multi-dimensional (max 8 dimensions) array unloaded, whereby the first item is the indices delimited by a comma and the second item is the array's item. I.e. a multi-dimensional array is unloaded in a "flat" 2-dim array.|
|***ArryUnloadToDict***|S|**Syntax**: `ArryUnloadToDict array, dictionary[, indices]`<br>Returns all items in a multi-dimensional *array* as *dictionary* with the items as the array's item and the array item's indices delimited by a comma as key. The procedure is called recursively for nested arrays collecting the *indices*.|

## Error handling/display services
The services are used in all my VB-projects. In Common Components they are used as Private .... copies in order to support their autonomy, i.e. the quality of providing their common services without the involvement of other components.

|Service               |Kind&nbsp;[^2]| Description  |
|----------------------|:------------:|-------------|
|***AppErr***          |F|Returns a positive [application error][8] integer value as a negative value which not conflicts with any _VB-Runtime-Error_ by having the constant _vbObjectError_ added. The ***ErrMsg*** service identifies it as an [application error][8] and uses the service to turn it back into its origin  positive [application error][8] number. In conjunction with the unique identification of the error causing procedure these application error number can range in any procedure from 1 to n. I.e. there is no need for a VB-Project global list of the used _Application Error numbers_.|
|***BoP/EoP***         |S|Indicates the <u>B</u>egin-<u>o</u>f-a-<u>P</u>rocedure and the <u>E</u>nd-<u>o</u>f-a-<u>P</u>rocedure. Used by the [Common VBA Message Service (fMsg/mMsg)][1] to display the path to the error. The indication is also used by the ***[Common VBA Execution Trace Service][3] - when installed.|
|***BoC/EoC***         |S|Analog to the above for to-be-execution-traced code sequences (only relevant when the ***[Common VBA Execution Trace Service][3]*** is installed and used.|
|***ErrMsg***          |F|Universal error message display service. Specifically as a Private copy in any VB-Component using the common error service but not having this component (mErH) installed. The function displays:<br>- a debugging option allowing to resume the error raising code line<br>- an optional additional "About:" section when the provided error description has an additional string concatenated by two vertical bars (\|\|)<br>- the error message by means of the [Common VBA Message Service (fMsg/mMsg)][1] when installed, and active (Cond. Comp. Arg. `mMsg = 1`) else by means of `VBA.MsgBox`. The function uses ***AppErr*** to identify/distinguish programmed [application errors][8] (`Err.Raise AppErr(n), ....`) and turn them back into their positive number. The function is obligatory as `Private` copy for any VB-Component using the ***[Common VBA Execution Trace Service][3]*** and/or the ***[Common VBA Error Services][2]*** but not having the mBasic common component installed. 1), 2).|
|***ErrSrc***         |F|Returns the full name of a provided procedure's name in the form<br>`<comp-name>.<proc-name>`<br>`Private Function ErrSrc(ByVal e_proc As String) As String`<br>&nbsp;&nbsp;&nbsp;&nbsp;`ErrSrc = "<compinent-name-goes-here>." & e_proc`<br>`End Function`<br>Note: This is available in the mBasic component as `Private Function` and for being used as such in any components using the common error handling service.|

 See also [Common error handling and display service][7].
 
## Alignment of string services
See ... for details.
|Service             |Kind&nbsp;[^2]|Description |
|---------------------|:------------:|-------------|
|***Align***          |F|Returns a provided string in a specified length, with optional margins, aligned left, right or centered, optionally filled which defaults to spaces.|
|***AlignCntr***       |S|Called by ***Align*** or directly.|
|***AlignLeft***       |S|Called by ***Align*** or directly.|
|***AlignRght***       |S|Called by ***Align*** or directly.|

## Other services
|Service             |Kind&nbsp;[^2]|Description |
|---------------------|:------------:|-------------|
|***AppIsInstalled***  |F|Returns TRUE when an application identified by its .exe is installed, i.e. available in the systems path.|
|***Center***          |F|Returns a string centered within a string with of a given length.|
|***CleanTrim***       |F|Returns a provide string cleaned from any non-printable characters.|
|***Coll***            |P|**Syntax**: `Coll(collection[, argument][, default]`<br>Universal **READ** from and **WRITE** to a *collection*.<br>- **READ**: When the provided *argument* is an integer which addresses a not existing index or an index of which the element = Empty, Empty is returned, else the element's content. When the provided *argument* is not an integer the index of the first element which is identical with the provided argument is returned. When no item is identical, a *default* - which defaults to Empty - is returned.<br>- **WRITE**/Let: Writes items with any index by filling/adding the gap with Empty items.|
|***Max***              |F|Returns the maximum of provided arguments which may be a numeric value, a string, an Arrays, or a Collections. Of strings items the length is considered the numeric value. When an argument is an Array or a Collection the items again may be a numeric value, a string, an Arrays, or a Collection.<br>**Constraint:**  When an argument is an Array or a Collection and it contains a single item which again is an Array or a Collection, an [application error][8] is raised. Nested Arrays and/or Collections may again have items which are an Array or a Collection but only among other numeric or string items. This constraint is caused by the fact that function calls with a single argument which is an Array or a Collection are considered recursive calls and thus will result in a loop.|
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


[^1]: It goes without saying that my VB-Projects use this _mBasic_ component. However, all my _Common Components_ use any of the provided services as `Private ...` copy in order to keep them 100% autonomous, i.e. independent from this and other components. The service copied most are: ***AppErr***, ***BoP/EoP***, ***ErrMsg***, and ***ErrSrc***.

[^1]: The term *hosted* I use for all my Common Components to indicate the Workbook which is "the home" of the component, which means where the development and maintenance mainly happens and where the means for testing including regression testing are located. 
[^2]: S=Sub, F=Function, P=Property (r=read/Get, w=write/Let)

[1]: https://github.com/warbe-maker/VBA-Message
[2]: https://github.com/warbe-maker/VBA-Error
[3]: https://github.com/warbe-maker/VBA-Trace
[4]: https://github.com/warbe-maker/VBA-Basics
[5]: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#the-arrydims-service
[6]: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#the-arry-service
[7]: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#error-handling-application-errors
[8]: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#application-error-apperr
[9]: https://github.com/warbe-maker/VBA-Basics/blob/master/CompMan/source/mBasic.bas
[10]: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#specifics-and-usage-examples-for-mbasic-vba-services