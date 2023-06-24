# Common VBA Basic Services
This is a (personal) collection of services used every now and then n VB-Projects. Some of them [^1] serve my personal common Excel VB-Project development needs but many of them had been collected over a long time.

# Services
| Service           |Kind&nbsp;[^2]| Description |
|-------------------|:------------:|-------------|
|_AppErr_           | S            | Ensures that the number of a programmed _Application-Error_ never conflicts with a _VB-Runtime-Error_ by adding the constant _vbObjectError_ which turns it into a negative value. In return, it translates the negative _Application-Error_ number back into its original positive number. When used in an error handling in conjunction with the unique identification of the error source (service ErrSrc in the mErH module) programmed error numbers can start from 1 to n in any procedure.|
|_AppInstalled_     | F           | Returns TRUE when a named program/application, usually an .exe file is installed.|
|_ArrayCompare_     | F           | Returns a Dictionary with the provided number of lines/items (defaults to all) which differ between two one dimensional arrays. When no difference is encountered the returned Dictionary is empty (Count = 0).|
|_ArrayDiffers_     | F           | Returns TRUE when two arrays are different (stops with the first difference).|
|_ArrayIsAlocated_  | F           | Returns TRUE when the provided array has at least one item.|
|_ArrayNumberOfDims_| F           | Returns the number of dimensions of a provided array.|
|_ArrayRemoveItems_ | S           | Returns a provided array with the identified item removed. The to-be-removed items may be provide as index plus the number of to be removed items.|
|_ArrayToRange_     | S           | Copies the content of a one or two dimensional array to a provided range.|
|_ArrayTrim_        | S           | Returns a provided array with all leading and trailing blank items removed. Any `vbCr`, `vbCrLf`, `vbLf` are ignored. When the array contains only blank items the returned array is erased.|
|_BaseName_         | F           | Returns the file name's BaseName. The argument may be a full file name, a file object or a  Workbook object.|
|_BoP/EoP_          | S           | Common **B**egin-**o**f-**P**rocedure/**E**nd-**o**f-**P**rocedure  interface for the optional [Common VBA Execution Trace Service][3]. Obligatory copy Private for any VB-Component using the _[Common VBA Execution Trace Service][3] but not having the mBasic common component installed. 1), 2).|
|_BoC/EoC_          | S           | Analog to the above for to-be-traced code sequences  1), 2) |
|_Center_           | F           | Returns a string centered within a string with of a given length.|
|_CleanTrim_        | F           | Returns a provide string cleaned from any non-printable characters.|
|_ElementOfIndex_   | F           | Returns the element number of index (i) in array (a).|
|_ErrMsg_           | F           | Universal error message display service. Obligatory Private copy for any VB-Component using the common error service but not having this component installed. The function displays:<br>- a debugging option button when the Cond. Comp. Arg. `Debugging = 1`<br>- an optional additional "About:" section when the provided error description has an additional string concatenated by two vertical bars (\|\|)<br>- the error message by means of the [Common VBA Message Service (fMsg/mMsg)][1] when installed, and active (Cond. Comp. Arg. `MsgComp = 1`) else by means of `VBA.MsgBox`. The function uses _AppErr_ to identify/distinguish programmed application errors (`Err.Raise AppErr(n), ....`) and turn them back into their positive number. The function is obligatory as `Private` copy for any VB-Component using the _Common VBA Execution Trace Service_ and/or the _[Common VBA Error Services][2]_ but not having the mBasic common component installed. 1), 2).|
|_Max_              | F          | Returns the maximum value of any number of values provided.|
|_Min_              | F          | Returns the minimum value of any number of values provided.|
|_KeySort_          | F          | Returns a given Dictionary sorted by key. |
|_PointsPerPixel_   | F          | Return the DPI of the current display/monitor.|
|_README_           | S | Displays a given url with a given bookmark in the computer's default browser. The url defaults to this 'component's README url in the public GitHub repo, the bookmark defaults to vbNullString.|
|_SelectFolder_     | ||
|_ShellRun_         | S          | Opens a folder, an email-app, a url, an Access instance, etc. all by means of the default application. Courtesy of: Dev Ashish.<br>Example:<br>`ShellRun "https://github.com/warbe-maker/VBA-Basics"`<br> displays this README in this _Common Component's_ public GiHub repo.|
|_Spaced_           | | Returns a provided string 'spaced' with non-braking spaces. Spaces already in the provided string are doubled.|
|_SysFrequency_     | P r        | Returns the current number of ticks per seconds which is the precision for the _TimeBegin_ end _TimerEnd_ function.|
|_TimedDoEvents_    | S          | Debug.Print the time delay of a performed `DoEvents`.<br>**Background:** `DoEvents` are often recommended to solve un-explainable problems. However, according to the `DoEvents` documentation the way how and why it may solve a problem is pretty miraculous. Anyhow, even when it helped it should be taken into consideration that `DoEvents`  enable keyboard interactions while a process executes. In case of an endless loop with embedded `DoEvents` this may be a godsend. But it as well may cause unpredictable timing  results. When, instead of performing `DoEvents` directly `TimedDoEvents` is used at least the resulting performance delay in milliseconds is printed to the _VBE Immediate Window_.|
|_TimerBegin_       | F          | Returns the current system ticks as the start of a timer. |
|_TimerEnd_         | F          | Returns, based on a provided _TimerBegin_:<br>'- the end-ticks<br>- the elapsed ticks<br>- the elapsed time in the provided format which defaults to `"hh:mm:ss.0000"`.)|


[^1]: It goes without saying that my VB-Projects use this _mBasic_ component. However, all my _Common Components_ use some services as `Private Sub` copy. This keeps them 100% autonomous, i.e. independent from this and other components but still serve my personal use of them. The service i am talking about are:  
-&nbsp;_BoP/EoP_, and _ErrMsg_ to keep them independent from the _[Common VBA Error Services][2]_  
-&nbsp;_BoP/EoP_, _BoC/EoC_ to keep the use of the _[Common VBA Execution Trace Service][3]_ optional

[^2]: S=Sub, F=Function, P=Property (r=read/Get, w=write/Let)

[1]: https://github.com/warbe-maker/VBA-Message
[2]: https://github.com/warbe-maker/VBA-Error
[3]: https://github.com/warbe-maker/VBA-Trace
[4]: https://github.com/warbe-maker/VBA-Basics