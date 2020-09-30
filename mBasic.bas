Attribute VB_Name = "mBasic"
Option Private Module
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mTest: Declarations, procedures, methods and function
'       likely to be required in any VB-Project.
'
' Note: 1. Procedures of the mBasic module do not use the Common VBA Error Handler.
'          However, this test module uses the mErrHndlr module for test purpose.
'
'       2. This module is developed, tested, and maintained in the dedicated
'          Common Component Workbook Basic.xlsm available on Github
'          https://Github.com/warbe-maker/VBA-Basic-Procedures
'
' Methods:
' - AppErr              Converts a positive error number into a negative one which
'                       ensures non conflicting application error numbers since
'                       they are not mixed up with positive VB error numbers. In
'                       return a negative error number is turned back into its
'                       original positive Application Error Number.
' - AppIsInstalled      Returns TRUE when a named exec is found in the system path
' - ArrayCompare        Compares two one-dimensional arrays. Returns an array with
'                       al different items
' - ArrayIsAllocated    Returns TRUE when the provided array has at least one item
' - ArrayNoOfDims       Returns the number of dimensions of an array.
' - ArrayRemoveItem     Removes an array's item by its index or element number
' - ArrayToRange        Transferres the content of a one- or two-dimensional array
'                       to a range
' - ArrayTrim           Removes any leading or trailing empty items.
' - CleanTrim           Clears a string from any unprinable characters.
' - ErrMsg              Displays a common error message by means of the VB MsgBox.
'
' Requires Reference to:
' - "Microsoft Scripting Runtime"
' - "Microsoft Visual Basic Application Extensibility .."
'
' W. Rauschenberger, Berlin Sept 2020
' ----------------------------------------------------------------------------
' Basic declarations potentially uesefull in any project
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

'Functions to get DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88               ' Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72  ' A point is defined as 1/72 inches
Private Declare PtrSafe Function GetForegroundWindow _
  Lib "User32.dll" () As Long

Private Declare PtrSafe Function GetWindowLongPtr _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hwnd As LongPtr, _
     ByVal nIndex As Long) _
  As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hwnd As LongPtr, _
     ByVal nIndex As LongPtr, _
     ByVal dwNewLong As LongPtr) _
  As LongPtr

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16
                
Public Const DCONCAT    As String = "||"    ' For concatenating and error with a general message (info) to the error description
Public Const DGT        As String = ">"
Public Const DLT        As String = "<"
Public Const DAPOST     As String = "'"
Public Const DKOMMA     As String = ","
Public Const DBSLASH    As String = "\"
Public Const DDOT       As String = "."
Public Const DCOLON     As String = ":"
Public Const DEQUAL     As String = "="
Public Const DSPACE     As String = " "
Public Const DEXCL      As String = "!"
Public Const DQUOTE     As String = """" ' one " character
Private vMsgReply       As Variant

' Common xl constants grouped ----------------------------
Public Enum YesNo   ' ------------------------------------
    xlYes = 1       ' System constants (identical values)
    xlNo = 2        ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------
Public Enum xlOnOff ' ------------------------------------
    xlOn = 1        ' System constants (identical values)
    xlOff = -4146   ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------

Public Property Get MsgReply() As Variant:          MsgReply = vMsgReply:   End Property

Public Property Let MsgReply(ByVal v As Variant):   vMsgReply = v:          End Property

Public Function AppErr(ByVal lNo As Long) As Long
' -------------------------------------------------------------------------------
' Attention: This function is dedicated for being used with Err.Raise AppErr()
'            in conjunction with the common error handling module mErrHndlr when
'            the call stack is supported. The error number passed on to the entry
'            procedure is interpreted when the error message is displayed.
' The function ensures that a programmed (application) error numbers never
' conflicts with VB error numbers by adding vbObjectError which turns it into a
' negative value. In return, translates a negative error number back into an
' Application error number. The latter is the reason why this function must never
' be used with a true VB error number.
' -------------------------------------------------------------------------------
    If lNo < 0 Then
        AppErr = lNo - vbObjectError
    Else
        AppErr = vbObjectError + lNo
    End If
End Function

Public Function AppIsInstalled(ByVal sApp As String) As Boolean
Dim i As Long: i = 1
    Do Until Left(Environ$(i), 5) = "Path="
        i = i + 1
    Loop
    AppIsInstalled = InStr(Environ$(i), sApp) <> 0
End Function

Public Function ArrayCompare(ByVal a1 As Variant, _
                             ByVal a2 As Variant, _
                    Optional ByVal lStopAfter As Long = 1) As Variant
' -------------------------------------------------------------------
' Returns an array of all lines which are different. Each element
' contains the corresponding elements of both arrays in form:
' linenumber: <line>||<line>. When a value for stop after
' (lStopAfter) is provided greater 0 the comparison stops after that.
' Note: Either or both arrays may not be assigned (=empty).
' -------------------------------------------------------------------
Dim l       As Long
Dim i       As Long
Dim va()    As Variant

    If Not mBasic.ArrayIsAllocated(a1) And mBasic.ArrayIsAllocated(a2) Then
        va = a2
    ElseIf mBasic.ArrayIsAllocated(a1) And Not mBasic.ArrayIsAllocated(a2) Then
        va = a1
    ElseIf Not mBasic.ArrayIsAllocated(a1) And Not mBasic.ArrayIsAllocated(a2) Then
        GoTo exit_proc
    End If
    
    l = 0
    For i = LBound(a1) To Min(UBound(a1), UBound(a2))
        If a1(i) <> a2(i) Then
            ReDim Preserve va(l)
            va(l) = i & ": " & DGT & a1(i) & DLT & DCONCAT & DGT & a2(i) & DLT
            l = l + 1
            If lStopAfter > 0 And l >= lStopAfter Then GoTo exit_proc
        End If
    Next i
    
    If UBound(a1) < UBound(a2) Then
        For i = UBound(a1) + 1 To UBound(a2)
            ReDim Preserve va(l)
            va(l) = i & ": " & DGT & DLT & DCONCAT & DGT & a2(i) & DLT
            l = l + 1
            If lStopAfter > 0 And l >= lStopAfter Then GoTo exit_proc
        Next i
        
    ElseIf UBound(a2) < UBound(a1) Then
        For i = UBound(a2) + 1 To UBound(a1)
            ReDim Preserve va(l)
            va(l) = i & ": " & DGT & a1(i) & DLT & DCONCAT & DGT & DLT
            l = l + 1
            If lStopAfter > 0 And l >= lStopAfter Then GoTo exit_proc
        Next i
    End If

exit_proc:
    ArrayCompare = va
    
End Function

Public Function ArrayDiffers(ByVal a1 As Variant, _
                             ByVal a2 As Variant) As Boolean
' ----------------------------------------------------------
' Returns TRUE when array (a1) differs from array (a2).
' ----------------------------------------------------------
Const PROC  As String = "ArrayDiffers"
Dim i       As Long
Dim va()    As Variant

    On Error GoTo on_error
    
    If Not mBasic.ArrayIsAllocated(a1) And mBasic.ArrayIsAllocated(a2) Then
        va = a2
    ElseIf mBasic.ArrayIsAllocated(a1) And Not mBasic.ArrayIsAllocated(a2) Then
        va = a1
    ElseIf Not mBasic.ArrayIsAllocated(a1) And Not mBasic.ArrayIsAllocated(a2) Then
        GoTo exit_proc
    End If
    
    On Error Resume Next
    ArrayDiffers = Join(a1) <> Join(a2)
    If Err.Number = 0 Then GoTo exit_proc
    
    '~~ At least one of the joins resulted in a string exeeding the maximum possible lenght
    For i = LBound(a1) To Min(UBound(a1), UBound(a2))
        If a1(i) <> a2(i) Then
            ArrayDiffers = True
            Exit Function
        End If
    Next i
    
exit_proc:
    Exit Function

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Function

Public Function ArrayIsAllocated(arr As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = IsArray(arr) And _
                       Not IsError(LBound(arr, 1)) And _
                       LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function ArrayNoOfDims(arr As Variant) As Integer
' ------------------------------------------------------
' Returns the number of dimensions of an array. An un-
' allocated dynamic array has 0 dimensions. This may as
' as well be tested by means of ArrayIsAllocated.
' ------------------------------------------------------
Dim Ndx As Integer
Dim Res As Integer

    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    ArrayNoOfDims = Ndx - 1

End Function

Public Sub ArrayRemoveItems(ByRef va As Variant, _
                   Optional ByVal Element As Variant, _
                   Optional ByVal Index As Variant, _
                   Optional ByVal NoOfElements = 1)
' ------------------------------------------------------
' Returns the array (va) with the number of elements
' (NoOfElements) removed whereby the start element may be
' indicated by the element number 1,2,... (vElement) or
' the index (Index) which must be within the array's
' LBound to Ubound.
' Any inapropriate provision of the parameters results
' in a clear error message.
' When the last item in an array is removed the returned
' arry is erased (no longer allocated).
'
' Restriction: Works only with one dimensional array.
'
' W. Rauschenberger, Berlin Jan 2020
' ------------------------------------------------------
Const PROC              As String = "ArrayRemoveItems"
Dim a                   As Variant
Dim iElement            As Long
Dim iIndex              As Long
Dim NoOfElementsInArray    As Long
Dim i                   As Long
Dim iNewUBound          As Long

    On Error GoTo on_error
    
    If Not IsArray(va) Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
    Else
        a = va
        NoOfElementsInArray = UBound(a) - LBound(a) + 1
    End If
    If Not ArrayNoOfDims(a) = 1 Then
        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
    End If
    If Not IsNumeric(Element) And Not IsNumeric(Index) Then
        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
    End If
    If IsNumeric(Element) Then
        iElement = Element
        If iElement < 1 _
        Or iElement > NoOfElementsInArray Then
            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & NoOfElementsInArray & " !"
        Else
            iIndex = LBound(a) + iElement - 1
        End If
    End If
    If IsNumeric(Index) Then
        iIndex = Index
        If iIndex < LBound(a) _
        Or iIndex > UBound(a) Then
            Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
        Else
            iElement = ElementOfIndex(a, iIndex)
        End If
    End If
    If iElement + NoOfElements - 1 > NoOfElementsInArray Then
        Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & NoOfElements & ") is beyond the number of elelemnts in the array (" & NoOfElementsInArray & ")!"
    End If
    
    For i = iIndex + NoOfElements To UBound(a)
        a(i - NoOfElements) = a(i)
    Next i
    
    iNewUBound = UBound(a) - NoOfElements
    If iNewUBound < 0 Then Erase a Else ReDim Preserve a(LBound(a) To iNewUBound)
    va = a
    
exit_proc:
    Exit Sub

on_error:
    '~~ Global error handling is used to seamlessly monitor error conditions
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Sub ArrayToRange(ByVal vArr As Variant, _
                        ByVal r As Range, _
               Optional ByVal bOneCol As Boolean = False)
' -------------------------------------------------------
' Copy the content of the Arry (vArr) to the range (r).
' -------------------------------------------------------
Dim rTarget As Range

    If bOneCol Then
        '~~ One column, n rows
        Set rTarget = r.Cells(1, 1).Resize(UBound(vArr), 1)
        rTarget.Value = Application.Transpose(vArr)
    Else
        '~~ One column, n rows
        Set rTarget = r.Cells(1, 1).Resize(1, UBound(vArr))
        rTarget.Value = vArr
    End If
    
End Sub

Public Sub ArrayTrimm(ByRef a As Variant)
' ---------------------------------------
' Return the array (a) with all leading
' and trailing blank items removed. Any
' vbCr, vbCrLf, vbLf are ignored.
' When the array contains only blank
' items the returned array is erased.
' ---------------------------------------
Const PROC  As String = "ArrayTrimm"
Dim i       As Long

    On Error GoTo on_error
    
    '~~ Eliminate leading blank lines
    If Not mBasic.ArrayIsAllocated(a) Then Exit Sub
    
    Do While (Len(Trim$(a(LBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
        mBasic.ArrayRemoveItems a, Index:=i
        If Not mBasic.ArrayIsAllocated(a) Then Exit Do
    Loop
    
    If mBasic.ArrayIsAllocated(a) Then
        Do While (Len(Trim$(a(UBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
            If UBound(a) = 0 Then
                Erase a
            Else
                ReDim Preserve a(UBound(a) - 1)
            End If
            If Not mBasic.ArrayIsAllocated(a) Then Exit Do
        Loop
    End If
exit_proc:
    Exit Sub
    
on_error:
    '~~ Global error handling is used to seamlessly monitor error conditions
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Function BaseName(ByVal v As Variant) As String
' -----------------------------------------------------
' Returns the file name without the extension. v may be
' a file name a file path (full name) a File object or
' a Workbook object.
' -----------------------------------------------------
Const PROC  As String = "BaseName"
Dim fso     As New FileSystemObject

    On Error GoTo on_error
    
    Select Case TypeName(v)
        Case "String":      BaseName = fso.GetBaseName(v)
        Case "Workbook":    BaseName = fso.GetBaseName(v.FullName)
        Case "File":        BaseName = fso.GetBaseName(v.ShortName)
        Case Else:          Err.Raise AppErr(1), ErrSrc(PROC), "The parameter (v) is neither a string nor a File or Workbook object (TypeName = '" & TypeName(v) & "')!"
    End Select

exit_proc:
    Exit Function
    
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Function

Public Function CleanTrim(ByVal s As String, _
                 Optional ByVal ConvertNonBreakingSpace As Boolean = True) As String
' ----------------------------------------------------------------------------------
' Returns the string 's' cleaned from any non-printable characters.
' ----------------------------------------------------------------------------------
Dim l           As Long
Dim asToClean   As Variant
    
    asToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                     21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then s = Replace(s, Chr$(160), " ")
    For l = LBound(asToClean) To UBound(asToClean)
        If InStr(s, Chr$(asToClean(l))) Then s = Replace(s, Chr$(asToClean(l)), vbNullString)
    Next
    CleanTrim = s

End Function

Public Sub ErrMsg(ByVal errnumber As Long, _
                  ByVal errsource As String, _
                  ByVal errdscrptn As String, _
                  ByVal errline As String)
' ----------------------------------------------------
' Display the error message by means of the VBA MsgBox
' ----------------------------------------------------
    
    Dim sErrMsg     As String
    Dim sErrPath    As String
    
    sErrMsg = "Description: " & vbLf & ErrMsgErrDscrptn(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrMsgErrLine(errline)
    sErrPath = vbNullString ' only available with the mErrHndlr module
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & sErrPath
    If ErrMsgInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrMsgInfo(errdscrptn)
    MsgBox sErrMsg, vbCritical, ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline)

End Sub

Private Function ErrMsgErrType( _
        ByVal errnumber As Long, _
        ByVal errsource As String) As String
' ------------------------------------------
' Return the kind of error considering the
' Err.Source and the error number.
' ------------------------------------------

   If InStr(1, Err.Source, "DAO") <> 0 _
   Or InStr(1, Err.Source, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, Err.Source, "ODBC") <> 0 _
   Or InStr(1, Err.Source, "Oracle") <> 0 Then
      ErrMsgErrType = "Database Error"
   Else
      If errnumber > 0 _
      Then ErrMsgErrType = "VB Runtime Error" _
      Else ErrMsgErrType = "Application Error"
   End If
   
End Function

Private Function ErrMsgErrLine(ByVal errline As Long) As String
    If errline <> 0 _
    Then ErrMsgErrLine = " (at line " & errline & ")" _
    Else ErrMsgErrLine = vbNullString
End Function

Private Function ErrMsgErrDscrptn(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string before a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgErrDscrptn = Split(s, DCONCAT)(0) _
    Else ErrMsgErrDscrptn = s
End Function

Private Function ErrMsgInfo(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string after a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgInfo = Split(s, DCONCAT)(1) _
    Else ErrMsgInfo = vbNullString
End Function

Public Function ElementOfIndex(ByVal a As Variant, _
                                ByVal i As Long) As Long
' ------------------------------------------------------
' Returns the element number of index (i) in array (a).
' ------------------------------------------------------
Dim ia  As Long
    For ia = LBound(a) To i
        ElementOfIndex = ElementOfIndex + 1
    Next ia
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & " mBasic." & sProc
End Function

Public Function IsCvName(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then IsCvName = True
End Function

Public Function IsCvObject(ByVal v As Variant) As Boolean

    If VarType(v) = vbObject Then
        If Not TypeName(v) = "Nothing" Then
            IsCvObject = TypeOf v Is CustomView
        End If
    End If
    
End Function

Public Function IsPath(ByVal v As Variant) As Boolean
    
    If VarType(v) = vbString Then
        If InStr(v, "\") <> 0 Then
            If InStr(Right(v, 6), ".") = 0 Then
                IsPath = True
            End If
        End If
    End If

End Function

Public Sub MakeFormResizable()
' ---------------------------------------------------------------------------
' This part is from Leith Ross                                              |
' Found this Code on:                                                       |
' https://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html |
'                                                                           |
' All credits belong to him                                                 |
' ---------------------------------------------------------------------------
Const WS_THICKFRAME = &H40000
Const GWL_STYLE As Long = (-16)
Dim lStyle As LongPtr
Dim hwnd As LongPtr
Dim RetVal

    hwnd = GetForegroundWindow
    
    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE Or WS_THICKFRAME)
    RetVal = SetWindowLongPtr(hwnd, GWL_STYLE, lStyle)
End Sub

Public Function Max(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = 0, _
           Optional ByVal v4 As Variant = 0, _
           Optional ByVal v5 As Variant = 0, _
           Optional ByVal v6 As Variant = 0, _
           Optional ByVal v7 As Variant = 0, _
           Optional ByVal v8 As Variant = 0, _
           Optional ByVal v9 As Variant = 0) As Variant
' -----------------------------------------------------
' Returns the maximum (biggest) of all provided values.
' -----------------------------------------------------
Dim dMax As Double
    dMax = v1
    If v2 > dMax Then dMax = v2
    If v3 > dMax Then dMax = v3
    If v4 > dMax Then dMax = v4
    If v5 > dMax Then dMax = v5
    If v6 > dMax Then dMax = v6
    If v7 > dMax Then dMax = v7
    If v8 > dMax Then dMax = v8
    If v9 > dMax Then dMax = v9
    Max = dMax
End Function

Public Function Min(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = Nothing, _
           Optional ByVal v4 As Variant = Nothing, _
           Optional ByVal v5 As Variant = Nothing, _
           Optional ByVal v6 As Variant = Nothing, _
           Optional ByVal v7 As Variant = Nothing, _
           Optional ByVal v8 As Variant = Nothing, _
           Optional ByVal v9 As Variant = Nothing) As Variant
' ------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------
Dim dMin As Double
    dMin = v1
    If v2 < dMin Then dMin = v2
    If TypeName(v3) <> "Nothing" Then If v3 < dMin Then dMin = v3
    If TypeName(v4) <> "Nothing" Then If v4 < dMin Then dMin = v4
    If TypeName(v5) <> "Nothing" Then If v5 < dMin Then dMin = v5
    If TypeName(v6) <> "Nothing" Then If v6 < dMin Then dMin = v6
    If TypeName(v7) <> "Nothing" Then If v7 < dMin Then dMin = v7
    If TypeName(v8) <> "Nothing" Then If v8 < dMin Then dMin = v8
    If TypeName(v9) <> "Nothing" Then If v9 < dMin Then dMin = v9
    Min = dMin
End Function

Public Function Msg(ByVal sTitle As String, _
           Optional ByVal sMsgText As String = vbNullString, _
           Optional ByVal bFixed As Boolean = False, _
           Optional ByVal sTitleFontName As String = vbNullString, _
           Optional ByVal lTitleFontSize As Long = 0, _
           Optional ByVal siFormWidth As Single = 0, _
           Optional ByVal sReplies As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------
' Custom message using the UserForm fMsg. The function returns the
' clicked reply button's caption or the corresponding vb variable
' (vbOk, vbYes, vbNo, etc.) or its caption string.
' ------------------------------------------------------------------
    Dim w, h        As Long
    Dim siHeight    As Single

    w = GetSystemMetrics32(0) ' Screen Resolution width in points
    h = GetSystemMetrics32(1) ' Screen Resolution height in points
    
    With fMsg
        .Title = sTitle
        .TitleFontName = sTitleFontName
        .TitleFontSize = lTitleFontSize
        
        If sMsgText <> vbNullString Then
            If bFixed = True _
            Then .Message1Fixed = sMsgText _
            Else .Message1Proportional = sMsgText
        End If
        
        .Replies = sReplies
        If siFormWidth <> 0 Then .Width = Max(.Width, siFormWidth)
        .StartupPosition = 1
        .Width = w * PointsPerPixel * 0.85 'Userform width= Width in Resolution * DPI * 85%
        siHeight = h * PointsPerPixel * 0.2
        .Height = Min(.Height, siHeight)
                     
        .Show
    End With
    
    Msg = vMsgReply
End Function

Public Function Msg3(ByVal sTitle As String, _
            Optional ByVal sLabel1 As String = vbNullString, _
            Optional ByVal sText1 As String = vbNullString, _
            Optional ByVal bFixed1 As Boolean = False, _
            Optional ByVal sLabel2 As String = vbNullString, _
            Optional ByVal sText2 As String = vbNullString, _
            Optional ByVal bFixed2 As Boolean = False, _
            Optional ByVal sLabel3 As String = vbNullString, _
            Optional ByVal sText3 As String = vbNullString, _
            Optional ByVal bFixed3 As Boolean = False, _
            Optional ByVal sTitleFontName As String = vbNullString, _
            Optional ByVal lTitleFontSize As Long = 0, _
            Optional ByVal siFormWidth As Single = 0, _
            Optional ByVal sReplies As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------
' Custom message allowing three sections, each with a label/haeder,
' using the UserForm fMsg. The function returns the clicked reply
' button's caption or the corresponding vb variable (vbOk, vbYes,
' vbNo, etc.) or its caption string.
' ------------------------------------------------------------------
    Dim w, h        As Long
    Dim siHeight    As Single

    w = GetSystemMetrics32(0) ' Screen Resolution width in points
    h = GetSystemMetrics32(1) ' Screen Resolution height in points
    
    With fMsg
        .Title = sTitle
        .TitleFontName = sTitleFontName
        .TitleFontSize = lTitleFontSize
        
        If sText1 <> vbNullString Then
            If bFixed1 = True _
            Then .Message1Fixed = sText1 _
            Else .Message1Proportional = sText1
            .LabelMessage1 = sLabel1
        End If
        
        If sText2 <> vbNullString Then
            If bFixed2 = True _
            Then .Message2Fixed = sText2 _
            Else .Message2Proportional = sText2
            .LabelMessage2 = sLabel2
        End If
        
        If sText3 <> vbNullString Then
            If bFixed3 = True _
            Then .Message3Fixed = sText3 _
            Else .Message3Proportional = sText3
            .LabelMessage3 = sLabel3
        End If
        
        .Replies = sReplies
        If siFormWidth <> 0 Then .Width = Max(.Width, siFormWidth)
        .StartupPosition = 1
        .Width = w * PointsPerPixel * 0.85 'Userform width= Width in Resolution * DPI * 85%
        siHeight = h * PointsPerPixel * 0.2
        .Height = Min(.Height, siHeight)
                     
        .Show
    End With
    
    Msg3 = vMsgReply

End Function

Public Function PointsPerPixel() As Double
' ----------------------------------------
' Return DPI
' ----------------------------------------
Dim hDC             As Long
Dim lDotsPerInch    As Long
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC
End Function

Public Function ProgramIsInstalled(ByVal sProgram As String) As Boolean
        ProgramIsInstalled = InStr(Environ$(18), sProgram) <> 0
End Function

Public Function SelectFolder( _
                Optional ByVal sTitle As String = "Select a Folder") As String
' ----------------------------------------------------------------------------
' Returns the selected folder or a vbNullString if none had been selected.
' ----------------------------------------------------------------------------
Dim sFolder As String
    SelectFolder = vbNullString
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = sTitle
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    SelectFolder = sFolder
End Function

Public Function Space(ByVal l As Long) As String
' --------------------------------------------------
' Unifies the VB differences SPACE$ and Space$ which
' lead to code diferences where there aren't any.
' --------------------------------------------------
    Space = VBA.Space$(l)
End Function

