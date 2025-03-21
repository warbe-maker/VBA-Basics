Attribute VB_Name = "mBasicTest"
Option Private Module
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mTest: Dedicate for the test of mBasic procedures.
'
' To be noticed: Procedures of the mBasic module do not use the
'                Common VBA Error Handler. However, since testing includes
'                testing of error conditions the Conditional Compile Argument
'                'CommonErHComp = 1' is essential in fact the following are
'                obligatory for this Development and Test Workbook, VB-Project
'                respectively:
'                Debugging = 1
'                ExecTrace = 1
'                CompMan = 1
'                ErHComp = 1
'
' W. Rauschenberger, Berlin Now 2017
' ----------------------------------------------------------------------------
Private bModeRegression As Boolean
Public Trc As clsTrc

Private TestAid As New clsTestAid

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mBasicTest." & s:  End Property

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | ErHComp = 1             | mErH                |
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If Not IsMissing(b_arguments) Then s = Join(b_arguments, ";")

#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | ErHComp = 1             | mErH                |
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays a debugging option
' button when the Conditional Compile Argument 'Debugging = 1', displays an
' optional additional "About:" section when the err_dscrptn has an additional
' string concatenated by two vertical bars (||), and displays the error message
' by means of VBA.MsgBox when neither the Common Component mErH (indicated by
' the Conditional Compile Argument "ErHComp = 1", nor the Common Component mMsg
' (idicated by the Conditional Compile Argument "MsgComp = 1") is installed.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'       ErrSrc  To provide an unambiguous procedure name by prefixing is with
'               the module name.
'
' W. Rauschenberger Berlin, Apr 2016
'
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
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
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Sub Prepare()
    
    mErH.Regression = bModeRegression

    If TestAid Is Nothing Then Set TestAid = New clsTestAid
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    If bModeRegression Then
'        TestAid.ModeRegression = True
        Trc.FileFullName = TestAid.TestFolder & "\TestExecution.trc"
    Else
        Trc.FileFullName = TestAid.TestFolder & "\RegressionTestExecution.trc"
    End If
    TestAid.TestedComp = "mBasic"
    
End Sub

Private Function TestArray(ParamArray t_specs() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns an array t_specs(0) with items like "Item(x,y,z)" whereby the integers
' in the brackets correspond with the Item's indices in a n-dim (up to 8).
' When a dim-ed array (t_arry) is provided its dimension specs are mandatory,
' else the provided specs (t_specs(1), t_specs(2), ... are used to provide the
' array dimension specifics.
'
' Uses: ArryDims
' ------------------------------------------------------------------------------
    
    Const PROC = "TestArray"
    
    Dim arr                 As Variant
    Dim arrSpecs(1 To 2)    As Long
    Dim cllSpecs            As New Collection
    Dim i                   As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long
    Dim lBase               As Long
    Dim lDims               As Long
    Dim v                   As Variant
    
    On Error GoTo xt ' no argument
    If UBound(t_specs) >= LBound(t_specs) Then
        lBase = LBound(Array("x"))
        If IsArray(t_specs(0)) Then
            arr = t_specs(0)
            lDims = ArryDims(arr)
        Else
            lDims = UBound(t_specs) - LBound(t_specs) + 1
            For Each v In t_specs
                If Not mBasic.IsInteger(v) Then Err.Raise AppErr(1), ErrSrc(PROC), "Dimesion spec not an integer!"
                arrSpecs(1) = lBase
                arrSpecs(2) = v
                cllSpecs.Add arrSpecs
            Next v
        End If
        
        '~~ Assert the number of dimensions less/equal 8
        If lDims > 8 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "Max number of possible t_specs of 8 exceeded!"
        
        If Not IsArray(t_specs(0)) Then
            '~~ When no array has been provided, one is created
            Select Case lDims
                Case 0: ReDim arr(lBase To t_specs(0))
                Case 1: ReDim arr(lBase To cllSpecs(1)(2))
                Case 2: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2))
                Case 3: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2))
                Case 4: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2) _
                                , lBase To cllSpecs(4)(2))
                Case 5: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2) _
                                , lBase To cllSpecs(4)(2) _
                                , lBase To cllSpecs(5)(2))
                Case 6: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2) _
                                , lBase To cllSpecs(4)(2) _
                                , lBase To cllSpecs(5)(2) _
                                , lBase To cllSpecs(6)(2))
                Case 7: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2) _
                                , lBase To cllSpecs(4)(2) _
                                , lBase To cllSpecs(5)(2) _
                                , lBase To cllSpecs(6)(2) _
                                , lBase To cllSpecs(7)(2))
                Case 8: ReDim arr(lBase To cllSpecs(1)(2) _
                                , lBase To cllSpecs(2)(2) _
                                , lBase To cllSpecs(3)(2) _
                                , lBase To cllSpecs(4)(2) _
                                , lBase To cllSpecs(5)(2) _
                                , lBase To cllSpecs(6)(2) _
                                , lBase To cllSpecs(7)(2) _
                                , lBase To cllSpecs(8)(2))
            End Select
        End If
        
        lDims = ArryDims(arr)
        On Error Resume Next
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                For k = LBound(arr, 3) To UBound(arr, 3)
                    For l = LBound(arr, 4) To UBound(arr, 4)
                        For m = LBound(arr, 5) To UBound(arr, 5)
                            For n = LBound(arr, 6) To UBound(arr, 6)
                                For o = LBound(arr, 7) To UBound(arr, 7)
                                    For p = LBound(arr, 8) To UBound(arr, 8)
                                        Select Case lDims
                                            Case 1: arr(i) = "Item(" & i & ")"
                                            Case 2: arr(i, j) = "Item(" & i & "," & j & ")"
                                            Case 3: arr(i, j, k) = "Item(" & i & "," & j & "," & k & ")"
                                            Case 4: arr(i, j, k, l) = "Item(" & i & "," & j & "," & k & "," & l & ")"
                                            Case 5: arr(i, j, k, l, m) = "Item(" & i & "," & j & "," & k & "," & l & "," & m & ")"
                                            Case 6: arr(i, j, k, l, m, n) = "Item(" & i & "," & j & "," & k & "," & l & "," & m & "," & n & ")"
                                            Case 7: arr(i, j, k, l, m, n, o) = "Item(" & i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o & ")"
                                            Case 8: arr(i, j, k, l, m, n, o, p) = "Item(" & i & "," & j & "," & k & "," & l & "," & m & "," & n & "," & o & "," & p & ")"
                                        End Select
                                    Next p
                                Next o
                            Next n
                        Next m
                    Next l
                Next k
            Next j
        Next i
    
    End If
    
xt: On Error GoTo -1
    TestArray = arr
    
End Function

Public Sub Regression()
' ------------------------------------------------------------------------------
' This Regression test:
' - uses the Common VBA Execution Trace service (mTrc) to trace/log the
'   performed tests
' - uses the Common VBA Message Service (fMsg/mMsg) to provide well designed
'   error messages (requires the Conditional Compile Argument MsgComp = 1)
' - requires the Conditional Compile Arguments:
'   Debugging = 1 : ErHComp = 1 : MsgComp = 1 : XcTrc_mTrc = 1
'   last but not least to run uninterrupted with all errors asserted.
'   When mErH.Regression is not set to True any mErH.Asserted does not have any
'   effect, i.e. all errors are displayed one by one.
' ------------------------------------------------------------------------------
    Const PROC = "Regression"
    
    On Error GoTo eh
    bModeRegression = True
    Prepare
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
#If mTrc = 1 Then
    mTrc.FileName = TestAid.TestFolder & "\RegressionTest.ExecTrace.log"
    mTrc.Title = "Execution Trace Result of the mBasic Regression test"
    mTrc.NewFile
#ElseIf clsTrc = 1 Then
    Set Trc = New clsTrc
    With Trc
        .FileFullName = TestAid.TestFolder & "\RegressionTest.ExecTrace.log"
        .Title = "Execution Trace Result of the mBasic Regression test"
        .NewFile
    End With
#End If
    
    BoP ErrSrc(PROC)
    mBasicTest.Test_0010_Fundamentals
    mBasicTest.Test_0110_Align_Simple
    mBasicTest.Test_0120_Align_column_arranged
    mBasicTest.Test_0200_ArryNextIndex
    mBasicTest.Test_0210_Arry_Get
    mBasicTest.Test_0215_Arry_Let
    mBasicTest.Test_0216_Arry_Let_Performance
    mBasicTest.Test_0220_ArryAsRnge_RngeAsArry
    mBasicTest.Test_0230_ArryCompare
    mBasicTest.Test_0240_ArryDiffers
    mBasicTest.Test_0250_Arry_Various
    mBasicTest.Test_0260_ArrayAsDict
    mBasicTest.Test_0270_ArryRemoveItems
    mBasicTest.Test_0275_ArryReDim
    mBasicTest.Test_0280_ArryTrimm
    mBasicTest.Test_0400_Spaced
'    mBasicTest.Test_0500_Stack
'    mBasicTest.Test_0600_TimedDoEvents
'    mBasicTest.Test_0700_Timer
     mBasicTest.Test_0800_Coll
            
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
#If mTrc = 1 Then
    mTrc.Dsply
#ElseIf clsTrc = 1 Then
    Trc.Dsply
#End If
    TestAid.ResultLogSummary
    TestAid.CleanUp
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0260_ArrayAsDict()
    Const PROC = "Test_0260_ArrayAsDict"

    On Error GoTo eh
    Dim arr     As Variant
    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim l       As Long
    Dim m       As Long
    Dim dctRes  As Dictionary
    Dim arrRes  As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0260"
        .Title = "Unload/reload multidimensional arry to Dictionary"
        
        ReDim arr(1 To 2, 2 To 3, 3 To 4, 4 To 5, 5 To 6)
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                For k = LBound(arr, 3) To UBound(arr, 3)
                    For l = LBound(arr, 4) To UBound(arr, 4)
                        For m = LBound(arr, 5) To UBound(arr, 5)
                            arr(i, j, k, l, m) = "Item(" & i & "," & j & "," & k & "," & l & "," & m & ")"
                        Next m
                    Next l
                Next k
            Next j
        Next i
        
        .Verification = "The number of items in the Dictionary are equal to those in the array"
            .TestedProc = "Function ArryAsDict"
            .TimerStart
            Set dctRes = ArryAsDict(arr)
            .TimerEnd
            .Result = dctRes.Count
            .ResultExpected = CInt(ArryItems(arr))
            
        .Verification = "The number of items in the array is equal to those in the Dictionary"
            .TestedProc = "Function DictAsArray"
            .TimerStart
            arrRes = mBasic.DictAsArry(dctRes)
            .TimerEnd
            .Result = mBasic.ArryItems(arrRes)
            .ResultExpected = mBasic.ArryItems(arr)
            
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0220_ArryAsRnge_RngeAsArry()
    Const PROC = "Test_0220_ArryAsRnge_RangeAsArray"
    
    On Error GoTo eh
    Dim arr   As Variant
    Dim rng   As Range
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0220"
            .Title = "Array as range, range as array"
            
            .Verification = "The 1-dim array is written to the range not transposed"
                .TestedProc = "Sub ArryAsRnge" & vbLf & "Function RngeAsArry"
                arr = TestArray(6)
                arr(3) = Empty
                Set rng = wsBasic.Range("celArryAsRangeTarget1")
                .TimerStart
                mBasic.ArryAsRnge arr, rng, False
                .TimerEnd
                .Result = rng
                .ResultExpected = wsBasic.Range("rngExpected1")
            
            .Verification = "The 1-dim array is written to the range transposed"
                .TestedProc = "Function RngeAsArry"
                arr = TestArray(6)
                arr(3) = Empty
                Set rng = wsBasic.Range("celArryAsRangeTarget2")
                .TimerStart
                mBasic.ArryAsRnge arr, rng, True
                .TimerEnd
                .Result = rng
                .ResultExpected = wsBasic.Range("rngExpected2")
    
            .Verification = "The 2-dim array is written to the range not transposed"
                .TestedProc = "Sub ArryAsRnge"
                arr = TestArray(3, 3)
                arr(0, 0) = Empty
                arr(1, 0) = Empty
                arr(2, 0) = Empty
                arr(3, 0) = Empty
                Set rng = wsBasic.Range("celArryAsRangeTarget3")
                .TimerStart
                mBasic.ArryAsRnge arr, rng, False
                .TimerEnd
                .Result = rng
                .ResultExpected = wsBasic.Range("rngExpected3")

            .Verification = "The 2-dim array is written to the range transposed"
                .TestedProc = "Function RngeAsArry"
                arr = TestArray(3, 3)
                arr(0, 0) = Empty
                arr(1, 0) = Empty
                arr(2, 0) = Empty
                arr(3, 0) = Empty
                Set rng = wsBasic.Range("celArryAsRangeTarget4")
                .TimerStart
                mBasic.ArryAsRnge arr, rng, True
                .TimerEnd
                .Result = rng
                .ResultExpected = wsBasic.Range("rngExpected4")
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0900_Max()

    Dim a As Variant
    
    Arry(a) = "5"
    Arry(a) = "xxxxxxxxxx"
    Arry(a) = 3
    
    Debug.Assert Max(2, 5) = 5
    Debug.Assert Max(3, "xxxxxx", 4) = 6
    Debug.Assert Max(2, a, 7) = 10
    
End Sub

Public Sub Test_0010_Fundamentals()
' ----------------------------------------------------------------------------
' Attention! This test cannot simply be repeated since the first
'            verification's Result is only properly computed when the Class
'            is initialized which is not the case once it had already been
'            initialized!
' ----------------------------------------------------------------------------
    Dim arr         As Variant
    Dim cll         As New Collection
    Dim arrLong()   As Long
    Dim l           As Long
    
    Set TestAid = Nothing
    Prepare
    With TestAid
        .TestId = "0010-1"
        .Title = "Fundamental services"
        
        .Verification = "Timer precision proof"
            .CommentOnVerification = "The difference to 500 is considered an unavoidable precision failure - which should usually be less than 10 milliseconds!"
            .TestedComp = "clsTestAid"
            .TestedProc = "Sub TimerStart" & vbLf & "Function TimerEnd"
            .TimerStart
            .SleepMsecs 500 '~~ Nothing executed
            Debug.Print "Sleep 500 = " & .TimerEnd & " exec time in msecs"
            .Result = True
            .ResultExpected = True

        .Verification = "A TestArray is provided conforming with expectations"
            .TestedComp = "mBasicTest"
            .TestedProc = "Function TestArray"
            Set arr = Nothing
            arr = TestArray(2, 2, 2)
            .Result = arr(2, 2, 2)
            .ResultExpected = "Item(2,2,2)"

        .TestId = "0010-2"
        .Title = "Fundamental services"
        .TestedComp = "mBasic"
        
        .Verification = "Indices provided as string are returned as array 1 to n"
            .TestedProc = "Function ArryNdcs"
            arr = mBasic.ArryNdcs(l, "1,2 , 3")
            .Result = l
            .ResultExpected = 3
        
        .Verification = "Indices provided as string are returned as array 1 to n"
            .Result = arr(1) & "," & arr(2) & "," & arr(3)
            .ResultExpected = "1,2,3"

        .Verification = "Indices provided as string are returned as array 1 to n"
            .TestedProc = "Function ArryNdcs"
            arr = mBasic.ArryNdcs(l, mBasic.ArryNdcs(l, "1,2 , 3"))
            .Result = arr(1) & "," & arr(2) & "," & arr(3)
            .ResultExpected = "1,2,3"
        
        .Verification = "Indices provided as string are returned as array 1 to n"
            .Result = CInt(arr(3))
            .ResultExpected = 3
            
        .Verification = "Indices provided as array are returned as array 1 to n"
            arr = mBasic.ArryNdcs(l, Array(3, 5, 7))
            .Result = CInt(arr(3))
            .ResultExpected = 7
        
        .Verification = "Indices provided as array are returned as array 1 to n"
            .Result = CInt(l)
            .ResultExpected = 3
            
        .Verification = "An aleady redimed TestArray is provided conforming with expectations"
            ReDim arr(3, 3, 3)
            arr = TestArray(arr)
            .Result = arr(1, 2, 3)
            .ResultExpected = "Item(1,2,3)"

        .Verification = "Max value without arguments is 0"
            .TestedProc = "Function Max"
            .Result = Max()
            .ResultExpected = 0

        .Verification = "Max returns the max value of provided arguments"
            .Result = Max(10, 50, 2)
            .ResultExpected = 50

        .Verification = "Max returns the max length when arguments are strings"
            .Result = Max("1x34", "1234y678", "1z3") ' numeric value as string is still considered a value
            .ResultExpected = CLng(8)

        .Verification = "Max returns the max value of provided arguments of which some are provided as array"
            .Result = Max(10, Array(3, 5, 60), 50, 2)
            .ResultExpected = 60
    
        .Verification = "Max returns the max value of provided arguments of which some are provided as array"
            cll.Add 5
            cll.Add 70
            .Result = Max(10, cll, 50, 2)
            .ResultExpected = 70
    
        .Verification = "An active array Item (not IsError and not = default is identified as such"
            .TestedProc = "Function ArryItemIsActive"
            ReDim arrLong(1 To 5)
            arrLong(1) = 1  ' 1 active
                            ' 2 error and thus inactive
            arrLong(3) = 0  ' 3 default and thus inactive
                            ' 4 error and thus inactive
            arrLong(5) = 5  ' 5 active
            .Result = ArryItemIsActive(arrLong, 0, 2)
            .ResultExpected = False
    
        .Verification = "An in-active array Item (= default) is identified as such"
            ReDim arrLong(1 To 5)
            arrLong(1) = 1  ' 1 active
                            ' 2 error and thus inactive
            arrLong(3) = 0  ' 3 default and thus inactive
                            ' 4 error and thus inactive
            arrLong(5) = 5  ' 5 active
            .Result = ArryItemIsActive(arrLong, 0, 3)
            .ResultExpected = False
            
        .Verification = "An in-active array Item (= default) is identified as such"
            .TestedProc = "Function ArryItems"
            ReDim arrLong(1 To 5)
            arrLong(1) = 1  ' 1 active
                            ' 2 error and thus inactive
            arrLong(3) = 0  ' 3 default and thus inactive
                            ' 4 error and thus inactive
            arrLong(5) = 5  ' 5 active
            .Result = ArryItems(arrLong, True)
            .ResultExpected = 2
    
        .Verification = "An active array Item (not IsError and not = default is identified as such"
            .TestedProc = "Function ArryItems"
            ReDim arrLong(1 To 5)
            arrLong(1) = 1  ' 1 active
                            ' 2 error and thus inactive
            arrLong(3) = 0  ' 3 default and thus inactive
                            ' 4 error and thus inactive
            arrLong(5) = 5  ' 5 active
            .Result = ArryItems(arrLong, False)
            .ResultExpected = 5
            
        .Verification = "Mixed multi-dimensional arry (with ""jagged"" array items"
            .TestedProc = "Function ArryItems"
            ReDim arr(1 To 3, 5 To 10)
            arr = TestArray(arr)            ' 3 x 6 = 18 - 1 = 17 items (one item is a "jagged" array!)
            arr(2, 8) = Array(2, 0, , 3)    ' + 2 active items = 19 items active
            arr(2, 6) = Empty               ' - 1 = 18 Items
            .Result = ArryItems(arr, True)
            .ResultExpected = 18
            
        .Verification = "Mixed multi-dimensional arry (with ""jagged"" array items"
            .TestedProc = "Function ArryItems"
            ReDim arr(1 To 3, 5 To 10)      ' 3 x 6 = 18 - 1 = 17 items (one item is a "jagged" array!)
            arr = TestArray(arr)
            arr(2, 8) = Array(2, 0, , 3)    ' + 4 "jagged array items = 21
            .Result = ArryItems(arr, False)
            .ResultExpected = 21
         
         If Not bModeRegression Then .ResultLogSummary
    End With
    
End Sub

Public Sub Test_0110_Align_Simple()
    Const PROC = "Test_0110_Align_Simple"
    
    Prepare
    mBasic.BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0100"
            .TestedProc = "Function Align"
            .Title = "Align strings"
            
            .Verification = "Align simple left filled with "" -"""
                .ResultExpected = "Abcde --"
                .Result = mBasic.Align("Abcde", enAlignLeft, 8, " -")
        
            .Verification = "Align simple right filled with ""- """
                .ResultExpected = "-- Abcde"
                .Result = mBasic.Align("Abcde", enAlignRight, 8, "- ")
        
            .Verification = "Align simple left centered with ""- "" (not exactly centered)"
                .ResultExpected = " Abcde -"
                .Result = mBasic.Align("Abcde", enAlignCentered, 8, " -")
        
            .Verification = "Align simple left centered with ""- "" (exactly centered)"
                .ResultExpected = "- Abcde -"
                .Result = mBasic.Align("Abcde", enAlignCentered, 9, " -")
        
            .Verification = "Align simple simple right filled with ""- """
                .ResultExpected = "- Abcde"
                .Result = mBasic.Align("Abcde", enAlignRight, 7, "- ")
                
            .Verification = "Align simple simple centered filled with ""-"""
                .ResultExpected = "-Abcde-"
                .Result = mBasic.Align("Abcde", enAlignCentered, 7, "-")
        
            .Verification = "Align simple simple left filled with ""-"""
                .ResultExpected = "Abcde--"
                .Result = mBasic.Align("Abcde", enAlignLeft, 7, "-")
        
            .Verification = "Align simple simple right filled with ""-"""
                .ResultExpected = "-Abcde"
                .Result = mBasic.Align("Abcde", enAlignRight, 6, "-")
        
            .Verification = "Align simple centered filled with ""-"" (not exactly centered)"
                .ResultExpected = "Abcde-"
                .Result = mBasic.Align("Abcde", enAlignCentered, 6, "-")
        
            .Verification = "Align simple centered filled with ""-"" (truncated to 4)"
                .ResultExpected = "Abcd"
                .Result = mBasic.Align("Abcde", enAlignLeft, 4, "-")
        
            .Verification = "Align simple right filled with ""-"""
                .ResultExpected = "Abcd"
                .Result = mBasic.Align("Abcde", enAlignRight, 4, "-")
        
            .Verification = "Align simple left filled with ""-"" (truncated to 4)"
                .ResultExpected = "Abcd"
                .Result = mBasic.Align("Abcde", enAlignLeft, 4, "-", " ")
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0120_Align_column_arranged()
  Const PROC = "Test_0120_Align_column_arranged"
  
  Prepare
  mBasic.BoP ErrSrc(PROC)
  
  With TestAid
    
    .TestId = "0120"
        .Title = "Align string as arranged in a column"
        .TestedProc = "Function Align"
    
        .Verification = "Align col arranged left filled with "" -"""
            .ResultExpected = " Abcde ---- "
            .Result = mBasic.Align("Abcde", enAlignLeft, 8, " -", " ", True)
        
        .Verification = "Align col arranged right filled with ""- """
            .ResultExpected = " ---- Abcde "
            .Result = mBasic.Align("Abcde", enAlignRight, 8, "- ", " ", True)
        
        .Verification = "Align col arranged left centered with ""- "" (not exactly centered)"
            .ResultExpected = " -- Abcde --- "
            .Result = mBasic.Align("Abcde", enAlignCentered, 8, " -", " ", True)
        
        .Verification = "Align col arranged left centered with ""- "" (exactly centered)"
            .ResultExpected = " -- Abcde -- "
            .Result = mBasic.Align("Abcde", enAlignCentered, 7, " -", " ", True)
        
        .Verification = "Align col arranged right filled with ""- """
            .ResultExpected = " --- Abcde "
            .Result = mBasic.Align("Abcde", enAlignRight, 7, "- ", " ", True)
            
        .Verification = "Align col arranged centered filled with ""-"""
            .ResultExpected = " --Abcde-- "
            .Result = mBasic.Align("Abcde", enAlignCentered, 7, "-", " ", True)
        
        .Verification = "Align col arranged left filled with ""-"""
            .ResultExpected = " Abcde--- "
            .Result = mBasic.Align("Abcde", enAlignLeft, 7, "-", " ", True)
        
        .Verification = "Align col arranged right filled with ""-"""
            .ResultExpected = " --Abcde "
            .Result = mBasic.Align("Abcde", enAlignRight, 6, "-", " ", True)
        
        .Verification = "Align col arranged centered filled with ""-"" (not exactly centered)"
            .ResultExpected = " -Abcde-- "
            .Result = mBasic.Align("Abcde", enAlignCentered, 6, "-", " ", True)
        
        .Verification = "Align col arranged centered filled with ""-"" (not exactly centered)"
            .ResultExpected = " Abcd- "
            .Result = mBasic.Align("Abcde", enAlignLeft, 4, "-", " ", True)
        
        .Verification = "Align col arranged right filled with ""-"""
            .ResultExpected = " -Abcd "
            .Result = mBasic.Align("Abcde", enAlignRight, 4, "-", " ", True)
        
        .Verification = "Align col arranged centered filled with ""-"" (exactly centered)"
            .ResultExpected = " -Abcd- "
            .Result = mBasic.Align("Abcde", enAlignCentered, 4, "-", " ", True)  ' Centered 4  "-"
  
  End With
  
xt: EoP ErrSrc(PROC)
  Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
    Case vbResume: Stop: Resume
    Case Else:   GoTo xt
  End Select
End Sub

Public Sub Test_0200_ArryNextIndex()
    Const PROC = "Test_0200_ArryNextIndex"

    On Error GoTo eh
    Dim arr             As Variant
    Dim arrNext()       As Long
    Dim arrNextExp()    As Long
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0200"
        .TestedProc = "Function ArryNextIndex"
        .Title = "Get next index of a multi-dim arry"
        
            .Verification = "Next index for a 1-dim array is provided"
            arr = TestArray(8)
            ReDim arrNext(1 To 1)
            arrNext(1) = 6
            .TimerStart
            .Result = ArryNextIndex(arr, arrNext)
            .TimerEnd
            .ResultExpected = True
    
            .Verification = "The next index for a 1-dim array is the next possible without ReDim"
            .Result = arrNext(1)
            .ResultExpected = 7
            
            .Verification = "Next index for a 1-dim array is not provided because the given has reached the upper bound"
            arr = TestArray(8)
            ReDim arrNext(1 To 1)
            arrNext(1) = 8
            .TimerStart
            .Result = ArryNextIndex(arr, arrNext)
            .TimerEnd
            .ResultExpected = False
        
            .Verification = "Given an index of a 3,4 for a 2-dim 0-4,0-4 array is provided as 4,0"
            arr = TestArray(4, 4)
            ReDim arrNext(1 To 2)
            ReDim arrNextExp(1 To 2)
            arrNext(1) = 3
            arrNext(2) = 4
            arrNextExp(1) = 4
            arrNextExp(2) = 0
            .TimerStart
            ArryNextIndex arr, arrNext
            .TimerEnd
            .Result = arrNext
            .ResultExpected = arrNextExp
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0210_Arry_Get()
    Const PROC = "Test_0210_Arry_Get"
    
    On Error GoTo eh
    Dim arr1    As Variant
    Dim arr2    As Variant
    Dim arr3()  As Object
    Dim arr4()  As Boolean
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0210-1"
        .Title = "Array read"
        .TestedProc = "Property Get Arry"
        
            .Verification = "Read from not-an-array returns Null"
                .Result = Arry(arr1, 1)
                .ResultExpected = Empty
                
            .Verification = "Read from not-an-array returns Nothing"
                ReDim arr1(3 To 4)
                .Result = Arry(arr1, 3)
                .ResultExpected = Empty
            
            .Verification = "Read from a 1-dim array"
                arr1 = TestArray(8) ' 9 elements
                .TimerStart
                .Result = Arry(arr1, 5)
                .TimerEnd
                .ResultExpected = "Item(5)"
            
            .Verification = "Empty is returned for an index outside the array's dimension specs"
                arr1 = TestArray(8) ' 9 elements
                .TimerStart
                .Result = Arry(arr1, 10)
                .TimerEnd
                .ResultExpected = Empty
            
            .Verification = "Read from a 3-dim array"
                arr1 = TestArray(3, 3, 3) ' 4 x 4 x 4 = 64 items
                .TimerStart
                .Result = Arry(arr1, "1, 2, 2")
                .TimerEnd
                .ResultExpected = "Item(1,2,2)"
            
            .Verification = "Read from a 3-dim array with an index outside the array's dimension specs"
                arr1 = TestArray(3, 3, 3) ' 27 elements
                .TimerStart
                .Result = Arry(arr1, Array(4, 2, 2))
                .TimerEnd
                .ResultExpected = Empty
            
            .Verification = "Return default from a non-existing or not allocated array"
                arr1 = TestArray(3, 3, 3) ' 27 elements
                .TimerStart
                Arry(arr1, Array(2, 2, 1)) = "Item(2,2,1) updated"
                .TimerEnd
                .Result = Arry(arr1, Array(2, 2, 1))
                .ResultExpected = "Item(2,2,1) updated"
        
            .Verification = "Return the array's type specific default for an index ouside the bounds of the provided array"
                arr1 = TestArray(10)
                .TimerStart
                .Result = Arry(arr1, 11)
                .TimerEnd
                .ResultExpected = Empty
        
            .Verification = "Returns the array's type specific default for an index outside the bounds of a 2-dim array"
                arr2 = TestArray(4, 4)
                .TimerStart
                .Result = Arry(arr2, "5, 4")
                .TimerEnd
                .ResultExpected = Empty
                
            .Verification = "An array of objects returns the object or Nothing"
                ReDim arr3(1 To 5)
                .Result = Arry(arr3, 2)
                .ResultExpected = Nothing
                
            .Verification = "An array of Booleans returns False for any Item out of bounds"
                ReDim arr4(1 To 5)
                .Result = Arry(arr4, 6)
                .ResultExpected = False
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0215_Arry_Let()
' ----------------------------------------------------------------------------
' Please note: Since the Arry-Let service is as universal as possible this
' test procedure has to cover/verify a great number of aspects in order to
' assure correctness and completeness.
' ----------------------------------------------------------------------------
    Const PROC = "Test_0215_Arry_Let"
    
    On Error GoTo eh
    Dim arr2    As Variant
    Dim arr3    As Variant
    Dim l       As Long
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0215-1"
        .Title = "Write 1-dim Array"
        .TestedProc = "Property Let Arry"
        
        .Verification = "Writing an element to a yet not allocated array without providing an index creates a 1-dim array with one Item"
            Set arr2 = Nothing
            .TimerStart
            Arry(arr2) = "Item(0)"
            .TimerEnd
            .Result = arr2(0)
            .ResultExpected = "Item(0)"
        
        .Verification = "Writing an element to a yet not allocated array returns a 1-dim array with one Item"
            .Comment = "Result is verified by means of Array(Get)"
            Set arr2 = Nothing
            .TimerStart
            Arry(arr2, 10) = "Item(10)"
            .TimerEnd
            .Result = arr2(10)
            .ResultExpected = "Item(10)"
        
        .Verification = "Adding an Item to a not yet allocated array with a certain index returns a 1-dim array with one Item at the given index"
            .Comment = "Result is verified by means of Array(Get)"
            .TimerStart
            ArryErase arr2
            Arry(arr2, 10) = "Item(10)"
            .TimerEnd
            .Result = Arry(arr2, 10)
            .ResultExpected = "Item(10)"
                    
        .Verification = "Adding an Item to an allocated array with a certain index expands the array as required"
            .Comment = "Result is verified by means of Array(Get)"
            .TimerStart
            Arry(arr2, 100) = "Item(100)"
            .TimerEnd
            .Result = Arry(arr2, 100)
            .ResultExpected = "Item(100)"
                    
        .TestId = "0215-2"
        .Title = "Write to multi-dim array with automated ReDim of any dimensions specs when the requested index is out of bounds"
        
        .Verification = "When an index is provided for a yet un-dimensioned multi-dim array it is dimensioned with the bounds provided by the indices"
            .TestedProc = "Property Let Arry" & vbLf & "Property Let Arry"
            Set arr2 = Nothing
            .Comment = "Result is verified by means of Array(Get)"
            Arry(arr2, "3,3") = "Item(3,3)"
            .Result = Arry(arr2, "3,3")
            .ResultExpected = "Item(3,3)"
        
        .Verification = "Write a new Item with the last dimensions index beyond its current boundary extents the array"
            .TestedProc = "Property Let Arry" & vbLf & "Property Let Arry"
            .Comment = "Result is verified by means of Array(Get)"
            arr3 = TestArray(4, 3, 2)
            .TimerStart
            Arry(arr3, ArryNdcs(l, 3, 3, 5)) = "Item(3,3,5)"
            .TimerEnd
            .Result = Arry(arr3, "3, 3, 5")
            .ResultExpected = "Item(3,3,5)"

        .Verification = "Write an Item to a yet not dimensioned arry"
            .Comment = "Result is verified by means of Array(Get)"
            .TestedProc = "Property Let Arry" & vbLf & "Property Let Arry"
            Set arr3 = Nothing
            .TimerStart
            Arry(arr3, "2,3,4") = "Item(2,3,4)"
            .Result = Arry(arr3, "2,3,4")
            .ResultExpected = "Item(2,3,4)"
        
        .Verification = "The multi-dim array created along with writing an Item has a from spec based on the ""Base Option"""
            .Result = LBound(arr3, 2)
            .ResultExpected = 0
            
        .Verification = "Write an Item to an already dimensioned but yet un-allocated array"
            .Comment = "Result is verified by means of Array(Get)"
            Set arr3 = Nothing
            ReDim arr3(1 To 2, 2 To 3, 3 To 4)
            .TimerStart
            Arry(arr3, "2,3,4") = "Item(2,3,4)"
            .Result = Arry(arr3, "2,3,4")
            .ResultExpected = "Item(2,3,4)"
            
        .Verification = "The multi-dim array created along with writing an Item has correctly considered the dimensions LBound"
            .Result = LBound(arr3, 2)
            .ResultExpected = 2
        
        .Verification = "When a multi-dim array's last dimension is expanded thi does not effect the lower bound of it"
            .TimerStart
            Arry(arr3, "2,3,6") = "Item(2,3,6)"
            .TimerEnd
            .Result = LBound(arr3, 2)
            .ResultExpected = 2
            
        .Verification = "Another but the last dimension's upper bound is extended by writing an item"
            .Comment = "Result is verified by means of Arry(Get)"
            Set arr3 = Nothing
            ReDim arr3(1 To 2, 2 To 3, 3 To 4)
            Arry(arr3, "2,3,4") = "Item(2,3,4)"
            Arry(arr3, "2,3,6") = "Item(2,3,6)" ' extends upper bound last dim
            .TimerStart
            Arry(arr3, "2,4,6") = "Item(2,4,6)" ' extends upper bound 2nd dim
            .TimerEnd
            .Result = Arry(arr3, "2,4,6")
            .ResultExpected = "Item(2,4,6)"
        
        .Verification = "Writing with the involment of ArryRedim did not effect any dimensions low bound"
            .Result = LBound(arr3, 2)
            .ResultExpected = 2
            
        If Not bModeRegression Then .ResultLogSummary
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0216_Arry_Let_Performance()
' ----------------------------------------------------------------------------
' The results of these tests are only displayed along with the regression test
' summary.
' ----------------------------------------------------------------------------
    Const PROC = "Test_0216_Arry_Let_Performance"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim i       As Long
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0216"
        .Title = "Performance comparison pure VBA versus Arry(Let) service"
        .TestedProc = "Property Let Arry"
        
            .Verification = "Writing 1000 elements into a correspondingly pre-dimensioned array by pure VBA"
                .Comment = "Test result is the difference of the exceuton time!"
                ReDim arr(1 To 1000)
                .TimerStart
                For i = 1 To 1000
                    arr(i) = "Item(" & i & ")"
                Next i
                .TimerEnd
                .Result = arr
                .ResultExpected = arr
            
            .Verification = "Writing 1000 elements into a correspondingly pre-dimensioned array by the Arry(Let) service"
                .Comment = "Test result is the difference of the exceuton time!"
                ReDim arr(1 To 1000)
                .TimerStart
                For i = 1 To 1000
                    Arry(arr, i) = "Item(" & i & ")"
                Next i
                .TimerEnd
                .Result = arr
                .ResultExpected = arr
            
        If Not bModeRegression Then .ResultLogSummary
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


Public Sub Test_0230_ArryCompare()
    Const PROC = "Test_0230_ArryCompare"
    
    On Error GoTo eh
    Dim aRes      As Variant
    Dim aExp      As Variant
    Dim dctDiff As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0230"
            .Title = "Compare a Result array with an expected array and return a Dictionary with the differences by irnoring empty items"
            .TestedProc = "Sub ArryCompare"
        
            .Verification = "Compare array with 4th Item different (stop at first difference)"
                aRes = Split("1,2,3,4,5,6,7", ",")
                aExp = Split("1,2,3,x,5,6,7", ",")
                .TimerStart
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .TimerEnd
                .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
                .ResultExpected = "'4' = 'x'"
            
            .Verification = "Compare array with 7th Item not exists in the first one (stop at first difference)"
                aRes = Split("1,2,3,4,5,6", ",")
                aExp = Split("1,2,3,4,5,6,7", ",")
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
                .ResultExpected = "'' =  '7'"
                
            .Verification = "Compare array with 7th Item not exists in the second one (stop at first difference)"
                aRes = Split("1,2,3,4,5,6,7", ",")
                aExp = Split("1,2,3,4,5,6", ",")
                .TimerStart
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .TimerEnd
                .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
                .ResultExpected = "'7' = ''"
            
            .Verification = "Compare array with 1st Item not exists in the second one (stop at first difference)"
                aRes = Split("1,2,3,4,5,6,7", ",")
                aExp = Split(",2,3,4,5,6,7", ",")
                
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .Result = dctDiff.Count
                .ResultExpected = 7
                        
            .Verification = "Compare array with 1st Item not exists in the first one (stop at first difference)"
                aRes = Split(",2,3,4,5,6,7", ",")
                aExp = Split("1,2,3,4,5,6,7", ",")
                .TimerStart
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .TimerEnd
                .Result = dctDiff.Count
                .ResultExpected = 7
            
            .Verification = "Compare array with several different items (stop at first difference)"
                aRes = Split("1,2,3,4,5,6,7", ",")
                aExp = Split("1,2,3,x,y,z,4,5,6,7", ",")
                .TimerStart
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .TimerEnd
                .Result = dctDiff.Count
                .ResultExpected = 7
            
            .Verification = "The arrays are equal, empty elements are ignored"
                aRes = Split("1,2,3,4,5,6,7,,,", ",")
                aExp = Split("1,2,3,4,5,6,7", ",")
                .TimerStart
                Set dctDiff = mBasic.ArryCompare(aRes, aExp)
                .TimerEnd
                .Result = dctDiff.Count
                .ResultExpected = 0
    
    End With
    
xt: EoP ErrSrc(PROC)

#If clsTrace = 1 Then
    If Not mErH.Regression Then
        Trc.Dsply
        Kill Trc.LogFile
    End If
#End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0240_ArryDiffers()
    Const PROC  As String = "Test_0240_ArryDiffers"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0240"
        .Title = "Compare arrays"
        .TestedProc = "Function ArryDiffers"
        
        .Verification = "Arrays not differ when only leading and trailing items are empty"
            a1 = Split(",1,2,3,4,5,6,7,,,,", ",")                   ' Test array
            a2 = Split(",,1,2,3,4,5,6,7,,", ",")                    ' Test array
            .Result = mBasic.ArryDiffers(a1, a2)
            .ResultExpected = True
        
        .Verification = "Items at different positions are empty, differ when not ignored"
            a1 = Split(",1,2,,,3,4,5,6,7,,,,", ",")                 ' Test array
            a2 = Split(",,1,,,,,,,,,2,3,4,,,5,6,7,,", ",")          ' Test array
            .Result = mBasic.ArryDiffers(a1, a2)
            .ResultExpected = True
        
        .Verification = "Items at different positions are empty, equal when ignored"
            a1 = Split("1,2,,,3,4,5,6,7,,,,", ",")                 ' Test array
            a2 = Split("1,2,,,3,4,5,6,7,,,,", ",")          ' Test array
            .Result = mBasic.ArryDiffers(a1, a2)
            .ResultExpected = False
        
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0245_ArryDims()

    Const PROC  As String = "Test_0240_ArryDiffers"
    
    On Error GoTo eh
    Dim arr         As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0245"
        .Title = "Array dimensions"
        
        .Verification = "ArryDims returns 1 for a 1-dim array"
            .TestedProc = "Function ArryDims"
            ReDim arr(3 To 5)
            .Result = ArryDims(arr)
            .ResultExpected = 1
            
        .Verification = "ArrySpecs()(1,1) returns the LBound of the first dimension"
            .TestedProc = "Function ArrySpecs"
            ReDim arr(3 To 5)
            .Result = ArrySpecs(arr, ArryDims(arr))(1, 1)
            .ResultExpected = 3
    
        .Verification = "ArryDims returns 1 for a 1-dim array"
            .TestedProc = "Function ArryDims"
            ReDim arr(3 To 5, 2 To 8, 1 To 4)
            .Result = ArryDims(arr)
            .ResultExpected = 3
            
        .Verification = "ArrySpecs()(2,2) returns the UBound of the second dimension"
            .TestedProc = "Function ArrySpecs"
            ReDim arr(3 To 5, 2 To 8, 1 To 4)
            .Result = ArrySpecs(arr, ArryDims(arr))(2, 2)
            .ResultExpected = 8
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0250_Arry_Various()
    Const PROC = "Test_0250_Arry_Various"
    
    Dim a As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "250"
        .Title = "Various array serviced"
        .TestedProc = "Function ArryIsAllocated"
        
        .Verification = "Number of dimensions"
        .Verification = "Array is not allocated"
            .Result = ArryIsAllocated(a)
            .ResultExpected = False
            
        .Verification = "Array is allocated"
            .Result = ArryIsAllocated(Array(1, 2, 3))
            .ResultExpected = True
    End With

xt: EoP ErrSrc(PROC)

End Sub

Public Sub Test_0270_ArryRemoveItems()
' ----------------------------------------------------------------------------
' Whitebox and regression test. Global error handling is used to monitor error
' condition tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_0270_ArryRemoveItems"

    On Error GoTo eh
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    Test_0271_ArryRemoveItems_Function
    Test_0272_ArryRemoveItems_Error_Conditions
    
xt: EoP ErrSrc(PROC)
#If clsTrace = 1 Then
    If Not mErH.Regression Then
        Trc.Dsply
        Kill Trc.LogFile
    End If
#End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0271_ArryRemoveItems_Function()
    Const PROC  As String = "Test_0271_ArryRemoveItems"
    
    On Error GoTo eh
    Dim arrExp  As Variant ' expected reult array
    Dim arrRes  As Variant ' result array
    
    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0271"
        .Title = "Remove indicated items from an array"
        .TestedProc = "Sub ArryRemoveItems"
        
        .Verification = "Items 4 and 5 are removed"
            ReDim arrRes(1 To 8):   arrRes = TestArray(arrRes)
            ReDim arrExp(1 To 6):   arrExp = TestArray(arrExp)
            arrExp(4) = "Item(6)"
            arrExp(5) = "Item(7)"
            arrExp(6) = "Item(8)"
            mBasic.ArryRemoveItems arrRes, 4, , 2
            .Result = arrRes
            .ResultExpected = arrExp
        
        .Verification = "Item 1 is removed (number = default = 1)"
            ReDim arrRes(1 To 8):   arrRes = TestArray(arrRes)
            ReDim arrExp(1 To 7):   arrExp = TestArray(arrExp)
            arrExp(1) = "Item(2)"
            arrExp(2) = "Item(3)"
            arrExp(3) = "Item(4)"
            arrExp(4) = "Item(5)"
            arrExp(5) = "Item(6)"
            arrExp(6) = "Item(7)"
            arrExp(7) = "Item(8)"
            mBasic.ArryRemoveItems arrRes, 1
            .Result = arrRes
            .ResultExpected = arrExp
        
        .Verification = "Remove Item 8 (the last Item)"
            ReDim arrRes(1 To 8):   arrRes = TestArray(arrRes)
            ReDim arrExp(1 To 7):   arrExp = TestArray(arrExp)
            arrExp(1) = "Item(1)"
            arrExp(2) = "Item(2)"
            arrExp(3) = "Item(3)"
            arrExp(4) = "Item(4)"
            arrExp(5) = "Item(5)"
            arrExp(6) = "Item(6)"
            arrExp(7) = "Item(7)"
            mBasic.ArryRemoveItems arrRes, 8
            .Result = arrRes
            .ResultExpected = arrExp
            
        .Verification = "Remove 2 beginning with the 3rd Item in an array -2 to 4"
            ReDim arrRes(-2 To 4):  arrRes = TestArray(arrRes)
            ReDim arrExp(-2 To 2):  arrExp = TestArray(arrExp)
            arrExp(-2) = "Item(-2)"
            arrExp(-1) = "Item(-1)"
            arrExp(0) = "Item(2)"
            arrExp(1) = "Item(3)"
            arrExp(2) = "Item(4)"
            mBasic.ArryRemoveItems arrRes, 3, , 2
            .Result = arrRes
            .ResultExpected = arrExp
            
        .Verification = "Remove from index 0 two items in an array -2 to 4"
            ReDim arrRes(-2 To 4):  arrRes = TestArray(arrRes)
            ReDim arrExp(-2 To 2):  arrExp = TestArray(arrExp)
            arrExp(-2) = "Item(-2)"
            arrExp(-1) = "Item(-1)"
            arrExp(0) = "Item(2)"
            arrExp(1) = "Item(3)"
            arrExp(2) = "Item(4)"
            mBasic.ArryRemoveItems arrRes, , 0, 2
            .Result = arrRes
            .ResultExpected = arrExp
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0275_ArryReDim()
    Const PROC = "Test_0275_ArryReDim"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim arrOut  As Variant

    Prepare
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0275-1"
        .Title = "Redim array basics"
                
        .Verification = "Number of items when any = Empty is ignored"
            .TestedProc = "Function ArryItems"
            ReDim arr(1 To 8)
            arr = TestArray(arr)
            arr(3) = Empty
            .TimerStart
            .Result = mBasic.ArryItems(arr, True)
            .TimerEnd
            .ResultExpected = 7
        
        .Verification = "Number of items when any Empty isn't ignored"
            .TestedProc = "Function ArryItems"
            .TimerStart
            .Result = mBasic.ArryItems(arr, False)
            .TimerEnd
            .ResultExpected = 8
        
        .TestId = "0275-2"
        .Title = "Redim array"
        
        '~~ Provide the expected Result array
        ReDim arrOut(1 To 2, 0 To 1, 1 To 6)
       
        .Verification = "The resulting array has two more (3) dimensions resulting in 24 items"
            .TestedProc = "Sub ArryReDim(S)" & vbLf & "Function ArryUnload(F)"
        
            ReDim arr(1 To 8)
            arr = TestArray(arr) ' filled with defaults
            arr(3) = Empty
            arrOut(1, 0, 1) = "Item(1)"
            arrOut(1, 0, 2) = "Item(2)"
            arrOut(1, 0, 4) = "Item(4)"
            arrOut(1, 0, 5) = "Item(5)"
            arrOut(1, 0, 6) = "Item(6)"
            .TimerStart
            mBasic.ArryReDim arr, "+:1 to 2", "+:0 to 1", "1:1 to 6"
            .TimerEnd
            .Result = mBasic.ArryDims(arr)
            .ResultExpected = 3
        
        .Verification = "The resulting array conforms with the expected"
            .TimerStart
            .Result = mBasic.ArryItems(arr, True)
            .TimerEnd
            .ResultExpected = mBasic.ArryItems(arrOut, True)
            
        .Verification = "The resulting array conforms with the expected"
            .Result = arr
            .ResultExpected = arrOut
            
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Sub

Public Sub Test_0272_ArryRemoveItems_Error_Conditions()
' ------------------------------------------------------------------------------
' Attention! Conditional Compile Argument 'CommonErHComp = 1' is required for
'            this test in order to have the raised error passed on to
'            the caller.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_0272_ArryRemoveItems_Error_Conditions"
    
    On Error GoTo eh
    Dim arrTest   As Variant
    Dim arr       As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    arrTest = Array(1, 2, 3, 4, 5, 6, 7, ",")
        
    ' Not an array
    Set arr = Nothing
    
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems arr, 2
    
    arr = arrTest
    ' Missing parameter
    mErH.Asserted AppErr(3) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems arr
    
    ' Element out of boundary
    mErH.Asserted AppErr(4) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems arr, 8
    
    ' Index out of boundary
    mErH.Asserted AppErr(5) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems arr, 7
    
    ' Element plus number of elements out of boundary
    mErH.Asserted AppErr(6) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems arr, 7, 2
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0280_ArryTrimm()
    Const PROC = "Test_0280_ArryTrimm"
    
    On Error GoTo eh
    Dim arrExp  As Variant
    Dim arrRes  As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    With TestAid
        .TestId = "0280"
        .Title = "Trim array by removing any empty items"
        .TestedProc = "ArryTrim (Sub)"
    
        .Verification = "The expected result has all empty items removed"
            arrRes = Array(, , 1, 2, 3, 4, 5, 6, 7, , , " ", " , ", vbCr, vbCrLf, vbLf)
            arrExp = Array(1, 2, 3, 4, 5, 6, 7, ",", vbCr, vbCrLf, vbLf)
            mBasic.ArryTrimm arrRes
            .Result = arrRes
            .ResultExpected = arrExp
                        
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0400_Spaced()
    Const PROC = "Test_0400_Spaced"
    
    Prepare
    mBasic.BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0400"
        .Title = "A provided string is returned spaced with non-breaking spaces `Chr$(160)`"
        .TestedProc = "Spaced (Function)"
        
        .Verification = "Result is the string with non-breaking spaces"
            .Result = mBasic.Spaced("Abc")
            .ResultExpected = "A" & Chr$(160) & "b" & Chr$(160) & "c"
        
        .Verification = "String spaced with non-breaking spaces, including existing spaces"
            .Result = mBasic.Spaced("Ab c")
            .ResultExpected = "A" & Chr$(160) & "b" & Chr$(160) & Chr$(160) & Chr$(160) & "c"
    End With

xt: mBasic.EoP ErrSrc(PROC)
End Sub

Public Sub Test_0500_Stack()
    Const PROC = "Test_0500_Stack"
    
    On Error GoTo eh
    Dim cllStack    As Collection:    Set cllStack = Nothing
    Dim iLevel      As Long
    Dim i           As Long
    Dim v           As Variant
    
    Prepare
    mBasic.BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0500"
        .Title = "Stack services"
        
        .Verification = "Returns TRUE when empty or stack not exists"
            .TestedProc = "StackIsEmpty (Function)"
            .ResultExpected = True
            .Result = mBasic.StackIsEmpty(cllStack)
               
        .Verification = "Returns FALSE when exists and not empty"
            .TestedProc = "StackIsEmpty (Function)"
            mBasic.StackPush cllStack, wsBasic
            .ResultExpected = False
            .Result = mBasic.StackIsEmpty(cllStack)
        
        .Verification = "Returns the object first stacked object"
            .TestedProc = "StackTop (Function)"
            Set cllStack = Nothing
            mBasic.StackPush cllStack, wsBasic
            Set v = mBasic.StackTop(cllStack)
            .ResultExpected = True
            .Result = v Is wsBasic
                
        .Verification = "Returns TRUE for a stacked object"
           .TestedProc = "StackEd (Function)"
            Set cllStack = Nothing
            mBasic.StackPush cllStack, wsBasic
            .ResultExpected = True
            .Result = mBasic.StackEd(cllStack, wsBasic, iLevel)
        
        .Verification = "Returns the stacked objects stack level"
            .ResultExpected = 1
            .Result = iLevel
                
        .Verification = "Returns TRUE when a given object is stacked at a given level"
            .TestedProc = "StackEd (Function)"
            Set cllStack = Nothing
            mBasic.StackPush cllStack, wsBasic
            .ResultExpected = True
            .Result = mBasic.StackEd(cllStack, wsBasic, 1)
                
        .Verification = "Returns the stacked object"
            .TestedProc = "StackTop (Function)"
            Set cllStack = Nothing
            mBasic.StackPush cllStack, wsBasic
            Set v = mBasic.StackTop(cllStack)
            .ResultExpected = True
            .Result = v Is wsBasic
        
        .Verification = "Returns the stacked object"
            .TestedProc = "StackPop (Function)"
            Set cllStack = Nothing
            mBasic.StackPush cllStack, wsBasic
            Set v = mBasic.StackPop(cllStack)
            .ResultExpected = True
            .Result = v Is wsBasic
            
        .Verification = "StackIsEmpty returns TRUE after StackPop"
            .ResultExpected = True
            .Result = mBasic.StackIsEmpty(cllStack)
        
        .Verification = "StackPop for an empty stack returns a vbNullString"
            .TestedProc = "StackPop (Function)"
            .ResultExpected = vbNullString
            .Result = mBasic.StackPop(cllStack)
        
        .Verification = "StackTop returns the stacked Item"
            .TestedProc = "StackTop (Function)"
            mBasic.StackPush cllStack, 10
            .ResultExpected = 10
            .Result = mBasic.StackTop(cllStack)
    
        .Verification = "StackEd returns the Item for a given level"
            .TestedProc = "StackPop (Function)"
            Set cllStack = Nothing
            For i = 1 To 10
                mBasic.StackPush cllStack, 10 * i
            Next i
            .ResultExpected = 80
            .Result = mBasic.StackEd(cllStack, , 8)
                
        .Verification = "StackPop end with an empty stack when all items are poped"
            .TestedProc = "StackPop (Function)"
            Set cllStack = Nothing
            For i = 1 To 10
                mBasic.StackPush cllStack, 10 * i
            Next i
            For i = 10 To 1 Step -1
                Debug.Assert mBasic.StackPop(cllStack) = 10 * i
            Next i
            .ResultExpected = True
            .Result = mBasic.StackIsEmpty(cllStack)
        
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_0600_TimedDoEvents()
    Const PROC = "Test_0600_TimedDoEvents"
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    mBasic.BoP ErrSrc(PROC)
    mBasic.TimedDoEvents ErrSrc(PROC)
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub Test_0700_Timer()
    Const PROC = "Test_0700_Timer"
    
    Dim i As Long
    Dim SecsMin     As Currency
    Dim SecsMax     As Currency
    Dim SecsWait    As Single
    Dim cBegin      As Currency
    Dim cEnd        As Currency
    Dim cElapsed    As Currency
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    mBasic.BoP ErrSrc(PROC)
    SecsWait = 0.000001
    
    SecsMin = 13000
    For i = 1 To 20
        mBasic.TimerBegin cBegin
        Application.Wait Now() + SecsWait
        mBasic.TimerEnd cBegin, , cEnd
        TimerEnd cBegin, cEnd, cElapsed, "00.0000"
        SecsMin = mBasic.Min(SecsMin, (cElapsed / mBasic.SysFrequency) * 130)
        SecsMax = mBasic.Max(SecsMax, (cElapsed / mBasic.SysFrequency) * 130)
    Next i
    Debug.Print """Application.Wait Now() + " & SecsWait & """ waits from min " & SecsMin * 130 & " to max " & SecsMax * 130 & " milliseconds (with a precision of " & mBasic.SysFrequency & " ticks per second)"
    mBasic.EoP ErrSrc(PROC), "Returned for ""Application.Wait Now() + 0,000001"": min=" & SecsMin & " milliseconds, max=" & SecsMax & " milliseconds"

End Sub

Public Sub Test_0800_Coll()
    Const PROC = "Test_0800_Coll"
    
    On Error GoTo eh
    Dim cll As Collection
    Dim obj As Object
    
    Prepare
    mBasic.BoP ErrSrc(PROC)
    
    With TestAid
        .TestId = "0800"
        .Title = "Collection service"
        
            .Verification = "Read from not existing Collection"
                .TestedProc = "Coll (Property Get)"
                .TimerStart
                .Result = Coll(cll)
                .TimerEnd
                .ResultExpected = Empty
                
            .Verification = "Read from not existing Collection index"
                .TestedProc = "Coll (Property Get)"
                Set cll = New Collection
                .TimerStart
                .Result = Coll(cll)
                .TimerEnd
                .ResultExpected = Empty
        
            .Verification = "Read with existing Collection index"
                .TestedProc = "Coll (Property Get)"
                Set cll = New Collection
                cll.Add "X"
                cll.Add "Y"
                .TimerStart
                .Result = Coll(cll, 2)
                .TimerEnd
                .ResultExpected = "Y"
        
            .Verification = "Read with not existing Collection index"
                .TestedProc = "Coll (Property Get)"
                Set cll = New Collection
                cll.Add "X"
                cll.Add "Y"
                .TimerStart
                .Result = Coll(cll, 3)
                .TimerEnd
                .ResultExpected = Empty
                
            .Verification = "Read with a non-integer argument returns the index of the found element"
                .TestedProc = "Coll (Property Get)"
                Set cll = New Collection
                cll.Add "X"
                cll.Add "Y"
                .TimerStart
                .Result = Coll(cll, "Y")
                .TimerEnd
                .ResultExpected = 2
            
            .Verification = "Read with a non-integer argument returns Empty for a not found element"
                .TestedProc = "Coll (Property Get)"
                Set cll = New Collection
                cll.Add "X"
                cll.Add "Y"
                .TimerStart
                .Result = Coll(cll, "Z")
                .TimerEnd
                .ResultExpected = Empty
            
            .Verification = "An object is returned as object"
                .TestedProc = "Coll (Property Let)" & vbLf & "Coll (Property Get)"
                Set cll = Nothing
                Coll(cll, 10) = ThisWorkbook
                .Result = Coll(cll, 10)
                .ResultExpected = ThisWorkbook
    
            .Verification = "A yet un-allocated object is returned as Nothing"
                .TestedProc = "Coll (Property Let)" & vbLf & "Coll (Property Get)"
                Set cll = New Collection
                Coll(cll, 10) = obj
                .Result = Coll(cll, 10)
                .ResultExpected = Nothing
        
            .Verification = "Write to not existing Collection"
                .TestedProc = "Coll (Property Let)"
                Set cll = Nothing
                Coll(cll) = "A"
                Coll(cll) = "B"
                .Result = cll.Count
                .ResultExpected = 2
            
            .Verification = "Write to a very high (10000) index (9998 indices below are filled with Empty)"
                .TestedProc = "Coll (Property Let)" & vbLf & "Coll (Property Get)"
                Set cll = Nothing
                Coll(cll) = "A"
                Coll(cll) = "B"
                .TimerStart
                Coll(cll, 10000) = "Z"
                .TimerEnd
                .Result = Coll(cll, 9999)
                .ResultExpected = Empty
                        
            .Verification = "Read from an existing index"
                .TimerStart
                .Result = Coll(cll, 10000)
                .TimerEnd
                .ResultExpected = "Z"
                
            .Verification = "Argument is not an integer/index results in the index of the found Item"
                .TimerStart
                .Result = Coll(cll, "Z")
                .TimerEnd
                .ResultExpected = 10000
    
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0850_Dict()

    Dim dct As Dictionary
    Dim i   As Long
    Dim k   As Long
    Dim l   As Long
    
    Dict(dct, "X") = 1                  ' add
    Dict(dct, "X") = 2                  ' ignore
    Debug.Assert Dict(dct, "X") = 1     ' Assert
    Dict(dct, "X", enReplace) = 2       ' replace
    Debug.Assert Dict(dct, "X") = 2     ' Assert
    
    Dict(dct, "X", enIncrement) = 5     ' Increment
    Debug.Assert Dict(dct, "X") = 7     ' Assert
    Dict(dct, "X", enIncrement) = -3    ' (de)increment
    Debug.Assert Dict(dct, "X") = 4     ' Assert
        
    '~~ Assert defaults when not existing
    Debug.Assert Dict(dct, "B") Is Nothing
    Dict(dct, "B") = vbNullString
    Debug.Assert Dict(dct, "B") = vbNullString
    
    '~~ Collect
    Set dct = Nothing
    Dict(dct, "A", enCollect) = "Axel"
    Dict(dct, "A", enCollect) = "Anton"
    Dict(dct, "A", enCollect) = "Abraham"
    Dict(dct, "B", enCollect) = "Bobo"
    Dict(dct, "B", enCollect) = "Bilbo"
    Dict(dct, "B", enCollect) = "Batman"
    Debug.Assert Dict(dct, "A").Count = 3
    Debug.Assert Dict(dct, "B").Count = 3
    Debug.Assert Dict(dct, "A")(1) = "Axel"
    Debug.Assert Dict(dct, "A")(2) = "Anton"
    Debug.Assert Dict(dct, "A")(3) = "Abraham"
    Debug.Assert Dict(dct, "B")(1) = "Bobo"
    Debug.Assert Dict(dct, "B")(2) = "Bilbo"
    Debug.Assert Dict(dct, "B")(3) = "Batman"
    
    '~~ Collect sorted
    Set dct = Nothing
    Dict(dct, "A", enCollectSorted) = "Axel"
    Dict(dct, "A", enCollectSorted) = "Anton"
    Dict(dct, "A", enCollectSorted) = "Abraham"
    Dict(dct, "B", enCollectSorted) = "Bobo"
    Dict(dct, "B", enCollectSorted) = "Bilbo"
    Dict(dct, "B", enCollectSorted) = "Batman"
    Debug.Assert Dict(dct, "A").Count = 3
    Debug.Assert Dict(dct, "A")(1) = "Abraham"
    Debug.Assert Dict(dct, "A")(2) = "Anton"
    Debug.Assert Dict(dct, "A")(3) = "Axel"
    Debug.Assert Dict(dct, "B")(1) = "Batman"
    Debug.Assert Dict(dct, "B")(2) = "Bilbo"
    Debug.Assert Dict(dct, "B")(3) = "Bobo"
    
    '~~ Performance
    l = 100000
    Set dct = Nothing
    With New clsTestAid
        .TimerStart
        For i = 1 To l
            k = Int((l * Rnd) + 1)
            Dict(dct, k) = vbNullString ' add, ignore duplicates
        Next i
        .TimerEnd
        Debug.Print vbLf & "Add ignore duplicates:"
        Debug.Print "======================"
        Debug.Print "Items added          = " & dct.Count
        Debug.Print "Duplicates ignored   = " & l - dct.Count
        Debug.Print "Elapsed milliseconds = " & .TimerExecTimeMsecs
    End With
    
    l = 50000
    Set dct = Nothing
    With New clsTestAid
        .TimerStart
        For i = 1 To l
            k = Int((l * Rnd) + 1)
            Dict(dct, k, enCollect) = k ' add, ignore duplicates
        Next i
        .TimerEnd
        Debug.Print vbLf & "Collect duplicates (unsorted):"
        Debug.Print "=============================="
        Debug.Print "Items added          = " & dct.Count
        Debug.Print "Items collected      = " & l - dct.Count
        Debug.Print "Elapsed milliseconds = " & .TimerExecTimeMsecs
    End With
       
    Set dct = Nothing
    
End Sub

Public Sub Test_ErrMsg()
    Const PROC = "Test_ErrMsg"
    
    On Error GoTo eh
    Err.Raise AppErr(10), ErrSrc(PROC), "This is an application error" & "||" & "This is an optional additional info about the error."
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

