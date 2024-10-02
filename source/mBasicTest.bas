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
Public Trc As clsTrc

Private TestAid As New clsTestAid

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mBasicTest." & s:  End Property

Private Sub BoC(ByVal b_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common '(B)egin-(o)f-(C)ode' interface for the Common VBA Execution Trace
' service.
' - To be copied into any module which makes use of this trace service.
' - When the used Conditional Compile Argument is 0 or not set at all.
' - Important! The begin id (b_id) has to be identical with the paired EoC
'              statement.
' ------------------------------------------------------------------------------
    Dim s As String
    If Not IsMissing(b_arguments) Then s = Join(b_arguments, ",")

#If mTrc = 1 Then
    mTrc.BoC b_id, s
#ElseIf clsTrc = 1 Then
    Trc.BoC b_id, s
#End If

End Sub

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

Private Sub EoC(ByVal e_id As String, ParamArray e_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common '(E)nd-(o)f-(C)ode' interface for the Common VBA Execution Trace
' service.
' - To be copied into any module which makes use of this trace service.
' - When the used Conditional Compile Argument is 0 or not set at all.
' - Important! The end id (b_id) has to be identical with the paired EoC
'              statement.
' ------------------------------------------------------------------------------
    Dim s As String
    If Not IsMissing(e_arguments) Then s = Join(e_arguments, ",")

#If mTrc = 1 Then
    mTrc.EoC e_id, s
#ElseIf clsTrc = 1 Then
    Trc.EoC e_id, s
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
    
    If Not mErH.Regression Then
        Set TestAid = Nothing: Set TestAid = New clsTestAid
    Else
        If TestAid Is Nothing Then Set TestAid = New clsTestAid
    End If
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    If TestAid.ModeRegression Then
        Trc.FileFullName = TestAid.TestFolder & "\TestExecution.trc"
    Else
        Trc.FileFullName = TestAid.TestFolder & "\RegressionTestExecution.trc"
    End If
    TestAid.TestedComp = "mBasic"
    
End Sub

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
    mErH.Regression = True
    Prepare
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
#If mTrc = 1 Then
    mTrc.FileName = TestAid.TestFolder & "\RegressionTest.ExecTrace.log"
    mTrc.Title = "Execution Trace result of the mBasic Regression test"
    mTrc.NewFile
#ElseIf clsTrc = 1 Then
    Set Trc = New clsTrc
    With Trc
        .FileFullName = TestAid.TestFolder & "\RegressionTest.ExecTrace.log"
        .Title = "Execution Trace result of the mBasic Regression test"
        .NewFile
    End With
#End If

    mErH.Regression = True
    TestAid.ModeRegression = True
    
    BoP ErrSrc(PROC)
    mBasicTest.Test_0100_Align
    
    mBasicTest.Test_0100_Align
    mBasicTest.Test_0100_Align
    mBasicTest.Test_0200_Arry_Get_Let
    mBasicTest.Test_0210_ArryAsRange_RangeAsArray
    mBasicTest.Test_0220_ArryCompare
    mBasicTest.Test_0230_ArryDiffers
    mBasicTest.Test_0240_ArryIsAllocated
    mBasicTest.Test_0250_ArryNoOfDims
    mBasicTest.Test_0260_ArryRemoveItems
    mBasicTest.Test_0270_ArryTrimm
    mBasicTest.Test_0300_BaseName
    mBasicTest.Test_0300_BaseName
    mBasicTest.Test_0400_Spaced
    mBasicTest.Test_0500_Stack
    mBasicTest.Test_0600_TimedDoEvents
    mBasicTest.Test_0700_Timer
        
    TestAid.ResultSummaryLog
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
#If mTrc = 1 Then
    mTrc.Dsply
#ElseIf clsTrc = 1 Then
    Trc.Dsply
#End If
    TestAid.ResultSummaryLog
    TestAid.CleanUp
    Set TestAid = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0210_ArryAsRange_RangeAsArray()
    Const PROC = "Test_0210_ArryAsRange_RangeAsArray"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    Dim a2      As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    Prepare
    BoP ErrSrc(PROC)
    aTest = Split("X,2,3,4,5,AAAA,Z", ",") ' Test array
    a = aTest

    With TestAid
        .TestedProc = "ArryAsRange, RangeAsArray"
        
        .TestNumber = "0210-01"
        .TestHeadLine = "Array to range, range to array"
        wsBasic.UsedRange.ClearContents
        mBasic.ArryAsRange a, wsBasic.rngArryAsRangeTarget, True
        a2 = mBasic.RangeAsArray(Intersect(wsBasic.rngArryAsRangeTarget, wsBasic.UsedRange))
        .ResultExpected = a
        .Result = a2
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0100_Align()
    Const PROC = "Test_0100_Align"
    
    Prepare
    
    mBasic.BoP ErrSrc(PROC)
    With TestAid
        .TestedProc = "Align"
        
        .TestNumber = "0100-01"
        .TestHeadLine = "Align ""Abcde"" left 8, filled with "" -"""      ' Left      8   " -"
        .ResultExpected = "Abcde ----"
        .Result = mBasic.Align("Abcde", enAlignLeft, 8, " -")
        
        .TestNumber = "0100-02"
        .TestHeadLine = "Align ""Abcde"" right 8, filled with ""- """     ' Right     8   "- "
        .ResultExpected = "---- Abcde"
        .Result = mBasic.Align("Abcde", enAlignRight, 8, "- ")
        
        .TestNumber = "0100-03"
        .TestHeadLine = "Align ""Abcde"" centered 8, filled with ""- """  ' Centered  8   "- "
        .ResultExpected = "-- Abcde ---"
        .Result = mBasic.Align("Abcde", enAlignCentered, 8, " -")
        
        .TestNumber = "0100-04"
        .TestHeadLine = "Align ""Abcde"" centered 7, filled with ""- """  ' Centered  7   " -"
        .ResultExpected = "-- Abcde --"
        .Result = mBasic.Align("Abcde", enAlignCentered, 7, " -")
        
        .TestNumber = "0100-05"
        .TestHeadLine = "Align ""Abcde"" right 7, filled with ""- """     ' Right     7   "- "
        .ResultExpected = "--- Abcde"
        .Result = mBasic.Align("Abcde", enAlignRight, 7, "- ")
                
        .TestNumber = "0100-06"
        .TestHeadLine = "Align ""Abcde"" centered 7, filled with ""-"""   ' Centered  7   "-"
        .ResultExpected = "--Abcde--"
        .Result = mBasic.Align("Abcde", enAlignCentered, 7, "-")
        
        .TestNumber = "0100-07"
        .TestHeadLine = "Align ""Abcde"" left 7, filled with ""-"""       ' Left      7   "-"
        .ResultExpected = "Abcde---"
        .Result = mBasic.Align("Abcde", enAlignLeft, 7, "-")
        
        .TestNumber = "0100-08"
        .TestHeadLine = "Align ""Abcde"" right 6, filled with ""-"""      ' Right     6   "-"
        .ResultExpected = "--Abcde"
        .Result = mBasic.Align("Abcde", enAlignRight, 6, "-")
        
        .TestNumber = "0100-09"
        .TestHeadLine = "Align ""Abcde"" centered 6, filled with ""-"""   ' Centered  6   "-"
        .ResultExpected = "-Abcde--"
        .Result = mBasic.Align("Abcde", enAlignCentered, 6, "-")
        
        .TestNumber = "0100-10"
        .TestHeadLine = "Align ""Abcde"" left 4, filled with ""-"""       ' Left      4   "-"
        .ResultExpected = "Abcd-"
        .Result = mBasic.Align("Abcde", enAlignLeft, 4, "-")
        
        .TestNumber = "0100-11"
        .TestHeadLine = "Align ""Abcde"" right 4, filled with ""-"""      ' Right     4   "-"
        .ResultExpected = "-Abcd"
        .Result = mBasic.Align("Abcde", enAlignRight, 4, "-")
        
        .TestNumber = "0100-12"
        .TestHeadLine = "Align ""Abcde"" centered 4, filled with ""-"", with margin"
        .ResultExpected = " -Abcd- "
        .Result = mBasic.Align("Abcde", enAlignCentered, 4, "-", " ")   ' Centered  4   "-"
    
    End With
    
xt: mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub Test_0200_Arry_Get_Let()
    Const PROC = "Test_0200_Arry_Get_Let"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim dctDiff As Variant
    Dim v       As Variant
    
    Prepare
    BoP ErrSrc(PROC)
    With TestAid
        .TestedProc = "Arry"
        
        .TestNumber = "0200-01"
        .TestHeadLine = "Write/Read array"
        Arry(a1) = "A"
        .ResultExpected = "A"
        v = Arry(a1, LBound(a1))
        .Result = v
        
        .TestNumber = "0200-02"
        .TestHeadLine = "Write at any index, and read"
        Arry(a1, 10) = "Z"
        .ResultExpected = "Z"
        v = Arry(a1, 10)
        .Result = v
        
        .TestNumber = "0200-03"
        .TestHeadLine = "Replace at any index, and read"
        Arry(a1, 10) = "Y"
        .ResultExpected = "Y"
        v = Arry(a1, 10)
        .Result = v
    
        .TestNumber = "0200-04"
        .TestHeadLine = "Return default from a non-existing or not allocated array"
        Arry(a1, 10) = "Y"
        .ResultExpected = vbNullString
        v = Arry(a1, 11)
        .Result = v
    
        .TestNumber = "0200-04"
        .TestHeadLine = "Return 0 as the default from a non-existing or not allocated array"
        .ResultExpected = 0
        v = Arry(a1, 11, 0)
        .Result = 0
    
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0220_ArryCompare()
    Const PROC = "Test_0220_ArryCompare"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim dctDiff As Variant
    Dim v       As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    Prepare
    With TestAid
        .TestedProc = "ArryCompare"
        .TestedType = "Sub"
        
        ' ==================================================================
        .TestNumber = "0220-01"
        .TestHeadLine = "One element is different, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
        a2 = Split("1,2,3,x,5,6,7", ",") ' Test array
        
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        
        .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
        .ResultExpected = "'4' = 'x'"
            
        ' ==================================================================
        .TestNumber = "0220-02"
        .TestHeadLine = "The first array has less elements, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6", ",") ' Test array
        a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
        
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        
        .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
        .ResultExpected = "'' =  '7'"
            
        ' ==================================================================
        .TestNumber = "0220-03"
        .TestHeadLine = "The second array has less elements, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
        a2 = Split("1,2,3,4,5,6", ",") ' Test array
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        .Result = Replace(dctDiff.Items()(0), vbLf, " = ")
        .ResultExpected = "'7' = ''"
        
        ' ==================================================================
        .TestNumber = "0220-04"
        .TestHeadLine = "The arrays first elements are different, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
        a2 = Split(",2,3,4,5,6,7", ",") ' Test array
        
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        
        .Result = dctDiff.Count
        .ResultExpected = 7
                    
        ' ==================================================================
        .TestNumber = "0220-05"
        .TestHeadLine = "The arrays first elements are different, empty elements are ignored"
        a1 = Split(",2,3,4,5,6,7", ",")     ' Test array
        a2 = Split("1,2,3,4,5,6,7", ",")    ' Test array
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        .Result = dctDiff.Count
        .ResultExpected = 7
        
        ' ==================================================================
        .TestNumber = "0220-06"
        .TestHeadLine = "The second array has additional inserted elements, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6,7", ",")        ' Test array
        a2 = Split("1,2,3,x,y,z,4,5,6,7", ",")  ' Test array
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
        .Result = dctDiff.Count
        .ResultExpected = 7
        
        ' ==================================================================
        .TestNumber = "0220-07"
        .TestHeadLine = "The arrays are equal, empty elements are ignored"
        a1 = Split("1,2,3,4,5,6,7,,,", ",") ' Test array
        a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
        BoC "mBasic.ArryCompare"
        Set dctDiff = mBasic.ArryCompare(a_v1:=a1 _
                                        , a_v2:=a2 _
                                         )
        EoC "mBasic.ArryCompare"
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

Public Sub Test_0230_ArryDiffers()
    Const PROC  As String = "Test_0230_ArryDiffers"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    
    '~~ Test 1: Only leading and trailing items are empty
    a1 = Split(",1,2,3,4,5,6,7,,,,", ",")                   ' Test array
    a2 = Split(",,1,2,3,4,5,6,7,,", ",")                    ' Test array
    Debug.Assert Not mBasic.ArryDiffers(ad_v1:=a1 _
                                      , ad_v2:=a2 _
                                      , ad_ignore_empty_items:=False)
    Debug.Assert Not mBasic.ArryDiffers(ad_v1:=a1 _
                                      , ad_v2:=a2 _
                                      , ad_ignore_empty_items:=True)
    
    '~~ Test 2: Various numbers of items at different positions are empty
    a1 = Split(",1,2,,,3,4,5,6,7,,,,", ",")                 ' Test array
    a2 = Split(",,1,,,,,,,,,2,3,4,,,5,6,7,,", ",")          ' Test array
    Debug.Assert mBasic.ArryDiffers(ad_v1:=a1 _
                                  , ad_v2:=a2 _
                                  , ad_ignore_empty_items:=False)
    Debug.Assert Not mBasic.ArryDiffers(ad_v1:=a1 _
                                      , ad_v2:=a2 _
                                      , ad_ignore_empty_items:=True)
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0240_ArryIsAllocated()

End Sub

Public Sub Test_0250_ArryNoOfDims()

End Sub

Public Sub Test_0260_ArryRemoveItems()
' ----------------------------------------------------------------------------
' Whitebox and regression test. Global error handling is used to monitor error
' condition tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_0260_ArryRemoveItems"

    On Error GoTo eh
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    Test_0261_ArryRemoveItems_Function
    Test_0262_ArryRemoveItems_Error_Conditions
    
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

Public Sub Test_0261_ArryRemoveItems_Function()
    Const PROC  As String = "Test_0261_ArryRemoveItems_Function"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    Dim v       As Variant
    Dim i       As Long
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    
    a = aTest
    mBasic.ArryRemoveItems a_va:=a, a_element:=3, a_no_of_elements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"
    
    a = aTest
    mBasic.ArryRemoveItems a_va:=a, a_index:=1
    Debug.Assert Join(a, ",") = "1,3,4,5,6,7"
    
    a = aTest
    mBasic.ArryRemoveItems a_va:=a, a_element:=7
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"
    
    ReDim a(-2 To 4)
    i = LBound(a)
    For Each v In aTest
        a(i) = v: i = i + 1
    Next v
    mBasic.ArryRemoveItems a_va:=a, a_element:=3, a_no_of_elements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"

    ReDim a(2 To 8):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArryRemoveItems a_va:=a, a_element:=3
    Debug.Assert Join(a, ",") = "1,2,4,5,6,7"

    ReDim a(0 To 6): i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArryRemoveItems a_va:=a, a_index:=0
    Debug.Assert Join(a, ",") = "2,3,4,5,6,7"

    ReDim a(1 To 7):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArryRemoveItems a_va:=a, a_index:=UBound(a)
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0262_ArryRemoveItems_Error_Conditions()
' ------------------------------------------------------------------------------
' Attention! Conditional Compile Argument 'CommonErHComp = 1' is required for
'            this test in order to have the raised error passed on to
'            the caller.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_0262_ArryRemoveItems_Error_Conditions"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
        
    ' Not an array
    Set a = Nothing
    
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems a_va:=a, a_element:=2
    
    a = aTest
    ' Missing parameter
    mErH.Asserted AppErr(3) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems a_va:=a
    
    ' Element out of boundary
    mErH.Asserted AppErr(4) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems a_va:=a, a_element:=8
    
    ' Index out of boundary
    mErH.Asserted AppErr(5) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems a_va:=a, a_index:=7
    
    ' Element plus number of elements out of boundary
    mErH.Asserted AppErr(6) ' skip display of error message when mErH.Regression = True
    mBasic.ArryRemoveItems a_va:=a, a_element:=7, a_no_of_elements:=2
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0270_ArryTrimm()
    Const PROC = "Test_0270_ArryTrimm"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    BoP ErrSrc(PROC)
    aTest = Split(" , ,1,2,3,4,5,6,7, , , ", ",") ' Test array
    a = aTest
    mBasic.ArryTrimm a
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6,7"
    
    a = Split(" , , , , ", ",")
    mBasic.ArryTrimm a
    Debug.Assert mBasic.ArryIsAllocated(a) = False
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0300_BaseName()
' ----------------------------------------------------------------------------
' Please note: The common error handler (module mErH) is used in order to
'              allow an "unattended regression test" because the ErH passes on
'              the error number to the (this) entry procedure
' ----------------------------------------------------------------------------
    Const PROC  As String = "Test_0300_BaseName"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim fl      As File
    
    '~~ Prepare for tests
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    Set wb = ThisWorkbook
    With New FileSystemObject
        Set fl = .GetFile(wb.FullName)
    End With
    
    BoP ErrSrc(PROC)
    Debug.Assert mBasic.BaseName(wb) = "Basic"                    ' Test with Workbook object
    Debug.Assert mBasic.BaseName(fl) = "Basic"                    ' Test with File object
    Debug.Assert mBasic.BaseName(ThisWorkbook.Name) = "Basic"     ' Test with a file's name
    Debug.Assert mBasic.BaseName(ThisWorkbook.FullName) = "Basic" ' Test with a file's full name
    
    '~~ Test unsupported object
    mErH.Asserted AppErr(1)
    mBasic.BaseName wb.Worksheets(1)
    
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

Public Sub Test_0400_Spaced()
    Const PROC = "Test_0400_Spaced"
    
    Dim s As String
    
    Prepare
    mBasic.BoP ErrSrc(PROC)
    With TestAid
        .TestedProc = "Spaced"
        .TestedType = "Function"
        
        .TestNumber = "0400-1"
        .TestHeadLine = "A provided string is returned spaced with non-breaking spaces `Chr$(160)`"
        .Result = Spaced("Ab c")
        .ResultExpected = "A" & Chr$(160) & "b" & Chr$(160) & Chr$(160) & "c"
    End With

xt: mBasic.EoP ErrSrc(PROC)
End Sub

Public Sub Test_0500_Stack()
    Const PROC = "Test_0500_Stack"
    
    On Error GoTo eh
    Dim Stack   As Collection:    Set Stack = Nothing
    Dim Level   As Long
    Dim i       As Long
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' when tested individually
    mBasic.BoP ErrSrc(PROC)
    ' Test 1: Push/Pop an object
    Debug.Assert mBasic.StackIsEmpty(Stack) = True
    mBasic.StackPush Stack, wsBasic
    Debug.Assert mBasic.StackIsEmpty(Stack) = False
    Debug.Assert mBasic.StackTop(Stack) Is wsBasic
    Debug.Assert mBasic.StackEd(Stack, wsBasic, Level) = True
    Debug.Assert Level = 1
    Debug.Assert mBasic.StackEd(Stack, wsBasic, 1) = True
    Debug.Assert mBasic.StackTop(Stack) Is wsBasic
    Debug.Assert mBasic.StackPop(Stack) Is wsBasic
    Debug.Assert mBasic.StackIsEmpty(Stack) = True
    Debug.Assert mBasic.StackPop(Stack) = vbNullString
    Debug.Assert mBasic.StackTop(Stack) = vbNullString
    
    ' Test 2: Push/Pop a numeric item
    Level = 0
    Debug.Assert mBasic.StackIsEmpty(Stack) = True
    mBasic.StackPush Stack, 10
    Debug.Assert mBasic.StackIsEmpty(Stack) = False
    Debug.Assert mBasic.StackTop(Stack) = 10
    Debug.Assert mBasic.StackEd(Stack, 10, Level) = True
    Debug.Assert Level = 1
    Debug.Assert mBasic.StackEd(Stack, 10, 1) = True
    Debug.Assert mBasic.StackPop(Stack) = 10
    Debug.Assert mBasic.StackIsEmpty(Stack) = True
    Debug.Assert mBasic.StackPop(Stack) = vbNullString
    Debug.Assert mBasic.StackTop(Stack) = vbNullString
    Set Stack = Nothing

    For i = 1 To 10
        mBasic.StackPush Stack, 10 * i
    Next i
    Debug.Assert mBasic.StackTop(Stack) = 10 * (i - 1)
    Debug.Assert mBasic.StackEd(Stack, , 8) = 80
    
    For i = 10 To 1 Step -1
        Debug.Assert mBasic.StackPop(Stack) = 10 * i
    Next i

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
    Dim SecsElapsed As Currency
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

