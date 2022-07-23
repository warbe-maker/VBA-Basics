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
' W. Rauschenberger, Berlin Now 2021
' ----------------------------------------------------------------------------
Dim dctTest As Dictionary

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mBasicTest." & s:  End Property

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' indication.
' - Serves for a comprehensive display of an error message when the Common VBA
'   Error Services Component mErH is installed and the Conditional Compile
'   Argument 'ErHComp = 1'
' - Serves for the execution trace when the Common VBA Execution Trace Service
'   Component mTrc is installed and the Conditional Compile Argument
'   'ExecTrace = 1'.
' - May alternatively be copied as a Private procedure into any component when
'   this mBasic component is not installed but the mErH and or mTrc services
'   are
' Note: Because the error handling also hands over an 'End of Procedure' to the
'       mTrc component (when installed and 'ExecTrace = 1') an explicit call of
'       mTrc.EoP is only performed when mErH is not installed/in use.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EnvironmentVariables()
Dim i As Long
    For i = 1 To 100
        On Error Resume Next
        Debug.Print i & ". : " & VBA.Environ$(i) & """"
        If Err.Number <> 0 Then Exit For
    Next i
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' indication.
' - Serves for a comprehensive display of an error message when the Common VBA
'   Error Services Component mErH is installed and the Conditional Compile
'   Argument 'ErHComp = 1'
' - Serves for the execution trace when the Common VBA Execution Trace Service
'   Component mTrc is installed and the Conditional Compile Argument
'   'ExecTrace = 1'.
' - May alternatively be copied as a Private procedure into any component when
'   this mBasic component is not installed but the mErH and or mTrc services
'   are
' Note: Because the error handling also hands over an 'End of Procedure' to the
'       mTrc component (when installed and 'ExecTrace = 1') an explicit call of
'       mTrc.EoP is only performed when mErH is not installed/in use.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. Displays a debugging option button
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section when information is concatenated with
' the error message by two vertical bars (||).
'
' May be copied as Private Function into any module. Considers the Common VBA
' Message Service and the Common VBA Error Services as optional components.
' When neither is installed the error message is displayed by the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
' Note:  The above may seem to be a lot of code but will be a godsend in case
'        of an error!
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn tem into negative and in the error mesaage back into a positive
'          number.
' - ErrSrc To provide an unambigous procedure name - prefixed by the module name
'
' W. Rauschenberger Berlin, Nov 2021
'
' See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
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
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Public Sub Regression()
' ------------------------------------------------------------------------------
' This Regression test requires the Conditional Compile Argument 'ErHComp = 1'
' to run uninterrupted with all errors asserted.
'
' When not set the mErH.Regression = True setting and the mErH.Asserted settings
' will have no effect and all errors are displayed.
' ------------------------------------------------------------------------------
    Const PROC = "Regression"
    
    On Error GoTo eh
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    mTrc.LogTitle = "Execution Trace result of the mBasic Regression test"
    mTrc.LogTitle = "Regression Test module mTrc"
    
    mErH.Regression = True
    
    BoP ErrSrc(PROC)
    mBasicTest.Test_09_ErrMsg
    mBasicTest.Test_01_ArrayCompare
    mBasicTest.Test_02_0_ArrayRemoveItems
    mBasicTest.Test_03_ArrayToRange
    mBasicTest.Test_04_ArrayTrimm
    mBasicTest.Test_05_BaseName
    mBasicTest.Test_05_BaseName
    mBasicTest.Test_06_Spaced
    mBasicTest.Test_07_Align
    mBasicTest.Test_08_Stack
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub RegressionKeepLog()
    Dim sFile As String

#If ExecTrace = 1 Then
#If MsgComp = 1 Or ErHComp = 1 Then
    '~~ avoid the error message when the Conditional Compile Argument 'MsgComp = 0'!
    mTrc.Dsply
#End If
    '~~ Keep the regression test result
    With New FileSystemObject
        sFile = .GetParentFolderName(mTrc.LogFile) & "\RegressionTest.log"
        If .FileExists(sFile) Then .DeleteFile (sFile)
        .GetFile(mTrc.LogFile).Name = "RegressionTest.log"
    End With
    mTrc.Terminate
#End If

End Sub
Public Sub Test_09_ErrMsg()
    Const PROC = "Test_09_ErrMsg"
    
    On Error GoTo eh
    Dim l As Long
    
    BoP ErrSrc(PROC)
    '~~ 1. Test: Display an Application Error
    Test_09_1_ErrMsg 0
    
    '~~ 2. Test: Display a VB-Runtime-Error
    mErH.Asserted 6 ' Skip display when mErh.Regression = True
    l = l / 0
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_09_1_ErrMsg(ByVal test_value As Long)
' ------------------------------------------------------------------
' Display an Application Error istead of a VB Runtime Error
' ------------------------------------------------------------------
    Const PROC = "Test_09_1_ErrMsg"
    
    On Error GoTo eh
    Dim l As Long
    
    BoP ErrSrc(PROC)
    
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    If test_value = 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Application Error test: The argument 'test_value' must not be 0!"
    l = l / test_value
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_ArrayCompare()
    Const PROC  As String = "Test_03_ArrayToRange"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim dctDiff As Variant
    Dim v       As Variant
    
    BoP ErrSrc(PROC)
    
    '~~ Test 1: One element is different, empty elements are ignored
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,x,5,6,7", ",") ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 1
    For Each v In dctDiff
        Debug.Print "Test 1: Item/line " & v & vbLf & dctDiff(v)
    Next v
    
    '~~ Test 2: The first array has less elements, empty elements are ignored
    a1 = Split("1,2,3,4,5,6", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 1
    For Each v In dctDiff
        Debug.Print "Test 2: Item/line " & v & vbLf & dctDiff(v)
    Next v
        
    '~~ Test 3: The second array has less elements, empty elements are ignored
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6", ",") ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 1
    For Each v In dctDiff
        Debug.Print "Test 3: Item/line " & v & vbLf & dctDiff(v)
    Next v
    
    '~~ Test 4: The arrays first elements are different, empty elements are ignored
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split(",2,3,4,5,6,7", ",") ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 7
    For Each v In dctDiff
        Debug.Print "Test 4: Item/line " & v & vbLf & dctDiff(v)
    Next v
        
    '~~ Test 5: The arrays first elements are different, empty elements are ignored
    a1 = Split(",2,3,4,5,6,7", ",")     ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",")    ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 7
    For Each v In dctDiff
        Debug.Print "Test 5: Item/line " & v & vbLf & dctDiff(v)
    Next v
    
    '~~ Test 6: The second array has additional inserted elements, empty elements are ignored
    a1 = Split("1,2,3,4,5,6,7", ",")        ' Test array
    a2 = Split("1,2,3,x,y,z,4,5,6,7", ",")  ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 7
    For Each v In dctDiff
        Debug.Print "Test 6: Item/line " & v & vbLf & dctDiff(v)
    Next v
    
    '~~ Test 7: The arrays are equal, empty elements are ignored
    a1 = Split("1,2,3,4,5,6,7,,,", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 0
    For Each v In dctDiff
        Debug.Print "Test 7: Item/line " & v & vbLf & dctDiff(v)
    Next v
    
xt: EoP ErrSrc(PROC)
#If ExecTrace = 1 Then
    If Not mErH.Regression Then
        mTrc.Dsply
        Kill mTrc.LogFile
    End If
#End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_1_ArrayRemoveItems_Function()
    Const PROC  As String = "Test_02-1_ArrayRemoveItems_Function"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    Dim v       As Variant
    Dim i       As Long
    
    BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    
    a = aTest
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=3, ri_no_of_elements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"
    
    a = aTest
    mBasic.ArrayRemoveItems ri_va:=a, ri_index:=1
    Debug.Assert Join(a, ",") = "1,3,4,5,6,7"
    
    a = aTest
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=7
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"
    
    ReDim a(-2 To 4)
    i = LBound(a)
    For Each v In aTest
        a(i) = v: i = i + 1
    Next v
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=3, ri_no_of_elements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"

    ReDim a(2 To 8):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=3
    Debug.Assert Join(a, ",") = "1,2,4,5,6,7"

    ReDim a(0 To 6): i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems ri_va:=a, ri_index:=0
    Debug.Assert Join(a, ",") = "2,3,4,5,6,7"

    ReDim a(1 To 7):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems ri_va:=a, ri_index:=UBound(a)
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_2_ArrayRemoveItems_Error_Conditions()
' ------------------------------------------------------------------------------
' Attention! Conditional Compile Argument 'CommonErHComp = 1' is required for
'            this test in order to have the raised error passed on to
'            the caller.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_02_2_ArrayRemoveItems_Error_Conditions"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    
    BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
        
    ' Not an array
    Set a = Nothing
    
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=2
    
    a = aTest
    ' Missing parameter
    mErH.Asserted AppErr(3) ' skip display of error message when mErH.Regression = True
    mBasic.ArrayRemoveItems ri_va:=a
    
    ' Element out of boundary
    mErH.Asserted AppErr(4) ' skip display of error message when mErH.Regression = True
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=8
    
    ' Index out of boundary
    mErH.Asserted AppErr(5) ' skip display of error message when mErH.Regression = True
    mBasic.ArrayRemoveItems ri_va:=a, ri_index:=7
    
    ' Element plus number of elements out of boundary
    mErH.Asserted AppErr(6) ' skip display of error message when mErH.Regression = True
    mBasic.ArrayRemoveItems ri_va:=a, ri_element:=7, ri_no_of_elements:=2
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_0_ArrayRemoveItems()
' ----------------------------------------------------------------------------
' Whitebox and regression test. Global error handling is used to monitor error
' condition tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_0_ArrayRemoveItems"

    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_02_1_ArrayRemoveItems_Function
    Test_02_2_ArrayRemoveItems_Error_Conditions
    
xt: EoP ErrSrc(PROC)
#If ExecTrace = 1 Then
    If Not mErH.Regression Then
        mTrc.Dsply
        Kill mTrc.LogFile
    End If
#End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_03_ArrayToRange()
    Const PROC = "Test_03_ArrayToRange"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    a = aTest

    wsBasic.UsedRange.ClearContents
    mBasic.ArrayToRange a, wsBasic.celArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasic.rngArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasic.celArrayToRangeTarget
    mBasic.ArrayToRange a, wsBasic.rngArrayToRangeTarget

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_04_ArrayTrimm()
    Const PROC = "Test_04_ArrayTrimm"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    BoP ErrSrc(PROC)
    aTest = Split(" , ,1,2,3,4,5,6,7, , , ", ",") ' Test array
    a = aTest
    mBasic.ArrayTrimm a
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6,7"
    
    a = Split(" , , , , ", ",")
    mBasic.ArrayTrimm a
    Debug.Assert mBasic.ArrayIsAllocated(a) = False
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_05_BaseName()
' ----------------------------------------------------------------------------
' Please note: The common error handler (module mErH) is used in order to
'              allow an "unattended regression test" because the ErH passes on
'              the error number to the (this) entry procedure
' ----------------------------------------------------------------------------
    Const PROC  As String = "Test_05_BaseName"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim fl      As File
    
    '~~ Prepare for tests
    Set wb = ThisWorkbook
    With New FileSystemObject
        Set fl = .GetFile(wb.FullName)
    End With
    
    BoP ErrSrc(PROC)
    Debug.Assert mBasic.BaseName(wb) = "Basic"                    ' Test with Workbook object
    Debug.Assert mBasic.BaseName(fl) = "Basic"                    ' Test with File object
    Debug.Assert mBasic.BaseName(ThisWorkbook.Name) = "Basic"     ' Test with a file's name
    Debug.Assert mBasic.BaseName(ThisWorkbook.FullName) = "Basic" ' Test with a file's full name
    Debug.Assert mBasic.BaseName("xxxx") = "xxxx"
    
    '~~ Test unsupported object
    mErH.Asserted AppErr(1)
    mBasic.BaseName wb.Worksheets(1)
    
xt: EoP ErrSrc(PROC)
#If ExecTrace = 1 Then
    If Not mErH.Regression Then
        mTrc.Dsply
        Kill mTrc.LogFile
    End If
#End If
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_06_Spaced()
    Dim s As String
    s = Spaced("Ab c")
    Debug.Assert Replace(s, Chr$(160), " ") = "A b  c"
End Sub

Public Sub Test_07_Align()

    Debug.Assert Align("Abcde", 8, AlignLeft, " ", "-") = "Abcde --"
    Debug.Assert Align("Abcde", 8, AlignRight, " ", "-") = "-- Abcde"
    Debug.Assert Align("Abcde", 8, AlignCentered, " ", "-") = " Abcde -"
    Debug.Assert Align("Abcde", 7, AlignLeft, " ", "-") = "Abcde -"
    Debug.Assert Align("Abcde", 7, AlignRight, " ", "-") = "- Abcde"
    Debug.Assert Align("Abcde", 7, AlignCentered, " ", "-") = " Abcde "
    Debug.Assert Align("Abcde", 6, AlignLeft, " ", "-") = "Abcde "
    Debug.Assert Align("Abcde", 6, AlignRight, " ", "-") = " Abcde"
    Debug.Assert Align("Abcde", 6, AlignCentered, " ", "-") = " Abcd "
    Debug.Assert Align("Abcde", 5, AlignLeft, " ", "-") = "Abcde"
    Debug.Assert Align("Abcde", 5, AlignRight, " ", "-") = " Abcd"
    Debug.Assert Align("Abcde", 5, AlignCentered, " ", "-") = " Abc "
    
End Sub

Public Sub Test_08_Stack()
    Const PROC = "Test_08_Stack"
    
    On Error GoTo eh
    Dim Stack   As Collection:    Set Stack = Nothing
    Dim Level   As Long
    Dim i       As Long
    
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

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test___Timer()

    Dim i As Long
    Dim SecsMin     As Currency
    Dim SecsMax     As Currency
    Dim SecsElapsed As Currency
    Dim SecsWait    As Single
    
    SecsWait = 0.000001
    
    SecsMin = 100000
    For i = 1 To 20
        DoEvents
        mBasic.TimerBegin
        Application.Wait Now() + SecsWait
        SecsElapsed = TimerEnd
        SecsMin = mBasic.Min(SecsMin, SecsElapsed)
        SecsMax = mBasic.Max(SecsMax, SecsElapsed)
    Next i
    Debug.Print "Application.Wait Now() + " & SecsWait & " waits from " & SecsMin * 1000 & " to " & SecsMax * 1000 & " msec"
    
End Sub

