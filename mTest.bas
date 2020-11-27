Attribute VB_Name = "mTest"
Option Private Module
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mTest: Dedicate fo the test of procedures in the mBasic
'       module.
'
' Note: Procedures of the mBasic module do not use the Common VBA Error Handler.
'       However, this test module uses the mErrHndlr module for test purpose.
'
' W. Rauschenberger, Berlin Sept 2020
' ----------------------------------------------------------------------------
Dim dctTest As Dictionary

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mTest." & s:  End Property

Private Sub EnvironmentVariables()
Dim i As Long
    For i = 1 To 100
        On Error Resume Next
        Debug.Print i & ". : " & Environ$(i) & """"
        If err.Number <> 0 Then Exit For
    Next i
End Sub

Public Sub Regression()
    Const PROC = "Regression"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    mTest.Test_01_ArrayCompare
    mTest.Test_02_ArrayRemoveItems
    mTest.Test_03_ArrayRemoveItems_Error_Conditions
    mTest.Test_04_ArrayRemoveItems_Error_Display
    mTest.Test_05_ArrayRemoveItems_Function
    mTest.Test_06_ArrayToRange
    mTest.Test_07_ArrayTrimm
    mTest.Test_08_BaseName
    mErH.EoP ErrSrc(PROC)

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
    
End Sub

Public Sub Test_01_ArrayCompare()
    Const PROC  As String = "Test_06_ArrayToRange"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim aDiff   As Variant
    
    mErH.BoP ErrSrc(PROC)
    '~~ Test 1: One element is different
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,x,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2 _
                                )
    Debug.Assert UBound(aDiff) = 0
                                
    
    '~~ Test 2: The first array has less elements
    a1 = Split("1,2,3,4,5,6", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    Debug.Assert UBound(aDiff) = 0
    
    
    '~~ Test 3: The second array has less elements
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    Debug.Assert UBound(aDiff) = 0
    
    '~~ Test 4: The arrays first elements are different (element in second array is empty)
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split(",2,3,4,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    Debug.Assert UBound(aDiff) >= 0
    
    
    '~~ Test 5: The arrays first elements are different (element in first array is empty)
    a1 = Split(",2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    Debug.Assert UBound(aDiff) >= 0
    
    '~~ Test 5: The second array has additional inserted elements
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,x,y,z,4,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    Debug.Assert UBound(aDiff) >= 0
    
    '~~ Test 6: The arrays are equal
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    aDiff = mBasic.ArrayCompare(ac_a1:=a1, _
                                ac_a2:=a2)
    On Error Resume Next
    Debug.Assert UBound(aDiff) >= 0
    Debug.Assert err.Number = 9
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_02_ArrayRemoveItems()
' ---------------------------------
' Whitebox and regression test.
' Global error handling is used to
' monitor error condition tests.
' ---------------------------------
Const PROC  As String = "Test_02_ArrayRemoveItems"

    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    Test_05_ArrayRemoveItems_Function
    Test_03_ArrayRemoveItems_Error_Conditions
    Test_04_ArrayRemoveItems_Error_Display
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_03_ArrayRemoveItems_Error_Conditions()
    Const PROC  As String = "Test_03_ArrayRemoveItems_Error_Conditions"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    
    mErH.BoP ErrSrc(PROC)

    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
        
    ' Not an array
    Set a = Nothing
    On Error Resume Next
    mBasic.ArrayRemoveItems a, 2
    Debug.Assert mErH.AppErr(err.Number) = 1
    
    a = aTest
    ' Missing parameter
    On Error Resume Next
    mBasic.ArrayRemoveItems a
    Debug.Assert mErH.AppErr(err.Number) = 3
    
    ' Element out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=8
    Debug.Assert mErH.AppErr(err.Number) = 4
    
    ' Index out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Index:=7
    Debug.Assert mErH.AppErr(err.Number) = 5
    
    ' Element plus number of elements out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=7, NoOfElements:=2
    Debug.Assert mErH.AppErr(err.Number) = 6

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_04_ArrayRemoveItems_Error_Display()
    Const PROC  As String = "Test_04_ArrayRemoveItems_Error_Display"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    Dim v       As Variant
    Dim i       As Long

    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    
    ReDim a(5, 2 To 8):    i = LBound(a, 2)
    For Each v In aTest
        a(1, i) = v:  i = i + 1
    Next v
    
    mErH.BoP ErrSrc(PROC)
    mBasic.ArrayRemoveItems a, Element:=3

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_05_ArrayRemoveItems_Function()
    Const PROC  As String = "Test_05_ArrayRemoveItems_Function"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    Dim v       As Variant
    Dim i       As Long
    
    mErH.BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    
    a = aTest
    mBasic.ArrayRemoveItems a, Element:=3, NoOfElements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"
    
    a = aTest
    mBasic.ArrayRemoveItems a, Index:=1
    Debug.Assert Join(a, ",") = "1,3,4,5,6,7"
    
    a = aTest
    mBasic.ArrayRemoveItems a, Element:=7
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"
    
    ReDim a(-2 To 4)
    i = LBound(a)
    For Each v In aTest
        a(i) = v: i = i + 1
    Next v
    mBasic.ArrayRemoveItems a, Element:=3, NoOfElements:=2
    Debug.Assert Join(a, ",") = "1,2,5,6,7"

    ReDim a(2 To 8):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems a, Element:=3
    Debug.Assert Join(a, ",") = "1,2,4,5,6,7"

    ReDim a(0 To 6): i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems a, Index:=0
    Debug.Assert Join(a, ",") = "2,3,4,5,6,7"

    ReDim a(1 To 7):    i = LBound(a)
    For Each v In aTest
        a(i) = v:   i = i + 1
    Next v
    mBasic.ArrayRemoveItems a, Index:=UBound(a)
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6"

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_06_ArrayToRange()
    Const PROC  As String = "Test_06_ArrayToRange"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    mErH.BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    a = aTest

    wsBasicTest.UsedRange.ClearContents
    mBasic.ArrayToRange a, wsBasicTest.celArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasicTest.rngArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasicTest.celArrayToRangeTarget
    mBasic.ArrayToRange a, wsBasicTest.rngArrayToRangeTarget

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_07_ArrayTrimm()
    Const PROC = "Test_07_ArrayTrimm"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    mErH.BoP ErrSrc(PROC)
    aTest = Split(" , ,1,2,3,4,5,6,7, , , ", ",") ' Test array
    a = aTest
    mBasic.ArrayTrimm a
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6,7"
    
    a = Split(" , , , , ", ",")
    mBasic.ArrayTrimm a
    Debug.Assert mBasic.ArrayIsAllocated(a) = False
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_08_BaseName()
' -----------------------------------------------------
' Please note:
' The common error handler (module mErrHndlr) is used
' in order to allow an "unattended regression test"
' because the ErrHndlr passes on the error number to
' the (this) entry procedure
' -----------------------------------------------------
    Const PROC  As String = "Test_08_BaseName"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim fl      As File
    
    '~~ Prepare for tests
    Set wb = ThisWorkbook
    With New FileSystemObject
        Set fl = .GetFile(wb.FullName)
    End With
    
    mErH.BoP ErrSrc(PROC)
    Debug.Assert mBasic.BaseName(wb) = "Basic"                    ' Test with Workbook object
    Debug.Assert mBasic.BaseName(fl) = "Basic"                    ' Test with File object
    Debug.Assert mBasic.BaseName(ThisWorkbook.Name) = "Basic"     ' Test with a file's name
    Debug.Assert mBasic.BaseName(ThisWorkbook.FullName) = "Basic" ' Test with a file's full name
    Debug.Assert mBasic.BaseName("xxxx") = "xxxx"
    
    '~~ Test unsupported object
    On Error Resume Next
    mBasic.BaseName wb.Worksheets(1)
    On Error GoTo eh
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub
