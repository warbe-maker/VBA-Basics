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
        Debug.Print i & ". : " & VBA.Environ$(i) & """"
        If Err.Number <> 0 Then Exit For
    Next i
End Sub

Public Sub Regression()
    Const PROC = "Regression"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    mTest.Test_01_ArrayCompare
    mTest.Test_02_ArrayRemoveItems
    mTest.Test_03_ArrayToRange
    mTest.Test_04_ArrayTrimm
    mTest.Test_05_BaseName
    mErH.EoP ErrSrc(PROC)

xt: Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
    
End Sub

Public Sub Test_01_ArrayCompare()
    Const PROC  As String = "Test_03_ArrayToRange"
    
    On Error GoTo eh
    Dim a1      As Variant
    Dim a2      As Variant
    Dim dctDiff As Variant
    Dim v       As Variant
    
    mErH.BoP ErrSrc(PROC)
    
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
    Debug.Assert dctDiff.Count = 1
    For Each v In dctDiff
        Debug.Print "Test 4: Item/line " & v & vbLf & dctDiff(v)
    Next v
        
    '~~ Test 5: The arrays first elements are different, empty elements are ignored
    a1 = Split(",2,3,4,5,6,7", ",")     ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",")    ' Test array
    Set dctDiff = mBasic.ArrayCompare(ac_a1:=a1 _
                                    , ac_a2:=a2 _
                                     )
    Debug.Assert dctDiff.Count = 1
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
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_02_1_ArrayRemoveItems_Function()
    Const PROC  As String = "Test_02-1_ArrayRemoveItems_Function"
    
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
        Case mErH.DebugOptResumeErrorLine:  Stop: Resume
        Case mErH.DebugOptResumeNext:       Resume Next
    End Select
End Sub

Public Sub Test_02_2_ArrayRemoveItems_Error_Conditions()
    Const PROC  As String = "Test_02_2_ArrayRemoveItems_Error_Conditions"
    
    On Error GoTo eh
    Dim aTest   As Variant
    Dim a       As Variant
    
    mErH.BoP ErrSrc(PROC)

    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
        
    ' Not an array
    Set a = Nothing
    On Error Resume Next
    mBasic.ArrayRemoveItems a, 2
    Debug.Assert mErH.AppErr(Err.Number) = 1
    
    a = aTest
    ' Missing parameter
    On Error Resume Next
    mBasic.ArrayRemoveItems a
    Debug.Assert mErH.AppErr(Err.Number) = 3
    
    ' Element out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=8
    Debug.Assert mErH.AppErr(Err.Number) = 4
    
    ' Index out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Index:=7
    Debug.Assert mErH.AppErr(Err.Number) = 5
    
    ' Element plus number of elements out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=7, NoOfElements:=2
    Debug.Assert mErH.AppErr(Err.Number) = 6

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
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
    Test_02_1_ArrayRemoveItems_Function
    Test_02_2_ArrayRemoveItems_Error_Conditions
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_03_ArrayToRange()
    Const PROC  As String = "Test_03_ArrayToRange"
    
    On Error GoTo eh
    Dim a       As Variant
    Dim aTest   As Variant
    
    mErH.BoP ErrSrc(PROC)
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    a = aTest

    wsBasic.UsedRange.ClearContents
    mBasic.ArrayToRange a, wsBasic.celArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasic.rngArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasic.celArrayToRangeTarget
    mBasic.ArrayToRange a, wsBasic.rngArrayToRangeTarget

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_04_ArrayTrimm()
    Const PROC = "Test_04_ArrayTrimm"
    
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
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_05_BaseName()
' -----------------------------------------------------
' Please note:
' The common error handler (module mErrHndlr) is used
' in order to allow an "unattended regression test"
' because the ErrHndlr passes on the error number to
' the (this) entry procedure
' -----------------------------------------------------
    Const PROC  As String = "Test_05_BaseName"
    
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
    Debug.Assert mBasic.BaseName(ThisWorkbook.name) = "Basic"     ' Test with a file's name
    Debug.Assert mBasic.BaseName(ThisWorkbook.FullName) = "Basic" ' Test with a file's full name
    Debug.Assert mBasic.BaseName("xxxx") = "xxxx"
    
    '~~ Test unsupported object
    On Error Resume Next
    mBasic.BaseName wb.Worksheets(1)
    On Error GoTo eh
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_06_NbSpd()
    Debug.Assert Replace(Nbspd("     Ab c      "), Chr$(160), " ") = "  A b c  "
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
