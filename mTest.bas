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
        If Err.Number <> 0 Then Exit For
    Next i
End Sub

Public Sub Test_ArrayCompare()
Const PROC  As String = "Test_ArrayToRange"
Dim a1      As Variant
Dim a2      As Variant

    On Error GoTo on_error
    
    '~~ Test 1: One element is different
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,x,5,6,7", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "3: " & DGT & "4" & DLT & DCONCAT & DGT & "x" & DLT
    
    '~~ Test 2: The first array has less elements
    a1 = Split("1,2,3,4,5,6", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "6: " & DGT & "" & DLT & DCONCAT & DGT & "7" & DLT
    
    '~~ Test 3: The second array has less elements
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "6: " & DGT & "7" & DLT & DCONCAT & DGT & "" & DLT
    
    '~~ Test 4: The arrays first elements are different (element in second array is empty)
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split(",2,3,4,5,6,7", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "0: " & DGT & "1" & DLT & DCONCAT & DGT & "" & DLT
    
    '~~ Test 5: The arrays first elements are different (element in first array is empty)
    a1 = Split(",2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,4,5,6,7", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "0: " & DGT & "" & DLT & DCONCAT & DGT & "1" & DLT
    
    '~~ Test 5: The second array has additional inserted elements
    a1 = Split("1,2,3,4,5,6,7", ",") ' Test array
    a2 = Split("1,2,3,x,y,z,4,5,6,7", ",") ' Test array
    Debug.Assert Join(mBasic.ArrayCompare(a1, a2), ",") = "3: " & DGT & "4" & DLT & DCONCAT & DGT & "x" & DLT & ",4: " & DGT & "5" & DLT & DCONCAT & DGT & "y" & DLT & ",5: " & DGT & "6" & DLT & DCONCAT & DGT & "z" & DLT & ",6: " & DGT & "7" & DLT & DCONCAT & DGT & "4" & DLT & ",7: " & DGT & "" & DLT & DCONCAT & DGT & "5" & DLT & ",8: " & DGT & "" & DLT & DCONCAT & DGT & "6" & DLT & ",9: " & DGT & "" & DLT & DCONCAT & DGT & "7" & DLT
    
exit_proc:
    Exit Sub
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mBasic.ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_ArrayRemoveItems()
' ---------------------------------
' Whitebox and regression test.
' Global error handling is used to
' monitor error condition tests.
' ---------------------------------
Const PROC  As String = "Test_ArrayRemoveItems"

    On Error GoTo on_error
    
    Test_ArrayRemoveItems_Function
    Test_ArrayRemoveItems_Error_Conditions
    Test_ArrayRemoveItems_Error_Display
    
exit_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mBasic.ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_DctDiffers()
' -------------------------------------------
' Precondition: DctAdd is tested
' -------------------------------------------
    Const PROC = "Test_DctDiffers"
    Dim dct1 As Dictionary
    Dim dct2 As Dictionary
    Dim vbc  As VBComponent
    
    BoP ErrSrc(PROC)
    Test_DctAdd
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dct1, dctnewkey:=vbc, dctnewitem:=vbc.Name, dctmode:=dct_ascending_bykey_casesensitive
    Next vbc
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dct2, dctnewkey:=vbc, dctnewitem:=vbc.Name, dctmode:=dct_ascending_bykey_casesensitive
    Next vbc
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    
    '~~ Test: Differs in keys
    Debug.Assert Not DctDiffers(dct1, dct2)
    dct1.Remove vbc
    Debug.Assert DctDiffers(dct1, dct2)
    EoP ErrSrc(PROC)
    
End Sub
Private Sub Test_ArrayRemoveItems_Error_Conditions()
Const PROC  As String = "Test_ArrayRemoveItems_Error_Conditions"
Dim aTest   As Variant
Dim a       As Variant

    On Error GoTo on_error
    
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
        
    ' Not an array
    Set a = Nothing
    On Error Resume Next
    mBasic.ArrayRemoveItems a, 2
    Debug.Assert mErrHndlr.AppErr(Err.Number) = 1
    
    a = aTest
    ' Missing parameter
    On Error Resume Next
    mBasic.ArrayRemoveItems a
    Debug.Assert mErrHndlr.AppErr(Err.Number) = 3
    
    ' Element out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=8
    Debug.Assert mErrHndlr.AppErr(Err.Number) = 4
    
    ' Index out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Index:=7
    Debug.Assert mErrHndlr.AppErr(Err.Number) = 5
    
    ' Element plus number of elements out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=7, NoOfElements:=2
    Debug.Assert mErrHndlr.AppErr(Err.Number) = 6

exit_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_ArrayRemoveItems_Error_Display()
Const PROC  As String = "Test_ArrayRemoveItems_Error_Display"
Dim aTest   As Variant
Dim a       As Variant
Dim v       As Variant
Dim i       As Long

    On Error GoTo on_error
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    
    ReDim a(5, 2 To 8):    i = LBound(a, 2)
    For Each v In aTest
        a(1, i) = v:  i = i + 1
    Next v
    mBasic.ArrayRemoveItems a, Element:=3

exit_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_ArrayRemoveItems_Function()
Const PROC  As String = "Test_ArrayRemoveItems_Function"
Dim aTest   As Variant
Dim a       As Variant
Dim v       As Variant
Dim i       As Long

    On Error GoTo on_error
    
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

exit_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub Test_ArrayToRange()
Const PROC  As String = "Test_ArrayToRange"
Dim a       As Variant
Dim aTest   As Variant

    On Error GoTo on_error
    
    aTest = Split("1,2,3,4,5,6,7", ",") ' Test array
    a = aTest

    wsBasicTest.UsedRange.ClearContents
    mBasic.ArrayToRange a, wsBasicTest.celArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasicTest.rngArrayToRangeTarget, True
    mBasic.ArrayToRange a, wsBasicTest.celArrayToRangeTarget
    mBasic.ArrayToRange a, wsBasicTest.rngArrayToRangeTarget

    Exit Sub
    
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Sub Test_ArrayTrimm()
Dim a       As Variant
Dim aTest   As Variant

    On Error GoTo on_error
    
    aTest = Split(" , ,1,2,3,4,5,6,7, , , ", ",") ' Test array
    a = aTest
    mBasic.ArrayTrimm a
    Debug.Assert Join(a, ",") = "1,2,3,4,5,6,7"
    
    a = Split(" , , , , ", ",")
    mBasic.ArrayTrimm a
    Debug.Assert mBasic.ArrayIsAllocated(a) = False
    
    Exit Sub
    
on_error:
    Stop: Resume
End Sub

Public Sub Test_BaseName()
' -----------------------------------------------------
' Please note:
' The common error handler (module mErrHndlr) is used
' in order to allow an "unattended regression test"
' because the ErrHndlr passes on the error number to
' the (this) entry procedure
' -----------------------------------------------------
Const PROC  As String = "Test_BaseName"
Dim wb      As Workbook
Dim fl      As File

    On Error GoTo on_error
    
    '~~ Prepare for tests
    Set wb = ThisWorkbook
    With New FileSystemObject
        Set fl = .GetFile(wb.FullName)
    End With
    
    Debug.Assert mBasic.BaseName(wb) = "Basic"                    ' Test with Workbook object
    Debug.Assert mBasic.BaseName(fl) = "Basic"                    ' Test with File object
    Debug.Assert mBasic.BaseName(ThisWorkbook.Name) = "Basic"     ' Test with a file's name
    Debug.Assert mBasic.BaseName(ThisWorkbook.FullName) = "Basic" ' Test with a file's full name
    Debug.Assert mBasic.BaseName("xxxx") = "xxxx"
    
    '~~ Test unsupported object
    On Error Resume Next
    mBasic.BaseName wb.Worksheets(1)
    On Error GoTo on_error
    
exit_proc:
    Exit Sub
    
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Public Sub Test_DctAdd()
    
    Const PROC = "Test_DctAdd"
    BoP ErrSrc(PROC)
    
    Test_DctAdd_KeyIsValue
    Test_DctAdd_KeyIsObjectWithNameProperty
    Test_DctAdd_ItemIsObjectWithNameProperty
    Test_DctAdd_InsertKeyBefore
    Test_DctAdd_InsertKeyAfter
    Test_DctAdd_InsertItemBefore
    Test_DctAdd_InsertItemAfter
    
    EoP ErrSrc(PROC)

End Sub

Private Sub Test_DctAddResult(ByVal dct As Dictionary)
    
    Dim v As Variant
    
    Debug.Print "---------------------------"
    For Each v In dct
        If IsNumeric(v) Or TypeName(v) = "String" _
        Then Debug.Print v _
        Else Debug.Print v.Name
    Next v
    Debug.Print "---------------------------"
    
End Sub

Private Sub Test_DctAdd_ItemIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_ItemIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dctTest, dctnewkey:=vbc.Name, dctnewitem:=vbc, dctmode:=dct_ascending_byitem_casesensitive
    Next vbc
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsBasicTest"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd dct:=dctTest, dctnewkey:=vbc.Name, dctnewitem:=vbc, dctmode:=dct_ascending_byitem_casesensitive
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1).Name = "wsBasicTest"
    EoP ErrSrc(PROC)
        
End Sub

Private Sub Test_DctAdd_InsertKeyAfter()
    
    Const PROC = "Test_DctAdd_InsertKeyAfter"
    Dim vbc1 As VBComponent
    Dim vbc2 As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStackItem"
    Set vbc1 = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc2 = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc2
    Debug.Assert dctTest.Count = 7
    
    '~~ Test
    DctAdd dctTest, vbc2, vbc1.Name, dct_addafter_key_caseignored, vbc1
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_InsertItemAfter()
    
    Const PROC = "Test_DctAdd_InsertItemAfter"
    Dim vbc1 As VBComponent
    Dim vbc2 As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "clsCallStack"
    Debug.Assert dctTest.Keys()(1) = "clsCallStackItem"
    Set vbc1 = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc2 = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc2.Name
    Debug.Assert dctTest.Count = 7
    
    '~~ Test
    DctAdd dctTest, vbc2.Name, vbc2, dct_addafter_item_caseignored, vbc1
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Items()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_InsertItemBefore()
    
    Const PROC = "Test_DctAdd_InsertItemBefore"
    Dim vbc1 As VBComponent
    Dim vbc2 As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_ItemIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0) = "clsCallStack"
    Debug.Assert dctTest.Keys()(1) = "clsCallStackItem"
    Set vbc1 = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc2 = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc1.Name
    Debug.Assert dctTest.Count = 7
    
    '~~ Test
    DctAdd dctTest, vbc1.Name, vbc1, dct_addbefore_item_caseignored, vbc2
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Items()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_InsertKeyBefore()
    
    Const PROC = "Test_DctAdd_InsertKeyBefore"
    Dim vbc1 As VBComponent
    Dim vbc2 As VBComponent
    
    BoP ErrSrc(PROC)
    
    '~~ Preparation
    Test_DctAdd_KeyIsObjectWithNameProperty
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStack"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStackItem"
    Set vbc1 = ThisWorkbook.VBProject.VBComponents("clsCallStackItem")
    Set vbc2 = ThisWorkbook.VBProject.VBComponents("clsCallStack")
    dctTest.Remove vbc1
    Debug.Assert dctTest.Count = 7
    
    '~~ Test
    DctAdd dctTest, vbc1, vbc1.Name, dct_addbefore_key_caseignored, vbc2
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Keys()(0).Name = "clsCallStackItem"
    Debug.Assert dctTest.Keys()(1).Name = "clsCallStack"
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctAdd_KeyIsObjectWithNameProperty()
' ----------------------------------------------------
' Added items with a key which is an object.
' The order by key uses the object's name property.
' ----------------------------------------------------
    Const PROC = "Test_DctAdd_KeyIsObjectWithNameProperty"
    Dim i   As Long
    Dim vbc As VBComponent
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        DctAdd dct:=dctTest, dctnewkey:=vbc, dctnewitem:=vbc.Name, dctmode:=dct_ascending_bykey_casesensitive
    Next vbc
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0) = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsBasicTest"
    
    '~~ Add an already existing key = update the item
    Set vbc = ThisWorkbook.VBProject.VBComponents("mTest")
    DctAdd dct:=dctTest, dctnewkey:=vbc, dctnewitem:=vbc.Name, dctmode:=dct_ascending_bykey_casesensitive
    Debug.Assert dctTest.Count = 8
    Debug.Assert dctTest.Items()(0) = "clsCallStack"
    Debug.Assert dctTest.Items()(dctTest.Count - 1) = "wsBasicTest"
    EoP ErrSrc(PROC)
    
'    Test_DctAddResult dctTest
    
End Sub

Private Sub Test_DctAdd_KeyIsValue()
' -----------------------------------------------
' Note: Since a 100% reverse key order added in mode ascending
' is the worst case regarding performance this test sorts 100 items
' with 50% already in sequence and the other 50% to be inserted.
' -----------------------------------------------
    Const PROC = "Test_DctAdd_KeyIsValue"
    Dim i   As Long
    
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    For i = 1 To 9 Step 2
        DctAdd dct:=dctTest, dctnewkey:=i, dctnewitem:=ThisWorkbook, dctmode:=dct_ascending_bykey_casesensitive
    Next i
    For i = 10 To 2 Step -2
        DctAdd dct:=dctTest, dctnewkey:=i, dctnewitem:=ThisWorkbook, dctmode:=dct_ascending_bykey_casesensitive
    Next i
    
    '~~ Add an already existing key, ignored when the item is neither numeric nor a string
    DctAdd dct:=dctTest, dctnewkey:=5, dctnewitem:=ThisWorkbook, dctmode:=dct_ascending_bykey_casesensitive
    
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_DctNumKey()
    Const PROC = "Test_DctNumKey"
    BoP ErrSrc(PROC)
    Set dctTest = Nothing
    
    DctAdd dctTest, 2, 5, dct_ascending_bykey_caseignored
    DctAdd dctTest, 5, 2, dct_ascending_bykey_caseignored
    DctAdd dctTest, 3, 4, dct_ascending_bykey_caseignored
    
    Debug.Assert dctTest.Count = 3
    Debug.Assert dctTest.Keys()(0) = 2
    Debug.Assert dctTest.Keys()(dctTest.Count - 1) = 5
    
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_Msg_1_Reply()
' ---------------------------
' ---------------------------
Dim sMsg1       As String
Dim sMsg2       As String
Dim sTitle      As String
Dim vReplies    As Variant
Dim vReply      As Variant

    sMsg1 = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog."
    sMsg2 = "Click Display Execution Trace!"
    sTitle = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over ..."
    vReplies = vbOKOnly
    vReply = vbOK
    
    Debug.Assert Msg( _
                 sTitle:=sTitle, _
                 sMsgText:="Fixed: " & sMsg1 & vbLf & sMsg1 & vbLf & sMsg1 & vbLf & sMsg1 & vbLf & sMsg1 & vbLf & vbLf & sMsg2 & vbLf & vbLf & "Form width is dertermined by the 4 reply buttons with maximized width.", _
                 bFixed:=True, _
                 sReplies:=vReplies _
                    ) = vReply
End Sub

Public Sub Test_Msg_5_Replies()
' -----------------------------
' -----------------------------
Dim sMsg1       As String
Dim sMsg2       As String
Dim sMsg3       As String
Dim sTitle      As String
Dim vReplies    As Variant
Dim vReply      As Variant
Dim sLabel1     As String
Dim sLabel2     As String
Dim sLabel3     As String
Dim vReply1     As Variant
Dim vReply2     As Variant
Dim vReply3     As Variant
Dim vReply4     As Variant
Dim vReply5     As Variant

    sTitle = "mBasic.mMsg and .mMsg3 guarantee that the title is never truncated!"
    
    sLabel1 = vbNullString
    sMsg1 = "mBasic.mMsg displays 1, msg3 displays up to 3 text strings/blocks" & vbLf & _
            "- either proportional or fixed" & vbLf & _
            "- each with an optional label/title above"
    
    sLabel2 = "optional label/titel for any of the 3 text strings"
    sMsg2 = "The message window width is adjusted to" & vbLf & _
            "- the required width for the title " & vbLf & _
            "- the required width for the longest fixed font text (like this one)." & vbLf & _
            "- the required width for the (max 5) displayed reply buttons (determined the width of this test)" & vbLf & _
            "- the specified minimum window width" & vbLf & _
            "Proportional text strings are adjusted to the final window width"
    
    sLabel3 = "For this test reply with <Reply button 5> !!!"
    sMsg3 = "By the way: The returned reply is equal to the content of the clicked reply button which may be any of the MsgBox values vbOk, vbCancel, ... or any free text string."
    
    vReply1 = "Reply button 1"
    vReply2 = "Reply button 2"
    vReply3 = "All reply buttons" & vbLf & "are adjusted" & vbLf & "to the biggest"
    vReply4 = "Reply button 4"
    vReply5 = "Reply button 5"
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    
    vReply = vReply5
    
    Debug.Assert mBasic.Msg3( _
                 sTitle:=sTitle, _
                 sLabel1:=sLabel1, sText1:=sMsg1, bFixed1:=False, _
                 sLabel2:=sLabel2, sText2:=sMsg2, bFixed2:=True, _
                 sLabel3:=sLabel3, sText3:=sMsg3, bFixed3:=False, _
                 sReplies:=vReplies _
                    ) = vReply
End Sub

