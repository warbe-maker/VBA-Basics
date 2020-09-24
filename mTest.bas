Attribute VB_Name = "mTest"
Option Private Module
Option Explicit
Option Compare Text
' -----------------------------------------------------
' Standard Module mBasic
'          Basic declarations, procedures, methods and
'          functions coomon im most VBProjects.
'
' Please note:
' Errors raised by the tested procedures cannot be
' asserted since they are not passed on to the calling
' /entry procedure. This would require the Common
' Standard Module mErrHndlr which is not used with this
' module by intention.
'
' W. Rauschenberger, Berlin Feb 2020
' -----------------------------------------------------

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

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
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    Debug.Assert AppErr(Err.Number) = 1
    
    a = aTest
    ' Missing parameter
    On Error Resume Next
    mBasic.ArrayRemoveItems a
    Debug.Assert AppErr(Err.Number) = 3
    
    ' Element out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=8
    Debug.Assert AppErr(Err.Number) = 4
    
    ' Index out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Index:=7
    Debug.Assert AppErr(Err.Number) = 5
    
    ' Element plus number of elements out of boundary
    On Error Resume Next
    mBasic.ArrayRemoveItems a, Element:=7, NoOfElements:=2
    Debug.Assert AppErr(Err.Number) = 6

exit_proc:
    Exit Sub

on_error:
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
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
    Debug.Assert AppErr(Err.Number) = 1
    On Error GoTo on_error
    
exit_proc:
    Exit Sub
    
on_error:
    mBasic.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_ErrMsg()
' -------------------------
' Common error message test
' -------------------------
Const PROC = "Test_ErrMsg"
    mBasic.ErrMsg lErrNo:=1, sErrSrc:=ErrSrc(PROC), sErrDesc:="The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog." & DCONCAT & "Optional possible additional info about the error", sErrLine:="0"
End Sub

Public Sub Test_Msg_1_Reply_Button_1_Message_Paragraph()
Dim sTitle      As String
Dim sMsg1       As String
Dim vReply1     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant
Dim vReplied    As Variant

    sTitle = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over ..."
    
    sMsg1 = "Test Message with 1 reply button. Reply with <Ok>!"
    
    vReplies = vbOKOnly
    vReply = vbOK
    
    vReplied = Msg( _
                   sTitle:=sTitle, _
                   sMsgText:=sMsg1, _
                   bFixed:=False, _
                   vReplies:=vReplies _
                  )
    Debug.Assert vReplied = vReply

End Sub

Public Sub Test_Msg_2_Reply_Buttons_1_Message_Paragraph()
Dim sTitle      As String
Dim sMsg1       As String
Dim vReply1     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant
Dim vReplied    As Variant

    sTitle = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over ..."
    
    sMsg1 = "Test Message with 2 standard MsgBox reply buttons. Reply with <Yes>!"
    
    vReplies = vbYesNo
    vReply = vbYes
    
    Debug.Assert Msg( _
                 sTitle:=sTitle, _
                 sMsgText:=sMsg1, _
                 bFixed:=False, _
                 vReplies:=vReplies _
                    ) = vReply
End Sub

Public Sub Test_Msg_3_Reply_Buttons_1_Message_Paragraph()
Dim sTitle      As String
Dim sMsg1       As String
Dim vReply1     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant
Dim vReplied    As Variant

    sTitle = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over ..."
    
    sMsg1 = "Test Message with 3 standard MsgBox reply buttons. Reply with <No>!"
    
    vReplies = vbYesNoCancel
    vReply = vbNo
    
    Debug.Assert Msg( _
                 sTitle:=sTitle, _
                 sMsgText:=sMsg1, _
                 bFixed:=False, _
                 vReplies:=vReplies _
                    ) = vReply
End Sub

Public Sub Test_Msg_4_Reply_Buttons_1_Message_Pragraph()
Dim sTitle      As String
Dim sLabel1     As String
Dim sLabel2     As String
Dim sLabel3     As String
Dim sText1      As String
Dim sText2      As String
Dim sText3      As String
Dim vReply1     As Variant
Dim vReply2     As Variant
Dim vReply3     As Variant
Dim vReply4     As Variant
Dim vReply5     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant
Dim vReplied    As Variant

    sTitle = "mBasic.Msg provides a message analogous to the MsgBox with some ""improvements"""
    
    sText1 = "1. The title is never trunctated (see 4. width adjustment)" & vbLf & _
             "2. The message text may optionally be displayed in a fixed font" & vbLf & _
             "   to support a message text with indented lines (like these here or the error-path in an error message)" & vbLf & _
             "3. There may be up to 5 reply buttons either like the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
             "   or any text with any number of lines (all buttons are adjusted accordingly) and the" & vbLf & _
             "   reply value corresponds with the button's content, i.e. either vbOk, vbYes, ..." & vbLf & _
             "   or the displayed text" & vbLf & _
             "4. The final message box width is dertermined by:" & vbLf & _
             "   - the width of the title (never truncated)" & vbLf & _
             "   - the longest fixed font message text line" & vbLf & _
             "   - the required space for the displayed reply buttons"
    sText2 = "Reply this test with <Display Execution Trace>!"
    
    vReply1 = "Update Target" & vbLf & "with Source"
    vReply2 = "Update Source" & vbLf & "with Target"
    vReply3 = "Display" & vbLf & "Execution Trace"
    vReply4 = "Ignore"
    vReply5 = vbNullString
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply3
    
    vReplied = Msg( _
                   sTitle:=sTitle, _
                   sMsgText:=sText1 & vbLf & vbLf & sText2, _
                   bFixed:=True, _
                   vReplies:=vReplies _
                  )
    Debug.Assert CStr(vReplied) = CStr(vReply)
    
End Sub

Public Sub Test_Msg_4_Reply_Buttons_3_Message_Paragraphs()
Dim sTitle      As String
Dim sLabel1     As String
Dim sLabel2     As String
Dim sLabel3     As String
Dim sText1      As String
Dim sText2      As String
Dim sText3      As String
Dim vReply1     As Variant
Dim vReply2     As Variant
Dim vReply3     As Variant
Dim vReply4     As Variant
Dim vReply5     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant

    sTitle = "mBasic.mMsg3 works pretty much like MsgBox but with significant enhancements (see below)"
    
    sLabel1 = "General"
    sText1 = "- The title will never be truncated" & vbLf & _
             "- There are up to 3 message paragraphs, each with an optional label/header" & vbLf & _
             "- Each message paragraph may be in a proportional or fixed font (like these two)" & vbLf & _
             "  supporting indented text like this one - or the display of the error-path in an error message" & vbLf & _
             "- There are up to 5 reply buttons. 3 work exactly like with the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
             "  but all may as well contain any string and the reply value corresponds with the clicked reply button"
    
    sLabel2 = "Window width adjustment"
    sText2 = "The message window width considers:" & vbLf & _
             "- the title width (never truncated anymore)" & vbLf & _
             "- the maximum length of any fixed font message block" & vbLf & _
             "- total width of the displayed reply buttons" & vbLf & _
             "- specified minimum window width"
    
    sLabel3 = vbNullString
    sText3 = "Reply this test with <Reply 4>"
    
    vReply1 = "Reply 1"
    vReply2 = "Reply 2"
    vReply3 = "Reply lines" & vbLf & "determine the" & vbLf & "button height"
    vReply4 = "Reply 4"
    vReply5 = vbNullString
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply4
    
    Debug.Assert Msg3( _
                      sTitle:=sTitle, _
                      sText1:=sText1, sLabel1:=sLabel1, bFixed1:=True, _
                      sText2:=sText2, sLabel2:=sLabel2, bFixed2:=False, _
                      sText3:=sText3, sLabel3:=sLabel3, bFixed3:=False, _
                      vReplies:=vReplies _
                     ) = vReply
End Sub

Public Sub Test_Msg_5_Reply_Buttons_3_Message_Paragraphs()
Dim sTitle      As String
Dim sLabel1     As String
Dim sLabel2     As String
Dim sLabel3     As String
Dim sText1      As String
Dim sText2      As String
Dim sText3      As String
Dim vReply1     As Variant
Dim vReply2     As Variant
Dim vReply3     As Variant
Dim vReply4     As Variant
Dim vReply5     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant

    sTitle = "mBasic.mMsg3 works pretty much like MsgBox but with significant enhancements (see below)"
    
    sLabel1 = "General"
    sText1 = "- The title will never be truncated" & vbLf & _
             "- There are up to 3 message paragraphs, each with an optional label/header" & vbLf & _
             "- Each message paragraph may be in a proportional or fixed font (like these two)" & vbLf & _
             "  supporting indented text like this one - or the display of the error-path in an error message" & vbLf & _
             "- There are up to 5 reply buttons. 3 work exactly like with the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
             "  but all may as well contain any string and the reply value corresponds with the clicked reply button"
    
    sLabel2 = "Window width adjustment"
    sText2 = "The message window width considers:" & vbLf & _
             "- the title width (never truncated anymore)" & vbLf & _
             "- the maximum length of any fixed font message block" & vbLf & _
             "- total width of the displayed reply buttons" & vbLf & _
             "- specified minimum window width"
    
    sLabel3 = vbNullString
    sText3 = "Reply this test with <Reply lines determine the button height>"
    
    vReply1 = "Reply 1"
    vReply2 = "Reply 2"
    vReply3 = "Reply lines" & vbLf & "determine the" & vbLf & "button height"
    vReply4 = "Reply 4"
    vReply5 = "Reply 5"
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply3
    
    Debug.Assert Msg3( _
                      sTitle:=sTitle, _
                      sText1:=sText1, sLabel1:=sLabel1, bFixed1:=True, _
                      sText2:=sText2, sLabel2:=sLabel2, bFixed2:=False, _
                      sText3:=sText3, sLabel3:=sLabel3, bFixed3:=False, _
                      vReplies:=vReplies _
                     ) = vReply
End Sub

