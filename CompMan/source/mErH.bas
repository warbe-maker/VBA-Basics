Attribute VB_Name = "mErH"
Option Explicit
' ------------------------------------------------------------------------------
' Standard  Module mErH: Common VBA Error Services with focus on the display of
' ====================== an error message plus some specific debugging support.
'
' Public services:
' ----------------
' Asserted   Only used with regression testing (Regression = True) to avoid the
'            display errors specifically tested. When Regression = False the
'            Asserted service is ignored and any error is displayed.
' BoP        Indicates the Begin of a Procedure and maintains the call stack.
' EoP        Indicates the End of a Procedure and maintains the call stack.
'            Triggers the display of the Execution Trace when the end of the
'            'Entry-Procedure' is reached.
' ErrMsg     Displays an error message either
'            - with the procedure the error had been raised when no BoP
'              statement had ever been executed
'            - passed on the error to the 'Entry Procedure' (which is the first
'              procedure with a BoP statement thereby assembling the 'path
'              to the error' displayed when the 'Entry Procedure' is reached.
' Regression When TRUE the ErrMsg only displays errors which are not regarded
'            'Asserted'.
'
' Uses components:
' ----------------
' fMsg/mMsg   Used only when installed and activated by the Cond. Comp. Arg.
'             `mMsg = 1`
' mTrc/clsTrc Used only by the test environment and only when activated by the
'             Cond. Comp. Arg. `mTrc = 1`
'
' Requires:
' ---------
' Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin, Jan 2024
'
' See https://github.com/warbe-maker/VBA-Error
' See https://warbe-maker.github.io/vba/common/2020/10/02/Comprehensive-Common-VBA-Error-Handler.html
' ------------------------------------------------------------------------------
Public Const CONCAT As String = "||"

' Begin of ShellRun declarations ---------------------------------------------
Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long
' Window Constants
Private Const WIN_NORMAL = 1         'Open Normal
' ShellRun Error Codes
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
' End of ShellRun declarations ---------------------------------------------


Private Const GITHUB_REPO_URL       As String = "https://github.com/warbe-maker/VBA-Error"

Private cllErrPath          As Collection   ' managed by ErrPath... procedures exclusively
Private ProcStack           As Collection   ' stack maintained by BoP (push) and EoP (pop)
Private lSubsequErrNo       As Long         ' possibly different from the initial error number if it changes when passed on
Private vErrsAsserted       As Variant      ' possibly provided with BoTP
Private vErrReply           As Variant
Private vArguments()        As String       ' Arguments passed with the BoP statement
Private cllRecentErrors     As Collection
Private bRegression         As Boolean
Private CurrentProc         As String

Private Property Get EntryProc():                               EntryProc = StackBottom(ProcStack):                         End Property

Private Property Get EntryProcIsKnown() As Boolean:             EntryProcIsKnown = Not StackIsEmpty(ProcStack):             End Property

Private Property Get EntryProcReached() As Boolean:             EntryProcReached = StackBottom(ProcStack) = CurrentProc:    End Property

Public Property Get ErrReply() As Variant:                      ErrReply = vErrReply:                                       End Property

Public Property Get MostRecentError() As Long
    If Not cllRecentErrors Is Nothing Then
        If cllRecentErrors.Count <> 0 Then
            MostRecentError = cllRecentErrors(cllRecentErrors.Count)
        End If
    End If
End Property

Private Property Let MostRecentError(ByVal lErrNo As Long)
    If cllRecentErrors Is Nothing Then Set cllRecentErrors = New Collection
    cllRecentErrors.Add lErrNo
End Property

Public Property Get RecentErrors() As Collection:               Set RecentErrors = cllRecentErrors:                         End Property

Public Property Get Regression() As Boolean:                    Regression = bRegression:                                   End Property

Public Property Let Regression(ByVal et_status As Boolean):     bRegression = et_status:                                    End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If

End Sub

Public Function ShellRun(ByVal sr_string As String, _
                Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Public Sub Asserted(ParamArray botp_errs_asserted() As Variant)
    vErrsAsserted = botp_errs_asserted
End Sub

Public Sub BoP(ByVal b_id As String, _
      Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Trace and push on proc-stack the 'Begin of a Procedure'.
' ------------------------------------------------------------------------------
    If StackIsEmpty(ProcStack) Then
        Set cllRecentErrors = Nothing: Set cllRecentErrors = New Collection
    End If
    StackPush ProcStack, b_id
#If mTrc = 1 Then               ' when mTrc is installed and active
    mTrc.BoP_ErH b_id, b_args
#ElseIf clsTrc = 1 Then         ' when clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP_ErH b_id, b_args
#End If
End Sub

Public Sub EoP(ByVal e_id As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Trace and pop from proc-stack the 'Eegin of a Procedure'.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then       ' when mTrc is installed and active
    mTrc.EoP e_id, e_args
#ElseIf clsTrc = 1 Then ' when clsTrc is installed and active
    Trc.EoP e_id, e_args
#End If
    If StackTop(ProcStack) = e_id Then StackPop ProcStack
End Sub

Private Function ErrArgName(ByVal s As String) As Boolean
    If Right(s, 1) = ":" _
    Or Right(s, 1) = "=" _
    Or Right(s, 2) = ": " _
    Or Right(s, 2) = " :" _
    Or Right(s, 2) = "= " _
    Or Right(s, 2) = " =" _
    Or Right(s, 3) = " : " _
    Or Right(s, 3) = " = " _
    Then ErrArgName = True
End Function

Private Property Let BoPArguments(ByVal v As String)
    If v <> vbNullString Then
        vArguments = Split(v, ";")
    End If
End Property

Private Property Get BoPArguments() As String
' ----------------------------------------------------------------------------
' Returns a string with the arguments which had been passed with the BoP
' statement in the procedure which raised the error. Any argument string
' ending with a ":" or "=" is an arguments name with its value in the
' subsequent Item.
' ----------------------------------------------------------------------------
    Const PROC = "BoPArguments(Get)"
        
    On Error GoTo xt
    If UBound(vArguments) >= LBound(vArguments) Then
        BoPArguments = "(" & Replace(Join(vArguments, ", "), "=, ", " = ") & ")"
    End If

xt: Exit Property

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Private Function ErrBttns(ByVal Bttns As Variant) As Long
' ------------------------------------------------------------------------------
' Returns the number of specified buttons in (bttns).
' ------------------------------------------------------------------------------
    Dim v   As Variant
    Dim s   As String
    Dim i   As Long
    Dim cll As Collection
    
    Select Case TypeName(Bttns)
        Case "Collection"
            Set cll = Bttns
            s = cll(1)
            For i = 2 To cll.Count
                s = s & "," & cll(i)
            Next i
            ErrBttns = ErrBttns(s)
        Case "String"
            i = 0
            For Each v In Split(Bttns, ",")
                If IsNumeric(v) Then
                    Select Case v
                        Case vbOKOnly:                                          i = i + 1
                        Case vbOKCancel, vbYesNo, vbRetryCancel, vbResumeOk:    i = i + 2
                        Case vbAbortRetryIgnore, vbYesNoCancel:                 i = i + 3
                    End Select
                Else
                    Select Case v
                        Case vbNullString, vbLf, vbCr, vbCrLf
                        Case Else:  i = 1
                    End Select
                End If
            Next v
    End Select
    ErrBttns = i
End Function

Private Function ErrHndlrFailed(ByVal err_number As Long, _
                                ByVal err_source As String, _
                                ByVal err_buttons As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the error handling had been invoked with an error number 0
' ----------------------------------------------------------------------------

    If err_number = 0 Then
        MsgBox Prompt:="The error handling has been called with an error number = 0 !" & vbLf & vbLf & _
                       "This indicates that in procedure" & vbLf & _
                       ">>>>> " & err_source & " <<<<<" & vbLf & _
                       "an ""Exit ..."" statement before the call of the error handling is missing!" _
             , Buttons:=vbExclamation _
             , Title:="Exit ... statement missing in " & err_source & "!"
        ErrHndlrFailed = True
        Exit Function
    End If
    
    If IsNumeric(err_buttons) Then
        '~~ When err_buttons is a numeric value, only the VBA MsgBox values for the button argument are supported
        Select Case err_buttons
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore
            Case Else
                MsgBox "When the argument for the ""err_buttons"" parameter is a numeric value  o n l y  the valid VBA.MsgBox vaulues are supported. " & _
                       "For valid values please refer to:" & vbLf & _
                       "https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/msgbox-function" _
                       , vbOKOnly, "Only the valid VBA MsgBox vaulues are supported!"
                ErrHndlrFailed = True
                Exit Function
        End Select
    End If

End Function

Private Function ErrIsAsserted(ByVal err_no As Long) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the error (err_no) is one of the error numbers regarded
' asserted (vErrsAsserted) because thes error numbers had already been
' anticipated with the BoTP (Begin of Test-Procedure) statement - which is a
' variant of the BoP (Begin of Procedure) statement.
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    On Error GoTo xt ' no asserted errors provided
    For Each v In vErrsAsserted
        If v = err_no Then
            ErrIsAsserted = True
            GoTo xt
        End If
    Next v

xt: Exit Function

End Function

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_number As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0, _
              Optional ByVal err_buttons As Variant = vbNullString, _
              Optional ByRef err_reply As Variant) As Variant
' ------------------------------------------------------------------------------
' When the buttons (err_buttons) argument specifies more than one button the
' error message is immediately displayed and the users choice is returned to the
' caller, else when the caller (err_source) is the 'Entry-Procedure' the error
' is displayed with the path to the error, else the error is passed on to the
' Entry Procedure whereby the path to the error is composed/assembled.
' ------------------------------------------------------------------------------
    Const PROC = "ErrMsg"
    
    Static lInitErrNo       As Long
    Static lInitErrLine     As Long
    Static sInitErrSource   As String
    Static sInitErrDscrptn  As String
    Static sInitErrInfo     As String
    Dim sDetails            As String
    Dim sType               As String
    Dim lNo                 As Long
    Dim sLine               As String
        
    If err_number = 0 Then err_number = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_line = 0 Then err_line = Erl
    
    If ErrHndlrFailed(err_number, err_source, err_buttons) Then GoTo xt
    CurrentProc = err_source
    
    ErrMsgMatter err_source:=err_source, err_no:=err_number, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_details:=sDetails
    
    If sInitErrSource = vbNullString Then
        '~~ This is the initial/first execution of the error handler within the error raising procedure.
        sInitErrInfo = sDetails
        lInitErrLine = err_line
        lInitErrNo = err_number
        sInitErrSource = err_source
        sInitErrDscrptn = err_dscrptn
        MostRecentError = err_number
    ElseIf err_number <> lInitErrNo _
        And err_number <> lSubsequErrNo _
        And err_source <> sInitErrSource Then
        '~~ In the rare case when the error number had changed during the process of passing it back up to the 'Entry-Procedure'
        lSubsequErrNo = err_number
        sInitErrInfo = sDetails
    End If
    
    ErrMsgButtons err_buttons
    ErrMsgMatter err_source:=sInitErrSource, _
                 err_no:=lInitErrNo, _
                 err_line:=lInitErrLine, _
                 err_dscrptn:=sInitErrDscrptn, _
                 msg_type:=sType, _
                 msg_line:=sLine, _
                 msg_no:=lNo
    
    '~~ ---------------------------------------------------------------------------
    '~~ The error is passed on to the 'Entry-Procedure' when
    '~~ 1. The 'Entry-Procedure' is know (EntryProcIsKnown) but yet not reached
    '~~ 2. the user has no choice to press another but the Ok button.
    '~~ ---------------------------------------------------------------------------
'    Debug.Print ErrSrc(PROC) & ": " &  "Current Proc              : '" & CurrentProc & "'"
'    Debug.Print ErrSrc(PROC) & ": " &  "EntryProc                 : '" & EntryProc & "'"
'    Debug.Print ErrSrc(PROC) & ": " &  "EntryProcReached          : " & EntryProcReached
'    Debug.Print ErrSrc(PROC) & ": " &  "ErrBttns(err_buttons) = 1 : " & ErrBttns(err_buttons)
    
    If EntryProcIsKnown And CurrentProc <> EntryProc And ErrBttns(err_buttons) = 1 Then
        '~~ When the Entry Procedure is known but yet not reached and there is just one reply
        '~~ button displayed with the error message, the current procedure is added to the
        '~~ 'path to the error' and the error is passed on to the caller.
        ' !! With the Cond. Comp. Arg. 'Debugging = 1' there will be an extra !!
        ' !! button for the error line is to be resumed. In this case the has to be       !!
        ' !! displayed immediately with the error raising procidure. A path to the error  !!
        ' !! will thus only be available to the extent the stack had been maintained by   !!
        ' !! BoP/EoP statements on the way down to the error raising procedure.           !!
        ErrPathAdd err_source
#If mTrc = 1 Then
        mTrc.EoP err_source, "!! " & sType & lNo & " " & sLine & " !!"
#ElseIf clsTrc = 1 Then
        Trc.EoP err_source, "!! " & sType & lNo & " " & sLine & " !!"
#End If
        sInitErrInfo = vbNullString
        Err.Raise err_number, err_source, err_dscrptn ' pass on erro to caller
    
    ElseIf EntryProcIsKnown = False Then
        ErrPathAdd err_source ' add the 'Entry Procedure' as the last one now to the error path
        
        '~~ The display of the error message will be skipped Regression = True and the error number is an Asserted one
        vErrReply = ErrMsgDsply(err_source:=sInitErrSource, err_number:=lInitErrNo, err_dscrptn:=sInitErrDscrptn, err_line:=lInitErrLine, err_buttons:=err_buttons, err_no_asserted:=lInitErrNo)
        ErrMsg = vErrReply
        err_reply = vErrReply
        
        If vErrReply <> vbResume Then ErrPathErase
        StackPop ProcStack
        sInitErrInfo = vbNullString
        sInitErrSource = vbNullString
        sInitErrDscrptn = vbNullString
        lInitErrNo = 0
    
    Else
        '~~ The error is displayed only when the displayed error number is not asserted.
        '~~ Note: It may haver been asserted by the BoTP (begin of Test Procedure service) when just an error condition is tested#
        vErrReply = ErrMsgDsply(err_source:=sInitErrSource, err_number:=lInitErrNo, err_dscrptn:=sInitErrDscrptn, err_line:=lInitErrLine, err_buttons:=err_buttons, err_no_asserted:=lInitErrNo)
        ErrMsg = vErrReply
        err_reply = vErrReply
        sInitErrInfo = vbNullString
        sInitErrSource = vbNullString
        sInitErrDscrptn = vbNullString
        lInitErrNo = 0
        StackErase ProcStack
    End If
    
xt: Exit Function
eh: MsgBox "Error in " & ErrSrc(PROC)
End Function

Private Sub ErrMsgButtons(ByRef err_buttons As Variant)
    Dim cll As New Collection

    Set cll = mMsg.Buttons(cll, vbResumeOk)
    Set err_buttons = cll
End Sub

Private Function ErrMsgDsply(ByVal err_source As String, _
                             ByVal err_number As Long, _
                             ByVal err_dscrptn As String, _
                             ByVal err_line As Long, _
                    Optional ByVal err_no_asserted As Long = 0, _
                    Optional ByVal err_buttons As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------------------
' Displays the error message. The displayed path to the error may be provided as
' the error is passed on to the 'Entry-Procedure' or based on all passed BoP/EoP
' services. In the first case the path to the error may be pretty complete, in
' the second case the extent of detail depends on which (how many) procedures do
' call the BoP/EoP service.
'
' W. Rauschenberger, Berlin, Jun 2023
' ------------------------------------------------------------------------------
    
    Dim sErrPath    As String
    Dim sTitle      As String
    Dim sLine       As String
    Dim sDetails    As String
    Dim sDscrptn    As String
    Dim sAbout      As String
    Dim sSource     As String
    Dim sType       As String
    Dim lNo         As Long
    Dim ErrMsgText  As udtMsg
    Dim sMsg        As String ' The MsgBox Prompt string
    Dim lBttns      As Long
    
#If mTrc = 1 Then
    '~~ When this component is used with clsTrc installed and activated (Cond. Comp.Arg. `clsTrc = 1`
    '~~ the using VB-Project must have `Public Trc As clsTrc` and `Set Trc = New clsTrc` codelines in
    '~~ one of its components! If not the below code line will cause an error.
    mTrc.Pause
#ElseIf clsTrc = 1 Then
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.Pause ' prevent useless timing values by exempting the display and wait time for the reply
#End If
    ErrMsgMatter err_source:=err_source _
               , err_no:=err_number _
               , err_line:=err_line _
               , err_dscrptn:=err_dscrptn _
               , msg_title:=sTitle _
               , msg_line:=sLine _
               , msg_details:=sDetails _
               , msg_source:=sSource _
               , msg_dscrptn:=sDscrptn _
               , msg_info:=sAbout _
               , msg_type:=sType _
               , msg_no:=lNo
    sErrPath = ErrPathErrMsg(sType & lNo & " " & sLine)
    
    '~~ Skip the display when this is a regression test with the error explicitly already asserted
    If bRegression And ErrIsAsserted(err_no_asserted) Then GoTo xt
                       
    '~~ Display the error message via the Common Component procedure mMsg.Dsply
    With ErrMsgText.Section(1)
        With .Label
            .Text = "Error:"
            .FontColor = rgbBlue
        End With
        .Text.Text = sDscrptn
        sMsg = "Error description:" & vbLf & sDscrptn
    End With
'    If BoPArguments <> vbNullString Then
        With ErrMsgText.Section(2)
            With .Label
                .Text = "Source:"
                .FontColor = rgbBlue
            End With
            If BoPArguments <> vbNullString Then
                sSource = sSource & " " & sLine & vbLf & BoPArguments
            Else
                sSource = sSource & " " & sLine
            End If
            .Text.Text = sSource
            sMsg = sMsg & vbLf & vbLf & "Error source:" & vbLf & sSource
        End With
'    End If
    With ErrMsgText.Section(3)
        With .Label
            .Text = "Path:"
            .FontColor = rgbBlue
            .OnClickAction = GITHUB_REPO_URL & "#the-path-to-the-error"
            sMsg = sMsg & vbLf & vbLf & "Error path:" & vbLf
        End With
        If sErrPath <> vbNullString Then
            .Text.Text = sErrPath
            .Text.MonoSpaced = True
            sMsg = sMsg & sErrPath
        Else
            .Text.Text = "A path to the error is not avialable. Click the Label above for more information"
            .Text.MonoSpaced = False
            sMsg = sMsg & "A path to the error is not avialable."
        End If
    End With
    With ErrMsgText.Section(4)
        If sAbout = vbNullString Then
            .Label.Text = vbNullString
            .Text.Text = vbNullString
        Else
            .Label.Text = "About:"
            .Text.Text = sAbout
            .Text.FontSize = 9
            sMsg = sMsg & vbLf & vbLf & "About the error:" & vbLf & sAbout
        End If
        .Label.FontColor = rgbBlue
    End With
    
    With ErrMsgText.Section(5)
        With .Label
            .Text = "Resume Error Line"
            .FontColor = rgbBlue
        End With
        .Text.Text = "Pressing this button and twice F8 leads straight to the code line which raised the error. " & _
                     "(button is displayed because the Cond. Comp. Argument 'Debugging = 1')."
    End With

    Application.EnableEvents = True ' set to TRUE in case FALSE
#If mMsg = 1 Then
    ErrMsgDsply = mMsg.Dsply(d_title:=sTitle _
                           , d_msg:=ErrMsgText _
                           , d_width_max:=60 _
                           , d_label_spec:="R50" _
                           , d_buttons:=mMsg.Buttons(err_buttons))
#Else
    lBttns = vbYesNo
    sMsg = sMsg & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsgDsply = VBA.MsgBox(Title:=sTitle _
                            , Prompt:=sMsg _
                            , Buttons:=lBttns)
#End If

xt:
#If mTrc = 1 Then
    mTrc.EoP err_source, sType & lNo & " " & sLine
    mTrc.Continue ' when the user has replied by pressinbg a button the execution timer continues
#ElseIf clsTrc = 1 Then
    Trc.EoP err_source, sType & lNo & " " & sLine
    Trc.Continue ' when the user has replied by pressinbg a button the execution timer continues
#End If

End Function

Private Sub ErrMsgMatter(ByVal err_source As String, _
                         ByVal err_no As Long, _
                         ByVal err_line As Long, _
                         ByVal err_dscrptn As String, _
                Optional ByRef msg_title As String, _
                Optional ByRef msg_type As String, _
                Optional ByRef msg_line As String, _
                Optional ByRef msg_no As Long, _
                Optional ByRef msg_details As String, _
                Optional ByRef msg_dscrptn As String, _
                Optional ByRef msg_info As String, _
                Optional ByRef msg_source As String)
' ----------------------------------------------------------------------------
' Returns all the matter to build a proper error message:
' - Error Message Title
' - Error Type
' - Error Line
' - Error Number (in case translated back to the original "application-error")
' - Error Details
' ----------------------------------------------------------------------------
                
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)     ' Message error line
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)                ' Message error number
    msg_title = msg_type & msg_no & " in " & err_source & " " & msg_line             ' Message title
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in " & err_source & " (at line " & err_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1)
    msg_source = err_source
    
End Sub

Private Sub ErrPathAdd(ByVal s As String)
    
    If cllErrPath Is Nothing Then Set cllErrPath = New Collection
    If Not ErrPathItemExists(s) Then
'        Debug.Print ErrSrc(PROC) & ": " &  "Add to ErrPath: " & s
        cllErrPath.Add s ' avoid duplicate recording of the same procedure/Item
    End If
End Sub

Private Sub ErrPathErase()
    Set cllErrPath = Nothing
End Sub

Private Function ErrPathErrMsg(ByVal msg_details As String) As String
' ----------------------------------------------------------------------------
' Returns the error path for being displayed in the error message. The path to
' the error has two possible sources:
' - cllErrPath Path collected when the error is passed on to the
'              entry procedure (bottom up approach).
' - ProcStack  maintained by BoP/EoP statements (top down approach).
' Either of the two or both may provide information about the path to the
' error. While the stack requires a BoP/EoP statement with each procedure
' executed, the path collected on the way up will contain only those
' procedures which are passed on the way up to the entry procedure - provided
' at least this one is known. Finally both have its pros and cons. The
' cllErrPath is given the first chance.
' ----------------------------------------------------------------------------
    Dim i   As Long
    Dim j   As Long
    
    ErrPathErrMsg = vbNullString
    If Not ErrPathIsEmpty Then
        '~~ When the error path is not empty this means that the error handling
        '~~ had identified an Entry Procedure (one with BoP/EoP statements.
        '~~ This is by far the best source because it has gathered the path to
        '~~ the error when the error had been passed on up to the Entry Procedure.
        '~~ The downside of this approach: It will not work when the debugging option
        '~~ is used because this option displays the error message immediately with
        '~~ the procedure which raised the error and had an error handling.
        ErrPathErrMsg = " " & cllErrPath(cllErrPath.Count)
        For i = cllErrPath.Count - 1 To 1 Step -1
            j = j + 1
            ErrPathErrMsg = ErrPathErrMsg & vbLf & Space$((j - 1) * 2) & " |_" & cllErrPath(i)
        Next i
    Else
        '~~ When the error path is empty because the error had not been passed on to
        '~~ the Entry Procedure - either because it was unknown or because the debugging
        '~~ option displayed the error immediately with the error raising procedure -
        '~~ the second best chance to get the path to the error is using the stack
        '~~ maintained with all the BoP/EoP statements passed on the way down to the
        '~~ error raising procedure.
        If Not StackIsEmpty(ProcStack) Then
            For i = 1 To ProcStack.Count
                If ErrPathErrMsg <> vbNullString Then
                   ErrPathErrMsg = ErrPathErrMsg & vbLf & Space$((i - 1) * 2) & "|_" & ProcStack(i)
                Else
                   ErrPathErrMsg = ProcStack(i)
                End If
            Next i
        End If
    End If
    If ErrPathErrMsg <> vbNullString Then
        ErrPathErrMsg = ErrPathErrMsg & ": " & msg_details
    End If

End Function

Private Function ErrPathIsEmpty() As Boolean
    ErrPathIsEmpty = cllErrPath Is Nothing
    If Not ErrPathIsEmpty Then ErrPathIsEmpty = cllErrPath.Count = 0
End Function

Private Function ErrPathItemExists(ByVal s As String) As Boolean

    Dim v As Variant
    
    For Each v In cllErrPath
        If InStr(v & " ", s & " ") <> 0 Then
            ErrPathItemExists = True
            Exit Function
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mErH." & sProc
End Function

Private Function StackBottom(ByVal s_stck As Collection) As String
    If Not StackIsEmpty(s_stck) Then StackBottom = ProcStack(1)
End Function

Private Sub StackErase(ByRef s_stck As Collection)
    If Not s_stck Is Nothing Then Set s_stck = Nothing
    Set s_stck = New Collection
End Sub

Public Function StackIsEmpty(ByVal s_stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Common Stack Empty check service. Returns True when either there is no stack
' (s_stck Is Nothing) or when the stack is empty (items count is 0).
' ----------------------------------------------------------------------------
    StackIsEmpty = s_stck Is Nothing
    If Not StackIsEmpty Then StackIsEmpty = s_stck.Count = 0
End Function

Public Function StackPop(ByVal s_stck As Collection, _
                Optional ByVal s_id As Variant = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last Item pushed on the stack (s_stck)
' and removes the Item from the stack. When the stack (s_stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StackIsEmpty(s_stck) Then GoTo xt
    
    If VarType(s_id) = vbObject Then
        If Not StackTop(ProcStack) Is s_id Then GoTo xt
    Else
        If Not s_id = vbNullString Then
            If StackTop(ProcStack) <> s_id Then GoTo xt
        End If
    End If
    
    If IsObject(s_stck(s_stck.Count)) _
    Then Set StackPop = s_stck(s_stck.Count) _
    Else StackPop = s_stck(s_stck.Count)
    s_stck.Remove s_stck.Count

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub StackPush(ByRef s_stck As Collection, _
                     ByVal s_item As Variant)
' ----------------------------------------------------------------------------
' Common Stack Push service. Pushes (adds) an Item (s_item) to the stack
' (s_stck). When the provided stack (s_stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "StckPush"
    
    On Error GoTo eh
    If s_stck Is Nothing Then Set s_stck = New Collection
    s_stck.Add s_item

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function StackTop(ByVal s_stck As Collection) As Variant
    If Not StackIsEmpty(s_stck) Then
        If VarType(s_stck.Count) = vbObject _
        Then Set StackTop = s_stck(s_stck.Count) _
        Else StackTop = s_stck(s_stck.Count)
    End If
End Function

