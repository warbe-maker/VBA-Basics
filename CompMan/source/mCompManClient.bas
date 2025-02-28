Attribute VB_Name = "mCompManClient"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mCompManClient: InterfaceCompMan services (Export Changed
' =============================== Components, Update Outdated CommonComponents,
' or the Synchronize VB-Projects). To be imported into any VBProject for
' making use of one or more services.
'
' W. Rauschenberger, Berlin Dec 2024
' See https://github.com/warbe-maker/VB-Components-Management
' ----------------------------------------------------------------------------
Public Const COMPMAN_DEVLP              As String = "CompMan.xlsb"
Public Const SRVC_EXPORT_ALL            As String = "ExportAll"
Public Const SRVC_EXPORT_ALL_DSPLY      As String = "Export All Components"
Public Const SRVC_EXPORT_CHANGED        As String = "ExportChangedComponents"
Public Const SRVC_EXPORT_CHANGED_DSPLY  As String = "Export Changed Components"
Public Const SRVC_RELEASE_PENDING       As String = "ReleaseService"
Public Const SRVC_RELEASE_PENDING_DSPLY As String = "Release pending changes"
Public Const SRVC_SYNCHRONIZE           As String = "SynchronizeVBProjects"
Public Const SRVC_SYNCHRONIZE_DSPLY     As String = "Synchronize VB-Projects"
Public Const SRVC_UPDATE_OUTDATED       As String = "UpdateOutdatedCommonComponents"
Public Const SRVC_UPDATE_OUTDATED_DSPLY As String = "Update Outdated Common Components"

Private Const COMPMAN_ADDIN             As String = "CompMan.xlam"
Private Const vbResume                  As Long = 6 ' return value (equates to vbYes)
Private Busy                            As Boolean ' prevent parallel execution of a service
Private sEventsLvl                      As String
Private bWbkExecChange                  As Boolean

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As LongPtr, ByRef riid As UUID, ByRef ppvObject As Object) As LongPtr

Type UUID 'GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Const OBJID_NATIVEOM As LongPtr = &HFFFFFFF0

Private Property Let DisplayedServiceStatus(ByVal s As String)
    With Application
        .StatusBar = vbNullString
        .StatusBar = s
    End With
End Property

Public Function IsAddinInstance() As Boolean
    IsAddinInstance = ThisWorkbook.Name = COMPMAN_ADDIN
End Function

Public Function IsDevInstance() As Boolean
    IsDevInstance = ThisWorkbook.Name = mCompManClient.COMPMAN_DEVLP
End Function

Public Function ServiceName(Optional ByVal s As String) As String
    Select Case s
        Case SRVC_EXPORT_CHANGED:   ServiceName = SRVC_EXPORT_CHANGED_DSPLY
        Case SRVC_UPDATE_OUTDATED:  ServiceName = SRVC_UPDATE_OUTDATED_DSPLY
    End Select
End Function

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error number never conflicts
' with VB runtime error. Thr function returns a given positive number
' (app_err_no) with the vbObjectError added - which turns it to negative. When
' the provided number is negative it returns the original positive "application"
' error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Sub CompManService(ByVal c_service_proc As String, _
                 Optional ByVal c_hosted_common_components As String = vbNullString, _
                 Optional ByVal c_public_procedure_copies As String = vbNullString)
' ----------------------------------------------------------------------------
' Execution of the CompMan service (c_service_proc) preferably via the "CompMan
' Development Instance" as the servicing Workbook. Only when not available the
' "CompMan AddIn Instance" (COMPMAN_ADDIN) becomes the servicing
' Workbook - which maynot be open either or open but paused.
' Note: c_unused is for backwards compatibility only
' ----------------------------------------------------------------------------
    Const PROC = "CompManService"
    
    On Error GoTo eh
    Dim sServicingWbkName   As String
           
    Progress p_service_name:=ServiceName(c_service_proc) _
           , p_serviced_wbk_name:=ThisWorkbook.Name
    On Error Resume Next
    If ActiveWindow.Caption <> ActiveWorkbook.Name Then
        Debug.Print "service denied because the caption is not identical with the active Workbook's name"
        Exit Sub ' Any restored, e.g. (Version ..) is ignored
    End If
    
    sEventsLvl = vbNullString
    mCompManClient.Events ErrSrc(PROC) & "." & c_service_proc, False
    If IsAddinInstance Then
        Application.StatusBar = "None of CompMan's services is applicable for CompMan's Add-in instance!"
        GoTo xt
    End If

    '~~ Avoid any trouble caused by DoEvents used throughout the execution of any service
    '~~ when a service is already currently busy. This may be the case when Workbook-Save
    '~~ is clicked twice.
    If Busy Then
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_service_info:="Terminated because a previous task is still busy!"
        Exit Sub
    End If
    Busy = True
    
    sServicingWbkName = ServicingWbkName(c_service_proc)
                                   
    If sServicingWbkName <> vbNullString Then
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_by_servicing_wbk_name:=sServicingWbkName
        Application.Run sServicingWbkName & "!mCompMan." & c_service_proc, ThisWorkbook, c_hosted_common_components, c_public_procedure_copies
    Else
        Progress p_service_name:=ServiceName(c_service_proc) _
               , p_serviced_wbk_name:=ThisWorkbook.Name _
               , p_service_info:="Workbook saved (CompMan-Service not applicable)"
    End If
    
xt: Busy = False
    mCompManClient.Events ErrSrc(PROC) & "." & c_service_proc, True
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
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
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManClient." & sProc
End Function

Public Sub Events(ByVal e_src As String, _
                  ByVal e_b As Boolean, _
         Optional ByVal e_reset As Boolean = False)
' ------------------------------------------------------------------------------
' Follow-Up (trace) of Application.EnableEvents False/True - proves consistency.
' Recognizes the execution chang from the initiating Workbook to the service
' executing Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "Events"
    
    On Error GoTo eh
    Static sLastExecWrkbk   As String
    
    If e_reset Then
        sEventsLvl = vbNullString
        bWbkExecChange = False
        sLastExecWrkbk = vbNullString
        Exit Sub
    End If
    
    If Not e_b Then
        EventsApp False
        Debug.Print ErrSrc(PROC) & ": " & sEventsLvl & ">> " & ThisWorkbook.Name & "." & e_src & " (Application.EnableEvents = False)"
        If sLastExecWrkbk <> vbNullString And ThisWorkbook.Name <> sLastExecWrkbk And Not bWbkExecChange Then
            bWbkExecChange = True
            sEventsLvl = sEventsLvl & "   "
        End If
        sEventsLvl = sEventsLvl & "   "
        sLastExecWrkbk = ThisWorkbook.Name
    Else
        sEventsLvl = Left(sEventsLvl, Len(sEventsLvl) - 3)
        sLastExecWrkbk = ThisWorkbook.Name
        EventsApp True
        Debug.Print ErrSrc(PROC) & ": " & sEventsLvl & "<< " & ThisWorkbook.Name & "." & e_src & " (Application.EnableEvents = True)"
    End If

    If sEventsLvl = vbNullString Then
        Events vbNullString, False, True
    End If
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub EventsApp(a_events As Boolean)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "EventsApp"

    On Error GoTo eh
#If Win64 Then
    Dim hWndMain As LongPtr
#Else
    Dim hWndMain As Long
#End If
    Dim appThis As Application
    Dim appNext As Application
    
    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)
    Do While hWndMain <> 0
        Set appNext = GetExcelObjectFromHwnd(hWndMain)
        If Not appNext Is Nothing Then
            If Not appNext Is appThis Then
                Set appThis = appNext
                If appThis.EnableEvents <> a_events Then
                    appThis.EnableEvents = a_events
                End If
            End If
        End If
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

#If Win64 Then
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As LongPtr) As Application
#Else
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As Long) As Application
#End If

#If Win64 Then
    Dim hWndDesk As LongPtr
    Dim hwnd As LongPtr
#Else
    Dim hWndDesk As Long
    Dim hwnd As Long
#End If
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Dim sText   As String
    Dim lRet    As Long
    Dim iid     As UUID
    Dim ob      As Object
    
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then
        hwnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Do While hwnd <> 0
            sText = String$(100, Chr$(0))
            lRet = CLng(GetClassName(hwnd, sText, 100))
            If Left$(sText, lRet) = "EXCEL7" Then
                Call IIDFromString(StrPtr(IID_IDispatch), iid)
                If AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, iid, ob) = 0 Then 'S_OK
                    Set GetExcelObjectFromHwnd = ob.Application
                    GoTo xt
                End If
            End If
            hwnd = FindWindowEx(hWndDesk, hwnd, vbNullString, vbNullString)
        Loop
        
    End If
    
xt:
End Function

Private Function IsString(ByVal v As Variant, _
                 Optional ByVal vbnullstring_is_a_string = False) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is neither an object nor numeric.
' ----------------------------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = v
    If Err.Number = 0 Then
        If Not IsNumeric(v) Then
            If (s = vbNullString And vbnullstring_is_a_string) _
            Or s <> vbNullString _
            Then IsString = True
        End If
    End If
End Function

Public Sub Progress(ByVal p_service_name As String, _
           Optional ByVal p_serviced_wbk_name As String = vbNullString, _
           Optional ByVal p_by_servicing_wbk_name As String = vbNullString, _
           Optional ByVal p_progress_figures As Boolean = False, _
           Optional ByVal p_service_op As String = vbNullString, _
           Optional ByVal p_no_comps_serviced As Long = 0, _
           Optional ByVal p_no_comps_outdated As Long = 0, _
           Optional ByVal p_no_comps_total As Long = 0, _
           Optional ByVal p_no_comps_skipped As Long = 0, _
           Optional ByVal p_service_info As String = vbNullString, _
           Optional ByVal p_sequ_no As Long = 0)
' --------------------------------------------------------------------------
' Progress display in the Application.StatusBar for CompMan services.
' Form: <service> (by <by>) for <serviced>: <n> of <m> <o> [<c> [, <c>] ..]
' <n> = Components the service has been provided for (p_items_serviced)
' <m> = Total number of components being serviced
' <o> = The performed operation
' <c> = Components processed, e.g. exported
' Whereby the progress is indicated in two ways: an increasing number of
' dots for the items collected for being serviced and a decreasing number
' of dots indication the items already serviced.
' --------------------------------------------------------------------------
    Const PROC                  As String = "Progress"
    Const SRVC_PROGRESS_SCHEME  As String = "<srvc> <by> <serviced>: <n> of <m> <op> <info> <dots>"
    
    On Error GoTo eh
    Dim sMsg    As String
    Dim lDots   As Long
    Dim sFormat As String
    
    sMsg = Replace(SRVC_PROGRESS_SCHEME, "<srvc>", p_service_name)
    sMsg = Replace(sMsg, "<serviced>", "for " & p_serviced_wbk_name)
    
    If p_by_servicing_wbk_name <> vbNullString _
    Then sMsg = Replace(sMsg, "<by>", "(by " & p_by_servicing_wbk_name & ")") _
    Else sMsg = Replace(sMsg, "<by>", vbNullString)
    
    If p_no_comps_total < 100 Then sFormat = "#0" Else sFormat = "##0"
    
    lDots = p_no_comps_total - p_no_comps_skipped - p_no_comps_serviced
    If lDots >= 0 Then
        sMsg = Replace(sMsg, "<dots>", String(lDots, "."))
    Else
        sMsg = Replace(sMsg, "<dots>", vbNullString)
    End If
    
    If p_service_op <> vbNullString Then
        If p_sequ_no = 0 _
        Then sMsg = Replace(sMsg, "<op>", p_service_op) _
        Else sMsg = Replace(sMsg, "<op>", p_service_op & " " & p_sequ_no & " ")
    Else
        sMsg = Replace(sMsg, "<op>", "Service initiating")
    End If
    
    If p_progress_figures Then
        sMsg = Replace(sMsg, "<n>", Format(p_no_comps_serviced, sFormat))
        If p_no_comps_outdated <> 0 _
        Then sMsg = Replace(sMsg, "<m>", Format(p_no_comps_outdated, sFormat)) _
        Else sMsg = Replace(sMsg, "<m>", Format(p_no_comps_total, sFormat))
    Else
        sMsg = Replace(sMsg, "<n> of <m>", vbNullString)
    End If
    
    sMsg = Replace(sMsg, "<info>", p_service_info)
    sMsg = Replace(sMsg, "  ", " ")
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 250) & " ..."
    With Application
        .StatusBar = vbNullString
        .StatusBar = Trim(sMsg)
    End With
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ServicingWbkName(ByVal s_service_proc As String) As String
' ----------------------------------------------------------------------------
' Returns the name of the Workbook providing the requested service which may
' be a vbNullString when the service neither can be provided by an open
' CompMan development instance Workbook nor by an available CompMan Add-in
' instance.
' Notes: - When the requested service is not "update" an available development
'          instance is given priority over an also available Add-in instance.
'        - When the requested service is "update" and the serviced Workbook
'          is the development instance the service is only available when the
'          Add-in instance is avaialble.
'        - Even when a servicing Workbook (the Add-in and or the development
'          instance is available, CompMan may still not be configured
'          correctly!
' Uses: mCompMan.RunTest
' ----------------------------------------------------------------------------
    Const PROC = "ServicingWbkName"
    
    Dim ServicedByAddinResult           As Long
    Dim ServicedByWrkbkResult           As Long
    Dim ServiceAvailableByAddIn         As Boolean
    Dim ServiceAvailableByDevInstance   As Boolean
    Dim ResultRequiredAddinNotAvailable As Long
    Dim ResultConfigInvalid             As Long
    Dim ResultOutsideCfgFolder          As Long
    Dim ResultRequiredDevInstncNotOpen  As Long
    Dim ResultServiceByAddinIsPaused    As Long
    
    ResultConfigInvalid = AppErr(1)             ' Configuration for the service is invalid
    ResultOutsideCfgFolder = AppErr(2)          ' Outside the for the service required folder
    ResultRequiredAddinNotAvailable = AppErr(3) ' Required Addin for DevInstance update paused or not open
    ResultRequiredDevInstncNotOpen = AppErr(4)  '
    ResultServiceByAddinIsPaused = AppErr(5)
    
    '~~ Availability check CompMan Add-in
    On Error Resume Next
    ServicedByAddinResult = Application.Run(COMPMAN_ADDIN & "!mCompMan.RunTest", s_service_proc, ThisWorkbook)
    ServiceAvailableByAddIn = Err.Number = 0
    '~~ Availability check CompMan Workbook
    On Error Resume Next
    ServicedByWrkbkResult = Application.Run(COMPMAN_DEVLP & "!mCompMan.RunTest", s_service_proc, ThisWorkbook)
    ServiceAvailableByDevInstance = Err.Number = 0
    
    Select Case True
        '~~ Display/indicate why the service cannot be provided
        Case ServicedByAddinResult = ResultServiceByAddinIsPaused _
         And ResultRequiredDevInstncNotOpen = AppErr(4)
            DisplayedServiceStatus = "The service has been denied because the AddIn is paused!"
        Case ServicedByWrkbkResult = ResultConfigInvalid
            Select Case s_service_proc
                Case SRVC_UPDATE_OUTDATED:  DisplayedServiceStatus = "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
                Case SRVC_EXPORT_CHANGED:   DisplayedServiceStatus = "The enabled/requested'" & SRVC_EXPORT_CHANGED_DSPLY & "' service had been denied due to an invalid or missing configuration (see Config Worksheet)!"
            End Select
        Case ServicedByWrkbkResult = ResultOutsideCfgFolder
            Progress p_service_name:=ServiceName(s_service_proc) _
                   , p_serviced_wbk_name:=ThisWorkbook.Name _
                   , p_service_info:="Service not applicable"
            Select Case s_service_proc
                Case SRVC_UPDATE_OUTDATED:  Debug.Print ErrSrc(PROC) & ": " & "The enabled/requested '" & SRVC_EXPORT_CHANGED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
                Case SRVC_EXPORT_CHANGED
                    Debug.Print ErrSrc(PROC) & ": " & "The enabled/requested '" & SRVC_UPDATE_OUTDATED_DSPLY & "' service had silently been denied! (Workbook has not been opened from within the configured 'Dev-and-Test-Folder')"
            End Select
        Case ServicedByWrkbkResult = ResultRequiredAddinNotAvailable
            DisplayedServiceStatus = "The required Add-in is not available for the 'Update' service for the Development-Instance!"
        Case ServicedByWrkbkResult = ResultRequiredDevInstncNotOpen
            DisplayedServiceStatus = mCompManClient.COMPMAN_DEVLP & " is the Workbook reqired for the " & SRVC_SYNCHRONIZE & " but it is not open!"
        
        '~~ When neither of the above is True the servicing Workbook instance is decided
        Case IsDevInstance _
         And s_service_proc = SRVC_UPDATE_OUTDATED _
         And ServiceAvailableByAddIn
            '~~ The development instance's outdated Common Components can only be updated
            '~~ by an available (open and not paused) Addin instance
                                                ServicingWbkName = COMPMAN_ADDIN
        Case Not IsDevInstance _
         And ServiceAvailableByDevInstance
                                                ServicingWbkName = COMPMAN_DEVLP
        Case Not IsDevInstance _
         And Not ServiceAvailableByDevInstance _
         And ServiceAvailableByAddIn
                                                ServicingWbkName = COMPMAN_ADDIN
        Case Not ServiceAvailableByDevInstance _
         And ServiceAvailableByAddIn
                                                ServicingWbkName = COMPMAN_ADDIN
        Case ServiceAvailableByDevInstance _
         And Not ServiceAvailableByAddIn
                                                ServicingWbkName = COMPMAN_DEVLP
        Case ServiceAvailableByDevInstance _
         And ServiceAvailableByAddIn
                                                ServicingWbkName = COMPMAN_DEVLP
        Case Else
            '~~ Silent service denial
            Debug.Print ErrSrc(PROC) & ": " & "CompMan services are not available because neither CompMan.xlsb nor the CompMan Add-in is open!"
    End Select
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

