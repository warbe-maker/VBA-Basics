VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const HOSTED_COMMON_COMPONENTS = "mBasic"
Private Const HOSTED_COMMON_PUBLIC_PROCS = "Align,AlignCntr,AlignLeft,AlignRght,AppErr,AppIsInstalled,Arry,ArryAsRnge,ArryBase,ArryCompare,ArryDiffers,ArryDims,ArryIsAllocated,ArryItems,ArryRemoveItems,ArryTrimm,BaseName,CleanTrim,DelayedAction,Dict,ErrMsg,KeySort,README,"

Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_COMMON_COMPONENTS, HOSTED_COMMON_PUBLIC_PROCS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_COMMON_COMPONENTS, HOSTED_COMMON_PUBLIC_PROCS
End Sub
