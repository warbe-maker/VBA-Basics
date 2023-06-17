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

Private Const HOSTED_RAWS = "mBasic"

Private Sub Workbook_Open()
    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Open service 'UpdateRawClones':
    '~~ Executed by the Addin *) or via the development instance when open
    '~~ *) automatically available only when referenced by the VB-Project
    mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
    '~~ ------------------------------------------------------------------
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Before_Save service 'ExportChangedComponents':
    '~~ Executed by the CompMan Addin when available or via CompMan's
    '~~ development instance Workbook when open. Neither open the service
    '~~ ends without error.
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
    '~~ ------------------------------------------------------------------
End Sub
