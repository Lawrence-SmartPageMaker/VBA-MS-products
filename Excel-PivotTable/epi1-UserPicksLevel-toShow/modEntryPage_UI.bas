Attribute VB_Name = "modEntryPage_UI"
'============================================================
'   module : modEntryPage_UI
' file name: modEntryPage_UI.bas
' --------------------------------
' The user interface for the Entry Page
'============================================================
Public Sub btnReset_Click()
    Dim objWorker As New clsHomeWorker
    objWorker.Execute
    Set objWorker = Nothing 'explicit Let it GO
    
    Dim objFormer As New clsPropagate
    Call objFormer.PrepareForm
    
End Sub

Public Sub btnConfig_Click()
    Load ConfigLevel
    ConfigLevel.Show
End Sub

Public Sub btnProcesss_Click()
    Dim objTransmit As New clsPropagate
    objTransmit.GrowTheOutput
End Sub
