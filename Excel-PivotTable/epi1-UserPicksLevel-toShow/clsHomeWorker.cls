VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHomeWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==Making the new Class - HomeWorker

'==Global Private variables
    Private mstrHomeSheet As String
    Private mstrRangeWorkOrigin As String
    Private mlngRowsToCopy As Long
Public Function Execute()
    Dim RowNum As Long
    Dim Result As String
    
    mstrHomeSheet = CreatedBlankWorker()
    mstrRangeWorkOrigin = DefineWorkOrigin()
    ' Got the sheet and the cell within it
    
    mlngRowsToCopy = EstablishRowDepth()
    
    Execute = FoundationDataForm()
End Function
Private Sub Class_Initialize()
    Dim item As Variant
    For Each item In ActiveWorkbook.Sheets
        If (item.Name = "WorkerPage") Then
            Call MsgBox("Forget it - You should have deleted THAT page", vbCritical, "CLASS ERROR")
        End If
    Next item
End Sub

Private Function numberToLetter(ByVal code As Long) As String
    numberToLetter = Chr$(64 + code)
End Function

Private Function CreatedBlankWorker() As String
    Dim sheetList() As String
    Dim i As Long, max As Long
    Dim blnUsedPage As Boolean
    blnUsedPage = False
    
    Sheets.Add After:=ActiveSheet
    
    max = Sheets.count
    For i = 1 To max
        If ((Sheets(i).Name <> "EntryPage") And (blnUsedPage = False)) Then
            blnUsedPage = True
            Sheets(i).Name = "BlankPage"
        End If
    Next i
    
    Sheets("BlankPage").Name = "WorkerPage"
    Sheets("WorkerPage").Select
    
    CreatedBlankWorker = ActiveSheet.Name
End Function
Private Function DefineWorkOrigin()
    Sheets(mstrHomeSheet).Select
    Range("K10").Select
    ActiveWorkbook.Names.Add Name:="BeginWorking", RefersToR1C1:= _
        "=WorkerPage!R10C11"
        

    DefineWorkOrigin = "BeginWorking"
End Function

Private Function EstablishRowDepth() As Long
'
' ContinueWorking Macro
'

'
    Dim RowNum As Long
    
    Sheets("EntryPage").Select
    ActiveSheet.Range("DontKnow_HowMany").Select
    RowNum = ActiveCell.Row
    
    ActiveCell.Offset(20000, 0).Select
    Selection.End(xlUp).Select
    EstablishRowDepth = ActiveCell.Row - RowNum
End Function
Private Function FoundationDataForm() As String
' Formally function CopiedBothParts(RowNum)

    Sheets(mstrHomeSheet).Select
    ActiveSheet.Range(mstrRangeWorkOrigin).Select
    ActiveCell.Offset(2, -1).Value = "Data"
    ActiveCell.Offset(2, -1).Select
    Call growBranch("Data")
    
    Sheets(mstrHomeSheet).Select
    ActiveSheet.Range(mstrRangeWorkOrigin).Select
    ActiveCell.Offset(2, 1).Value = "FieldSourseForHolder"
    ActiveCell.Offset(2, 1).Select
    Call growBranch("FieldHolder")
    
    FoundationDataForm = "Done Both"
End Function
Private Sub growBranch(ByVal switchSource As String)
'
' Formally CopyAccross Macro
'

'
    Sheets("EntryPage").Select
    ActiveSheet.Range("DontKnow_HowMany").Select
    Dim lngColSource As Long
    
    Select Case switchSource
        Case "Data":
            lngColSource = ActiveCell.Column + 1
        Case "FieldHolder":
            lngColSource = ActiveCell.Column
        Case Else:
            Call MsgBox("Cant find - " & switchSource, vbCritical, "CASE ERROR growBranch")
    End Select
    
    Dim strC As String
    strC = numberToLetter(lngColSource)
    Dim i As Long
    Dim j As Long
    
    i = ActiveCell.Row
    j = i + mlngRowsToCopy
    
    Dim strData As String
    strData = strC & CStr(i) & ":" & strC & CStr(j)
    
    Range(strData).Select
    Selection.Copy
    Sheets(mstrHomeSheet).Select
    ActiveSheet.Columns("L:L").ColumnWidth = 45 'so we can see it

    ActiveSheet.Paste
End Sub
Private Sub Class_Terminate()
    Application.CutCopyMode = False
    Application.Goto Reference:="BeginWorking"
    Sheets("EntryPage").Select
    Application.Goto Reference:="DontKnow_HowMany"
End Sub
