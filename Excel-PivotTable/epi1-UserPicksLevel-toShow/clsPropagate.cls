VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPropagate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==Making the new Class - Propagate

'==Global Private variables
    Private mlngRowsToCopy As Long
    Private mlngMaxWidth As Long
    Private mstrStartSample As String
Public Sub PrepareForm()
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    ActiveCell.Value = 1
    ' needs its one method !
    Call BeginTheSample
    Call ContinueFormingFirstRow
    Call BeginTheSample
    Call DetermineColumnEnds
    Call ExtendPrepareCopy
    Call ExtendPreparePaste
    Call FormulaMaker
    Call GrowFinalProduction ' Should be enough ?

End Sub
Private Sub GrowBeforeOutput()
    Dim i As Long
    Application.Goto Reference:="BeginWorking"
    ActiveCell.Offset(2, 1).Range("A1").Select
    i = ActiveCell.Row
    ActiveCell.Offset(20000, 0).Range("A1").Select
    Selection.End(xlUp).Select
    mlngRowsToCopy = ActiveCell.Row - i
End Sub
Public Sub GrowTheOutput()
    Call GrowBeforeOutput
    Call GrowFinalOutput
    
'==== This puts its hat on!
    Range("L12").Select
    Selection.Copy
    Range("I12").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
End Sub

Private Sub BeginTheSample()
    ActiveSheet.Range(mstrStartSample).Select
End Sub
Private Sub ContinueFormingFirstRow()
    ActiveCell.FormulaR1C1 = "=FIND(""-"",+RC[-2],1)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    ActiveCell.FormulaR1C1 = "=FIND(""-"",RC[-3],1+RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    ActiveCell.FormulaR1C1 = "=FIND(""-"",RC[-4],1+RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    ActiveCell.FormulaR1C1 = "=FIND(""-"",RC[-5],1+RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    ActiveCell.FormulaR1C1 = "=FIND(""-"",RC[-6],1+RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    ActiveCell.FormulaR1C1 = "=FIND(""-"",RC[-7],1+RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select

End Sub
Private Sub Class_Initialize()
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    ' # # DONT DO ActiveCell.Value = 1
    ActiveCell.Offset(3, 3).Select
    mstrStartSample = ActiveCell.Address
End Sub
Sub DetermineColumnEnds()
    Dim i As Long, j As Long, k As Long
  
    k = -1
    For i = 1 To 7
        j = CLng(ActiveCell.Offset(0, i).Value)
        If j = xlErrValue Then
            ActiveCell.Offset(0, i).Value = "E*R*R*O*R"
            k = IIf(k < 0, i, k) 'first bad value
        End If
    Next i
    mlngMaxWidth = k
End Sub
Private Sub ExtendPrepareCopy()
'
' ExtendPrepareCopy Macro
'

'
    Dim widest As String, strCopySource As String
    widest = Chr$(mlngMaxWidth + 64)
    strCopySource = "A1:" & widest & "1"
    ActiveCell.Range(strCopySource).Select
    Selection.Copy
End Sub

Private Sub ExtendPreparePaste()
    Dim i As Long
    Application.Goto Reference:="BeginWorking"
    ActiveCell.Offset(2, 1).Range("A1").Select
    i = ActiveCell.Row
    ActiveCell.Offset(20000, 0).Range("A1").Select
    Selection.End(xlUp).Select
    mlngRowsToCopy = ActiveCell.Row - i
    Call BeginTheSample
    
    Dim strPasteDest As String
    strPasteDest = "A1:A" & CStr(mlngRowsToCopy)
    ActiveCell.Range(strPasteDest).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    
    'MsgBox ("Finised pasting I hope")
End Sub
Private Sub FormulaMaker()
    Dim s As String
    s = FormulaRun()
    
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    ActiveCell.Offset(3, 2).Select
    
    ActiveCell.Formula = Trim$(s)
End Sub
Private Function FormulaRun() As String
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    
    Dim i As Long
    Dim baseCol As Long, fixedRow As Long
    
    fixedRow = ActiveCell.Row + 3
    baseCol = 1 + 64 + ActiveCell.Column
    
    Dim strElephant(12) As String
    Dim a As String, b As String
    Dim c As String, d As String
    
    Dim DoubledQuote As String
    DoubledQuote = Chr$(34) & Chr$(34)
        
    
    For i = 0 To mlngMaxWidth
        a = "IF(+BeginWorking=" & CStr(i + 1) & ","
        b = "LEFT( " & Chr$(baseCol) & CStr(fixedRow) & ","
        c = Chr$(2 + baseCol + i) & CStr(fixedRow) & ")"
        d = "," & DoubledQuote & ")"
        strElephant(i) = a & b & c & d
    Next i
    
    a = "IF(+BeginWorking=" & CStr(mlngMaxWidth + 1) & ","
    b = Chr$(baseCol) & CStr(fixedRow) & ","
    d = DoubledQuote & ")"
    
    strElephant(mlngMaxWidth) = a & b & d
    
    a = "=" & strElephant(0)
    For i = 1 To mlngMaxWidth
        a = a & "&" & strElephant(i)
    Next i
    
    FormulaRun = a
End Function
Private Sub GrowFinalOutput()
    Dim s As String
    
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    ActiveCell.Offset(3, 2).Select 'This is First Member of Production ?
    
    Dim strProduced As String
    strProduced = "A1:A" & CStr(mlngRowsToCopy)
    ActiveCell.Range(strProduced).Select
    Selection.Copy
    
    Sheets("WorkerPage").Select
    Application.Range("BeginWorking").Select
    ActiveCell.Offset(3, -2).Select 'This is First member of Output ?
    
    s = "GrowTheRest is at: " & ActiveCell.Address
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
End Sub
Private Sub GrowFinalProduction()
    Dim s As String
    s = "GrowTheRest is at: " & ActiveCell.Address



    'MsgBox (s)
'
'
    Dim strC As String
    strC = Chr$(64 + ActiveCell.Column)
    Dim i As Long
    i = ActiveCell.Row

    strRangeToPaste = strC & CStr(i) & ":" & strC & CStr(i + mlngRowsToCopy - 1)
    Selection.Copy
    Range(strRangeToPaste).Select

    ActiveSheet.Paste
    Application.CutCopyMode = False


End Sub

