VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFeedTheConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####
'
Option Explicit
'##### module Global Variables
Private mstrSampleRoot As String
Private mVectbean() As String
Public Function count() As Long
'
    count = UBound(mVectbean) + 1
End Function
Public Function serveLevel(ByVal index As Long) As String
    Dim s As String
    s = mVectbean(index)
    serveLevel = Right$(s, Len(s) - 1)
End Function
Private Sub Class_Initialize()
    Dim Result As String
    'First Start Feeding
    Result = acceptRoot
    'Second Grow the Stalk
    Result = beginGrowth
End Sub
Private Function acceptRoot() As String
    Sheets("EntryPage").Select
    ActiveSheet.Range("DontKnow_HowMany").Select
    mstrSampleRoot = CStr(ActiveCell.Offset(1, 0).Value)
    acceptRoot = "Delivered the sample."
End Function

Private Function beginGrowth() As String
    Dim strBeanstalk As String
    Dim beenRoot As String
    Dim i As Long, iMax As Long
    
    strBeanstalk = mstrSampleRoot & "-"
    
    Dim stalk As Variant
    stalk = Split(strBeanstalk, "-")
    iMax = UBound(stalk)
    ReDim mVectbean(iMax - 1)
    
    Dim strBean As String
    strBean = ""
    For i = iMax - 1 To 0 Step -1
        strBean = strBean & "-" & stalk(iMax - i - 1)
        mVectbean(iMax - i - 1) = strBean
    Next i
    
    beginGrowth = "Beanshoot has grown"

End Function

