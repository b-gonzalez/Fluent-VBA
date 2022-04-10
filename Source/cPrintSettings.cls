VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPrintSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPrintSettings

Private pPrintTestsToImmediate As Boolean
Private pPrintTotalTests As Boolean
Private pPrintTestsToSheet As Boolean
Private pTestName As String
Private pTestCollection As Collection

Public Property Let PrintTestsToImmediate(value As Boolean)
    pPrintTestsToImmediate = value
End Property

Public Property Get PrintTestsToImmediate() As Boolean
    PrintTestsToImmediate = pPrintTestsToImmediate
End Property

Public Property Let PrintTestsToSheet(value As Boolean)
    pPrintTestsToSheet = value
End Property

Public Property Get PrintTestsToSheet() As Boolean
    PrintTestsToSheet = pPrintTestsToSheet
End Property

Public Property Let PrintTotalTests(value As Boolean)
    pPrintTotalTests = value
End Property

Public Property Get PrintTotalTests() As Boolean
    PrintTotalTests = pPrintTotalTests
End Property

Private Property Let IPrintSettings_testname(value As String)
    pTestName = value
End Property

Private Property Get IPrintSettings_testname() As String
    IPrintSettings_testname = pTestName
End Property

Private Property Get IPrintSettings_TestCollection() As Collection
    If pTestCollection Is Nothing Then
        Set pTestCollection = New Collection
    End If
    Set IPrintSettings_TestCollection = pTestCollection
End Property

Private Sub IPrintSettings_checkIfPrint(ByVal value As Boolean, testMessage As String, Optional negateValue As Boolean = False)
'    Dim arr() As Variant
'
'    arr = getArrayFromCollection(pTestCollection)
'
'    If negateValue Then
'        value = Not value
'    End If
'
'    If pPrintTestsToImmediate Then
'        Call PrintToImmediate(value, arr)
'    ElseIf pPrintTestsToSheet Then
'        Call PrintToSheet(value, arr)
'    End If
End Sub

Public Sub PrintToImmediate()
Dim arr() As Variant
    Dim i As Long
'    Dim Result As String
'    Dim TestPassed As Boolean
'
'    Result = ""
'    TestPassed = value
'
'    If TestPassed Then
'        Result = "Passed"
'        If pTestName <> Empty Then
'            Debug.Print Result & " - " & pTestName & testMessage
'        Else
'            Debug.Print Result & " - " & testMessage
'        End If
'    Else
'        Result = "Failed"
'        If pTestName <> Empty Then
'            Debug.Print Result & " - " & pTestName & testMessage
'        Else
'            Debug.Print Result & " - " & testMessage
'        End If
'    End If

arr = getArrayFromCollection(pTestCollection)

For i = LBound(arr, 1) To UBound(arr, 1)
    Debug.Print arr(i, 1)
Next i

End Sub

Public Sub PrintToSheet()
    Dim lo As ListObject
    Dim testResultCol As ListColumn
    Dim testNameCol As ListColumn
    Dim arr() As Variant
    
    arr = getArrayFromCollection(pTestCollection)

    Set lo = getTableFromWorkbook
    
    Set testResultCol = lo.ListColumns("Test result")
    
    Set testNameCol = lo.ListColumns("Test name")
    
    Call addResultsIconSets(testResultCol)
    
    Call addValueConditColors(testResultCol, "-1", "0")
    
    Call addValueConditColors(testNameCol, "Passed", "Failed")
    
    lo.ListRows.Add.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).value = arr

End Sub

Private Function getTableFromWorkbook() As ListObject
    Dim xl As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    If Not TypeOf Application Is Excel.Application Then
        Set xl = New Excel.Application
    Else
        Set xl = Application
    End If
    
    xl.Visible = True
    
    Set wb = xl.Workbooks.Add
    Set ws = wb.ActiveSheet
    ws.Range("A1").value = "Test result"
    ws.Range("B1").value = "Test name"
    Set lo = ws.ListObjects.Add(xlSrcRange, Range("$A$1:$B$1"), , xlYes)
    lo.Name = "Table1"
        
    Set getTableFromWorkbook = lo
End Function

Private Sub addResultsIconSets(lc As ListColumn)
    lc.Range.Select
    
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Symbols2)
    End With
    Selection.FormatConditions(1).IconCriteria(1).Icon = xlIconYellowExclamation
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .value = -1
        .Operator = 7
        .Icon = xlIconGreenCheck
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .value = 0
        .Operator = 7
        .Icon = xlIconRedCross
    End With
End Sub

Private Sub addValueConditColors(lc As ListColumn, passedMsg As String, failedMsg As String)
    lc.Range.Select
    
    Selection.FormatConditions.Add Type:=xlTextString, String:=failedMsg, _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:=passedMsg, _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Private Function getArrayFromCollection(col As Collection) As Variant()
    Dim arr() As Variant
    Dim i As Long
    Dim temp As Variant
    
    ReDim arr(col.Count - 1, 1)
    
    For i = 1 To col.Count
       temp = Split(col(i), "=")
       arr(i - 1, 0) = temp(0)
        If temp(0) = -1 Then
            arr(i - 1, 1) = "PASSED - " & temp(1)
        ElseIf temp(0) = 0 Then
            arr(i - 1, 1) = "FAILED - " & temp(1)
        End If
    Next i
    
    getArrayFromCollection = arr
End Function
