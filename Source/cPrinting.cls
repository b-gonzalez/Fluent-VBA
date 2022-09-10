VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPrinting

Private pPrintTotalTests As Boolean
Private pTestName As String
Private pTestCollection As Collection
Private pTestCount As Long
Private pCategory As String
Private iPrint As IPrinting
Private pPassedMessage As String
Private pFailedMessage As String

Private Sub Class_Initialize()
    pPassedMessage = "PASSED"
    pFailedMessage = "FAILED"
End Sub

Private Property Get IPrinting_TestCollection() As Collection
    If pTestCollection Is Nothing Then
        Set pTestCollection = New Collection
    End If
    Set IPrinting_TestCollection = pTestCollection
End Property

Private Property Let IPrinting_testCount(value As Long)
    pTestCount = value
End Property

Private Property Get IPrinting_testCount() As Long
    IPrinting_testCount = pTestCount
End Property

Public Property Let Category(value As String)
    pCategory = value
End Property

Public Property Get Category() As String
    Category = pCategory
End Property

Public Property Let TestName(value As String)
    pTestName = value
End Property

Public Property Get TestName() As String
    TestName = pTestName
End Property

Public Property Let PassedMessage(value As String)
    pPassedMessage = value
End Property

Public Property Get PassedMessage() As String
    PassedMessage = pPassedMessage
End Property

Public Property Let FailedMessage(value As String)
    pFailedMessage = value
End Property

Public Property Get FailedMessage() As String
    FailedMessage = pFailedMessage
End Property

Public Sub PrintToImmediate()
    Dim arr() As Variant
    Dim i As Long
    Dim TestName As Long
    
    If pCategory = "" Then
        TestName = 1
    Else
        TestName = 2
    End If
    
    arr = getArrayFromCollection(pTestCollection)
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If pCategory <> "" Then
            Debug.Print pCategory & " - " & arr(i, TestName)
        Else
            Debug.Print arr(i, TestName)
        End If
    Next i
End Sub

Public Sub PrintToSheet()
    Dim lo As ListObject
    Dim testResultCol As ListColumn
    Dim testNameCol As ListColumn
    Dim categoryCol As ListColumn
    Dim arr() As Variant
    
    arr = getArrayFromCollection(pTestCollection)

    Set lo = getTableFromWorkbook
    
    Set testResultCol = lo.ListColumns("Test result")
    
    Set testNameCol = lo.ListColumns("Test name")
    
    Call addResultsIconSets(testResultCol)
    
    If pCategory <> "" Then
        Set categoryCol = lo.ListColumns("Category")
        lo.ListRows.Add.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).value = arr
        Columns("C:C").EntireColumn.AutoFit
        Call addCategoryStyle(testResultCol, categoryCol)
        'Call addValueConditColorsNew(categoryCol)
    Else
        lo.ListRows.Add.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).value = arr
        Columns("B:B").EntireColumn.AutoFit
    End If

    Call addValueConditColorsNew(testResultCol, "-1", "0")
    
    Call addValueConditColorsNew(testNameCol, "*" & pPassedMessage & "*", "*" & pFailedMessage & "*")

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

    If pCategory <> "" Then
        ws.Range("A1").value = "Category"
        ws.Range("B1").value = "Test result"
        ws.Range("C1").value = "Test name"
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("$A$1:$C$1"), , xlYes)
    Else
        ws.Range("A1").value = "Test result"
        ws.Range("B1").value = "Test name"
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("$A$1:$B$1"), , xlYes)
    End If
    
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

Private Function getArrayFromCollection(Col As Collection) As Variant()
    Dim Category As Long
    Dim TestResult As Long
    Dim TestName As Long
    Dim arr() As Variant
    Dim i As Long
    Dim temp As Variant
    
    If pCategory = "" Then
        TestResult = 0
        TestName = 1
        ReDim arr(Col.Count - 1, 1)
    Else
        Category = 0
        TestResult = 1
        TestName = 2
        ReDim arr(Col.Count - 1, 2)
    End If
    
    For i = 1 To Col.Count
       temp = Split(Col(i), "=")
       arr(i - 1, Category) = temp(Category)
       arr(i - 1, TestResult) = temp(TestResult)
        If temp(TestResult) = -1 Then
            arr(i - 1, TestName) = pPassedMessage & " - " & temp(TestName)
        ElseIf temp(TestResult) = 0 Then
            arr(i - 1, TestName) = pFailedMessage & " - " & temp(TestName)
        End If
    Next i
    
    getArrayFromCollection = arr
End Function

Private Sub addCategoryStyle(resultCol As ListColumn, categoryCol As ListColumn)
    Dim i As Long
    
    For i = 1 To resultCol.Range.Rows.Count
        If resultCol.Range.Cells(i, 1).value = -1 Then
            'categoryCol.Range.Cells(i, 1).Style = "Good"
            Call setGood(categoryCol.Range.Cells(i, 1))
        ElseIf resultCol.Range.Cells(i, 1).value = 0 Then
            'categoryCol.Range.Cells(i, 1).Style = "Bad"
            Call setBad(categoryCol.Range.Cells(i, 1))
        End If
    Next i
End Sub

Private Sub addValueConditColorsNew(lc As ListColumn, passedValue As String, failedValue As String)
    Dim cell As Range
    
    For Each cell In lc.DataBodyRange
        If cell.value Like passedValue Then
            Call setGood(cell)
        ElseIf cell.value Like failedValue Then
            Call setBad(cell)
        End If
    Next cell
End Sub

Private Sub setGood(cell As Variant)
    cell.Interior.Color = 13561798
    cell.Font.Color = 24832
End Sub

Private Sub setBad(cell As Variant)
    cell.Interior.Color = 13551615
    cell.Font.Color = 393372
End Sub


