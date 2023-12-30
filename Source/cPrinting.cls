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
Private pFluentPath As String
Private pExcel As Excel.Application
Private pWorkbook As Excel.Workbook

Private Property Get IPrinting_TestCollection() As Collection
    If pTestCollection Is Nothing Then
        Set pTestCollection = New Collection
    End If
    Set IPrinting_TestCollection = pTestCollection
End Property

Private Property Let IPrinting_FluentPath(value As String)
    pFluentPath = value
End Property

Public Property Let Category(ByVal value As String)
    pCategory = value
End Property

Public Property Get Category() As String
    Category = pCategory
End Property

Public Property Let TestName(ByVal value As String)
    pTestName = value
End Property

Public Property Get TestName() As String
    TestName = pTestName
End Property

Public Property Let PassedMessage(ByVal value As String)
    pPassedMessage = value
End Property

Public Property Get PassedMessage() As String
    PassedMessage = pPassedMessage
End Property

Public Property Let FailedMessage(ByVal value As String)
    pFailedMessage = value
End Property

Public Property Get FailedMessage() As String
    FailedMessage = pFailedMessage
End Property

Public Property Get FluentPath() As String
    FluentPath = pFluentPath
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
    Dim lo As Excel.ListObject
    Dim testResultCol As Excel.ListColumn
    Dim testNameCol As Excel.ListColumn
    Dim categoryCol As Excel.ListColumn
    Dim arr() As Variant
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    
    arr = getArrayFromCollection(pTestCollection)

    Set wb = getWorkbook
    
    Call addTableToWorkbook(wb)
    
    Set ws = wb.ActiveSheet
    
    Set lo = ws.ListObjects(1)
    
    Set testResultCol = lo.ListColumns("Test result")
    
    Set testNameCol = lo.ListColumns("Test name")
    
    If pCategory <> "" Then
        Set categoryCol = lo.ListColumns("Category")
        lo.ListRows.Add.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).value = arr
        ws.Columns("C:C").EntireColumn.AutoFit
        Call addCategoryStyle(categoryCol, "-1", "0")
    Else
        lo.ListRows.Add.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).value = arr
        ws.Columns("B:B").EntireColumn.AutoFit
    End If

    Call addValueConditColors(testResultCol, "-1", "0")
    
    Call addValueConditColors(testNameCol, "*" & pPassedMessage & "*", "*" & pFailedMessage & "*")
    
    Call addResultsIconSets(testResultCol)

End Sub

Private Function getWorkbook() As Workbook
    Dim xl As Excel.Application
    
    If pExcel Is Nothing Then
        If Not TypeOf Application Is Excel.Application Then
            Set pExcel = New Excel.Application
        Else
            Set pExcel = Application
        End If
    End If
    
    pExcel.Visible = True
    
    If pWorkbook Is Nothing Then
        Set pWorkbook = pExcel.Workbooks.Add
    End If
        
    Set getWorkbook = pWorkbook
End Function

Private Sub addTableToWorkbook(wb As Workbook)
    Dim ws As Excel.Worksheet
    Dim lo As Excel.ListObject
    
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
End Sub

Private Sub addResultsIconSets(ByVal lc As Excel.ListColumn)
    Dim wb As Workbook

    lc.Range.Select
    
    Set wb = getWorkbook()
    
    With lc.DataBodyRange
        .FormatConditions.AddIconSetCondition
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1)
            .ReverseOrder = False
            .ShowIconOnly = False
            .IconSet = wb.IconSets(xl3Symbols2)
        End With
        .FormatConditions(1).IconCriteria(1).Icon = xlIconYellowExclamation
        With .FormatConditions(1).IconCriteria(2)
            .Type = xlConditionValueNumber
            .value = -1
            .Operator = 7
            .Icon = xlIconGreenCheck
        End With
        With .FormatConditions(1).IconCriteria(3)
            .Type = xlConditionValueNumber
            .value = 0
            .Operator = 7
            .Icon = xlIconRedCross
        End With
    End With
End Sub

Private Function getArrayFromCollection(ByVal Col As Collection) As Variant()
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

Private Sub addCategoryStyleOld(ByVal resultCol As Excel.ListColumn, ByVal categoryCol As Excel.ListColumn)
    Dim i As Long
    
    For i = 1 To resultCol.Range.Rows.Count
        If resultCol.Range.Cells(i, 1).value = -1 Then
            Call setGood(categoryCol.Range.Cells(i, 1))
        ElseIf resultCol.Range.Cells(i, 1).value = 0 Then
            Call setBad(categoryCol.Range.Cells(i, 1))
        End If
    Next i
End Sub

Sub addCategoryStyle(ByVal lc As Excel.ListColumn, ByVal passedMsg As String, ByVal failedMsg As String)
    
    With lc.Range
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OFFSET(A1,0,1)=" & failedMsg
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
        
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OFFSET(A1,0,1)=" & passedMsg
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

Private Sub addValueConditColorsNew(ByVal lc As Excel.ListColumn, ByVal passedValue As String, ByVal failedValue As String)
    Dim cell As Excel.Range
    
    For Each cell In lc.DataBodyRange
        If cell.value Like passedValue Then
            Call setGood(cell)
        ElseIf cell.value Like failedValue Then
            Call setBad(cell)
        End If
    Next cell
End Sub

Private Sub addValueConditColors(ByVal lc As Excel.ListColumn, ByVal passedMsg As String, ByVal failedMsg As String)
    lc.Range.Select
    
    With lc.DataBodyRange
        .FormatConditions.Add Type:=xlTextString, String:=failedMsg, _
            TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlTextString, String:=passedMsg, _
            TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

Private Sub setGood(ByVal cell As Variant)
    cell.Interior.Color = 13561798
    cell.Font.Color = 24832
End Sub

Private Sub setBad(ByVal cell As Variant)
    cell.Interior.Color = 13551615
    cell.Font.Color = 393372
End Sub

Private Sub Class_Initialize()
    pPassedMessage = "PASSED"
    pFailedMessage = "FAILED"
End Sub
