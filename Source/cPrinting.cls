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

Implements IPrintExpr
Implements IPrinting

Private pPrintTotalTests As Boolean
Private pName As String
Private pTestCollection As Collection
Private pTestCount As Long
Private pCategory As String
Private iPrint As IPrintExpr
Private pPassedMessage As String
Private pFailedMessage As String
Private pUnexpectedMessage As String
Private pFluentPath As String
Private pExcel As Excel.Application
Private pWorkbook As Excel.Workbook
Private pExpressions As IExpressions
Private pCategorySet As Boolean
Private pNameSet As Boolean

Private Property Get IPrintExpr_TestCollection() As Collection
    If pTestCollection Is Nothing Then
        Set pTestCollection = New Collection
    End If
    Set IPrintExpr_TestCollection = pTestCollection
End Property

Private Property Let IPrintExpr_FluentPath(ByVal value As String)
    pFluentPath = value
End Property

Private Property Get IPrintExpr_Expressions() As IExpressions
    Set IPrintExpr_Expressions = pExpressions
End Property

Private Property Let IPrintExpr_CategorySet(ByVal value As Boolean)
    pCategorySet = value
End Property

Private Property Let IPrintExpr_NameSet(ByVal value As Boolean)
    pNameSet = value
End Property

Private Property Let IPrinting_Category(ByVal value As String)
    pCategory = value
End Property

Private Property Get IPrinting_Category() As String
    IPrinting_Category = pCategory
End Property

Private Property Let IPrinting_Name(ByVal value As String)
    pName = value
End Property

Private Property Get IPrinting_Name() As String
    IPrinting_Name = pName
End Property

Private Property Let IPrinting_PassedMessage(ByVal value As String)
    pPassedMessage = value
End Property

Private Property Get IPrinting_PassedMessage() As String
    IPrinting_PassedMessage = pPassedMessage
End Property

Private Property Let IPrinting_FailedMessage(ByVal value As String)
    pFailedMessage = value
End Property

Private Property Get IPrinting_FailedMessage() As String
    IPrinting_FailedMessage = pFailedMessage
End Property

Private Property Let IPrinting_UnexpectedMessage(ByVal value As String)
    pUnexpectedMessage = value
End Property

Private Property Get IPrinting_UnexpectedMessage() As String
    IPrinting_UnexpectedMessage = pUnexpectedMessage
End Property

Private Property Get IPrinting_FluentPath() As String
    IPrinting_FluentPath = pFluentPath
End Property

Private Function IPrinting_ArrayFromTests(col As Collection, tp As TestProperty)
    Dim test As ITest
    Dim testDev As ITestDev
    Dim elem As Variant
    Dim arr() As Variant
    Dim i As Long
    
    ReDim arr(1 To col.Count, 1 To 1)
    i = 1
    
    For Each elem In col
        Set test = elem
        Set testDev = elem
        
        Select Case tp
            Case TestProperty.TestCategory
                arr(i, 1) = test.Category
            Case TestProperty.TestInfo
                arr(i, 1) = test.FluentPath
            Case TestProperty.TestName
                arr(i, 1) = test.TestName
            Case TestProperty.TestResultNumber
                arr(i, 1) = testDev.resultNumber
        End Select
        i = i + 1
    Next elem
    
    IPrinting_ArrayFromTests = arr
End Function

Private Sub IPrinting_PrintToImmediate()
    Dim arr() As Variant
    Dim i As Long
    Dim temp As String
    Dim TestName As Long
    Dim testResultArr() As Variant
    Dim testInfoArr() As Variant
    Dim testNameArr() As Variant
    Dim testCategoryArr() As Variant
    Dim test As ITest
    Dim testDev As ITestDev
    
    TestName = 1
    temp = ""
    
    If pCategorySet Then testCategoryArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestCategory)
    
    If pNameSet Then testNameArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestName)
    
    testInfoArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestInfo)
    
    testResultArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestResultNumber)
    
    For Each test In pTestCollection
        Set testDev = test
        
        If test.Category <> "" Then
            temp = temp & test.Category & " - "
        End If
    
        If test.TestName <> "" Then
            temp = temp & test.TestName & " - "
        End If
        
        temp = temp & testDev.resultNumber & " - " & test.FluentPath
        
        Debug.Print temp
        
        temp = ""
    Next test
    
End Sub

Private Sub IPrinting_PrintToSheet()
    Dim lo As Excel.ListObject
    Dim lc As Excel.ListColumn
    Dim testResultCol As Excel.ListColumn
    Dim testNameCol As Excel.ListColumn
    Dim testCategoryCol As Excel.ListColumn
    Dim testInfoCol As Excel.ListColumn
    Dim arr() As Variant
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim testResultArr() As Variant
    Dim testInfoArr() As Variant
    Dim testNameArr() As Variant
    Dim testCategoryArr() As Variant

    Set wb = IPrinting_getWorkbook
    
    testInfoArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestInfo)
    
    wb.ActiveSheet.Range("A2").Resize(UBound(testInfoArr, 1), UBound(testInfoArr, 2)).value = testInfoArr
    
    Call IPrinting_addTableToWorkbook(wb)
    
    Set ws = wb.ActiveSheet
    
    Set lo = ws.ListObjects(1)
    
    Set testInfoCol = lo.ListColumns("Test info")

    testResultArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestResultNumber)
    
    Set testResultCol = lo.ListColumns.Add(Position:=1)

    testResultCol.Name = "Test result"
    
    testResultCol.DataBodyRange.value = testResultArr
    
    If pNameSet Then
        Set testNameCol = lo.ListColumns.Add(Position:=1)
        testNameCol.Name = "Test name"
        testNameArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestName)
        testNameCol.DataBodyRange.value = testNameArr
    End If

    If pCategorySet Then
        Set testCategoryCol = lo.ListColumns.Add(Position:=1)
        testCategoryCol.Name = "Category"
        testCategoryArr = IPrinting_ArrayFromTests(IPrintExpr_TestCollection, TestProperty.TestCategory)
        testCategoryCol.DataBodyRange.value = testCategoryArr
    End If
    
    Call IPrinting_addValueConditColors(lo, "-1", "0", "1")
    
    Call IPrinting_addResultsIconSets(testResultCol)
    
    testInfoCol.Range.EntireColumn.AutoFit
    
    Set pExcel = Nothing
    
    Set pWorkbook = Nothing

End Sub

Private Function IPrinting_getWorkbook() As Workbook
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
        
    Set IPrinting_getWorkbook = pWorkbook
End Function

Private Sub IPrinting_addTableToWorkbook(wb As Workbook)
    Dim ws As Excel.Worksheet
    Dim lo As Excel.ListObject
    Dim lc As Excel.ListColumn
    
    Set ws = wb.ActiveSheet
    
    ws.Range("A1").value = "Test info"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("$A$1").CurrentRegion, , xlYes)
    
    lo.Name = "Table1"
End Sub

Private Sub IPrinting_addResultsIconSets(ByVal lc As Excel.ListColumn)
    Dim wb As Workbook

    lc.Range.Select
    
    Set wb = IPrinting_getWorkbook()
    
    With lc.DataBodyRange
        .FormatConditions.AddIconSetCondition
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1)
            .ReverseOrder = False
            .ShowIconOnly = False
            .IconSet = wb.IconSets(xl3TrafficLights1)
        End With
        .FormatConditions(1).IconCriteria(1).Icon = xlIconGreenCheck
        With .FormatConditions(1).IconCriteria(2)
            .Type = xlConditionValueNumber
            .value = 0
            .Operator = 7
            .Icon = xlIconRedCross
        End With
        With .FormatConditions(1).IconCriteria(3)
            .Type = xlConditionValueNumber
            .value = 1
            .Operator = 7
            .Icon = xlIconYellowExclamation
        End With
    End With
End Sub

Private Sub IPrinting_addValueConditColors(ByVal lo As Excel.ListObject, ByVal passedMsg As String, ByVal failedMsg As String, ByVal unexpectedMsg As String)
    Dim lc As Excel.ListColumn
    Dim index As Long
    Dim temp As Variant
        
    Set lc = lo.ListColumns("Test result")
    index = lc.index
        
    lo.Range.AutoFilter Field:=index, Criteria1:="1"
    On Error Resume Next
        temp = lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        
        If Err.Number = 0 Then
            lo.DataBodyRange.Select
            Selection.Style = "Neutral"
        End If
        lo.Range.AutoFilter Field:=index
    On Error GoTo 0
    
    lo.Range.AutoFilter Field:=index, Criteria1:="-1"
    On Error Resume Next
        temp = lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        
        If Err.Number = 0 Then
            lo.DataBodyRange.Select
            Selection.Style = "Good"
        End If
        
        lo.Range.AutoFilter Field:=index
    On Error GoTo 0
    
    lo.Range.AutoFilter Field:=index, Criteria1:="0"
    On Error Resume Next
        temp = lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        
        If Err.Number = 0 Then
            lo.DataBodyRange.Select
            Selection.Style = "Bad"
        End If
        
        lo.Range.AutoFilter Field:=index
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    pPassedMessage = "PASSED"
    pFailedMessage = "FAILED"
    pUnexpectedMessage = "UNEXPECTED"
    
    Set pExpressions = New cExpressions
    Set pExpressions.setPrinting = Me
End Sub
