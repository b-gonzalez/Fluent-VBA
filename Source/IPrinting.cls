VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum TestProperty
    TestCategory
    TestInfo
    TestName
    TestResultNumber
End Enum

Public Property Let Category(ByVal value As String)
End Property

Public Property Get Category() As String
End Property

Public Property Let Name(ByVal value As String)
End Property

Public Property Get Name() As String
End Property

Public Property Let PassedMessage(ByVal value As String)
End Property

Public Property Get PassedMessage() As String
End Property

Public Property Let FailedMessage(ByVal value As String)
End Property

Public Property Get FailedMessage() As String
End Property

Public Property Let UnexpectedMessage(ByVal value As String)
End Property

Public Property Get UnexpectedMessage() As String
End Property

Public Property Get FluentPath() As String
End Property

Public Sub PrintToImmediate()
End Sub

Public Sub PrintToSheet()
End Sub

Private Function getWorkbook() As Workbook
End Function

Private Sub addTableToWorkbook(wb As Workbook)
End Sub

Private Sub addResultsIconSets(ByVal lc As Excel.ListColumn)
End Sub

Private Sub addValueConditColors( _
    ByVal lo As Excel.ListObject, _
    ByVal passedMsg As String, _
    ByVal failedMsg As String, _
    ByVal unexpectedMsg As String)
End Sub

Private Function ArrayFromTests(col As Collection, tp As TestProperty) As Variant()
End Function

Private Function getTestMessageFromResultNumber(ResultNumber As Long)
End Function
