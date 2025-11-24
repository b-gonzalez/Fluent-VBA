VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrintingFluentFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Public Property Set PathDict(value As Scripting.Dictionary)
End Property

Public Property Set FunctionDict(value As Scripting.Dictionary)
End Property

Public Sub printPathDict()
End Sub

Public Sub queryAndPrintFunctionDict(query As String)
End Sub
