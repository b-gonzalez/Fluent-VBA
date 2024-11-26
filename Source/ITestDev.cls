VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let negateValue(ByVal value As Boolean)
End Property

Public Property Get negateValue() As Boolean
End Property

Public Property Let ResultNumber(ByVal value As Long)
End Property

Public Property Get ResultNumber() As Long
End Property

Public Property Let TestValueRecur(value As String)
End Property

Public Property Get TestValueRecur() As String
End Property

Public Property Let TestValueIter(value As String)
End Property

Public Property Get TestValueIter() As String
End Property

Public Property Let TestInputRecur(value As String)
End Property

Public Property Get TestInputRecur() As String
End Property

Public Property Let TestInputIter(value As String)
End Property

Public Property Get TestInputIter() As String
End Property
