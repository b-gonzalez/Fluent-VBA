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

'Add TestValueRecur, TestValueIter, TestInputRecur, and TestInputIter properties to ITestDev

    'These just need to be added as read-only properties whose implementation calls prepopulated values
    
    'A boolean property can also be added to determine whether the additional recursive or iterative
    'function needs to be called so that they aren't automatically populated everytime a test completes

Public Property Let NegateValue(ByVal value As Boolean)
End Property

Public Property Get NegateValue() As Boolean
End Property

Public Property Let resultNumber(ByVal value As Long)
End Property

Public Property Get resultNumber() As Long
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
