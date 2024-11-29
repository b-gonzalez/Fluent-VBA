VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingFunctionsSubInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let Count(ByVal value As Long)
End Property

Public Property Get Count() As Long
End Property

Public Property Let Passed(ByVal value As Long)
End Property

Public Property Get Passed() As Long
End Property

Public Property Let Failed(ByVal value As Long)
End Property

Public Property Get Failed() As Long
End Property

Public Property Let Unexpected(ByVal value As Long)
End Property

Public Property Get Unexpected() As Long
End Property
