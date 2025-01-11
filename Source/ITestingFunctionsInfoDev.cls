VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingFunctionsInfoDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get DepthCountOfRecur() As ITestingFunctionsSubInfo
End Property

Public Property Get DepthCountOfIter() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructureRecur() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructureIter() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructuresRecur() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructuresIter() As ITestingFunctionsSubInfo
End Property

Public Property Get NestedCountOfRecur() As ITestingFunctionsSubInfo
End Property

Public Property Get NestedCountOfIter() As ITestingFunctionsSubInfo
End Property

Public Sub addRecurIterFuncName(value As String)
End Sub

Public Function getRecurIterFuncNameCol() As VBA.Collection
End Function
