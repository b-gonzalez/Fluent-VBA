VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get result() As Boolean
End Property

Public Property Get item(ByVal OneBasedIndex As Long) As ITest
Attribute item.VB_UserMemId = 0
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Property

Public Property Let Algorithm(ByVal value As flAlgorithm)
End Property

Public Property Get Algorithm() As flAlgorithm
End Property

Public Property Get AlgorithmValueSet() As Boolean
End Property

Public Property Let ToStrDev(ByVal value As Boolean)
End Property

Public Property Get ToStrDev() As Boolean
End Property

Public Property Get Count() As Long
End Property

Public Property Let ApproximateEqual(ByVal value As Boolean)
End Property

Public Property Get ApproximateEqual() As Boolean
End Property

Public Property Let Epsilon(ByVal value As Double)
End Property

Public Property Get Epsilon() As Double
End Property

Public Property Get DataStructures() As VBA.Collection
End Property

Public Property Let SkipDupCheck(ByVal value As Boolean)
End Property

Public Property Get SkipDupCheck() As Boolean
End Property

Public Property Get TestingFunctionsInfos() As ITestingFunctionsInfos
End Property

Public Property Get TestStrings() As ITestStrings
End Property

Public Sub AddDataStructure(ByVal dataStructure As Variant)
End Sub

Public Function IsDataStructure(ByVal dataStructure As Variant) As Boolean
End Function

Public Sub CheckTest(ByVal test As ITest)
End Sub

Public Sub resetTestingInfo()
End Sub

Public Sub ResetCounter()
End Sub

Public Function DatastructureIsSelfReferential(ByVal ds As Variant, ByVal Algorithm As flAlgorithm) As Boolean
End Function

Public Sub AddFromFluent(ByVal fluent As cFluent)
End Sub

Public Sub AddFromFluentOf(ByVal fluentOf As cFluentOf)
End Sub

Public Sub AddFromTests(ByVal tests As cTests)
End Sub
