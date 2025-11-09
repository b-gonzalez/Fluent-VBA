VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingFunctionsInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get item(ByVal OneBasedIndex As Long) As ITest
Attribute item.VB_UserMemId = 0
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Property

Public Property Get Alphabetic() As ITestingFunctionsInfo
End Property

Public Property Get Alphanumeric() As ITestingFunctionsInfo
End Property

Public Property Get Between() As ITestingFunctionsInfo
End Property

Public Property Get Contain() As ITestingFunctionsInfo
End Property

Public Property Get DepthCountOf() As ITestingFunctionsInfo
End Property

Public Property Get Elements() As ITestingFunctionsInfo
End Property

Public Property Get ElementsInDataStructure() As ITestingFunctionsInfo
End Property

Public Property Get EndWith() As ITestingFunctionsInfo
End Property

Public Property Get EqualTo() As ITestingFunctionsInfo
End Property

Public Property Get Erroneous() As ITestingFunctionsInfo
End Property

Public Property Get ErrorDescriptionOf() As ITestingFunctionsInfo
End Property

Public Property Get ErrorNumberOf() As ITestingFunctionsInfo
End Property

Public Property Get EvaluateTo() As ITestingFunctionsInfo
End Property

Public Property Get ExactSameElementsAs() As ITestingFunctionsInfo
End Property

Public Property Get GreaterThan() As ITestingFunctionsInfo
End Property

Public Property Get GreaterThanOrEqualTo() As ITestingFunctionsInfo
End Property

Public Property Get IdenticalTo() As ITestingFunctionsInfo
End Property

Public Property Get InDataStructure() As ITestingFunctionsInfo
End Property

Public Property Get InDataStructures() As ITestingFunctionsInfo
End Property

Public Property Get LengthBetween() As ITestingFunctionsInfo
End Property

Public Property Get LengthOf() As ITestingFunctionsInfo
End Property

Public Property Get LessThan() As ITestingFunctionsInfo
End Property

Public Property Get LessThanOrEqualTo() As ITestingFunctionsInfo
End Property

Public Property Get MaxLengthOf() As ITestingFunctionsInfo
End Property

Public Property Get MinLengthOf() As ITestingFunctionsInfo
End Property

Public Property Get NestedCountOf() As ITestingFunctionsInfo
End Property

Public Property Get Numeric() As ITestingFunctionsInfo
End Property

Public Property Get OneOf() As ITestingFunctionsInfo
End Property

Public Property Get Procedure() As ITestingFunctionsInfo
End Property

Public Property Get SameElementsAs() As ITestingFunctionsInfo
End Property

Public Property Get SameTypeAs() As ITestingFunctionsInfo
End Property

Public Property Get SameUniqueElementsAs() As ITestingFunctionsInfo
End Property

Public Property Get Something() As ITestingFunctionsInfo
End Property

Public Property Get StartWith() As ITestingFunctionsInfo
End Property

Public Property Let AddDevTestingFunctionsInfo(ByVal value As Boolean)
End Property

Public Property Get AddDevTestingFunctionsInfo() As Boolean
End Property

Public Sub initTestingFunctionsInfo()
End Sub

Public Sub PrintTestFunInfo()
End Sub

Public Function TestFuncInfoToDict() As Scripting.Dictionary
End Function

Public Function TestFuncInfoDictsAreIdentical(ByVal tfiDict1 As Scripting.Dictionary, ByVal tfiDict2 As Scripting.Dictionary) As Boolean
End Function

Public Function validateTfiDictCounters(ByVal tfiDict As Scripting.Dictionary, Optional ByVal counter As Long = 0) As Boolean
End Function

Public Sub AddTestingFunctionsInfo(ByVal value As ITestingFunctionsInfo)
End Sub

Public Function getTestingFunction(ByVal functionName As String) As ITestingFunctionsInfo
End Function
