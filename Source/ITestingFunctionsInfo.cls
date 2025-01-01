VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingFunctionsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Alphabetic() As ITestingFunctionsSubInfo
End Property

Public Property Get Alphanumeric() As ITestingFunctionsSubInfo
End Property

Public Property Get Between() As ITestingFunctionsSubInfo
End Property

Public Property Get Contain() As ITestingFunctionsSubInfo
End Property

Public Property Get DepthCountOf() As ITestingFunctionsSubInfo
End Property

Public Property Get Elements() As ITestingFunctionsSubInfo
End Property

Public Property Get ElementsInDataStructure() As ITestingFunctionsSubInfo
End Property

Public Property Get EndWith() As ITestingFunctionsSubInfo
End Property

Public Property Get EqualTo() As ITestingFunctionsSubInfo
End Property

Public Property Get Erroneous() As ITestingFunctionsSubInfo
End Property

Public Property Get ErrorDescriptionOf() As ITestingFunctionsSubInfo
End Property

Public Property Get ErrorNumberOf() As ITestingFunctionsSubInfo
End Property

Public Property Get EvaluateTo() As ITestingFunctionsSubInfo
End Property

Public Property Get ExactSameElementsAs() As ITestingFunctionsSubInfo
End Property

Public Property Get GreaterThan() As ITestingFunctionsSubInfo
End Property

Public Property Get GreaterThanOrEqualTo() As ITestingFunctionsSubInfo
End Property

Public Property Get IdenticalTo() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructure() As ITestingFunctionsSubInfo
End Property

Public Property Get InDataStructures() As ITestingFunctionsSubInfo
End Property

Public Property Get LengthBetween() As ITestingFunctionsSubInfo
End Property

Public Property Get LengthOf() As ITestingFunctionsSubInfo
End Property

Public Property Get LessThan() As ITestingFunctionsSubInfo
End Property

Public Property Get LessThanOrEqualTo() As ITestingFunctionsSubInfo
End Property

Public Property Get MaxLengthOf() As ITestingFunctionsSubInfo
End Property

Public Property Get MinLengthOf() As ITestingFunctionsSubInfo
End Property

Public Property Get NestedCountOf() As ITestingFunctionsSubInfo
End Property

Public Property Get Numeric() As ITestingFunctionsSubInfo
End Property

Public Property Get OneOf() As ITestingFunctionsSubInfo
End Property

Public Property Get Procedure() As ITestingFunctionsSubInfo
End Property

Public Property Get SameElementsAs() As ITestingFunctionsSubInfo
End Property

Public Property Get SameTypeAs() As ITestingFunctionsSubInfo
End Property

Public Property Get SameUniqueElementsAs() As ITestingFunctionsSubInfo
End Property

Public Property Get Something() As ITestingFunctionsSubInfo
End Property

Public Property Get StartWith() As ITestingFunctionsSubInfo
End Property

Public Sub initTestingFunctionsInfo()
End Sub

Public Function getTestingFunction(ByVal functionName As String) As ITestingFunctionsSubInfo
End Function

Public Sub PrintTestFunInfo()
End Sub

Public Function TestFuncInfoToDict() As Scripting.Dictionary
End Function

Public Function TestFuncInfoDictsAreIdentical(ByVal tfiDict1 As Scripting.Dictionary, ByVal tfiDict2 As Scripting.Dictionary) As Boolean
End Function

Public Function validateTfiDictCounters(ByVal tfiDict As Scripting.Dictionary, Optional ByVal counter As Long = 0) As Boolean
End Function

Private Sub populateDictWithTestFuncInfo()
End Sub
