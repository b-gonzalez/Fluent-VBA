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

Public Function Alphabetic() As ITestingFunctionsSubInfo
End Function

Public Function Alphanumeric() As ITestingFunctionsSubInfo
End Function

Public Function Between() As ITestingFunctionsSubInfo
End Function

Public Function Contain() As ITestingFunctionsSubInfo
End Function

Public Function EndWith() As ITestingFunctionsSubInfo
End Function

Public Function EqualTo() As ITestingFunctionsSubInfo
End Function

Public Function Erroneous() As ITestingFunctionsSubInfo
End Function

Public Function ErrorDescriptionOf() As ITestingFunctionsSubInfo
End Function

Public Function ErrorNumberOf() As ITestingFunctionsSubInfo
End Function

Public Function EvaluateTo() As ITestingFunctionsSubInfo
End Function

Public Function ExactSameElementsAs() As ITestingFunctionsSubInfo
End Function

Public Function GreaterThan() As ITestingFunctionsSubInfo
End Function

Public Function GreaterThanOrEqualTo() As ITestingFunctionsSubInfo
End Function

Public Function IdenticalTo() As ITestingFunctionsSubInfo
End Function

Public Function InDataStructure() As ITestingFunctionsSubInfo
End Function

Public Function InDataStructures() As ITestingFunctionsSubInfo
End Function

Public Function LengthBetween() As ITestingFunctionsSubInfo
End Function

Public Function LengthOf() As ITestingFunctionsSubInfo
End Function

Public Function LessThan() As ITestingFunctionsSubInfo
End Function

Public Function LessThanOrEqualTo() As ITestingFunctionsSubInfo
End Function

Public Function MaxLengthOf() As ITestingFunctionsSubInfo
End Function

Public Function MinLengthOf() As ITestingFunctionsSubInfo
End Function

Public Function Numeric() As ITestingFunctionsSubInfo
End Function

Public Function OneOf() As ITestingFunctionsSubInfo
End Function

Public Function SameElementsAs() As ITestingFunctionsSubInfo
End Function

Public Function SameTypeAs() As ITestingFunctionsSubInfo
End Function

Public Function SameUniqueElementsAs() As ITestingFunctionsSubInfo
End Function

Public Function Something() As ITestingFunctionsSubInfo
End Function

Public Function StartWith() As ITestingFunctionsSubInfo
End Function

Public Sub initTestingFunctionsInfo()
End Sub

Public Function getTestingFunction(functionName As String) As ITestingFunctionsSubInfo
End Function

Public Sub PrintTestFunInfo()
End Sub

Public Function TestFuncInfoToDict() As Scripting.Dictionary
End Function

Public Function TestFuncInfoDictsAreIdentical(tfiDict1 As Scripting.Dictionary, tfiDict2 As Scripting.Dictionary) As Boolean
End Function

Public Function validateTfiDictCounters(tfiDict As Scripting.Dictionary, Optional counter As Long = 0) As Boolean
End Function

Private Sub populateDictWithTestFuncInfo()
End Sub
