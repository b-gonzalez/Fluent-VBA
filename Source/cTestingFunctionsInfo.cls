VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestingFunctionsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestingFunctionsInfo

Private pAlphabetic As ITestingFunctionsSubInfo
Private pAlphanumeric As ITestingFunctionsSubInfo
Private pBetween As ITestingFunctionsSubInfo
Private pContain As ITestingFunctionsSubInfo
Private pElements As ITestingFunctionsSubInfo
Private pEndWith As ITestingFunctionsSubInfo
Private pEqualTo As ITestingFunctionsSubInfo
Private pErroneous As ITestingFunctionsSubInfo
Private pErrorDescriptionOf As ITestingFunctionsSubInfo
Private pErrorNumberOf As ITestingFunctionsSubInfo
Private pEvaluateTo As ITestingFunctionsSubInfo
Private pExactSameElementsAs As ITestingFunctionsSubInfo
Private pGreaterThan As ITestingFunctionsSubInfo
Private pGreaterThanOrEqualTo As ITestingFunctionsSubInfo
Private pIdenticalTo As ITestingFunctionsSubInfo
Private pInDataStructure As ITestingFunctionsSubInfo
Private pInDataStructures As ITestingFunctionsSubInfo
Private pLengthBetween As ITestingFunctionsSubInfo
Private pLengthOf As ITestingFunctionsSubInfo
Private pLessThan As ITestingFunctionsSubInfo
Private pLessThanOrEqualTo As ITestingFunctionsSubInfo
Private pMaxLengthOf As ITestingFunctionsSubInfo
Private pMinLengthOf As ITestingFunctionsSubInfo
Private pNumeric As ITestingFunctionsSubInfo
Private pOneOf As ITestingFunctionsSubInfo
Private pProcedure As ITestingFunctionsInfo
Private pSameElementsAs As ITestingFunctionsSubInfo
Private pSameTypeAs As ITestingFunctionsSubInfo
Private pSameUniqueElementsAs As ITestingFunctionsSubInfo
Private pSomething As ITestingFunctionsSubInfo
Private pStartWith As ITestingFunctionsSubInfo

Private pTestingFunctionsDict As Scripting.Dictionary

Private Property Get ITestingFunctionsInfo_Alphabetic() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Alphabetic = pAlphabetic
End Property

Private Property Get ITestingFunctionsInfo_Alphanumeric() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Alphanumeric = pAlphanumeric
End Property

Private Property Get ITestingFunctionsInfo_Between() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Between = pBetween
End Property

Private Property Get ITestingFunctionsInfo_Contain() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Contain = pContain
End Property

Private Property Get ITestingFunctionsInfo_Elements() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Elements = pElements
End Property

Private Property Get ITestingFunctionsInfo_EndWith() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_EndWith = pEndWith
End Property

Private Property Get ITestingFunctionsInfo_EqualTo() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_EqualTo = pEqualTo
End Property

Private Property Get ITestingFunctionsInfo_Erroneous() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Erroneous = pErroneous
End Property

Private Property Get ITestingFunctionsInfo_ErrorDescriptionOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_ErrorDescriptionOf = pErrorDescriptionOf
End Property

Private Property Get ITestingFunctionsInfo_ErrorNumberOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_ErrorNumberOf = pErrorNumberOf
End Property

Private Property Get ITestingFunctionsInfo_EvaluateTo() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_EvaluateTo = pEvaluateTo
End Property

Private Property Get ITestingFunctionsInfo_ExactSameElementsAs() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_ExactSameElementsAs = pExactSameElementsAs
End Property

Private Property Get ITestingFunctionsInfo_GreaterThan() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_GreaterThan = pGreaterThan
End Property

Private Property Get ITestingFunctionsInfo_GreaterThanOrEqualTo() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_GreaterThanOrEqualTo = pGreaterThanOrEqualTo
End Property

Private Property Get ITestingFunctionsInfo_IdenticalTo() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_IdenticalTo = pIdenticalTo
End Property

Private Property Get ITestingFunctionsInfo_InDataStructure() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_InDataStructure = pInDataStructure
End Property

Private Property Get ITestingFunctionsInfo_InDataStructures() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_InDataStructures = pInDataStructures
End Property

Private Property Get ITestingFunctionsInfo_LengthBetween() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_LengthBetween = pLengthBetween
End Property

Private Property Get ITestingFunctionsInfo_LengthOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_LengthOf = pLengthOf
End Property

Private Property Get ITestingFunctionsInfo_LessThan() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_LessThan = pLessThan
End Property

Private Property Get ITestingFunctionsInfo_LessThanOrEqualTo() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_LessThanOrEqualTo = pLessThanOrEqualTo
End Property

Private Property Get ITestingFunctionsInfo_MaxLengthOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_MaxLengthOf = pMaxLengthOf
End Property

Private Property Get ITestingFunctionsInfo_MinLengthOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_MinLengthOf = pMinLengthOf
End Property

Private Property Get ITestingFunctionsInfo_Numeric() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Numeric = pNumeric
End Property

Private Property Get ITestingFunctionsInfo_OneOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_OneOf = pOneOf
End Property

Private Property Get ITestingFunctionsInfo_Procedure() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Procedure = pProcedure
End Property

Private Property Get ITestingFunctionsInfo_SameElementsAs() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_SameElementsAs = pSameElementsAs
End Property

Private Property Get ITestingFunctionsInfo_SameTypeAs() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_SameTypeAs = pSameTypeAs
End Property

Private Property Get ITestingFunctionsInfo_SameUniqueElementsAs() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_SameUniqueElementsAs = pSameUniqueElementsAs
End Property

Private Property Get ITestingFunctionsInfo_Something() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Something = pSomething
End Property

Private Property Get ITestingFunctionsInfo_StartWith() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_StartWith = pStartWith
End Property

Private Function ITestingFunctionsInfo_getTestingFunction(functionName As String) As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_getTestingFunction = pTestingFunctionsDict(functionName)
End Function

Private Sub ITestingFunctionsInfo_populateDictWithTestFuncInfo()
    pTestingFunctionsDict.Add "Alphabetic", pAlphabetic
    pTestingFunctionsDict.Add "Alphanumeric", pAlphanumeric
    pTestingFunctionsDict.Add "Between", pBetween
    pTestingFunctionsDict.Add "Contain", pContain
    pTestingFunctionsDict.Add "Elements", pElements
    pTestingFunctionsDict.Add "EndWith", pEndWith
    pTestingFunctionsDict.Add "EqualTo", pEqualTo
    pTestingFunctionsDict.Add "Erroneous", pErroneous
    pTestingFunctionsDict.Add "ErrorDescriptionOf", pErrorDescriptionOf
    pTestingFunctionsDict.Add "ErrorNumberOf", pErrorNumberOf
    pTestingFunctionsDict.Add "EvaluateTo", pEvaluateTo
    pTestingFunctionsDict.Add "ExactSameElementsAs", pExactSameElementsAs
    pTestingFunctionsDict.Add "GreaterThan", pGreaterThan
    pTestingFunctionsDict.Add "GreaterThanOrEqualTo", pGreaterThanOrEqualTo
    pTestingFunctionsDict.Add "IdenticalTo", pIdenticalTo
    pTestingFunctionsDict.Add "InDataStructure", pInDataStructure
    pTestingFunctionsDict.Add "InDataStructures", pInDataStructures
    pTestingFunctionsDict.Add "LengthBetween", pLengthBetween
    pTestingFunctionsDict.Add "LengthOf", pLengthOf
    pTestingFunctionsDict.Add "LessThan", pLessThan
    pTestingFunctionsDict.Add "LessThanOrEqualTo", pLessThanOrEqualTo
    pTestingFunctionsDict.Add "MaxLengthOf", pMaxLengthOf
    pTestingFunctionsDict.Add "MinLengthOf", pMinLengthOf
    pTestingFunctionsDict.Add "Numeric", pNumeric
    pTestingFunctionsDict.Add "OneOf", pOneOf
    pTestingFunctionsDict.Add "Procedure", pProcedure
    pTestingFunctionsDict.Add "SameElementsAs", pSameElementsAs
    pTestingFunctionsDict.Add "SameTypeAs", pSameTypeAs
    pTestingFunctionsDict.Add "SameUniqueElementsAs", pSameUniqueElementsAs
    pTestingFunctionsDict.Add "Something", pSomething
    pTestingFunctionsDict.Add "StartWith", pStartWith
End Sub

Private Sub ITestingFunctionsInfo_initTestingFunctionsInfo()
    Dim i As Long
    Dim elem As Variant
    
    For Each elem In pTestingFunctionsDict.Keys
        Set pTestingFunctionsDict(elem) = New cTestingFunctionsSubInfo
    Next elem
End Sub

Private Sub ITestingFunctionsInfo_PrintTestFunInfo()
    Dim elem As Variant
    Dim tempStr As String
    Dim testFuncSubInfo As ITestingFunctionsSubInfo
    
    For Each elem In pTestingFunctionsDict.Keys
        Set testFuncSubInfo = pTestingFunctionsDict(elem)
    
        If testFuncSubInfo.Count > 0 Then
            Debug.Print "Test info for " & elem & ": "
        
            tempStr = vbTab & vbTab
        
            tempStr = tempStr & "Count: " & testFuncSubInfo.Count & "; "
    
            If testFuncSubInfo.Passed > 0 Then tempStr = tempStr & "Passed: " & testFuncSubInfo.Passed
    
            If testFuncSubInfo.Failed > 0 Then tempStr = tempStr & "; "
    
            If testFuncSubInfo.Failed > 0 Then tempStr = tempStr & "Failed: " & testFuncSubInfo.Failed
    
            If testFuncSubInfo.Passed > 0 Or testFuncSubInfo.Failed > 0 Then tempStr = tempStr & "; "
    
            If testFuncSubInfo.Unexpected > 0 Then tempStr = tempStr & "Unexpected: " & testFuncSubInfo.Unexpected
    
            Debug.Print tempStr
    
        End If
        
        tempStr = vbTab & vbTab
    Next elem
End Sub

Private Function ITestingFunctionsInfo_TestFuncInfoToDict() As Scripting.Dictionary
    Set ITestingFunctionsInfo_TestFuncInfoToDict = pTestingFunctionsDict
End Function

Private Function ITestingFunctionsInfo_TestFuncInfoDictsAreIdentical(tfiDict1 As Scripting.Dictionary, tfiDict2 As Scripting.Dictionary) As Boolean
    Dim b As Boolean
    Dim counter As Long
    Dim elem As Variant
    Dim testFuncSubInfo1 As ITestingFunctionsSubInfo
    Dim testFuncSubInfo2 As ITestingFunctionsSubInfo
    
    b = False
    counter = 0
    
    If UBound(tfiDict1.Keys) = UBound(tfiDict2.Keys) Then
        For Each elem In tfiDict1.Keys
            Set testFuncSubInfo1 = tfiDict1(elem)
            Set testFuncSubInfo2 = tfiDict2(elem)
            
            If testFuncSubInfo1.Count <> testFuncSubInfo2.Count And _
               testFuncSubInfo1.Passed <> testFuncSubInfo2.Passed And _
               testFuncSubInfo1.Failed <> testFuncSubInfo2.Failed And _
               testFuncSubInfo1.Unexpected <> testFuncSubInfo2.Unexpected _
            Then
                Exit For
            Else
                counter = counter + 1
            End If
               
        Next elem
    End If
    
    b = (counter = tfiDict1.Count And counter = tfiDict2.Count)
    
    ITestingFunctionsInfo_TestFuncInfoDictsAreIdentical = b
End Function

Private Function ITestingFunctionsInfo_validateTfiDictCounters(tfiDict As Scripting.Dictionary, Optional counter As Long = 0) As Boolean
    Dim b As Boolean
    Dim elem As Variant
    Dim testFuncSubInfo As ITestingFunctionsSubInfo
    
    For Each elem In tfiDict.Keys
        Set testFuncSubInfo = tfiDict(elem)
        
        If testFuncSubInfo.Count > 0 Then
            counter = counter + 1
        Else
            Debug.Print "Not in dict: " & elem
        End If
           
    Next elem
    
    b = (counter = tfiDict.Count)
    ITestingFunctionsInfo_validateTfiDictCounters = b
End Function

Private Sub Class_Initialize()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    
    Call ITestingFunctionsInfo_populateDictWithTestFuncInfo
    Call ITestingFunctionsInfo_initTestingFunctionsInfo
End Sub
