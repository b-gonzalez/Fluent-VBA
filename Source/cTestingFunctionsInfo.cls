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
Implements ITestingFunctionsInfoDev

Private pRecurIterFuncCol As VBA.Collection

Private pAlphabetic As New cTestingFunctionsSubInfo
Private pAlphanumeric As New cTestingFunctionsSubInfo
Private pBetween As New cTestingFunctionsSubInfo
Private pContain As New cTestingFunctionsSubInfo
Private pDepthCountOf As New cTestingFunctionsSubInfo
Private pElements As New cTestingFunctionsSubInfo
Private pElementsInDataStructure As New cTestingFunctionsSubInfo
Private pEndWith As New cTestingFunctionsSubInfo
Private pEqualTo As New cTestingFunctionsSubInfo
Private pErroneous As New cTestingFunctionsSubInfo
Private pErrorDescriptionOf As New cTestingFunctionsSubInfo
Private pErrorNumberOf As New cTestingFunctionsSubInfo
Private pEvaluateTo As New cTestingFunctionsSubInfo
Private pExactSameElementsAs As New cTestingFunctionsSubInfo
Private pGreaterThan As New cTestingFunctionsSubInfo
Private pGreaterThanOrEqualTo As New cTestingFunctionsSubInfo
Private pIdenticalTo As New cTestingFunctionsSubInfo
Private pInDataStructure As New cTestingFunctionsSubInfo
Private pInDataStructures As New cTestingFunctionsSubInfo
Private pLengthBetween As New cTestingFunctionsSubInfo
Private pLengthOf As New cTestingFunctionsSubInfo
Private pLessThan As New cTestingFunctionsSubInfo
Private pLessThanOrEqualTo As New cTestingFunctionsSubInfo
Private pMaxLengthOf As New cTestingFunctionsSubInfo
Private pMinLengthOf As New cTestingFunctionsSubInfo
Private pNestedCountOf As New cTestingFunctionsSubInfo
Private pNumeric As New cTestingFunctionsSubInfo
Private pOneOf As New cTestingFunctionsSubInfo
Private pProcedure As New cTestingFunctionsSubInfo
Private pSameElementsAs As New cTestingFunctionsSubInfo
Private pSameTypeAs As New cTestingFunctionsSubInfo
Private pSameUniqueElementsAs As New cTestingFunctionsSubInfo
Private pSomething As New cTestingFunctionsSubInfo
Private pStartWith As New cTestingFunctionsSubInfo

'ITestingFunctionsInfoDev

Private pDepthCountOfRecur As New cTestingFunctionsSubInfo
Private pDepthCountOfIter As New cTestingFunctionsSubInfo
Private pInDataStructureRecur As New cTestingFunctionsSubInfo
Private pInDataStructureIter As New cTestingFunctionsSubInfo
Private pInDataStructuresRecur As New cTestingFunctionsSubInfo
Private pInDataStructuresIter As New cTestingFunctionsSubInfo
Private pNestedCountOfRecur As New cTestingFunctionsSubInfo
Private pNestedCountOfIter As New cTestingFunctionsSubInfo

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

Private Property Get ITestingFunctionsInfo_DepthCountOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_DepthCountOf = pDepthCountOf
End Property

Private Property Get ITestingFunctionsInfo_Elements() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_Elements = pElements
End Property

Private Property Get ITestingFunctionsInfo_ElementsInDataStructure() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_ElementsInDataStructure = pElementsInDataStructure
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

Private Property Get ITestingFunctionsInfo_NestedCountOf() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_NestedCountOf = pNestedCountOf
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

Private Function ITestingFunctionsInfo_getTestingFunction(ByVal functionName As String) As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfo_getTestingFunction = pTestingFunctionsDict(functionName)
End Function

'ITestingFunctionsInfoDev settings

Private Property Get ITestingFunctionsInfoDev_DepthCountOfRecur() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_DepthCountOfRecur = pDepthCountOfRecur
End Property

Private Property Get ITestingFunctionsInfoDev_DepthCountOfIter() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_DepthCountOfIter = pDepthCountOfIter
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructureRecur() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_InDataStructureRecur = pInDataStructureRecur
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructureIter() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_InDataStructureIter = pInDataStructureIter
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructuresRecur() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_InDataStructuresRecur = pInDataStructuresRecur
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructuresIter() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_InDataStructuresIter = pInDataStructuresIter
End Property

Private Property Get ITestingFunctionsInfoDev_NestedCountOfRecur() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_NestedCountOfRecur = pNestedCountOfRecur
End Property

Private Property Get ITestingFunctionsInfoDev_NestedCountOfIter() As ITestingFunctionsSubInfo
    Set ITestingFunctionsInfoDev_NestedCountOfIter = pNestedCountOfIter
End Property

Private Sub ITestingFunctionsInfoDev_addRecurIterFuncName(value As String)
    Dim i As Long
    
    For i = 1 To pRecurIterFuncCol.Count
        If pRecurIterFuncCol(i) = value Then
            Exit Sub
        End If
    Next i
    
    pRecurIterFuncCol.Add value
End Sub

Private Function ITestingFunctionsInfoDev_getRecurIterFuncNameCol() As VBA.Collection
    Set ITestingFunctionsInfoDev_getRecurIterFuncNameCol = pRecurIterFuncCol
End Function

Private Sub ITestingFunctionsInfo_populateDictWithTestFuncInfo()
    pTestingFunctionsDict.Add "Alphabetic", pAlphabetic
    pTestingFunctionsDict.Add "Alphanumeric", pAlphanumeric
    pTestingFunctionsDict.Add "Between", pBetween
    pTestingFunctionsDict.Add "Contain", pContain
    pTestingFunctionsDict.Add "DepthCountOf", pDepthCountOf
    pTestingFunctionsDict.Add "Elements", pElements
    pTestingFunctionsDict.Add "ElementsInDataStructure", pElementsInDataStructure
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
    pTestingFunctionsDict.Add "NestedCountOf", pNestedCountOf
    pTestingFunctionsDict.Add "Numeric", pNumeric
    pTestingFunctionsDict.Add "OneOf", pOneOf
    pTestingFunctionsDict.Add "Procedure", pProcedure
    pTestingFunctionsDict.Add "SameElementsAs", pSameElementsAs
    pTestingFunctionsDict.Add "SameTypeAs", pSameTypeAs
    pTestingFunctionsDict.Add "SameUniqueElementsAs", pSameUniqueElementsAs
    pTestingFunctionsDict.Add "Something", pSomething
    pTestingFunctionsDict.Add "StartWith", pStartWith
    
    'ITestingFunctionsInfoDev
    
    pTestingFunctionsDict.Add "DepthCountOfRecur", pDepthCountOfRecur
    pTestingFunctionsDict.Add "DepthCountOfIter", pDepthCountOfIter
    pTestingFunctionsDict.Add "InDataStructureRecur", pInDataStructureRecur
    pTestingFunctionsDict.Add "InDataStructureIter", pInDataStructureIter
    pTestingFunctionsDict.Add "InDataStructuresRecur", pInDataStructuresRecur
    pTestingFunctionsDict.Add "InDataStructuresIter", pInDataStructuresIter
    pTestingFunctionsDict.Add "NestedCountOfRecur", pNestedCountOfRecur
    pTestingFunctionsDict.Add "NestedCountOfIter", pNestedCountOfIter
End Sub

Private Sub ITestingFunctionsInfo_initTestingFunctionsInfo()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    Call initTestingFunctionsSubInfoObjects
    Call ITestingFunctionsInfo_populateDictWithTestFuncInfo

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

Private Function ITestingFunctionsInfo_TestFuncInfoDictsAreIdentical(ByVal tfiDict1 As Scripting.Dictionary, ByVal tfiDict2 As Scripting.Dictionary) As Boolean
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

Private Function ITestingFunctionsInfo_validateTfiDictCounters(ByVal tfiDict As Scripting.Dictionary, Optional ByVal counter As Long = 0) As Boolean
    Dim b As Boolean
    Dim elem As Variant
    Dim testFuncSubInfo As ITestingFunctionsSubInfo
    Dim recurIterFuncscount As Long
    
    recurIterFuncscount = 0
    
    For Each elem In tfiDict.Keys
        If Not elem Like "*Recur*" And Not elem Like "*Iter*" Then
            Set testFuncSubInfo = tfiDict(elem)
            
            If testFuncSubInfo.Count > 0 Then
                counter = counter + 1
            Else
                Debug.Print "Not in dict: " & elem
            End If
        Else
            recurIterFuncscount = recurIterFuncscount + 1
        End If
    Next elem
    
    b = (counter = (tfiDict.Count - recurIterFuncscount))
    ITestingFunctionsInfo_validateTfiDictCounters = b
End Function

Private Sub initTestingFunctionsSubInfoObjects()
    Set pRecurIterFuncCol = New VBA.Collection
End Sub

Private Sub Class_Initialize()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    
    Call initTestingFunctionsSubInfoObjects
    Call ITestingFunctionsInfo_populateDictWithTestFuncInfo
End Sub
