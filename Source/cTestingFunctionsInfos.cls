VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestingFunctionsInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestingFunctionsInfoDev
Implements ITestingFunctionsInfos

Private pAlphabetic As cTestingFunctionsInfo
Private pAlphanumeric As cTestingFunctionsInfo
Private pBetween As cTestingFunctionsInfo
Private pContain As cTestingFunctionsInfo
Private pDepthCountOf As cTestingFunctionsInfo
Private pElements As cTestingFunctionsInfo
Private pElementsInDataStructure As cTestingFunctionsInfo
Private pEndWith As cTestingFunctionsInfo
Private pEqualTo As cTestingFunctionsInfo
Private pErroneous As cTestingFunctionsInfo
Private pErrorDescriptionOf As cTestingFunctionsInfo
Private pErrorNumberOf As cTestingFunctionsInfo
Private pEvaluateTo As cTestingFunctionsInfo
Private pExactSameElementsAs As cTestingFunctionsInfo
Private pGreaterThan As cTestingFunctionsInfo
Private pGreaterThanOrEqualTo As cTestingFunctionsInfo
Private pIdenticalTo As cTestingFunctionsInfo
Private pInDataStructure As cTestingFunctionsInfo
Private pInDataStructures As cTestingFunctionsInfo
Private pLengthBetween As cTestingFunctionsInfo
Private pLengthOf As cTestingFunctionsInfo
Private pLessThan As cTestingFunctionsInfo
Private pLessThanOrEqualTo As cTestingFunctionsInfo
Private pMaxLengthOf As cTestingFunctionsInfo
Private pMinLengthOf As cTestingFunctionsInfo
Private pNestedCountOf As cTestingFunctionsInfo
Private pNumeric As cTestingFunctionsInfo
Private pOneOf As cTestingFunctionsInfo
Private pProcedure As cTestingFunctionsInfo
Private pSameElementsAs As cTestingFunctionsInfo
Private pSameTypeAs As cTestingFunctionsInfo
Private pSameUniqueElementsAs As cTestingFunctionsInfo
Private pSomething As cTestingFunctionsInfo
Private pStartWith As cTestingFunctionsInfo

'ITestingFunctionsInfoDev

Private pDepthCountOfRecur As cTestingFunctionsInfo
Private pDepthCountOfIter As cTestingFunctionsInfo
Private pInDataStructureRecur As cTestingFunctionsInfo
Private pInDataStructureIter As cTestingFunctionsInfo
Private pInDataStructuresRecur As cTestingFunctionsInfo
Private pInDataStructuresIter As cTestingFunctionsInfo
Private pNestedCountOfRecur As cTestingFunctionsInfo
Private pNestedCountOfIter As cTestingFunctionsInfo

Private pTestingFunctionsDict As Scripting.Dictionary
Private pRecurIterFuncCol As VBA.Collection
Private pTestingFunctionsInfoCol As VBA.Collection
Private pUtilities As IUtilities
Private pAddDevTestingFunctionsInfo As Boolean

Private Property Get ITestingFunctionsInfos_item(ByVal OneBasedIndex As Long) As ITestingFunctionsInfo
    Dim tempCol As Collection
    
    Set tempCol = pUtilities.getElementFromCollection(pTestingFunctionsInfoCol, OneBasedIndex)
    
    If Not VBA.Information.IsNull(tempCol(1)) Then
        Set ITestingFunctionsInfos_item = tempCol(1)
    End If
End Property

Private Property Get ITestingFunctionsInfos_NewEnum() As IUnknown
    Set ITestingFunctionsInfos_NewEnum = pTestingFunctionsInfoCol.[_NewEnum]
End Property

Private Property Get ITestingFunctionsInfos_Alphabetic() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Alphabetic = pAlphabetic
End Property

Private Property Get ITestingFunctionsInfos_Alphanumeric() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Alphanumeric = pAlphanumeric
End Property

Private Property Get ITestingFunctionsInfos_Between() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Between = pBetween
End Property

Private Property Get ITestingFunctionsInfos_Contain() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Contain = pContain
End Property

Private Property Get ITestingFunctionsInfos_DepthCountOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_DepthCountOf = pDepthCountOf
End Property

Private Property Get ITestingFunctionsInfos_Elements() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Elements = pElements
End Property

Private Property Get ITestingFunctionsInfos_ElementsInDataStructure() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_ElementsInDataStructure = pElementsInDataStructure
End Property

Private Property Get ITestingFunctionsInfos_EndWith() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_EndWith = pEndWith
End Property

Private Property Get ITestingFunctionsInfos_EqualTo() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_EqualTo = pEqualTo
End Property

Private Property Get ITestingFunctionsInfos_Erroneous() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Erroneous = pErroneous
End Property

Private Property Get ITestingFunctionsInfos_ErrorDescriptionOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_ErrorDescriptionOf = pErrorDescriptionOf
End Property

Private Property Get ITestingFunctionsInfos_ErrorNumberOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_ErrorNumberOf = pErrorNumberOf
End Property

Private Property Get ITestingFunctionsInfos_EvaluateTo() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_EvaluateTo = pEvaluateTo
End Property

Private Property Get ITestingFunctionsInfos_ExactSameElementsAs() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_ExactSameElementsAs = pExactSameElementsAs
End Property

Private Property Get ITestingFunctionsInfos_GreaterThan() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_GreaterThan = pGreaterThan
End Property

Private Property Get ITestingFunctionsInfos_GreaterThanOrEqualTo() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_GreaterThanOrEqualTo = pGreaterThanOrEqualTo
End Property

Private Property Get ITestingFunctionsInfos_IdenticalTo() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_IdenticalTo = pIdenticalTo
End Property

Private Property Get ITestingFunctionsInfos_InDataStructure() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_InDataStructure = pInDataStructure
End Property

Private Property Get ITestingFunctionsInfos_InDataStructures() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_InDataStructures = pInDataStructures
End Property

Private Property Get ITestingFunctionsInfos_LengthBetween() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_LengthBetween = pLengthBetween
End Property

Private Property Get ITestingFunctionsInfos_LengthOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_LengthOf = pLengthOf
End Property

Private Property Get ITestingFunctionsInfos_LessThan() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_LessThan = pLessThan
End Property

Private Property Get ITestingFunctionsInfos_LessThanOrEqualTo() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_LessThanOrEqualTo = pLessThanOrEqualTo
End Property

Private Property Get ITestingFunctionsInfos_MaxLengthOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_MaxLengthOf = pMaxLengthOf
End Property

Private Property Get ITestingFunctionsInfos_MinLengthOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_MinLengthOf = pMinLengthOf
End Property

Private Property Get ITestingFunctionsInfos_NestedCountOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_NestedCountOf = pNestedCountOf
End Property

Private Property Get ITestingFunctionsInfos_Numeric() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Numeric = pNumeric
End Property

Private Property Get ITestingFunctionsInfos_OneOf() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_OneOf = pOneOf
End Property

Private Property Get ITestingFunctionsInfos_Procedure() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Procedure = pProcedure
End Property

Private Property Get ITestingFunctionsInfos_SameElementsAs() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_SameElementsAs = pSameElementsAs
End Property

Private Property Get ITestingFunctionsInfos_SameTypeAs() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_SameTypeAs = pSameTypeAs
End Property

Private Property Get ITestingFunctionsInfos_SameUniqueElementsAs() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_SameUniqueElementsAs = pSameUniqueElementsAs
End Property

Private Property Get ITestingFunctionsInfos_Something() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_Something = pSomething
End Property

Private Property Get ITestingFunctionsInfos_StartWith() As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_StartWith = pStartWith
End Property

Private Property Let ITestingFunctionsInfos_AddDevTestingFunctionsInfo(ByVal value As Boolean)
    ITestingFunctionsInfos_AddDevTestingFunctionsInfo = value
End Property

Private Property Get ITestingFunctionsInfos_AddDevTestingFunctionsInfo() As Boolean
    ITestingFunctionsInfos_AddDevTestingFunctionsInfo = pAddDevTestingFunctionsInfo
End Property


'ITestingFunctionsInfoDev settings

Private Property Get ITestingFunctionsInfoDev_DepthCountOfRecur() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_DepthCountOfRecur = pDepthCountOfRecur
End Property

Private Property Get ITestingFunctionsInfoDev_DepthCountOfIter() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_DepthCountOfIter = pDepthCountOfIter
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructureRecur() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_InDataStructureRecur = pInDataStructureRecur
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructureIter() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_InDataStructureIter = pInDataStructureIter
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructuresRecur() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_InDataStructuresRecur = pInDataStructuresRecur
End Property

Private Property Get ITestingFunctionsInfoDev_InDataStructuresIter() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_InDataStructuresIter = pInDataStructuresIter
End Property

Private Property Get ITestingFunctionsInfoDev_NestedCountOfRecur() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_NestedCountOfRecur = pNestedCountOfRecur
End Property

Private Property Get ITestingFunctionsInfoDev_NestedCountOfIter() As ITestingFunctionsInfo
    Set ITestingFunctionsInfoDev_NestedCountOfIter = pNestedCountOfIter
End Property

Private Function ITestingFunctionsInfos_getTestingFunction(ByVal functionName As String) As ITestingFunctionsInfo
    Set ITestingFunctionsInfos_getTestingFunction = pTestingFunctionsDict(functionName)
End Function

Private Sub populateDictWithTestFuncInfo()
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

Private Sub ITestingFunctionsInfos_initTestingFunctionsInfo()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    Call initTestingFunctionsInfoObjects
    Call populateDictWithTestFuncInfo
End Sub

Private Sub ITestingFunctionsInfos_PrintTestFunInfo()
    Dim elem As Variant
    Dim tempStr As String
    Dim testFuncSubInfo As ITestingFunctionsInfo
    
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

Private Function ITestingFunctionsInfos_TestFuncInfoToDict() As Scripting.Dictionary
    Set ITestingFunctionsInfos_TestFuncInfoToDict = pTestingFunctionsDict
End Function

Private Function ITestingFunctionsInfos_TestFuncInfoDictsAreIdentical(ByVal tfiDict1 As Scripting.Dictionary, ByVal tfiDict2 As Scripting.Dictionary) As Boolean
    Dim b As Boolean
    Dim counter As Long
    Dim elem As Variant
    Dim testFuncSubInfo1 As ITestingFunctionsInfo
    Dim testFuncSubInfo2 As ITestingFunctionsInfo
    
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
    
    ITestingFunctionsInfos_TestFuncInfoDictsAreIdentical = b
End Function

Private Function ITestingFunctionsInfos_validateTfiDictCounters(ByVal tfiDict As Scripting.Dictionary, Optional ByVal counter As Long = 0) As Boolean
    Dim b As Boolean
    Dim elem As Variant
    Dim testFuncSubInfo As ITestingFunctionsInfo
    Dim testFuncSubInfo2 As cTestingFunctionsInfo
    Dim recurIterFuncscount As Long
    Dim recurIterFuncscount2 As Long
    Dim tempCounter As Long
    
    recurIterFuncscount = 0
    recurIterFuncscount2 = 0
    tempCounter = 0
    
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
    
    For Each elem In tfiDict.Keys
        If Not elem Like "*Recur*" And Not elem Like "*Iter*" Then
            Set testFuncSubInfo2 = tfiDict(elem)
            
            If testFuncSubInfo2.Count > 0 Then
                tempCounter = tempCounter + 1
            Else
                Debug.Print "Not in dict: " & elem
            End If
        Else
            recurIterFuncscount2 = recurIterFuncscount2 + 1
        End If
    Next elem
    
    b = (counter = (tfiDict.Count - recurIterFuncscount) And tempCounter = (tfiDict.Count - recurIterFuncscount2))
    ITestingFunctionsInfos_validateTfiDictCounters = b
End Function

Private Sub initTestingFunctionsInfoObjects()
    Set pRecurIterFuncCol = New VBA.Collection
    Set pTestingFunctionsInfoCol = New VBA.Collection
End Sub

Private Sub ITestingFunctionsInfos_AddTestingFunctionsInfo(ByVal value As ITestingFunctionsInfo)
    Dim testInfoPreviouslyAdded As Boolean
    Dim elem As ITestingFunctionsInfo
    
    For Each elem In pTestingFunctionsInfoCol
        If elem.Name = value.Name Then
            testInfoPreviouslyAdded = True
            Exit For
        End If
    Next elem
    
    If Not testInfoPreviouslyAdded Then pTestingFunctionsInfoCol.Add value
End Sub

Private Sub ITestingFunctionsInfoDev_addRecurIterFuncName(ByVal value As String)
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

Private Sub Class_Initialize()
    Set pAlphabetic = New cTestingFunctionsInfo
    Set pAlphanumeric = New cTestingFunctionsInfo
    Set pBetween = New cTestingFunctionsInfo
    Set pContain = New cTestingFunctionsInfo
    Set pDepthCountOf = New cTestingFunctionsInfo
    Set pElements = New cTestingFunctionsInfo
    Set pElementsInDataStructure = New cTestingFunctionsInfo
    Set pEndWith = New cTestingFunctionsInfo
    Set pEqualTo = New cTestingFunctionsInfo
    Set pErroneous = New cTestingFunctionsInfo
    Set pErrorDescriptionOf = New cTestingFunctionsInfo
    Set pErrorNumberOf = New cTestingFunctionsInfo
    Set pEvaluateTo = New cTestingFunctionsInfo
    Set pExactSameElementsAs = New cTestingFunctionsInfo
    Set pGreaterThan = New cTestingFunctionsInfo
    Set pGreaterThanOrEqualTo = New cTestingFunctionsInfo
    Set pIdenticalTo = New cTestingFunctionsInfo
    Set pInDataStructure = New cTestingFunctionsInfo
    Set pInDataStructures = New cTestingFunctionsInfo
    Set pLengthBetween = New cTestingFunctionsInfo
    Set pLengthOf = New cTestingFunctionsInfo
    Set pLessThan = New cTestingFunctionsInfo
    Set pLessThanOrEqualTo = New cTestingFunctionsInfo
    Set pMaxLengthOf = New cTestingFunctionsInfo
    Set pMinLengthOf = New cTestingFunctionsInfo
    Set pNestedCountOf = New cTestingFunctionsInfo
    Set pNumeric = New cTestingFunctionsInfo
    Set pOneOf = New cTestingFunctionsInfo
    Set pProcedure = New cTestingFunctionsInfo
    Set pSameElementsAs = New cTestingFunctionsInfo
    Set pSameTypeAs = New cTestingFunctionsInfo
    Set pSameUniqueElementsAs = New cTestingFunctionsInfo
    Set pSomething = New cTestingFunctionsInfo
    Set pStartWith = New cTestingFunctionsInfo
    
    'ITestingFunctionsInfoDev
    
    Set pDepthCountOfRecur = New cTestingFunctionsInfo
    Set pDepthCountOfIter = New cTestingFunctionsInfo
    Set pInDataStructureRecur = New cTestingFunctionsInfo
    Set pInDataStructureIter = New cTestingFunctionsInfo
    Set pInDataStructuresRecur = New cTestingFunctionsInfo
    Set pInDataStructuresIter = New cTestingFunctionsInfo
    Set pNestedCountOfRecur = New cTestingFunctionsInfo
    Set pNestedCountOfIter = New cTestingFunctionsInfo

    Set pTestingFunctionsDict = New Scripting.Dictionary
    Set pRecurIterFuncCol = New VBA.Collection
    Set pTestingFunctionsInfoCol = New VBA.Collection
    Set pUtilities = New cUtilities
    
    Call initTestingFunctionsInfoObjects
    Call populateDictWithTestFuncInfo
End Sub

Private Sub Class_Terminate()
    Set pAlphabetic = Nothing
    Set pAlphanumeric = Nothing
    Set pBetween = Nothing
    Set pContain = Nothing
    Set pDepthCountOf = Nothing
    Set pElements = Nothing
    Set pElementsInDataStructure = Nothing
    Set pEndWith = Nothing
    Set pEqualTo = Nothing
    Set pErroneous = Nothing
    Set pErrorDescriptionOf = Nothing
    Set pErrorNumberOf = Nothing
    Set pEvaluateTo = Nothing
    Set pExactSameElementsAs = Nothing
    Set pGreaterThan = Nothing
    Set pGreaterThanOrEqualTo = Nothing
    Set pIdenticalTo = Nothing
    Set pInDataStructure = Nothing
    Set pInDataStructures = Nothing
    Set pLengthBetween = Nothing
    Set pLengthOf = Nothing
    Set pLessThan = Nothing
    Set pLessThanOrEqualTo = Nothing
    Set pMaxLengthOf = Nothing
    Set pMinLengthOf = Nothing
    Set pNestedCountOf = Nothing
    Set pNumeric = Nothing
    Set pOneOf = Nothing
    Set pProcedure = Nothing
    Set pSameElementsAs = Nothing
    Set pSameTypeAs = Nothing
    Set pSameUniqueElementsAs = Nothing
    Set pSomething = Nothing
    Set pStartWith = Nothing
    
    'ITestingFunctionsInfoDev
    
    Set pDepthCountOfRecur = Nothing
    Set pDepthCountOfIter = Nothing
    Set pInDataStructureRecur = Nothing
    Set pInDataStructureIter = Nothing
    Set pInDataStructuresRecur = Nothing
    Set pInDataStructuresIter = Nothing
    Set pNestedCountOfRecur = Nothing
    Set pNestedCountOfIter = Nothing
    
    Set pTestingFunctionsDict = Nothing
    Set pRecurIterFuncCol = Nothing
    
    Set pTestingFunctionsInfoCol = Nothing
    
    Set pUtilities = Nothing
End Sub
