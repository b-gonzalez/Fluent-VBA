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

Private pAlphabetic As New cTestingFunctionsInfo
Private pAlphanumeric As New cTestingFunctionsInfo
Private pBetween As New cTestingFunctionsInfo
Private pContain As New cTestingFunctionsInfo
Private pDepthCountOf As New cTestingFunctionsInfo
Private pElements As New cTestingFunctionsInfo
Private pElementsInDataStructure As New cTestingFunctionsInfo
Private pEndWith As New cTestingFunctionsInfo
Private pEqualTo As New cTestingFunctionsInfo
Private pErroneous As New cTestingFunctionsInfo
Private pErrorDescriptionOf As New cTestingFunctionsInfo
Private pErrorNumberOf As New cTestingFunctionsInfo
Private pEvaluateTo As New cTestingFunctionsInfo
Private pExactSameElementsAs As New cTestingFunctionsInfo
Private pGreaterThan As New cTestingFunctionsInfo
Private pGreaterThanOrEqualTo As New cTestingFunctionsInfo
Private pIdenticalTo As New cTestingFunctionsInfo
Private pInDataStructure As New cTestingFunctionsInfo
Private pInDataStructures As New cTestingFunctionsInfo
Private pLengthBetween As New cTestingFunctionsInfo
Private pLengthOf As New cTestingFunctionsInfo
Private pLessThan As New cTestingFunctionsInfo
Private pLessThanOrEqualTo As New cTestingFunctionsInfo
Private pMaxLengthOf As New cTestingFunctionsInfo
Private pMinLengthOf As New cTestingFunctionsInfo
Private pNestedCountOf As New cTestingFunctionsInfo
Private pNumeric As New cTestingFunctionsInfo
Private pOneOf As New cTestingFunctionsInfo
Private pProcedure As New cTestingFunctionsInfo
Private pSameElementsAs As New cTestingFunctionsInfo
Private pSameTypeAs As New cTestingFunctionsInfo
Private pSameUniqueElementsAs As New cTestingFunctionsInfo
Private pSomething As New cTestingFunctionsInfo
Private pStartWith As New cTestingFunctionsInfo

'ITestingFunctionsInfoDev

Private pDepthCountOfRecur As New cTestingFunctionsInfo
Private pDepthCountOfIter As New cTestingFunctionsInfo
Private pInDataStructureRecur As New cTestingFunctionsInfo
Private pInDataStructureIter As New cTestingFunctionsInfo
Private pInDataStructuresRecur As New cTestingFunctionsInfo
Private pInDataStructuresIter As New cTestingFunctionsInfo
Private pNestedCountOfRecur As New cTestingFunctionsInfo
Private pNestedCountOfIter As New cTestingFunctionsInfo

Private pTestingFunctionsDict As Scripting.Dictionary
Private pRecurIterFuncCol As VBA.Collection
Private pTestingFunctionsInfoCol As VBA.Collection
Private pAddDevTestingFunctionsInfo As Boolean

Private Property Get ITestingFunctionsInfos_item(ByVal OneBasedIndex As Long) As ITest
    Set ITestingFunctionsInfos_item = pTestingFunctionsInfoCol(OneBasedIndex)
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

Public Property Get item(ByVal OneBasedIndex As Long) As ITest
Attribute item.VB_UserMemId = 0
    Set item = pTestingFunctionsInfoCol(OneBasedIndex)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pTestingFunctionsInfoCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    Set pRecurIterFuncCol = New VBA.Collection
    Set pTestingFunctionsInfoCol = New VBA.Collection
    
    Call initTestingFunctionsInfoObjects
    Call populateDictWithTestFuncInfo
End Sub
