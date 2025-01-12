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

Public Property Get Item(ByVal OneBasedIndex As Long) As ITest
Attribute Item.VB_UserMemId = 0
    Set Item = pTestingFunctionsInfoCol(OneBasedIndex)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pTestingFunctionsInfoCol.[_NewEnum]
End Property

Public Property Get Alphabetic() As ITestingFunctionsInfo
    Set Alphabetic = pAlphabetic
End Property

Public Property Get Alphanumeric() As ITestingFunctionsInfo
    Set Alphanumeric = pAlphanumeric
End Property

Public Property Get Between() As ITestingFunctionsInfo
    Set Between = pBetween
End Property

Public Property Get Contain() As ITestingFunctionsInfo
    Set Contain = pContain
End Property

Public Property Get DepthCountOf() As ITestingFunctionsInfo
    Set DepthCountOf = pDepthCountOf
End Property

Public Property Get Elements() As ITestingFunctionsInfo
    Set Elements = pElements
End Property

Public Property Get ElementsInDataStructure() As ITestingFunctionsInfo
    Set ElementsInDataStructure = pElementsInDataStructure
End Property

Public Property Get EndWith() As ITestingFunctionsInfo
    Set EndWith = pEndWith
End Property

Public Property Get EqualTo() As ITestingFunctionsInfo
    Set EqualTo = pEqualTo
End Property

Public Property Get Erroneous() As ITestingFunctionsInfo
    Set Erroneous = pErroneous
End Property

Public Property Get ErrorDescriptionOf() As ITestingFunctionsInfo
    Set ErrorDescriptionOf = pErrorDescriptionOf
End Property

Public Property Get ErrorNumberOf() As ITestingFunctionsInfo
    Set ErrorNumberOf = pErrorNumberOf
End Property

Public Property Get EvaluateTo() As ITestingFunctionsInfo
    Set EvaluateTo = pEvaluateTo
End Property

Public Property Get ExactSameElementsAs() As ITestingFunctionsInfo
    Set ExactSameElementsAs = pExactSameElementsAs
End Property

Public Property Get GreaterThan() As ITestingFunctionsInfo
    Set GreaterThan = pGreaterThan
End Property

Public Property Get GreaterThanOrEqualTo() As ITestingFunctionsInfo
    Set GreaterThanOrEqualTo = pGreaterThanOrEqualTo
End Property

Public Property Get IdenticalTo() As ITestingFunctionsInfo
    Set IdenticalTo = pIdenticalTo
End Property

Public Property Get InDataStructure() As ITestingFunctionsInfo
    Set InDataStructure = pInDataStructure
End Property

Public Property Get InDataStructures() As ITestingFunctionsInfo
    Set InDataStructures = pInDataStructures
End Property

Public Property Get LengthBetween() As ITestingFunctionsInfo
    Set LengthBetween = pLengthBetween
End Property

Public Property Get LengthOf() As ITestingFunctionsInfo
    Set LengthOf = pLengthOf
End Property

Public Property Get LessThan() As ITestingFunctionsInfo
    Set LessThan = pLessThan
End Property

Public Property Get LessThanOrEqualTo() As ITestingFunctionsInfo
    Set LessThanOrEqualTo = pLessThanOrEqualTo
End Property

Public Property Get MaxLengthOf() As ITestingFunctionsInfo
    Set MaxLengthOf = pMaxLengthOf
End Property

Public Property Get MinLengthOf() As ITestingFunctionsInfo
    Set MinLengthOf = pMinLengthOf
End Property

Public Property Get NestedCountOf() As ITestingFunctionsInfo
    Set NestedCountOf = pNestedCountOf
End Property

Public Property Get Numeric() As ITestingFunctionsInfo
    Set Numeric = pNumeric
End Property

Public Property Get OneOf() As ITestingFunctionsInfo
    Set OneOf = pOneOf
End Property

Public Property Get Procedure() As ITestingFunctionsInfo
    Set Procedure = pProcedure
End Property

Public Property Get SameElementsAs() As ITestingFunctionsInfo
    Set SameElementsAs = pSameElementsAs
End Property

Public Property Get SameTypeAs() As ITestingFunctionsInfo
    Set SameTypeAs = pSameTypeAs
End Property

Public Property Get SameUniqueElementsAs() As ITestingFunctionsInfo
    Set SameUniqueElementsAs = pSameUniqueElementsAs
End Property

Public Property Get Something() As ITestingFunctionsInfo
    Set Something = pSomething
End Property

Public Property Get StartWith() As ITestingFunctionsInfo
    Set StartWith = pStartWith
End Property

Public Property Let AddDevTestingFunctionsInfo(value As Boolean)
    AddDevTestingFunctionsInfo = value
End Property

Public Property Get AddDevTestingFunctionsInfo() As Boolean
    AddDevTestingFunctionsInfo = pAddDevTestingFunctionsInfo
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

Public Function getTestingFunction(ByVal functionName As String) As ITestingFunctionsInfo
    Set getTestingFunction = pTestingFunctionsDict(functionName)
End Function

Public Sub populateDictWithTestFuncInfo()
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

Public Sub initTestingFunctionsInfo()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    Call initTestingFunctionsInfoObjects
    Call populateDictWithTestFuncInfo

End Sub

Public Sub PrintTestFunInfo()
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

Public Function TestFuncInfoToDict() As Scripting.Dictionary
    Set TestFuncInfoToDict = pTestingFunctionsDict
End Function

Public Function TestFuncInfoDictsAreIdentical(ByVal tfiDict1 As Scripting.Dictionary, ByVal tfiDict2 As Scripting.Dictionary) As Boolean
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
    
    TestFuncInfoDictsAreIdentical = b
End Function

Public Function validateTfiDictCounters(ByVal tfiDict As Scripting.Dictionary, Optional ByVal counter As Long = 0) As Boolean
    Dim b As Boolean
    Dim elem As Variant
    Dim testFuncSubInfo As ITestingFunctionsInfo
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
    validateTfiDictCounters = b
End Function

Private Sub initTestingFunctionsInfoObjects()
    Set pRecurIterFuncCol = New VBA.Collection
    Set pTestingFunctionsInfoCol = New VBA.Collection
End Sub

Public Sub AddTestingFunctionsInfo(value As ITestingFunctionsInfo)
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

Private Sub Class_Initialize()
    Set pTestingFunctionsDict = New Scripting.Dictionary
    Set pRecurIterFuncCol = New VBA.Collection
    Set pTestingFunctionsInfoCol = New VBA.Collection
    
    Call initTestingFunctionsInfoObjects
    Call populateDictWithTestFuncInfo
End Sub
