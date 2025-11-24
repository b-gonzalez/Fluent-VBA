VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'AlgorithmValueSet

Implements ITests

Public Event TestPassed(ByVal test As ITest)
Public Event TestFailed(ByVal test As ITest)
Public Event TestCompleted(ByVal test As ITest)
Public Event TestUnexpected(ByVal test As ITest)
Public Event DuplicateTest(ByVal test As ITest)

Private pAlgorithm As Long
Private pAlgorithmValueSet As Boolean
Private pApproximateEqual As Boolean
Private pCol As VBA.Collection
Private pCount As Long
Private pDataStructures As VBA.Collection
Private pEpsilon As Double
Private pFluentPathDict As Scripting.Dictionary
Private pResult As Variant
Private pSkipDupCheck As Boolean
Private pTestDictCounter As Scripting.Dictionary
Private pTestingFunctionsInfos As ITestingFunctionsInfos
Private pToStrDev As Boolean
Private pTestStrings As ITestStrings
Private pContinueWithSelfReferentialIfPossible As Boolean
Private pUtilities As IUtilities

Private Const defaultEpsilon As Double = 0.000001
Private Const INVALID_ENUM_VALUE_ERROR As Long = vbObjectError + 516

Private Property Get ITests_result() As Boolean
    ITests_result = pResult
End Property

Private Property Get ITests_item(ByVal OneBasedIndex As Long) As ITest
    Dim tempCol As Collection
    
    Set tempCol = pUtilities.getElementFromCollection(pCol, OneBasedIndex)
    
    If Not VBA.Information.IsNull(tempCol(1)) Then
        Set ITests_item = tempCol(1)
    End If
End Property

Private Property Get ITests_NewEnum() As IUnknown
    Set ITests_NewEnum = pCol.[_NewEnum]
End Property

Private Property Let ITests_Algorithm(ByVal value As Long)
    'This following if statement uses bitwise flags
    
    If value And flAlgorithm.flIterative Or value And flAlgorithm.flRecursive Then
        pAlgorithm = value
    Else
        Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
    End If
    
    pAlgorithmValueSet = True
End Property

Private Property Get ITests_Algorithm() As Long
    ITests_Algorithm = pAlgorithm
End Property

Private Property Get ITests_AlgorithmValueSet() As Boolean
    ITests_AlgorithmValueSet = pAlgorithmValueSet
End Property

Private Property Let ITests_ToStrDev(ByVal value As Boolean)
    pToStrDev = value
End Property

Private Property Get ITests_ToStrDev() As Boolean
    ITests_ToStrDev = pToStrDev
End Property

Private Property Get ITests_Count() As Long
    ITests_Count = pCount
End Property

Private Property Let ITests_ApproximateEqual(ByVal value As Boolean)
    pApproximateEqual = value
End Property

Private Property Get ITests_ApproximateEqual() As Boolean
    ITests_ApproximateEqual = pApproximateEqual
End Property

Private Property Let ITests_Epsilon(ByVal value As Double)
    pEpsilon = value
End Property

Private Property Get ITests_Epsilon() As Double
    ITests_Epsilon = pEpsilon
End Property

Private Property Get ITests_DataStructures() As VBA.Collection
    Set ITests_DataStructures = pDataStructures
End Property

Private Property Let ITests_SkipDupCheck(ByVal value As Boolean)
    pSkipDupCheck = value
End Property

Private Property Get ITests_SkipDupCheck() As Boolean
    ITests_SkipDupCheck = pSkipDupCheck
End Property

Private Property Get ITests_TestingFunctionsInfos() As ITestingFunctionsInfos
    Set ITests_TestingFunctionsInfos = pTestingFunctionsInfos
End Property

Private Property Get ITests_TestStrings() As ITestStrings
    Set ITests_TestStrings = pTestStrings
End Property

Private Property Let ITests_ContinueWithSelfReferentialIfPossible(value As Boolean)
    pContinueWithSelfReferentialIfPossible = value
End Property

Private Property Get ITests_ContinueWithSelfReferentialIfPossible() As Boolean
    ITests_ContinueWithSelfReferentialIfPossible = pContinueWithSelfReferentialIfPossible
End Property

Private Sub ITests_AddDataStructure(ByVal dataStructure As Variant)
    Dim dsType As String

    dsType = VBA.Information.TypeName(dataStructure)

    If IsIterable(dataStructure) And Not datastructureInCollection(dsType, pDataStructures) Then
        pDataStructures.Add dsType
    End If
End Sub

Private Function datastructureInCollection(ByVal dsName As String, ByVal col As VBA.Collection) As Boolean
    Dim elem As Variant
    Dim b As Boolean
    
    b = False
    
    For Each elem In col
        If elem = dsName Then
            b = True
            Exit For
        End If
    Next elem
    
    datastructureInCollection = b
End Function

Private Function ITests_IsDataStructure(ByVal dataStructure As Variant) As Boolean
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    dsTypeName = VBA.Information.TypeName(dataStructure)
    tempBool = False
    
    If VBA.Information.IsArray(dataStructure) Then
        tempBool = True
    ElseIf datastructureInCollection(dsTypeName, pDataStructures) Then
        tempBool = IsIterable(dataStructure)
    End If
    
    ITests_IsDataStructure = tempBool
End Function

Private Function IsIterable(ByVal dataStructure As Variant) As Boolean
    Dim tempBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    On Error Resume Next
        For Each elem In dataStructure
            Exit For
        Next elem
        
        tempErr = Err.Number
        
        tempBool = (tempErr = 0)
        
    On Error GoTo 0
    
    IsIterable = tempBool
End Function

Private Sub ITests_CheckTest(ByVal test As ITest)
    pResult = test.result
    
    pCol.Add test
    
    pCount = pCount + 1
    
    RaiseEvent TestCompleted(test)
    
    If test.result = True Then
        RaiseEvent TestPassed(test)
    ElseIf test.result = False Then
        RaiseEvent TestFailed(test)
    ElseIf VBA.Information.IsNull(test.result) Or VBA.Information.IsEmpty(test.result) Or test.HasSelfReferential Then
        RaiseEvent TestUnexpected(test)
    End If
    
    If pFluentPathDict.Exists(test.FluentPath) And Not pSkipDupCheck Then
        RaiseEvent DuplicateTest(test)
    Else
        pFluentPathDict(test.FluentPath) = 0
    End If
End Sub

Private Sub ITests_resetTestingInfo()
    pTestingFunctionsInfos.initTestingFunctionsInfo
End Sub

Private Sub ITests_ResetCounter()
    pCount = 0
End Sub

Private Function ITests_DatastructureIsSelfReferential(ByVal ds As Variant, ByVal Algorithm As Long) As Boolean
    Dim iterFuncCalled As Boolean
    Dim recurFuncCalled As Boolean
    Dim hasSelfReferentialIter As Boolean
    Dim hasSelfReferentialRecur As Boolean
    Dim b As Boolean
    
    If Algorithm And flAlgorithm.flIterative Then
        hasSelfReferentialIter = datastructureIsSelfReferentialIter(ds)
        iterFuncCalled = True
    End If
    
    If Algorithm And flAlgorithm.flRecursive Then
        hasSelfReferentialRecur = datastructureIsSelfReferentialRecur(ds)
        recurFuncCalled = True
    End If
    
    If iterFuncCalled And recurFuncCalled Then
        b = hasSelfReferentialIter And hasSelfReferentialRecur
    ElseIf iterFuncCalled Then
        b = hasSelfReferentialIter
    ElseIf recurFuncCalled Then
        b = hasSelfReferentialRecur
    End If
    
    ITests_DatastructureIsSelfReferential = b
        
End Function

Private Function datastructureIsSelfReferentialRecur( _
    ByVal ds As Variant, _
    Optional ByVal objPtrCol As VBA.Collection = Nothing, _
    Optional ByVal objCol As VBA.Collection = Nothing _
) As Boolean
    Dim elem As Variant
    Dim elem2 As Variant
    Dim objPtrAddress As Variant

    If objPtrCol Is Nothing Then
        Set objPtrCol = New VBA.Collection

        If IsObject(ds) Then
            objPtrCol.Add VBA.[_HiddenModule].ObjPtr(ds)
        End If
    End If

    For Each elem In ds
        If ITests_IsDataStructure(ds) Then
            If VBA.Information.IsObject(elem) Then
                objPtrAddress = VBA.[_HiddenModule].ObjPtr(elem)

                For Each elem2 In objPtrCol
                    If objPtrAddress = elem2 Then
                        datastructureIsSelfReferentialRecur = True
                        Exit Function
                    End If
                Next elem2

                If objPtrCol.Count = 0 Then objPtrCol.Add objPtrAddress
            ElseIf VBA.Information.IsArray(elem) Then
                datastructureIsSelfReferentialRecur = datastructureIsSelfReferentialRecur(elem, objPtrCol)
            End If
        End If
    Next elem
End Function

Private Function datastructureIsSelfReferentialIter(ByVal ds As Variant) As Boolean
    Dim elem As Variant
    Dim elem2 As Variant
    Dim col As VBA.Collection
    Dim tempCol As VBA.Collection
    Dim objPtrCol As VBA.Collection
    Dim isSelfReferential As Boolean

    Set col = New VBA.Collection
    Set tempCol = New VBA.Collection
    Set objPtrCol = New VBA.Collection
    isSelfReferential = False

    If IsObject(ds) Then objPtrCol.Add VBA.[_HiddenModule].ObjPtr(ds)

    For Each elem In ds
        col.Add elem
    Next elem

    If Not isSelfReferential Then
        Do
            For Each elem In col
                If ITests_IsDataStructure(elem) Then
                    If VBA.Information.IsObject(elem) Then
                        isSelfReferential = itemInObjPtrCol(elem, objPtrCol)
                        If isSelfReferential Then Exit Do
                    Else
                        For Each elem2 In elem
                            If VBA.Information.IsObject(elem2) Then
                                isSelfReferential = itemInObjPtrCol(elem2, objPtrCol)
                                If isSelfReferential Then Exit Do
                            ElseIf ITests_IsDataStructure(elem2) Then
                                tempCol.Add elem2
                            End If
                        Next elem2
                    End If
                Else
                    If ITests_IsDataStructure(elem) Then
                        For Each elem2 In elem
                            If VBA.Information.IsObject(elem2) Then
                                objPtrCol.Add VBA.[_HiddenModule].ObjPtr(elem2)
                            End If

                            If ITests_IsDataStructure(elem2) Then tempCol.Add elem2
                        Next elem2
                    End If
                End If
            Next elem

            Set col = tempCol
            Set tempCol = New Collection
        Loop While col.Count > 0
    End If

    datastructureIsSelfReferentialIter = isSelfReferential
End Function

Private Function itemInObjPtrCol(ByVal item As Variant, ByVal objPtrCol As VBA.Collection) As Boolean
    Dim elem As Variant
    Dim b As Boolean
    Dim objPtrAddress As Variant
    
    b = False
    objPtrAddress = VBA.[_HiddenModule].ObjPtr(item)

    For Each elem In objPtrCol
         If elem = objPtrAddress Then
             b = True
         End If
     Next elem
     
     itemInObjPtrCol = b
End Function

Private Function elementIsInCollection(ByVal elem As Variant, ByVal col As VBA.Collection) As Boolean
    Dim b As Boolean
    Dim e As Variant
    
    For Each e In col
        If VBA.Information.IsObject(e) Then
            If e Is elem Then
                b = True
                Exit For
            End If
        End If
    Next e
    
    elementIsInCollection = b
End Function

Private Sub ITests_AddFromFluent(ByVal fluent As cFluent)
    Dim test As cTest
    
    For Each test In fluent.Meta.tests
        pCol.Add test
        pCount = pCount + 1
    Next test
End Sub

Private Sub ITests_AddFromFluentOf(ByVal fluentOf As cFluentOf)
    Dim test As cTest
    
    For Each test In fluentOf.Meta.tests
        pCol.Add test
        pCount = pCount + 1
    Next test
End Sub

Private Sub ITests_AddFromTests(ByVal tests As ITests)
    Dim test As ITest
    
    For Each test In tests
        pCol.Add test
        pCount = pCount + 1
    Next test
End Sub

Private Sub Class_Initialize()
    Set pCol = New VBA.Collection
    pCount = 0
    
    pAlgorithm = flAlgorithm.flRecursive
    pEpsilon = defaultEpsilon
    
    Set pFluentPathDict = New Scripting.Dictionary
    Set pTestingFunctionsInfos = New cTestingFunctionsInfos
    Set pTestStrings = New cTestStrings
    Set pDataStructures = New VBA.Collection
    Set pUtilities = New cUtilities
    
    pDataStructures.Add "Dictionary"
    pDataStructures.Add "Collection"
    pDataStructures.Add "ArrayList"
    
    pContinueWithSelfReferentialIfPossible = False
End Sub

Private Sub Class_Terminate()
    Set pFluentPathDict = Nothing
    Set pTestDictCounter = Nothing
    Set pTestingFunctionsInfos = Nothing
    Set pTestStrings = Nothing
    Set pDataStructures = Nothing
    Set pCol = Nothing
    Set pUtilities = Nothing
End Sub
