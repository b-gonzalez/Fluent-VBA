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

Public Event TestPassed(ByVal test As ITest)
Public Event TestFailed(ByVal test As ITest)
Public Event TestCompleted(ByVal test As ITest)
Public Event TestUnexpected(ByVal test As ITest)
Public Event DuplicateTest(ByVal test As ITest)

Private pAlgorithm As flAlgorithm
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
Private pTestingFunctionsInfos As cTestingFunctionsInfos
Private pToStrDev As Boolean
Private pTestStrings As ITestStrings

Private Const defaultEpsilon As Double = 0.000001
Private Const INVALID_ENUM_VALUE_ERROR As Long = vbObjectError + 516

Public Property Get result() As Boolean
    result = pResult
End Property

Public Property Get Item(ByVal OneBasedIndex As Long) As ITest
Attribute Item.VB_UserMemId = 0
    Set Item = pCol(OneBasedIndex)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Property

Public Property Let Algorithm(ByVal value As flAlgorithm)
    If value And flAlgorithm.flIterative Or value And flAlgorithm.flRecursive Then
        pAlgorithm = value
    Else
        Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
    End If
    
    pAlgorithmValueSet = True
End Property

Public Property Get Algorithm() As flAlgorithm
    Algorithm = pAlgorithm
End Property

Public Property Get AlgorithmValueSet() As Boolean
    AlgorithmValueSet = pAlgorithmValueSet
End Property

Public Property Let ToStrDev(ByVal value As Boolean)
    pToStrDev = value
End Property

Public Property Get ToStrDev() As Boolean
    ToStrDev = pToStrDev
End Property

Public Property Get Count() As Long
    Count = pCount
End Property

Public Property Let ApproximateEqual(ByVal value As Boolean)
    pApproximateEqual = value
End Property

Public Property Get ApproximateEqual() As Boolean
    ApproximateEqual = pApproximateEqual
End Property

Public Property Let Epsilon(ByVal value As Double)
    pEpsilon = value
End Property

Public Property Get Epsilon() As Double
    Epsilon = pEpsilon
End Property

Public Property Get DataStructures() As VBA.Collection
    Set DataStructures = pDataStructures
End Property

Public Property Let SkipDupCheck(ByVal value As Boolean)
    pSkipDupCheck = value
End Property

Public Property Get SkipDupCheck() As Boolean
    SkipDupCheck = pSkipDupCheck
End Property

Public Property Get TestingFunctionsInfos() As cTestingFunctionsInfos
    Set TestingFunctionsInfos = pTestingFunctionsInfos
End Property

Public Property Get TestStrings() As ITestStrings
    Set TestStrings = pTestStrings
End Property

Public Sub AddDataStructure(ByVal dataStructure As Variant)
    Dim dsType As String

    dsType = VBA.Information.TypeName(dataStructure)

    If IsIterable(dataStructure) And Not datastructureInCollection(dsType, pDataStructures) Then
        pDataStructures.Add dsType
    End If
End Sub

Public Function datastructureInCollection(ByVal dsName As String, ByVal col As VBA.Collection) As Boolean
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

Public Function IsDataStructure(ByVal dataStructure As Variant) As Boolean
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
    
    IsDataStructure = tempBool
End Function

Public Function IsIterable(ByVal dataStructure As Variant) As Boolean
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

Public Sub CheckTest(ByVal test As ITest)
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

Public Sub resetTestingInfo()
    pTestingFunctionsInfos.initTestingFunctionsInfo
End Sub

Public Sub ResetCounter()
    pCount = 0
End Sub

Public Function DatastructureIsSelfReferential(ByVal ds As Variant, ByVal Algorithm As flAlgorithm) As Boolean
    Dim iterFuncCalled As Boolean
    Dim recurFuncCalled As Boolean
    Dim hasSelfReferentialIter As Boolean
    Dim hasSelfReferentialRecur As Boolean
    Dim b As Boolean
    
    If Algorithm And flIterative Then
        hasSelfReferentialIter = datastructureIsSelfReferentialIter(ds)
        iterFuncCalled = True
    End If
    
    If Algorithm And flRecursive Then
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
    
    DatastructureIsSelfReferential = b
        
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
        If IsDataStructure(ds) Then
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
    Dim IsSelfReferential As Boolean

    Set col = New VBA.Collection
    Set tempCol = New VBA.Collection
    Set objPtrCol = New VBA.Collection
    IsSelfReferential = False

    If IsObject(ds) Then objPtrCol.Add VBA.[_HiddenModule].ObjPtr(ds)

    For Each elem In ds
        col.Add elem
    Next elem

    If Not IsSelfReferential Then
        Do
            For Each elem In col
                If IsDataStructure(elem) Then
                    If VBA.Information.IsObject(elem) Then
                        IsSelfReferential = itemInObjPtrCol(elem, objPtrCol)
                        If IsSelfReferential Then Exit Do
                    Else
                        For Each elem2 In elem
                            If VBA.Information.IsObject(elem2) Then
                                IsSelfReferential = itemInObjPtrCol(elem2, objPtrCol)
                                If IsSelfReferential Then Exit Do
                            ElseIf IsDataStructure(elem2) Then
                                tempCol.Add elem2
                            End If
                        Next elem2
                    End If
                Else
                    If IsDataStructure(elem) Then
                        For Each elem2 In elem
                            If VBA.Information.IsObject(elem2) Then
                                objPtrCol.Add VBA.[_HiddenModule].ObjPtr(elem2)
                            End If

                            If IsDataStructure(elem2) Then tempCol.Add elem2
                        Next elem2
                    End If
                End If
            Next elem

            Set col = tempCol
            Set tempCol = New Collection
        Loop While col.Count > 0
    End If

    datastructureIsSelfReferentialIter = IsSelfReferential
End Function

Private Function itemInObjPtrCol(ByVal Item As Variant, ByVal objPtrCol As VBA.Collection) As Boolean
    Dim elem As Variant
    Dim b As Boolean
    Dim objPtrAddress As Variant
    
    b = False
    objPtrAddress = VBA.[_HiddenModule].ObjPtr(Item)

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

Private Sub Class_Initialize()
    Set pCol = New VBA.Collection
    pCount = 0
    
    pAlgorithm = flAlgorithm.flRecursive
    pEpsilon = defaultEpsilon
    
    Set pFluentPathDict = New Scripting.Dictionary
    
    Set pTestingFunctionsInfos = New cTestingFunctionsInfos
    
    Set pTestStrings = New cTestStrings
        
    Set pDataStructures = New VBA.Collection
    
    pDataStructures.Add "Dictionary"
    pDataStructures.Add "Collection"
    pDataStructures.Add "ArrayList"
    
End Sub
