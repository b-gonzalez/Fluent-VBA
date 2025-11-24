VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestingFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestingFunctions
Implements IShouldTestFuncs
Implements IHaveTestFuncs
Implements IBeTestFuncs

Private pMeta As IMeta
Private pTests As cTests
Private pPrinting As cPrinting
Private iPrint As IPrintExpr
Private pApproximateEqual As Boolean
Private pEpsilon As Double

Private pTestDictCounter As Scripting.Dictionary

Private pTestValue As Variant
Private pTestValueSet As Boolean
Private pFluentPath As String
Private pExpressions As IExpressions
Private pTestValueClean As String

Private Const INVALID_ENUM_VALUE_ERROR As Long = vbObjectError + 516
Private Const RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR As Long = vbObjectError + 517

Private Function inputIsValid(ByVal inputValue As Variant, ByVal inputConditions As Long) As Boolean
    Dim b As Boolean
    Dim col As VBA.Collection
    Dim elem As Variant
    Dim testTypeCount As Long
    Dim boolCount As Long
    Dim HasSelfReferential As Boolean
    
    b = False
    Set col = New VBA.Collection
    testTypeCount = 0
    
    If pTestValueSet Then
        If inputConditions And flTestType.flIsDataStructure Then
            If pMeta.tests.IsDataStructure(inputValue) And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsNotDataStructure Then
            If Not pMeta.tests.IsDataStructure(inputValue) And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsNumeric Then
            If VBA.Information.IsNumeric(inputValue) And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsObject Then
            If VBA.Information.IsObject(inputValue) Then
                If pMeta.tests.IsDataStructure(inputValue) Then
                    HasSelfReferential = pMeta.tests.DatastructureIsSelfReferential(inputValue, pMeta.tests.Algorithm)
                    
                    If Not HasSelfReferential Or pMeta.tests.ContinueWithSelfReferentialIfPossible Then col.Add True
                Else
                    If VBA.Information.IsObject(inputValue) And pTestValueSet Then col.Add True
                End If
            End If
            
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsNotObject Then
            If Not VBA.Information.IsObject(inputValue) And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
    
        If inputConditions And flTestType.flIsNotNull Then
            If Not VBA.Information.IsNull(inputValue) And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsErrObject Then
            If VBA.Information.TypeName(inputValue) = "ErrObject" And pTestValueSet Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsString Then
            If VBA.Information.TypeName(inputValue) = "String" Then col.Add True
            testTypeCount = testTypeCount + 1
        End If
        
        If inputConditions And flTestType.flIsNotSelfReferential Then
            If pMeta.tests.IsDataStructure(inputValue) Then
                HasSelfReferential = pMeta.tests.DatastructureIsSelfReferential(inputValue, pMeta.tests.Algorithm)
                If Not HasSelfReferential Then col.Add True
                testTypeCount = testTypeCount + 1
            End If
        End If
    End If
    
    inputIsValid = ((testTypeCount = col.Count) And pTestValueSet)
End Function

Private Property Let ITestingFunctions_TestValue(ByVal value As Variant)
    pTestValue = value
    pTestValueSet = True
End Property

Private Property Set ITestingFunctions_TestValue(ByVal value As Variant)
    Set pTestValue = value
    pTestValueSet = True
End Property

Private Property Get ITestingFunctions_TestValue() As Variant
    If VBA.Information.IsObject(pTestValue) Then
        Set ITestingFunctions_TestValue = pTestValue
    Else
        ITestingFunctions_TestValue = pTestValue
    End If
End Property

Private Property Let ITestingFunctions_TestValueClean(ByVal value As Variant)
    pTestValueClean = pMeta.tests.TestStrings.CleanString(value)
End Property

Private Property Let ITestingFunctions_TestValueSet(ByVal value As Boolean)
    pTestValueSet = value
End Property

Private Function ITestingFunctions_SetDefaultFuncVal() As Variant
    Dim val As Variant
    
    If pTestValueSet Then
        val = Null
    Else
        val = Empty
    End If
        
    ITestingFunctions_SetDefaultFuncVal = val
End Function

Private Property Set ITestingFunctions_setMeta(ByVal value As cMeta)
    Set pMeta = value
End Property

Private Property Set ITestingFunctions_Expressions(ByVal value As IExpressions)
    Set pExpressions = value
End Property

Private Function ITestingFunctions_GetCleanedString(ByVal value As Variant) As String
    Dim tempVal As String
    
    tempVal = value
    
    tempVal = VBA.Strings.Replace(tempVal, " ", "")
    
    If VBA.Strings.InStr(1, tempVal, """") = 1 Then
        tempVal = VBA.Strings.Replace(tempVal, """", "")
    End If
    
    ITestingFunctions_GetCleanedString = tempVal
End Function

Private Function ITestingFunctions_getTempTestValue(ByVal testingValue As Variant) As Variant
    Dim tempTestValue As Variant
    
    If Not VBA.Information.IsObject(testingValue) Then
        tempTestValue = testingValue
        
        With pMeta.tests.TestStrings
            If .CleanTestValueStr Or .CleanTestStrings Then
                tempTestValue = pTestValueClean
            End If
        End With
        
        ITestingFunctions_getTempTestValue = tempTestValue
    Else
        Set ITestingFunctions_getTempTestValue = testingValue
    End If
End Function

Private Function ITestingFunctions_getTempTestInput(ByVal testingInput As Variant) As Variant
    Dim tempTestingInput As Variant
    
    If Not VBA.Information.IsObject(testingInput) Then
        tempTestingInput = testingInput
        
        With pMeta.tests.TestStrings
            If .CleanTestInputStr Or .CleanTestStrings Then
                tempTestingInput = pMeta.tests.TestStrings.CleanString(testingInput)
            End If
        End With
        
        ITestingFunctions_getTempTestInput = tempTestingInput
    Else
        Set ITestingFunctions_getTempTestInput = testingInput
    End If
End Function

Private Function IBeTestFuncs_InDataStructures( _
    ByVal testingValue As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingInput As Variant _
) As Variant
    Dim col As VBA.Collection
    Dim i As Long
    Dim funcVal As Variant
    Dim method As Long
    Dim recurCol As VBA.Collection
    Dim iterCol As VBA.Collection
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim tempTestingInput As VBA.Collection
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.tests.Algorithm
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure)
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    
    If IsDataStructure And testingInputIsNotSelfReferential Then
'        For i = LBound(testingInput) To UBound(testingInput)
'            If inputIsValid(testingInput(i), flTestType.flIsDataStructure + flTestType.flIsNotSelfReferential) Then
'
'                If method And flAlgorithm.flRecursive Then
'                   Set recurCol = getNestedElementsRecur(testingInput(i))
'                End If
'
'                If method And flAlgorithm.flIterative Then
'                    Set iterCol = getNestedElementsIter(testingInput(i))
'                End If
'
'                If (Not recurCol Is Nothing And Not iterCol Is Nothing) Then
'                    If IBeTestFuncs_IdenticalTo(recurCol, negateValue, iterCol, False) Then
'                        Set col = recurCol
'                    Else
'                        Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative collections are not identical"
'                    End If
'                Else
'                    If Not recurCol Is Nothing Then
'                        Set col = getNestedElementsRecur(testingInput)
'                    ElseIf Not iterCol Is Nothing Then
'                        Set col = getNestedElementsIter(testingInput)
'                    Else
'                        Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
'                    End If
'                End If
'
'                If Not col Is Nothing Then
'                    funcVal = IBeTestFuncs_InDataStructure(testingValue, col, negateValue, updateFluentPath:=False)
'                Else
'                    Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
'                End If
'            End If
'        Next i
        Set tempTestingInput = getElementsFromDsUpToFirstMissing(testingInput)
        
        For i = 1 To tempTestingInput.Count
            If inputIsValid(tempTestingInput(i), flTestType.flIsDataStructure + flTestType.flIsNotSelfReferential) Then
            
                If method And flAlgorithm.flRecursive Then
                   Set recurCol = getNestedElementsRecur(tempTestingInput(i))
                End If
                
                If method And flAlgorithm.flIterative Then
                    Set iterCol = getNestedElementsIter(tempTestingInput(i))
                End If
                
                If (Not recurCol Is Nothing And Not iterCol Is Nothing) Then
                    If IBeTestFuncs_IdenticalTo(recurCol, negateValue, iterCol, False) Then
                        Set col = recurCol
                    Else
                        Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative collections are not identical"
                    End If
                Else
                    If Not recurCol Is Nothing Then
                        Set col = getNestedElementsRecur(tempTestingInput)
                    ElseIf Not iterCol Is Nothing Then
                        Set col = getNestedElementsIter(tempTestingInput)
                    Else
                        Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
                    End If
                End If
                    
                If Not col Is Nothing Then
                    funcVal = IBeTestFuncs_InDataStructure(testingValue, col, negateValue, updateFluentPath:=False)
                Else
                    Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
                End If
            End If
        Next i
    End If
    
    IBeTestFuncs_InDataStructures = funcVal
    
    Call pExpressions.addToTestAndFluentPath("InDataStructures", funcVal, VBA.Conversion.CBool(negateValue), testingValue, tempTestingInput, recurIterFunc:=True, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)
    
End Function

Private Function IBeTestFuncs_InDataStructure( _
    ByVal testingValue As Variant, _
    ByVal dataStructure As Variant, _
    Optional ByVal negateValue As Boolean = False, _
    Optional ByVal updateFluentPath As Boolean = False _
) As Variant
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    Dim funcVal As Variant
    Dim method As Long
    Dim recurCol As VBA.Collection
    Dim iterCol As VBA.Collection
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.tests.Algorithm
    
    IsDataStructure = inputIsValid(dataStructure, flTestType.flIsDataStructure)
    testingInputIsNotSelfReferential = inputIsValid(dataStructure, flTestType.flIsNotSelfReferential)
        
    If IsDataStructure And testingInputIsNotSelfReferential Then
        tempBool = False
        
        If method And flAlgorithm.flRecursive Then
           Set recurCol = getNestedElementsRecur(dataStructure)
        End If
        
        If method And flAlgorithm.flIterative Then
            Set iterCol = getNestedElementsIter(dataStructure)
        End If
        
        If (Not recurCol Is Nothing And Not iterCol Is Nothing) Then
            If IBeTestFuncs_IdenticalTo(recurCol, negateValue, iterCol, False) Then
                Set col = recurCol
            Else
                Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative collections are not identical"
            End If
        Else
            If Not recurCol Is Nothing Then
                Set col = recurCol
            ElseIf Not iterCol Is Nothing Then
                Set col = iterCol
            Else
                Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
            End If
        End If
        
        If Not col Is Nothing Then
            For Each elem In col
                If testingValue = elem Then
                    tempBool = True
                    Exit For
                End If
            Next elem
            
            funcVal = tempBool
        Else
            Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
        End If
    End If
    
    IBeTestFuncs_InDataStructure = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("InDataStructure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, dataStructure, recurIterFunc:=True, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)
    End If
    
End Function

Private Function getNestedElementsRecur(ByVal dataStructure As Variant, Optional ByRef col As VBA.Collection = Nothing) As VBA.Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New VBA.Collection
    End If
    
    If pMeta.tests.IsDataStructure(dataStructure) Then
        If TypeOf dataStructure Is Scripting.Dictionary Then
            tempDataStructure = dataStructure.Items
        Else
            If VBA.Information.IsObject(dataStructure) Then
                Set tempDataStructure = dataStructure
            Else
                tempDataStructure = dataStructure
            End If
        End If
        
        For Each elem In tempDataStructure
            If pMeta.tests.IsDataStructure(elem) Then
                Set col = getNestedElementsRecur(elem, col)
            Else
                col.Add elem
            End If
        Next elem
    Else
        col.Add elem
    End If
    
    Set getNestedElementsRecur = col
End Function

Private Function getNestedElementsIter(ByVal dataStructure As Variant) As VBA.Collection
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim tempCol As Variant
    Dim tempDataStructure As Variant
    Dim i As Long
    Dim hasDataStructures As Boolean

    Set col = New VBA.Collection

    If TypeOf dataStructure Is Scripting.Dictionary Then
        tempDataStructure = dataStructure.Items
    Else
        If VBA.Information.IsObject(dataStructure) Then
            Set tempDataStructure = dataStructure
        Else
            tempDataStructure = dataStructure
        End If
    End If

    For Each elem In tempDataStructure
        col.Add elem
    Next elem
    
    Do
        hasDataStructures = False
        
        For i = 1 To col.Count
            If pMeta.tests.IsDataStructure(col(i)) Then
                If pExpressions.getDsCount(col(i)) > 0 Then
                    hasDataStructures = True
                    
                    If TypeOf col(i) Is Scripting.Dictionary Then
                        tempCol = col(i).Items
                    ElseIf IsObject(col(i)) Then
                        Set tempCol = col(i)
                    Else
                        tempCol = col(i)
                    End If
                    
                    For Each elem In tempCol
                        col.Add elem
                    Next elem
                    
                    col.Remove i
                    Exit For
                Else
                    col.Remove i
                End If
            End If
        Next i
    Loop While hasDataStructures = True

    Set getNestedElementsIter = col
End Function

Private Function IBeTestFuncs_GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If inputIsValid(testingValue, flTestType.flIsNumeric) And inputIsValid(testingInput, flTestType.flIsNumeric) Then
        funcVal = (testingValue > testingInput)
    End If
    
    IBeTestFuncs_GreaterThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("GreaterThan", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_GreaterThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNumeric) And inputIsValid(testingInput, flTestType.flIsNumeric) Then
        funcVal = (testingValue >= testingInput)
    End If
    
    IBeTestFuncs_GreaterThanOrEqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("GreaterThanOrEqualTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IBeTestFuncs_LessThan( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If inputIsValid(testingValue, flTestType.flIsNumeric) And inputIsValid(testingInput, flTestType.flIsNumeric) Then
        funcVal = (testingValue < testingInput)
    End If
    
    IBeTestFuncs_LessThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LessThan", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)

End Function

Private Function IBeTestFuncs_LessThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNumeric) And inputIsValid(testingInput, flTestType.flIsNumeric) Then
        funcVal = (testingValue <= testingInput)
    End If
    
    IBeTestFuncs_LessThanOrEqualTo = funcVal
       
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("LessThanOrEqualTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

'Original function found here:
'https://stackoverflow.com/a/72108415/16864869

Private Function Roughly(ByVal a As Double, ByVal b As Double, Optional ByVal within As Double = 0.00001) As Boolean
    Dim d As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim final As Boolean
    
    Const TINY As Double = 1.17549435E-38    'SINGLE_MIN
    
    final = False
    
    x = VBA.Math.Abs(a)
    y = VBA.Math.Abs(b)
    d = VBA.Math.Abs(a - b)
    
    If a = b Then
        final = True
    ElseIf a <> 0 Then
        If b <> 0 Then
            z = x + y
            
            If z > TINY Then
                final = d / z < within
            Else
                final = d < within * TINY
            End If
        End If
    End If
        
    Roughly = final
End Function

Private Function IBeTestFuncs_EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim tempBool As Boolean
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    tempBool = False
    
    If inputIsValid(testingValue, flTestType.flIsNotObject + flTestType.flIsNotDataStructure) And _
    inputIsValid(testingInput, flTestType.flIsNotObject + flTestType.flIsNotDataStructure) Then
    
        If VBA.Information.IsNull(testingValue) Or VBA.Information.IsNull(testingInput) Then
            tempBool = (VBA.Information.TypeName(testingValue) = VBA.Information.TypeName(testingInput))
            
            testingValue = VBA.Information.TypeName(testingValue)
            testingInput = VBA.Information.TypeName(testingInput)
        Else
            If VBA.Information.TypeName(testingValue) = "String" And VBA.Information.TypeName(testingInput) = "String" Then
                tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
                
                tempBool = (tempTestValue = tempTestInput)
            Else
                tempBool = (testingValue = testingInput)
            End If
            
            If pMeta.tests.ApproximateEqual Then
                If VBA.Information.TypeName(testingValue) <> VBA.Information.TypeName(testingInput) Then
                    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) Then
                        tempBool = Roughly(CDbl(testingValue), VBA.Conversion.CDbl(testingInput), pMeta.tests.Epsilon)
                    ElseIf VBA.Information.VarType(testingValue) = vbString Or VBA.Information.VarType(testingInput) = vbString Then
                        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
                        tempBool = (VBA.Strings.StrComp(VBA.Conversion.CStr(tempTestValue), VBA.Conversion.CStr(tempTestInput), vbTextCompare) = 0)
                    End If
                End If
            End If
        End If
        
        funcVal = tempBool
    End If
    
    IBeTestFuncs_EqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("EqualTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
    End If
        
End Function

Private Function IShouldTestFuncs_Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        funcVal = CStr(tempTestValue) Like "*" & CStr(tempTestInput) & "*"
    End If
    
    IShouldTestFuncs_Contain = funcVal
        
    Call pExpressions.addToTestAndFluentPath("Contain", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
        
End Function

Private Function IShouldTestFuncs_StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        valLength = VBA.Strings.Len(tempTestInput)
        funcVal = VBA.Strings.Left$(tempTestValue, valLength) = VBA.Conversion.CStr(tempTestInput)
    End If
    
    IShouldTestFuncs_StartWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("StartWith", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
        
End Function

Private Function IShouldTestFuncs_EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        valLength = VBA.Strings.Len(tempTestInput)
        funcVal = VBA.Strings.Right$(tempTestValue, valLength) = VBA.Conversion.CStr(tempTestInput)
    End If
    
    IShouldTestFuncs_EndWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EndWith", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
    
End Function

Private Function IHaveTestFuncs_LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = (VBA.Strings.Len(VBA.Conversion.CStr(testingValue)) = testingInput)
    End If
    
    IHaveTestFuncs_LengthOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LengthOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
            
End Function

Private Function IHaveTestFuncs_MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = (VBA.Strings.Len(VBA.Conversion.CStr(tempTestValue)) <= testingInput)
    End If
    
    IHaveTestFuncs_MaxLengthOf = funcVal
    
    If updateFluentPath Then
        If tempTestValue <> "" Then
            Call pExpressions.addToTestAndFluentPath("MaxLengthOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
        Else
            Call pExpressions.addToTestAndFluentPath("MaxLengthOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
        End If
    End If
            
End Function

Private Function IHaveTestFuncs_MinLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        funcVal = (VBA.Strings.Len(VBA.Conversion.CStr(tempTestValue)) >= testingInput)
    End If
    
    IHaveTestFuncs_MinLengthOf = funcVal
        
    If updateFluentPath Then
        If tempTestValue <> "" Then
            Call pExpressions.addToTestAndFluentPath("MinLengthOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
        Else
            Call pExpressions.addToTestAndFluentPath("MinLengthOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
        End If
    End If
        
End Function

Private Function IBeTestFuncs_Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim testingValueIsNotSelfReferential As Boolean
    Dim IsDataStructure As Boolean
    Dim IsObj As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsObj = inputIsValid(testingValue, flTestType.flIsObject)
    
    If IsObj Then
        If pMeta.tests.IsDataStructure(testingValue) Then
            IsDataStructure = inputIsValid(testingValue, flTestType.flIsDataStructure)
            testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
            
            If testingValueIsNotSelfReferential And IsDataStructure Then
                funcVal = (Not testingValue Is Nothing)
            End If
        Else
            funcVal = (Not testingValue Is Nothing)
        End If
    End If
    
    IBeTestFuncs_Something = funcVal
    
    If testingValueIsNotSelfReferential And IsDataStructure Then
        Call pExpressions.addToTestAndFluentPath("Something", funcVal, VBA.Conversion.CBool(negateValue), testingValue, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential)
    Else
        Call pExpressions.addToTestAndFluentPath("Something", funcVal, VBA.Conversion.CBool(negateValue), testingValue)
    End If
        
End Function

Private Function IBeTestFuncs_Between( _
ByVal testingValue As Variant, _
ByVal lowerVal As Double, _
ByVal higherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNumeric) Then
        funcVal = _
        (IBeTestFuncs_GreaterThanOrEqualTo(testingValue, lowerVal, updateFluentPath:=False) And _
        IBeTestFuncs_LessThanOrEqualTo(testingValue, higherVal, updateFluentPath:=False))
    End If
    
    IBeTestFuncs_Between = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Between", funcVal, VBA.Conversion.CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal)
End Function

Private Function IBeTestFuncs_OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim tempBool As Boolean
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As VBA.Collection
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If pTestValueSet Then
        If VBA.Information.TypeName(testingValue) = "String" Then
            tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        End If
        
        Set tempTestInput = getElementsFromDsUpToFirstMissing(testingInput)
    
        tempBool = False
        
'        For i = LBound(testingInput) To UBound(testingInput)
'            If VBA.Information.IsObject(testingInput(i)) Or VBA.Information.IsObject(testingValue) Then
'                If VBA.Information.TypeName(testingInput(i)) = VBA.Information.TypeName(testingValue) Then
'                    tempBool = True
'                End If
'            ElseIf testingInput(i) = testingValue Then
'                tempBool = True
'            ElseIf VBA.Information.IsNull(testingInput(i)) Or VBA.Information.IsNull(testingValue) Then
'                If VBA.Information.TypeName(testingInput(i)) = VBA.Information.TypeName(testingValue) Then
'                    tempBool = True
'                End If
'            End If
'
'            If tempBool Then Exit For
'        Next i

        For i = 1 To tempTestInput.Count
            If VBA.Information.IsObject(tempTestInput(i)) Or VBA.Information.IsObject(testingValue) Then
                If VBA.Information.TypeName(tempTestInput(i)) = VBA.Information.TypeName(testingValue) Then
                    tempBool = True
                End If
            ElseIf tempTestInput(i) = testingValue Then
                tempBool = True
            ElseIf VBA.Information.IsNull(tempTestInput(i)) Or VBA.Information.IsNull(testingValue) Then
                If VBA.Information.TypeName(tempTestInput(i)) = VBA.Information.TypeName(testingValue) Then
                    tempBool = True
                End If
            End If
            
            If tempBool Then Exit For
        Next i
        
        funcVal = tempBool
    End If
        
    IBeTestFuncs_OneOf = funcVal
    
    If TypeOf testingValue Is Object  Then testingValue = VBA.Information.TypeName(testingValue)
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, tempTestInput, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, tempTestInput)
    End If
        
End Function

Private Function IHaveTestFuncs_LengthBetween( _
ByVal testingValue As Variant, _
ByVal lowerVal As Double, _
ByVal higherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim tempTestVal As Double
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = _
        (IHaveTestFuncs_MinLengthOf(tempTestValue, lowerVal, updateFluentPath:=False) And _
        IHaveTestFuncs_MaxLengthOf(tempTestValue, higherVal, updateFluentPath:=False))
    End If
    
    IHaveTestFuncs_LengthBetween = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("LengthBetween", funcVal, VBA.Conversion.CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("LengthBetween", funcVal, VBA.Conversion.CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal)
    End If

End Function

Private Function IShouldTestFuncs_EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal

    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
        
        funcVal = (Excel.Evaluate(tempTestValue) = tempTestInput)
    End If
    
    IShouldTestFuncs_EvaluateTo = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EvaluateTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
    
End Function

Private Function IBeTestFuncs_Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        For i = 1 To VBA.Strings.Len(tempTestValue)
            tempChar = VBA.Strings.Mid$(tempTestValue, i, 1)
            
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            End If
        Next i
        
        funcVal = (tempTestValue = tempAlph)
    End If
    
    IBeTestFuncs_Alphabetic = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Alphabetic", funcVal, VBA.Conversion.CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Alphabetic", funcVal, VBA.Conversion.CBool(negateValue), testingValue)
    End If
End Function

Private Function IBeTestFuncs_Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        For i = 1 To VBA.Strings.Len(tempTestValue)
            tempChar = VBA.Strings.Mid$(tempTestValue, i, 1)
            If tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        funcVal = (tempTestValue = tempNum)
    End If
    
    IBeTestFuncs_Numeric = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Numeric", funcVal, VBA.Conversion.CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Numeric", funcVal, VBA.Conversion.CBool(negateValue), testingValue)
    End If
     
End Function

Private Function IBeTestFuncs_Alphanumeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    b = False
    
    If inputIsValid(testingValue, flTestType.flIsNotDataStructure + flTestType.flIsNotNull) Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        For i = 1 To VBA.Strings.Len(tempTestValue)
            tempChar = VBA.Strings.Mid$(tempTestValue, i, 1)
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            ElseIf tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        If VBA.Strings.Len(tempAlph) > 1 And VBA.Strings.Len(tempNum) > 1 Then
            If VBA.Strings.Len(tempAlph) + VBA.Strings.Len(tempNum) = VBA.Strings.Len(tempTestValue) Then
                b = True
            End If
        End If
        
        funcVal = b
    End If
    
    IBeTestFuncs_Alphanumeric = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Alphanumeric", funcVal, VBA.Conversion.CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Alphanumeric", funcVal, VBA.Conversion.CBool(negateValue), testingValue)
    End If
    
End Function

Private Function IBeTestFuncs_Erroneous(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim temp As Variant
    Dim tempErrType As Variant
    Dim tempErrObj As VBA.ErrObject
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    temp = False
    
    If inputIsValid(testingValue, flTestType.flIsErrObject) Or inputIsValid(testingValue, flTestType.flIsString) Then
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            temp = (tempErrObj.Number <> 0)
        ElseIf VBA.Information.TypeName(testingValue) = "String" Then
            On Error Resume Next
                tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                temp = Excel.Evaluate(tempTestValue)
                temp = (VBA.Information.TypeName(temp) = "Error")
            On Error GoTo 0
        End If
        
        funcVal = temp
        
        If TypeOf testingValue Is Object  Then
            tempErrType = VBA.Information.TypeName(testingValue)
        Else
            tempErrType = testingValue
        End If
    End If
    
    IBeTestFuncs_Erroneous = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Erroneous", funcVal, VBA.Conversion.CBool(negateValue), tempErrType, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Erroneous", funcVal, VBA.Conversion.CBool(negateValue), tempErrType)
    End If
End Function

Private Function IHaveTestFuncs_ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim temp As Variant
    Dim tempErrObj As VBA.ErrObject
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsErrObject) Or inputIsValid(testingValue, flTestType.flIsString) Then
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            temp = tempErrObj.Number
        ElseIf VBA.Information.TypeName(testingValue) = "String" Then
            On Error Resume Next
                tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                temp = Excel.Evaluate(tempTestValue)
                Call Err.Raise(testingInput)
                temp = Err.Number
            On Error GoTo 0
        End If
        
        funcVal = (temp = testingInput)
    End If
    
    IHaveTestFuncs_ErrorNumberOf = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    End If
End Function

Private Function IHaveTestFuncs_ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim tempTestVal As Variant
    Dim tempErrObj As VBA.ErrObject
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If inputIsValid(testingValue, flTestType.flIsErrObject) Or inputIsValid(testingValue, flTestType.flIsString) Then
            tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            tempTestVal = tempErrObj.Description
        ElseIf VBA.Information.TypeName(testingValue) = "String" Then
            tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
            
            On Error Resume Next
                tempTestVal = VBA.Conversion.CStr(Excel.Evaluate(tempTestValue))
                tempTestVal = VBA.Strings.Replace(tempTestVal, "Error ", "")
                Err.Raise tempTestVal
                tempTestVal = Err.Description
            On Error GoTo 0
        End If
        
        funcVal = (tempTestVal = tempTestInput)
    End If
    
    IHaveTestFuncs_ErrorDescriptionOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ErrorDescriptionOf", funcVal, VBA.Conversion.CBool(negateValue), tempTestVal, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
End Function

Private Function IHaveTestFuncs_SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim t1 As Variant
    Dim t2 As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If pTestValueSet Then
        If Not VBA.Information.IsNull(testingValue) Then
            t1 = VBA.Information.TypeName(testingValue)
        Else
            t1 = testingValue
        End If
        
        If Not VBA.Information.IsNull(testingInput) Then
            t2 = VBA.Information.TypeName(testingInput)
        Else
            t2 = testingInput
        End If
        
        funcVal = IBeTestFuncs_EqualTo(t1, t2, updateFluentPath:=False)
    End If
    
    IHaveTestFuncs_SameTypeAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameTypeAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
End Function

Private Function IBeTestFuncs_IdenticalTo( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant, _
ByVal updateFluentPath As Boolean) As Variant
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure) And inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingInputIsNotSelfReferential And testingValueIsNotSelfReferential And IsDataStructure Then

        funcVal = _
        IBeTestFuncs_EqualTo( _
            pExpressions.InputToString(testingValue), _
            pExpressions.InputToString(testingInput), _
            updateFluentPath:=False _
        )
    End If
    
    IBeTestFuncs_IdenticalTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("IdenticalTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential)
    End If

End Function

Private Function IHaveTestFuncs_ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim col2 As VBA.Collection
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure) And inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingInputIsNotSelfReferential And testingValueIsNotSelfReferential And IsDataStructure Then
        Set col = New VBA.Collection
        Set col2 = New VBA.Collection
        
        For Each elem In testingValue
            col.Add elem
        Next elem
        
        For Each elem In testingInput
            col2.Add elem
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(col, negateValue, col2, False)
    End If
    
    IHaveTestFuncs_ExactSameElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ExactSameElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)

End Function

Private Function IHaveTestFuncs_SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim tempStr As String
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure) And inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingInputIsNotSelfReferential And testingValueIsNotSelfReferential And IsDataStructure Then
        Set d = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = pExpressions.InputToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = pExpressions.InputToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            Else
                d(tempStr) = tempStr
            End If
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(d.Keys, negateValue, d.Items, False)
    End If
    
    IHaveTestFuncs_SameUniqueElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameUniqueElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)

End Function

Private Function IHaveTestFuncs_SameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim tempStr As String
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure) And inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingInputIsNotSelfReferential And testingValueIsNotSelfReferential And IsDataStructure Then
        i = 1
        elem = Empty
        Set d = New Scripting.Dictionary
        Set d2 = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = pExpressions.InputToString(elem)
            
            If d.Exists(tempStr) Then
                d(tempStr) = d(tempStr) + 1
            Else
                d.Add pExpressions.InputToString(elem), i
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = pExpressions.InputToString(elem)
            
            If d2.Exists(tempStr) Then
                d2(tempStr) = d2(tempStr) + 1
            Else
                d2.Add pExpressions.InputToString(elem), i
            End If
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(d, negateValue, d2, False)
    End If
    
    IHaveTestFuncs_SameElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)

End Function

Private Function assignOrConcatStr(ByVal assignStr As String, ByVal valueStr As String) As String
    Dim finalStr As String
    
    If assignStr = "" Then
        finalStr = valueStr
    Else
        finalStr = assignStr & ", " & valueStr
    End If
    
    assignOrConcatStr = finalStr
End Function

Private Function IHaveTestFuncs_Procedure( _
    ByVal testingValue As Variant, _
    ByVal procName As String, _
    ByVal callType As VBA.VbCallType, _
    ByVal negateValue As Boolean _
) As Variant
    Dim funcVal As Variant
    Dim callCounter As Long
    Dim errCounter As Long
    Dim callTypeStr As String
    Dim testingInputStr As String
    Dim actualStr As String
    Dim tempStr As String
        
    callCounter = 0
    errCounter = 0
    callTypeStr = ""
    
    Const ARG_NOT_OPTIONAL As Long = 449
    
    testingInputStr = procName & "("
    actualStr = testingInputStr
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
'    If VBA.Information.IsObject(testingValue) And pTestValueSet Then
    If inputIsValid(testingValue, flTestType.flIsObject) Then
        If Not testingValue Is Nothing Then
            On Error Resume Next
            'This section uses bitwise flags / composite values
            'with the bitand operator
            
            'You can see more here in the Bitwise Flags section:
            'https://developer.mozilla.org/en-US/docs/Glossary/Bitwise_flags
            
            'You can see more here in the composite values section:
            'and in the composite values  section here:
            'http://www.cpearson.com/excel/Enums.aspx
            
                If callType And VbCallType.VbGet Then
                    tempStr = "VbGet"
                    callCounter = callCounter + 1
                    
                    Call VBA.Interaction.CallByName(testingValue, procName, VbCallType.VbGet)

                    If Err.Number = 0 Or Err.Number = ARG_NOT_OPTIONAL Then
                        errCounter = errCounter + 1
                        actualStr = tempStr
                    Else
                        Err.clear
                    End If
                    
                    callTypeStr = tempStr
                End If
                
                If callType And VbCallType.VbLet Then
                    tempStr = "VbLet"
                    callCounter = callCounter + 1
                    
                    Call VBA.Interaction.CallByName(testingValue, procName, VbCallType.VbLet, Empty)
                    
                    If (Err.Number = 0) Then
                        errCounter = errCounter + 1
                        
                        actualStr = assignOrConcatStr(actualStr, tempStr)
                    Else
                        Err.clear
                    End If
                    
                    callTypeStr = assignOrConcatStr(callTypeStr, tempStr)
                End If
                
                If callType And VbCallType.VbMethod Then
                    tempStr = "VbMethod"
                    callCounter = callCounter + 1
                    
                    Call VBA.Interaction.CallByName(testingValue, procName, VbCallType.VbMethod)
                    
                    If Err.Number = 0 Or Err.Number = ARG_NOT_OPTIONAL Then
                        errCounter = errCounter + 1
                        
                        actualStr = assignOrConcatStr(actualStr, tempStr)
                    Else
                        Err.clear
                    End If
                    
                    callTypeStr = assignOrConcatStr(callTypeStr, tempStr)
                End If
                
                If callType And VbCallType.VbSet Then
                    tempStr = "VbSet"
                    
                    callCounter = callCounter + 1
                    
                    Call VBA.Interaction.CallByName(testingValue, procName, VbCallType.VbSet, Nothing)
                    
                    If (Err.Number = 0) Then
                        errCounter = errCounter + 1
                        
                        actualStr = assignOrConcatStr(actualStr, tempStr)
                    Else
                        Err.clear
                    End If
                    
                    callTypeStr = assignOrConcatStr(callTypeStr, tempStr)
                End If
            On Error GoTo 0
            
            actualStr = testingInputStr & actualStr & ")"
            
            testingInputStr = testingInputStr & callTypeStr & ")"
            
            funcVal = (callCounter > 0 And errCounter > 0 And callCounter = errCounter)
        End If
    End If
    
    IHaveTestFuncs_Procedure = funcVal
    
    If testingInputStr <> actualStr Then
        Call pExpressions.addToTestAndFluentPath("Procedure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInputStr, Actual:=actualStr)
    Else
        Call pExpressions.addToTestAndFluentPath("Procedure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInputStr)
    End If

End Function

Private Function IHaveTestFuncs_Elements( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Variant, _
    ByVal negateValue As Boolean, _
    ByVal updateFluentPath As Boolean _
) As Variant
    Dim i As Long
    Dim boolCounter As Long
    Dim elemCount As Long
    Dim elem As Variant
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
'    Dim tempTestingValue As Variant
    Dim tempTestingInput As VBA.Collection
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    boolCounter = 0
    elemCount = 0
    
    IsDataStructure = inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingValueIsNotSelfReferential And IsDataStructure Then
'        elemCount = pExpressions.getDsCount(testingInput)
'
'        For Each elem In testingInput
'            If IBeTestFuncs_InDataStructure(elem, testingValue, False, False) Then
'                boolCounter = boolCounter + 1
'            End If
'        Next elem

        Set tempTestingInput = getElementsFromDsUpToFirstMissing(testingInput)

        elemCount = pExpressions.getDsCount(tempTestingInput)

        For Each elem In tempTestingInput
            If IBeTestFuncs_InDataStructure(elem, testingValue, False, False) Then
                boolCounter = boolCounter + 1
            End If
        Next elem

        funcVal = (boolCounter = elemCount)
    End If
    
    IHaveTestFuncs_Elements = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("Elements", funcVal, VBA.Conversion.CBool(negateValue), testingValue, tempTestingInput, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential)
    End If
    
End Function

Public Function IHaveTestFuncs_ElementsInDataStructure( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Variant, _
    ByVal negateValue As Boolean _
) As Variant
    Dim funcVal As Variant
    Dim IsDataStructure As Boolean
    Dim testingInputIsNotSelfReferential As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    IsDataStructure = inputIsValid(testingInput, flTestType.flIsDataStructure) And inputIsValid(testingValue, flTestType.flIsDataStructure)
    
    testingInputIsNotSelfReferential = inputIsValid(testingInput, flTestType.flIsNotSelfReferential)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If testingInputIsNotSelfReferential And testingValueIsNotSelfReferential And IsDataStructure Then
        funcVal = IHaveTestFuncs_Elements(testingValue, testingInput, False, False)
    End If
    
    IHaveTestFuncs_ElementsInDataStructure = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ElementsInDataStructure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential, TestingInputIsSelfReferential:=testingInputIsNotSelfReferential)
End Function

Function getDepthCountRecur(ByVal ds As Variant, Optional ByVal counter As Long = 1, Optional ByRef depthCounter As Long = 0) As Long
    Dim elem As Variant

    For Each elem In ds
        If counter > depthCounter Then depthCounter = counter

        If TypeOf ds Is Scripting.Dictionary Then
            If pMeta.tests.IsDataStructure(ds(elem)) Then
                Call getDepthCountRecur(ds(elem), counter + 1, depthCounter)
            End If
        Else
            If pMeta.tests.IsDataStructure(elem) Then
                Call getDepthCountRecur(elem, counter + 1, depthCounter)
            End If
        End If
    Next elem

    getDepthCountRecur = depthCounter
End Function

Private Function getDepthCountIter(ByVal ds As Variant) As Long
    Dim counter As Long
    Dim depthCounter As Long
    Dim hasNestedElements As Boolean
    Dim elem As Variant
    Dim elem2 As Variant
    Dim col As VBA.Collection
    Dim dsCounter As Long
    Dim tempCol As VBA.Collection
    
    counter = 0
    depthCounter = 0
    hasNestedElements = False
    dsCounter = 1
    
    Set col = New VBA.Collection
    Set tempCol = New VBA.Collection

    If pExpressions.getDsCount(ds) > 0 Then
        counter = counter + 1
        
        For Each elem In ds
'            col.Add elem
            If Not TypeOf ds Is Scripting.Dictionary Then
                col.Add elem
            Else
                col.Add ds(elem)
            End If
        Next elem
    End If
    
    Do
        hasNestedElements = False
        
        'dsCounter gets a count of how many data structures
        If dsCounter = 1 Then
            For Each elem In col
                If pMeta.tests.IsDataStructure(elem) Then
                    If pExpressions.getDsCount(elem) > 0 Then
                        dsCounter = dsCounter + 1
                        tempCol.Add elem
                    End If
                End If
            Next elem
        End If
        
        Set col = tempCol
        Set tempCol = New VBA.Collection
        
        For Each elem In col
            If pMeta.tests.IsDataStructure(elem) Then
                If pExpressions.getDsCount(elem) > 0 Then
                    counter = counter + 1
                    
                    hasNestedElements = True
                                        
                    For Each elem2 In elem
'                        tempCol.Add elem2
                        If Not TypeOf elem Is Scripting.Dictionary Then
                            tempCol.Add elem2
                        Else
                            tempCol.Add elem(elem2)
                        End If
                    Next elem2
                End If
            End If
        Next elem
                
        If counter = dsCounter And counter > 0 And dsCounter > 0 Then
            depthCounter = depthCounter + 1
            counter = 1
            dsCounter = 1
            Set col = tempCol
            Set tempCol = New VBA.Collection
        End If
    Loop While hasNestedElements = True
    
    getDepthCountIter = depthCounter
End Function

Private Function IHaveTestFuncs_DepthCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
) As Variant
    Dim funcVal As Variant
    Dim method As Long
    Dim tempCount As Long
    Dim recurCol As VBA.Collection
    Dim iterCol As VBA.Collection
    Dim IsDataStructure As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.tests.Algorithm
    
    IsDataStructure = inputIsValid(testingValue, flTestType.flIsDataStructure)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If IsDataStructure And testingValueIsNotSelfReferential Then
        If method And flAlgorithm.flRecursive Then
           Set recurCol = getNestedElementsRecur(testingValue)
        End If
        
        If method And flAlgorithm.flIterative Then
            Set iterCol = getNestedElementsIter(testingValue)
        End If
        
        If (Not recurCol Is Nothing And Not iterCol Is Nothing) Then
            If IBeTestFuncs_IdenticalTo(recurCol, negateValue, iterCol, False) Then
                tempCount = getDepthCountRecur(testingValue)
            Else
                Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative collections are not identical"
            End If
        Else
            If Not recurCol Is Nothing Then
                tempCount = getDepthCountRecur(testingValue)
            ElseIf Not iterCol Is Nothing Then
                tempCount = getDepthCountIter(testingValue)
            Else
                Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
            End If
        End If
    
        funcVal = (tempCount = testingInput)
    End If
    
    IHaveTestFuncs_DepthCountOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("DepthCountOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, recurIterFunc:=True, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential)
End Function

Private Function IHaveTestFuncs_NestedCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
    ) As Variant
    Dim funcVal As Variant
    Dim method As Long
    Dim tempCount As Long
    Dim tempCol As VBA.Collection
    Dim recurCol As VBA.Collection
    Dim iterCol As VBA.Collection
    Dim IsDataStructure As Boolean
    Dim testingValueIsNotSelfReferential As Boolean
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.tests.Algorithm
    
    IsDataStructure = inputIsValid(testingValue, flTestType.flIsDataStructure)
    testingValueIsNotSelfReferential = inputIsValid(testingValue, flTestType.flIsNotSelfReferential)
    
    If IsDataStructure And testingValueIsNotSelfReferential Then
        If method And flAlgorithm.flRecursive Then
           Set recurCol = getNestedElementsRecur(testingValue)
        End If
        
        If method And flAlgorithm.flIterative Then
            Set iterCol = getNestedElementsIter(testingValue)
        End If
        
        If (Not recurCol Is Nothing And Not iterCol Is Nothing) Then
            If IBeTestFuncs_IdenticalTo(recurCol, negateValue, iterCol, False) Then
                Set tempCol = getNestedElementsRecur(testingValue)
            Else
                Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative collections are not identical"
            End If
        Else
            If Not recurCol Is Nothing Then
                Set tempCol = getNestedElementsRecur(testingValue)
            ElseIf Not iterCol Is Nothing Then
                Set tempCol = getNestedElementsIter(testingValue)
            Else
                Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
            End If
        End If
        
        tempCount = tempCol.Count
    
        funcVal = (tempCount = testingInput)
    End If
    
    IHaveTestFuncs_NestedCountOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("NestedCountOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, recurIterFunc:=True, TestingValueIsSelfReferential:=testingValueIsNotSelfReferential)
End Function

Private Function getElementsFromDsUpToFirstMissing(ds As Variant) As VBA.Collection
    Dim tempCol As VBA.Collection
    Dim elem As Variant
    
    Set tempCol = New VBA.Collection
    
    For Each elem In ds
        If VBA.Information.IsMissing(elem) Then
            Exit For
        Else
            tempCol.Add elem
        End If
    Next elem
    
    Set getElementsFromDsUpToFirstMissing = tempCol
End Function

Private Sub Class_Terminate()
    Set pMeta = Nothing
    Set pTests = Nothing
    Set pPrinting = Nothing
    Set iPrint = Nothing
    Set pTestDictCounter = Nothing
    Set pExpressions = Nothing
End Sub
