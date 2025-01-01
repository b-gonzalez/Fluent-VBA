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
    pTestValueClean = pMeta.Tests.TestStrings.CleanString(value)
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
        
        With pMeta.Tests.TestStrings
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
        
        With pMeta.Tests.TestStrings
            If .CleanTestInputStr Or .CleanTestStrings Then
                tempTestingInput = pMeta.Tests.TestStrings.CleanString(testingInput)
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
    ByVal testingInput As Variant) As Variant
    
    Dim col As VBA.Collection
    Dim i As Long
    Dim funcVal As Variant
    Dim method As flAlgorithm
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.Tests.Algorithm
    
    If pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
        For i = LBound(testingInput) To UBound(testingInput)
            If pMeta.Tests.IsDataStructure(testingInput(i)) Then
                If method = flRecursive Then
                    Set col = IBeTestFuncs_getNestedElementsRecur(testingInput)
                ElseIf method = flIterative Then
                    Set col = IBeTestFuncs_getNestedElementsIter(testingInput)
                End If
            
                funcVal = IBeTestFuncs_InDataStructure(testingValue, col, negateValue, updateFluentPath:=False)
            End If
        Next i
    End If
    
    IBeTestFuncs_InDataStructures = funcVal
    
    Call pExpressions.addToTestAndFluentPath("InDataStructures", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_InDataStructure( _
ByVal testingValue As Variant, _
ByVal dataStructure As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    Dim funcVal As Variant
    Dim method As flAlgorithm
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.Tests.Algorithm
    
    If pMeta.Tests.IsDataStructure(dataStructure) And pTestValueSet Then
        tempBool = False
    
        If method = flRecursive Then
            Set col = IBeTestFuncs_getNestedElementsRecur(dataStructure)
        ElseIf method = flIterative Then
            Set col = IBeTestFuncs_getNestedElementsIter(dataStructure)
        End If
            
        For Each elem In col
            If testingValue = elem Then
                tempBool = True
                Exit For
            End If
        Next elem
        
        funcVal = tempBool
    End If
    
    IBeTestFuncs_InDataStructure = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("InDataStructure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, dataStructure)
    End If
    
End Function

Private Function IBeTestFuncs_getNestedElementsRecur(ByVal dataStructure As Variant, Optional ByRef col As VBA.Collection = Nothing) As VBA.Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New VBA.Collection
    End If
    
    If pMeta.Tests.IsDataStructure(dataStructure) Then
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
            If pMeta.Tests.IsDataStructure(elem) Then
                Set col = IBeTestFuncs_getNestedElementsRecur(elem, col)
            Else
                col.Add elem
            End If
        Next elem
    Else
        col.Add elem
    End If
    
    Set IBeTestFuncs_getNestedElementsRecur = col
End Function

Private Function IBeTestFuncs_getNestedElementsIter(ByVal dataStructure As Variant) As VBA.Collection
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
            If pMeta.Tests.IsDataStructure(col(i)) Then
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
                End If
            End If
        Next i
    Loop While hasDataStructures = True

    Set IBeTestFuncs_getNestedElementsIter = col
End Function

Private Function IBeTestFuncs_GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) And pTestValueSet Then
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
        
    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) And pTestValueSet Then
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
    
    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) And pTestValueSet Then
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
        
    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) And pTestValueSet Then
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
    
    If Not VBA.Information.IsObject(testingValue) And Not pMeta.Tests.IsDataStructure(testingValue) And _
    Not VBA.Information.IsObject(testingInput) And Not pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
    
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
            
            If pMeta.Tests.ApproximateEqual Then
                If VBA.Information.TypeName(testingValue) <> VBA.Information.TypeName(testingInput) Then
                    If VBA.Information.IsNumeric(testingValue) And VBA.Information.IsNumeric(testingInput) Then
                        tempBool = Roughly(CDbl(testingValue), VBA.Conversion.CDbl(testingInput), pMeta.Tests.Epsilon)
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If VBA.Information.IsObject(testingValue) And pTestValueSet Then
        funcVal = (Not testingValue Is Nothing)
    End If
    
    IBeTestFuncs_Something = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Something", funcVal, VBA.Conversion.CBool(negateValue), testingValue)
        
End Function

Private Function IBeTestFuncs_Between( _
ByVal testingValue As Variant, _
ByVal lowerVal As Double, _
ByVal higherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If VBA.Information.IsNumeric(testingValue) And pTestValueSet Then
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
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If pTestValueSet Then
        If VBA.Information.TypeName(testingValue) = "String" Then
            tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        End If
    
        tempBool = False
        
        For i = LBound(testingInput) To UBound(testingInput)
            If VBA.Information.IsObject(testingInput(i)) Or VBA.Information.IsObject(testingValue) Then
                If VBA.Information.TypeName(testingInput(i)) = VBA.Information.TypeName(testingValue) Then
                    tempBool = True
                End If
            ElseIf testingInput(i) = testingValue Then
                tempBool = True
            ElseIf VBA.Information.IsNull(testingInput(i)) Or VBA.Information.IsNull(testingValue) Then
                If VBA.Information.TypeName(testingInput(i)) = VBA.Information.TypeName(testingValue) Then
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
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
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
    
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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

    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
        
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
    
    If Not VBA.Information.IsObject(testingValue) And Not VBA.Information.IsArray(testingValue) And Not VBA.Information.IsNull(testingValue) And pTestValueSet Then
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
    
    If VBA.Information.TypeName(testingValue) = "ErrObject" Or VBA.Information.TypeName(testingValue) = "String" And pTestValueSet Then
    
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
        
    If VBA.Information.TypeName(testingValue) = "ErrObject" Or VBA.Information.TypeName(testingValue) = "String" And pTestValueSet Then
    
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
        
    If VBA.Information.TypeName(testingValue) = "ErrObject" Or VBA.Information.TypeName(testingValue) = "String" And pTestValueSet Then
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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
        funcVal = _
        IBeTestFuncs_EqualTo( _
            pExpressions.InputToString(testingValue), _
            pExpressions.InputToString(testingInput), _
            updateFluentPath:=False _
        )
    End If
    
    IBeTestFuncs_IdenticalTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("IdenticalTo", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
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
    
    Call pExpressions.addToTestAndFluentPath("ExactSameElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)

End Function

Private Function IHaveTestFuncs_SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim tempStr As String
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
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
    
    Call pExpressions.addToTestAndFluentPath("SameUniqueElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)

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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
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
    
    Call pExpressions.addToTestAndFluentPath("SameElementsAs", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)

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
    
    If VBA.Information.IsObject(testingValue) And pTestValueSet Then
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
                        Err.Clear
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
                        Err.Clear
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
                        Err.Clear
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
                        Err.Clear
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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    boolCounter = 0
    elemCount = 0
    
    If pMeta.Tests.IsDataStructure(testingValue) And pTestValueSet Then
        elemCount = pExpressions.getDsCount(testingInput)

        For Each elem In testingInput
            If IBeTestFuncs_InDataStructure(elem, testingValue, False, False) Then
                boolCounter = boolCounter + 1
            End If
        Next elem

        funcVal = (boolCounter = elemCount)
    End If
    
    IHaveTestFuncs_Elements = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("Elements", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
    End If
    
    
End Function

Public Function IHaveTestFuncs_ElementsInDataStructure( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Variant, _
    ByVal negateValue As Boolean _
) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
        funcVal = IHaveTestFuncs_Elements(testingValue, testingInput, False, False)
    End If
    
    IHaveTestFuncs_ElementsInDataStructure = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ElementsInDataStructure", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
End Function

Private Function IHaveTestFuncs_getDepthCountRecur(ds As Variant, Optional counter As Long = 1, Optional depthCounter As Long = 0) As Long
    Dim elem As Variant

    For Each elem In ds
        If counter > depthCounter Then depthCounter = counter

        If TypeOf ds Is Scripting.Dictionary Then
            If pMeta.Tests.IsDataStructure(ds(elem)) Then
                Call IHaveTestFuncs_getDepthCountRecur(ds(elem), counter + 1, depthCounter)
            End If
        Else
            If pMeta.Tests.IsDataStructure(elem) Then
                Call IHaveTestFuncs_getDepthCountRecur(elem, counter + 1, depthCounter)
            End If
        End If
    Next elem

    IHaveTestFuncs_getDepthCountRecur = depthCounter
End Function

Private Function IHaveTestFuncs_getDepthCountIter(ds As Variant) As Long
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
                If pMeta.Tests.IsDataStructure(elem) Then
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
            If pMeta.Tests.IsDataStructure(elem) Then
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
    
    IHaveTestFuncs_getDepthCountIter = depthCounter
End Function

Private Function IHaveTestFuncs_DepthCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
) As Variant
    Dim funcVal As Variant
    Dim method As flAlgorithm
    Dim tempCount As Long
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.Tests.Algorithm
    
    If pMeta.Tests.IsDataStructure(testingValue) And pTestValueSet Then
        If method = flRecursive Then
            tempCount = IHaveTestFuncs_getDepthCountRecur(testingValue)
        ElseIf method = flIterative Then
            tempCount = IHaveTestFuncs_getDepthCountIter(testingValue)
        End If
    
        funcVal = (tempCount = testingInput)
    End If
    
    IHaveTestFuncs_DepthCountOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("DepthCountOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
End Function

Private Function IHaveTestFuncs_NestedCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
    ) As Variant
    Dim funcVal As Variant
    Dim method As flAlgorithm
    Dim tempCount As Long
    Dim tempCol As VBA.Collection
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.Tests.Algorithm
    
    If pMeta.Tests.IsDataStructure(testingValue) And pTestValueSet Then
        If method = flRecursive Then
            Set tempCol = IBeTestFuncs_getNestedElementsRecur(testingValue)
        ElseIf method = flIterative Then
            Set tempCol = IBeTestFuncs_getNestedElementsIter(testingValue)
        End If
        
        tempCount = tempCol.Count
    
        funcVal = (tempCount = testingInput)
    End If
    
    IHaveTestFuncs_NestedCountOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("NestedCountOf", funcVal, VBA.Conversion.CBool(negateValue), testingValue, testingInput)
End Function

