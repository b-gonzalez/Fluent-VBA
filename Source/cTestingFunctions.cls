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
    If Information.IsObject(pTestValue) Then
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
    
    tempVal = Replace(tempVal, " ", "")
    
    If InStr(1, tempVal, """") = 1 Then
        tempVal = Replace(tempVal, """", "")
    End If
    
    ITestingFunctions_GetCleanedString = tempVal
End Function

Private Function ITestingFunctions_getTempTestValue(ByVal testingValue As Variant) As Variant
    Dim tempTestValue As Variant
    
    If Not Information.IsObject(testingValue) Then
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
    
    If Not Information.IsObject(testingInput) Then
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
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    For i = LBound(testingInput) To UBound(testingInput)
        If pMeta.Tests.IsDataStructure(testingInput(i)) Then
            Set col = IBeTestFuncs_getNestedElementsRecur(testingInput)
        
            funcVal = IBeTestFuncs_InDataStructure(testingValue, col, negateValue, updateFluentPath:=False)
        End If
    Next i
    
    IBeTestFuncs_InDataStructures = funcVal
    
    Call pExpressions.addToTestAndFluentPath("InDataStructures", funcVal, CBool(negateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_InDataStructure( _
ByVal testingValue As Variant, _
ByVal datastructure As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    Dim funcVal As Variant
    Dim method As flAlgorithm
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    method = pMeta.Tests.Algorithm
    
    If pMeta.Tests.IsDataStructure(datastructure) And pTestValueSet Then
        tempBool = False
    
        If method = flRecursive Then
            Set col = IBeTestFuncs_getNestedElementsRecur(datastructure)
        ElseIf method = flIterative Then
            Set col = IBeTestFuncs_getNestedElementsIter(datastructure)
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
        Call pExpressions.addToTestAndFluentPath("InDataStructure", funcVal, CBool(negateValue), testingValue, datastructure)
    End If
    
End Function

Private Function IBeTestFuncs_getNestedElementsRecur(ByVal datastructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If pMeta.Tests.IsDataStructure(datastructure) Then
        If TypeOf datastructure Is Scripting.Dictionary Then
            tempDataStructure = datastructure.Items
        Else
            If Information.IsObject(datastructure) Then
                Set tempDataStructure = datastructure
            Else
                tempDataStructure = datastructure
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

Private Function IBeTestFuncs_getNestedElementsIter(ByVal datastructure As Variant) As Collection
    Dim elem As Variant
    Dim col As Collection
    Dim tempDataStructure As Variant
    Dim i As Long
    Dim hasDataStructures As Boolean

    Set col = New Collection

    If TypeOf datastructure Is Scripting.Dictionary Then
        tempDataStructure = datastructure.Items
    Else
        If Information.IsObject(datastructure) Then
            Set tempDataStructure = datastructure
        Else
            tempDataStructure = datastructure
        End If
    End If

    For Each elem In tempDataStructure
        col.Add elem
    Next elem
    
    Do
        hasDataStructures = False
        For i = 1 To col.Count
            If pMeta.Tests.IsDataStructure(col(i)) Then
                hasDataStructures = True
                For Each elem In col(i)
                    col.Add elem
                Next elem
                col.Remove i
                Exit For
            End If
        Next i
    Loop While hasDataStructures = True

    Set IBeTestFuncs_getNestedElementsIter = col
End Function

Private Function IBeTestFuncs_GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) And pTestValueSet Then
        funcVal = (testingValue > testingInput)
    End If
    
    IBeTestFuncs_GreaterThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("GreaterThan", funcVal, CBool(negateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_GreaterThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If IsNumeric(testingValue) And IsNumeric(testingInput) And pTestValueSet Then
        funcVal = (testingValue >= testingInput)
    End If
    
    IBeTestFuncs_GreaterThanOrEqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("GreaterThanOrEqualTo", funcVal, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IBeTestFuncs_LessThan( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) And pTestValueSet Then
        funcVal = (testingValue < testingInput)
    End If
    
    IBeTestFuncs_LessThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LessThan", funcVal, CBool(negateValue), testingValue, testingInput)

End Function

Private Function IBeTestFuncs_LessThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If IsNumeric(testingValue) And IsNumeric(testingInput) And pTestValueSet Then
        funcVal = (testingValue <= testingInput)
    End If
    
    IBeTestFuncs_LessThanOrEqualTo = funcVal
       
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("LessThanOrEqualTo", funcVal, CBool(negateValue), testingValue, testingInput)
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
    
    x = Abs(a)
    y = Abs(b)
    d = Abs(a - b)
    
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
    
    If Not Information.IsObject(testingValue) And Not pMeta.Tests.IsDataStructure(testingValue) And _
    Not Information.IsObject(testingInput) And Not pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
    
        If Information.IsNull(testingValue) Or Information.IsNull(testingInput) Then
            tempBool = (Information.TypeName(testingValue) = Information.TypeName(testingInput))
            
            testingValue = Information.TypeName(testingValue)
            testingInput = Information.TypeName(testingInput)
        Else
            If Information.TypeName(testingValue) = "String" And Information.TypeName(testingInput) = "String" Then
                tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
                
                tempBool = (tempTestValue = tempTestInput)
            Else
                tempBool = (testingValue = testingInput)
            End If
            
            If pMeta.Tests.ApproximateEqual Then
                If Information.TypeName(testingValue) <> Information.TypeName(testingInput) Then
                    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
                        tempBool = Roughly(CDbl(testingValue), CDbl(testingInput), pMeta.Tests.Epsilon)
                    ElseIf VarType(testingValue) = vbString Or VarType(testingInput) = vbString Then
                        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
                        tempBool = (StrComp(CStr(tempTestValue), CStr(tempTestInput), vbTextCompare) = 0)
                    End If
                End If
            End If
        End If
        
        funcVal = tempBool
    End If
    
    IBeTestFuncs_EqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("EqualTo", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
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
        
    If Not Information.IsObject(testingValue) And Not IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        funcVal = CStr(tempTestValue) Like "*" & CStr(tempTestInput) & "*"
    End If
    
    IShouldTestFuncs_Contain = funcVal
        
    Call pExpressions.addToTestAndFluentPath("Contain", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
        
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
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        valLength = Len(tempTestInput)
        funcVal = Left$(tempTestValue, valLength) = CStr(tempTestInput)
    End If
    
    IShouldTestFuncs_StartWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("StartWith", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
        
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
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
    
        valLength = Len(tempTestInput)
        funcVal = Right$(tempTestValue, valLength) = CStr(tempTestInput)
    End If
    
    IShouldTestFuncs_EndWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EndWith", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
    
End Function

Private Function IHaveTestFuncs_LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = (Len(CStr(testingValue)) = testingInput)
    End If
    
    IHaveTestFuncs_LengthOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LengthOf", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
            
End Function

Private Function IHaveTestFuncs_MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = (Len(CStr(tempTestValue)) <= testingInput)
    End If
    
    IHaveTestFuncs_MaxLengthOf = funcVal
    
    If updateFluentPath Then
        If tempTestValue <> "" Then
            Call pExpressions.addToTestAndFluentPath("MaxLengthOf", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
        Else
            Call pExpressions.addToTestAndFluentPath("MaxLengthOf", funcVal, CBool(negateValue), testingValue, testingInput)
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
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        funcVal = (Len(CStr(tempTestValue)) >= testingInput)
    End If
    
    IHaveTestFuncs_MinLengthOf = funcVal
        
    If updateFluentPath Then
        If tempTestValue <> "" Then
            Call pExpressions.addToTestAndFluentPath("MinLengthOf", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
        Else
            Call pExpressions.addToTestAndFluentPath("MinLengthOf", funcVal, CBool(negateValue), testingValue, testingInput)
        End If
    End If
        
End Function

Private Function IBeTestFuncs_Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If Information.IsObject(testingValue) And pTestValueSet Then
        funcVal = (Not testingValue Is Nothing)
    End If
    
    IBeTestFuncs_Something = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Something", funcVal, CBool(negateValue), testingValue)
        
End Function

Private Function IBeTestFuncs_Between( _
ByVal testingValue As Variant, _
ByVal lowerVal As Double, _
ByVal higherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If Information.IsNumeric(testingValue) And pTestValueSet Then
        funcVal = _
        (IBeTestFuncs_GreaterThanOrEqualTo(testingValue, lowerVal, updateFluentPath:=False) And _
        IBeTestFuncs_LessThanOrEqualTo(testingValue, higherVal, updateFluentPath:=False))
    End If
    
    IBeTestFuncs_Between = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Between", funcVal, CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal)
End Function

Private Function IBeTestFuncs_OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim tempBool As Boolean
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    Dim tempTestInput As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
    
    If pTestValueSet Then
        If Information.TypeName(testingValue) = "String" Then
            tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        End If
    
        tempBool = False
        
        For i = LBound(testingInput) To UBound(testingInput)
            If Information.IsObject(testingInput(i)) Or Information.IsObject(testingValue) Then
                If Information.TypeName(testingInput(i)) = Information.TypeName(testingValue) Then
                    tempBool = True
                End If
            ElseIf testingInput(i) = testingValue Then
                tempBool = True
            ElseIf Information.IsNull(testingInput(i)) Or Information.IsNull(testingValue) Then
                If Information.TypeName(testingInput(i)) = Information.TypeName(testingValue) Then
                    tempBool = True
                End If
            End If
            
            If tempBool Then Exit For
        Next i
        
        funcVal = tempBool
    End If
        
    IBeTestFuncs_OneOf = funcVal
    
    If TypeOf testingValue Is Object  Then testingValue = Information.TypeName(testingValue)
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, CBool(negateValue), testingValue, testingInput)
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
    
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        funcVal = _
        (IHaveTestFuncs_MinLengthOf(tempTestValue, lowerVal, updateFluentPath:=False) And _
        IHaveTestFuncs_MaxLengthOf(tempTestValue, higherVal, updateFluentPath:=False))
    End If
    
    IHaveTestFuncs_LengthBetween = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("LengthBetween", funcVal, CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("LengthBetween", funcVal, CBool(negateValue), testingValue:=testingValue, lowerVal:=lowerVal, higherVal:=higherVal)
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

    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
        
        funcVal = (Excel.Evaluate(tempTestValue) = tempTestInput)
    End If
    
    IShouldTestFuncs_EvaluateTo = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EvaluateTo", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
    
End Function

Private Function IBeTestFuncs_Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
    
        For i = 1 To Len(tempTestValue)
            tempChar = Mid$(tempTestValue, i, 1)
            
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            End If
        Next i
        
        funcVal = (tempTestValue = tempAlph)
    End If
    
    IBeTestFuncs_Alphabetic = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Alphabetic", funcVal, CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Alphabetic", funcVal, CBool(negateValue), testingValue)
    End If
End Function

Private Function IBeTestFuncs_Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    Dim funcVal As Variant
    Dim tempTestValue As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        For i = 1 To Len(tempTestValue)
            tempChar = Mid$(tempTestValue, i, 1)
            If tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        funcVal = (tempTestValue = tempNum)
    End If
    
    IBeTestFuncs_Numeric = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Numeric", funcVal, CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Numeric", funcVal, CBool(negateValue), testingValue)
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
    
    If Not Information.IsObject(testingValue) And Not Information.IsArray(testingValue) And Not Information.IsNull(testingValue) And pTestValueSet Then
        tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
        
        For i = 1 To Len(tempTestValue)
            tempChar = Mid$(tempTestValue, i, 1)
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            ElseIf tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        If Len(tempAlph) > 1 And Len(tempNum) > 1 Then
            If Len(tempAlph) + Len(tempNum) = Len(tempTestValue) Then
                b = True
            End If
        End If
        
        funcVal = b
    End If
    
    IBeTestFuncs_Alphanumeric = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Alphanumeric", funcVal, CBool(negateValue), testingValue, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Alphanumeric", funcVal, CBool(negateValue), testingValue)
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
    
    If Information.TypeName(testingValue) = "ErrObject" Or Information.TypeName(testingValue) = "String" And pTestValueSet Then
    
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            temp = (tempErrObj.Number <> 0)
        ElseIf Information.TypeName(testingValue) = "String" Then
            On Error Resume Next
                tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
                temp = Excel.Evaluate(tempTestValue)
                temp = (Information.TypeName(temp) = "Error")
            On Error GoTo 0
        End If
        
        funcVal = temp
        
        If TypeOf testingValue Is Object  Then
            tempErrType = Information.TypeName(testingValue)
        Else
            tempErrType = testingValue
        End If
    End If
    
    IBeTestFuncs_Erroneous = funcVal
    
    If tempTestValue <> "" Then
        Call pExpressions.addToTestAndFluentPath("Erroneous", funcVal, CBool(negateValue), tempErrType, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("Erroneous", funcVal, CBool(negateValue), tempErrType)
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
        
    If Information.TypeName(testingValue) = "ErrObject" Or Information.TypeName(testingValue) = "String" And pTestValueSet Then
    
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            temp = tempErrObj.Number
        ElseIf Information.TypeName(testingValue) = "String" Then
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
        Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", funcVal, CBool(negateValue), testingValue, testingInput, cleanedTestValue:=tempTestValue)
    Else
        Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", funcVal, CBool(negateValue), testingValue, testingInput)
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
        
    If Information.TypeName(testingValue) = "ErrObject" Or Information.TypeName(testingValue) = "String" And pTestValueSet Then
            tempTestInput = ITestingFunctions_getTempTestInput(testingInput)
        If TypeOf testingValue Is VBA.ErrObject Then
            Set tempErrObj = testingValue
            tempTestVal = tempErrObj.Description
        ElseIf Information.TypeName(testingValue) = "String" Then
            tempTestValue = ITestingFunctions_getTempTestValue(testingValue)
            
            On Error Resume Next
                tempTestVal = CStr(Excel.Evaluate(tempTestValue))
                tempTestVal = Replace(tempTestVal, "Error ", "")
                Err.Raise tempTestVal
                tempTestVal = Err.Description
            On Error GoTo 0
        End If
        
        funcVal = (tempTestVal = tempTestInput)
    End If
    
    IHaveTestFuncs_ErrorDescriptionOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ErrorDescriptionOf", funcVal, CBool(negateValue), tempTestVal, testingInput, cleanedTestValue:=tempTestValue, cleanedTestInput:=tempTestInput)
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
        If Not Information.IsNull(testingValue) Then
            t1 = Information.TypeName(testingValue)
        Else
            t1 = testingValue
        End If
        
        If Not Information.IsNull(testingInput) Then
            t2 = Information.TypeName(testingInput)
        Else
            t2 = testingInput
        End If
        
        funcVal = IBeTestFuncs_EqualTo(t1, t2, updateFluentPath:=False)
    End If
    
    IHaveTestFuncs_SameTypeAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameTypeAs", funcVal, CBool(negateValue), testingValue, testingInput)
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
        Call pExpressions.addToTestAndFluentPath("IdenticalTo", funcVal, CBool(negateValue), testingValue, testingInput)
    End If

End Function

Private Function IHaveTestFuncs_ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim col2 As Collection
    Dim funcVal As Variant
    
    funcVal = ITestingFunctions_SetDefaultFuncVal
        
    If pMeta.Tests.IsDataStructure(testingValue) And pMeta.Tests.IsDataStructure(testingInput) And pTestValueSet Then
        Set col = New Collection
        Set col2 = New Collection
        
        For Each elem In testingValue
            col.Add elem
        Next elem
        
        For Each elem In testingInput
            col2.Add elem
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(col, negateValue, col2, False)
    End If
    
    IHaveTestFuncs_ExactSameElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ExactSameElementsAs", funcVal, CBool(negateValue), testingValue, testingInput)

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
    
    Call pExpressions.addToTestAndFluentPath("SameUniqueElementsAs", funcVal, CBool(negateValue), testingValue, testingInput)

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
    
    Call pExpressions.addToTestAndFluentPath("SameElementsAs", funcVal, CBool(negateValue), testingValue, testingInput)

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
    
    If Information.IsObject(testingValue) And pTestValueSet Then
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
        Call pExpressions.addToTestAndFluentPath("Procedure", funcVal, CBool(negateValue), testingValue, testingInputStr, Actual:=actualStr)
    Else
        Call pExpressions.addToTestAndFluentPath("Procedure", funcVal, CBool(negateValue), testingValue, testingInputStr)
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
        Call pExpressions.addToTestAndFluentPath("Elements", funcVal, CBool(negateValue), testingValue, testingInput)
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
    
    Call pExpressions.addToTestAndFluentPath("ElementsInDataStructure", funcVal, CBool(negateValue), testingValue, testingInput)
End Function

