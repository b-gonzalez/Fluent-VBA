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
'Private pExprDict As Scripting.Dictionary

Private pTestDictCounter As Scripting.Dictionary

Private pTestValue As Variant
Private pFluentPath As String
Private pExpressions As IExpressions

Private Property Let ITestingFunctions_TestValue(ByVal value As Variant)
    pTestValue = value
End Property

Private Property Set ITestingFunctions_TestValue(ByVal value As Variant)
    Set pTestValue = value
End Property

Private Property Get ITestingFunctions_TestValue() As Variant
    If IsObject(pTestValue) Then
        Set ITestingFunctions_TestValue = pTestValue
    Else
        ITestingFunctions_TestValue = pTestValue
    End If
End Property

Private Property Set ITestingFunctions_setMeta(ByVal value As cMeta)
    Set pMeta = value
End Property

Private Property Set ITestingFunctions_Expressions(ByVal value As IExpressions)
    Set pExpressions = value
End Property

Private Function IBeTestFuncs_InDataStructures( _
ByVal testingValue As Variant, _
ByVal method As flAlgorithm, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim col As VBA.Collection
    Dim i As Long
    Dim funcVal As Variant
    
    funcVal = Null
    
    For i = LBound(testingInput) To UBound(testingInput)
        If pExpressions.IsDataStructure(testingInput(i)) Then
            Set col = IBeTestFuncs_getNestedElementsRecur(testingInput)
        
            funcVal = IBeTestFuncs_InDataStructure(testingValue, col, method, NegateValue, updateFluentPath:=False)
        End If
    Next i
    
    IBeTestFuncs_InDataStructures = funcVal
    
    Call pExpressions.addToTestAndFluentPath("InDataStructures", funcVal, CBool(NegateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_InDataStructure( _
ByVal testingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flAlgorithm = flRecursive, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    Dim funcVal As Variant
    
    funcVal = Null
    
    If pExpressions.IsDataStructure(DataStructure) Then
        tempBool = False
    
        If method = flRecursive Then
            Set col = IBeTestFuncs_getNestedElementsRecur(DataStructure)
        ElseIf method = flIterative Then
            Set col = IBeTestFuncs_getNestedElementsIter(DataStructure)
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
        Call pExpressions.addToTestAndFluentPath("InDataStructure", funcVal, CBool(NegateValue), testingValue, DataStructure)
    End If
    
End Function

Private Function IBeTestFuncs_getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If pExpressions.IsDataStructure(DataStructure) Then
        If TypeName(DataStructure) = "Dictionary" Then
            tempDataStructure = DataStructure.Items
        Else
            If IsObject(DataStructure) Then
                Set tempDataStructure = DataStructure
            Else
                tempDataStructure = DataStructure
            End If
        End If
        
        For Each elem In tempDataStructure
            If pExpressions.IsDataStructure(elem) Then
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

Private Function IBeTestFuncs_getNestedElementsIter(ByVal DataStructure As Variant) As Collection
    Dim elem As Variant
    Dim col As Collection
    Dim tempDataStructure As Variant
    Dim i As Long
    Dim hasDataStructures As Boolean

    Set col = New Collection

    If TypeName(DataStructure) = "Dictionary" Then
        tempDataStructure = DataStructure.Items
    Else
        If IsObject(DataStructure) Then
            Set tempDataStructure = DataStructure
        Else
            tempDataStructure = DataStructure
        End If
    End If

    For Each elem In tempDataStructure
        col.Add elem
    Next elem
    
    Do
        hasDataStructures = False
        For i = 1 To col.Count
            If pExpressions.IsDataStructure(col(i)) Then
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

Private Function IBeTestFuncs_GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        funcVal = (testingValue > testingInput)
    End If
    
    IBeTestFuncs_GreaterThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("GreaterThan", funcVal, CBool(NegateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_GreaterThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        funcVal = (testingValue >= testingInput)
    End If
    
    IBeTestFuncs_GreaterThanOrEqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("GreaterThanOrEqualTo", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IBeTestFuncs_LessThan( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        funcVal = (testingValue < testingInput)
    End If
    
    IBeTestFuncs_LessThan = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LessThan", funcVal, CBool(NegateValue), testingValue, testingInput)

End Function

Private Function IBeTestFuncs_LessThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        funcVal = (testingValue <= testingInput)
    End If
    
    IBeTestFuncs_LessThanOrEqualTo = funcVal
       
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("LessThanOrEqualTo", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IBeTestFuncs_EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim tempBool As Boolean
    Dim funcVal As Variant
    
    funcVal = Null
    
    tempBool = False
    
    If Not IsObject(testingValue) And Not pExpressions.IsDataStructure(testingValue) And _
    Not IsObject(testingInput) And Not pExpressions.IsDataStructure(testingInput) Then
        If IsNull(testingValue) Or IsNull(testingInput) Then
            tempBool = (TypeName(testingValue) = TypeName(testingInput))
            
            testingValue = TypeName(testingValue)
            testingInput = TypeName(testingInput)
        Else
            tempBool = (testingValue = testingInput)
            
            If pMeta.ApproximateEqual Then
                If TypeName(testingValue) <> TypeName(testingInput) Then
                    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
                        tempBool = (Abs(testingValue - testingInput) <= pMeta.Epsilon)
                    ElseIf VarType(testingValue) = vbString Or VarType(testingInput) = vbString Then
                        tempBool = (StrComp(CStr(testingValue), CStr(testingInput), vbTextCompare) = 0)
                    End If
                End If
            End If
        End If
        
        funcVal = tempBool
    End If
    
    IBeTestFuncs_EqualTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("EqualTo", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IShouldTestFuncs_Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        funcVal = CStr(testingValue) Like "*" & CStr(testingInput) & "*"
    End If
    
    IShouldTestFuncs_Contain = funcVal
        
    Call pExpressions.addToTestAndFluentPath("Contain", funcVal, CBool(NegateValue), testingValue, testingInput)
        
End Function

Private Function IShouldTestFuncs_StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim valLength As Long
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        valLength = Len(testingInput)
        funcVal = Left$(testingValue, valLength) = CStr(testingInput)
    End If
    
    IShouldTestFuncs_StartWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("StartWith", funcVal, CBool(NegateValue), testingValue, testingInput)
        
End Function

Private Function IShouldTestFuncs_EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim valLength As Long
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        valLength = Len(testingInput)
        funcVal = Right$(testingValue, valLength) = CStr(testingInput)
    End If
    
    IShouldTestFuncs_EndWith = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EndWith", funcVal, CBool(NegateValue), testingValue, testingInput)
    
End Function

Private Function IHaveTestFuncs_LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        funcVal = (Len(CStr(testingValue)) = testingInput)
    End If
    
    IHaveTestFuncs_LengthOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LengthOf", funcVal, CBool(NegateValue), testingValue, testingInput)
            
End Function

Private Function IHaveTestFuncs_MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        funcVal = (Len(CStr(testingValue)) <= testingInput)
    End If
    
    IHaveTestFuncs_MaxLengthOf = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("MaxLengthOf", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If
            
End Function

Private Function IHaveTestFuncs_MinLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        funcVal = (Len(CStr(testingValue)) >= testingInput)
    End If
    
    IHaveTestFuncs_MinLengthOf = funcVal
        
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("MinLengthOf", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If
        
End Function

Private Function IBeTestFuncs_Something(ByVal testingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
    
    If IsObject(testingValue) Then
        funcVal = (Not testingValue Is Nothing)
    End If
    
    IBeTestFuncs_Something = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Something", funcVal, CBool(NegateValue), testingValue)
        
End Function

Private Function IBeTestFuncs_Between( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If IsNumeric(testingValue) Then
        funcVal = _
        (IBeTestFuncs_GreaterThanOrEqualTo(testingValue, LowerVal, updateFluentPath:=False) And _
        IBeTestFuncs_LessThanOrEqualTo(testingValue, HigherVal, updateFluentPath:=False))
    End If
    
    IBeTestFuncs_Between = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Between", funcVal, CBool(NegateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
End Function

Private Function IBeTestFuncs_OneOf(ByVal testingValue As Variant, ByVal NegateValue As Boolean, ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim tempBool As Boolean
    Dim funcVal As Variant
    
    funcVal = Null
            
    tempBool = False
    
    For i = LBound(testingInput) To UBound(testingInput)

        If IsObject(testingInput(i)) Or IsObject(testingValue) Then
            If TypeName(testingInput(i)) = TypeName(testingValue) Then
                tempBool = True
            End If
        ElseIf testingInput(i) = testingValue Then
            tempBool = True
        ElseIf IsNull(testingInput(i)) Or IsNull(testingValue) Then
            If TypeName(testingInput(i)) = TypeName(testingValue) Then
                tempBool = True
            End If
        End If
        
        If tempBool Then Exit For
    Next i
    
    funcVal = tempBool
    
    IBeTestFuncs_OneOf = funcVal
    
    If TypeOf testingValue Is Object  Then testingValue = TypeName(testingValue)
    
    Call pExpressions.addToTestAndFluentPath("OneOf", funcVal, CBool(NegateValue), testingValue, testingInput)
        
End Function

Private Function IHaveTestFuncs_LengthBetween( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal NegateValue As Boolean = False) As Variant
    Dim tempTestVal As Double
    Dim funcVal As Variant
    
    funcVal = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
'        tempTestVal = Len(CStr(testingValue))
        funcVal = _
        (IHaveTestFuncs_MinLengthOf(testingValue, LowerVal, updateFluentPath:=False) And _
        IHaveTestFuncs_MaxLengthOf(testingValue, HigherVal, updateFluentPath:=False))
    End If
    
    IHaveTestFuncs_LengthBetween = funcVal
    
    Call pExpressions.addToTestAndFluentPath("LengthBetween", funcVal, CBool(NegateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Private Function IShouldTestFuncs_EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal NegateValue = False) As Variant
    Dim funcVal As Variant
    
    funcVal = Null

    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        funcVal = (Excel.Evaluate(testingValue) = testingInput)
    End If
    
    IShouldTestFuncs_EvaluateTo = funcVal
    
    Call pExpressions.addToTestAndFluentPath("EvaluateTo", funcVal, CBool(NegateValue), testingValue, testingInput)
    
End Function

Private Function IBeTestFuncs_Alphabetic(ByVal testingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            End If
        Next i
        
        funcVal = (testingValue = tempAlph)
    End If
    
    IBeTestFuncs_Alphabetic = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Alphabetic", funcVal, CBool(NegateValue), testingValue)
End Function

Private Function IBeTestFuncs_Numeric(ByVal testingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    Dim funcVal As Variant
    
    funcVal = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            If tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        funcVal = (testingValue = tempNum)
    End If
    
    IBeTestFuncs_Numeric = funcVal
     
    Call pExpressions.addToTestAndFluentPath("Numeric", funcVal, CBool(NegateValue), testingValue)
End Function

Private Function IBeTestFuncs_Alphanumeric(ByVal testingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    Dim funcVal As Variant
    
    funcVal = Null
        
    b = False
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            ElseIf tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        If Len(tempAlph) > 1 And Len(tempNum) > 1 Then
            If Len(tempAlph) + Len(tempNum) = Len(testingValue) Then
                b = True
            End If
        End If
        
        funcVal = b
    End If
    
    IBeTestFuncs_Alphanumeric = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Alphanumeric", funcVal, CBool(NegateValue), testingValue)
End Function

Private Function IBeTestFuncs_Erroneous(ByVal testingValue As Variant, Optional ByVal NegateValue = False) As Variant
    Dim tempTestingVal As Variant
    Dim tempErrType As Variant
    Dim tempErrObj As ErrObject
    Dim funcVal As Variant
    
    funcVal = Null
        
    tempTestingVal = False
    
    If TypeName(testingValue) = "ErrObject" Or TypeName(testingValue) = "String" Then
        If TypeOf testingValue Is ErrObject Then
            Set tempErrObj = testingValue
            tempTestingVal = (tempErrObj.Number <> 0)
        ElseIf TypeName(testingValue) = "String" Then
            On Error Resume Next
                tempTestingVal = Excel.Evaluate(testingValue)
                tempTestingVal = (TypeName(tempTestingVal) = "Error")
            On Error GoTo 0
        End If
        
        funcVal = tempTestingVal
        
        If TypeOf testingValue Is Object  Then
            tempErrType = TypeName(testingValue)
        Else
            tempErrType = testingValue
        End If
    End If
    
    IBeTestFuncs_Erroneous = funcVal
    
    Call pExpressions.addToTestAndFluentPath("Erroneous", funcVal, CBool(NegateValue), tempErrType)
End Function

Private Function IHaveTestFuncs_ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue = False) As Variant
    Dim temp As Variant
    Dim tempErrObj As ErrObject
    Dim funcVal As Variant
    
    funcVal = Null
        
    If TypeName(testingValue) = "ErrObject" Or TypeName(testingValue) = "String" Then
        If TypeOf testingValue Is ErrObject Then
            Set tempErrObj = testingValue
            temp = tempErrObj.Number
        ElseIf TypeName(testingValue) = "String" Then
            On Error Resume Next
                temp = Excel.Evaluate(testingValue)
                Call Err.Raise(testingInput)
                temp = Err.Number
            On Error GoTo 0
        End If
        
        funcVal = (temp = testingInput)
    End If
    
    IHaveTestFuncs_ErrorNumberOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", funcVal, CBool(NegateValue), testingValue, testingInput)
End Function

Private Function IHaveTestFuncs_ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue = False) As Variant
    Dim tempTestVal As Variant
    Dim tempErrObj As ErrObject
    Dim funcVal As Variant
    
    funcVal = Null
        
    If TypeName(testingValue) = "ErrObject" Or TypeName(testingValue) = "String" Then
        If TypeOf testingValue Is ErrObject Then
            Set tempErrObj = testingValue
            tempTestVal = tempErrObj.Description
        ElseIf TypeName(testingValue) = "String" Then
            On Error Resume Next
                tempTestVal = CStr(Excel.Evaluate(testingValue))
                tempTestVal = Replace(tempTestVal, "Error ", "")
                Err.Raise tempTestVal
                tempTestVal = Err.Description
            On Error GoTo 0
        End If
        
        funcVal = (tempTestVal = testingInput)
    End If
    
    IHaveTestFuncs_ErrorDescriptionOf = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ErrorDescriptionOf", funcVal, CBool(NegateValue), tempTestVal, testingInput)
End Function

Private Function IHaveTestFuncs_SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal NegateValue = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim t1 As String
    Dim t2 As String
    Dim funcVal As Variant
    
    funcVal = Null
    
    t1 = TypeName(testingValue)
    t2 = TypeName(testingInput)
    
    funcVal = IBeTestFuncs_EqualTo(t1, t2, updateFluentPath:=False)
    
    IHaveTestFuncs_SameTypeAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameTypeAs", funcVal, CBool(NegateValue), testingValue, testingInput)
End Function

Private Function IBeTestFuncs_IdenticalTo( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant, _
ByVal updateFluentPath As Boolean) As Variant
    Dim funcVal As Variant
    
    funcVal = Null
        
    If pExpressions.IsDataStructure(testingValue) And pExpressions.IsDataStructure(testingInput) Then
        funcVal = _
        IBeTestFuncs_EqualTo( _
            pExpressions.InputToStringRecur(testingValue), _
            pExpressions.InputToStringRecur(testingInput), _
            updateFluentPath:=False _
        )
    End If
    
    IBeTestFuncs_IdenticalTo = funcVal
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("IdenticalTo", funcVal, CBool(NegateValue), testingValue, testingInput)
    End If

End Function

Private Function IHaveTestFuncs_ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim col2 As Collection
    Dim funcVal As Variant
    
    funcVal = Null
        
    If pExpressions.IsDataStructure(testingValue) And pExpressions.IsDataStructure(testingInput) Then
        Set col = New Collection
        Set col2 = New Collection
        
        For Each elem In testingValue
            col.Add elem
        Next elem
        
        For Each elem In testingInput
            col2.Add elem
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(col, NegateValue, col2, False)
    End If
    
    IHaveTestFuncs_ExactSameElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("ExactSameElementsAs", funcVal, CBool(NegateValue), testingValue, testingInput)

End Function

Private Function IHaveTestFuncs_SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim tempStr As String
    Dim funcVal As Variant
    
    funcVal = Null
        
    If pExpressions.IsDataStructure(testingValue) And pExpressions.IsDataStructure(testingInput) Then
        Set d = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = pExpressions.ToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = pExpressions.ToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            Else
                d(tempStr) = tempStr
            End If
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(d.Keys, NegateValue, d.Items, False)
    End If
    
    IHaveTestFuncs_SameUniqueElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameUniqueElementsAs", funcVal, CBool(NegateValue), testingValue, testingInput)

End Function

Private Function IHaveTestFuncs_SameElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim i As Long
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim tempStr As String
    Dim funcVal As Variant
    
    funcVal = Null
        
    If pExpressions.IsDataStructure(testingValue) And pExpressions.IsDataStructure(testingInput) Then
        i = 1
        elem = Empty
        Set d = New Scripting.Dictionary
        Set d2 = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = pExpressions.ToString(elem)
            
            If d.Exists(tempStr) Then
                d(tempStr) = d(tempStr) + 1
            Else
                d.Add pExpressions.ToString(elem), i
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = pExpressions.ToString(elem)
            
            If d2.Exists(tempStr) Then
                d2(tempStr) = d2(tempStr) + 1
            Else
                d2.Add pExpressions.ToString(elem), i
            End If
        Next elem
        
        funcVal = IBeTestFuncs_IdenticalTo(d, NegateValue, d2, False)
    End If
    
    IHaveTestFuncs_SameElementsAs = funcVal
    
    Call pExpressions.addToTestAndFluentPath("SameElementsAs", funcVal, CBool(NegateValue), testingValue, testingInput)

End Function
