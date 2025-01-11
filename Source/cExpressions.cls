VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IExpressions

Private pTestingInfo As ITestingFunctionsInfo
Private pTestingSubInfo As ITestingFunctionsSubInfo
Private pPrinting As IPrinting
Private iPrint As IPrintExpr
Private pMeta As IMeta
Private pTests As cTests
Private pTestDictCounter As Scripting.Dictionary

'Visual Basic errors (both Visual Basic-defined and user-defined errors) are in the range 0–65535.
'The range 0–512 is reserved for system errors; the range 513–65535 is available for user-defined errors.
'
'When setting the Number property to your own error code in a class module, you add your error code number
'to the vbObjectError constant. For example, to generate the error number 513, assign vbObjectError + 513
'to the Number property.

'source: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raise-method

Private Const NOT_IMPLEMENTED_ERROR As Long = vbObjectError + 514
Private Const SPACE_IN_FUNC_NAME_ERROR As Long = vbObjectError + 515
Private Const INVALID_ENUM_VALUE_ERROR As Long = vbObjectError + 516
Private Const RECUR_ITER_FUNC_ERROR As Long = vbObjectError + 517
Private Const RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR As Long = vbObjectError + 517

Private Property Set IExpressions_setPrinting(ByVal value As cPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Private Property Set IExpressions_setTests(ByVal value As cTests)
    Set pTests = value
End Property

Private Property Set IExpressions_setMeta(ByVal value As cMeta)
    Set pMeta = value
    Set pTestDictCounter = pMeta.Tests.TestingInfo.TestFuncInfoToDict
    Set pTestingInfo = pMeta.Tests.TestingInfo
End Property

Private Function setTestResultProps(ByVal test As ITest, ByVal result As Variant, ByVal negateValue As Boolean) As ITest
    
    If VBA.Information.IsNull(result) Then
        test.HasNull = True
    ElseIf VBA.Information.IsEmpty(result) Then
        test.HasEmpty = True
        test.TestValueSet = False
    End If

    If negateValue And Not VBA.Information.IsNull(result) And Not VBA.Information.IsEmpty(result) Then
        result = Not result
    End If
    
    test.result = result
    
    Set setTestResultProps = test
End Function

Private Function initTest(ByVal test As ITest, ByVal functionName As String, ByVal pPrinting As IPrinting) As ITest

    With test
        .functionName = functionName
        .testingInput = Empty
        .lowerVal = Null
        .higherVal = Null
        .StrTestInput = ""
        .strTestValue = ""
        .cleanedTestValue = ""
        .cleanedTestInput = ""
        .Category = pPrinting.Category
        .TestName = pPrinting.Name
        .HasCleanTestInput = False
        .HasCleanTestStrings = False
        .HasCleanTestValue = False
        .Actual = ""
    End With
    
    Set initTest = test
End Function

Private Function setTestValueTestInput(ByVal test As ITest, ByVal testingValue As Variant, ByVal testingInput As Variant) As ITest
    If VBA.Information.TypeName(testingValue) <> "Nothing" Then
        If TypeOf testingValue Is Object  Then
            Set test.testingValue = testingValue
        Else
            test.testingValue = testingValue
        End If
    End If
    
    If Not VBA.Information.IsMissing(testingInput) Then
        If VBA.Information.TypeName(testingInput) <> "Nothing" Then
            If TypeOf testingInput Is Object  Then
                Set test.testingInput = testingInput
            Else
                test.testingInput = testingInput
            End If
        End If
    End If
    
    Set setTestValueTestInput = test
End Function

Private Function setLowerValHigherVal(ByVal test As ITest, ByVal lowerVal As Variant, ByVal higherVal As Variant) As ITest
    test.lowerVal = lowerVal

    test.higherVal = higherVal
    
    test.testingInput = VBA.[_HiddenModule].Array(lowerVal, higherVal)
    
    Set setLowerValHigherVal = test
End Function

Private Function getCleanedTest(ByVal test As ITest, ByVal cleanedTestValue As Variant, ByVal cleanedTestInput As Variant) As ITest
    If Not VBA.Information.IsMissing(cleanedTestValue) Then
        If VBA.Information.TypeName(cleanedTestValue) = "String" Then
            If test.testingValue <> cleanedTestValue Then
                test.cleanedTestValue = cleanedTestValue
                test.HasCleanTestValue = True
            End If
        End If
    End If
    
    If Not VBA.Information.IsMissing(cleanedTestInput) Then
        If VBA.Information.TypeName(cleanedTestInput) = "String" Then
            If test.testingInput <> cleanedTestInput Then
                test.cleanedTestInput = cleanedTestInput
                test.HasCleanTestInput = True
            End If
        End If
    End If
    
    If test.HasCleanTestInput And test.HasCleanTestValue Then
        test.HasCleanTestStrings = True
    End If
    
    Set getCleanedTest = test
End Function

Private Function setStrTestValueStrTestInput(ByVal test As ITest, ByVal Meta As IMeta, ByVal negateValue As Boolean) As ITest
    Dim testDev As ITestDev
    Dim algo As flAlgorithm
    
    With test
        .StrTestInput = IExpressions_InputToString(.testingInput)
        .strTestValue = IExpressions_InputToString(.testingValue)
    End With
    
    Set testDev = test
    
    With pMeta.Tests
        If .ToStrDev Then
            algo = .Algorithm
            
            If algo = flAlgorithm.flRecursive Then
                testDev.TestValueRecur = test.strTestValue
                testDev.TestInputRecur = test.StrTestInput
                
                testDev.TestValueIter = IExpressions_InputToString(test.testingValue, flAlgorithm.flIterative)
                testDev.TestInputIter = IExpressions_InputToString(test.testingInput, flAlgorithm.flIterative)
            ElseIf algo = flAlgorithm.flIterative Then
                testDev.TestValueIter = test.strTestValue
                testDev.TestInputIter = test.StrTestInput
                
                testDev.TestValueRecur = IExpressions_InputToString(test.testingValue, flAlgorithm.flRecursive)
                testDev.TestInputRecur = IExpressions_InputToString(test.testingInput, flAlgorithm.flRecursive)
            Else
                Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
            End If
        End If
    End With
    
    testDev.negateValue = negateValue
    
    Set setStrTestValueStrTestInput = testDev
End Function

Private Sub IExpressions_addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal result As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant, _
    Optional ByVal cleanedTestValue As Variant, _
    Optional ByVal cleanedTestInput As Variant, _
    Optional ByVal Actual As Variant, _
    Optional ByVal recurIterFunc As Variant _
)

    'Consider refactoring to add a method to sets all test properties if not missing.
    'This can probably consolidate the actual property with the lowerval and higherval
    'properties into one method. And perhaps similar properties from some of the other
    'functions as well.
    
    Dim test As ITest
    Dim testDev As ITestDev
    
    Set test = New cTest
    Set testDev = test
    
    test.TestValueSet = True
    
    Set test = setTestResultProps(test, result, negateValue)

    Set test = initTest(test, functionName, pPrinting)
    
    If Not VBA.Information.IsMissing(Actual) Then
        test.Actual = Actual
    End If
    
    If pPrinting.Category <> "" Then iPrint.CategorySet = True
    If pPrinting.Name <> "" Then iPrint.NameSet = True

    Set test = setTestValueTestInput(test, testingValue, testingInput)

    If Not VBA.Information.IsMissing(lowerVal) And Not VBA.Information.IsMissing(higherVal) Then
        Set test = setLowerValHigherVal(test, lowerVal, higherVal)
    End If
    
    If Not VBA.Information.IsMissing(cleanedTestValue) Or Not VBA.Information.IsMissing(higherVal) Then
        Set test = getCleanedTest(test, cleanedTestValue, cleanedTestInput)
    End If
        
    Set test = setStrTestValueStrTestInput(test, pMeta, negateValue)
    
    If Not VBA.Information.IsMissing(recurIterFunc) Then
        If recurIterFunc Then
            testDev.Algorithm = pMeta.Tests.Algorithm
        
            testDev.AlgorithmValueSet = pMeta.Tests.AlgorithmValueSet
            
            testDev.IsRecurIterFunc = recurIterFunc
        End If
    End If
    
    Call IExpressions_addFluentPathNew(test, negateValue)

End Sub

Private Function IExpressions_Should(ByVal shouldType As flShould, Optional ByVal testingInput As Variant) As String
    Dim shouldStr As String
    
    shouldStr = ""
    
    Select Case shouldType
        Case Contain_
            shouldStr = "contain"
        Case EndWith_
            shouldStr = "endWith"
        Case EvaluateTo_
            shouldStr = "evaluate to"
        Case StartWith_
            shouldStr = "start with"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    If Not VBA.Information.IsMissing(testingInput) Then
        shouldStr = shouldStr & " " & testingInput
    End If

    IExpressions_Should = shouldStr
End Function

Private Function IExpressions_Be( _
ByVal beType As flBe, _
Optional ByVal TesingInput As String = "", _
Optional ByVal lowerVal As Variant, _
Optional ByVal higherVal As Variant) As String
    Dim beStr As String
    Dim funcStr As String
    
    beStr = "be "
    
    Select Case beType
        Case Alphabetic_
            funcStr = "alphabetic"
        Case Alphanumeric_
            funcStr = "alphanumeric"
        Case Between_
            funcStr = "between"
        Case Erroneous_
            funcStr = "erroneous"
        Case EqualTo_
            funcStr = "equal to"
        Case GreaterThan_
            funcStr = "greater than"
        Case GreaterThanOrEqualTo_
            funcStr = "greater than or equal to"
        Case IdenticalTo_
            funcStr = "identical to"
        Case InDataStructure_
            funcStr = "in data structure"
        Case InDataStructures_
            funcStr = "in data structures"
        Case LessThan_
            funcStr = "less than"
        Case LessThanOrEqualTo_
            funcStr = "less than or equal to"
        Case Numeric_
            funcStr = "numeric"
        Case OneOf_
            funcStr = "one of"
        Case Something_
            funcStr = "something"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
            
    End Select
    
    beStr = beStr & funcStr
    
    If TesingInput <> "" And VBA.Information.IsMissing(lowerVal) And VBA.Information.IsMissing(higherVal) Then
        beStr = beStr & " " & TesingInput
    End If
    
    If TesingInput = "" And Not VBA.Information.IsMissing(lowerVal) And Not VBA.Information.IsMissing(higherVal) Then
        beStr = beStr & " " & lowerVal & " and " & higherVal
    End If

    IExpressions_Be = beStr
End Function

Private Function IExpressions_Have( _
ByVal haveType As flHave, _
Optional ByVal TesingInput As String = "", _
Optional ByVal lowerVal As Variant, _
Optional ByVal higherVal As Variant) As String
    Dim haveStr As String
    Dim funcStr As String
    
    haveStr = "have "
    
    Select Case haveType
        Case DepthCountOf_
            funcStr = "depth count of"
        Case Elements_
            funcStr = "elements"
        Case ElementsInDataStructure_
            funcStr = "elements in data structure"
        Case ErrorDescriptionOf_
            funcStr = "error description of"
        Case ErrorNumberOf_
            funcStr = "error number of"
        Case ExactSameElementsAs_
            funcStr = "exact same elements as"
        Case LengthBetween_
            funcStr = "length between"
        Case LengthOf_
            funcStr = "length of"
        Case MaxLengthOf_
            funcStr = "max length of"
        Case MinLengthOf_
            funcStr = "min length of"
        Case NestedCountOf_
            funcStr = "nested count of"
        Case Procedure_
            funcStr = "procedure"
        Case SameElementsAs_
            funcStr = "same elements as"
        Case SameTypeAs_
            funcStr = "same type as"
        Case SameUniqueElementsAs_
            funcStr = "same unique elements as"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    haveStr = haveStr & funcStr
    
    If TesingInput <> "" And VBA.Information.IsMissing(lowerVal) And VBA.Information.IsMissing(higherVal) Then
        haveStr = haveStr & " " & TesingInput
    End If
    
    If TesingInput = "" And Not VBA.Information.IsMissing(lowerVal) And Not VBA.Information.IsMissing(higherVal) Then
        haveStr = haveStr & " " & lowerVal & " and " & higherVal
    End If

    IExpressions_Have = haveStr

End Function

Private Function IExpressions_buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
    Dim finalFluentPath As String
    
    If VBA.Information.IsNull(test.result) Or VBA.Information.IsEmpty(test.result) Then
        test.result = VBA.Information.TypeName(test.result)
    End If

    With test
        finalFluentPath = FluentPath & _
            "Testing value: " & .strTestValue & ";"
            
            If pMeta.Tests.TestStrings.CleanTestValueStr Then
                If .cleanedTestValue <> "" Then
                    finalFluentPath = finalFluentPath & " Cleaned test value: " & pMeta.ToString.WrapIfString(.cleanedTestValue)
                End If
            End If

            If .StrTestInput <> "" Then finalFluentPath = finalFluentPath & " Testing input: " & .StrTestInput & ";"
            
            If pMeta.Tests.TestStrings.CleanTestInputStr Then
                If .cleanedTestInput <> "" Then
                    finalFluentPath = finalFluentPath & " Cleaned test input: " & pMeta.ToString.WrapIfString(.cleanedTestInput)
                End If
            End If

        finalFluentPath = finalFluentPath & _
            " Expectation: Testing value " & tempFluentPath & _
            ";" & IIf(.Actual <> "", " Actual: " & pMeta.ToString.WrapIfString(.Actual) & ";", "") & _
            " Result: " & .result
    End With
    
    IExpressions_buildFluentPath = finalFluentPath
End Function

Private Function IExpressions_getFunctionExpr(ByVal test As ITest) As String
    Dim tempFluentPath As String
    Dim testingInput As String
    
    tempFluentPath = ""
    testingInput = ""
    
    testingInput = test.StrTestInput
    
    Select Case test.functionName
        Case "Alphabetic"
            tempFluentPath = IExpressions_Be(Alphabetic_)
        Case "Alphanumeric"
            tempFluentPath = IExpressions_Be(Alphanumeric_)
        Case "Between"
            tempFluentPath = IExpressions_Be(Between_, lowerVal:=test.lowerVal, higherVal:=test.higherVal)
        Case "Contain"
            tempFluentPath = IExpressions_Should(Contain_, testingInput)
        Case "DepthCountOf"
            tempFluentPath = IExpressions_Have(DepthCountOf_, testingInput)
        Case "Elements"
            tempFluentPath = IExpressions_Have(Elements_, testingInput)
        Case "ElementsInDataStructure"
            tempFluentPath = IExpressions_Have(ElementsInDataStructure_, testingInput)
        Case "Erroneous"
            tempFluentPath = IExpressions_Be(Erroneous_)
        Case "ErrorNumberOf"
            tempFluentPath = IExpressions_Have(ErrorNumberOf_, testingInput)
        Case "ErrorDescriptionOf"
            tempFluentPath = IExpressions_Have(ErrorDescriptionOf_, testingInput)
        Case "EqualTo"
            tempFluentPath = IExpressions_Be(EqualTo_, testingInput)
        Case "EndWith"
            tempFluentPath = IExpressions_Should(EndWith_, testingInput)
        Case "EvaluateTo"
            tempFluentPath = IExpressions_Should(EvaluateTo_, testingInput)
        Case "ExactSameElementsAs"
            tempFluentPath = IExpressions_Have(ExactSameElementsAs_, testingInput)
        Case "GreaterThan"
            tempFluentPath = IExpressions_Be(GreaterThan_, testingInput)
        Case "GreaterThanOrEqualTo"
            tempFluentPath = IExpressions_Be(GreaterThanOrEqualTo_, testingInput)
        Case "IdenticalTo"
            tempFluentPath = IExpressions_Be(IdenticalTo_, testingInput)
        Case "InDataStructure"
            tempFluentPath = IExpressions_Be(InDataStructure_, testingInput)
        Case "InDataStructures"
            tempFluentPath = IExpressions_Be(InDataStructures_, testingInput)
        Case "LengthBetween"
            tempFluentPath = IExpressions_Have(LengthBetween_, lowerVal:=test.lowerVal, higherVal:=test.higherVal)
        Case "LengthOf"
            tempFluentPath = IExpressions_Have(LengthOf_, testingInput)
        Case "LessThan"
            tempFluentPath = IExpressions_Be(LessThan_, testingInput)
        Case "LessThanOrEqualTo"
            tempFluentPath = IExpressions_Be(LessThanOrEqualTo_, testingInput)
        Case "MaxLengthOf"
            tempFluentPath = IExpressions_Have(MaxLengthOf_, testingInput)
        Case "MinLengthOf"
            tempFluentPath = IExpressions_Have(MinLengthOf_, testingInput)
        Case "NestedCountOf"
            tempFluentPath = IExpressions_Have(NestedCountOf_, testingInput)
        Case "Numeric"
            tempFluentPath = IExpressions_Be(Numeric_)
        Case "OneOf"
            tempFluentPath = IExpressions_Be(OneOf_, testingInput)
        Case "Procedure"
            tempFluentPath = IExpressions_Have(Procedure_, testingInput)
        Case "SameElementsAs"
            tempFluentPath = IExpressions_Have(SameElementsAs_, testingInput)
        Case "SameTypeAs"
            tempFluentPath = IExpressions_Have(SameTypeAs_, testingInput)
        Case "SameUniqueElementsAs"
            tempFluentPath = IExpressions_Have(SameUniqueElementsAs_, testingInput)
        Case "Something"
            tempFluentPath = IExpressions_Be(Something_)
        Case "StartWith"
            tempFluentPath = IExpressions_Should(StartWith_, testingInput)
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    IExpressions_getFunctionExpr = tempFluentPath
End Function

Private Function IExpressions_getFluentPath(ByVal FluentPath As String, ByVal test As ITest, Optional ByVal notShould As Boolean = False) As String
    Dim tempFluentPath As String
    
    tempFluentPath = ""
    
    tempFluentPath = IExpressions_getFunctionExpr(test)
    
    If notShould Then
        tempFluentPath = "should not " & tempFluentPath
    Else
        tempFluentPath = "should " & tempFluentPath
    End If
    
    tempFluentPath = IExpressions_buildFluentPath(test, FluentPath, tempFluentPath)
    
    IExpressions_getFluentPath = tempFluentPath
End Function

Private Sub IExpressions_addFluentPathNew(ByVal test As ITest, ByVal negateValue As Boolean)
    Dim FluentPath As String
    Dim finalFluentPath As String
    Dim fluentElement As Long
    Dim ResultNumber As Long
    Dim testDev As ITestDev
    Dim testInfoDev As ITestingFunctionsInfoDev
    Dim recurName As String
    Dim iterName As String
    Dim TestingSubInfoRecur As ITestingFunctionsSubInfo
    Dim TestingSubInfoIter As ITestingFunctionsSubInfo
    
    FluentPath = ""
    fluentElement = 1
    Set testDev = test
    recurName = ""
    iterName = ""
    
    If pTestDictCounter.Exists(test.functionName) Then
        Set pTestingSubInfo = pTestingInfo.getTestingFunction(test.functionName)
        
        If testDev.IsRecurIterFunc Then
            recurName = test.functionName & "Recur"
            iterName = test.functionName & "Iter"
            If pTestDictCounter.Exists(recurName) And pTestDictCounter.Exists(iterName) Then
                Set testInfoDev = pTestingInfo
                
                Call testInfoDev.addRecurIterFuncName(test.functionName)
            
                Set TestingSubInfoRecur = pTestingInfo.getTestingFunction(recurName)
                Set TestingSubInfoIter = pTestingInfo.getTestingFunction(iterName)
                
                TestingSubInfoRecur.Count = TestingSubInfoRecur.Count + 1
                TestingSubInfoIter.Count = TestingSubInfoIter.Count + 1
            Else
                Err.Raise Number:=RECUR_ITER_FUNC_ERROR, Description:="Recursive and/or iterative function does not entries in pTestDictCounter"
            End If
        End If
        
        pTestingSubInfo.Count = pTestingSubInfo.Count + 1
        
        If Not VBA.Information.IsNull(test.result) Then
            If test.result Then
                pTestingSubInfo.Passed = pTestingSubInfo.Passed + 1
            Else
                pTestingSubInfo.Failed = pTestingSubInfo.Failed + 1
            End If
        Else
            pTestingSubInfo.Unexpected = pTestingSubInfo.Unexpected + 1
        End If
    ElseIf pTestDictCounter.Exists(VBA.Strings.Trim(test.functionName)) Then
        Err.Raise Number:=SPACE_IN_FUNC_NAME_ERROR, Description:="Test name not matching in TestDictCounter due to extra space in the function name"
    Else
        Err.Raise Number:=NOT_IMPLEMENTED_ERROR, Description:="Test name not implemented in TestDictCounter"
    End If
    
    testDev.ResultNumber = IExpressions_getResultNumber(test.result)
    
    FluentPath = FluentPath & testDev.ResultNumber & "="
    
    FluentPath = IExpressions_getFluentPath(FluentPath, test, negateValue)
    
    finalFluentPath = VBA.Strings.Trim(VBA.Strings.Split(FluentPath, "=")(fluentElement))
    
    iPrint.FluentPath = finalFluentPath
    
    test.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add test
    
    Call pTests.CheckTest(test)
    
End Sub

Private Function IExpressions_InputToString(ByVal nv As Variant, Optional algo As Variant) As String
    Dim strArgs() As String
    Dim strObj As String
    Dim tempNVTS As String
    Dim i As Long
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim counter As Long
    Dim method As flAlgorithm
    Dim recurStr As String
    Dim iterStr As String
    Dim recurBool As Boolean
    Dim iterBool As Boolean
    
    i = 0
    counter = 0
    Set col = New VBA.Collection
    
    If IsMissing(algo) Then 'a value of 0 is equivalent to
        method = pMeta.Tests.Algorithm
    Else
        method = algo
    End If
    
    Select Case True
        Case pMeta.Tests.IsDataStructure(nv)
            If method And flAlgorithm.flRecursive Then
                recurStr = IExpressions_InputToStringRecur(nv)
            End If

            If method And flAlgorithm.flIterative Then
                iterStr = IExpressions_inputToStringIter(nv)
            End If
            
            If recurStr <> "" And iterStr <> "" Then
                If recurStr = iterStr Then
                    tempNVTS = recurStr
                Else
                    Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative strings are not identical"
                End If
            Else
                If recurStr <> "" Then
                    tempNVTS = IExpressions_InputToStringRecur(nv)
                ElseIf iterStr <> "" Then
                    tempNVTS = IExpressions_inputToStringIter(nv)
                Else
                    Err.Raise INVALID_ENUM_VALUE_ERROR, Description:="Enum value is not valid!"
                End If
            End If
        Case VBA.Information.IsNull(nv) Or VBA.Information.IsEmpty(nv)
            tempNVTS = VBA.Information.TypeName(nv)
        Case Not TypeOf nv Is Object
            If VBA.Information.TypeName(nv) = "String" Then
                nv = pMeta.ToString.WrapIfString(nv)
            End If
            tempNVTS = VBA.Conversion.CStr(nv)
        Case TypeOf nv Is VBA.ErrObject
            tempNVTS = VBA.Information.TypeName(nv)
        Case TypeOf nv Is Object
            If Not nv Is Nothing Then
                tempNVTS = VBA.Information.TypeName(nv)
            Else
                tempNVTS = "Nothing"
            End If
    End Select
    
    IExpressions_InputToString = tempNVTS
        
End Function

Private Function IExpressions_DatastructureIsEmpty(ByVal v As Variant) As Boolean
    Dim elem As Variant
    Dim counter As Long
    Dim b As Boolean
    Dim temp As Variant

    Select Case True
        Case TypeOf v Is Scripting.Dictionary Or TypeOf v Is Collection
            b = (v.Count = 0)
        
        Case VBA.Information.IsArray(v)
            On Error Resume Next
            
                '// will throw an error if v is initialized with no elements
                temp = UBound(v)
                
                b = (Err.Number <> 0)
                
                If Not b Then
                    For Each elem In v
                        If elem <> Empty Then
                            counter = counter + 1
                        End If
                    Next elem
                    
                    b = (counter = 0)
                End If
                
            On Error GoTo 0
    End Select
    
    IExpressions_DatastructureIsEmpty = b
End Function

Private Function IExpressions_InputToStringRecur(ByVal v As Variant) As String
    Dim temp As String
    Dim elem As Variant
    Dim emptyDS As Boolean

    temp = IExpressions_getDsType(v)

    Select Case True
        Case pMeta.Tests.IsDataStructure(v) And Not IExpressions_DatastructureIsEmpty(v)
            If Not TypeOf v Is Scripting.Dictionary Then
                For Each elem In v
                    If pMeta.Tests.IsDataStructure(elem) Then
                        temp = temp & IExpressions_InputToStringRecur(elem) & ", "
                    ElseIf VBA.Information.IsObject(elem) Then
                        temp = temp & VBA.Information.TypeName(elem) & ", "
                    ElseIf VBA.Information.TypeName(elem) = "String" Then
                        temp = temp & pMeta.ToString.WrapIfString(elem) & ", "
                    Else
                        temp = temp & elem & ", "
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If pMeta.Tests.IsDataStructure(v(elem)) Then
                        temp = temp & elem & ": " & IExpressions_InputToStringRecur(v(elem)) & ", "
                    ElseIf VBA.Information.IsObject(v(elem)) Then
                        temp = temp & elem & ": " & VBA.Information.TypeName(v(elem)) & ", "
                    ElseIf VBA.Information.TypeName(v(elem)) = "String" Then
                        temp = temp & elem & ": " & pMeta.ToString.WrapIfString(v(elem)) & ", "
                    Else
                        temp = temp & elem & ": " & v(elem) & ", "
                    End If
                Next elem
            End If
        Case Else
            temp = temp & elem & ", "
    End Select

    If VBA.Right$(temp, 2) = ", " Then
        temp = VBA.Left$(temp, VBA.Strings.Len(temp) - 2)
    End If

    temp = temp & ")"

    IExpressions_InputToStringRecur = temp
End Function

Function IExpressions_inputToStringIter(ByVal v As Variant) As String
    Dim elem As Variant
    Dim tempCol As VBA.Collection
    Dim tempCount As Long
    Dim finalCol As VBA.Collection
    Dim elemStr As String
    Dim temp As String
    Dim vCounter As String
    Dim dsCount As Long
    
    Set tempCol = New VBA.Collection
    tempCount = 1
    
    vCounter = IExpressions_getDsCount(v)
    
    temp = IExpressions_getDsType(v)
    
    tempCol.Add temp
    
    For Each elem In v
        If Not pMeta.Tests.IsDataStructure(elem) Then
            
            'This section is needed in case v is a dictionary
            'if so, then elem will be a key. And that key is
            'concatenated with a colon and the corresponding
            'value or object it is associated with
            
            If Not TypeOf v Is Scripting.Dictionary Then
                If VBA.Information.IsObject(elem) Then
                    tempCol.Add VBA.Information.TypeName(elem)
                ElseIf VBA.Information.TypeName(elem) = "String" Then
                    tempCol.Add pMeta.ToString.WrapIfString(elem)
                Else
                    tempCol.Add elem
                End If
            Else
                'v(elem) is a dictionary
                tempCol.Add elem & ": "
                
                If Not TypeOf v(elem) Is Scripting.Dictionary Then
                    tempCol.Add pMeta.ToString.WrapIfString(v(elem))
                Else
                    tempCol.Add v(elem)
                End If
            End If
        Else
            dsCount = IExpressions_getDsCount(elem)
            
            If dsCount > 0 Then
                tempCol.Add elem
            ElseIf dsCount = 0 Then
                tempCol.Add IExpressions_getDsType(elem) & ")"
            End If
        End If
        
        If tempCount < vCounter Then
            tempCol.Add ", "
        ElseIf tempCount = vCounter Then
            tempCol.Add ")"
        End If
        
        tempCount = tempCount + 1
    Next elem
    
    If tempCol.Count > 1 Then
        Set finalCol = IExpressions_getNestedElementsFromCollection(tempCol)
    ElseIf tempCol.Count = 1 Then
        'If tempCol.count equals 1, that means that the data structure
        'has no elements except for the typename of the data structure
        'and an open parentheses. So all you need to set finalCol to
        'tempCol and add a closing parentheses
        
        Set finalCol = tempCol
        finalCol.Add ")"
    End If
    
    For Each elem In finalCol
        elemStr = elemStr & elem
    Next elem
    
    IExpressions_inputToStringIter = elemStr
    
End Function

Function IExpressions_getNestedElementsFromCollection(ByVal col As VBA.Collection) As VBA.Collection
    Dim elem As Variant
    Dim elem2 As Variant
    Dim counter As Long
    Dim elemCount As Long
    Dim elemCounter As Long
    Dim tempCol As VBA.Collection
    Dim hasDataStructure As Boolean
    Dim d As Object
    
    Set tempCol = col
    counter = 1
    elemCounter = 1
    
    Do
        hasDataStructure = False

        If VBA.Information.IsObject(tempCol(counter)) Then
            Set elem = tempCol(counter)
        Else
            elem = tempCol(counter)
        End If
        
        If pMeta.Tests.IsDataStructure(elem) Then
            If IExpressions_getDsCount(elem) > 0 Then
                hasDataStructure = True
                
                tempCol.Add IExpressions_getDsType(elem), before:=counter
                
                'counter needs to get incremented every time an element is added
                'to tempCol or when the loop continues to the next element
                counter = counter + 1
                
                elemCount = IExpressions_getDsCount(elem)
                
                For Each elem2 In elem
                    If Not TypeOf elem Is Scripting.Dictionary Then
                        If VBA.Information.TypeName(elem2) = "String" Then
                            tempCol.Add pMeta.ToString.WrapIfString(elem2), before:=counter
                        Else
                            tempCol.Add elem2, before:=counter
                        End If
                        
                        'counter incremented towards end of if statement
                    Else
                        Set d = elem
                        
                        If pMeta.Tests.IsDataStructure(d(elem2)) Then
                            tempCol.Add elem2 & ": ", before:=counter
                            counter = counter + 1
                            
                            'counter incremented towards end of if statement
                            tempCol.Add d(elem2), before:=counter
                        ElseIf VBA.Information.TypeName(d(elem2)) = "String" Then
                            'counter incremented towards end of if statement
                            tempCol.Add elem2 & ": " & pMeta.ToString.WrapIfString(d(elem2)), before:=counter
                        Else
                            'counter incremented towards end of if statement
                            tempCol.Add elem2 & ": " & d(elem2), before:=counter
                        End If
                    End If
                
                    counter = counter + 1
                    
                    If elemCounter < elemCount Then
                    
                        tempCol.Add ", ", before:=counter
                    
                    ElseIf elemCounter >= elemCount Then
                    
                        tempCol.Add ")", before:=counter
                    
                    End If
                    
                    elemCounter = elemCounter + 1
                    
                    counter = counter + 1
                Next elem2
                
                tempCol.Remove counter
    
                counter = 0
                
                elemCounter = 1
            Else
                tempCol.Add IExpressions_getDsType(elem), before:=counter
                counter = counter + 1
                tempCol.Add ")", before:=counter
                counter = counter + 1
                tempCol.Remove counter
            End If
        End If
        
        counter = counter + 1

        If counter >= tempCol.Count And hasDataStructure = False Then Exit Do
    Loop While counter < tempCol.Count
    
    Set IExpressions_getNestedElementsFromCollection = tempCol
End Function

Public Function IExpressions_getDsCount(ByVal dataStructure As Variant) As Long
    Dim counter As Long
    Dim elem As Variant
    
    counter = 0
    
    For Each elem In dataStructure
        counter = counter + 1
    Next elem
    
    IExpressions_getDsCount = counter
End Function

Private Function IExpressions_getDsType(ByVal v As Variant) As String
    Dim counter As Long
    Dim elem As Variant
    Dim temp As String
    
    If VBA.Information.IsArray(v) Then
        temp = VBA.Strings.Left$(VBA.Information.TypeName(v), VBA.Strings.Len(VBA.Information.TypeName(v)) - 2) & "("
    ElseIf pMeta.Tests.IsDataStructure(v) Then
        temp = VBA.Information.TypeName(v) & "("
    ElseIf VBA.Information.IsObject(v) Then
        temp = VBA.Information.TypeName(v)
    End If
    
    IExpressions_getDsType = temp
End Function


Private Function IExpressions_getResultNumber(ByVal result As Variant) As Long
    Dim ResultNumber As Variant
    
    ResultNumber = Null
    
    If VBA.Information.IsNull(result) Or VBA.Information.IsEmpty(result) Then
        ResultNumber = 1
    ElseIf VBA.Information.TypeName(result) = "Boolean" Then
        ResultNumber = VBA.Conversion.CLng(result)
    End If
    
    IExpressions_getResultNumber = ResultNumber
End Function
