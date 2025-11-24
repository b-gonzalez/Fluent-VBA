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

Private pTestingInfos As ITestingFunctionsInfos
Private pTestingInfo As ITestingFunctionsInfo
Private pPrinting As IPrinting
Private iPrint As IPrintExpr
Private pMeta As IMeta
Private pTests As ITests
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

Private Property Set IExpressions_setTests(ByVal value As ITests)
    Set pTests = value
End Property

Private Property Set IExpressions_setMeta(ByVal value As IMeta)
    Set pMeta = value
    
    If Not pMeta Is Nothing Then
        Set pTestDictCounter = pMeta.tests.TestingFunctionsInfos.TestFuncInfoToDict
        Set pTestingInfos = pMeta.tests.TestingFunctionsInfos
    End If
End Property

Private Function setTestResultProps(ByVal test As ITestSet, ByVal result As Variant, ByVal negateValue As Boolean) As ITest
    
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

Private Function initTest(ByVal test As ITestSet, ByVal functionName As String, ByVal pPrinting As IPrinting) As ITest

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
    Dim testSet As ITestSet
    
    Set testSet = test
    
    If VBA.Information.TypeName(testingValue) <> "Nothing" Then
        If TypeOf testingValue Is Object  Then
            Set testSet.testingValue = testingValue
        Else
            testSet.testingValue = testingValue
        End If
    End If
    
    If Not VBA.Information.IsMissing(testingInput) Then
        If VBA.Information.TypeName(testingInput) <> "Nothing" Then
            If TypeOf testingInput Is Object  Then
                Set testSet.testingInput = testingInput
            Else
                testSet.testingInput = testingInput
            End If
        End If
    End If
    
    Set setTestValueTestInput = test
End Function

Private Function setLowerValHigherVal(ByVal test As ITestSet, ByVal lowerVal As Variant, ByVal higherVal As Variant) As ITest
    test.lowerVal = lowerVal

    test.higherVal = higherVal
    
    test.testingInput = VBA.[_HiddenModule].Array(lowerVal, higherVal)
    
    Set setLowerValHigherVal = test
End Function

Private Function getCleanedTest(ByVal test As ITest, ByVal cleanedTestValue As Variant, ByVal cleanedTestInput As Variant) As ITest
    Dim testSet As ITestSet
    
    Set testSet = test
    
    If Not VBA.Information.IsMissing(cleanedTestValue) Then
        If VBA.Information.TypeName(cleanedTestValue) = "String" Then
            If test.testingValue <> cleanedTestValue Then
                testSet.cleanedTestValue = cleanedTestValue
                testSet.HasCleanTestValue = True
            End If
        End If
    End If
    
    If Not VBA.Information.IsMissing(cleanedTestInput) Then
        If VBA.Information.TypeName(cleanedTestInput) = "String" Then
            If test.testingInput <> cleanedTestInput Then
                testSet.cleanedTestInput = cleanedTestInput
                testSet.HasCleanTestInput = True
            End If
        End If
    End If
    
    If test.HasCleanTestInput And test.HasCleanTestValue Then
        testSet.HasCleanTestStrings = True
    End If
    
    Set getCleanedTest = test
End Function

Private Function setStrTestValueStrTestInput(ByVal test As ITest, ByVal Meta As IMeta, ByVal negateValue As Boolean) As ITest
    Dim testDev As ITestDev
    Dim algo As Long
    Dim testSet As ITestSet
    
    Set testSet = test
    
    With testSet
        .StrTestInputPretty = IExpressions_InputToString(test.testingInput, indent:=True)
        .strTestValuePretty = IExpressions_InputToString(test.testingValue, indent:=True)
        .StrTestInput = deprettify(test.StrTestInputPretty)
        .strTestValue = deprettify(test.strTestValuePretty)
    End With
    
    Set testDev = test
    
    With pMeta.tests
        If .ToStrDev Then
            algo = .Algorithm
            
            If algo And flAlgorithm.flRecursive Then
                testDev.TestValueRecur = test.strTestValue
                testDev.TestInputRecur = test.StrTestInput
                
                testDev.TestValueIter = IExpressions_InputToString(test.testingValue, flAlgorithm.flIterative, False)
                testDev.TestInputIter = IExpressions_InputToString(test.testingInput, flAlgorithm.flIterative, False)
            ElseIf algo And flAlgorithm.flIterative Then
                testDev.TestValueIter = test.strTestValue
                testDev.TestInputIter = test.StrTestInput
                
                testDev.TestValueRecur = IExpressions_InputToString(test.testingValue, flAlgorithm.flRecursive, False)
                testDev.TestInputRecur = IExpressions_InputToString(test.testingInput, flAlgorithm.flRecursive, False)
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
    Optional ByVal recurIterFunc As Variant, _
    Optional ByVal TestingValueIsSelfReferential As Variant, _
    Optional ByVal TestingInputIsSelfReferential As Variant _
)

    'Consider refactoring to add a method to sets all test properties if not missing.
    'This can probably consolidate the actual property with the lowerval and higherval
    'properties into one method. And perhaps similar properties from some of the other
    'functions as well.
    
'    Dim test As ITest
    Dim test As ITestSet
    Dim testDev As ITestDev
    Dim tempTestingInput As Variant
    Dim tempTestingValue As Variant
    
    Set test = New cTest
    Set testDev = test
    
    test.TestingInputIsSelfReferential = Null
    test.TestingValueIsSelfReferential = Null
    test.HasSelfReferential = Null
    
    If VBA.Information.IsObject(testingInput) Then
        Set tempTestingInput = testingInput
    Else
        tempTestingInput = testingInput
    End If

    If VBA.Information.IsObject(testingValue) Then
        Set tempTestingValue = testingValue
    Else
        tempTestingValue = testingValue
    End If
    
    If Not VBA.Information.IsMissing(TestingValueIsSelfReferential) Then
        If Not TestingValueIsSelfReferential Then
            tempTestingValue = Null
        End If
    Else
        If pMeta.tests.IsDataStructure(testingValue) Then
            If pMeta.tests.DatastructureIsSelfReferential(testingValue, pMeta.tests.Algorithm) Then
                tempTestingValue = Null
            End If
        End If
    End If
    
    If VBA.Information.IsNull(tempTestingValue) Then test.TestingValueIsSelfReferential = True

    If Not VBA.Information.IsMissing(TestingInputIsSelfReferential) Then
        If Not TestingInputIsSelfReferential Then
            tempTestingInput = Null
        End If
    Else
        If pMeta.tests.IsDataStructure(testingInput) Then
            If pMeta.tests.DatastructureIsSelfReferential(testingInput, pMeta.tests.Algorithm) Then
                tempTestingInput = Null
            End If
        End If
    End If
    
    If VBA.Information.IsNull(tempTestingInput) Then test.TestingInputIsSelfReferential = True
    
    If VBA.Information.IsNull(tempTestingValue) Or VBA.Information.IsNull(tempTestingInput) Then
        test.HasSelfReferential = True
    End If
    
    test.TestValueSet = True
    
    Set test = setTestResultProps(test, result, negateValue)

    Set test = initTest(test, functionName, pPrinting)
    
    If Not VBA.Information.IsMissing(Actual) Then
        test.Actual = Actual
    End If
    
    If pPrinting.Category <> "" Then iPrint.CategorySet = True
    If pPrinting.Name <> "" Then iPrint.NameSet = True

    Set test = setTestValueTestInput(test, tempTestingValue, tempTestingInput)

    If Not VBA.Information.IsMissing(lowerVal) And Not VBA.Information.IsMissing(higherVal) Then
        Set test = setLowerValHigherVal(test, lowerVal, higherVal)
    End If
    
    If Not VBA.Information.IsMissing(cleanedTestValue) Or Not VBA.Information.IsMissing(higherVal) Then
        Set test = getCleanedTest(test, cleanedTestValue, cleanedTestInput)
    End If
        
    Set test = setStrTestValueStrTestInput(test, pMeta, negateValue)
    
    If Not VBA.Information.IsMissing(recurIterFunc) Then
        If recurIterFunc Then
            testDev.Algorithm = pMeta.tests.Algorithm
        
            testDev.AlgorithmValueSet = pMeta.tests.AlgorithmValueSet
            
            If testDev.Algorithm And flAlgorithm.flIterative Then
                testDev.IsBaseCaseIter = True
            ElseIf testDev.Algorithm And flAlgorithm.flRecursive Then
                testDev.IsBaseCaseRecur = True
            End If
            
            testDev.IsRecurIterFunc = recurIterFunc
                
            If Not VBA.Information.IsMissing(TestingInputIsSelfReferential) Then
                If Not TestingInputIsSelfReferential Then
                    test.HasSelfReferential = True
                End If
                
                test.TestingInputIsSelfReferential = Not TestingInputIsSelfReferential
            End If
            
            If Not VBA.Information.IsMissing(TestingValueIsSelfReferential) Then
                If Not TestingValueIsSelfReferential Then
                    test.HasSelfReferential = True
                End If
                
                test.TestingValueIsSelfReferential = Not TestingValueIsSelfReferential
            End If
        End If
    End If
    
    Call IExpressions_addFluentPathNew(test, negateValue)

End Sub

Private Function IExpressions_Should(ByVal shouldType As Long, Optional ByVal testingInput As Variant) As String
    Dim shouldStr As String
    
    shouldStr = ""
    
    Select Case shouldType
        Case flShould.Contain_
            shouldStr = "contain"
        Case flShould.EndWith_
            shouldStr = "endWith"
        Case flShould.EvaluateTo_
            shouldStr = "evaluate to"
        Case flShould.StartWith_
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
    ByVal beType As Long, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant _
) As String
    Dim beStr As String
    Dim funcStr As String
    
    beStr = "be "
    
    Select Case beType
        Case flBe.Alphabetic_
            funcStr = "alphabetic"
        Case flBe.Alphanumeric_
            funcStr = "alphanumeric"
        Case flBe.Between_
            funcStr = "between"
        Case flBe.Erroneous_
            funcStr = "erroneous"
        Case flBe.EqualTo_
            funcStr = "equal to"
        Case flBe.GreaterThan_
            funcStr = "greater than"
        Case flBe.GreaterThanOrEqualTo_
            funcStr = "greater than or equal to"
        Case flBe.IdenticalTo_
            funcStr = "identical to"
        Case flBe.InDataStructure_
            funcStr = "in data structure"
        Case flBe.InDataStructures_
            funcStr = "in data structures"
        Case flBe.LessThan_
            funcStr = "less than"
        Case flBe.LessThanOrEqualTo_
            funcStr = "less than or equal to"
        Case flBe.Numeric_
            funcStr = "numeric"
        Case flBe.OneOf_
            funcStr = "one of"
        Case flBe.Something_
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
    ByVal haveType As Long, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant _
) As String
    Dim haveStr As String
    Dim funcStr As String
    
    haveStr = "have "
    
    Select Case haveType
        Case flHave.DepthCountOf_
            funcStr = "depth count of"
        Case flHave.Elements_
            funcStr = "elements"
        Case flHave.ElementsInDataStructure_
            funcStr = "elements in data structure"
        Case flHave.ErrorDescriptionOf_
            funcStr = "error description of"
        Case flHave.ErrorNumberOf_
            funcStr = "error number of"
        Case flHave.ExactSameElementsAs_
            funcStr = "exact same elements as"
        Case flHave.LengthBetween_
            funcStr = "length between"
        Case flHave.LengthOf_
            funcStr = "length of"
        Case flHave.MaxLengthOf_
            funcStr = "max length of"
        Case flHave.MinLengthOf_
            funcStr = "min length of"
        Case flHave.NestedCountOf_
            funcStr = "nested count of"
        Case flHave.Procedure_
            funcStr = "procedure"
        Case flHave.SameElementsAs_
            funcStr = "same elements as"
        Case flHave.SameTypeAs_
            funcStr = "same type as"
        Case flHave.SameUniqueElementsAs_
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
    Dim testResult As String
    
    If VBA.Information.IsNull(test.result) Or VBA.Information.IsEmpty(test.result) Then
        testResult = VBA.Information.TypeName(test.result)
    Else
        testResult = test.result
    End If

    With test
        finalFluentPath = FluentPath & _
            "Testing value: " & .strTestValue & ";"
            
            If pMeta.tests.TestStrings.CleanTestValueStr Then
                If .cleanedTestValue <> "" Then
                    finalFluentPath = finalFluentPath & " Cleaned test value: " & pMeta.ToString.WrapIfString(.cleanedTestValue)
                End If
            End If

            If .StrTestInput <> "" Then finalFluentPath = finalFluentPath & " Testing input: " & .StrTestInput & ";"
            
            If pMeta.tests.TestStrings.CleanTestInputStr Then
                If .cleanedTestInput <> "" Then
                    finalFluentPath = finalFluentPath & " Cleaned test input: " & pMeta.ToString.WrapIfString(.cleanedTestInput)
                End If
            End If

        finalFluentPath = finalFluentPath & _
            " Expectation: Testing value " & tempFluentPath & _
            ";" & IIf(.Actual <> "", " Actual: " & pMeta.ToString.WrapIfString(.Actual) & ";", "") & _
            " Result: " & testResult
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
            tempFluentPath = IExpressions_Be(flBe.Alphabetic_)
        Case "Alphanumeric"
            tempFluentPath = IExpressions_Be(flBe.Alphanumeric_)
        Case "Between"
            tempFluentPath = IExpressions_Be(flBe.Between_, lowerVal:=test.lowerVal, higherVal:=test.higherVal)
        Case "Contain"
            tempFluentPath = IExpressions_Should(flShould.Contain_, testingInput)
        Case "DepthCountOf"
            tempFluentPath = IExpressions_Have(flHave.DepthCountOf_, testingInput)
        Case "Elements"
            tempFluentPath = IExpressions_Have(flHave.Elements_, testingInput)
        Case "ElementsInDataStructure"
            tempFluentPath = IExpressions_Have(flHave.ElementsInDataStructure_, testingInput)
        Case "Erroneous"
            tempFluentPath = IExpressions_Be(flBe.Erroneous_)
        Case "ErrorNumberOf"
            tempFluentPath = IExpressions_Have(flHave.ErrorNumberOf_, testingInput)
        Case "ErrorDescriptionOf"
            tempFluentPath = IExpressions_Have(flHave.ErrorDescriptionOf_, testingInput)
        Case "EqualTo"
            tempFluentPath = IExpressions_Be(flBe.EqualTo_, testingInput)
        Case "EndWith"
            tempFluentPath = IExpressions_Should(flShould.EndWith_, testingInput)
        Case "EvaluateTo"
            tempFluentPath = IExpressions_Should(flShould.EvaluateTo_, testingInput)
        Case "ExactSameElementsAs"
            tempFluentPath = IExpressions_Have(flHave.ExactSameElementsAs_, testingInput)
        Case "GreaterThan"
            tempFluentPath = IExpressions_Be(flBe.GreaterThan_, testingInput)
        Case "GreaterThanOrEqualTo"
            tempFluentPath = IExpressions_Be(flBe.GreaterThanOrEqualTo_, testingInput)
        Case "IdenticalTo"
            tempFluentPath = IExpressions_Be(flBe.IdenticalTo_, testingInput)
        Case "InDataStructure"
            tempFluentPath = IExpressions_Be(flBe.InDataStructure_, testingInput)
        Case "InDataStructures"
            tempFluentPath = IExpressions_Be(flBe.InDataStructures_, testingInput)
        Case "LengthBetween"
            tempFluentPath = IExpressions_Have(flHave.LengthBetween_, lowerVal:=test.lowerVal, higherVal:=test.higherVal)
        Case "LengthOf"
            tempFluentPath = IExpressions_Have(flHave.LengthOf_, testingInput)
        Case "LessThan"
            tempFluentPath = IExpressions_Be(flBe.LessThan_, testingInput)
        Case "LessThanOrEqualTo"
            tempFluentPath = IExpressions_Be(flBe.LessThanOrEqualTo_, testingInput)
        Case "MaxLengthOf"
            tempFluentPath = IExpressions_Have(flHave.MaxLengthOf_, testingInput)
        Case "MinLengthOf"
            tempFluentPath = IExpressions_Have(flHave.MinLengthOf_, testingInput)
        Case "NestedCountOf"
            tempFluentPath = IExpressions_Have(flHave.NestedCountOf_, testingInput)
        Case "Numeric"
            tempFluentPath = IExpressions_Be(flBe.Numeric_)
        Case "OneOf"
            tempFluentPath = IExpressions_Be(flBe.OneOf_, testingInput)
        Case "Procedure"
            tempFluentPath = IExpressions_Have(flHave.Procedure_, testingInput)
        Case "SameElementsAs"
            tempFluentPath = IExpressions_Have(flHave.SameElementsAs_, testingInput)
        Case "SameTypeAs"
            tempFluentPath = IExpressions_Have(flHave.SameTypeAs_, testingInput)
        Case "SameUniqueElementsAs"
            tempFluentPath = IExpressions_Have(flHave.SameUniqueElementsAs_, testingInput)
        Case "Something"
            tempFluentPath = IExpressions_Be(flBe.Something_)
        Case "StartWith"
            tempFluentPath = IExpressions_Should(flShould.StartWith_, testingInput)
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
    Dim TestingInfoRecur As ITestingFunctionsInfo
    Dim TestingInfoIter As ITestingFunctionsInfo
    Dim testSet As ITestSet
    
    Set testSet = test
    
    FluentPath = ""
    fluentElement = 1
    Set testDev = test
    recurName = ""
    iterName = ""
    
    If pTestDictCounter.Exists(test.functionName) Then
        Set pTestingInfo = pTestingInfos.getTestingFunction(test.functionName)
        
        pTestingInfo.Name = test.functionName
        
        If testDev.IsRecurIterFunc Then
            recurName = test.functionName & "Recur"
            iterName = test.functionName & "Iter"
            If pTestDictCounter.Exists(recurName) And pTestDictCounter.Exists(iterName) Then
                Set testInfoDev = pTestingInfos
                
                Call testInfoDev.addRecurIterFuncName(test.functionName)
            
                Set TestingInfoRecur = pTestingInfos.getTestingFunction(recurName)
                Set TestingInfoIter = pTestingInfos.getTestingFunction(iterName)
                
                TestingInfoRecur.Count = TestingInfoRecur.Count + 1
                TestingInfoIter.Count = TestingInfoIter.Count + 1
                
                If pTestingInfos.AddDevTestingFunctionsInfo Then
                    With pTestingInfos
                        .AddTestingFunctionsInfo .getTestingFunction(recurName)
                        .AddTestingFunctionsInfo .getTestingFunction(iterName)
                    End With
                End If
            Else
                Err.Raise Number:=RECUR_ITER_FUNC_ERROR, Description:="Recursive and/or iterative function does not entries in pTestDictCounter"
            End If
        End If
        
        pTestingInfo.Count = pTestingInfo.Count + 1
        
        If Not VBA.Information.IsNull(test.result) Then
            If test.result Then
                pTestingInfo.Passed = pTestingInfo.Passed + 1
            Else
                pTestingInfo.Failed = pTestingInfo.Failed + 1
            End If
        Else
            pTestingInfo.Unexpected = pTestingInfo.Unexpected + 1
        End If
        
        pTestingInfos.AddTestingFunctionsInfo pTestingInfo
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
    
    testSet.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add test
    
    Call pTests.CheckTest(test)
    
End Sub

Private Function IExpressions_InputToString( _
    ByVal nv As Variant, _
    Optional ByVal algo As Variant, _
    Optional ByVal indent As Boolean = False _
) As String
    Dim strArgs() As String
    Dim strObj As String
    Dim tempNVTS As String
    Dim i As Long
    Dim elem As Variant
    Dim col As VBA.Collection
    Dim counter As Long
    Dim method As Long
    Dim recurStr As String
    Dim iterStr As String
    Dim recurBool As Boolean
    Dim iterBool As Boolean
    Dim recurCol As Collection
    Dim iterCol As Collection
    
    i = 0
    counter = 0
    Set col = New VBA.Collection
    
    If IsMissing(algo) Then 'a value of 0 is equivalent to
        method = pMeta.tests.Algorithm
    Else
        method = algo
    End If
    
    Select Case True
        Case pMeta.tests.IsDataStructure(nv)
            If method And flAlgorithm.flRecursive Then
                Set recurCol = IExpressions_InputToStringRecur(nv, indent:=indent)
                recurStr = Join(collectionToStrArray(recurCol), "")
            End If

            If method And flAlgorithm.flIterative Then
                Set iterCol = IExpressions_inputToStringIter(nv, indent:=indent)
                iterStr = Join(collectionToStrArray(iterCol), "")
            End If
            
            If recurStr <> "" And iterStr <> "" Then
                If recurStr = iterStr Then
                    tempNVTS = recurStr
                Else
                    Err.Raise RECUR_ITER_INPUTS_NOT_IDENTICAL_ERROR, Description:="Recursive and Iterative strings are not identical"
                End If
            Else
                If recurStr <> "" Then
                    Set recurCol = IExpressions_InputToStringRecur(nv, indent:=indent)
                    tempNVTS = Join(collectionToStrArray(recurCol), "")
                ElseIf iterStr <> "" Then
                    Set iterCol = IExpressions_inputToStringIter(nv, indent:=indent)
                    tempNVTS = Join(collectionToStrArray(iterCol), "")
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

Private Function IExpressions_InputToStringRecur( _
    ByVal v As Variant, _
    Optional ByVal col As Collection = Nothing, _
    Optional ByVal counter As Long = 1, _
    Optional ByVal indent As Boolean = False _
) As Collection
    Dim elem As Variant
    Dim emptyDS As Boolean
    Dim dsCount As Long
    Dim tempCount As Long
    Dim arr() As String
    Dim i As Long
    
    If col Is Nothing Then
        Set col = New Collection
    End If

    col.Add getDsType(v)

    If indent Then
        col.Add vbNewLine
    End If
    
    tempCount = 1

    Select Case True
        Case pMeta.tests.IsDataStructure(v) And Not IExpressions_DatastructureIsEmpty(v)
            dsCount = IExpressions_getDsCount(v)
                        
            If Not TypeOf v Is Scripting.Dictionary Then
                For Each elem In v
                    If indent Then
                        col.Add getIndent(counter)
                    End If
                                        
                    If pMeta.tests.IsDataStructure(elem) Then
                        Set col = IExpressions_InputToStringRecur(elem, col, counter + 1, indent)
                    ElseIf VBA.Information.IsObject(elem) Then
                        col.Add VBA.Information.TypeName(elem)
                    ElseIf VBA.Information.TypeName(elem) = "String" Then
                        col.Add pMeta.ToString.WrapIfString(elem)
                    ElseIf VBA.Information.TypeName(elem) = "Date" Then
                        col.Add "#" & elem & "#"
                    Else
                        col.Add elem
                    End If
                    
                    If tempCount < dsCount Then
                        col.Add ", "
                        tempCount = tempCount + 1
                    End If
                                        
                    If indent Then
                        col.Add vbNewLine
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If indent Then
                        col.Add getIndent(counter)
                    End If
                    
                    If VBA.Information.TypeName(elem) = "String" Then
                        col.Add pMeta.ToString.WrapIfString(elem)
                    Else
                        col.Add elem
                    End If
                    
                    col.Add ":"
                    
                    If pMeta.tests.IsDataStructure(v(elem)) Then
                        Set col = IExpressions_InputToStringRecur(v(elem), col, counter + 1, indent)
                    ElseIf VBA.Information.IsObject(v(elem)) Then
                        col.Add VBA.Information.TypeName(v(elem))
                    ElseIf VBA.Information.TypeName(v(elem)) = "String" Then
                        col.Add pMeta.ToString.WrapIfString(v(elem))
                    ElseIf VBA.Information.TypeName(v(elem)) = "Date" Then
                        col.Add "#" & v(elem) & "#"
                    Else
                        col.Add v(elem)
                    End If
                    
                    If tempCount < dsCount Then
                        col.Add ", "
                        tempCount = tempCount + 1
                    End If
                    
                    If indent Then
                        col.Add vbNewLine
                    End If
                    
                Next elem
            End If
        Case Else
            col.Add elem
            
            If tempCount < dsCount Then
                col.Add ", "
                tempCount = tempCount + 1
            End If
    End Select

    If indent And counter > 1 Then
        col.Add getIndent(counter - 1)
    End If
    
    col.Add ")"

    Set IExpressions_InputToStringRecur = col
End Function

Private Function getCollectionFromDataStructure( _
    ByVal ds As Variant, _
    ByVal col1 As VBA.Collection, _
    Optional ByVal counter As Long = 1, _
    Optional ByVal indent As Boolean = False _
) As VBA.Collection
    Dim elem As Variant
    Dim tempCol As VBA.Collection
    Dim parentDsCount As Long
    Dim childDsCount As Long

    parentDsCount = 1
    childDsCount = 1
    parentDsCount = IExpressions_getDsCount(ds)
    
    Set tempCol = col1
    
    If indent Then
        tempCol.Add getDsType(ds) & vbNewLine
    Else
        tempCol.Add getDsType(ds)
    End If

    If parentDsCount > 0 Then
        If TypeOf ds Is Scripting.Dictionary Then
            For Each elem In ds.Keys
                If indent Then
                    tempCol.Add getIndent(counter)
                End If
                
                If VBA.Information.TypeName(elem) = "String" Then
                    tempCol.Add pMeta.ToString.WrapIfString(elem) & ":"
                Else
                    tempCol.Add elem & ":"
                End If
                                
                If VBA.Information.TypeName(ds(elem)) = "String" Then
                    tempCol.Add pMeta.ToString.WrapIfString(ds(elem))
                ElseIf VBA.Information.TypeName(ds(elem)) = "Date" Then
                    tempCol.Add "#" & ds(elem) & "#"
                Else
                    tempCol.Add ds(elem)
                End If

                If childDsCount < parentDsCount Then
                    tempCol.Add ", "
                End If
                
                If indent Then
                    tempCol.Add vbNewLine
                End If

                childDsCount = childDsCount + 1
            Next elem
        Else
            For Each elem In ds
                If indent Then
                    tempCol.Add getIndent(counter)
                End If
                
                If VBA.Information.TypeName(elem) = "String" Then
                    tempCol.Add pMeta.ToString.WrapIfString(elem)
                ElseIf VBA.Information.TypeName(elem) = "Date" Then
                    tempCol.Add "#" & elem & "#"
                Else
                    tempCol.Add elem
                End If

                If childDsCount < parentDsCount Then
                    tempCol.Add ", "
                End If
                
                If indent Then
                    tempCol.Add vbNewLine
                End If

                childDsCount = childDsCount + 1
            Next elem
        End If
    End If
    
    If counter > 1 Then
        If indent Then
            tempCol.Add getIndent(counter - 1)
        End If
    End If
    
    tempCol.Add ")"

    Set getCollectionFromDataStructure = tempCol
End Function

Function IExpressions_inputToStringIter(ByVal ds As Variant, Optional ByVal indent As Boolean = False) As Collection
    Dim col1 As VBA.Collection
    Dim tempCol As VBA.Collection
    Dim elem As Variant
    Dim hasDataStructure As Boolean
    Dim finalStr As String
    Dim indentCounter As Long
    
    Set col1 = New VBA.Collection
    Set tempCol = New VBA.Collection
    indentCounter = 0

    col1.Add ds
    
    Do
        hasDataStructure = False
            
        If IExpressions_getDsCount(col1) > 0 Then
            For Each elem In col1
                If pMeta.tests.IsDataStructure(elem) Then
                    hasDataStructure = True
                    
                    indentCounter = indentCounter + 1
                    
                    Set tempCol = getCollectionFromDataStructure(elem, tempCol, indentCounter, indent)
                ElseIf IsObject(elem) Then
                    tempCol.Add getDsType(elem)
                Else
                    tempCol.Add elem
                End If

            Next elem
        End If
    
        Set col1 = tempCol
        Set tempCol = New VBA.Collection
    Loop While hasDataStructure = True
    
'    For Each elem In col1
'        finalStr = finalStr & elem
'    Next elem
    
    Set IExpressions_inputToStringIter = col1
End Function

Function getNestedElementsFromCollection(ByVal col As VBA.Collection) As VBA.Collection
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
        
        If pMeta.tests.IsDataStructure(elem) Then
            If IExpressions_getDsCount(elem) > 0 Then
                hasDataStructure = True
                
                tempCol.Add getDsType(elem), before:=counter
                
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
                        
                        If pMeta.tests.IsDataStructure(d(elem2)) Then
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
                tempCol.Add getDsType(elem), before:=counter
                counter = counter + 1
                tempCol.Add ")", before:=counter
                counter = counter + 1
                tempCol.Remove counter
            End If
        End If
        
        counter = counter + 1

        If counter >= tempCol.Count And hasDataStructure = False Then Exit Do
    Loop While counter < tempCol.Count
    
    Set getNestedElementsFromCollection = tempCol
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

Private Function getDsType(ByVal v As Variant) As String
    Dim counter As Long
    Dim elem As Variant
    Dim temp As String
    
    If VBA.Information.IsArray(v) Then
        temp = VBA.Strings.Left$(VBA.Information.TypeName(v), VBA.Strings.Len(VBA.Information.TypeName(v)) - 2) & "("
    ElseIf pMeta.tests.IsDataStructure(v) Then
        temp = VBA.Information.TypeName(v) & "("
    ElseIf VBA.Information.IsObject(v) Then
        temp = VBA.Information.TypeName(v)
    End If
    
    getDsType = temp
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

Private Function getIndent(ByVal num As Long) As String
    ReDim arr(num - 1) As Variant
    
    Dim i As Long
    
    For i = 0 To num - 1
'        arr(i) = String$(4, " ")
        arr(i) = vbTab
    Next i
    
    getIndent = Join(arr, "")
End Function

Private Function deprettify(ByVal str As String) As String
    Dim temp As String

    temp = str
    temp = Replace(temp, vbNewLine, "")
    temp = Replace(temp, vbTab, "")

    deprettify = temp
End Function

Private Function collectionToStrArray(ByVal col As Collection) As String()
    Dim arr() As String
    Dim i As Long
    
    ReDim arr(col.Count - 1)
    
    For i = 1 To col.Count
        If VBA.Information.IsNull(col(i)) Then
            arr(i - 1) = "Null"
        Else
            arr(i - 1) = col(i)
        End If
    Next i
    
    collectionToStrArray = arr
End Function

Private Sub Class_Terminate()
    Set pTestDictCounter = Nothing
End Sub
