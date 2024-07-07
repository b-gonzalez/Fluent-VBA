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

Private Const NOT_IMPLEMENTED_ERROR As Long = 514

Private Property Set IExpressions_setPrinting(ByVal value As cPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Private Property Set IExpressions_setTests(ByVal value As cTests)
    Set pTests = value
End Property

Private Property Set IExpressions_setMeta(ByVal value As cMeta)
    Set pMeta = value
    Set pTestDictCounter = pMeta.Tests.testingInfo.TestFuncInfoToDict
    Set pTestingInfo = pMeta.Tests.testingInfo
End Property

Private Sub IExpressions_addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal Result As Variant, _
    ByVal NegateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal LowerVal As Variant, _
    Optional ByVal HigherVal As Variant, _
    Optional cleanedTestValue As Variant, _
    Optional cleanedTestInput As Variant)
    
    Dim test As ITest
    Dim testDev As ITestDev
    Dim shouldBool As String
    Dim algo As flAlgorithm
    
    Set test = New cTest
    Set testDev = test
    
    test.TestValueSet = True
    
    If IsNull(Result) Then
'        result = TypeName(result)
        test.HasNull = True
    ElseIf IsEmpty(Result) Then
        test.HasEmpty = True
        test.TestValueSet = False
    End If

    If NegateValue And Not IsNull(Result) And Not IsEmpty(Result) Then
        Result = Not Result
    End If

    With test
        .functionName = functionName
        .Result = Result
        .testingInput = Empty
        .LowerVal = Empty
        .HigherVal = Empty
        .StrTestInput = ""
        .StrTestValue = ""
        .cleanedTestValue = ""
        .cleanedTestInput = ""
        .Category = pPrinting.Category
        .TestName = pPrinting.Name
    End With
    
    If pPrinting.Category <> "" Then iPrint.CategorySet = True
    If pPrinting.Name <> "" Then iPrint.NameSet = True
        
    If TypeName(testingValue) <> "Nothing" Then
        If TypeOf testingValue Is Object  Then
            Set test.testingValue = testingValue
        Else
            test.testingValue = testingValue
        End If
    End If
    
    If Not IsMissing(testingInput) Then
        If TypeName(testingInput) <> "Nothing" Then
            If TypeOf testingInput Is Object  Then
                Set test.testingInput = testingInput
            Else
                test.testingInput = testingInput
            End If
        End If
    End If

    If Not IsMissing(LowerVal) Then
        test.LowerVal = LowerVal
    End If

    If Not IsMissing(HigherVal) Then
        test.HigherVal = HigherVal
    End If
    
    If Not IsMissing(cleanedTestValue) Then
        If TypeName(cleanedTestValue) = "String" Then
            If test.testingValue <> cleanedTestValue Then test.cleanedTestValue = cleanedTestValue
        End If
    End If
    
    If Not IsMissing(cleanedTestInput) Then
        If TypeName(cleanedTestInput) = "String" Then
            If test.testingInput <> cleanedTestInput Then test.cleanedTestInput = cleanedTestInput
        End If
    End If
    
    If Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        test.testingInput = Array(LowerVal, HigherVal)
    End If
    
    With test
        .StrTestInput = IExpressions_ToString(.testingInput)
        .StrTestValue = IExpressions_ToString(.testingValue)
    End With
    

    
    '//TestDev
    
    With pMeta.Tests
        If .ToStrDev Then
            algo = .Algorithm
            
            If algo = flRecursive Then
                testDev.TestValueRecur = test.StrTestValue
                testDev.TestInputRecur = test.StrTestInput
                
                pMeta.Tests.Algorithm = flIterative
                
                testDev.TestValueIter = IExpressions_ToString(test.testingValue)
                testDev.TestInputIter = IExpressions_ToString(test.testingInput)
            ElseIf algo = flIterative Then
                testDev.TestValueIter = test.StrTestValue
                testDev.TestInputIter = test.StrTestInput
                
                pMeta.Tests.Algorithm = flRecursive
                
                testDev.TestValueRecur = IExpressions_ToString(test.testingValue)
                testDev.TestInputRecur = IExpressions_ToString(test.testingInput)
            End If
        End If
    End With
    
    pMeta.Tests.Algorithm = algo
    
    testDev.NegateValue = NegateValue
    
    Call IExpressions_addFluentPathNew(test, NegateValue)

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
    
    If Not IsMissing(testingInput) Then
        shouldStr = shouldStr & " " & testingInput
    End If

    IExpressions_Should = shouldStr
End Function

Private Function IExpressions_Be( _
ByVal beType As flBe, _
Optional ByVal TesingInput As String = "", _
Optional ByVal LowerVal As Variant, _
Optional ByVal HigherVal As Variant) As String
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
    
    If TesingInput <> "" And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        beStr = beStr & " " & TesingInput
    End If
    
    If TesingInput = "" And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        beStr = beStr & " " & LowerVal & " and " & HigherVal
    End If

    IExpressions_Be = beStr
End Function

Private Function IExpressions_Have( _
ByVal haveType As flHave, _
Optional ByVal TesingInput As String = "", _
Optional ByVal LowerVal As Variant, _
Optional ByVal HigherVal As Variant) As String
    Dim haveStr As String
    Dim funcStr As String
    
    haveStr = "have "
    
    Select Case haveType
        Case ErrorDescriptionOf_
            funcStr = "error description of"
        Case ErrorNumberOf_
            funcStr = "error number of"
        Case LengthBetween_
            funcStr = "length between"
        Case LengthOf_
            funcStr = "length of"
        Case MaxLengthOf_
            funcStr = "max length of"
        Case MinLengthOf_
            funcStr = "min length of"
        Case ExactSameElementsAs_
            funcStr = "exact same elements as"
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
    
    If TesingInput <> "" And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        haveStr = haveStr & " " & TesingInput
    End If
    
    If TesingInput = "" And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        haveStr = haveStr & " " & LowerVal & " and " & HigherVal
    End If

    IExpressions_Have = haveStr

End Function

Private Function IExpressions_buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
    Dim finalFluentPath As String
    
    If Information.IsNull(test.Result) Or Information.IsEmpty(test.Result) Then
        test.Result = TypeName(test.Result)
    End If

    With test
        finalFluentPath = FluentPath & _
            "Testing value: " & .StrTestValue & ";"
            
            If pMeta.Tests.CleanTestValStr Then
                If .cleanedTestValue <> "" Then finalFluentPath = finalFluentPath & " Cleaned test value: `" & .cleanedTestValue & "`;"
            End If

            If .StrTestInput <> "" Then finalFluentPath = finalFluentPath & " Testing input: " & .StrTestInput & ";"
            
            If pMeta.Tests.CleanTestInputStr Then
                If .cleanedTestInput <> "" Then finalFluentPath = finalFluentPath & " Cleaned test input: `" & .cleanedTestInput & "`;"
            End If

        finalFluentPath = finalFluentPath & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .Result
    End With
    
    IExpressions_buildFluentPath = finalFluentPath
End Function

Private Function IExpressions_getFunctionExpr(ByVal test As ITest) As String
    Dim tempFluentPath As String
    Dim testingInput As String
    
    tempFluentPath = ""
    testingInput = ""
    
    If TypeName(test.testingInput) = "String" Then
        testingInput = "`" & test.StrTestInput & "`"
    Else
        testingInput = test.StrTestInput
    End If
    
    Select Case test.functionName
        Case "Alphabetic"
            tempFluentPath = IExpressions_Be(Alphabetic_)
        Case "Alphanumeric"
            tempFluentPath = IExpressions_Be(Alphanumeric_)
        Case "Between"
            tempFluentPath = IExpressions_Be(Between_, LowerVal:=test.LowerVal, HigherVal:=test.HigherVal)
        Case "Contain"
            tempFluentPath = IExpressions_Should(Contain_, testingInput)
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
            tempFluentPath = IExpressions_Have(LengthBetween_, LowerVal:=test.LowerVal, HigherVal:=test.HigherVal)
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
        Case "Numeric"
            tempFluentPath = IExpressions_Be(Numeric_)
        Case "OneOf"
            tempFluentPath = IExpressions_Be(OneOf_, testingInput)
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
'    Dim LowerVal As Variant
'    Dim HigherVal As Variant
    Dim tempFluentPath As String
    
    tempFluentPath = ""
    
'    If test.LowerVal <> Empty And test.HigherVal <> Empty Then
'        LowerVal = test.LowerVal
'        HigherVal = test.HigherVal
'    End If

    tempFluentPath = IExpressions_getFunctionExpr(test)
    
    If notShould Then
        tempFluentPath = "should not " & tempFluentPath
    Else
        tempFluentPath = "should " & tempFluentPath
    End If
    
    tempFluentPath = IExpressions_buildFluentPath(test, FluentPath, tempFluentPath)
    
    IExpressions_getFluentPath = tempFluentPath
End Function

Private Sub IExpressions_addFluentPathNew(ByVal test As ITest, ByVal NegateValue As Boolean)
    Dim FluentPath As String
    Dim finalFluentPath As String
    Dim fluentElement As Long
    Dim resultNumber As Long
    Dim testDev As ITestDev
    
    FluentPath = ""
    fluentElement = 1
    Set testDev = test
    
    If pTestDictCounter.Exists(test.functionName) Then
        Set pTestingSubInfo = pTestingInfo.getTestingFunction(test.functionName)
        
        pTestingSubInfo.Count = pTestingSubInfo.Count + 1
        
        If Not IsNull(test.Result) Then
            If test.Result Then
                pTestingSubInfo.Passed = pTestingSubInfo.Passed + 1
            Else
                pTestingSubInfo.Failed = pTestingSubInfo.Failed + 1
            End If
        Else
            pTestingSubInfo.Unexpected = pTestingSubInfo.Unexpected + 1
        End If
    Else
        Err.Raise Number:=NOT_IMPLEMENTED_ERROR, Description:="Test name not implemented in TestDictCounter"
    End If
    
    testDev.resultNumber = IExpressions_getResultNumber(test.Result)
    
    FluentPath = FluentPath & testDev.resultNumber & "="
    
    FluentPath = IExpressions_getFluentPath(FluentPath, test, NegateValue)
    
    finalFluentPath = Trim(Split(FluentPath, "=")(fluentElement))
    
    iPrint.FluentPath = finalFluentPath
    
    test.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add test
    
    Call pTests.CheckTest(test)
    
End Sub

Private Function IExpressions_ToString(ByVal nv As Variant) As String
    Dim strArgs() As String
    Dim strObj As String
    Dim tempNVTS As String
    Dim i As Long
    Dim elem As Variant
    Dim col As Collection
    Dim counter As Long
    
    i = 0
    counter = 0
    Set col = New Collection
    
    Select Case True
        Case pMeta.Tests.IsDataStructure(nv)
            If pMeta.Tests.Algorithm = flRecursive Then
                tempNVTS = IExpressions_InputToStringRecur(nv)
            ElseIf pMeta.Tests.Algorithm = flIterative Then
                tempNVTS = IExpressions_inputToStringIter(nv)
            End If
        Case IsNull(nv) Or IsEmpty(nv)
            tempNVTS = TypeName(nv)
        Case Not TypeOf nv Is Object
            If TypeName(nv) = "String" Then
                nv = "`" & nv & "`"
            End If
            tempNVTS = CStr(nv)
        Case TypeOf nv Is VBA.ErrObject
            tempNVTS = TypeName(nv)
        Case TypeOf nv Is Object
            If Not nv Is Nothing Then
                tempNVTS = TypeName(nv)
            Else
                tempNVTS = "Nothing"
            End If
    End Select
    
    IExpressions_ToString = tempNVTS
        
End Function

Private Function IExpressions_DatastructureIsEmpty(ByVal v As Variant) As Boolean
    Dim elem As Variant
    Dim counter As Long
    Dim b As Boolean
    Dim temp As Variant

    Select Case True
        Case TypeOf v Is Scripting.Dictionary Or TypeOf v Is Collection
            b = (v.Count = 0)
        
        Case IsArray(v)
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
                    ElseIf IsObject(elem) Then
                        temp = temp & TypeName(elem) & ", "
                    ElseIf TypeName(elem) = "String" Then
                        temp = temp & "`" & elem & "`" & ", "
                    Else
                        temp = temp & elem & ", "
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If pMeta.Tests.IsDataStructure(v(elem)) Then
                        temp = temp & elem & ": " & IExpressions_InputToStringRecur(v(elem)) & ", "
                    ElseIf IsObject(v(elem)) Then
                        temp = temp & elem & ": " & TypeName(v(elem)) & ", "
                    ElseIf TypeName(v(elem)) = "String" Then
                        temp = temp & elem & ": " & "`" & TypeName(v(elem)) & "`" & ", "
                    Else
                        temp = temp & elem & ": " & v(elem) & ", "
                    End If
                Next elem
            End If
        Case Else
            temp = temp & elem & ", "
    End Select
    
    If Right(temp, 2) = ", " Then
        temp = Left(temp, Len(temp) - 2)
    End If

    temp = temp & ")"

    IExpressions_InputToStringRecur = temp
End Function

Function IExpressions_inputToStringIter(v As Variant) As String
    Dim elem As Variant
    Dim tempCol As Collection
    Dim tempCount As Long
    Dim finalCol As Collection
    Dim elemStr As String
    Dim temp As String
    Dim vCounter As String
    Dim dsCount As Long
    
    Set tempCol = New Collection
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
                If IsObject(elem) Then
                    tempCol.Add TypeName(elem)
                ElseIf TypeName(elem) = "String" Then
                    tempCol.Add "`" & elem & "`"
                Else
                    tempCol.Add elem
                End If
            Else
                tempCol.Add elem & ": "
                tempCol.Add v(elem)
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

Function IExpressions_getNestedElementsFromCollection(col As Collection) As Collection
    Dim elem As Variant
    Dim elem2 As Variant
    Dim counter As Long
    Dim elemCount As Long
    Dim elemCounter As Long
    Dim tempCol As Collection
    Dim hasDataStructure As Boolean
    Dim d As Object
    
    Set tempCol = col
    counter = 1
    elemCounter = 1
    
    Do
        hasDataStructure = False

        If IsObject(tempCol(counter)) Then
            Set elem = tempCol(counter)
        Else
            elem = tempCol(counter)
        End If
        
        If pMeta.Tests.IsDataStructure(elem) Then
            hasDataStructure = True
            
            tempCol.Add IExpressions_getDsType(elem), before:=counter
            
            'counter needs to get incremented every time an element is added
            'to tempCol or when the loop continues to the next element
            counter = counter + 1
            
            elemCount = IExpressions_getDsCount(elem)
            
            For Each elem2 In elem
                If Not TypeOf elem Is Scripting.Dictionary Then
                    If TypeName(elem2) = "String" Then
                        tempCol.Add "`" & elem2 & "`", before:=counter
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
                    ElseIf TypeName(d(elem2)) = "String" Then
                        'counter incremented towards end of if statement
                        tempCol.Add elem2 & ": " & "`" & d(elem2) & "`", before:=counter
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

        End If
        
        counter = counter + 1

        If counter >= tempCol.Count And hasDataStructure = False Then Exit Do
    Loop While counter < tempCol.Count
    
    Set IExpressions_getNestedElementsFromCollection = tempCol
End Function

Function IExpressions_getDsCount(dataStructure As Variant) As Long
    Dim counter As Long
    Dim elem As Variant
    
    counter = 0
    
    For Each elem In dataStructure
        counter = counter + 1
    Next elem
    
    IExpressions_getDsCount = counter
End Function

Function IExpressions_getDsType(v As Variant) As String
    Dim counter As Long
    Dim elem As Variant
    Dim temp As String
    
    If IsArray(v) Then
        temp = Left(TypeName(v), Len(TypeName(v)) - 2) & "("
    ElseIf pMeta.Tests.IsDataStructure(v) Then
        temp = TypeName(v) & "("
    ElseIf IsObject(v) Then
        temp = TypeName(v)
    End If
    
    IExpressions_getDsType = temp
End Function


Private Function IExpressions_getResultNumber(ByVal Result As Variant) As Long
    Dim resultNumber As Variant
    
    resultNumber = Null
    
    If IsNull(Result) Or IsEmpty(Result) Then
        resultNumber = 1
    ElseIf TypeName(Result) = "Boolean" Then
        resultNumber = CLng(Result)
    End If
    
    IExpressions_getResultNumber = resultNumber
End Function

Public Function IExpressions_CleanString(TestStr As Variant, CleanDict As Scripting.Dictionary) As String
    Dim elem As Variant
    Dim tempStr As Variant
    
    tempStr = TestStr
    
    For Each elem In CleanDict.Keys
        tempStr = Replace(tempStr, elem, CleanDict(elem))
    Next elem
    
    IExpressions_CleanString = tempStr
End Function

'Private Sub Class_Initialize()
'
'End Sub
