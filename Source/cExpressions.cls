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

Private pTestDictCounter As Scripting.Dictionary
Private pPrinting As IPrinting
Private iPrint As IPrintExpr
Private pTests As cTests
Private pMeta As IMeta

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
    Set pMeta.Tests.TestDictCounter = pTestDictCounter
End Property

Private Sub IExpressions_addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal Result As Variant, _
    ByVal NegateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal LowerVal As Variant, _
    Optional ByVal HigherVal As Variant)
    
    Dim test As ITest
    Dim testDev As ITestDev
    Dim shouldBool As String
    
    Set test = New cTest
    Set testDev = test
    
    If IsNull(Result) Then
'        result = TypeName(result)
        test.HasNull = True
    End If

    If NegateValue Then
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
        .Category = pPrinting.Category
        .TestName = pPrinting.Name
    End With
    
    If pPrinting.Category <> "" Then iPrint.CategorySet = True
    If pPrinting.Name <> "" Then iPrint.NameSet = True
    
    testDev.NegateValue = NegateValue
        
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
    
    If Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        test.testingInput = Array(CStr(LowerVal), CStr(HigherVal))
    End If
    
    With test
        .StrTestInput = IExpressions_ToString(.testingInput)
        .StrTestValue = IExpressions_ToString(.testingValue)
    End With
    
    
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
Optional ByVal HigherVal As Variant)
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
Optional ByVal HigherVal As Variant)
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
    
    If Not IsMissing(TesingInput) And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        haveStr = haveStr & " " & TesingInput
    End If
    
    If IsMissing(TesingInput) And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        haveStr = haveStr & " " & LowerVal & " and " & HigherVal
    End If

    IExpressions_Have = haveStr

End Function

Private Function IExpressions_buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
    If IsNull(test.Result) Then
        test.Result = TypeName(test.Result)
    End If

    With test
        If .StrTestInput = "" Then
            tempFluentPath = FluentPath & _
            "Testing value: " & .StrTestValue & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .Result
        Else
            tempFluentPath = FluentPath & _
            "Testing value: " & .StrTestValue & ";" & _
            " Testing input: " & .StrTestInput & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .Result
        End If
    End With
    
    IExpressions_buildFluentPath = tempFluentPath
End Function

Private Function IExpressions_getFunctionExpr(ByVal test As ITest)
    Dim tempFluentPath As String
    
    With test
        Select Case .functionName
            Case "Alphabetic"
                tempFluentPath = IExpressions_Be(Alphabetic_)
            Case "Alphanumeric"
                tempFluentPath = IExpressions_Be(Alphanumeric_)
            Case "Between"
                tempFluentPath = IExpressions_Be(Between_, LowerVal:=.LowerVal, HigherVal:=.HigherVal)
            Case "Contain"
                tempFluentPath = IExpressions_Should(Contain_, .testingInput)
            Case "Erroneous"
                tempFluentPath = IExpressions_Be(Erroneous_)
            Case "ErrorNumberOf"
                tempFluentPath = IExpressions_Have(ErrorNumberOf_, .testingInput)
            Case "ErrorDescriptionOf"
                tempFluentPath = IExpressions_Have(ErrorDescriptionOf_, .testingInput)
            Case "EqualTo"
                tempFluentPath = IExpressions_Be(EqualTo_, .testingInput)
            Case "EndWith"
                tempFluentPath = IExpressions_Should(EndWith_, .testingInput)
            Case "EvaluateTo"
                tempFluentPath = IExpressions_Should(EvaluateTo_, .StrTestInput)
            Case "ExactSameElementsAs"
                tempFluentPath = IExpressions_Have(ExactSameElementsAs_, .StrTestInput)
            Case "GreaterThan"
                tempFluentPath = IExpressions_Be(GreaterThan_, .testingInput)
            Case "GreaterThanOrEqualTo"
                tempFluentPath = IExpressions_Be(GreaterThanOrEqualTo_, .testingInput)
            Case "IdenticalTo"
                tempFluentPath = IExpressions_Be(IdenticalTo_, .StrTestInput)
            Case "InDataStructure"
                tempFluentPath = IExpressions_Be(InDataStructure_, .StrTestInput)
            Case "InDataStructures"
                tempFluentPath = IExpressions_Be(InDataStructures_, .StrTestInput)
            Case "LengthBetween"
                tempFluentPath = IExpressions_Have(LengthBetween_, .StrTestInput)
            Case "LengthOf"
                tempFluentPath = IExpressions_Have(LengthOf_, .testingInput)
            Case "LessThan"
                tempFluentPath = IExpressions_Be(LessThan_, .testingInput)
            Case "LessThanOrEqualTo"
                tempFluentPath = IExpressions_Be(LessThanOrEqualTo_, .testingInput)
            Case "MaxLengthOf"
                tempFluentPath = IExpressions_Have(MaxLengthOf_, .testingInput)
            Case "MinLengthOf"
                tempFluentPath = IExpressions_Have(MinLengthOf_, .testingInput)
            Case "Numeric"
                tempFluentPath = IExpressions_Be(Numeric_)
            Case "OneOf"
                tempFluentPath = IExpressions_Be(OneOf_, .StrTestInput)
            Case "SameElementsAs"
                tempFluentPath = IExpressions_Have(SameElementsAs_, .StrTestInput)
            Case "SameTypeAs"
                tempFluentPath = IExpressions_Have(SameTypeAs_, .StrTestInput)
            Case "SameUniqueElementsAs"
                tempFluentPath = IExpressions_Have(SameUniqueElementsAs_, .StrTestInput)
            Case "Something"
                tempFluentPath = IExpressions_Be(Something_)
            Case "StartWith"
                tempFluentPath = IExpressions_Should(StartWith_, .testingInput)
            Case Else
                Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
        End Select
    End With
    
    IExpressions_getFunctionExpr = tempFluentPath
End Function

Private Function IExpressions_getFluentPath(ByVal FluentPath As String, ByVal test As ITest, Optional ByVal notShould As Boolean = False)
    Dim LowerVal As Variant
    Dim HigherVal As Variant
    Dim tempFluentPath As String
    
    tempFluentPath = ""
    
    If test.LowerVal <> Empty And test.HigherVal <> Empty Then
        LowerVal = test.LowerVal
        HigherVal = test.HigherVal
    End If

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
        pTestDictCounter(test.functionName) = pTestDictCounter(test.functionName) + 1
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
        Case IsArray(nv)
            For Each elem In nv
                    If TypeName(elem) = "Nothing" Or TypeName(elem) = "Null" Then
                        col.Add TypeName(elem)
                    ElseIf TypeOf elem Is Object  Or IsArray(elem) Then
                        If IExpressions_IsDataStructure(elem) Then
                            col.Add IExpressions_InputToStringRecur(elem)
                        ElseIf Not elem Is Nothing Then
                            col.Add TypeName(elem)
                        End If
                    Else
                        col.Add elem
                    End If
                counter = counter + 1
            Next elem
            
            
            If counter > 0 Then
                ReDim strArgs(1 To col.Count)
                    
                For i = 1 To col.Count
                    strArgs(i) = col(i)
                Next i
                
                tempNVTS = Join(strArgs, ", ")
            End If
        Case IsNull(nv)
            tempNVTS = TypeName(nv)
        Case Not TypeOf nv Is Object
            tempNVTS = CStr(nv)
        Case TypeOf nv Is ErrObject
            tempNVTS = TypeName(nv)
        Case TypeOf nv Is Object
            If IExpressions_IsDataStructure(nv) Then
                tempNVTS = IExpressions_InputToStringRecur(nv)
            ElseIf Not nv Is Nothing Then
                tempNVTS = TypeName(nv)
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
        Case TypeOf v Is Dictionary Or TypeOf v Is Collection
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

    temp = ""

    Select Case True
        Case IsArray(v) And IExpressions_DatastructureIsEmpty(v)
            temp = Left(TypeName(v), Len(TypeName(v)) - 2) & "("
        Case IsObject(v) Or IsArray(v)
            If IsArray(v) Then
                temp = Left(TypeName(v), Len(TypeName(v)) - 2) & "("
            Else
                temp = TypeName(v) & "("
            End If
            
            If Not TypeOf v Is Dictionary Then
                For Each elem In v
                    If IsObject(elem) Or IsArray(elem) Then
                        temp = temp & IExpressions_InputToStringRecur(elem) & ", "
                    Else
                        temp = temp & elem & ", "
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If IsObject(v(elem)) Or IsArray(v(elem)) Then
                        temp = temp & elem & ": " & IExpressions_InputToStringRecur(v(elem)) & ", "
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

Private Function IExpressions_getResultNumber(ByVal Result As Variant)
    Dim resultNumber As Variant
    
    resultNumber = Null
    
    If IsNull(Result) Then
        resultNumber = 1
    ElseIf TypeName(Result) = "Boolean" Then
        resultNumber = CLng(Result)
    End If
    
    IExpressions_getResultNumber = resultNumber
End Function

Private Function IExpressions_datastructureInCollection(ByVal dsName As String, ByVal col As Collection)
    Dim elem As Variant
    Dim b As Boolean
    
    b = False
    
    For Each elem In col
        If elem = dsName Then
            b = True
            Exit For
        End If
    Next elem
    
    IExpressions_datastructureInCollection = b
End Function

Private Function IExpressions_IsDataStructure(ByVal DataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    dsTypeName = TypeName(DataStructure)
    tempBool = False
    
    If IsArray(DataStructure) Then
        tempBool = True
    ElseIf IExpressions_datastructureInCollection(dsTypeName, pMeta.DataStructures) Then
        On Error Resume Next
            For Each elem In DataStructure
                Exit For
            Next elem
            
            tempErr = Err.Number
            
            tempBool = (tempErr = 0)
            
        On Error GoTo 0
    End If
    
    IExpressions_IsDataStructure = tempBool
End Function

Private Function IExpressions_getTestDictCounter() As Scripting.Dictionary
    Dim TestDictCounter As Scripting.Dictionary

    Set TestDictCounter = New Scripting.Dictionary

    TestDictCounter.Add "Alphabetic", 0
    TestDictCounter.Add "Alphanumeric", 0
    TestDictCounter.Add "Between", 0
    TestDictCounter.Add "Contain", 0
    TestDictCounter.Add "EndWith", 0
    TestDictCounter.Add "EqualTo", 0
    TestDictCounter.Add "Erroneous", 0
    TestDictCounter.Add "ErrorDescriptionOf", 0
    TestDictCounter.Add "ErrorNumberOf", 0
    TestDictCounter.Add "EvaluateTo", 0
    TestDictCounter.Add "ExactSameElementsAs", 0
    TestDictCounter.Add "GreaterThan", 0
    TestDictCounter.Add "GreaterThanOrEqualTo", 0
    TestDictCounter.Add "IdenticalTo", 0
    TestDictCounter.Add "InDataStructure", 0
    TestDictCounter.Add "InDataStructures", 0
    TestDictCounter.Add "LengthBetween", 0
    TestDictCounter.Add "LengthOf", 0
    TestDictCounter.Add "LessThan", 0
    TestDictCounter.Add "LessThanOrEqualTo", 0
    TestDictCounter.Add "MaxLengthOf", 0
    TestDictCounter.Add "MinLengthOf", 0
    TestDictCounter.Add "Numeric", 0
    TestDictCounter.Add "OneOf", 0
    TestDictCounter.Add "SameElementsAs", 0
    TestDictCounter.Add "SameTypeAs", 0
    TestDictCounter.Add "SameUniqueElementsAs", 0
    TestDictCounter.Add "Something", 0
    TestDictCounter.Add "StartWith", 0

    Set IExpressions_getTestDictCounter = TestDictCounter
End Function

Private Sub Class_Initialize()
    Set pTestDictCounter = IExpressions_getTestDictCounter
End Sub
