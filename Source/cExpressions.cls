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

Private pTestDictCounter As Scripting.Dictionary
Private pPrinting As cPrinting
Private iPrint As IPrinting
Private pTests As cTests
Private pMeta As cMeta

Private Const NOT_IMPLEMENTED_ERROR As Long = 514

Public Enum flShould
    Contain_
    EndWith_
    EvaluateTo_
    StartWith_
End Enum

Public Enum flHave
    ErrorDescriptionOf_
    ErrorNumberOf_
    ExactSameElementsAs_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
    SameElementsAs_
    SameTypeAs_
    SameUniqueElementsAs_
End Enum

Public Enum flBe
    Alphabetic_
    Alphanumeric_
    Between_
    Erroneous_
    EqualTo_
    GreaterThan_
    GreaterThanOrEqualTo_
    IdenticalTo_
    InDataStructure_
    InDataStructures_
    LessThan_
    LessThanOrEqualTo_
    Numeric_
    OneOf_
    Something_
End Enum

Public Property Set setPrinting(ByVal value As cPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Public Property Set setTests(ByVal value As cTests)
    Set pTests = value
End Property

Public Property Set setMeta(ByVal value As cMeta)
    Set pMeta = value
    Set pMeta.Tests.TestDictCounter = pTestDictCounter
End Property

Public Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal result As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional LowerVal As Variant, _
    Optional HigherVal As Variant)
    
    Dim test As ITest
    Dim shouldBool As String
    
    Set test = New cTest
    
    If IsNull(result) Then
'        result = TypeName(result)
        test.HasNull = True
    End If

    If negateValue Then
        result = Not result
    End If

    With test
        .functionName = functionName
        .result = result
        .testingInput = Empty
        .LowerVal = Empty
        .HigherVal = Empty
        .StrTestInput = ""
        .StrTestValue = ""
    End With
        
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
        .StrTestInput = ToString(.testingInput)
        .StrTestValue = ToString(.testingValue)
    End With
    
    
    Call addFluentPathNew(test, negateValue)

End Sub

Public Function Should_(shouldType As flShould, Optional testingInput As Variant) As String
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

    Should_ = shouldStr
End Function

Public Function Be_(beType As flBe, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
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

    Be_ = beStr
End Function

Public Function Have_(haveType As flHave, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
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

    Have_ = haveStr

End Function

Public Function buildFluentPath(test As ITest, FluentPath As String, tempFluentPath As String) As String
    If IsNull(test.result) Then
        test.result = TypeName(test.result)
    End If

    With test
        If .StrTestInput = "" Then
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .result
        Else
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Testing input: " & .StrTestInput & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .result
        End If
    End With
    
    buildFluentPath = tempFluentPath
End Function

Public Function getFunctionExpr(test As ITest)
    Dim tempFluentPath As String
    
    With test
        Select Case .functionName
            Case "Alphabetic"
                tempFluentPath = Be_(Alphabetic_)
            Case "Alphanumeric"
                tempFluentPath = Be_(Alphanumeric_)
            Case "Between"
                tempFluentPath = Be_(Between_, LowerVal:=.LowerVal, HigherVal:=.HigherVal)
            Case "Contain"
                tempFluentPath = Should_(Contain_, .testingInput)
            Case "Erroneous"
                tempFluentPath = Be_(Erroneous_)
            Case "ErrorNumberOf"
                tempFluentPath = Have_(ErrorNumberOf_, .testingInput)
            Case "ErrorDescriptionOf"
                tempFluentPath = Have_(ErrorDescriptionOf_, .testingInput)
            Case "EqualTo"
                tempFluentPath = Be_(EqualTo_, .testingInput)
            Case "EndWith"
                tempFluentPath = Should_(EndWith_, .testingInput)
            Case "EvaluateTo"
                tempFluentPath = Should_(EvaluateTo_, .StrTestInput)
            Case "ExactSameElementsAs"
                tempFluentPath = Have_(ExactSameElementsAs_, .StrTestInput)
            Case "GreaterThan"
                tempFluentPath = Be_(GreaterThan_, .testingInput)
            Case "GreaterThanOrEqualTo"
                tempFluentPath = Be_(GreaterThanOrEqualTo_, .testingInput)
            Case "IdenticalTo"
                tempFluentPath = Be_(IdenticalTo_, .StrTestInput)
            Case "InDataStructure"
                tempFluentPath = Be_(InDataStructure_, .StrTestInput)
            Case "InDataStructures"
                tempFluentPath = Be_(InDataStructures_, .StrTestInput)
            Case "LengthBetween"
                tempFluentPath = Be_(InDataStructures_, .StrTestInput)
            Case "LengthOf"
                tempFluentPath = Have_(LengthOf_, .testingInput)
            Case "LessThan"
                tempFluentPath = Be_(LessThan_, .testingInput)
            Case "LessThanOrEqualTo"
                tempFluentPath = Be_(LessThanOrEqualTo_, .testingInput)
            Case "MaxLengthOf"
                tempFluentPath = Have_(MaxLengthOf_, .testingInput)
            Case "MinLengthOf"
                tempFluentPath = Have_(MinLengthOf_, .testingInput)
            Case "Numeric"
                tempFluentPath = Be_(Numeric_)
            Case "OneOf"
                tempFluentPath = Be_(OneOf_, .StrTestInput)
            Case "SameElementsAs"
                tempFluentPath = Have_(SameElementsAs_, .StrTestInput)
            Case "SameTypeAs"
                tempFluentPath = Have_(SameTypeAs_, .StrTestInput)
            Case "SameUniqueElementsAs"
                tempFluentPath = Have_(SameUniqueElementsAs_, .StrTestInput)
            Case "Something"
                tempFluentPath = Be_(Something_)
            Case "StartWith"
                tempFluentPath = Should_(StartWith_, .testingInput)
            Case Else
                Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
        End Select
    End With
    
    getFunctionExpr = tempFluentPath
End Function

Public Function getFluentPath(FluentPath As String, test As ITest, Optional notShould As Boolean = False)
    Dim LowerVal As Variant
    Dim HigherVal As Variant
    Dim tempFluentPath As String
    
    tempFluentPath = ""
    
    If test.LowerVal <> Empty And test.HigherVal <> Empty Then
        LowerVal = test.LowerVal
        HigherVal = test.HigherVal
    End If

    tempFluentPath = getFunctionExpr(test)
    
    If notShould Then
        tempFluentPath = "should not " & tempFluentPath
    Else
        tempFluentPath = "should " & tempFluentPath
    End If
    
    tempFluentPath = buildFluentPath(test, FluentPath, tempFluentPath)
    
    getFluentPath = tempFluentPath
End Function

Public Sub addFluentPathNew(ByVal test As ITest, ByVal negateValue As Boolean)
    Dim FluentPath As String
    Dim finalFluentPath As String
    Dim fluentElement As Long
    
    FluentPath = ""
    
    If pTestDictCounter.Exists(test.functionName) Then
        pTestDictCounter(test.functionName) = pTestDictCounter(test.functionName) + 1
    Else
        Err.Raise Number:=NOT_IMPLEMENTED_ERROR, Description:="Test name not implemented in TestDictCounter"
    End If
    
    If pPrinting.Category <> "" Then
        FluentPath = FluentPath & pPrinting.Category & "="
        fluentElement = 2
    Else
        fluentElement = 1
    End If
    
    FluentPath = FluentPath & getResultNumber(test.result) & "="
    
    If pPrinting.TestName <> "" Then FluentPath = FluentPath & pPrinting.TestName & " "
    
    FluentPath = getFluentPath(FluentPath, test, negateValue)
    
    finalFluentPath = Trim(Split(FluentPath, "=")(fluentElement))
    
    iPrint.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add FluentPath
    
    test.FluentPath = finalFluentPath
    
    Call pTests.CheckTest(test)
    
End Sub

Public Function ToString(ByVal nv As Variant) As String
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
                        If IsDataStructure(elem) Then
                            col.Add InputToStringRecur(elem)
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
            If IsDataStructure(nv) Then
                tempNVTS = InputToStringRecur(nv)
            ElseIf Not nv Is Nothing Then
                tempNVTS = TypeName(nv)
            End If
    End Select
    
    ToString = tempNVTS
        
End Function

Public Function DatastructureIsEmpty(v As Variant) As Boolean
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
    
    DatastructureIsEmpty = b
End Function

Public Function InputToStringRecur(v As Variant) As String
    Dim temp As String
    Dim elem As Variant
    Dim emptyDS As Boolean

    temp = ""

    Select Case True
        Case IsArray(v) And DatastructureIsEmpty(v)
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
                        temp = temp & InputToStringRecur(elem) & ", "
                    Else
                        temp = temp & elem & ", "
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If IsObject(v(elem)) Or IsArray(v(elem)) Then
                        temp = temp & elem & ": " & InputToStringRecur(v(elem)) & ", "
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

    InputToStringRecur = temp
End Function

Public Function getResultNumber(result As Variant)
    Dim resultNumber As Variant
    
    resultNumber = Null
    
    If IsNull(result) Then
        resultNumber = 1
    ElseIf TypeName(result) = "Boolean" Then
        resultNumber = CLng(result)
    End If
    
    getResultNumber = resultNumber
End Function

Public Function IsDataStructure(ByVal DataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    dsTypeName = TypeName(DataStructure)
    tempBool = False

    On Error Resume Next
        For Each elem In DataStructure
            Exit For
        Next elem
        
        tempErr = Err.Number
        
        tempBool = (tempErr = 0)
        
    On Error GoTo 0
    
    IsDataStructure = tempBool
End Function

Public Function getTestDictCounter() As Scripting.Dictionary
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

    Set getTestDictCounter = TestDictCounter
End Function

Private Sub Class_Initialize()
    Set pTestDictCounter = getTestDictCounter
End Sub
