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

Private pMeta As cMeta
Private pTests As cTests
Private pPrinting As cPrinting
Private iPrint As IPrinting
Private pExprDict As Scripting.Dictionary

Private pTestValue As Variant
Private pFluentPath As String

Private Const NOT_IMPLEMENTED_ERROR As Long = 514

Public Enum flSearch
    flRecursive
    flIterative
End Enum

Enum flShould
    Contain_
    EndWith_
    EvaluateTo_
    StartWith_
End Enum

Enum flHave
    ErrorDescriptionOf_
    ErrorNumberOf_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
End Enum

Enum flBe
    Alphabetic_
    Alphanumeric_
    Between_
    Erroneous_
    EqualTo_
    GreaterThan_
    GreaterThanOrEqualTo_
    InDataStructure_
    InDataStructures_
    LessThan_
    LessThanOrEqualTo_
    Numeric_
    OneOf_
    Something_
End Enum

Public Property Let TestValue(ByVal value As Variant)
    pTestValue = value
End Property

Public Property Set TestValue(ByVal value As Variant)
    Set pTestValue = value
End Property

Public Property Get TestValue() As Variant
    If IsObject(pTestValue) Then
        Set TestValue = pTestValue
    Else
        TestValue = pTestValue
    End If
End Property

Public Property Set setMeta(ByVal value As cMeta)
    Set pMeta = value
End Property

Public Property Set setTests(ByVal value As cTests)
    Set pTests = value
End Property

Public Property Set setPrinting(ByVal value As IPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Public Function InDataStructures(ByVal TestingValue As Variant, ByVal method As flSearch, negateValue As Boolean, ParamArray args() As Variant) As Boolean
    Dim i As Long
    Dim temp_bool As Boolean
    Dim IsLastElem As Boolean
    Dim col As VBA.Collection
    Dim strArgs() As String
    Dim strTestingVal As String
    
    ReDim strArgs(UBound(args))
    
    Set col = getNestedElementsRecur(args)
    
    temp_bool = False
    
    temp_bool = Me.InDataStructure(TestingValue, col, method, negateValue, updateFluentPath:=False)
    
    InDataStructures = temp_bool
    
    strTestingVal = ToString(args)
    
    Call addToTestAndFluentPath("InDataStructures", CBool(InDataStructures), CBool(negateValue), TestingValue, args)
    
End Function

Public Function InDataStructure( _
ByVal TestingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flSearch = flRecursive, _
Optional negateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    Dim strTestingVal As String
    
    tempBool = False
    
    If isDataStructure(DataStructure) Then
        If method = flRecursive Then
            Set col = getNestedElementsRecur(DataStructure)
        ElseIf method = flIterative Then
            Set col = getNestedElementsIter(DataStructure)
        End If
            
        For Each elem In col
            If TestingValue = elem Then
                tempBool = True
                Exit For
            End If
        Next elem
    End If
    
    InDataStructure = tempBool
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("InDataStructure", CBool(InDataStructure), CBool(negateValue), TestingValue, DataStructure)
    End If
    
End Function

Private Function isDataStructure(ByVal DataStructure As Variant)
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
    
    isDataStructure = tempBool
End Function

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If isDataStructure(DataStructure) Then
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
            If isDataStructure(elem) Then
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

Private Function getNestedElementsIter(ByVal DataStructure As Variant) As Collection
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
            If isDataStructure(col(i)) Then
                hasDataStructures = True
                For Each elem In col(i)
                    col.Add elem
                Next elem
                col.Remove i
                Exit For
            End If
        Next i
    Loop While hasDataStructures = True

    Set getNestedElementsIter = col
End Function

Public Function GreaterThan(ByVal TestingValue As Double, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThan = (TestingValue > testingInput)
    
    Call addToTestAndFluentPath("GreaterThan", CBool(GreaterThan), CBool(negateValue), TestingValue, testingInput)
    
End Function

Public Function GreaterThanOrEqualTo( _
ByVal TestingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (TestingValue >= testingInput)
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(negateValue), TestingValue, testingInput)
    End If
        
End Function

Public Function LessThan(ByVal TestingValue As Double, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (TestingValue < testingInput)
    
    Call addToTestAndFluentPath("LessThan", CBool(LessThan), CBool(negateValue), TestingValue, testingInput)

End Function

Public Function LessThanOrEqualTo( _
ByVal TestingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (TestingValue <= testingInput)
       
    If updateFluentPath Then
        Call addToTestAndFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(negateValue), TestingValue, testingInput)
    End If
        
End Function

Public Function EqualTo(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = (TestingValue = testingInput)
    
    If pMeta.ApproximateEqual Then
        If TypeName(TestingValue) <> TypeName(testingInput) Then
            If IsNumeric(TestingValue) And IsNumeric(testingInput) Then
                tempBool = (Abs(TestingValue - testingInput) <= pMeta.Epsilon)
            ElseIf VarType(TestingValue) = vbString Or VarType(testingInput) = vbString Then
                tempBool = (StrComp(CStr(TestingValue), CStr(testingInput), vbTextCompare) = 0)
            End If
        End If
    End If
    
    EqualTo = tempBool
    
    Call addToTestAndFluentPath("EqualTo", CBool(EqualTo), CBool(negateValue), TestingValue, testingInput)
        
End Function

Public Function Contain(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If TestingValue Like "*" & testingInput & "*" Then
        Contain = True
    End If
    
    Call addToTestAndFluentPath("Contain", CBool(Contain), CBool(negateValue), TestingValue, testingInput)
        
End Function

Public Function StartWith(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(testingInput)
    
    If Left(TestingValue, valLength) = CStr(testingInput) Then
        StartWith = True
    End If
    
    Call addToTestAndFluentPath("StartWith", CBool(StartWith), CBool(negateValue), TestingValue, testingInput)
        
End Function

Public Function EndWith(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(testingInput)
    
    If Right(TestingValue, valLength) = CStr(testingInput) Then
        EndWith = True
    End If
    
    Call addToTestAndFluentPath("EndWith", CBool(EndWith), CBool(negateValue), TestingValue, testingInput)
    
End Function

Public Function LengthOf(ByVal TestingValue As Double, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(TestingValue)) = testingInput)
    
    Call addToTestAndFluentPath("LengthOf", CBool(LengthOf), CBool(negateValue), TestingValue, testingInput)
            
End Function

Public Function MaxLengthOf( _
ByVal TestingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(TestingValue)) <= testingInput)
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(negateValue), TestingValue, testingInput)
    End If
            
End Function

Public Function MinLengthOf( _
ByVal TestingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(TestingValue)) >= testingInput)
        
    If updateFluentPath Then
        Call addToTestAndFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(negateValue), TestingValue, testingInput)
    End If
        
End Function

Public Function Something(ByVal TestingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not TestingValue Is Nothing)
    
    Call addToTestAndFluentPath("Something", CBool(Something), CBool(negateValue), TestingValue)
        
End Function

Public Function Between(ByVal TestingValue As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(TestingValue, LowerVal, updateFluentPath:=False) And LessThanOrEqualTo(TestingValue, HigherVal, updateFluentPath:=False))
    
    Call addToTestAndFluentPath("Between", CBool(Between), CBool(negateValue), TestingValue:=TestingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
        
End Function

Public Function OneOf(ByVal TestingValue As Variant, ByVal negateValue As Boolean, ParamArray args() As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    Dim TestingInputStr As String
        
    arr = args(0)
    tempBool = False
    
    For i = LBound(arr) To UBound(arr)
'        If IsObject(arr(i)) Then
'            TestingInputStr = TestingInputStr & TypeName(arr(i)) & ", "
'        Else
'            TestingInputStr = TestingInputStr & arr(i) & ", "
'        End If

        If IsObject(arr(i)) Or IsObject(TestingValue) Then
            If TypeName(arr(i)) = TypeName(TestingValue) Then
                tempBool = True
            End If
        ElseIf arr(i) = TestingValue Then
            tempBool = True
            'Exit For
        End If
    Next i
    
'    TestingInputStr = Left(TestingInputStr, Len(TestingInputStr) - 2)
    
    OneOf = tempBool
    
    If TypeOf TestingValue Is Object  Then TestingValue = TypeName(TestingValue)
    
    Call addToTestAndFluentPath("OneOf", CBool(OneOf), CBool(negateValue), TestingValue, arr)
        
End Function

Public Function LengthBetween(ByVal TestingValue As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(TestingValue))
    
    LengthBetween = (MinLengthOf(tempOrigVal, LowerVal, updateFluentPath:=False) And MaxLengthOf(tempOrigVal, HigherVal, updateFluentPath:=False))
    
    Call addToTestAndFluentPath("LengthBetween", CBool(LengthBetween), CBool(negateValue), TestingValue:=TestingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue = False) As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(TestingValue) = testingInput)

    Call addToTestAndFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(negateValue), TestingValue, testingInput)
    
End Function

Public Function Alphabetic(ByVal TestingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    
    For i = 1 To Len(TestingValue)
        tempChar = Mid(TestingValue, i, 1)
        If tempChar Like "[A-Za-z]" Then
            tempAlph = tempAlph & tempChar
        End If
    Next i
    
    Alphabetic = (TestingValue = tempAlph)
    
    Call addToTestAndFluentPath("Alphabetic", CBool(Alphabetic), CBool(negateValue), TestingValue)
End Function

Public Function Numeric(ByVal TestingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    
    For i = 1 To Len(TestingValue)
        tempChar = Mid(TestingValue, i, 1)
        If tempChar Like "[0-9]" Then
            tempNum = tempNum & tempChar
        End If
    Next i
    
    Numeric = (TestingValue = tempNum)
     
    Call addToTestAndFluentPath("Numeric", CBool(Numeric), CBool(negateValue), TestingValue)
End Function

Public Function Alphanumeric(ByVal TestingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    
    b = False
    
    For i = 1 To Len(TestingValue)
        tempChar = Mid(TestingValue, i, 1)
        If tempChar Like "[A-Za-z]" Then
            tempAlph = tempAlph & tempChar
        ElseIf tempChar Like "[0-9]" Then
            tempNum = tempNum & tempChar
        End If
    Next i
    
    If Len(tempAlph) > 1 And Len(tempNum) > 1 Then
        If Len(tempAlph) + Len(tempNum) = Len(TestingValue) Then
            b = True
        End If
    End If
    
    Alphanumeric = b
    
    Call addToTestAndFluentPath("Alphanumeric", CBool(Alphanumeric), CBool(negateValue), TestingValue)
End Function

Public Function Erroneous(ByVal TestingValue As Variant, Optional ByVal negateValue = False) As Boolean
    Dim tempTestingVal As Variant
    Dim tempErrType As Variant
    
    tempTestingVal = False
    
    If TypeOf TestingValue Is ErrObject Then
        tempTestingVal = (TestingValue.Number <> 0)
    Else
        On Error Resume Next
            tempTestingVal = Application.Evaluate(TestingValue)
            tempTestingVal = (TypeName(tempTestingVal) = "Error")
        On Error GoTo 0
    End If
    
    Erroneous = tempTestingVal
    
    If TypeOf TestingValue Is Object  Then
        tempErrType = TypeName(TestingValue)
    Else
        tempErrType = TestingValue
    End If
    
    Call addToTestAndFluentPath("Erroneous", CBool(Erroneous), CBool(negateValue), tempErrType)
End Function

Public Function ErrorNumberOf(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue = False) As Boolean
    Dim temp As Variant
    Dim e1 As Variant
    
    On Error Resume Next
        temp = Application.Evaluate(TestingValue)
        Err.Raise (testingInput)
        temp = Err.Number
        ErrorNumberOf = (temp = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorNumberOf", CBool(ErrorNumberOf), CBool(negateValue), TestingValue, testingInput)
End Function

Public Function ErrorDescriptionOf(ByVal TestingValue As Variant, ByVal testingInput As Variant, Optional ByVal negateValue = False) As Boolean
    Dim tempTestVal As Variant
    
    On Error Resume Next
        tempTestVal = CStr(Application.Evaluate(TestingValue))
        tempTestVal = Replace(tempTestVal, "Error ", "")
        Err.Raise tempTestVal
        tempTestVal = Err.Description
        ErrorDescriptionOf = (tempTestVal = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorDescriptionOf", CBool(ErrorDescriptionOf), CBool(negateValue), tempTestVal, testingInput)
End Function

Private Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal Result As Boolean, _
    ByVal negateValue As Boolean, _
    ByVal TestingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional LowerVal As Variant, _
    Optional HigherVal As Variant)
    
    Dim test As cTest
    Dim shouldBool As String
    
    Set test = New cTest

    If negateValue Then
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
    End With
        
    If TypeName(TestingValue) <> "Nothing" Then
        If TypeOf TestingValue Is Object  Then
            Set test.TestingValue = TestingValue
        Else
            test.TestingValue = TestingValue
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
        .StrTestValue = ToString(.TestingValue)
    End With
    
    
    Call addFluentPathNew(test, negateValue)

End Sub

Function Should_(shouldType As flShould, Optional testingInput As Variant) As String
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

Function Be_(beType As flBe, Optional tesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
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
    
    If tesingInput <> "" And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        beStr = beStr & " " & tesingInput
    End If
    
    If tesingInput = "" And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        beStr = beStr & " " & LowerVal & " and " & HigherVal
    End If

    Be_ = beStr
End Function

Function Have_(haveType As flHave, Optional tesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
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
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    haveStr = haveStr & funcStr
    
    If Not IsMissing(tesingInput) And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        haveStr = haveStr & " " & tesingInput
    End If
    
    If IsMissing(tesingInput) And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        haveStr = haveStr & " " & LowerVal & " and " & HigherVal
    End If

    Have_ = haveStr

End Function

Public Function buildFluentPath(test As cTest, FluentPath As String, tempFluentPath As String) As String
    With test
        If .StrTestInput = "" Then
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .Result
        Else
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Testing input: " & .StrTestInput & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .Result
        End If
    End With
    
    buildFluentPath = tempFluentPath
End Function

Public Function getFunctionExpr(test As cTest)
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
            Case "GreaterThan"
                tempFluentPath = Be_(GreaterThan_, .testingInput)
            Case "GreaterThanOrEqualTo"
                tempFluentPath = Be_(GreaterThanOrEqualTo_, .testingInput)
            Case "InDataStructure"
                tempFluentPath = Be_(InDataStructure_, .StrTestInput)
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
            Case "Something"
                tempFluentPath = Be_(Something_)
            Case "StartWith"
                tempFluentPath = Should_(StartWith_, .testingInput)
        End Select
    End With
    
    getFunctionExpr = tempFluentPath
End Function

Function getFluentPath(FluentPath As String, test As cTest, Optional notShould As Boolean = False)
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

Private Sub addFluentPathNew(ByVal test As cTest, ByVal negateValue As Boolean)
    Dim FluentPath As String
    Dim finalFluentPath As String
    Dim fluentElement As Long
    
    FluentPath = ""
    
    If pPrinting.Category <> "" Then
        FluentPath = FluentPath & pPrinting.Category & "="
        fluentElement = 2
    Else
        fluentElement = 1
    End If
    
    FluentPath = FluentPath & CLng(test.Result) & "="
    
    If pPrinting.TestName <> "" Then FluentPath = FluentPath & pPrinting.TestName & " "
    
    FluentPath = getFluentPath(FluentPath, test, negateValue)
    
    finalFluentPath = Split(FluentPath, "=")(fluentElement)
    
    iPrint.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add FluentPath
    
    test.FluentPath = finalFluentPath
    
    Call pTests.CheckTest(test)
    
End Sub

Private Function ToString(ByVal nv As Variant) As String
    Dim strArgs() As String
    Dim strObj As String
    Dim tempNVTS As String
    Dim i As Long
    Dim elem As Variant
    Dim col As Collection
    
    i = 0
    Set col = New Collection
    
    Select Case True
        Case IsArray(nv)
            For Each elem In nv
                    If TypeName(elem) = "Nothing" Then
                        col.Add TypeName(elem)
                    ElseIf TypeOf elem Is Object  Then
                        If Not elem Is Nothing Then
                            col.Add TypeName(elem)
                        End If
                    ElseIf IsArray(elem) Then
                        col.Add TypeName(elem)
                    Else
                        col.Add elem
                    End If
                    
                    ReDim strArgs(1 To col.Count)
                    
                    For i = 1 To col.Count
                        strArgs(i) = col(i)
                    Next i
            Next elem
            
            tempNVTS = Join(strArgs, ", ")
        Case Not TypeOf nv Is Object
            tempNVTS = CStr(nv)
        Case TypeOf nv Is ErrObject
            tempNVTS = TypeName(nv)
        Case TypeOf nv Is Object
            tempNVTS = TypeName(nv)
    End Select
    
    ToString = tempNVTS
        
End Function

Private Sub Class_Initialize()
    Set pExprDict = New Scripting.Dictionary
End Sub
