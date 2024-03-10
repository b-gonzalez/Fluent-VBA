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

Public Enum flAlgorithm
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
    SameTypeAs_
    SameElementsAs_
End Enum

Enum flBe
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

Public Function InDataStructures( _
ByVal testingValue As Variant, _
ByVal method As flAlgorithm, _
ByVal negateValue As Boolean, _
ParamArray args() As Variant) As Boolean
    Dim i As Long
    Dim temp_bool As Boolean
    Dim IsLastElem As Boolean
    Dim col As VBA.Collection
    Dim strArgs() As String
    Dim strTestingVal As String
    
    ReDim strArgs(UBound(args))
    
    Set col = getNestedElementsRecur(args)
    
    temp_bool = False
    
    temp_bool = Me.InDataStructure(testingValue, col, method, negateValue, updateFluentPath:=False)
    
    InDataStructures = temp_bool
    
    strTestingVal = ToString(args)
    
    Call addToTestAndFluentPath("InDataStructures", CBool(InDataStructures), CBool(negateValue), testingValue, args)
    
End Function

Public Function InDataStructure( _
ByVal testingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flAlgorithm = flRecursive, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    Dim strTestingVal As String
    
    tempBool = False
    
    If IsDataStructure(DataStructure) Then
        If method = flRecursive Then
            Set col = getNestedElementsRecur(DataStructure)
        ElseIf method = flIterative Then
            Set col = getNestedElementsIter(DataStructure)
        End If
            
        For Each elem In col
            If testingValue = elem Then
                tempBool = True
                Exit For
            End If
        Next elem
    End If
    
    InDataStructure = tempBool
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("InDataStructure", CBool(InDataStructure), CBool(negateValue), testingValue, DataStructure)
    End If
    
End Function

Private Function IsDataStructure(ByVal DataStructure As Variant)
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

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If IsDataStructure(DataStructure) Then
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
            If IsDataStructure(elem) Then
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
            If IsDataStructure(col(i)) Then
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

Public Function GreaterThan(ByVal testingValue As Double, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThan = (testingValue > testingInput)
    
    Call addToTestAndFluentPath("GreaterThan", CBool(GreaterThan), CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function GreaterThanOrEqualTo( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (testingValue >= testingInput)
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function LessThan( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (testingValue < testingInput)
    
    Call addToTestAndFluentPath("LessThan", CBool(LessThan), CBool(negateValue), testingValue, testingInput)

End Function

Public Function LessThanOrEqualTo( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (testingValue <= testingInput)
       
    If updateFluentPath Then
        Call addToTestAndFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean

    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
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
    
    EqualTo = tempBool
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("EqualTo", CBool(EqualTo), CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If testingValue Like "*" & testingInput & "*" Then
        Contain = True
    End If
    
    Call addToTestAndFluentPath("Contain", CBool(Contain), CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(testingInput)
    
    If Left(testingValue, valLength) = CStr(testingInput) Then
        StartWith = True
    End If
    
    Call addToTestAndFluentPath("StartWith", CBool(StartWith), CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(testingInput)
    
    If Right(testingValue, valLength) = CStr(testingInput) Then
        EndWith = True
    End If
    
    Call addToTestAndFluentPath("EndWith", CBool(EndWith), CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function LengthOf( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(testingValue)) = testingInput)
    
    Call addToTestAndFluentPath("LengthOf", CBool(LengthOf), CBool(negateValue), testingValue, testingInput)
            
End Function

Public Function MaxLengthOf( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(testingValue)) <= testingInput)
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(negateValue), testingValue, testingInput)
    End If
            
End Function

Public Function MinLengthOf( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(testingValue)) >= testingInput)
        
    If updateFluentPath Then
        Call addToTestAndFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not testingValue Is Nothing)
    
    Call addToTestAndFluentPath("Something", CBool(Something), CBool(negateValue), testingValue)
        
End Function

Public Function Between( _
ByVal testingValue As Double, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(testingValue, LowerVal, updateFluentPath:=False) And LessThanOrEqualTo(testingValue, HigherVal, updateFluentPath:=False))
    
    Call addToTestAndFluentPath("Between", CBool(Between), CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
        
End Function

Public Function OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ParamArray args() As Variant) As Boolean
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

        If IsObject(arr(i)) Or IsObject(testingValue) Then
            If TypeName(arr(i)) = TypeName(testingValue) Then
                tempBool = True
            End If
        ElseIf arr(i) = testingValue Then
            tempBool = True
            'Exit For
        End If
    Next i
    
'    TestingInputStr = Left(TestingInputStr, Len(TestingInputStr) - 2)
    
    OneOf = tempBool
    
    If TypeOf testingValue Is Object  Then testingValue = TypeName(testingValue)
    
    Call addToTestAndFluentPath("OneOf", CBool(OneOf), CBool(negateValue), testingValue, arr)
        
End Function

Public Function LengthBetween( _
ByVal testingValue As Double, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(testingValue))
    
    LengthBetween = (MinLengthOf(tempOrigVal, LowerVal, updateFluentPath:=False) And MaxLengthOf(tempOrigVal, HigherVal, updateFluentPath:=False))
    
    Call addToTestAndFluentPath("LengthBetween", CBool(LengthBetween), CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False) As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(testingValue) = testingInput)

    Call addToTestAndFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    
    For i = 1 To Len(testingValue)
        tempChar = Mid(testingValue, i, 1)
        If tempChar Like "[A-Za-z]" Then
            tempAlph = tempAlph & tempChar
        End If
    Next i
    
    Alphabetic = (testingValue = tempAlph)
    
    Call addToTestAndFluentPath("Alphabetic", CBool(Alphabetic), CBool(negateValue), testingValue)
End Function

Public Function Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    
    For i = 1 To Len(testingValue)
        tempChar = Mid(testingValue, i, 1)
        If tempChar Like "[0-9]" Then
            tempNum = tempNum & tempChar
        End If
    Next i
    
    Numeric = (testingValue = tempNum)
     
    Call addToTestAndFluentPath("Numeric", CBool(Numeric), CBool(negateValue), testingValue)
End Function

Public Function Alphanumeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Boolean
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    
    b = False
    
    For i = 1 To Len(testingValue)
        tempChar = Mid(testingValue, i, 1)
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
    
    Alphanumeric = b
    
    Call addToTestAndFluentPath("Alphanumeric", CBool(Alphanumeric), CBool(negateValue), testingValue)
End Function

Public Function Erroneous(ByVal testingValue As Variant, Optional ByVal negateValue = False) As Boolean
    Dim tempTestingVal As Variant
    Dim tempErrType As Variant
    
    tempTestingVal = False
    
    If TypeOf testingValue Is ErrObject Then
        tempTestingVal = (testingValue.Number <> 0)
    Else
        On Error Resume Next
            tempTestingVal = Application.Evaluate(testingValue)
            tempTestingVal = (TypeName(tempTestingVal) = "Error")
        On Error GoTo 0
    End If
    
    Erroneous = tempTestingVal
    
    If TypeOf testingValue Is Object  Then
        tempErrType = TypeName(testingValue)
    Else
        tempErrType = testingValue
    End If
    
    Call addToTestAndFluentPath("Erroneous", CBool(Erroneous), CBool(negateValue), tempErrType)
End Function

Public Function ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False) As Boolean
    Dim temp As Variant
    Dim e1 As Variant
    
    On Error Resume Next
        temp = Application.Evaluate(testingValue)
        Err.Raise (testingInput)
        temp = Err.Number
        ErrorNumberOf = (temp = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorNumberOf", CBool(ErrorNumberOf), CBool(negateValue), testingValue, testingInput)
End Function

Public Function ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False) As Boolean
    Dim tempTestVal As Variant
    
    On Error Resume Next
        tempTestVal = CStr(Application.Evaluate(testingValue))
        tempTestVal = Replace(tempTestVal, "Error ", "")
        Err.Raise tempTestVal
        tempTestVal = Err.Description
        ErrorDescriptionOf = (tempTestVal = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorDescriptionOf", CBool(ErrorDescriptionOf), CBool(negateValue), tempTestVal, testingInput)
End Function

Public Function SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False, _
Optional ByVal updateFluentPath As Boolean = False) As Boolean

    Dim t1 As String
    Dim t2 As String
    
    t1 = TypeName(testingValue)
    t2 = TypeName(testingInput)
    
'    SameTypeAs = (t1 = t2)
    
    SameTypeAs = EqualTo(t1, t2, updateFluentPath:=False)
    
    Call addToTestAndFluentPath("SameTypeAs", CBool(SameTypeAs), CBool(negateValue), testingValue, testingInput)
End Function

Public Function IdenticalTo( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ParamArray args() As Variant) As Boolean

Dim i As Long
Dim temp As Variant
Dim b As Boolean

b = False

For i = LBound(args(0)) To UBound(args(0))
    If IsObject(args(0)(i)) Then
        Set temp = args(0)(i)
    Else
        temp = args(0)(i)
    End If

    b = EqualTo(InputToStringRecur(testingValue), InputToStringRecur(temp), updateFluentPath:=False)
    
    If Not b Then Exit For
Next i

IdenticalTo = b

Call addToTestAndFluentPath("IdenticalTo", CBool(IdenticalTo), CBool(negateValue), testingValue, args)

End Function

Public Function SameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ParamArray args() As Variant) As Boolean

Dim i As Long
Dim temp As Variant
Dim b As Boolean
Dim elem As Variant
Dim col As Collection
Dim col2 As Collection

b = False
Set col = New Collection

For Each elem In testingValue
    col.Add elem
Next elem

For i = LBound(args(0)) To UBound(args(0))
    If IsObject(args(0)(i)) Then
        Set temp = args(0)(i)
    Else
        temp = args(0)(i)
    End If
    
    Set col2 = New Collection
    
    For Each elem In temp
        col2.Add elem
    Next elem

    b = EqualTo(InputToStringRecur(col), InputToStringRecur(col2), updateFluentPath:=False)
    
    If Not b Then Exit For
Next i

SameElementsAs = b

Call addToTestAndFluentPath("SameElementsAs", CBool(SameElementsAs), CBool(negateValue), testingValue, args)

End Function

Private Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal Result As Boolean, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
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
        Case SameElementsAs_
            funcStr = "same elements as"
        Case SameTypeAs_
            funcStr = "same type as"
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
    
    finalFluentPath = Trim(Split(FluentPath, "=")(fluentElement))
    
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
                    ElseIf TypeOf elem Is Object  Or IsArray(elem) Then
                        If IsDataStructure(elem) Then
                            col.Add InputToStringRecur(elem)
                        ElseIf Not elem Is Nothing Then
                            col.Add TypeName(elem)
                        End If
                    
                    Else
                        col.Add elem
                    End If
                    
            Next elem
            
            ReDim strArgs(1 To col.Count)
                    
            For i = 1 To col.Count
                strArgs(i) = col(i)
            Next i
            
            tempNVTS = Join(strArgs, ", ")
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

Function DatastructureIsEmpty(v As Variant) As Boolean
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

Private Function InputToStringRecur(v As Variant) As String
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
                For Each elem In v.keys
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

Private Sub Class_Initialize()
    Set pExprDict = New Scripting.Dictionary
End Sub
