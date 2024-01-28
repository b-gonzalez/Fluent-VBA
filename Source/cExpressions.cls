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

Enum flMethod
    flRecursive
    flIterative
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

Public Function InDataStructures(ByVal TestingValue As Variant, ByVal Method As flMethod, NegateValue As Boolean, ParamArray args() As Variant) As Boolean
    Dim i As Long
    Dim temp_bool As Boolean
    Dim IsLastElem As Boolean
    Dim col As VBA.Collection
    
    Set col = getNestedElementsRecur(args)
    
    temp_bool = False
    
'    For i = 1 To col.Count
'
    temp_bool = Me.InDataStructure(TestingValue, col, Method, NegateValue, updateFluentPath:=False)
'
'        If temp_bool Then
'            Exit For
'        End If
'    Next i
    
    InDataStructures = temp_bool
    
    Call addToDictAndFluentPath("InDataStructures", CBool(InDataStructures), CBool(NegateValue), TestingValue)
    
End Function

Public Function InDataStructure( _
ByVal TestingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal Method As flMethod = flRecursive, _
Optional NegateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = False
    
    If isDataStructure(DataStructure) Then
        If Method = flRecursive Then
            Set col = getNestedElementsRecur(DataStructure)
        ElseIf Method = flIterative Then
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
    
'    If Not IsMissing(IsLastElem) Then
'        If IsLastElem Then
'            Call addToDictAndFluentPath("InDataStructure", CBool(InDataStructure), CBool(NegateValue), TestingValue)
'        End If
'    Else
    If updateFluentPath Then
        Call addToDictAndFluentPath("InDataStructure", CBool(InDataStructure), CBool(NegateValue), TestingValue)
    End If
'    End If
    
End Function

Private Function isDataStructure(ByVal DataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    dsTypeName = TypeName(DataStructure)
    tempBool = False
    
'    Select Case True
'        Case IsArray(DataStructure)
'            tempBool = True
'        Case TypeName(DataStructure) = "Collection", TypeName(DataStructure) = "Dictionary", TypeName(DataStructure) = "ArrayList"
'            tempBool = True
'    End Select

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

Public Function GreaterThan(ByVal TestingValue As Double, ByVal TestingInput As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThan = (TestingValue > TestingInput)
    
    Call addToDictAndFluentPath("GreaterThan", CBool(GreaterThan), CBool(NegateValue), TestingValue, TestingInput)
    
End Function

Public Function GreaterThanOrEqualTo( _
ByVal TestingValue As Double, _
ByVal TestingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (TestingValue >= TestingInput)
    
    If updateFluentPath Then
        Call addToDictAndFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(NegateValue), TestingValue, TestingInput)
    End If
        
End Function

Public Function LessThan(ByVal TestingValue As Double, ByVal TestingInput As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (TestingValue < TestingInput)
    
    Call addToDictAndFluentPath("LessThan", CBool(LessThan), CBool(NegateValue), TestingValue, TestingInput)

End Function

Public Function LessThanOrEqualTo( _
ByVal TestingValue As Double, _
ByVal TestingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (TestingValue <= TestingInput)
       
    If updateFluentPath Then
        Call addToDictAndFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(NegateValue), TestingValue, TestingInput)
    End If
        
End Function

Public Function EqualTo(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = (TestingValue = TestingInput)
    
    If pMeta.ApproximateEqual Then
        If TypeName(TestingValue) <> TypeName(TestingInput) Then
            If IsNumeric(TestingValue) And IsNumeric(TestingInput) Then
                tempBool = (Abs(TestingValue - TestingInput) <= pMeta.Epsilon)
            ElseIf VarType(TestingValue) = vbString Or VarType(TestingInput) = vbString Then
                tempBool = (StrComp(CStr(TestingValue), CStr(TestingInput), vbTextCompare) = 0)
            End If
        End If
    End If
    
    EqualTo = tempBool
    
    Call addToDictAndFluentPath("EqualTo", CBool(EqualTo), CBool(NegateValue), TestingValue, TestingInput)
        
End Function

Public Function Contain(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If TestingValue Like "*" & TestingInput & "*" Then
        Contain = True
    End If
    
    Call addToDictAndFluentPath("Contain", CBool(Contain), CBool(NegateValue), TestingValue, TestingInput)
        
End Function

Public Function StartWith(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(TestingInput)
    
    If Left(TestingValue, valLength) = CStr(TestingInput) Then
        StartWith = True
    End If
    
    Call addToDictAndFluentPath("StartWith", CBool(StartWith), CBool(NegateValue), TestingValue, TestingInput)
        
End Function

Public Function EndWith(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(TestingInput)
    
    If Right(TestingValue, valLength) = CStr(TestingInput) Then
        EndWith = True
    End If
    
    Call addToDictAndFluentPath("EndWith", CBool(EndWith), CBool(NegateValue), TestingValue, TestingInput)
    
End Function

Public Function LengthOf(ByVal TestingValue As Double, ByVal TestingInput As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(TestingValue)) = TestingInput)
    
    Call addToDictAndFluentPath("LengthOf", CBool(LengthOf), CBool(NegateValue), TestingValue, TestingInput)
            
End Function

Public Function MaxLengthOf( _
ByVal TestingValue As Double, _
ByVal TestingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(TestingValue)) <= TestingInput)
    
    If updateFluentPath Then
        Call addToDictAndFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(NegateValue), TestingValue, TestingInput)
    End If
            
End Function

Public Function MinLengthOf( _
ByVal TestingValue As Double, _
ByVal TestingInput As Double, _
Optional ByVal NegateValue As Boolean = False, _
Optional updateFluentPath As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(TestingValue)) >= TestingInput)
        
    If updateFluentPath Then
        Call addToDictAndFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(NegateValue), TestingValue, TestingInput)
    End If
        
End Function

Public Function Something(ByVal TestingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not TestingValue Is Nothing)
    
    Call addToDictAndFluentPath("Something", CBool(Something), CBool(NegateValue), TypeName(TestingValue))
        
End Function

Public Function Between(ByVal TestingValue As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(TestingValue, LowerVal, updateFluentPath:=False) And LessThanOrEqualTo(TestingValue, HigherVal, updateFluentPath:=False))
    
    Call addToDictAndFluentPath("Between", CBool(Between), CBool(NegateValue), TestingValue:=TestingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
        
End Function

Public Function OneOf(ByVal TestingValue As Variant, ByVal NegateValue As Boolean, ParamArray args() As Variant) As Boolean
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
        If IsObject(arr(i)) Then
            TestingInputStr = TestingInputStr & TypeName(arr(i)) & ", "
        Else
            TestingInputStr = TestingInputStr & arr(i) & ", "
        End If
    
        If IsObject(arr(i)) Or IsObject(TestingValue) Then
            If TypeName(arr(i)) = TypeName(TestingValue) Then
                tempBool = True
            End If
        ElseIf arr(i) = TestingValue Then
            tempBool = True
            'Exit For
        End If
    Next i
    
    TestingInputStr = Left(TestingInputStr, Len(TestingInputStr) - 2)
    
    OneOf = tempBool
    
    If TypeOf TestingValue Is Object  Then TestingValue = TypeName(TestingValue)
    
    Call addToDictAndFluentPath("OneOf", CBool(OneOf), CBool(NegateValue), TestingValue, TestingInputStr)
        
End Function

Public Function LengthBetween(ByVal TestingValue As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(TestingValue))
    
    LengthBetween = (MinLengthOf(tempOrigVal, LowerVal, updateFluentPath:=False) And MaxLengthOf(tempOrigVal, HigherVal, updateFluentPath:=False))
    
    Call addToDictAndFluentPath("LengthBetween", CBool(LengthBetween), CBool(NegateValue), TestingValue:=TestingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue = False) As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(TestingValue) = TestingInput)

    Call addToDictAndFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(NegateValue), TestingValue, TestingInput)
    
End Function

Public Function Alphabetic(ByVal TestingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
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
    
    Call addToDictAndFluentPath("Alphabetic", CBool(Alphabetic), CBool(NegateValue), TypeName(TestingValue))
End Function

Public Function Numeric(ByVal TestingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
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
     
    Call addToDictAndFluentPath("Numeric", CBool(Numeric), CBool(NegateValue), TypeName(TestingValue))
End Function

Public Function Alphanumeric(ByVal TestingValue As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
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
    
    Call addToDictAndFluentPath("Alphanumeric", CBool(Alphanumeric), CBool(NegateValue), TypeName(TestingValue))
End Function

Public Function Erroneous(ByVal TestingValue As Variant, Optional ByVal NegateValue = False) As Boolean
    Dim temp As Variant
    
    temp = False
    
    If TypeOf TestingValue Is ErrObject Then
        temp = (TestingValue.Number <> 0)
    Else
        On Error Resume Next
            temp = Application.Evaluate(TestingValue)
            temp = (TypeName(temp) = "Error")
        On Error GoTo 0
    End If
    
    Erroneous = temp
    
    Call addToDictAndFluentPath("Erroneous", CBool(Erroneous), CBool(NegateValue), TypeName(temp), "Error")
End Function

Public Function ErrorNumberOf(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue = False) As Boolean
    Dim temp As Variant
    Dim e1 As Variant
    
    On Error Resume Next
        temp = Application.Evaluate(TestingValue)
        Err.Raise (TestingInput)
        temp = Err.Number
        ErrorNumberOf = (temp = TestingInput)
    On Error GoTo 0
    
    Call addToDictAndFluentPath("ErrorNumberOf", CBool(ErrorNumberOf), CBool(NegateValue), TestingInput, temp)
End Function

Public Function ErrorDescriptionOf(ByVal TestingValue As Variant, ByVal TestingInput As Variant, Optional ByVal NegateValue = False) As Boolean
    Dim temp As Variant
    
    On Error Resume Next
        temp = CStr(Application.Evaluate(TestingValue))
        temp = Replace(temp, "Error ", "")
        Err.Raise temp
        temp = Err.Description
        ErrorDescriptionOf = (Err.Description = TestingInput)
    On Error GoTo 0
    
    Call addToDictAndFluentPath("ErrorDescriptionOf", CBool(ErrorDescriptionOf), CBool(NegateValue), TestingInput, temp)
End Function

Private Sub addToDictAndFluentPath( _
    ByVal FunctionName As String, _
    ByVal Result As Boolean, _
    ByVal NegateValue As Boolean, _
    ByVal TestingValue As Variant, _
    Optional ByVal TestingInput As Variant, _
    Optional LowerVal As Variant, _
    Optional HigherVal As Variant)
    
    Dim test As cTest
    Dim shouldStr As String
    
    Set test = New cTest
    shouldStr = "Should"

'    Set pExprDict = New Scripting.Dictionary
'
'    pExprDict.Add "functionName", FunctionName
'    pExprDict.Add "result", Result
'    pExprDict.Add "negateValue", NegateValue
'    pExprDict.Add "testingValue", TestingValue
'
'    If Not IsMissing(TestingInput) Then
'        pExprDict.Add "testingInput", TestingInput
'    End If
'
'    If Not IsMissing(LowerVal) Then
'        pExprDict.Add "lowerVal", LowerVal
'    End If
'
'    If Not IsMissing(HigherVal) Then
'        pExprDict.Add "higherVal", HigherVal
'    End If

    If NegateValue Then
        Result = Not Result
        shouldStr = "should not"
    End If

    With test
        .FunctionName = FunctionName
        .Result = Result
        .TestingValue = TestingValue
        .TestingInput = Empty
    End With
    
    If Not IsMissing(TestingInput) Then
        test.TestingInput = TestingInput
    End If

    If Not IsMissing(LowerVal) Then
        test.LowerVal = LowerVal
    End If

    If Not IsMissing(HigherVal) Then
        test.HigherVal = HigherVal
    End If
    
    
    Call addFluentPathNew(test, shouldStr)

End Sub

Private Sub addFluentPathNew(ByVal test As cTest, ByVal shouldStr As String)
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
    
    With test
        FluentPath = iPrint.getFluentPath(FluentPath, shouldStr, test)
    End With
    
    finalFluentPath = Split(FluentPath, "=")(fluentElement)
    
    iPrint.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add FluentPath
    
    test.FluentPath = finalFluentPath
    
    Call pTests.CheckTest(test)
    
End Sub

Private Sub Class_Initialize()
    Set pExprDict = New Scripting.Dictionary
End Sub
