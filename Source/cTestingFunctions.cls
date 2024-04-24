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

Private pMeta As cMeta
Private pTests As cTests
Private pPrinting As cPrinting
Private iPrint As IPrinting
Private pApproximateEqual As Boolean
Private pEpsilon As Double
'Private pExprDict As Scripting.Dictionary

Private pTestDictCounter As Scripting.Dictionary

Private pTestValue As Variant
Private pFluentPath As String
Private pExpressions As cExpressions

Public Enum flAlgorithm
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
'    Set pMeta.Tests.TestDictCounter = pTestDictCounter
End Property

Public Property Set Expressions(value As cExpressions)
    Set pExpressions = value
End Property

Public Function InDataStructures( _
ByVal testingValue As Variant, _
ByVal method As flAlgorithm, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim col As VBA.Collection
    Dim i As Long
    
    InDataStructures = Null
    
    For i = LBound(testingInput) To UBound(testingInput)
        If pExpressions.IsDataStructure(testingInput(i)) Then
            Set col = getNestedElementsRecur(testingInput)
        
            InDataStructures = Me.InDataStructure(testingValue, col, method, negateValue, updateFluentPath:=False)
        End If
    Next i
    
    
    
    Call pExpressions.addToTestAndFluentPath("InDataStructures", InDataStructures, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function InDataStructure( _
ByVal testingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flAlgorithm = flRecursive, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    
    InDataStructure = Null
    
    If pExpressions.IsDataStructure(DataStructure) Then
        tempBool = False
    
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
        
        InDataStructure = tempBool

    End If
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("InDataStructure", InDataStructure, CBool(negateValue), testingValue, DataStructure)
    End If
    
End Function

'Private Function IsDataStructure(ByVal DataStructure As Variant)
'    Dim dsTypeName As String
'    Dim tempBool As Boolean
'    Dim fluentBool As Boolean
'    Dim elem As Variant
'    Dim tempErr As Long
'
'    dsTypeName = TypeName(DataStructure)
'    tempBool = False
'
'    On Error Resume Next
'        For Each elem In DataStructure
'            Exit For
'        Next elem
'
'        tempErr = Err.Number
'
'        tempBool = (tempErr = 0)
'
'    On Error GoTo 0
'
'    IsDataStructure = tempBool
'End Function

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
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

    Set getNestedElementsIter = col
End Function

Public Function GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
    
    GreaterThan = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        GreaterThan = (testingValue > testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("GreaterThan", GreaterThan, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function GreaterThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    GreaterThanOrEqualTo = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        GreaterThanOrEqualTo = (testingValue >= testingInput)
    End If
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("GreaterThanOrEqualTo", GreaterThanOrEqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function LessThan( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    LessThan = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        LessThan = (testingValue < testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("LessThan", LessThan, CBool(negateValue), testingValue, testingInput)

End Function

Public Function LessThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    LessThanOrEqualTo = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        LessThanOrEqualTo = (testingValue <= testingInput)
    End If
       
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("LessThanOrEqualTo", LessThanOrEqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant

    Dim tempBool As Boolean
    
    EqualTo = Null
    
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
        
        EqualTo = tempBool
    End If
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("EqualTo", EqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant

    Contain = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        Contain = CStr(testingValue) Like "*" & CStr(testingInput) & "*"
    End If
        
    Call pExpressions.addToTestAndFluentPath("Contain", Contain, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    
    StartWith = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        valLength = Len(testingInput)
        StartWith = Left$(testingValue, valLength) = CStr(testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("StartWith", StartWith, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    
    EndWith = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        valLength = Len(testingInput)
        EndWith = Right$(testingValue, valLength) = CStr(testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("EndWith", EndWith, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    LengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        LengthOf = (Len(CStr(testingValue)) = testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("LengthOf", LengthOf, CBool(negateValue), testingValue, testingInput)
            
End Function

Public Function MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    MaxLengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        MaxLengthOf = (Len(CStr(testingValue)) <= testingInput)
    End If
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("MaxLengthOf", MaxLengthOf, CBool(negateValue), testingValue, testingInput)
    End If
            
End Function

Public Function MinLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    MinLengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        MinLengthOf = (Len(CStr(testingValue)) >= testingInput)
    End If
        
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("MinLengthOf", MinLengthOf, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    
    Something = Null
    
    If IsObject(testingValue) Then
        Something = (Not testingValue Is Nothing)
    End If
    
    Call pExpressions.addToTestAndFluentPath("Something", Something, CBool(negateValue), testingValue)
        
End Function

Public Function Between( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    Between = Null
    
    If IsNumeric(testingValue) Then
        Between = (GreaterThanOrEqualTo(testingValue, LowerVal, updateFluentPath:=False) And LessThanOrEqualTo(testingValue, HigherVal, updateFluentPath:=False))
    End If
    
    Call pExpressions.addToTestAndFluentPath("Between", Between, CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
End Function

Public Function OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim tempBool As Boolean
    
    OneOf = Null
        
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
    
    OneOf = tempBool
    
    If TypeOf testingValue Is Object  Then testingValue = TypeName(testingValue)
    
    Call pExpressions.addToTestAndFluentPath("OneOf", OneOf, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function LengthBetween( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim tempTestVal As Double
    
    LengthBetween = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        tempTestVal = Len(CStr(testingValue))
        LengthBetween = (MinLengthOf(tempTestVal, LowerVal, updateFluentPath:=False) And MaxLengthOf(tempTestVal, HigherVal, updateFluentPath:=False))
    End If
    
    Call pExpressions.addToTestAndFluentPath("LengthBetween", LengthBetween, CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False) As Variant

    EvaluateTo = Null

    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        EvaluateTo = (Excel.Evaluate(testingValue) = testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("EvaluateTo", EvaluateTo, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    
    Alphabetic = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            End If
        Next i
        
        Alphabetic = (testingValue = tempAlph)
    End If
    
    Call pExpressions.addToTestAndFluentPath("Alphabetic", Alphabetic, CBool(negateValue), testingValue)
End Function

Public Function Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    
    Numeric = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) And Not IsNull(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            If tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        Numeric = (testingValue = tempNum)
    End If
     
    Call pExpressions.addToTestAndFluentPath("Numeric", Numeric, CBool(negateValue), testingValue)
End Function

Public Function Alphanumeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    
    Alphanumeric = Null
    
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
        
        Alphanumeric = b
    End If
    
    Call pExpressions.addToTestAndFluentPath("Alphanumeric", Alphanumeric, CBool(negateValue), testingValue)
End Function

Public Function Erroneous(ByVal testingValue As Variant, Optional ByVal negateValue = False) As Variant
    Dim tempTestingVal As Variant
    Dim tempErrType As Variant
    Dim tempErrObj As ErrObject
    
    Erroneous = Null
    
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
        
        Erroneous = tempTestingVal
        
        If TypeOf testingValue Is Object  Then
            tempErrType = TypeName(testingValue)
        Else
            tempErrType = testingValue
        End If
    End If
    
    Call pExpressions.addToTestAndFluentPath("Erroneous", Erroneous, CBool(negateValue), tempErrType)
End Function

Public Function ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue = False) As Variant
    Dim temp As Variant
    Dim tempErrObj As ErrObject
    
    ErrorNumberOf = Null
    
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
        
        ErrorNumberOf = (temp = testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("ErrorNumberOf", ErrorNumberOf, CBool(negateValue), testingValue, testingInput)
End Function

Public Function ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue = False) As Variant
    Dim tempTestVal As Variant
    Dim tempErrObj As ErrObject
    
    ErrorDescriptionOf = Null
    
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
        
        ErrorDescriptionOf = (tempTestVal = testingInput)
    End If
    
    Call pExpressions.addToTestAndFluentPath("ErrorDescriptionOf", ErrorDescriptionOf, CBool(negateValue), tempTestVal, testingInput)
End Function

Public Function SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant

    Dim t1 As String
    Dim t2 As String
    
    SameTypeAs = Null
    
    t1 = TypeName(testingValue)
    t2 = TypeName(testingInput)
    
    SameTypeAs = EqualTo(t1, t2, updateFluentPath:=False)
    
    Call pExpressions.addToTestAndFluentPath("SameTypeAs", SameTypeAs, CBool(negateValue), testingValue, testingInput)
End Function

Public Function IdenticalTo( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant, _
ByVal updateFluentPath As Boolean) As Variant
    
    IdenticalTo = Null
    
    If pExpressions.IsDataStructure(testingInput) And pExpressions.IsDataStructure(testingInput) Then
        IdenticalTo = EqualTo(pExpressions.InputToStringRecur(testingValue), pExpressions.InputToStringRecur(testingInput), updateFluentPath:=False)
    End If
    
    If updateFluentPath Then
        Call pExpressions.addToTestAndFluentPath("IdenticalTo", IdenticalTo, CBool(negateValue), testingValue, testingInput)
    End If

End Function

Public Function ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim elem As Variant
    Dim col As Collection
    Dim col2 As Collection
    
    ExactSameElementsAs = Null
    
    If pExpressions.IsDataStructure(testingValue) And pExpressions.IsDataStructure(testingInput) Then
        Set col = New Collection
        Set col2 = New Collection
        
        For Each elem In testingValue
            col.Add elem
        Next elem
        
        For Each elem In testingInput
            col2.Add elem
        Next elem
        
        ExactSameElementsAs = IdenticalTo(col, negateValue, col2, False)
    End If
    
    Call pExpressions.addToTestAndFluentPath("ExactSameElementsAs", ExactSameElementsAs, CBool(negateValue), testingValue, testingInput)

End Function

Public Function SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim tempStr As String
    
    SameUniqueElementsAs = Null
    
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
        
        SameUniqueElementsAs = IdenticalTo(d.Keys, negateValue, d.Items, False)
    End If
    
    Call pExpressions.addToTestAndFluentPath("SameUniqueElementsAs", SameUniqueElementsAs, CBool(negateValue), testingValue, testingInput)

End Function

Public Function SameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim i As Long
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim tempStr As String
    
    SameElementsAs = Null
    
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
        
        SameElementsAs = IdenticalTo(d, negateValue, d2, False)
    End If
    
    Call pExpressions.addToTestAndFluentPath("SameElementsAs", SameElementsAs, CBool(negateValue), testingValue, testingInput)

End Function
