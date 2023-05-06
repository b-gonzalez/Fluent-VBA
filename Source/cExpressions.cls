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
Private pTestResult As cTestResult
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

Public Property Set setTestResult(ByVal value As cTestResult)
    Set pTestResult = value
End Property

Public Property Set setPrinting(ByVal value As IPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Public Function InDataStructure(ByVal ExpectedVal As Variant, ByVal DataStructure As Variant, Optional ByVal Method As flMethod = flRecursive, Optional NegateValue As Boolean = False) As Boolean
    Dim elem As Variant
    Dim Col As Collection
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = False
    
    If isDataStructure(DataStructure) Then
        If Method = flRecursive Then
            Set Col = getNestedElementsRecur(DataStructure)
        ElseIf Method = flIterative Then
            Set Col = getNestedElementsIter(DataStructure)
        End If
    End If
    
    For Each elem In Col
        If ExpectedVal = elem Then
            tempBool = True
            Exit For
        End If
    Next elem
    
    InDataStructure = tempBool
    
    Call addToDictAndFluentPath("InDataStructure", CBool(InDataStructure), CBool(NegateValue), ExpectedVal)
    
End Function

Private Function isDataStructure(ByVal DataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    
    dsTypeName = TypeName(DataStructure)
    tempBool = False
    
    Select Case True
        Case IsArray(DataStructure)
            tempBool = True
        Case TypeName(DataStructure) = "Collection", TypeName(DataStructure) = "Dictionary", TypeName(DataStructure) = "ArrayList"
            tempBool = True
    End Select
    
    isDataStructure = tempBool
End Function

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef Col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If Col Is Nothing Then
        Set Col = New Collection
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
                Set Col = getNestedElementsRecur(elem, Col)
            Else
                Col.Add elem
            End If
        Next elem
    Else
        Col.Add elem
    End If
    
    Set getNestedElementsRecur = Col
End Function

Private Function getNestedElementsIter(ByVal DataStructure As Variant) As Collection
    Dim elem As Variant
    Dim Col As Collection
    Dim tempCol As Collection
    Dim tempDataStructure As Variant
    Dim i As Long

    If Col Is Nothing Then
        Set Col = New Collection
        Set tempCol = New Collection
    End If

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
            tempCol.Add elem
        Else
            Col.Add elem
        End If
    Next elem

    Do While tempCol.Count > 0
        For i = 1 To tempCol.Count
            If isDataStructure(tempCol(i)) Then
                For Each elem In tempCol(i)
                    tempCol.Add elem
                Next elem
                tempCol.Remove i
                Exit For
            Else
                Col.Add tempCol(i)
                tempCol.Remove i
                Exit For
            End If
        Next i
    Loop

    Set getNestedElementsIter = Col
End Function

Public Function GreaterThan(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThan = (ExpectedVal > ActualVal)
    
    Call addToDictAndFluentPath("GreaterThan", CBool(GreaterThan), CBool(NegateValue), ExpectedVal, ActualVal)
    
End Function

Public Function GreaterThanOrEqualTo(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (ExpectedVal >= ActualVal)
    
    Call addToDictAndFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function LessThan(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (ExpectedVal < ActualVal)
    
    Call addToDictAndFluentPath("LessThan", CBool(LessThan), CBool(NegateValue), ExpectedVal, ActualVal)

End Function

Public Function LessThanOrEqualTo(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (ExpectedVal <= ActualVal)
    
    Call addToDictAndFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function EqualTo(ByVal ExpectedVal As Variant, ByVal ActualVal As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim tempBool As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = (ExpectedVal = ActualVal)
    
    If pMeta.ApproximateEqual Then
        If TypeName(ExpectedVal) <> TypeName(ActualVal) Then
            If IsNumeric(ExpectedVal) And IsNumeric(ActualVal) Then
                tempBool = (Abs(ExpectedVal - ActualVal) <= pMeta.Epsilon)
            ElseIf VarType(ExpectedVal) = vbString Or VarType(ActualVal) = vbString Then
                tempBool = (StrComp(CStr(ExpectedVal), CStr(ActualVal), vbTextCompare) = 0)
            End If
        End If
    End If
    
    EqualTo = tempBool
    
    Call addToDictAndFluentPath("EqualTo", CBool(EqualTo), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function Contain(ByVal ExpectedVal As Variant, ByVal ActualVal As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If ExpectedVal Like "*" & ActualVal & "*" Then
        Contain = True
    End If
    
    Call addToDictAndFluentPath("Contain", CBool(Contain), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function StartWith(ByVal ExpectedVal As Variant, ByVal ActualVal As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(ActualVal)
    
    If Left(ExpectedVal, valLength) = CStr(ActualVal) Then
        StartWith = True
    End If
    
    Call addToDictAndFluentPath("StartWith", CBool(StartWith), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function EndWith(ByVal ExpectedVal As Variant, ByVal ActualVal As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(ActualVal)
    
    If Right(ExpectedVal, valLength) = CStr(ActualVal) Then
        EndWith = True
    End If
    
    Call addToDictAndFluentPath("EndWith", CBool(EndWith), CBool(NegateValue), ExpectedVal, ActualVal)
    
End Function

Public Function LengthOf(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(ExpectedVal)) = ActualVal)
    
    Call addToDictAndFluentPath("LengthOf", CBool(LengthOf), CBool(NegateValue), ExpectedVal, ActualVal)
            
End Function

Public Function MaxLengthOf(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(ExpectedVal)) <= ActualVal)
    
    Call addToDictAndFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(NegateValue), ExpectedVal, ActualVal)
            
End Function

Public Function MinLengthOf(ByVal ExpectedVal As Double, ByVal ActualVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(ExpectedVal)) >= ActualVal)
    
    Call addToDictAndFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(NegateValue), ExpectedVal, ActualVal)
        
End Function

Public Function Something(ByVal ExpectedVal As Variant, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not ExpectedVal Is Nothing)
    
    Call addToDictAndFluentPath("Something", CBool(Something), CBool(NegateValue), TypeName(ExpectedVal))
        
End Function

Public Function Between(ByVal ExpectedVal As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(ExpectedVal, LowerVal) And LessThanOrEqualTo(ExpectedVal, HigherVal))
    
    Call addToDictAndFluentPath("Between", CBool(Between), CBool(NegateValue), Expected:=ExpectedVal, LowerVal:=LowerVal, HigherVal:=HigherVal)
        
End Function

Public Function OneOf(ByVal ExpectedVal As Variant, ByVal NegateValue As Boolean, ParamArray Args() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    Dim expectedStr As String
        
    arr = Args(0)
    tempBool = False
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            expectedStr = expectedStr & TypeName(arr(i)) & ", "
        Else
            expectedStr = expectedStr & arr(i) & ", "
        End If
    
        If IsObject(arr(i)) Or IsObject(ExpectedVal) Then
            If TypeName(arr(i)) = TypeName(ExpectedVal) Then
                tempBool = True
            End If
        ElseIf arr(i) = ExpectedVal Then
            tempBool = True
            'Exit For
        End If
    Next i
    
    expectedStr = Left(expectedStr, Len(expectedStr) - 2)
    
    OneOf = tempBool
    
    Call addToDictAndFluentPath("OneOf", CBool(OneOf), CBool(NegateValue), expectedStr)
        
End Function

Public Function LengthBetween(ByVal ExpectedVal As Double, ByVal LowerVal As Double, ByVal HigherVal As Double, Optional ByVal NegateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(ExpectedVal))
    
    LengthBetween = (MinLengthOf(tempOrigVal, LowerVal) And MaxLengthOf(tempOrigVal, HigherVal))
    
    Call addToDictAndFluentPath("LengthBetween", CBool(LengthBetween), CBool(NegateValue), Expected:=ExpectedVal, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo(ByVal ExpectedVal As Variant, ByVal ActualVal As Variant, Optional ByVal NegateValue = False) As Variant
    Dim FluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(ExpectedVal) = ActualVal)

    Call addToDictAndFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(NegateValue), ExpectedVal, ActualVal)
    
End Function

Public Function Alphabetic(ByVal ExpectedVal As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    
    For i = 1 To Len(ExpectedVal)
        tempChar = Mid(ExpectedVal, i, 1)
        If tempChar Like "[A-Za-z]" Then
            tempAlph = tempAlph & tempChar
        End If
    Next i
    
    Alphabetic = (ExpectedVal = tempAlph)
    
    Call addToDictAndFluentPath("Alphabetic", CBool(Alphabetic), CBool(NegateValue), TypeName(ExpectedVal))
End Function

Public Function Numeric(ByVal ExpectedVal As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    
    For i = 1 To Len(ExpectedVal)
        tempChar = Mid(ExpectedVal, i, 1)
        If tempChar Like "[0-9]" Then
            tempNum = tempNum & tempChar
        End If
    Next i
    
    Numeric = (ExpectedVal = tempNum)
     
    Call addToDictAndFluentPath("Numeric", CBool(Numeric), CBool(NegateValue), TypeName(ExpectedVal))
End Function

Public Function Alphanumeric(ByVal ExpectedVal As Variant, Optional ByVal NegateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    
    b = False
    
    For i = 1 To Len(ExpectedVal)
        tempChar = Mid(ExpectedVal, i, 1)
        If tempChar Like "[A-Za-z]" Then
            tempAlph = tempAlph & tempChar
        ElseIf tempChar Like "[0-9]" Then
            tempNum = tempNum & tempChar
        End If
    Next i
    
    If Len(tempAlph) > 1 And Len(tempNum) > 1 Then
        If Len(tempAlph) + Len(tempNum) = Len(ExpectedVal) Then
            b = True
        End If
    End If
    
    Alphanumeric = b
    
    Call addToDictAndFluentPath("Alphanumeric", CBool(Alphanumeric), CBool(NegateValue), TypeName(ExpectedVal))
End Function


Private Sub addToDictAndFluentPath(ByVal FunctionName As String, ByVal Result As Boolean, ByVal NegateValue As Boolean, ByVal Expected As Variant, Optional ByVal actual As Variant, Optional LowerVal As Variant, Optional HigherVal As Variant)

    Set pExprDict = New Scripting.Dictionary

    pExprDict.Add "functionName", FunctionName
    pExprDict.Add "result", Result
    pExprDict.Add "negateValue", NegateValue
    pExprDict.Add "expected", Expected
    
    If Not IsMissing(actual) Then
        pExprDict.Add "actual", actual
    End If
    
    If Not IsMissing(LowerVal) Then
        pExprDict.Add "lowerVal", LowerVal
    End If
    
    If Not IsMissing(HigherVal) Then
        pExprDict.Add "higherVal", HigherVal
    End If
    
    Call addFluentPathNew(pExprDict)

End Sub

Private Sub addFluentPathNew(ByVal dict As Scripting.Dictionary)
    Dim FluentPath As String
    Dim shouldStr As String
    Dim fluentBool As Boolean
    Dim fluentTestName As String
    Dim fluentCategory As String
    Dim newVal As Variant
    Dim ExpectedVal As Variant
    Dim LowerVal As Variant
    Dim HigherVal As Variant
    Dim FunctionName As String
    Dim finalFluentPath
    Dim fluentElement As Long
    
    
    FunctionName = dict("functionName")
    
    If dict("negateValue") Then
        dict("result") = Not dict("result")
'        dict("expected") = Not dict("expected")
        dict("actual") = Not dict("actual")
        fluentBool = dict("result")
        shouldStr = "should not"
    Else
        fluentBool = dict("result")
        shouldStr = "should"
    End If
    
    FluentPath = ""
    
    If pPrinting.Category <> "" Then
        FluentPath = FluentPath & pPrinting.Category & "="
        fluentElement = 2
    Else
        fluentElement = 1
    End If
    
    FluentPath = FluentPath & CLng(fluentBool) & "="
    
    If pPrinting.TestName <> "" Then FluentPath = FluentPath & pPrinting.TestName & " "
    
    ExpectedVal = dict("expected")
    
    If Not dict.exists("actual") Then
        dict.Add "actual", ""
    End If
    
    newVal = dict("actual")
        
    Select Case FunctionName
        Case "Alphabetical"
            FluentPath = FluentPath & shouldStr & " be alphabetical "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "Alphanumeric"
            FluentPath = FluentPath & shouldStr & " be alphanumeric "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "Between"
            LowerVal = dict("lowerVal")
            HigherVal = dict("higherVal")
            FluentPath = FluentPath & shouldStr & " be between "
            FluentPath = FluentPath & LowerVal & " and " & HigherVal & ". Actual: " & ExpectedVal
        Case "Contain"
            FluentPath = FluentPath & shouldStr & " contain "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "EqualTo"
            FluentPath = FluentPath & shouldStr & " be equal to "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "EndWith"
            FluentPath = FluentPath & shouldStr & " end with "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "EvaluateTo"
            FluentPath = FluentPath & shouldStr & " evaluate to "
            FluentPath = FluentPath & CStr(newVal) & ". Actual: " & ExpectedVal
        Case "GreaterThan"
            FluentPath = FluentPath & shouldStr & " be greater than "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "GreaterThanOrEqualTo"
            FluentPath = FluentPath & shouldStr & " be greater than or equal to "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "InDataStructure"
            FluentPath = FluentPath & shouldStr & " be in datastructure. "
        Case "LengthBetween"
            LowerVal = dict("lowerVal")
            HigherVal = dict("higherVal")
            FluentPath = FluentPath & shouldStr & " have length between "
            FluentPath = FluentPath & LowerVal & " and " & HigherVal & ". Actual: " & Len(ExpectedVal)
        Case "LengthOf"
            FluentPath = FluentPath & shouldStr & " have length of "
            FluentPath = FluentPath & newVal & ". Actual: " & Len(ExpectedVal)
        Case "LessThan"
            FluentPath = FluentPath & shouldStr & " be less than "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "LessThanOrEqualTo"
            FluentPath = FluentPath & shouldStr & " be less than or equal to "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "MaxLengthOf"
            FluentPath = FluentPath & shouldStr & " have max length of "
            FluentPath = FluentPath & newVal & ". Actual: " & Len(ExpectedVal)
        Case "MinLengthOf"
            FluentPath = FluentPath & shouldStr & " have min length of "
            FluentPath = FluentPath & newVal & ". Actual: " & Len(ExpectedVal)
        Case "Numeric"
            FluentPath = FluentPath & shouldStr & " be numeric "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
        Case "OneOf"
            FluentPath = FluentPath & shouldStr & " be one of: " & dict("expected")
        Case "Something"
            FluentPath = FluentPath & shouldStr & " be something: " & dict("expected")
        Case "StartWith"
            FluentPath = FluentPath & shouldStr & " start with "
            FluentPath = FluentPath & newVal & ". Actual: " & ExpectedVal
    End Select
    
    finalFluentPath = Split(FluentPath, "=")(fluentElement)
    
    iPrint.TestCount = iPrint.TestCount + 1
    
    iPrint.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add FluentPath
    
    pExprDict.Add "fluentPath", finalFluentPath
'    pExprDict.Add "fluentPath", fluentPath
    
    Call pTestResult.CheckTest(fluentBool, pExprDict)
    
    'Set pExprDict = CreateObject("Scripting.Dictionary")
    
    Set pExprDict = Nothing
End Sub

Private Sub Class_Initialize()
    Set pExprDict = CreateObject("Scripting.Dictionary")
End Sub
