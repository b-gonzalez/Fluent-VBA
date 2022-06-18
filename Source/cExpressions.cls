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
Private pExprDict As Object

Private pTestValue As Variant
Private pFluentPath As String

Enum flMethod
    flRecursive
    flIterative
End Enum

Public Property Let TestValue(value As Variant)
    pTestValue = value
End Property

Public Property Set TestValue(value As Variant)
    Set pTestValue = value
End Property

Public Property Get TestValue() As Variant
    If IsObject(pTestValue) Then
        Set TestValue = pTestValue
    Else
        TestValue = pTestValue
    End If
End Property

Public Property Set setMeta(value As cMeta)
    Set pMeta = value
End Property

Public Property Set setTestResult(value As cTestResult)
    Set pTestResult = value
End Property

Public Property Set setPrinting(value As IPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Public Function InDataStructure(origVal As Variant, dataStructure As Variant, Optional method As flMethod = flRecursive, Optional negateValue As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = False
    
    If isDataStructure(dataStructure) Then
        If method = flRecursive Then
            Set col = getNestedElementsRecur(dataStructure)
        ElseIf method = flIterative Then
            Set col = getNestedElementsIter(dataStructure)
        End If
    End If
    
    For Each elem In col
        If origVal = elem Then
            tempBool = True
            Exit For
        End If
    Next elem
    
    InDataStructure = tempBool
    
    pExprDict.Add "functionName", "InDataStructure"
    pExprDict.Add "result", CBool(InDataStructure)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    'pExprDict.Add "actual", newVal
    
    'Call addFluentPath("InDataStructure", CBool(InDataStructure), CBool(negateValue), origVal)
    Call addFluentPathNew(pExprDict)
    
End Function

Private Function isDataStructure(ByVal dataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    
    dsTypeName = TypeName(dataStructure)
    tempBool = False
    
    Select Case True
        Case IsArray(dataStructure)
            tempBool = True
        Case TypeName(dataStructure) = "Collection", TypeName(dataStructure) = "Dictionary", TypeName(dataStructure) = "ArrayList"
            tempBool = True
    End Select
    
    isDataStructure = tempBool
End Function

Private Function getNestedElementsRecur(dataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If isDataStructure(dataStructure) Then
        If TypeName(dataStructure) = "Dictionary" Then
            tempDataStructure = dataStructure.Items
        Else
            If IsObject(dataStructure) Then
                Set tempDataStructure = dataStructure
            Else
                tempDataStructure = dataStructure
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

Private Function getNestedElementsIter(dataStructure As Variant) As Collection
    Dim elem As Variant
    Dim col As Collection
    Dim tempCol As Collection
    Dim tempDataStructure As Variant
    Dim i As Long

    If col Is Nothing Then
        Set col = New Collection
        Set tempCol = New Collection
    End If

    If TypeName(dataStructure) = "Dictionary" Then
        tempDataStructure = dataStructure.Items
    Else
        If IsObject(dataStructure) Then
            Set tempDataStructure = dataStructure
        Else
            tempDataStructure = dataStructure
        End If
    End If

    For Each elem In tempDataStructure
        If isDataStructure(elem) Then
            tempCol.Add elem
        Else
            col.Add elem
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
                col.Add tempCol(i)
                tempCol.Remove i
                Exit For
            End If
        Next i
    Loop

    Set getNestedElementsIter = col
End Function

Public Function GreaterThan(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThan = (origVal > newVal)
            
    'Call addFluentPath("GreaterThan", CBool(GreaterThan), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "GreaterThan"
    pExprDict.Add "result", CBool(GreaterThan)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
    
End Function

Public Function GreaterThanOrEqualTo(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (origVal >= newVal)
    
    'Call addFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "GreaterThanOrEqualTo"
    pExprDict.Add "result", CBool(GreaterThanOrEqualTo)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function LessThan(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (origVal < newVal)
    
    'Call addFluentPath("LessThan", CBool(LessThan), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "LessThan"
    pExprDict.Add "result", CBool(LessThan)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)

End Function

Public Function LessThanOrEqualTo(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (origVal <= newVal)
    
    'Call addFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "LessThanOrEqualTo"
    pExprDict.Add "result", CBool(LessThanOrEqualTo)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function EqualTo(origVal As Variant, newVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim tempBool As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    tempBool = (origVal = newVal)
    
    If pMeta.ApproximateEqual Then
        If TypeName(origVal) <> TypeName(newVal) Then
            If IsNumeric(origVal) And IsNumeric(newVal) Then
                tempBool = (Abs(origVal - newVal) <= pMeta.Epsilon)
            ElseIf VarType(origVal) = vbString Or VarType(newVal) = vbString Then
                tempBool = (StrComp(CStr(origVal), CStr(newVal), vbTextCompare) = 0)
            End If
        End If
    End If
    
    EqualTo = tempBool
    
    'Call addFluentPath("EqualTo", CBool(EqualTo), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "EqualTo"
    pExprDict.Add "result", CBool(EqualTo)
    pExprDict.Add "negateValue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function Contain(origVal As Variant, newVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If origVal Like "*" & newVal & "*" Then
        Contain = True
    End If
    
    'Call addFluentPath("Contain", CBool(Contain), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "Contain"
    pExprDict.Add "result", CBool(Contain)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function StartWith(origVal As Variant, newVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(newVal)
    
    If Left(origVal, valLength) = CStr(newVal) Then
        StartWith = True
    End If
    
    'Call addFluentPath("StartWith", CBool(StartWith), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "StartWith"
    pExprDict.Add "result", CBool(StartWith)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function EndWith(origVal As Variant, newVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    valLength = Len(newVal)
    
    If Right(origVal, valLength) = CStr(newVal) Then
        EndWith = True
    End If
    
    'Call addFluentPath("EndWith", CBool(EndWith), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "EndWith"
    pExprDict.Add "result", CBool(EndWith)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
    
End Function

Public Function LengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(origVal)) = newVal)
    
    'Call addFluentPath("LengthOf", CBool(LengthOf), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "LengthOf"
    pExprDict.Add "result", CBool(LengthOf)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
            
End Function

Public Function MaxLengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(origVal)) <= newVal)
    
    'Call addFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "MaxLengthOf"
    pExprDict.Add "result", CBool(MaxLengthOf)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
            
End Function

Public Function MinLengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(origVal)) >= newVal)
    
    'Call addFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "MinLengthOf"
    pExprDict.Add "result", CBool(MinLengthOf)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function Something(origVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not origVal Is Nothing)
    
    'Call addFluentPath("Something", CBool(Something), CBool(negateValue), origVal)
    
    pExprDict.Add "functionName", "Something"
    pExprDict.Add "result", CBool(Something)
    pExprDict.Add "negatevalue", CBool(negateValue)
    'pExprDict.Add "expected", ""
    pExprDict.Add "expected", TypeName(origVal)
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function Between(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(origVal, lowerVal) And LessThanOrEqualTo(origVal, higherVal))
    
    'Call addFluentPath("Between", CBool(Between), CBool(negateValue), origVal:=origVal, lowerVal:=lowerVal, higherVal:=higherVal)
    
    pExprDict.Add "functionName", "Between"
    pExprDict.Add "result", CBool(Between)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "lowerVal", lowerVal
    pExprDict.Add "higherVal", higherVal
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function OneOf(origVal As Variant, negateValue As Boolean, ParamArray args() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    Dim expectedStr As String
        
    arr = args(0)
    tempBool = False
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            expectedStr = expectedStr & TypeName(arr(i)) & ", "
        Else
            expectedStr = expectedStr & arr(i) & ", "
        End If
    
        If IsObject(arr(i)) Or IsObject(origVal) Then
            If TypeName(arr(i)) = TypeName(origVal) Then
                tempBool = True
            End If
        ElseIf arr(i) = origVal Then
            tempBool = True
            'Exit For
        End If
    Next i
    
    expectedStr = Left(expectedStr, Len(expectedStr) - 2)
    
    OneOf = tempBool
    
    'Call addFluentPath("OneOf", CBool(OneOf), CBool(negateValue))
    
    pExprDict.Add "functionName", "OneOf"
    pExprDict.Add "result", CBool(OneOf)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", expectedStr
    
    Call addFluentPathNew(pExprDict)
        
End Function

Public Function LengthBetween(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(origVal))
    
    LengthBetween = (MinLengthOf(tempOrigVal, lowerVal) And MaxLengthOf(tempOrigVal, higherVal))
    
    'Call addFluentPath("LengthBetween", CBool(LengthBetween), CBool(negateValue), origVal:=origVal, lowerVal:=lowerVal, higherVal:=higherVal)
    
    pExprDict.Add "functionName", "LengthBetween"
    pExprDict.Add "result", CBool(LengthBetween)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "lowerVal", lowerVal
    pExprDict.Add "higherVal", higherVal
    
    Call addFluentPathNew(pExprDict)

End Function

Public Function EvaluateTo(origVal As Variant, newVal As Variant, Optional negateValue = False) As Variant
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(origVal) = newVal)
    
    'Call addFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(negateValue), origVal, newVal)
    
    pExprDict.Add "functionName", "EvaluateTo"
    pExprDict.Add "result", CBool(EvaluateTo)
    pExprDict.Add "negatevalue", CBool(negateValue)
    pExprDict.Add "expected", origVal
    pExprDict.Add "actual", newVal
    
    Call addFluentPathNew(pExprDict)
    
End Function

'Private Sub addFluentPath(functionName As String, functionBool As Boolean, negateValue As Boolean, _
'Optional origVal As Variant, Optional newVal As Variant, Optional lowerVal As Variant, Optional higherVal As Variant)
'    Dim fluentPath As String
'    Dim shouldStr As String
'    Dim fluentBool As Boolean
'    Dim fluentTestName As String
'    Dim fluentCategory As String
'
'    If negateValue Then
'        fluentBool = Not functionBool
'        shouldStr = "should not"
'    Else
'        fluentBool = functionBool
'        shouldStr = "should"
'    End If
'
'    fluentPath = ""
'
'    If pPrinting.Category <> "" Then fluentPath = fluentPath & pPrinting.Category & "="
'
'    fluentPath = fluentPath & CLng(fluentBool) & "="
'
'    If pPrinting.TestName <> "" Then fluentPath = fluentPath & pPrinting.TestName & " "
'
'    Select Case functionName
'        Case "InDataStructure"
'            fluentPath = fluentPath & shouldStr & " be in datastructure. "
'        Case "GreaterThan"
'            fluentPath = fluentPath & shouldStr & " be greater than "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "GreaterThanOrEqualTo"
'            fluentPath = fluentPath & shouldStr & " be greater than or equal to "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "LessThan"
'            fluentPath = fluentPath & shouldStr & " be less than "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "LessThanOrEqualTo"
'            fluentPath = fluentPath & shouldStr & " be less than or equal to "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "EqualTo"
'            fluentPath = fluentPath & shouldStr & " be equal to "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "Contain"
'            fluentPath = fluentPath & shouldStr & " contain "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "StartWith"
'            fluentPath = fluentPath & shouldStr & " start with "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "EndWith"
'            fluentPath = fluentPath & shouldStr & " end with "
'            fluentPath = fluentPath & newVal & ". Actual: " & origVal
'        Case "LengthOf"
'            fluentPath = fluentPath & shouldStr & " have length of "
'            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
'        Case "MaxLengthOf"
'            fluentPath = fluentPath & shouldStr & " have max length of "
'            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
'        Case "MinLengthOf"
'            fluentPath = fluentPath & shouldStr & " have min length of "
'            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
'        Case "Something"
'            fluentPath = fluentPath & shouldStr & " be something. "
'        Case "Between"
'            fluentPath = fluentPath & shouldStr & " be between "
'            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & origVal
'        Case "OneOf"
'            fluentPath = fluentPath & shouldStr & " be one of "
'        Case "LengthBetween"
'            fluentPath = fluentPath & shouldStr & " have length between "
'            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & Len(origVal)
'        Case "EvaluateTo"
'            fluentPath = fluentPath & shouldStr & " evaluate to "
'            fluentPath = fluentPath & CStr(newVal) & ". Actual: " & origVal
'    End Select
'
'    pFluentPath = fluentPath
'
'    iPrint.TestCount = iPrint.TestCount + 1
'
'    iPrint.TestCollection.Add fluentPath
'
'    Call pTestResult.CheckTest(fluentBool, pExprDict)
'End Sub

Private Sub addFluentPathNew(dict As Object)
    Dim fluentPath As String
    Dim shouldStr As String
    Dim fluentBool As Boolean
    Dim fluentTestName As String
    Dim fluentCategory As String
    Dim newVal As Variant
    Dim origVal As Variant
    Dim lowerVal As Variant
    Dim higherVal As Variant
    Dim functionName As String
    
    functionName = dict("functionName")
    
    If dict("negateValue") Then
        fluentBool = Not dict("result")
        shouldStr = "should not"
    Else
        fluentBool = dict("result")
        shouldStr = "should"
    End If
    
    fluentPath = ""
    
    If pPrinting.Category <> "" Then fluentPath = fluentPath & pPrinting.Category & "="
    
    fluentPath = fluentPath & CLng(fluentBool) & "="
    
    If pPrinting.TestName <> "" Then fluentPath = fluentPath & pPrinting.TestName & " "
    
    origVal = dict("expected")
    
    If Not dict.exists("actual") Then
        dict.Add "actual", ""
    End If
    
    newVal = dict("actual")
        
    Select Case functionName
        Case "Between"
            lowerVal = dict("lowerVal")
            higherVal = dict("higherVal")
            fluentPath = fluentPath & shouldStr & " be between "
            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & origVal
        Case "Contain"
            fluentPath = fluentPath & shouldStr & " contain "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "EqualTo"
            fluentPath = fluentPath & shouldStr & " be equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "EndWith"
            fluentPath = fluentPath & shouldStr & " end with "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "EvaluateTo"
            fluentPath = fluentPath & shouldStr & " evaluate to "
            fluentPath = fluentPath & CStr(newVal) & ". Actual: " & origVal
        Case "GreaterThan"
            fluentPath = fluentPath & shouldStr & " be greater than "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "GreaterThanOrEqualTo"
            fluentPath = fluentPath & shouldStr & " be greater than or equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "InDataStructure"
            fluentPath = fluentPath & shouldStr & " be in datastructure. "
        Case "LengthBetween"
            lowerVal = dict("lowerVal")
            higherVal = dict("higherVal")
            fluentPath = fluentPath & shouldStr & " have length between "
            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & Len(origVal)
        Case "LengthOf"
            fluentPath = fluentPath & shouldStr & " have length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "LessThan"
            fluentPath = fluentPath & shouldStr & " be less than "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "LessThanOrEqualTo"
            fluentPath = fluentPath & shouldStr & " be less than or equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "MaxLengthOf"
            fluentPath = fluentPath & shouldStr & " have max length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "MinLengthOf"
            fluentPath = fluentPath & shouldStr & " have min length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "OneOf"
            fluentPath = fluentPath & shouldStr & " be one of: " & dict("expected")
        Case "Something"
            fluentPath = fluentPath & shouldStr & " be something: " & dict("expected")
        Case "StartWith"
            fluentPath = fluentPath & shouldStr & " start with "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
    End Select
    
    'pFluentPath = fluentPath
    
    iPrint.TestCount = iPrint.TestCount + 1
    
    iPrint.TestCollection.Add fluentPath
    
    pExprDict.Add "fluentPath", Split(fluentPath, "=")(1)
    
    Call pTestResult.CheckTest(fluentBool, pExprDict)
    
    Set pExprDict = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Initialize()
    Set pExprDict = CreateObject("Scripting.Dictionary")
End Sub
