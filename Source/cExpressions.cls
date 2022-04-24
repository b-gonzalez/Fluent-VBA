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

Private pTestValue As Variant
Private pMeta As cMeta
Private pPrintSettings As IPrintSettings

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

Public Property Set setPrintSettings(value As IPrintSettings)
    Set pPrintSettings = value
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
    
    Call addFluentPath("InDataStructure", CBool(InDataStructure), CBool(negateValue), origVal)
    
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
            
    Call addFluentPath("GreaterThan", CBool(GreaterThan), CBool(negateValue), origVal, newVal)
    
End Function

Public Function GreaterThanOrEqualTo(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    GreaterThanOrEqualTo = (origVal >= newVal)
    
    Call addFluentPath("GreaterThanOrEqualTo", CBool(GreaterThanOrEqualTo), CBool(negateValue), origVal, newVal)
        
End Function

Public Function LessThan(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LessThan = (origVal < newVal)
    
    Call addFluentPath("LessThan", CBool(LessThan), CBool(negateValue), origVal, newVal)

End Function

Public Function LessThanOrEqualTo(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
       
    LessThanOrEqualTo = (origVal <= newVal)
    
    Call addFluentPath("LessThanOrEqualTo", CBool(LessThanOrEqualTo), CBool(negateValue), origVal, newVal)
        
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
    
    Call addFluentPath("EqualTo", CBool(EqualTo), CBool(negateValue), origVal, newVal)
        
End Function

Public Function Contain(origVal As Variant, newVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    If origVal Like "*" & newVal & "*" Then
        Contain = True
    End If
    
    Call addFluentPath("Contain", CBool(Contain), CBool(negateValue), origVal, newVal)
        
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
    
    Call addFluentPath("StartWith", CBool(StartWith), CBool(negateValue), origVal, newVal)
        
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
    
    Call addFluentPath("EndWith", CBool(EndWith), CBool(negateValue), origVal, newVal)
    
End Function

Public Function LengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    LengthOf = (Len(CStr(origVal)) = newVal)
    
    Call addFluentPath("LengthOf", CBool(LengthOf), CBool(negateValue), origVal, newVal)
            
End Function

Public Function MaxLengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MaxLengthOf = (Len(CStr(origVal)) <= newVal)
    
    Call addFluentPath("MaxLengthOf", CBool(MaxLengthOf), CBool(negateValue), origVal, newVal)
            
End Function

Public Function MinLengthOf(origVal As Double, newVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    MinLengthOf = (Len(CStr(origVal)) >= newVal)
    
    Call addFluentPath("MinLengthOf", CBool(MinLengthOf), CBool(negateValue), origVal, newVal)
        
End Function

Public Function Something(origVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Something = (Not origVal Is Nothing)
    
    Call addFluentPath("Something", CBool(Something), CBool(negateValue), origVal)
        
End Function

Public Function Between(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    Between = (GreaterThanOrEqualTo(origVal, lowerVal) And LessThanOrEqualTo(origVal, higherVal))
    
    Call addFluentPath("Between", CBool(Between), CBool(negateValue), origVal:=origVal, lowerVal:=lowerVal, higherVal:=higherVal)
        
End Function

Public Function OneOf(origVal As Variant, negateValue As Boolean, ParamArray args() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    arr = args(0)
    tempBool = False
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            If TypeName(arr(i)) = TypeName(origVal) Then
                tempBool = True
            End If
        ElseIf arr(i) = origVal Then
            tempBool = True
            Exit For
        End If
    Next i
    
    OneOf = tempBool
    
    Call addFluentPath("OneOf", CBool(OneOf), CBool(negateValue))
        
End Function

Public Function LengthBetween(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim tempOrigVal As Double
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
        
    tempOrigVal = Len(CStr(origVal))
    
    LengthBetween = (MinLengthOf(tempOrigVal, lowerVal) And MaxLengthOf(tempOrigVal, higherVal))
    
    Call addFluentPath("LengthBetween", CBool(LengthBetween), CBool(negateValue), origVal:=origVal, lowerVal:=lowerVal, higherVal:=higherVal)

End Function

Public Function EvaluateTo(origVal As Variant, newVal As Variant, Optional negateValue = False) As Variant
    Dim fluentPath As String
    Dim fluentBool As Boolean
    Dim shouldStr As String
    
    EvaluateTo = (Evaluate(origVal) = newVal)
    
    Call addFluentPath("EvaluateTo", CBool(EvaluateTo), CBool(negateValue), origVal, newVal)
    
End Function

Private Sub addFluentPath(functionName As String, functionBool As Boolean, negateValue As Boolean, _
Optional origVal As Variant, Optional newVal As Variant, Optional lowerVal As Variant, Optional higherVal As Variant)
    Dim fluentPath As String
    Dim shouldStr As String
    Dim fluentBool As Boolean
    Dim fluentTestName As String
    Dim fluentCategory As String
    
    If negateValue Then
        fluentBool = Not functionBool
        shouldStr = "should not"
    Else
        fluentBool = functionBool
        shouldStr = "should"
    End If
    
    fluentPath = ""
    
    If pPrintSettings.Category <> "" Then fluentPath = fluentPath & pPrintSettings.Category & "="
    
    fluentPath = fluentPath & CLng(fluentBool) & "="
    
    If pPrintSettings.testName <> "" Then fluentPath = fluentPath & pPrintSettings.testName & " "
        
    Select Case functionName
        Case "InDataStructure"
            fluentPath = fluentPath & shouldStr & " be in datastructure. "
        Case "GreaterThan"
            fluentPath = fluentPath & shouldStr & " be greater than "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "GreaterThanOrEqualTo"
            fluentPath = fluentPath & shouldStr & " be greater than or equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "LessThan"
            fluentPath = fluentPath & shouldStr & " be less than "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "LessThanOrEqualTo"
            fluentPath = fluentPath & shouldStr & " be less than or equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "EqualTo"
            fluentPath = fluentPath & shouldStr & " be equal to "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "Contain"
            fluentPath = fluentPath & shouldStr & " contain "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "StartWith"
            fluentPath = fluentPath & shouldStr & " start with "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "EndWith"
            fluentPath = fluentPath & shouldStr & " end with "
            fluentPath = fluentPath & newVal & ". Actual: " & origVal
        Case "LengthOf"
            fluentPath = fluentPath & shouldStr & " have length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "MaxLengthOf"
            fluentPath = fluentPath & shouldStr & " have max length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "MinLengthOf"
            fluentPath = fluentPath & shouldStr & " have min length of "
            fluentPath = fluentPath & newVal & ". Actual: " & Len(origVal)
        Case "Something"
            fluentPath = fluentPath & shouldStr & " be something. "
        Case "Between"
            fluentPath = fluentPath & shouldStr & " be between "
            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & origVal
        Case "OneOf"
            fluentPath = fluentPath & shouldStr & " be one of "
        Case "LengthBetween"
            fluentPath = fluentPath & shouldStr & " have length between "
            fluentPath = fluentPath & lowerVal & " and " & higherVal & ". Actual: " & Len(origVal)
        Case "EvaluateTo"
            fluentPath = fluentPath & shouldStr & " evaluate to "
            fluentPath = fluentPath & CStr(newVal) & ". Actual: " & origVal
    End Select
    
    pPrintSettings.TestCollection.Add fluentPath
End Sub
