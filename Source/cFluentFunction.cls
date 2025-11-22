VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFluentFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFluentFunction

Private pUnit As IUnitFluent
Private pMeta As IMeta
Private pTestingFunctions As cTestingFunctions
Private pFluentOfVal As Variant

Public Function OfCalledFunction( _
    ByVal functionName As String, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant _
) As IUnitFluentOf
    Set OfCalledFunction = IFluentFunction_OfCalledFunction(functionName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
End Function

Public Function OfCalledMethod( _
    methodObject As Object, _
    methodName As String, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant _
) As IUnitFluentOf
    Set OfCalledMethod = IFluentFunction_OfCalledMethod(methodObject, methodName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
End Function

Public Property Get Meta() As IMeta
    Set Meta = IFluentFunction_Meta
End Property

Private Property Get IFluentFunction_Meta() As IMeta
    Set IFluentFunction_Meta = pMeta
End Property

Private Function IFluentFunction_OfCalledFunction( _
    ByVal functionName As String, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant _
) As IUnitFluentOf
    Dim tf As ITestingFunctions
    Dim tempResult As Variant
    Dim All As Variant
    Dim errNum As Long
    
    
    If TypeName(testingInput1) = "ErrObject" Then
            Set tempResult = Application.run(functionName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
    Else
        On Error Resume Next
            If IsObject(testingInput1) Then
                Set tempResult = Application.run(functionName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            Else
                tempResult = Application.run(functionName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            End If
            
            errNum = Err.Number
        On Error GoTo 0
    End If
    
    If errNum <> 1004 Then
        If IsObject(tempResult) Then
            Set All = tempResult
        Else
            All = tempResult
        End If
    Else
        Debug.Print "Error: " & Err.Description
    End If
    
    If Not VBA.Information.IsMissing(All) Or VBA.Information.IsArray(All) Then
        If VBA.Information.IsObject(All) Then
            Set pFluentOfVal = All
            Set pUnit.testValue = pFluentOfVal
        Else
            pFluentOfVal = All
            pUnit.testValue = pFluentOfVal
        End If
    Else
        Set tf = pUnit.TestingFunctions
        tf.TestValueSet = False
    End If
    
    Set IFluentFunction_OfCalledFunction = pUnit
End Function

Private Function IFluentFunction_OfCalledMethod( _
    methodObject As Object, _
    methodName As String, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant _
) As IUnitFluentOf
    Dim tf As ITestingFunctions
    Dim tempResult As Variant
    Dim All As Variant
    
    On Error Resume Next
        tempResult = callByNameWithOptionalArgs(methodObject, methodName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
        
        If Err.Number = 0 Then
            All = tempResult
        End If
    On Error GoTo 0
    
    If Not VBA.Information.IsMissing(All) Or VBA.Information.IsArray(All) Then
        If VBA.Information.IsObject(All) Then
            Set pFluentOfVal = All
            Set pUnit.testValue = pFluentOfVal
        Else
            pFluentOfVal = All
            pUnit.testValue = pFluentOfVal
        End If
    Else
        Set tf = pUnit.TestingFunctions
        tf.TestValueSet = False
    End If
    
    Set IFluentFunction_OfCalledMethod = pUnit
End Function

Private Function callByNameWithOptionalArgs( _
    methodObj As Object, _
    methodName As String, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant _
)
    Dim temp As Variant
    Dim elem As Variant
    Dim tempArr As Variant
    
    temp = Null
    
    Select Case True
        Case Not VBA.Information.IsMissing(testingInput30)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            End If
        Case Not VBA.Information.IsMissing(testingInput29)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29)
            End If
        Case Not VBA.Information.IsMissing(testingInput28)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28)
            End If
        Case Not VBA.Information.IsMissing(testingInput27)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27)
            End If
        Case Not VBA.Information.IsMissing(testingInput26)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26)
            End If
        Case Not VBA.Information.IsMissing(testingInput25)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25)
            End If
        Case Not VBA.Information.IsMissing(testingInput24)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24)
            End If
        Case Not VBA.Information.IsMissing(testingInput23)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23)
            End If
        Case Not VBA.Information.IsMissing(testingInput22)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22)
            End If
        Case Not VBA.Information.IsMissing(testingInput21)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21)
            End If
        Case Not VBA.Information.IsMissing(testingInput20)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20)
            End If
        Case Not VBA.Information.IsMissing(testingInput19)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19)
            End If
        Case Not VBA.Information.IsMissing(testingInput18)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18)
            End If
        Case Not VBA.Information.IsMissing(testingInput17)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17)
            End If
        Case Not VBA.Information.IsMissing(testingInput16)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16)
            End If
        Case Not VBA.Information.IsMissing(testingInput15)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15)
            End If
        Case Not VBA.Information.IsMissing(testingInput14)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14)
            End If
        Case Not VBA.Information.IsMissing(testingInput13)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13)
            End If
        Case Not VBA.Information.IsMissing(testingInput12)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12)
            End If
        Case Not VBA.Information.IsMissing(testingInput11)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11)
            End If
        Case Not VBA.Information.IsMissing(testingInput10)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10)
            End If
        Case Not VBA.Information.IsMissing(testingInput9)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9)
            End If
        Case Not VBA.Information.IsMissing(testingInput8)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8)
            End If
        Case Not VBA.Information.IsMissing(testingInput7)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7)
            End If
        Case Not VBA.Information.IsMissing(testingInput6)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6)
            End If
        Case Not VBA.Information.IsMissing(testingInput5)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5)
            End If
        Case Not VBA.Information.IsMissing(testingInput4)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3, testingInput4)
            End If
        Case Not VBA.Information.IsMissing(testingInput3)
            tempArr = Array(testingInput1, testingInput2, testingInput3)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2, testingInput3)
            End If
        Case Not VBA.Information.IsMissing(testingInput2)
            tempArr = Array(testingInput1, testingInput2)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1, testingInput2)
            End If
        Case Not VBA.Information.IsMissing(testingInput1)
            tempArr = Array(testingInput1)
        
            If Not elementsAreMissing(tempArr) Then
                temp = VBA.Interaction.CallByName(methodObj, methodName, VbMethod, testingInput1)
            End If
    End Select
    
    callByNameWithOptionalArgs = temp
End Function

Private Function elementsAreMissing(arr As Variant)
    Dim counter As Long
    Dim elem As Variant
    
    For Each elem In arr
        If Not VBA.Information.IsMissing(elem) Then
            counter = counter + 1
        End If
    Next elem
    
    elementsAreMissing = (counter <> UBound(arr) + 1)
End Function

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pTestingFunctions = pUnit.TestingFunctions
    Set pMeta = pUnit.Meta
End Sub

Private Sub Class_Terminate()
    Set pUnit = Nothing
    Set pTestingFunctions = Nothing
    Set pMeta = Nothing
End Sub
