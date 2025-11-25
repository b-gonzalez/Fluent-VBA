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
Private pMeta As IMetaFluentFunction
Private pTestingFunctions As cTestingFunctions
Private pFluentOfVal As Variant
Private pPath As String

Const P_VB_METHOD_CALL_TYPE = 1
Const P_VB_GET_CALL_TYPE = 2

Public Property Get Meta() As IMetaFluentFunction
    Set Meta = IFluentFunction_Meta
End Property

Private Property Get IFluentFunction_Meta() As IMetaFluentFunction
    Set IFluentFunction_Meta = pMeta
End Property

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

Public Function OfCalledProcedure( _
    procedureObject As Object, _
    procedureName As String, _
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
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant, _
    Optional ByVal callType As VbCallType = 1 _
) As IUnitFluentOf
    Set OfCalledProcedure = IFluentFunction_OfCalledProcedure(procedureObject, procedureName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30, callType)
End Function

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
    Dim appRunCol As VBA.Collection
    Dim errObj As VBA.ErrObject
    Dim tiErrNum As Long
    Dim tiErrDesc As String
    
    Set appRunCol = New VBA.Collection

    'There is one global error object. If you get to this point and the
    'error number is greater than zero and the error description is not
    'a blank string, then that means that you're dealing with an error
    'that's been provided as a testign input. These values are then stored
    'in variables because they are cleared from the Err object in the error
    'hander in application.run. If there are no additional errors, then
    'they are restored. And the error handling functions are able to run
    'and work as expected.
    
    If Err.Number > 0 And Err.Description <> "" Then
        tiErrNum = Err.Number
        tiErrDesc = Err.Description
    End If
    
    
    On Error Resume Next
        'All 30 testing inputs are provided to Application.run.
        'Application.run automatically filters out missing parameters.
        
        appRunCol.Add Application.run(functionName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
        
        Set errObj = Err
        errNum = Err.Number
    On Error GoTo 0
    
    If errNum = 0 Then
        If IsObject(appRunCol(1)) Then
            Set All = appRunCol(1)
        Else
            All = appRunCol(1)
        End If
        
        If tiErrDesc <> "" And tiErrNum <> 0 Then
            Err.Number = tiErrNum
            Err.Description = tiErrDesc
        End If
        
        Call pMeta.addFunction(functionName, pPath, increment:=1)
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

Private Function IFluentFunction_OfCalledProcedure( _
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
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant, _
    Optional ByVal callType As VbCallType = 1 _
) As IUnitFluentOf
    Dim tf As ITestingFunctions
    Dim tempResult As Variant
    Dim All As Variant
    Dim cbnwoaCol As VBA.Collection
    
    'Integrate the same functionality here for dealing with error objects that I integrated in OfCalledFunction.
    
    Set cbnwoaCol = New Collection
    
    On Error Resume Next
        If callType = P_VB_GET_CALL_TYPE Or callType = P_VB_METHOD_CALL_TYPE Then
            Set cbnwoaCol = callByNameWithOptionalArgs(methodObject, methodName, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30, callType)
        End If
        
        If Err.Number = 0 And cbnwoaCol.Count > 0 Then
            If VBA.Information.IsObject(cbnwoaCol(1)) Then
                Set All = cbnwoaCol(1)
            Else
                All = cbnwoaCol(1)
            End If
            
            Call pMeta.addFunction(methodName, pPath, increment:=1)
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
    
    Set IFluentFunction_OfCalledProcedure = pUnit
End Function

Private Function callByNameWithOptionalArgs( _
    ByVal methodObj As Object, _
    ByVal methodName As String, _
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
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant, _
    Optional ByVal callType As VbCallType = 1 _
) As VBA.Collection
    Dim temp As Variant
    Dim elem As Variant
    Dim tempArr As Variant
    Dim col As VBA.Collection
    
    'This sub uses the ByRef returnedResult as the returned function value. This is done because
    'callByNameWithOptionalArgs does not know the return type in advance. So it does not know
    'whether to use or omit the set keyword. Without using ByRef this way, there'd be no alternative
    'but to call the function twice.
    
    temp = Null
    Set col = New VBA.Collection
    
    Select Case True
        Case Not VBA.Information.IsMissing(testingInput30)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            End If
        Case Not VBA.Information.IsMissing(testingInput29)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29)
            End If
        Case Not VBA.Information.IsMissing(testingInput28)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28)
            End If
        Case Not VBA.Information.IsMissing(testingInput27)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27)
            End If
        Case Not VBA.Information.IsMissing(testingInput26)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26)
            End If
        Case Not VBA.Information.IsMissing(testingInput25)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25)
            End If
        Case Not VBA.Information.IsMissing(testingInput24)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24)
            End If
        Case Not VBA.Information.IsMissing(testingInput23)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23)
            End If
        Case Not VBA.Information.IsMissing(testingInput22)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22)
            End If
        Case Not VBA.Information.IsMissing(testingInput21)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21)
            End If
        Case Not VBA.Information.IsMissing(testingInput20)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20)
            End If
        Case Not VBA.Information.IsMissing(testingInput19)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19)
            End If
        Case Not VBA.Information.IsMissing(testingInput18)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18)
            End If
        Case Not VBA.Information.IsMissing(testingInput17)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17)
            End If
        Case Not VBA.Information.IsMissing(testingInput16)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16)
            End If
        Case Not VBA.Information.IsMissing(testingInput15)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15)
            End If
        Case Not VBA.Information.IsMissing(testingInput14)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14)
            End If
        Case Not VBA.Information.IsMissing(testingInput13)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13)
            End If
        Case Not VBA.Information.IsMissing(testingInput12)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12)
            End If
        Case Not VBA.Information.IsMissing(testingInput11)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11)
            End If
        Case Not VBA.Information.IsMissing(testingInput10)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10)
            End If
        Case Not VBA.Information.IsMissing(testingInput9)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9)
            End If
        Case Not VBA.Information.IsMissing(testingInput8)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8)
            End If
        Case Not VBA.Information.IsMissing(testingInput7)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7)
            End If
        Case Not VBA.Information.IsMissing(testingInput6)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6)
            End If
        Case Not VBA.Information.IsMissing(testingInput5)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4, testingInput5)
            End If
        Case Not VBA.Information.IsMissing(testingInput4)
            tempArr = Array(testingInput1, testingInput2, testingInput3, testingInput4)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3, testingInput4)
            End If
        Case Not VBA.Information.IsMissing(testingInput3)
            tempArr = Array(testingInput1, testingInput2, testingInput3)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2, testingInput3)
            End If
        Case Not VBA.Information.IsMissing(testingInput2)
            tempArr = Array(testingInput1, testingInput2)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1, testingInput2)
            End If
        Case Not VBA.Information.IsMissing(testingInput1)
            tempArr = Array(testingInput1)
        
            If Not elementsAreMissing(tempArr) Then
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType, testingInput1)
            End If
        Case VBA.Information.IsMissing(testingInput1)
                col.Add VBA.Interaction.CallByName(methodObj, methodName, callType)
    End Select
    
    Set callByNameWithOptionalArgs = col
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
