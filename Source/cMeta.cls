VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IMeta
Implements IMetaFluentFunction

Private pPrinting As IPrinting
Private pPrintExpr As IPrintExpr
Private pPrintingFF As IPrintingFluentFunction
Private pTests As ITests
Private pUtilities As IUtilities
Private pToString As IToString
Private pFunctionDict As Scripting.Dictionary
Private pPathDict As Scripting.Dictionary
Private pFunctionOrMethodPath As String

Private Property Get IMeta_Printing() As IPrinting
    Set IMeta_Printing = pPrinting
End Property

Private Property Get IMeta_Tests() As ITests
    Set IMeta_Tests = pTests
End Property

Private Property Get IMeta_Utilities() As IUtilities
    Set IMeta_Utilities = pUtilities
End Property

Private Property Get IMeta_ToString() As IToString
    Set IMeta_ToString = pToString
End Property

Private Property Get IMetaFluentFunction_Printing() As IPrintingFluentFunction
    Set IMetaFluentFunction_Printing = pPrinting
End Property

Private Property Get IMetaFluentFunction_Utilities() As IUtilities
    Set IMetaFluentFunction_Utilities = pUtilities
End Property

Private Property Get IMetaFluentFunction_tests() As ITests
    Set IMetaFluentFunction_tests = pTests
End Property

Private Property Get IMetaFluentFunction_ToString() As IToString
    Set IMetaFluentFunction_ToString = pToString
End Property

Private Property Let IMetaFluentFunction_FunctionOrMethodPath(value As String)
    pFunctionOrMethodPath = value
End Property

Private Property Get IMetaFluentFunction_FunctionOrMethodPath() As String
    IMetaFluentFunction_FunctionOrMethodPath = pFunctionOrMethodPath
End Property

Private Sub IMetaFluentFunction_addFunction(functionName As String, Optional FunctionOrMethodPath As String = "", Optional increment As Long = 0)
    Dim tempPath As String
    
    If FunctionOrMethodPath = "" And pFunctionOrMethodPath = "" Then
        tempPath = "Root"
    Else
        If FunctionOrMethodPath <> "" Then
            tempPath = FunctionOrMethodPath
        ElseIf pFunctionOrMethodPath <> "" Then
            tempPath = pFunctionOrMethodPath
        End If
    End If
    
    If pTests.Algorithm = flAlgorithm.flRecursive Then
        Set pPathDict = getApiDictRecur(functionName, tempPath, increment, pPathDict)
    ElseIf pTests.Algorithm = flAlgorithm.flIterative Then
        Set pPathDict = getApiDictIter(functionName, tempPath, increment, pPathDict)
    End If
End Sub

Private Property Get IMetaFluentFunction_FunctionDict() As Scripting.Dictionary
    Set IMetaFluentFunction_FunctionDict = pFunctionDict
End Property

Private Function getApiDictRecur( _
    functionName As String, _
    Path As String, _
    increment As Long, _
    Optional PathDict As Scripting.Dictionary = Nothing, _
    Optional counter As Long = 0 _
) As Scripting.Dictionary
    Dim temp As Variant
    Dim i As Long
    Dim tempDict As Scripting.Dictionary
    Dim finalDict As Scripting.Dictionary
    Dim funcDictPath As String
    
    temp = Split(Path, ".")
    
    If PathDict Is Nothing Then
        Set PathDict = New Scripting.Dictionary
    End If
    
    If Not PathDict.Exists(temp(counter)) Then
        Set tempDict = New Scripting.Dictionary
    Else
        Set tempDict = PathDict(temp(counter))
    End If
        
    If counter < UBound(temp) Then
        Set PathDict(temp(counter)) = getApiDictRecur(functionName, Path, increment, tempDict, counter + 1)
    Else
        If Not PathDict.Exists(temp(counter)) Then
            Set finalDict = New Scripting.Dictionary
            finalDict.Add functionName, 0
        Else
            Set finalDict = PathDict(temp(counter))
        End If
        
        finalDict(functionName) = finalDict(functionName) + increment
        
        Set PathDict(temp(counter)) = finalDict
        
        funcDictPath = Path & "." & functionName
        
        If Not pFunctionDict.Exists(funcDictPath) Then
            pFunctionDict.Add funcDictPath, 0
        End If
        
        pFunctionDict(funcDictPath) = pFunctionDict(funcDictPath) + increment
    End If
    
    Set getApiDictRecur = PathDict
End Function

Public Function getApiDictIter( _
    functionName As String, _
    Path As String, _
    increment As Long, _
    PathDict As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim temp As Variant
    Dim i As Long
    Dim tempDict As Scripting.Dictionary
'    Dim pathDict As Scripting.Dictionary
    Dim workingDict As Scripting.Dictionary
    Dim funcDictPath As String

    temp = Split(Path, ".")
    Set workingDict = New Scripting.Dictionary
    
'    If pathDict Is Nothing Then
'        Set pathDict = New Scripting.Dictionary
'    End If

    For i = UBound(temp) To LBound(temp) Step -1
        If i = UBound(temp) Then
            Set tempDict = New Scripting.Dictionary
            tempDict.Add functionName, 0
            
            tempDict(functionName) = tempDict(functionName) + increment
            
            If i = LBound(temp) Then
                Set PathDict(temp(i)) = tempDict
            Else
                Set workingDict(temp(i)) = tempDict
            End If
            
            funcDictPath = Path & "." & functionName
            
            If Not pFunctionDict.Exists(funcDictPath) Then
                pFunctionDict.Add funcDictPath, 0
            End If
            
            pFunctionDict(funcDictPath) = pFunctionDict(funcDictPath) + increment
        ElseIf i > LBound(temp) Then
            Set tempDict = workingDict
            Set workingDict = New Scripting.Dictionary
            Set workingDict(temp(i)) = tempDict
        ElseIf i = LBound(temp) Then
            Set PathDict(temp(i)) = workingDict
        End If
    Next i
    
    Set getApiDictIter = PathDict
End Function

Private Sub Class_Initialize()
    Set pPrinting = New cPrinting
    Set pPrintExpr = pPrinting
    Set pPrintingFF = pPrintExpr
    Set pTests = New cTests
    Set pUtilities = New cUtilities
    Set pToString = New cToString
    Set pFunctionDict = New Scripting.Dictionary
    Set pPathDict = New Scripting.Dictionary
    
    Set pPrintingFF.PathDict = pPathDict
    Set pPrintingFF.FunctionDict = pFunctionDict
End Sub

Private Sub Class_Terminate()
    Set pPrinting = Nothing
    Set pPrintExpr = Nothing
    Set pTests = Nothing
    Set pUtilities = Nothing
    Set pToString = Nothing
    Set pFunctionDict = Nothing
    Set pPathDict = Nothing
End Sub
