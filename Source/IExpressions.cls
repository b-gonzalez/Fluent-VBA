VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Set setPrinting(ByVal value As cPrinting)
End Property

Public Property Set setTests(ByVal value As ITests)
End Property

Public Property Set setMeta(ByVal value As IMeta)
End Property

Public Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal result As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant, _
    Optional ByVal cleanedTestValue As Variant, _
    Optional ByVal cleanedTestInput As Variant, _
    Optional ByVal Actual As Variant, _
    Optional ByVal recurIterFunc As Variant, _
    Optional ByVal TestingValueIsSelfReferential As Variant, _
    Optional ByVal TestingInputIsSelfReferential As Variant _
)
End Sub

Public Function Should(ByVal shouldType As Long, Optional ByVal testingInput As Variant) As String
End Function

Public Function Be( _
    ByVal beType As Long, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant) As String
End Function

Public Function Have( _
    ByVal haveType As Long, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant) As String
End Function

Public Function buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
End Function

Public Function getFunctionExpr(ByVal test As ITest) As String
End Function

Public Function getFluentPath( _
    ByVal FluentPath As String, _
    ByVal test As ITest, _
    Optional ByVal notShould As Boolean = False) As String
End Function

Public Sub addFluentPathNew(ByVal test As ITest, ByVal negateValue As Boolean)
End Sub

Public Function InputToString(ByVal nv As Variant, Optional ByVal algo As Variant, Optional ByVal indent As Boolean = False) As String
End Function

Public Function DatastructureIsEmpty(ByVal v As Variant) As Boolean
End Function

Public Function InputToStringRecur( _
    ByVal v As Variant, _
    Optional ByVal col As Collection = Nothing, _
    Optional ByVal counter As Long = 1, _
    Optional ByVal indent As Boolean = False _
) As Collection
End Function

Public Function getResultNumber(ByVal result As Variant) As Long
End Function

Public Function inputToStringIter(ByVal v As Variant, Optional ByVal indent As Boolean = False) As Collection
End Function

Public Function getDsCount(ByVal dataStructure As Variant) As Long
End Function
