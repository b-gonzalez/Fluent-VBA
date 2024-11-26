VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cToString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IToString

Private pWrapChar As Long

Private Const DEFAULT_WRAP_CHAR As String = "`"

Private Property Let IToString_WrapChar(value As String)
    pWrapChar = VBA.Asc(value)
End Property

Private Property Get IToString_WrapChar() As String
    IToString_WrapChar = VBA.Chr(pWrapChar)
End Property

Private Function IToString_WrapIfString(value As Variant) As Variant
    Dim tempVal As Variant
    
    tempVal = value
    
    If TypeName(tempVal) = "String" Then
'        If Left(tempVal, 1) <> pWrapChar And Right(tempVal, 1) <> pWrapChar Then
        If VBA.Left(tempVal, 1) <> pWrapChar And VBA.Right(tempVal, 1) <> pWrapChar Then
            tempVal = VBA.Chr(pWrapChar) & tempVal & VBA.Chr(pWrapChar)
        End If
    End If
    
    IToString_WrapIfString = tempVal
End Function

Private Sub Class_Initialize()
    pWrapChar = VBA.Asc(DEFAULT_WRAP_CHAR)
End Sub
