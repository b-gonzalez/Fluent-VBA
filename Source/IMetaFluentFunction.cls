VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMetaFluentFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Printing() As IPrintingFluentFunction
End Property

Public Property Get Utilities() As IUtilities
End Property

Public Property Get tests() As ITests
End Property

Public Property Get ToString() As IToString
End Property

Public Sub addFunction(functionName As String, Optional Path As String = "", Optional increment As Long = 0)
End Sub

Public Property Let FunctionOrMethodPath(value As String)
End Property

Public Property Get FunctionOrMethodPath() As String
End Property

Public Property Get FunctionDict() As Scripting.Dictionary
End Property
