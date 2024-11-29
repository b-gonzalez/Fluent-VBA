VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let CleanTestValueStr(ByVal value As Boolean)
End Property

Public Property Get CleanTestValueStr() As Boolean
End Property

Public Property Let CleanTestInputStr(ByVal value As Boolean)
End Property

Public Property Get CleanTestInputStr() As Boolean
End Property

Public Property Let CleanTestStrings(ByVal value As Boolean)
End Property

Public Property Get CleanTestStrings() As Boolean
End Property

Public Property Get CleanStringDict() As Scripting.Dictionary
End Property

Public Sub AddToCleanStringDict(ByVal key As String, Optional ByVal value As String = "", Optional ByVal replaceIfExists As Boolean = False)
End Sub

Public Function CleanString(ByVal TestStr As Variant, Optional ByVal CleanDict As Scripting.Dictionary = Nothing) As String
End Function
