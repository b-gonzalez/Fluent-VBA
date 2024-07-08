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

Public Property Let CleanTestValStr(value As Boolean)
End Property

Public Property Get CleanTestValStr() As Boolean
End Property

Public Property Let CleanTestInputStr(value As Boolean)
End Property

Public Property Get CleanTestInputStr() As Boolean
End Property

Public Property Let CleanTestStrings(value As Boolean)
End Property

Public Property Get CleanTestStrings() As Boolean
End Property

Public Property Get CleanStringDict() As Scripting.Dictionary
End Property

Public Sub AddToCleanStringDict(key As String, Optional value As String = "", Optional replaceIfExists As Boolean = False)
End Sub

Public Function CleanString(TestStr As Variant, Optional CleanDict As Scripting.Dictionary = Nothing) As String
End Function
