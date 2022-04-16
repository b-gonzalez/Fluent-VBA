VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShould"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Be() As IBe

End Property

Public Property Get Have() As IHave

End Property


Public Function Contain(value As Variant) As Boolean

End Function

Public Function StartWith(value As Variant) As Boolean

End Function

Public Function EndWith(value As Variant) As Boolean

End Function

Public Function EvaluateTo(value As Variant) As Boolean

End Function
