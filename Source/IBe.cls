VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Alphabetic() As Boolean

End Function

Public Function Alphanumeric() As Boolean

End Function

Public Function Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Boolean

End Function

Public Function Erroneous() As Boolean

End Function

Public Function EqualTo(ByVal value As Variant) As Boolean

End Function

Public Function GreaterThan(ByVal value As Double) As Boolean

End Function

Public Function GreaterThanOrEqualTo(ByVal value As Double) As Boolean

End Function

Public Function InDataStructure(ByVal DataStructure As Variant, Optional ByVal Method As flMethod = flRecursive) As Boolean

End Function

Public Function InDataStructures(ParamArray DataStructures() As Variant) As Boolean

End Function

Public Function LessThan(ByVal value As Double) As Boolean

End Function

Public Function LessThanOrEqualTo(ByVal value As Double) As Boolean

End Function

Public Function Numeric() As Boolean

End Function

Public Function OneOf(ParamArray Vals() As Variant)

End Function

Public Function Something() As Boolean

End Function
