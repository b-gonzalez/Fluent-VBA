VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LengthOf(ByVal value As Long) As Variant
End Function

Public Function MaxLengthOf(ByVal value As Long) As Variant
End Function

Public Function MinLengthOf(ByVal value As Long) As Variant
End Function

Public Function LengthBetween(ByVal LowValue As Long, ByVal HighValue As Long) As Variant
End Function

Public Function ErrorNumberOf(ByVal value As Long) As Variant
End Function

Public Function ErrorDescriptionOf(ByVal value As String) As Variant
End Function

Public Function SameTypeAs(ByVal All As Variant) As Variant
End Function

Public Function ExactSameElementsAs(ByVal dataStructure As Variant) As Variant
End Function

Public Function SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
End Function

Public Function SameElementsAs(ByVal dataStructure As Variant) As Variant
End Function

Public Function Procedure( _
    ByVal procName As String, _
    ByVal callType As VBA.VbCallType _
) As Variant
End Function

Public Function Elements( _
    ParamArray args() As Variant _
) As Variant
End Function

Public Function ElementsInDataStructure( _
    ByVal dataStructure As Variant _
) As Variant
End Function

Public Function DepthCountOf( _
    ByVal value As Long _
) As Variant
End Function

Public Function NestedCountOf( _
    ByVal value As Long _
) As Variant
End Function

