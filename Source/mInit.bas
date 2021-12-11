Attribute VB_Name = "mInit"
Option Explicit

Public Function MakeFluent() As IFluent
    Set MakeFluent = New cFluent
End Function

Public Function MakeFluentOf() As IFluentOf
    Set MakeFluent = New cFluent
End Function
