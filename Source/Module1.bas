Attribute VB_Name = "Module1"
Option Explicit

Sub subby()
    Dim fluent As cFluent
    Dim testFluent As cFluentOf
    Dim arr As Variant
    
    Set fluent = New cFluent
    Set testFluent = New cFluentOf

    arr = Array(9, Array(10, Array(11)))
    With testFluent.Of(10).Should.Be
        Debug.Assert .InDataStructure(arr, flIterative)
    End With

End Sub
