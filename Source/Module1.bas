Attribute VB_Name = "Module1"
Option Explicit

Sub subby()
    Dim f As IFluent
    
    Set f = New cFluent
    
    f.Meta.ApproximateEqual = True
    f.TestValue = 5.0000001
    With f.Should.Be
        Debug.Assert .EqualTo(5)
    End With
End Sub
