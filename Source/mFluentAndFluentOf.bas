Attribute VB_Name = "FluentAndFluentOf"
Option Explicit

Sub FluentAndFluentOfExample()
    'NOTE: In order to create Fluent and FluentOf
    'objects, you have to add a reference to the
    'appropriate office file. This is done by:
    '1. Going to Tools > References in the VBIDE.
    '2. Clicking the browse button
    '3. Navigating to the Distibution folder
    '4. Filtering for the office files for the
    'application in the droodown (e.g. Microsoft Excel)
    '5. Selecting the appropriate office file (e.g. Fluent VBA.xlsm)
    
    'NOTE: For PowerPoint, you should add a reference to the ppa file
    'and for Word, you should add a reference to the dotm file.

    Dim fluent As IFluent
    Dim fluentOf As IFluentOf
    
    Set fluent = MakeFluent
    Set fluentOf = MakeFluentOf
    
    fluent.TestValue = fluentOf.Of(True).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    
    Debug.Print "All tests passed!"
End Sub
