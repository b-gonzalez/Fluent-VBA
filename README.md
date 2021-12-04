# Fluent-VBA
Fluent VBA is a fluent unit testing framework for VBA

Fluent frameworks are intended to be read like natural language. So instead of having something like:

    Sub NormalUnitTestExample
        Dim result as long
        Dim Assert as cUnitTester

        result = returnsFive() ‘returns the number 5
        Set Assert = New cUnitTester

        Assert.Equal(Result,5)
    End Sub
 
You can have code that reads more naturally like so:

    Sub FluentUnitTestExample
        Dim Result as cFluent

        Set Result = new cFluent
        Result.TestValue = ReturnsFive()

        Debug.Assert Result.Should.Be.EqualTo(5)
    End Sub
    
# Testing notes

Most of the tests utilize the IFluent interface. This is because the tests were written before I introduced the new IFluentOf interface (see notes on this interface below)
    
# Meta tests

The fluent unit testing framework uses itself to test itself. These set of tests are contained in the mTests module in the MetaTest sub

# Documentation tests

The mTests module has a DocumentationTests sub that contains several dozen tests. These tests document the various objects and methods in the API.

# Additional tests

Several other tests are implemented documenting the flexibility with which these unit tests can be created. These tests can be found in the mTests module.

# IFluentOf interface

One new big change is the addition of the IFluentOf interface. This new interface allows you to enter the test value in the testing line itself. Using this interface has several advantages: 

1. It removes the need for you to assign the test value using the TestValue property.
2. Writing the test value in the same line as the test can make debugging easier
3. Writing the test value in the test this way can also can read more naturally for certain types of tests.

For you to be able to use this, you need to use the IFluentOf interface for the cFluent object instead of the IFluent interface. You can see an example of the difference between the two interfaces below:

    Sub FluentOfExample()
        Dim Result As IFluentOf
        Dim Result2 As IFluent

        Set Result = New cFluent
        Set Result2 = New cFluent
        Result2.TestValue = True

        Debug.Assert Result2.Should.Be.EqualTo(True) '//true
        Debug.Assert Result.Of(True).Should.Be.EqualTo(True) '//true
        Debug.Assert Result.Of(True).Should.Be.EqualTo(False) '//false
    End Sub

# Notes
This framework is currently in beta. The design of the API is subject to change.
