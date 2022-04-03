# Fluent-VBA
Fluent VBA is a fluent unit testing framework for VBA

Fluent frameworks are intended to be read like natural language. So instead of having something like:

    Option Explicit

    Sub NormalUnitTestExample
        Dim result as long
        Dim Assert as cUnitTester
        
        '//Arrange
        Set Assert = New cUnitTester
        
        '//Act
        result = returnVal(5) ‘returns the value provided as an argument
        
        '//Assert
        Assert.Equal(Result,5)
    End Sub
    
    Public Function returnVal(value As Variant)
        returnVal = value
    End Function
 
You can have code that reads more naturally like so:

    Option Explicit

    Sub FluentUnitTestExample1
        Dim Result As cFluent
        Dim returnedResult As Variant
        
        '//arrange
        Set Result = New cFluent
        returnedResult = returnVal(5)
        
        '//Act
        Result.TestValue = returnedResult
        
        '//Assert
        Debug.Assert Result.Should.Be.EqualTo(5)
    End Sub

Or, arguably, even more naturally using cFluentOf objects like this:

    Option Explicit

    Sub FluentUnitTestExample2()
        Dim Result As cFluentOf
        Dim returnedResult As Variant
        
        '//arrange
        Set Result = New cFluentOf
        returnedResult = returnVal(5)
        
        '//Act
        With Result.Of(returnedResult)
            '//Assert
            Debug.Assert .Should.Be.EqualTo(5)
        End With
    End Sub
    
Or like this:

    Option Explicit

    Sub FluentUnitTestExample3()
        Dim Result As cFluentOf
        Dim returnedResult As Variant
        
        '//arrange
        Set Result = New cFluentOf
        returnedResult = returnVal(5)
        
        '//Act & Assert
        Debug.Assert Result.Of(returnedResult).Should.Be.EqualTo(5)
    End Sub

# Testing notes

All of the tests are written in the mTests.bas. There are 100+ tests within this file. Most of the tests utilize the IFluent interface. This is because the tests were written before I introduced the new IFluentOf interface (see notes on this interface below). The Meta tests do include additional tests using the IFluentOf interface. And I will probably refactor most of the tests to also use IFluentOf at a later time.
    
# Meta tests

The fluent unit testing framework uses itself to test itself. The mTests module has a MetaTests sub that the framework uses to accomplish this.

# Documentation tests

The mTests module has a DocumentationTests sub that contains several dozen tests. These tests document the various objects and methods in the API by using them in the tests.

# Additional tests

Several other tests are implemented documenting the flexibility with which these unit tests can be created.

# IFluentOf interface

One new big change is the addition of the IFluentOf interface. This new interface allows you to enter the test value in the testing line itself. You can read more about using the IFluentOf interface [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/IFluentOf-interface)

# Using Fluent VBA in an external project

All of the class modules in Fluent VBA are PublicNotCreatable. So the project can be used as a reference in other projects. If you plan on doing this I'd recommend doing the following:

1. Create a testing file that will reference the Fluent VBA workbook and the file containing the code to be tested
2. In the VBA projects references for the testing file, reference both the Fluent VBA workbook and the file containing the code to be tested.
3. Create a testing procedure that has a variable that has the type of IFluent or IFluentOf.
4. Instantiate this variable using the MakeFluent function or the MakeFluentOf function for IFluent or IFluentOf types respectively.
5. Write your tests.

# TODO: High level API design overview

A high level design of the API. This is mostly been completed previously. You can find a post of mine describing an older version of the API's structure on CodeReview on StackExchange [here](https://codereview.stackexchange.com/questions/267836/a-fluent-unit-testing-framework-in-vba). It is almost certainly at least a bit outdated. So when I have some time I will take some time to create a post with an updated API design on this project.

# Final notes

The API is in a good and usable state. Overall I'm pretty happy with the API's internal and external design. As of right now, I only anticipate internal changes and feature enhancements. So the design of the API should be relatively stable. I'd be open to changing certain design aspects of the API if I found a good reason to do so however.
