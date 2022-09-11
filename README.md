# Fluent-VBA
Fluent VBA is a fluent unit testing framework for VBA. This project was inspired by [Fluent Assertions](https://fluentassertions.com/introduction).

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
    
    Public Function returnVal(value As Variant) As Variant
        returnVal = value
    End Function
 
You can have code that reads more naturally like so:

    Option Explicit

    Sub FluentUnitTestExample1
        Dim Result As cFluent
        Dim returnedResult As Variant
        
        '//Arrange
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

All of the tests are written in the mTests.bas module. There are over 150 tests within this file. The tests use a combination of cFluent and cFluentOf objects to test the framework. You can see more information regarding the testing notes [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/Testing-notes)

# External projects

Fluent VBA supports being used in external projects. However, I'm still in the process of testing this. You can read more about using Fluent VBA in external projects [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/External-projects)

# TODO: High level API design overview

I'd like to write a high-level overview of the API's design. This had been completed previously but is now outdated. You can find a post of mine describing an older version of the API's structure on CodeReview on StackExchange [here](https://codereview.stackexchange.com/questions/267836/a-fluent-unit-testing-framework-in-vba). I will likely be updating this within the coming weeks.

# Contacting me

You can contact me at b.gonzalez.programming@gmail.com. Feel free to contact me regarding bug fixes, contributions (see more below), or other topics.

# Contributing to Fluent VBA

I'm open to external contributions for Fluent VBA. I do need to work on a style guide to determine how I'd like such contributions to be implemented. I also expect any contributions to have unit tests using the Fluent VBA framework.

# Feature requests

You are free to contact me regarding feature requests as long as you understand that I'm not obligated to implement them. I expect messages to be polite, respectful, and without a sense of entitlement. As long as you do those things, I'm happy to hear what you have to say.

# Final notes

The API is in a good and usable state. Overall I'm pretty happy with the API's internal and external design. As of right now, I only anticipate internal changes and feature enhancements. So the design of the API should be relatively stable. I'd be open to changing certain design aspects of the API if I found a good reason to do so however. Naturally, this is dependant on time and availability on my part however.
