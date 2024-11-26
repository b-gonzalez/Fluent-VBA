# Fluent-VBA
Fluent VBA is an object-oriented [fluent](https://en.wikipedia.org/wiki/Fluent_interface) unit testing library for VBA. This project was inspired by [Fluent Assertions](https://fluentassertions.com/introduction) in C#.

Fluent APIs are intended to be read like natural language. So instead of having something like:

```vba
Option Explicit

Sub NormalUnitTestExample
    Dim result as long
    Dim Assert as cUnitTester
    
    '//Arrange
    Set Assert = New cUnitTester
    
    '//Act
    result = returnVal(5) 'returns the value provided as an argument
    
    '//Assert
    Assert.Equal(Result,5)
End Sub

Public Function returnVal(value As Variant) As Variant
    returnVal = value
End Function
```
 
You can have code that reads more naturally like so:

```vba
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
    Debug.Assert Result.Should.Be.GreaterThan(4)
End Sub
```

Or, arguably, even more naturally using cFluentOf objects like this:

```vba
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
        Debug.Assert .Should.Be.GreaterThan(4)
    End With
    
    '//Or like this
    
    '//Act & Assert
    Debug.Assert Result.Of(returnedResult).Should.Be.EqualTo(5)
    Debug.Assert Result.Of(returnedResult).Should.Be.GreaterThan(4)
End Sub
```

# Getting started

To get started with Fluent VBA you can see the [getting started page](https://github.com/b-gonzalez/Fluent-VBA/wiki/Getting-started) on the wiki.

# Contacting me

You can contact me at b.gonzalez.programming@gmail.com.

# Contributing to Fluent VBA

I'm open to external contributions for Fluent VBA. I do need to work on a style guide to determine how I'd like such contributions to be implemented. I also expect any contributions to have unit tests using the Fluent VBA library.

# Feature requests

You are free to contact me regarding feature requests as long as you understand that I'm not obligated to implement them. I expect messages to be polite, respectful, and without a sense of entitlement. As long as you do those things, I'm happy to hear what you have to say.

# Final notes

The API is in a good and usable state. Overall I'm pretty happy with the API's internal and external design. As of right now, I only anticipate internal changes and feature enhancements. So the design of the API should be relatively stable. I'd be open to changing certain design aspects of the API if I found a good reason to do so however. Naturally, this is dependant on time and availability on my part however.
