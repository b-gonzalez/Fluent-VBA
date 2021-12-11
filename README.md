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
        Dim Result as IFluent

        Set Result = new cFluent
        Result.TestValue = ReturnsFive()

        Debug.Assert Result.Should.Be.EqualTo(5)
    End Sub

Or, arguably, even more naturally using the IFluentOf interface:

    Sub FluentUnitTestExample2
        Dim Result as IFluentOf

        Set Result = new cFluent

        Debug.Assert Result.Of(ReturnsFive).Should.Be.EqualTo(5)
    End Sub

# Testing notes

Most of the tests utilize the IFluent interface. This is because the tests were written before I introduced the new IFluentOf interface (see notes on this interface below). The Meta tests (see below) do include an additional procedure using the IFluentOf interface.
    
# Meta tests

The fluent unit testing framework uses itself to test itself. These set of tests are contained in the mTests module in the MetaTests sub. You can also find a link to the meta tests [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/Meta-Tests).

# Documentation tests

The mTests module has a DocumentationTests sub that contains several dozen tests. These tests document the various objects and methods in the API by using them in the tests. You can also find a link to the documentation tests [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/Documentation-Tests).

# Additional tests

Several other tests are implemented documenting the flexibility with which these unit tests can be created. These tests can be found in the mTests module. You can also find a link to the additional tests [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/Additional-tests).

# IFluentOf interface

One new big change is the addition of the IFluentOf interface. This new interface allows you to enter the test value in the testing line itself. You can read more about using the IFluentOf interface [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/IFluentOf-interface)

# Using Fluent VBA in an external project

All of the class modules in Fluent VBA are PublicNotCreatable. So the project can be used as a reference in other projects. You'd start by adding the project as a reference to whatever workbook you'd like to use. After you did that, you'd create a cFluent object using the MakeFluent() method in the mInit module. Once the object is created you should be able to execute the tests as normal.

# Scrapped additional features

There were a number of features I considered implementing. They were scrapped for a variety of reasons. You can see a detailed breakdown of some featues I considered (but didn't implemnt) as well as my reasoning [here](https://github.com/b-gonzalez/Fluent-VBA/wiki/Scrapped-additional-features)

# TODO: High level API design overview

A high level design of the API. This is mostly been completed previously. You can find a post of mine describing an older version of the API's structure on CodeExchange [here](https://codereview.stackexchange.com/questions/267836/a-fluent-unit-testing-framework-in-vba). It is almost certainly at least a bit outdated. So when I have some time I will take some time to create a post with an updated API design on this project.

# Final notes

I consider this API to be finished. I'm very happy with the API's internal and external design. There's only one small minor final change I'm currently considering making. And like the other recent changes, they're all related to the internal design.

Unless a large bug is discovered or a very good feature is requested, I don't anticipate other future updates to this project.
