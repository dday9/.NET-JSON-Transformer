# .NET JSON Transformer

A Visual Basic .NET JSON transformer that converts a JSON literal into a strongly structured XML ([XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument)) representation.

This project provides a custom JSON lexer and recursive-descent parser, producing a 1-to-1, lossless XML representation of JSON values, objects, and arrays.

## Key Features

-   Written in **VB.NET**    
-   Targets .NET Standard 2.0 (usable from .NET Framework, .NET Core, .NET 5+)
-   Fully custom JSON lexer + parser (no external JSON libraries)
-   Produces safe, schema-stable XML
-   Supports:
    -   Strings (including all escape sequences)
    -   Numbers (culture-aware or custom numeric formats)
    -   Booleans
    -   Null
    -   Arrays
    -   Objects
-   Extensive unit test coverage (test project targets .NET 10)

### Breaking Changes (Rewrite)

This version is a **complete rewrite** of the original implementation.

1. Architecture rewrite
	 - Old monolithic parser replaced with:
		 - JsonLexer (tokenization)
		 - JsonParser (grammar enforcement)
		 - JsonTransformer (public API)
2.  Boolean values are semantic
	- Booleans are now emitted as real boolean values, not string literals.
3. Whitespace handling
	- Leading whitespace is preserved for accurate error positions.

## Adding the Library to Your Project

### Option 1: Add the compiled DLL

The transformer is built as a .NET Standard 2.0 class library and included in the v2.0 release: [v2.0](../../releases/tag/v2.0)

1. Right-click your project in the _Solution Explorer_
2. Add Project Reference
3. Go to the _Browse_ tab
4. Click on the _Browse..._ button
5. Select the DLL
6. Click on the _OK_ button

### Option 2: Reference the Class Library

The transformer is built as a .NET Standard 2.0 class library.

 1. Build the project
 2. Reference the compiled DLL from your project
 3. Import the namespace:
```vb
Imports Json
```

This works from:

 - .NET Framework
 - .NET Core
 - .NET 5+

### Option 3: Add Source Files Directly

If you prefer source inclusion:

 1. Copy the following files into your project:
	 - JsonLexer.vb
	 - JsonParser.vb
	 - JsonTransformer.vb
	 - All the custom exception and token class files
 2. Ensure your project references:
   - `Microsoft.VisualBasic`
	 - `System.Globalization`
   - `System.Collections.Generic`
   - `System.Linq`
	 - `System.Text`
	 - `System.Xml.Linq`

## JsonTransformer.Parse Method

### Definition
Namespace: Json
Creates a new XDocument from a JSON literal

### Overloads

| Name                                  | Description                                                                                        |
|---------------------------------------|----------------------------------------------------------------------------------------------------|
| Parse(String)                         | Creates a new XDocument from a JSON literal using the official JSON EBNF for parsing numbers.      |
| Parse(String, CultureInfo)            | Creates a new XDocument from a JSON literal using culture specific characters for parsing numbers. |
| Parse(String, String, String, String) | Creates a new XDocument from a JSON literal using custom values for parsing numbers.               |

### Remarks

Using one of the overloads of this method, you can transform a JSON literal into an XDocument while optionally specifying how number parsing should be handled.

#### Parse(String)

Creates a new XDocument from a JSON literal using the official JSON EBNF for parsing numbers.

```vb
Public Shared Function Parse(source As String) As XDocument
```

##### Parameters

`source` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string)
The JSON literal to be transformed.

##### Returns

[XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument)
An [XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument) object that contains the transformed JSON data.

##### Remarks

This overload follows the EBNF defined in [JSON.org](https://www.json.org/) with the following deviations:

 - Booleans
	 - Case-insensitive (`true`, `TRUE`, `True`, etc.)
 - Null
	 - Case-insensitive (`null`, `NULL`)

##### Example
```vb
Dim literal = "{ ""Property1"": 1, ""Property2"": false }"
Dim document As XDocument = JsonTransformer.Parse(literal)

Debug.WriteLine(document.ToString())
```

#### Parse(String, CultureInfo)

Creates a new XDocument from a JSON literal using culture specific characters for parsing numbers.

```vb
Public Shared Function Parse(source As String, culture As CultureInfo) As XDocument
```

##### Parameters

`source` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string)
The JSON literal to be transformed.

`culture` [CultureInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)
The culture to use when parsing numbers.

##### Returns

[XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument)
An [XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument) object that contains the transformed JSON data.

##### Remarks

This overload follows the EBNF defined in [JSON.org](https://www.json.org/) with the following deviations:

 - Numbers
	 - Unary positive sign supported at beginning of numbers
	 - Decimal place derived from:
		 - [CultureInfo.NumberFormat.NumberDecimalSeparator](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo.numberdecimalseparator)
	 - Unary positive and negative signs derived from:
		 - [CultureInfo.NumberFormat.NegativeSign](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo.negativesign)
		 - [CultureInfo.NumberFormat.PositiveSign](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo.positivesign)
 - Booleans
	 - Case-insensitive (`true`, `TRUE`, `True`, etc.)
 - Null
	 - Case-insensitive (`null`, `NULL`)

##### Example
```vb
Dim literal = "123,45"
Dim culture As New Globalization.CultureInfo("de-DE")
Dim document As XDocument = JsonTransformer.Parse(literal, culture)

Debug.WriteLine(document.ToString())
```

#### Parse(String, String, String, String)

Creates a new XDocument from a JSON literal using custom values for parsing numbers.

```vb
Public Shared Function Parse(source As String, Optional decimalSeparator As String = DEFAULT_DECIMAL_SEPARATOR, Optional negativeSign As String = DEFAULT_UNARY_NEGATIVE_OPERATOR, Optional positiveSign As String = DEFAULT_UNARY_POSITIVE_OPERATOR) As XDocument
```

##### Parameters

`source` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string)
The JSON literal to be transformed.

`decimalSeparator` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string) (*optional*)
The custom decimal separator to be used.

`negativeSign` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string) (*optional*)
The negative sign character to be used

`positiveSign` [String](https://learn.microsoft.com/en-us/dotnet/api/system.string) (*optional*)
The positive sign character to be used

##### Returns

[XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument)
An [XDocument](https://learn.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument) object that contains the transformed JSON data.

##### Remarks

This overload follows the EBNF defined in [JSON.org](https://www.json.org/) with the following deviations:

 - Numbers
	 - Unary positive sign supported at beginning of numbers
	 - If any of the optional parameters are not set, they default to the official [JSON.org](https://www.json.org/) EBNF.
 - Booleans
	 - Case-insensitive (`true`, `TRUE`, `True`, etc.)
 - Null
	 - Case-insensitive (`null`, `NULL`)

##### Example
```vb
Dim decimalSeparator As String = "|"
Dim negativeSign As String = "~"
Dim positiveSign As String = "^"
Dim literal = "123,45"
Dim document As XDocument = JsonTransformer.Parse(literal, decimalSeparator, negativeSign, positiveSign)

Debug.WriteLine(document.ToString())
```

## Kitchen Sink Example

```vb
Dim source As String =
"{
    ""glossary"": {
        ""title"": ""example glossary"",
        ""GlossDiv"": {
            ""title"": ""S"",
            ""GlossList"": {
                ""GlossEntry"": {
                    ""ID"": ""SGML"",
                    ""SortAs"": ""SGML"",
                    ""GlossTerm"": ""Standard Generalized Markup Language"",
                    ""Acronym"": ""SGML"",
                    ""Abbrev"": ""ISO 8879:1986"",
                    ""GlossDef"": {
                        ""para"": ""A meta-markup language, used to create markup languages such as DocBook."",
                        ""GlossSeeAlso"": [""GML"", ""XML""]
                    },
                    ""GlossSee"": ""markup""
                }
            }
        }
    }
}"
Dim results As XDocument = Json.JsonTransformer.Parse(source)

Debug.WriteLine(results.ToString())
```

Live Demo: [https://dotnetfiddle.net/9E9vDo](https://dotnetfiddle.net/9E9vDo)

## ❤️ Support

This project is free and open source.

Show your support! Your (non-tax deductible) donation of Monero cryptocurrency is a sign of solidarity among web developers.

Being self-taught, I have come a long way over the years. I certainly do not intend on making a living from this free feature, but my hope is to earn a few dollars to validate all of my hard work.

Monero Address: 447SPi8XcexZnF7kYGDboKB6mghWQzRfyScCgDP2r4f2JJTfLGeVcFpKEBT9jazYuW2YG4qn51oLwXpQJ3oEXkeXUsd6TCF

![447SPi8XcexZnF7kYGDboKB6mghWQzRfyScCgDP2r4f2JJTfLGeVcFpKEBT9jazYuW2YG4qn51oLwXpQJ3oEXkeXUsd6TCF](monero.png)
