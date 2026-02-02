Imports System.Globalization
Imports Json.Engine

Public Class JsonTransformer

    Public Shared Function Parse(source As String) As XDocument
        If (String.IsNullOrWhiteSpace(source)) Then
            Throw New ArgumentNullException(source)
        End If

        source = source.TrimEnd() ' not trimming start to preserve index

        Dim lexer As New JsonLexer(source)
        Dim parser As New JsonParser(lexer)
        Dim parsedJson As XElement = parser.Parse()
        Dim document As New XDocument(New XDeclaration("1.0", "utf-8", "yes"), parsedJson)

        Return document
    End Function

    Public Shared Function Parse(source As String, culture As CultureInfo) As XDocument
        If (String.IsNullOrWhiteSpace(source)) Then
            Throw New ArgumentNullException(source)
        End If

        source = source.TrimEnd() ' not trimming start to preserve index

        Dim lexer As New JsonLexer(source, culture)
        Dim parser As New JsonParser(lexer)
        Dim parsedJson As XElement = parser.Parse()
        Dim document As New XDocument(New XDeclaration("1.0", "utf-8", "yes"), parsedJson)

        Return document
    End Function

    Public Shared Function Parse(source As String, decimalSeparator As String, negativeSign As String, positiveSign As String) As XDocument
        If (String.IsNullOrWhiteSpace(source)) Then
            Throw New ArgumentNullException(source)
        End If

        source = source.TrimEnd() ' not trimming start to preserve index

        Dim lexer As New JsonLexer(source, decimalSeparator, negativeSign, positiveSign)
        Dim parser As New JsonParser(lexer)
        Dim parsedJson As XElement = parser.Parse()
        Dim document As New XDocument(New XDeclaration("1.0", "utf-8", "yes"), parsedJson)

        Return document
    End Function

End Class
