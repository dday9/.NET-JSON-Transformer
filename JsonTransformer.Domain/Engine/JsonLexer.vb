Imports System.Globalization
Imports System.Text
Imports Json.Exceptions.Lexer
Imports Json.Token

Namespace Engine

    Public Class JsonLexer

        ' constants
        Private Const COLON As Char = ":"c
        Private Const COMMA As Char = ","c
        Private Const CLOSE_BRACE As Char = "}"c
        Private Const CLOSE_BRACKET As Char = "]"c
        Private Const DEFAULT_DECIMAL_SEPARATOR As String = "."
        Private Const DEFAULT_UNARY_NEGATIVE_OPERATOR As String = "-"
        Private Const DEFAULT_UNARY_POSITIVE_OPERATOR As String = ""
        Private Const DEFAULT_UNARY_POSITIVE_EXPONENT_OPERATOR As String = "+"
        Private Const DOUBLE_QUOTE As Char = """"c
        Private Const ESCAPING_CHARACTER As Char = "\"c
        Private Const EXPONENT_LOWERCASE As Char = "e"c
        Private Const EXPONENT_UPPERCASE As Char = "E"c
        Private Const FALSE_LITERAL As String = "false"
        Private Const NULL_LITERAL As String = "null"
        Private Const OPEN_BRACE As Char = "{"c
        Private Const OPEN_BRACKET As Char = "["c
        Private Const TRUE_LITERAL As String = "true"

        ' token mappings
        Private Shared ReadOnly _escapeMap As New Dictionary(Of Char, String) From {
            {DOUBLE_QUOTE, DOUBLE_QUOTE},
            {ESCAPING_CHARACTER, ESCAPING_CHARACTER},
            {"/"c, "/"},
            {"b"c, Char.ConvertFromUtf32(8)},
            {"f"c, Char.ConvertFromUtf32(12)},
            {"n"c, Char.ConvertFromUtf32(10)},
            {"r"c, Char.ConvertFromUtf32(13)},
            {"t"c, Char.ConvertFromUtf32(9)}
        }
        Private Shared ReadOnly _tokenMap As New Dictionary(Of Char, JsonTokenType) From {
            {COLON, JsonTokenType.Colon},
            {COMMA, JsonTokenType.Comma},
            {OPEN_BRACE, JsonTokenType.OpenBrace},
            {OPEN_BRACKET, JsonTokenType.OpenBracket},
            {CLOSE_BRACE, JsonTokenType.CloseBrace},
            {CLOSE_BRACKET, JsonTokenType.CloseBracket}
        }

        ' constants with optional overrides
        Private ReadOnly _decimalSeparator As String
        Private ReadOnly _unaryNegativeOperator As String
        Private ReadOnly _unaryPositiveOperator As String
        Private ReadOnly _unaryPositiveExponentOperator As String

        Private ReadOnly _source As String
        Private _index As Integer

        Public Sub New(source As String)
            _source = source
            _index = 0

            _decimalSeparator = DEFAULT_DECIMAL_SEPARATOR
            _unaryNegativeOperator = DEFAULT_UNARY_NEGATIVE_OPERATOR
            _unaryPositiveOperator = DEFAULT_UNARY_POSITIVE_OPERATOR
            _unaryPositiveExponentOperator = DEFAULT_UNARY_POSITIVE_EXPONENT_OPERATOR
        End Sub

        Public Sub New(source As String, cultureOverride As CultureInfo)
            _source = source
            _index = 0

            _decimalSeparator = cultureOverride.NumberFormat.NumberDecimalSeparator
            _unaryNegativeOperator = cultureOverride.NumberFormat.NegativeSign
            _unaryPositiveOperator = cultureOverride.NumberFormat.PositiveSign
            _unaryPositiveExponentOperator = _unaryPositiveOperator
        End Sub

        Public Sub New(source As String, Optional decimalSeparator As String = DEFAULT_DECIMAL_SEPARATOR, Optional negativeSign As String = DEFAULT_UNARY_NEGATIVE_OPERATOR, Optional positiveSign As String = DEFAULT_UNARY_POSITIVE_OPERATOR)
            _source = source
            _index = 0

            _decimalSeparator = decimalSeparator
            _unaryNegativeOperator = negativeSign
            _unaryPositiveOperator = positiveSign
            _unaryPositiveExponentOperator = _unaryPositiveOperator
            If (String.IsNullOrWhiteSpace(_unaryPositiveExponentOperator)) Then
                _unaryPositiveExponentOperator = DEFAULT_UNARY_POSITIVE_EXPONENT_OPERATOR
            End If
        End Sub

        Public Function NextToken() As JsonToken
            SkipWhitespace()

            If (_index >= _source.Length) Then
                Return New JsonToken(JsonTokenType.EOF, Nothing, _index)
            End If

            Dim currentCharacter As Char = _source(_index)
            Dim mappedToken As JsonTokenType
            If (_tokenMap.TryGetValue(currentCharacter, mappedToken)) Then
                Return Token(currentCharacter, mappedToken)
            End If
            If (currentCharacter.Equals(DOUBLE_QUOTE)) Then
                Return ReadString()
            End If
            Dim currentCharacterAsString As String = currentCharacter.ToString()
            If (Char.IsDigit(currentCharacter) OrElse currentCharacterAsString.Equals(_unaryNegativeOperator) OrElse (Not String.IsNullOrWhiteSpace(_unaryPositiveOperator) AndAlso currentCharacterAsString.Equals(_unaryPositiveOperator))) Then
                Return ReadNumber()
            End If

            Return ReadLiteral()
        End Function

        Private Function Token(character As Char, tokenType As JsonTokenType) As JsonToken
            Dim startingIndex As Integer = _index
            _index += 1

            Dim conversion As New JsonToken(tokenType, character, startingIndex)
            Return conversion
        End Function

        Private Function ReadString() As JsonToken
            Dim startingIndex As Integer = _index
            _index += 1

            Dim jsonStringBuilder As New StringBuilder()
            While (_index < _source.Length)
                Dim currentCharacter As Char = _source(_index)

                If (currentCharacter.Equals(DOUBLE_QUOTE)) Then
                    _index += 1
                    Return New JsonToken(JsonTokenType.StringLiteral, jsonStringBuilder.ToString(), startingIndex)
                End If

                If (currentCharacter.Equals(ESCAPING_CHARACTER)) Then
                    _index += 1
                    If _index >= _source.Length Then
                        Throw New JsonLexerException("Unterminated string")
                    End If

                    Dim escapedChar As Char = _source(_index)
                    Dim mappedString As String = String.Empty
                    If (_escapeMap.TryGetValue(escapedChar, mappedString)) Then
                        jsonStringBuilder.Append(mappedString)
                    ElseIf (escapedChar.Equals("u"c)) Then
                        If (_index + 4 >= _source.Length) Then
                            Throw New JsonLexerInvalidEscapeException("Invalid unicode escape")
                        End If

                        Dim hex As String = _source.Substring(_index + 1, 4)
                        If (Not hex.All(Function(c) Uri.IsHexDigit(c))) Then
                            Throw New JsonLexerInvalidEscapeException("Invalid unicode escape")
                        End If

                        jsonStringBuilder.Append(ChrW(Convert.ToInt32(hex, 16)))
                        _index += 4
                    Else
                        Throw New JsonLexerInvalidEscapeException("Invalid escape")
                    End If
                Else
                    jsonStringBuilder.Append(currentCharacter)
                End If

                _index += 1
            End While

            Throw New JsonLexerException("Unterminated string")
        End Function

        Private Function ReadNumber() As JsonToken
            Dim startingIndex As Integer = _index

            If (_source(_index) = _unaryNegativeOperator OrElse (Not String.IsNullOrWhiteSpace(_unaryPositiveOperator) AndAlso _source(_index) = _unaryPositiveOperator)) Then
                _index += 1
            End If

            Dim integerStart As Integer = _index
            While (_index < _source.Length AndAlso Char.IsDigit(_source(_index)))
                _index += 1
            End While

            If (integerStart = _index) Then
                Throw New JsonLexerInvalidNumberException("Expected digit")
            End If

            If (_index < _source.Length AndAlso _source(_index) = _decimalSeparator) Then
                _index += 1

                Dim fractionStart As Integer = _index
                While (_index < _source.Length AndAlso Char.IsDigit(_source(_index)))
                    _index += 1
                End While

                If (fractionStart = _index) Then
                    Throw New JsonLexerInvalidNumberException("Decimal point must be followed by digits")
                End If
            End If

            If (_index < _source.Length AndAlso (_source(_index) = EXPONENT_LOWERCASE OrElse _source(_index) = EXPONENT_UPPERCASE)) Then
                _index += 1

                If (_index < _source.Length AndAlso (_source(_index) = _unaryNegativeOperator OrElse _source(_index) = _unaryPositiveExponentOperator)) Then
                    _index += 1
                End If

                Dim exponentStart As Integer = _index
                While (_index < _source.Length AndAlso Char.IsDigit(_source(_index)))
                    _index += 1
                End While

                If (exponentStart = _index) Then
                    Throw New JsonLexerInvalidNumberException("Exponent must contain digits")
                End If
            End If

            Return New JsonToken(JsonTokenType.NumberLiteral, _source.Substring(startingIndex, _index - startingIndex), startingIndex)
        End Function

        Private Function ReadLiteral() As JsonToken
            Dim startingIndex As Integer = _index

            While (_index < _source.Length AndAlso Char.IsLetter(_source(_index)))
                _index += 1
            End While

            Dim word As String = _source.Substring(startingIndex, _index - startingIndex).ToLowerInvariant()

            Select Case word
                Case TRUE_LITERAL
                    Return New JsonToken(JsonTokenType.TrueLiteral, word, startingIndex)
                Case FALSE_LITERAL
                    Return New JsonToken(JsonTokenType.FalseLiteral, word, startingIndex)
                Case NULL_LITERAL
                    Return New JsonToken(JsonTokenType.NullLiteral, word, startingIndex)
                Case Else
            End Select

            Throw New JsonLexerException($"Unexpected literal '{word}'")
        End Function

        Private Sub SkipWhitespace()
            Do While (_index < _source.Length AndAlso Char.IsWhiteSpace(_source(_index)))
                _index += 1
            Loop
        End Sub

    End Class

End Namespace
