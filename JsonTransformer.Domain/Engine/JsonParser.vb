Imports Json.Exceptions.Lexer
Imports Json.Token

Namespace Engine

    Public Class JsonParser

        ' constants
        Private Const ARRAY_NAME As String = "array"
        Private Const BOOLEAN_NAME As String = "boolean"
        Private Const ITEM_NAME As String = "item"
        Private Const KEY_NAME As String = "key"
        Private Const NULL_NAME As String = "null"
        Private Const NUMBER_NAME As String = "number"
        Private Const OBJECT_NAME As String = "object"
        Private Const STRING_NAME As String = "string"
        Private Const VALUE_NAME As String = "value"

        ' token mappings
        Private Shared ReadOnly _tokenMap As New Dictionary(Of JsonTokenType, String) From {
            {JsonTokenType.FalseLiteral, BOOLEAN_NAME},
            {JsonTokenType.NullLiteral, NULL_NAME},
            {JsonTokenType.NumberLiteral, NUMBER_NAME},
            {JsonTokenType.StringLiteral, STRING_NAME},
            {JsonTokenType.TrueLiteral, BOOLEAN_NAME}
        }

        Private ReadOnly _lexer As JsonLexer
        Private _current As JsonToken

        Public Sub New(lexer As JsonLexer)
            _lexer = lexer
            _current = _lexer.NextToken()
        End Sub

        Public Function Parse() As XElement
            Dim root As XElement = ParseValue()

            If (_current.Type <> JsonTokenType.EOF) Then
                Throw New JsonParserException(_current.Position, $"Unexpected token after document end: {_current.Type}")
            End If

            Return root
        End Function

        Private Function ParseValue() As XElement
            If (IsToken(JsonTokenType.OpenBrace)) Then
                Return ParseObject()
            End If
            If (IsToken(JsonTokenType.OpenBracket)) Then
                Return ParseArray()
            End If

            Dim mappedTokenName As String = String.Empty
            If (_tokenMap.TryGetValue(_current.Type, mappedTokenName)) Then
                Dim value As Object = _current.Value
                If (IsToken(JsonTokenType.TrueLiteral)) Then
                    value = True
                ElseIf (IsToken(JsonTokenType.FalseLiteral)) Then
                    value = False
                ElseIf (IsToken(JsonTokenType.NullLiteral)) Then
                    value = Nothing
                End If

                Advance()
                Return New XElement(mappedTokenName, value)
            End If

            Throw New JsonParserException(_current.Position, $"Unexpected token: {_current.Type}.")
        End Function

        Private Function ParseObject() As XElement
            Expect(JsonTokenType.OpenBrace)

            Dim jsonObject As New XElement(OBJECT_NAME)

            If (Not IsToken(JsonTokenType.CloseBrace)) Then
                Do
                    Expect(JsonTokenType.StringLiteral, False)

                    Dim key As String = _current.Value
                    Advance()

                    Expect(JsonTokenType.Colon)

                    Dim keyElement As New XElement(KEY_NAME, key)
                    Dim valueElement As New XElement(VALUE_NAME, ParseValue())
                    Dim item As New XElement(ITEM_NAME)
                    item.Add(keyElement)
                    item.Add(valueElement)
                    jsonObject.Add(item)

                    If (Not IsToken(JsonTokenType.Comma)) Then
                        Exit Do
                    End If

                    Advance()
                Loop
            End If

            Expect(JsonTokenType.CloseBrace)
            Return jsonObject
        End Function

        Private Function ParseArray() As XElement
            Expect(JsonTokenType.OpenBracket)

            Dim jsonArray As New XElement(ARRAY_NAME)

            If (Not IsToken(JsonTokenType.CloseBracket)) Then
                Do
                    jsonArray.Add(ParseValue())

                    If (Not IsToken(JsonTokenType.Comma)) Then
                        Exit Do
                    End If

                    Advance()
                Loop
            End If

            Expect(JsonTokenType.CloseBracket)
            Return jsonArray
        End Function

        Private Sub Advance()
            _current = _lexer.NextToken()
        End Sub

        Private Sub Expect(expected As JsonTokenType)
            Expect(expected, True)
        End Sub

        Private Sub Expect(expected As JsonTokenType, shouldAdvance As Boolean)
            If (Not IsToken(expected)) Then
                Throw New JsonParserException(_current.Position, $"Expected {expected} but found {_current.Type}.")
            End If

            If (shouldAdvance) Then
                Advance()
            End If
        End Sub

        Private Function IsToken(expectedToken As JsonTokenType) As Boolean
            Return _current.Type = expectedToken
        End Function

    End Class

End Namespace
