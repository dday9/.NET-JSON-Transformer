Public Class Json

    ''' <summary>
    ''' Creates a new <see cref="XDocument"/> from a JSON literal
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <returns><see cref="XDocument"/>
    ''' An <see cref="XDocument"/> populated from the string that contains JSON.</returns>
    Public Shared Function Parse(source As String) As XDocument
        'Remove any whitespace
        source = source.Trim()
        If String.IsNullOrWhiteSpace(source) Then
            Return Nothing
        End If

        Dim value = ParseValue(source, 0)
        Dim document = If(value IsNot Nothing, New XDocument(New XDeclaration("1.0", "utf-8", "yes"), value), Nothing)

        Return document
    End Function

    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.</remarks>
    Private Shared Function ParseValue(source As String, ByRef index As Integer) As XElement
        'Declare a temporary placeholder and skip any whitespace
        Dim tempIndex = SkipWhitespace(source, index)

        'Go through each available value until one returns something that isn't null
        Dim value = ParseObject(source, tempIndex)
        If (value Is Nothing) Then
            value = ParseArray(source, tempIndex)
            If (value Is Nothing) Then
                value = ParseString(source, tempIndex)
                If (value Is Nothing) Then
                    value = ParseNumber(source, tempIndex)
                    If (value Is Nothing) Then
                        value = ParseBoolean(source, tempIndex)
                        If (value Is Nothing) Then
                            value = ParseNull(source, tempIndex)
                        End If
                    End If
                End If
            End If
        End If

        'Change the index
        index = tempIndex

        'Return the value
        Return value
    End Function

    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is an Object
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.</remarks>
    Private Shared Function ParseObject(source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a temporary placeholder
        Dim tempIndex = index

        'Check for the starting opening curly bracket
        If source(tempIndex).Equals("{"c) Then
            'Increment the index
            tempIndex += 1

            'Declare a collection that will make up the nodes' value
            Dim nodes = New List(Of Tuple(Of String, XElement))

            'Declare an XElement to store the key (aka - name) of the KeyValuePair
            Dim key As XElement

            'Declare an XElement which will represent the value of the KeyValuePair
            Dim item As XElement

            'Loop until we've reached the end of the source or until we've hit the ending bracket
            Do While (tempIndex < source.Length AndAlso Not source(tempIndex).Equals("}"c))
                'Increment the index and skip any unneeded whitespace
                tempIndex = SkipWhitespace(source, tempIndex)

                'Attempt to parse the String
                key = ParseString(source, tempIndex)

                'Check if the parse was successful
                If (key Is Nothing) Then
                    Throw New Exception($"Expected a String instead of a '{source(tempIndex)}' at position: {tempIndex}.")
                Else
                    'Skip any unneeded whitespace
                    tempIndex = SkipWhitespace(source, tempIndex)

                    If (tempIndex < source.Length) Then
                        'Check if the currently iterated character is a object separator ':'
                        If (source(tempIndex) = ":"c) Then
                            'Increment the index and skip any unneeded whitespace
                            tempIndex = SkipWhitespace(source, tempIndex + 1)

                            If (tempIndex < source.Length) Then
                                'Assign the item to the parsed value
                                item = ParseValue(source, tempIndex)

                                'Check if the parse was successful
                                If (item Is Nothing) Then
                                    Throw New Exception($"Unexpected character '{source(tempIndex)}' at position: {tempIndex}.")
                                Else
                                    'Add the item to the collection
                                    nodes.Add(New Tuple(Of String, XElement)(key.Value, item))

                                    'Skip any unneeded whitespace
                                    tempIndex = SkipWhitespace(source, tempIndex)

                                    'Check if we can continue
                                    If (tempIndex < source.Length) Then
                                        'Check if the currently iterated character is either a item separator (comma) or ending curly bracket
                                        If (source(tempIndex).Equals(","c)) Then
                                            'Increment the index and skip any unneeded whitespace
                                            tempIndex = SkipWhitespace(source, tempIndex + 1)
                                        ElseIf (source(tempIndex) <> "}"c) Then
                                            Throw New Exception($"Expected a ',' instead of a '{source(tempIndex)}' at position: {tempIndex}.")
                                        End If
                                    End If
                                End If
                            Else
                                Throw New Exception("Expected an Object, Array, String, Number, Boolean, or Null instead I reached the end of the source code.")
                            End If
                        Else
                            Throw New Exception($"Expected a ':' instead of a '{source(tempIndex)}' at position: {tempIndex}.")
                        End If
                    Else
                        Throw New Exception("Expected a ',' instead I reached the end of the source code.")
                    End If
                End If
            Loop

            'Check if the currently iterated value is an ending curly bracket
            If (tempIndex < source.Length AndAlso source(tempIndex) = "}"c) Then
                'Increment the index
                tempIndex += 1

                'Set the new index
                index = tempIndex

                'Create the Object
                value = New XElement("object")

                'Iterate through each item in the nodes
                Dim objectItem, objectKey, objectValue As XElement
                For Each n As Tuple(Of String, XElement) In nodes
                    objectKey = New XElement("key", n.Item1)
                    objectValue = New XElement("value", n.Item2)
                    
                    objectItem = New XElement("item", {objectKey, objectValue})
                    value.Add(objectItem)
                Next
            End If
        End If

        'Return the value
        Return value
    End Function
    
    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is an array
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.</remarks>
    Private Shared Function ParseArray(source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a temporary placeholder
        Dim tempIndex As Integer = index

        'Check for the starting opening bracket
        If (source(tempIndex).Equals("["c)) Then
            'Increment the index
            tempIndex += 1

            'Declare a collection that will make up the nodes' value
            Dim nodes = New List(Of XElement)

            'Declare an XElement which will represent the currently iterated item in the array
            Dim item As XElement

            'Loop until we've reached the end of the source or until we've hit the ending bracket
            Do While (tempIndex < source.Length AndAlso Not source(tempIndex).Equals("]"c))
                'Assign the item to the parsed value
                item = ParseValue(source, tempIndex)

                'Check if the parse was successful
                If (item Is Nothing) Then
                    Throw New Exception($"Unexpected character '{source(tempIndex)}' at position: {tempIndex}.")
                Else
                    'Add the item to the collection
                    nodes.Add(item)

                    'Skip any unneeded whitespace
                    tempIndex = SkipWhitespace(source, tempIndex)

                    'Check if we can continue
                    If (tempIndex < source.Length) Then
                        'Check if the currently iterated character is either a item separator (comma) or ending bracket
                        If (source(tempIndex).Equals(","c)) Then
                            'Increment the index and skip any unneeded whitespace
                            tempIndex = SkipWhitespace(source, tempIndex + 1)
                        ElseIf (source(tempIndex) <> "]"c) Then
                            Throw New Exception($"Expected a ',' instead of a '{source(tempIndex)}' at position: {tempIndex}.")
                        End If
                    End If
                End If
            Loop

            'Check if the currently iterated value is an ending bracket
            If (tempIndex < source.Length AndAlso source(tempIndex) = "]"c) Then
                'Increment the index
                tempIndex += 1

                'Set the new index
                index = tempIndex

                'Create the Array
                value = New XElement("array", nodes)
            End If
        End If

        Return value
    End Function
    
    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is a String
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.
    ''' 3. The parser deviates from ECMA-404 by not checking for "\u" followed by 4 hexadecimal characters as an escape character</remarks>
    Private Shared Function ParseString(source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a CONST to store the double-quote character and escaped characters
        Const doubleQuote = """"c
        Const escapedCharacters As String = doubleQuote & "\/bfnrt"

        'Declare a temporary placeholder
        Dim tempIndex As Integer = index

        'Check for the starting double-quote
        If (source(tempIndex).Equals(doubleQuote)) Then
            'Increment the index
            tempIndex += 1

            'Loop until we've reached the end of the source or until we've hit the ending double-quote
            Do While (tempIndex < source.Length AndAlso Not source(tempIndex).Equals(doubleQuote))
                'Check if we're at an escaped character
                If (source(tempIndex) = "\"c AndAlso
                    tempIndex + 1 < source.Length AndAlso
                    escapedCharacters.IndexOf(source(tempIndex + 1)) <> -1) Then
                    tempIndex += 1
                ElseIf (source(tempIndex) = "\"c) Then
                    Throw New Exception("Unescaped backslash in a String. Position: " & index)
                End If

                'Increment the index
                tempIndex += 1
            Loop

            'Check if the currently iterated character is a double-quote
            If (tempIndex < source.Length AndAlso source(tempIndex).Equals(doubleQuote)) Then
                'Increment the index
                tempIndex += 1

                'Create the String
                value = New XElement("string", source.Substring(index + 1, tempIndex - index - 2))

                'Set the new index
                index = tempIndex
            End If
        End If

        Return value
    End Function

    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is a number
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.
    ''' 3. The parser deviates from ECMA-404 by checking for an optional unary positive sign operator</remarks>
    Private Shared Function ParseNumber(source As String, ByRef index As Integer) As XElement
        'Get the current culture information
        Dim culture = Globalization.CultureInfo.CurrentCulture

        'Declare a temporary placeholder
        Dim tempIndex = index

        'Check for the optional unary operator
        If (source.IndexOf(culture.NumberFormat.NegativeSign, tempIndex, StringComparison.OrdinalIgnoreCase) = tempIndex OrElse
            source.IndexOf(culture.NumberFormat.PositiveSign, tempIndex, StringComparison.OrdinalIgnoreCase) = tempIndex) Then
            tempIndex += 1
        End If

        'Match one or more digits
        If (tempIndex < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex).ToString()) <> -1) Then
            Do While (tempIndex < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex).ToString()) <> -1)
                tempIndex += 1
            Loop
        Else
            Return Nothing
        End If

        'Optionally match a decimal separator followed by one or more digits
        If (tempIndex + 1 < source.Length AndAlso
            source.IndexOf(culture.NumberFormat.NumberDecimalSeparator, tempIndex, StringComparison.OrdinalIgnoreCase) = tempIndex AndAlso
            Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex + 1).ToString()) <> -1) Then

            tempIndex += 1
            Do While (tempIndex < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex).ToString()) <> -1)
                tempIndex += 1
            Loop
        End If

        'Optionally match an exponent, followed by an optional unary operator, followed by 1 or more digits
        If (tempIndex + 1 < source.Length AndAlso
            source.IndexOf("e", tempIndex, StringComparison.OrdinalIgnoreCase) = tempIndex) Then

            If (tempIndex + 2 < source.Length) AndAlso
                (source.IndexOf(culture.NumberFormat.NegativeSign, tempIndex + 1, StringComparison.OrdinalIgnoreCase) = tempIndex + 1 OrElse
                source.IndexOf(culture.NumberFormat.PositiveSign, tempIndex + 1, StringComparison.OrdinalIgnoreCase) = tempIndex + 1) AndAlso
                Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex + 2).ToString()) <> -1 Then
                tempIndex += 2
            ElseIf (tempIndex + 1 < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex + 1).ToString()) <> -1) Then
                tempIndex += 1
            End If

            Do While (tempIndex < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(tempIndex).ToString()) <> -1)
                tempIndex += 1
            Loop
        End If

        'Create the number
        Dim value = New XElement("number", source.Substring(index, tempIndex - index))
        index = tempIndex

        Return value
    End Function
    
    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is a boolean
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.
    ''' 3. The parser deviates from ECMA-404 by ignoring the casing</remarks>
    Private Shared Function ParseBoolean(source As String, ByRef index As Integer) As XElement
        Dim value As XElement = Nothing

        'Literally match 'true' or 'false'
        If source.IndexOf("true", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("boolean", True)
            index += 4
        ElseIf source.IndexOf("false", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("boolean", False)
            index += 5
        End If

        Return value
    End Function
    
    ''' <summary>
    ''' Creates a new <see cref="XElement"/> from a JSON literal at given index where the expected JSON type is the literal "null"
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the parsing will begin.</param>
    ''' <returns>An <see cref="XElement"/> populated from the string that contains JSON.</returns>
    ''' <remarks>1. <paramref name="index"/> will increment if the parse is successful.
    ''' 2. Nothing will be returned if the parse is not successful.
    ''' 3. The parser deviates from ECMA-404 by ignoring the casing</remarks>
    Private Shared Function ParseNull(source As String, ByRef index As Integer) As XElement
        Dim value As XElement = Nothing

        'Literally match 'null' in the source starting at the index
        If source.IndexOf("null", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("null")
            index += 4
        End If

        Return value
    End Function

    ''' <summary>
    ''' Starting at a given index, skip any character that is whitespace
    ''' </summary>
    ''' <param name="source">A string that contains JSON.</param>
    ''' <param name="index">The position of the JSON where the whitespace check will begin</param>
    ''' <returns>An Integer where the first character of <paramref name="source"/>, starting at <paramref name="index"/>, is not whitespace.</returns>
    Private Shared Function SkipWhitespace(source As String, index As Integer) As Integer
        Do While (index < source.Length AndAlso Char.IsWhiteSpace(source(index)))
            index += 1
        Loop

        Return index
    End Function

End Class
