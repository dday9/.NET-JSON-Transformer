Public Module JSON

    Public Function Parse(ByVal source As String) As XDocument
        'Remove any whitespace
        source = source.Trim()
        If String.IsNullOrWhiteSpace(source) Then Return Nothing

        'Declare a document to return
        Dim document As XDocument = Nothing

        'Declare a value that will make up the document
        Dim value As XElement = Parse_Value(source, 0)
        If value IsNot Nothing Then
            document = New XDocument(New XDeclaration("1.0", "utf-8", "yes"), value)
        End If

        Return document
    End Function

    Private Function Parse_Value(ByVal source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a temporary placeholder and skip any whitespace
        Dim temp_index As Integer = SkipWhitespace(source, index)

        'Go through each available value until one returns something that isn't null
        value = Parse_Object(source, temp_index)
        If value Is Nothing Then
            value = Parse_Array(source, temp_index)
            If value Is Nothing Then
                value = Parse_String(source, temp_index)
                If value Is Nothing Then
                    value = Parse_Number(source, temp_index)
                    If value Is Nothing Then
                        value = Parse_Boolean(source, temp_index)
                        If value Is Nothing Then
                            value = Parse_Null(source, temp_index)
                        End If
                    End If
                End If
            End If
        End If

        'Change the index
        index = temp_index

        'Return the value
        Return value
    End Function

    Private Function Parse_Object(ByVal source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a temporary placeholder
        Dim temp_index As Integer = index

        'Check for the starting opening curly bracket
        If source(temp_index).Equals("{"c) Then
            'Increment the index
            temp_index += 1

            'Declare a collection that will make up the nodes' value
            Dim nodes As List(Of Tuple(Of String, XElement)) = New List(Of Tuple(Of String, XElement))

            'Declare an XElement to store the key (aka - name) of the KeyValuePair
            Dim key As XElement = Nothing

            'Declare an XElement which will represent the value of the KeyValuePair
            Dim item As XElement = Nothing

            'Loop until we've reached the end of the source or until we've hit the ending bracket
            Do While temp_index < source.Length AndAlso Not source(temp_index).Equals("}"c)
                'Increment the index and skip any unneeded whitespace
                temp_index = SkipWhitespace(source, temp_index)

                'Attempt to parse the String
                key = Parse_String(source, temp_index)

                'Check if the parse was successful
                If key Is Nothing Then
                    Throw New Exception($"Expected a String instead of a '{source(temp_index)}' at position: {temp_index}.")
                Else
                    'Skip any unneeded whitespace
                    temp_index = SkipWhitespace(source, temp_index)

                    If temp_index < source.Length Then
                        'Check if the currently iterated character is a object separator ':'
                        If source(temp_index) = ":"c Then
                            'Increment the index and skip any unneeded whitespace
                            temp_index = SkipWhitespace(source, temp_index + 1)

                            If temp_index < source.Length Then
                                'Assign the item to the parsed value
                                item = Parse_Value(source, temp_index)

                                'Check if the parse was successful
                                If item Is Nothing Then
                                    Throw New Exception($"Unexpected character '{source(temp_index)}' at position: {temp_index}.")
                                Else
                                    'Add the item to the collection
                                    nodes.Add(New Tuple(Of String, XElement)(key.Value, item))

                                    'Skip any unneeded whitespace
                                    temp_index = SkipWhitespace(source, temp_index)

                                    'Check if we can continue
                                    If temp_index < source.Length Then
                                        'Check if the currently iterated character is either a item separator (comma) or ending curly bracket
                                        If source(temp_index).Equals(","c) Then
                                            'Increment the index and skip any unneeded whitespace
                                            temp_index = SkipWhitespace(source, temp_index + 1)
                                        ElseIf source(temp_index) <> "}"c Then
                                            Throw New Exception($"Expected a ',' instead of a '{source(temp_index)}' at position: {temp_index}.")
                                        End If
                                    End If
                                End If
                            Else
                                Throw New Exception("Expected an Object, Array, String, Number, Boolean, or Null instead I reached the end of the source code.")
                            End If
                        Else
                            Throw New Exception($"Expected a ':' instead of a '{source(temp_index)}' at position: {temp_index}.")
                        End If
                    Else
                        Throw New Exception("Expected a ',' instead I reached the end of the source code.")
                    End If
                End If
            Loop

            'Check if the currently iterated value is an ending curly bracket
            If temp_index < source.Length AndAlso source(temp_index) = "}"c Then
                'Increment the index
                temp_index += 1

                'Set the new index
                index = temp_index

                'Create the Object
                value = New XElement("object")

                'Iterate through each item in the nodes
                For Each n As Tuple(Of String, XElement) In nodes
                    'Set the name attribute and then add the element to the Object
                    n.Item2.SetAttributeValue("name", n.Item1)
                    value.Add(n.Item2)
                Next
            End If
        End If

        'Return the value
        Return value
    End Function

    Private Function Parse_Array(ByVal source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a temporary placeholder
        Dim temp_index As Integer = index

        'Check for the starting opening bracket
        If source(temp_index).Equals("["c) Then
            'Increment the index
            temp_index += 1

            'Declare a collection that will make up the nodes' value
            Dim nodes As List(Of XElement) = New List(Of XElement)

            'Declare an XElement which will represent the currently iterated item in the array
            Dim item As XElement = Nothing

            'Loop until we've reached the end of the source or until we've hit the ending bracket
            Do While temp_index < source.Length AndAlso Not source(temp_index).Equals("]"c)
                'Assign the item to the parsed value
                item = Parse_Value(source, temp_index)

                'Check if the parse was successful
                If item Is Nothing Then
                    Throw New Exception($"Unexpected character '{source(temp_index)}' at position: {temp_index}.")
                Else
                    'Add the item to the collection
                    nodes.Add(item)

                    'Skip any unneeded whitespace
                    temp_index = SkipWhitespace(source, temp_index)

                    'Check if we can continue
                    If temp_index < source.Length Then
                        'Check if the currently iterated character is either a item separator (comma) or ending bracket
                        If source(temp_index).Equals(","c) Then
                            'Increment the index and skip any unneeded whitespace
                            temp_index = SkipWhitespace(source, temp_index + 1)
                        ElseIf source(temp_index) <> "]"c Then
                            Throw New Exception($"Expected a ',' instead of a '{source(temp_index)}' at position: {temp_index}.")
                        End If
                    End If
                End If
            Loop

            'Check if the currently iterated value is an ending bracket
            If temp_index < source.Length AndAlso source(temp_index) = "]"c Then
                'Increment the index
                temp_index += 1

                'Set the new index
                index = temp_index

                'Create the Array
                value = New XElement("array", nodes)
            End If
        End If

        Return value
    End Function

    Private Function Parse_String(ByVal source As String, ByRef index As Integer) As XElement
        'Declare a value to return
        Dim value As XElement = Nothing

        'Declare a CONST to store the double-quote character and escaped characters
        Const double_quote As Char = """"c
        Const escaped_characters As String = double_quote & "\/bfnrt"

        'Declare a temporary placeholder
        Dim temp_index As Integer = index

        'Check for the starting double-quote
        If source(temp_index).Equals(double_quote) Then
            'Increment the index
            temp_index += 1

            'Loop until we've reached the end of the source or until we've hit the ending double-quote
            Do While temp_index < source.Length AndAlso Not source(temp_index).Equals(double_quote)
                'Check if we're at an escaped character
                If source(temp_index) = "\"c AndAlso
                    temp_index + 1 < source.Length AndAlso
                    escaped_characters.IndexOf(source(temp_index + 1)) <> -1 Then
                    temp_index += 1
                ElseIf source(temp_index) = "\"c Then
                    Throw New Exception("Unescaped backslash in a String. Position: " & index)
                End If

                'Increment the index
                temp_index += 1
            Loop

            'Check if the currently iterated character is a double-quote
            If temp_index < source.Length AndAlso source(temp_index).Equals(double_quote) Then
                'Increment the index
                temp_index += 1

                'Create the String
                value = New XElement("string", source.Substring(index + 1, temp_index - index - 2))

                'Set the new index
                index = temp_index
            End If
        End If

        Return value
    End Function

    Private Function Parse_Number(ByVal source As String, ByRef index As Integer) As XElement
        'Get the current culture information
        Dim culture As Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture

        'Declare a temporary placeholder
        Dim temp_index As Integer = index

        'Check for the optional unary operator
        If source.IndexOf(culture.NumberFormat.NegativeSign, temp_index, StringComparison.OrdinalIgnoreCase) = temp_index OrElse
            source.IndexOf(culture.NumberFormat.PositiveSign, temp_index, StringComparison.OrdinalIgnoreCase) = temp_index Then
            temp_index += 1
        End If

        'Match one or more digits
        If temp_index < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index).ToString()) <> -1 Then
            Do While temp_index < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index).ToString()) <> -1
                temp_index += 1
            Loop
        Else
            Return Nothing
        End If

        'Optionally match a decimal separator followed by one or more digits
        If temp_index + 1 < source.Length AndAlso
            source.IndexOf(culture.NumberFormat.NumberDecimalSeparator, temp_index, StringComparison.OrdinalIgnoreCase) = temp_index AndAlso
            Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index + 1).ToString()) <> -1 Then

            temp_index += 1
            Do While temp_index < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index).ToString()) <> -1
                temp_index += 1
            Loop
        End If

        'Optionally match an exponent, followed by an optional unary operator, followed by 1 or more digits
        If temp_index + 1 < source.Length AndAlso
            source.IndexOf("e", temp_index, StringComparison.OrdinalIgnoreCase) = temp_index Then

            If temp_index + 2 < source.Length AndAlso
                (source.IndexOf(culture.NumberFormat.NegativeSign, temp_index + 1, StringComparison.OrdinalIgnoreCase) = temp_index + 1 OrElse
                source.IndexOf(culture.NumberFormat.PositiveSign, temp_index + 1, StringComparison.OrdinalIgnoreCase) = temp_index + 1) AndAlso
                Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index + 2).ToString()) <> -1 Then
                temp_index += 2
            ElseIf temp_index + 1 < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index + 1).ToString()) <> -1 Then
                temp_index += 1
            End If

            Do While temp_index < source.Length AndAlso Array.IndexOf(culture.NumberFormat.NativeDigits, source(temp_index).ToString()) <> -1
                temp_index += 1
            Loop
        End If

        'Create the number
        Dim value As XElement = New XElement("number", source.Substring(index, temp_index - index))
        index = temp_index

        Return value
    End Function

    Private Function Parse_Boolean(ByVal source As String, ByRef index As Integer) As XElement
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

    Private Function Parse_Null(ByVal source As String, ByRef index As Integer) As XElement
        Dim value As XElement = Nothing

        'Literally match 'null' in the source starting at the index
        If source.IndexOf("null", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("null")
            index += 4
        End If

        Return value
    End Function

    Private Function SkipWhitespace(ByVal source As String, ByVal index As Integer) As Integer
        Do While index < source.Length AndAlso Char.IsWhiteSpace(source(index))
            index += 1
        Loop

        Return index
    End Function

End Module
