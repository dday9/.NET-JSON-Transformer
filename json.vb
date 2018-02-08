Public Module JSON

    ''' <summary>Converts a JSON literal into an XDocument.</summary>
    ''' <param name="literal">The JSON literal to convert.</param>
    ''' <returns>XDocument of parsed JSON literal.</returns>
    ''' <remarks>Returns Nothing if the conversion fails.</remarks>
    Public Function Parse(ByVal literal As String) As XDocument
        If String.IsNullOrWhiteSpace(literal) Then Return Nothing

        'Declare a document to return
        Dim document As XDocument = Nothing

        'Declare a value that will make up the document
        Dim value As XElement = Nothing
        If Parse_Value(literal, 0, value) Then
            document = New XDocument(New XDeclaration("1.0", "utf-8", "yes"), value)
        End If

        Return document
    End Function

    Private Function Parse_Value(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'Skip any whitespace
        index = JSON.SkipWhitespace(source, index)

        'Go through each available value until one returns something that ain't nothing
        If Not Parse_String(source, index, value) AndAlso
                Not Parse_Number(source, index, value) AndAlso
                Not Parse_Object(source, index, value) AndAlso
                Not Parse_Array(source, index, value) AndAlso
                Not Parse_Boolean(source, index, value) AndAlso
                Not Parse_Null(source, index, value) Then
        End If

        Return value IsNot Nothing
    End Function

    Private Function Parse_Object(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'Match the opening curly bracket
        If source(index) = "{"c Then
            'Declare collections that will make up the node's key:value
            Dim keys, values As New List(Of XElement)

            'Declare a temporary index placeholder in case the parsing fails
            Dim tempIndex As Integer = JSON.SkipWhitespace(source, index + 1)

            'Declare a String which would represent the key of key/value pair
            Dim key As XElement = Nothing

            'Declare an XElement which would represent the value of the key/value pair
            Dim item As XElement = Nothing

            'Loop until there is a closing curly bracket
            Do While tempIndex < source.Length AndAlso source(tempIndex) <> "}"c
                'Reset the key
                key = Nothing

                'Match a String which should be the key
                If Parse_String(source, tempIndex, key) Then
                    'Add the item to the collection that will ultimately represent the key of the key/value pair
                    keys.Add(key)

                    'Skip any whitespace
                    tempIndex = JSON.SkipWhitespace(source, tempIndex)

                    'Ensure a separator
                    If source(tempIndex) = ":"c Then
                        'Skip any whitespace
                        tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)

                        'Reset the value
                        item = Nothing

                        If Parse_Value(source, tempIndex, item) Then
                            'Add the item to the collection that will ultimately represent the value of the key/value pair
                            values.Add(item)

                            'Skip any whitespace
                            tempIndex = JSON.SkipWhitespace(source, tempIndex)

                            'Ensure a separator or ending bracket
                            If source(tempIndex) = ","c Then
                                tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)
                            ElseIf source(tempIndex) <> "}"c Then
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                End If

            Loop

            'Valid parse
            If tempIndex < source.Length AndAlso source(tempIndex) = "}"c Then
                value = New XElement("object")
                For i As Integer = 0 To keys.Count - 1
                    value.Add(New XElement(keys.Item(i).Value, values.Item(i)))
                Next
                index = tempIndex + 1
            End If
        End If

        Return value IsNot Nothing
    End Function

    Private Function Parse_Array(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'Match the opening square bracket
        If source(index) = "["c Then
            'Declare a collection that will make up the node's value
            Dim nodes As List(Of XElement) = New List(Of XElement)

            'Declare a temporary index placeholder in case the parsing fails
            Dim tempIndex As Integer = JSON.SkipWhitespace(source, index + 1)

            'Declare an XElement which would represent an item in the array
            Dim item As XElement = Nothing

            'Loop until there is a closing square bracket
            Do While tempIndex < source.Length AndAlso source(tempIndex) <> "]"c
                'Reset the value
                item = Nothing

                If Parse_Value(source, tempIndex, item) Then
                    'Add the item to the collection that will ultimately represent the node's value
                    nodes.Add(item)

                    'Skip any whitespace
                    tempIndex = JSON.SkipWhitespace(source, tempIndex)

                    'Ensure a separator or ending bracket
                    If source(tempIndex) = ","c Then
                        tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)
                    ElseIf source(tempIndex) <> "]"c Then
                        Return False
                    End If
                Else
                    Return False
                End If
            Loop

            'Valid parse
            If tempIndex < source.Length AndAlso source(tempIndex) = "]"c Then
                value = New XElement("array", nodes)
                index = tempIndex + 1
            End If
        End If

        Return value IsNot Nothing
    End Function

    Private Function Parse_String(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'The pattern to match a number is:
        'Double-Quote
        'Any unicode character except for (\ or ") or an escaped character zero or more times
        'Double-Quote
        Dim pattern As Regex = New Regex("""([^""\\]|\\[""\\\/bfnrt])*""", RegexOptions.IgnoreCase)
        Dim m As Match = pattern.Match(source, index)

        'A match will only be valid if it matches at the beginning of the string
        If m.Success AndAlso m.Index = index Then
            value = New XElement("string", m.Value.Substring(1, m.Value.Length - 2))
            index += m.Value.Length
        End If

        Return value IsNot Nothing
    End Function

    Private Function Parse_Number(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
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
        If temp_index < source.Length AndAlso Char.IsDigit(source(temp_index)) Then
            Do While temp_index < source.Length AndAlso Char.IsDigit(source(temp_index))
                temp_index += 1
            Loop
        Else
            Return False
        End If

        'Optionally match a mantissa followed by one or more digits
        If temp_index + 1 < source.Length AndAlso
            source.IndexOf(culture.NumberFormat.NumberDecimalSeparator, temp_index, StringComparison.OrdinalIgnoreCase) = temp_index AndAlso
            Char.IsDigit(source(temp_index + 1)) Then

            temp_index += 1
            Do While temp_index < source.Length AndAlso Char.IsDigit(source(temp_index))
                temp_index += 1
            Loop
        End If

        'Optionally match an exponent, followed by an optional unary operator, followed by 1 or more digits
        If temp_index + 1 < source.Length AndAlso
            source.IndexOf("e", temp_index, StringComparison.OrdinalIgnoreCase) = temp_index Then

            If temp_index + 2 < source.Length AndAlso
                (source.IndexOf(culture.NumberFormat.NegativeSign, temp_index + 1, StringComparison.OrdinalIgnoreCase) = temp_index + 1 OrElse
                source.IndexOf(culture.NumberFormat.PositiveSign, temp_index + 1, StringComparison.OrdinalIgnoreCase) = temp_index + 1) AndAlso
                Char.IsDigit(source(temp_index + 2)) Then

                temp_index += 2
                Do While temp_index < source.Length AndAlso Char.IsDigit(source(temp_index))
                    temp_index += 1
                Loop
            ElseIf temp_index + 1 < source.Length AndAlso Char.IsDigit(source(temp_index + 1)) Then
                temp_index += 1
                Do While temp_index < source.Length AndAlso Char.IsDigit(source(temp_index))
                    temp_index += 1
                Loop
            End If
        End If

        'Convert everything up to the index
        value = New XElement("number", source.Substring(index, temp_index - index))
        index = temp_index

        Return value IsNot Nothing
    End Function

    Private Function Parse_Boolean(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'Literally match 'true' or 'false'
        If source.IndexOf("true", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("boolean", True)
            index += 4
        ElseIf source.IndexOf("false", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("boolean", False)
            index += 5
        End If

        Return value IsNot Nothing
    End Function

    Private Function Parse_Null(ByVal source As String, ByRef index As Integer, ByRef value As XElement) As Boolean
        'Literally match 'null' in the source starting at the index
        If source.IndexOf("null", index, StringComparison.OrdinalIgnoreCase) = index Then
            value = New XElement("null")
            index += 4
        End If

        Return value IsNot Nothing
    End Function

    Private Function SkipWhitespace(ByVal source As String, ByVal index As Integer) As Integer
        Do While index < source.Length AndAlso Char.IsWhiteSpace(source(index))
            index += 1
        Loop

        Return index
    End Function
End Module
