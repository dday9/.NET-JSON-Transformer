Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq
Public Module JSON

	''' <summary>Converts a JSON literal into an XDocument.</summary>
	''' <param name="literal">The JSON literal to convert.</param>
	''' <value>XDocument</value>
	''' <returns>XDocument of parsed JSON literal.</returns>
	''' <remarks>Returns Nothing if the conversion fails.</remarks>
	Public Function Parse(ByVal literal As String) As XDocument
		If String.IsNullOrWhitespace(literal) Then Return Nothing

		'Declare a document to return
		Dim document As XDocument = Nothing

		'Declare a value that will make up the document
		Dim value As XElement = Parse_Value(literal, 0)
		If value IsNot Nothing Then
			document = New XDocument(New XDeclaration("1.0", "utf-8", "yes"), value)
		End If

		Return document
	End Function

	Private Function Parse_Value(ByVal source As String, ByRef index As Integer) As XElement
		'Skip any whitespace
		index = JSON.SkipWhitespace(source, index)

		'Go through each available value until one returns something that ain't nothing
		Dim node As XElement = JSON.Parse_Null(source, index)
		If node Is Nothing Then
			node = JSON.Parse_Boolean(source, index)
			If node Is Nothing Then
				node = JSON.Parse_Number(source, index)
				If node Is Nothing Then
					node = JSON.Parse_String(source, index)
					If node Is Nothing Then
						node = JSON.Parse_Array(source, index)
						If node Is Nothing Then
							node = JSON.Parse_Object(source, index)
						End If
					End If
				End If
			End If
		End If

		Return node
	End Function

	Private Function Parse_Object(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

		'Match the opening curly bracket
		If source(index) = "{"c Then	
			'Declare collections that will make up the node's key:value
			Dim keys,values As New List(Of XElement)

			'Declare a temporary index placeholder in case the parsing fails
			Dim tempIndex As Integer = JSON.SkipWhitespace(source, index + 1)

			'Declare a String which would represent the key of key/value pair
			Dim key As XElement = Nothing

			'Declare an XElement which would represent the value of the key/value pair
			Dim item As XElement = Nothing

			'Loop until there is a closing curly bracket
			Do While tempIndex < source.Length AndAlso source(tempIndex) <> "}"c
				'Match a String which should be the key
				key = Parse_String(source, tempIndex)

				'Ensure that there was a valid key
				If key IsNot Nothing Then
					'Add the item to the collection that will ultimately represent the key of the key/value pair
					keys.Add(key)

					'Skip any whitespace
					tempIndex = JSON.SkipWhitespace(source, tempIndex)

					'Ensure a separator
					If source(tempIndex) = ":"c Then
						'Skip any whitespace
						tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)

						item = Parse_Value(source, tempIndex)

						'Ensure that there was a valid item
						If item IsNot Nothing Then
							'Add the item to the collection that will ultimately represent the value of the key/value pair
							values.Add(item)

							'Skip any whitespace
							tempIndex = JSON.SkipWhitespace(source, tempIndex)

							'Ensure a separator or ending bracket
							If source(tempIndex) = ","c Then
								tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)
							ElseIf source(tempIndex) <> ","c AndAlso source(tempIndex) <> "}"c Then
								Throw New Exception("Unexpected token at position: " & tempIndex + 1 & ". Expected a comma to separate Object items.")
							End If
						Else
							Throw New Exception("Invalid item in array at position: " & tempIndex + 1)
						End If
					Else
						Throw New Exception("Unexpected token at position: " & tempIndex + 1 & ". Expected a colon to separate Array items.")
					End If
				Else
					Throw New Exception("Unexpected token at position: " & tempIndex + 1 & ". Expected a String to represent the key/value pair of an Object.")
				End If
			Loop

			'Valid parse
			If tempIndex < source.Length AndAlso source(tempIndex) = "}"c Then
				node = New XElement("object")
				For i As Integer = 0 To keys.Count - 1
					node.Add(New XElement(keys.Item(i).Value, values.Item(i)))
				Next
				index = tempIndex + 1
			End If
		End If
		
		Return node
	End Function

	Private Function Parse_Array(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

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
				item = Parse_Value(source, tempIndex)

				'Ensure that there was a valid item
				If item IsNot Nothing Then
					'Add the item to the collection that will ultimately represent the node's value
					nodes.Add(item)

					'Skip any whitespace
					tempIndex = JSON.SkipWhitespace(source, tempIndex)

					'Ensure a separator or ending bracket
					If source(tempIndex) = ","c Then
						tempIndex = JSON.SkipWhitespace(source, tempIndex + 1)
					ElseIf source(tempIndex) <> ","c AndAlso source(tempIndex) <> "]"c Then
						Throw New Exception("Unexpected token at position: " & tempIndex + 1 & ". Expected a comma to separate Array items.")
					End If
				Else
					Throw New Exception("Invalid item in array at position: " & tempIndex + 1)
				End If
			Loop

			'Valid parse
			If tempIndex < source.Length AndAlso source(tempIndex) = "]"c Then
				node = New XElement("array", nodes)
				index = tempIndex + 1
			End If
		End If

		Return node
	End Function

	Private Function Parse_String(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

		'The pattern to match a number is:
		'Double-Quote
		'Any unicode character except for (\ or ") or an escaped character zero or more times
		'Double-Quote
		Dim pattern As Regex = New Regex("""([^""\\]|\\[""\\\/bfnrt])*""", RegexOptions.IgnoreCase)
		Dim m As Match = pattern.Match(source, index)

		'A match will only be valid if it matches at the beginning of the string
		If m.Success AndAlso m.Index = index Then
			node = New XElement("string", m.Value.Substring(1, m.Value.Length - 2))
			index += m.Value.Length
		End If

		Return node
	End Function

	Private Function Parse_Number(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

		'The pattern to match a number is:
		'Optional unary negative
		'0 or 1-9 followed by zero or more digits
		'Optional mantissa followed by one or more digits
		'Optional euler's number followed by optional unary operator followed by one or more digits
		Dim pattern As Regex = New Regex("-?([1-9]\d*|0)(.\d+)?([eE][-+]?\d+)?")
		Dim m As Match = pattern.Match(source, index)

		'A match will only be valid if it matches at the beginning of the string
		If m.Success AndAlso m.Index = index Then
			node = New XElement("number", m.Value)
			index += m.Value.Length
		End If

		Return Node
	End Function

	Private Function Parse_Boolean(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

		'Literally match 'true' or 'false'
		Dim startSubstring As String = source.Substring(index)

		If startSubstring.IndexOf("true", StringComparison.OrdinalIgnoreCase) = 0 Then
			node = New XElement("boolean", startSubstring.Substring(0, 4))
			index += 4
		ElseIf startSubstring.IndexOf("false", StringComparison.OrdinalIgnoreCase) = 0 Then
			node = New XElement("boolean", startSubstring.Substring(0, 5))
			index += 5
		End If

		Return node
	End Function

	Private Function Parse_Null(ByVal source As String, ByRef index As Integer) As XElement
		'Declare a value to return (default is Nothing)
		Dim node As XElement = Nothing

		'Literally match 'null'
		If source.Substring(index).IndexOf("null", StringComparison.OrdinalIgnoreCase) = 0 Then
			node = New XElement("null")
			index += 4
		End If

		Return node
	End Function

	Private Function SkipWhitespace(ByVal source As String, ByVal index As Integer) As Integer
		Do While index < source.Length AndAlso Char.IsWhiteSpace(source(index))
			index += 1
		Loop
		
		Return index
	End Function
End Module
