Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq
Public Module JSON

    Public Function Parse(ByVal source As String) As XDocument
        Dim d As XDocument = Nothing
        Dim t As XElement = Parse_Value(source)

        If t IsNot Nothing Then
            d = New XDocument(New XDeclaration("1.0", "utf-8", "yes"), t)
        End If

        Return d
    End Function

    Private Function Parse_Object(ByRef source As String) As XElement
        Dim t As XElement = Nothing
        Dim tempSource As String = source.Trim()

        If tempSource.Length >= 2 AndAlso tempSource.StartsWith("{") Then
            tempSource = tempSource.Remove(0, 1).Trim()

            If tempSource.StartsWith("}") Then
                t = New XElement("object")
                source = tempSource.Remove(0, 1).Trim()
            Else
                Dim values As New List(Of XElement)
                Dim key As XElement = Parse_String(tempSource)
                tempSource = tempSource.Trim()

                If tempSource.StartsWith(":") Then
                    tempSource = tempSource.Remove(0, 1).Trim()

                    Dim value As XElement = Parse_Value(tempSource)
                    If value IsNot Nothing Then
                        value.SetAttributeValue("key", key.Value)
                        values.Add(value)
                        tempSource = tempSource.Trim()
                        Do While value IsNot Nothing AndAlso tempSource.Length > 0 AndAlso Not tempSource.StartsWith("}") AndAlso tempSource.StartsWith(",")
                            tempSource = tempSource.Remove(0, 1).Trim()
                            key = Parse_String(tempSource)
                            tempSource = tempSource.Trim()
                            If tempSource.StartsWith(":") Then
                                tempSource = tempSource.Remove(0, 1).Trim()
                                value = Parse_Value(tempSource)
                                If value IsNot Nothing Then
                                    value.SetAttributeValue("key", key.Value)
                                    values.Add(value)
                                    tempSource = tempSource.Trim()
                                End If
                            End If
                        Loop

                        tempSource = tempSource.Trim()
                        If value IsNot Nothing AndAlso tempSource.StartsWith("}") Then
                            tempSource = tempSource.Remove(0, 1).Trim()
                            t = New XElement("object", values)
                            source = New String(tempSource.ToArray())
                        End If
                    End If
                End If
            End If
        End If

        Return t
    End Function

    Private Function Parse_Array(ByRef source As String) As XElement
        Dim t As XElement = Nothing
        Dim tempSource As String = source.Trim()

        If tempSource.Length >= 2 AndAlso tempSource.StartsWith("[") Then
            tempSource = tempSource.Remove(0, 1).Trim()

            If tempSource.StartsWith("]") Then
                t = New XElement("array")
                source = tempSource.Remove(0, 1).Trim()
            Else
                Dim values As New List(Of XElement)
                Dim value As XElement = Parse_Value(tempSource)

                If value IsNot Nothing Then
                    values.Add(value)
                    tempSource = tempSource.Trim()
                    Do While value IsNot Nothing AndAlso tempSource.Length > 0 AndAlso Not tempSource.StartsWith("]") AndAlso tempSource.StartsWith(",")
                        tempSource = tempSource.Remove(0, 1).Trim()
                        value = Parse_Value(tempSource)
                        tempSource = tempSource.Trim()
                        If value IsNot Nothing Then
                            values.Add(value)
                        End If
                    Loop

                    tempSource = tempSource.Trim()
                    If value IsNot Nothing AndAlso tempSource.StartsWith("]") Then
                        tempSource = tempSource.Remove(0, 1).Trim()
                        t = New XElement("array", values)
                        source = New String(tempSource.ToArray())
                    End If
                End If
            End If
        End If

        Return t
    End Function

    Private Function Parse_Value(ByRef source As String) As XElement
        Dim t As XElement = Parse_String(source)

        If t Is Nothing Then
            t = Parse_Number(source)

            If t Is Nothing Then
                t = Parse_Boolean(source)

                If t Is Nothing Then
                    t = Parse_Array(source)

                    If t Is Nothing Then
                        t = Parse_Object(source)
                    End If
                End If
            End If
        End If

        Return t
    End Function

    Private Function Parse_String(ByRef source As String) As XElement
        Dim m As Match = Regex.Match(source, """([^""\\]|\\[""\\\/bfnrt])*""")
        Dim t As XElement = Nothing

        If m.Success AndAlso m.Index = 0 Then
            Dim literal As String = m.Value.Remove(0, 1)
            literal = literal.Remove(literal.Length - 1)
            t = New XElement("string", literal)
            source = source.Substring(m.Value.Length)
        End If

        Return t
    End Function

    Private Function Parse_Number(ByRef source As String) As XElement
        Dim m As Match = Regex.Match(source, "-?\d+((\.\d+)?((e|E)[+-]?\d+)?)")
        Dim t As XElement = Nothing

        If m.Success AndAlso m.Index = 0 Then
            t = New XElement("number", m.Value)
            source = source.Substring(m.Value.Length)
        End If

        Return t
    End Function

    Private Function Parse_Boolean(ByRef source As String) As XElement
        Dim t As XElement = Nothing

        If source.StartsWith("true", StringComparison.OrdinalIgnoreCase) OrElse source.StartsWith("null", StringComparison.OrdinalIgnoreCase) Then
            t = New XElement("boolean", source.Substring(0, 4))
            source = source.Substring(4)
        ElseIf source.StartsWith("false", StringComparison.OrdinalIgnoreCase) Then
            t = New XElement("boolean", source.Substring(0, 5))
            source = source.Substring(5)
        End If

        Return t
    End Function

    Private Function RemoveWhitespace(ByVal value As String) As String
        Do While value.Length > 0 AndAlso String.IsNullOrWhiteSpace(value(0))
            value = value.Remove(0, 1)
        Loop

        Return value
    End Function
End Module