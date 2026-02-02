Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace JsonTransformer.Test

    <TestClass>
    Public Class JsonParserTests

        <TestMethod>
        Public Sub TestString()
            Dim source As String = """hello"""
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("string", root.Name.LocalName)
            Assert.AreEqual("hello", root.Value)
        End Sub

        <TestMethod>
        Public Sub TestStringWithEscapes()
            Dim source As String = """\""\\\/\b\f\n\r\t\u0041"""
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Dim expected As String = """" & "\" & "/" & ChrW(8) & ChrW(12) & ChrW(10) & ChrW(13) & ChrW(9) & "A"
            Assert.AreEqual("string", root.Name.LocalName)
            Assert.AreEqual(expected, root.Value)
        End Sub

        <TestMethod>
        Public Sub TestNumber()
            Dim source As String = "123"
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("number", root.Name.LocalName)
            Assert.AreEqual("123", root.Value)

            source = "123.45"
            results = Json.JsonTransformer.Parse(source)
            root = results.Root
            Assert.AreEqual("number", root.Name.LocalName)
            Assert.AreEqual("123.45", root.Value)

            source = "123.45E-2"
            results = Json.JsonTransformer.Parse(source)
            root = results.Root
            Assert.AreEqual("number", root.Name.LocalName)
            Assert.AreEqual("123.45E-2", root.Value)
        End Sub

        <TestMethod>
        Public Sub TestNumberWithCulture()
            Dim culture As New Globalization.CultureInfo("de-DE")
            Dim source As String = "123,45"
            Dim results As XDocument = Json.JsonTransformer.Parse(source, culture)
            Dim root As XElement = results.Root

            Assert.AreEqual("number", root.Name.LocalName)
            Assert.AreEqual("123,45", root.Value)
        End Sub

        <TestMethod>
        Public Sub TestNumberWithCustomSeparators()
            Dim decimalSeparator As String = "|"
            Dim negativeSign As String = "~"
            Dim positiveSign As String = "^"
            Dim source As String = "~123|45E^2"
            Dim results As XDocument = Json.JsonTransformer.Parse(
                source,
                decimalSeparator,
                negativeSign,
                positiveSign
            )
            Dim root As XElement = results.Root
            Assert.AreEqual("number", root.Name.LocalName)
            Assert.AreEqual("~123|45E^2", root.Value)
        End Sub

        <TestMethod>
        Public Sub TestBoolean()
            Dim source As String = "true"
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("boolean", root.Name.LocalName)
            Assert.IsTrue(Boolean.Parse(root.Value))

            source = "false"
            results = Json.JsonTransformer.Parse(source)
            root = results.Root
            Assert.AreEqual("boolean", root.Name.LocalName)
            Assert.IsFalse(Boolean.Parse(root.Value))
        End Sub

        <TestMethod>
        Public Sub TestNull()
            Dim source As String = "null"
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("null", root.Name.LocalName)
            Assert.IsEmpty(root.Value)
        End Sub

        <TestMethod>
        Public Sub TestArray()
            Dim source As String = "[1, ""two"", false]"
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("array", root.Name.LocalName)
            Assert.AreEqual(3, root.Elements().Count())

            Assert.AreEqual("number", root.Elements()(0).Name.LocalName)
            Assert.AreEqual("1", root.Elements()(0).Value)

            Assert.AreEqual("string", root.Elements()(1).Name.LocalName)
            Assert.AreEqual("two", root.Elements()(1).Value)

            Assert.AreEqual("boolean", root.Elements()(2).Name.LocalName)
            Assert.IsFalse(Boolean.Parse(root.Elements()(2).Value))
        End Sub

        <TestMethod>
        Public Sub TestObject()
            Dim source As String = "{""a"":1,""b"":""two"",""c"":false}"
            Dim results As XDocument = Json.JsonTransformer.Parse(source)
            Dim root As XElement = results.Root
            Assert.AreEqual("object", root.Name.LocalName)

            Dim items = root.Elements("item").ToList()
            Assert.HasCount(3, items)

            Assert.AreEqual("a", items(0).Element("key").Value)
            Assert.AreEqual("number", items(0).Element("value").Elements().First().Name.LocalName)
            Assert.AreEqual("1", items(0).Element("value").Value)

            Assert.AreEqual("b", items(1).Element("key").Value)
            Assert.AreEqual("string", items(1).Element("value").Elements().First().Name.LocalName)
            Assert.AreEqual("two", items(1).Element("value").Value)

            Assert.AreEqual("c", items(2).Element("key").Value)
            Assert.AreEqual("boolean", items(2).Element("value").Elements().First().Name.LocalName)
            Assert.IsFalse(Boolean.Parse(items(2).Element("value").Value))
        End Sub

        <TestMethod>
        Public Sub TestKitchenSink()
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
            Dim root As XElement = results.Root

            ' root object
            Assert.AreEqual("object", root.Name.LocalName)
            Assert.AreEqual(1, root.Elements("item").Count())

            ' glossary
            Dim glossaryItem = root.Elements("item").First()
            Dim glossaryItemValue = glossaryItem.
                Element("key").
                Value
            Assert.AreEqual("glossary", glossaryItemValue)

            Dim glossaryObject = glossaryItem.
                Element("value").
                Element("object")
            Assert.IsNotNull(glossaryObject)

            ' glossary.title
            Dim titleItem = glossaryObject.Elements("item").
                First(Function(i) i.Element("key").Value = "title")
            Dim titleItemValue = titleItem.
                Element("value").
                Value
            Assert.AreEqual("example glossary", titleItem.Element("value").Value)

            ' GlossDiv
            Dim glossDivItem = glossaryObject.Elements("item").
                First(Function(i) i.Element("key").Value = "GlossDiv")

            Dim glossDivObject = glossDivItem.
                Element("value").
                Element("object")
            Assert.IsNotNull(glossDivObject)

            ' GlossDiv.title
            Dim glossDivTitle = glossDivObject.Elements("item").
                First(Function(i) i.Element("key").Value = "title")
            Dim glossDivTitleValue = glossDivTitle.
                Element("value").
                Value

            Assert.AreEqual("S", glossDivTitleValue)

            ' GlossList → GlossEntry
            Dim glossList = glossDivObject.Elements("item").
                First(Function(i) i.Element("key").Value = "GlossList").
                Element("value").
                Element("object")

            Dim glossEntry = glossList.Elements("item").
                First(Function(i) i.Element("key").Value = "GlossEntry").
                Element("value").
                Element("object")

            ' ID
            Dim id = glossEntry.Elements("item").
                First(Function(i) i.Element("key").Value = "ID").
                Element("value").
                Value
            Assert.AreEqual("SGML", id)

            ' GlossDef
            Dim glossDef = glossEntry.Elements("item").
                First(Function(i) i.Element("key").Value = "GlossDef").
                Element("value").
                Element("object")

            Dim glossDefValue = glossDef.Elements("item").
                First(Function(i) i.Element("key").Value = "para").
                Element("value").
                Value

            ' GlossDef.para
            Assert.AreEqual("A meta-markup language, used to create markup languages such as DocBook.", glossDefValue)

            ' GlossDef.GlossSeeAlso (array)
            Dim seeAlsoArray = glossDef.Elements("item").
                First(Function(i) i.Element("key").Value = "GlossSeeAlso").
                Element("value").
                Element("array")
            Dim seeAlsoValue1 = seeAlsoArray.Elements()(0).Value
            Dim seeAlsoValue2 = seeAlsoArray.Elements()(1).Value

            Assert.AreEqual(2, seeAlsoArray.Elements().Count())
            Assert.AreEqual("GML", seeAlsoValue1)
            Assert.AreEqual("XML", seeAlsoValue2)
        End Sub

    End Class

End Namespace
