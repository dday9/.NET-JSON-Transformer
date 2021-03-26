
# .NET-JSON-Transformer
A Visual Basic .NET (VB.NET) implementation that converts a JSON literal into a .NET XDocument

## Add to project
The [json.vb](json.vb) code file is uncompiled. Follow these instructions to add the file to your project:

 1. Project > Add Existing Item (shift + alt + a)
 2. Select the file in the
    browse dialog
 3. Add

## Json.Parse Method
Creates a new [XDocument](https://docs.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument) from a JSON literal

``` vb
Public Shared Function Parse(ByVal source As String) As XDocument
```

### Parameters
`source` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)
A string that contains JSON.

### Returns
[XDocument](https://docs.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument)
An [XDocument](https://docs.microsoft.com/en-us/dotnet/api/system.xml.linq.xdocument) populated from the string that contains JSON.

### Example
The following example illustrates the Json.Parse method:
``` vb
Const literal = "{ ""Property1"": 1, ""Property2"": false }"
Dim parsedJson = Json.Parse(literal)
Console.WriteLine(parsedJson.ToString())

' <object>
'   <item>
'     <key>Property1</key>
'     <value>
'       <number>1</number>
'     </value>
'   </item>
'   <item>
'     <key>Property2</key>
'     <value>
'       <boolean>false</boolean>
'     </value>
'   </item>
' </object>
```

## Remarks
  * The parser ignores whitespace, essentially minfying the JSON. For example, if the JSON literal is:
    ``` json
    {
        "glossary": {
            "title": "example glossary",
            "GlossDiv": {
                "title": "S",
                "GlossList": {
                    "GlossEntry": {
                        "ID": "SGML",
                        "SortAs": "SGML",
                        "GlossTerm": "Standard Generalized Markup Language",
                        "Acronym": "SGML",
                        "Abbrev": "ISO 8879:1986",
                        "GlossDef": {
                            "para": "A meta-markup language, used to create markup languages such as DocBook.",
                            "GlossSeeAlso": ["GML", "XML"]
                        },
                        "GlossSee": "markup"
                    }
                }
            }
        }
    }
    ```
  Then it gets parsed as:
  ``` json
  {"glossary":{"title":"example glossary","GlossDiv":{"title":"S","GlossList":{"GlossEntry":{"ID":"SGML","SortAs":"SGML","GlossTerm":"Standard Generalized Markup Language","Acronym":"SGML","Abbrev":"ISO 8879:1986","GlossDef":{"para":"A meta-markup language, used to create markup languages such as DocBook.","GlossSeeAlso":["GML","XML"]},"GlossSee":"markup"}}}}}
  ```
  
  * The returned XML is a 1-to-1 translation from the JSON. Using the same example as above, the resulting XML would be:
    ```
    <object>
      <item>
        <key>glossary</key>
        <value>
          <object>
            <item>
              <key>title</key>
              <value>
                <string>example glossary</string>
              </value>
            </item>
            <item>
              <key>GlossDiv</key>
              <value>
                <object>
                  <item>
                    <key>title</key>
                    <value>
                      <string>S</string>
                    </value>
                  </item>
                  <item>
                    <key>GlossList</key>
                    <value>
                      <object>
                        <item>
                          <key>GlossEntry</key>
                          <value>
                            <object>
                              <item>
                                <key>ID</key>
                                <value>
                                  <string>SGML</string>
                                </value>
                              </item>
                              <item>
                                <key>SortAs</key>
                                <value>
                                  <string>SGML</string>
                                </value>
                              </item>
                              <item>
                                <key>GlossTerm</key>
                                <value>
                                  <string>Standard Generalized Markup Language</string>
                                </value>
                              </item>
                              <item>
                                <key>Acronym</key>
                                <value>
                                  <string>SGML</string>
                                </value>
                              </item>
                              <item>
                                <key>Abbrev</key>
                                <value>
                                  <string>ISO 8879:1986</string>
                                </value>
                              </item>
                              <item>
                                <key>GlossDef</key>
                                <value>
                                  <object>
                                    <item>
                                      <key>para</key>
                                      <value>
                                        <string>A meta-markup language, used to create markup languages such as DocBook.</string>
                                      </value>
                                    </item>
                                    <item>
                                      <key>GlossSeeAlso</key>
                                      <value>
                                        <array>
                                          <string>GML</string>
                                          <string>XML</string>
                                        </array>
                                      </value>
                                    </item>
                                  </object>
                                </value>
                              </item>
                              <item>
                                <key>GlossSee</key>
                                <value>
                                  <string>markup</string>
                                </value>
                              </item>
                            </object>
                          </value>
                        </item>
                      </object>
                    </value>
                  </item>
                </object>
              </value>
            </item>
          </object>
        </value>
      </item>
    </object>
    ```

  * The parser does not parse to the exact specifications of the EBNF found on http://www.json.org/ the following list the deviations in this parser:
    * Number: Checks if the number starts with either a positive sign or negative sign
    * Boolean: Checks for "true" or "false" based on case-insensitivity
    * Null: Checks for "null" based on case-insensitivity
    * String: Does not check for "\u" followed by 4 hexadecimal characters as an escape character

## Examples and Demo
  The following example demonstrates the Parse method.
  
  ``` vb.net
 Public Module Module1

    Public Sub Main()
        Const literal = "{
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
        Dim parsedJson = Json.Parse(literal)
        Console.WriteLine(parsedJson.ToString())
        Console.ReadLine()
    End Sub

End Module
  ```
Fiddle: https://dotnetfiddle.net/bYcMYm
