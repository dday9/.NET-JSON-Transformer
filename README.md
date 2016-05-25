# .NET-JSON-Transformer
A Visual Basic .NET implementation that converts JSON Strings into a .NET XDocument

To use the file simply add it to your project and call the JSON.Parse method.

##Syntax
`Public Function Parse(ByVal source As String) As XDocument`


**Parameters**
- *source*
  - Type: System.String
  - The JSON literal to be converted.

- Return Value
  - Type: System.Xml.Linq.XDocument
  - An XML representation of the JSON literal. If the JSON is not parsable, then the method returns Nothing.
  
##Remarks
  1. The parser ignores whitespace. So if the JSON literal is:
    ``` json
{
    "string key": [1, 2, {
        "nested": true
    }]
}
    ```
  Then it gets parsed as:
  `{"string key":[1,2,{"nested":true}]}`
  
  2. The returned XML is a 1-to-1 translation from the JSON. Using the same example as above, the resulting XML would be:
    ``` xml
<object>
  <array key="string key">
    <number>1</number>
    <number>2</number>
    <object>
      <boolean key="nested">true</boolean>
    </object>
  </array>
</object>
    ```

##Examples:
  The following example demonstrates the Parse method.
  
  ``` vb.net
  Option Strict On
  Public Module Module1
      Public Sub main()
          Dim jsonLiteral As String = String.Format("{0}{1}key1{1}: [1, -2.12345E6, [true, false]], {1}key2{1}: {0}{1}nested object{1}: null{2}{2}", "{", """", "}")
          Console.WriteLine(jsonLiteral)
          '{"key1": [1, -2.12345E6, [true, false]], "key2": {"nested object": null}}
  
          Dim jsonDocument As XDocument = JSON.Parse(jsonLiteral)
          Console.WriteLine(jsonDocument.ToString)
          '<object>
          '  <array key = "key1" >
          '    <number>1</number>
          '    <number>-2.12345E6</number>
          '    <array>
          '      <boolean>true</Boolean>
          '      <boolean>false</Boolean>
          '    </array>
          '  </array>
          '  <object key = "key2" >
          '    <boolean key="nested object">null</boolean>
          '  </object>
          '</object>
  
          Console.ReadLine()
      End Sub
  
  End Module
  ```
