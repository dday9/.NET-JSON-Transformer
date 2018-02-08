# .NET-JSON-Transformer
A Visual Basic .NET implementation that converts JSON Strings into a .NET XDocument

To use the file simply add it to your project and call the JSON.Parse method.

## Syntax
`Public Function Parse(ByVal literal As String) As XDocument`

**Parameters**
- *literal*
  - Type: System.String
  - The JSON literal to be converted.

- Return Value
  - Type: System.Xml.Linq.XDocument
  - An XML representation of the JSON literal. If the JSON is not parsable, then the method returns Nothing.
  
## Remarks
  * The parser ignores whitespace. So if the JSON literal is:
    ``` json
    {
      "string_key": [
        1,
        2,
        {
          "nested": true
        }
      ]
    }
    ```
  Then it gets parsed as:
  ``` json
  {"string_key":[1,2,{"nested":true}]}
  ```
  
  * The returned XML is a 1-to-1 translation from the JSON. Using the same example as above, the resulting XML would be:
    ```
    <object>
      <string_key>
        <array>
          <number>1</number>
          <number>2</number>
          <object>
            <nested>
              <boolean>true</boolean>
            </nested>
          </object>
        </array>
      </string_key>
    </object>
    ```

## Examples:
  The following example demonstrates the Parse method.
  
  ``` vb.net
  Option Strict On
  Public Module Module1
	  Public Sub Main()
		  Dim jsonLiteral As String = "{""key"": [1, 2, 3], ""nested"": {""object"": true}}"
		  Console.WriteLine(JSON.Parse(jsonLiteral))
	  End Sub
  End Module
  ```
Fiddle: https://dotnetfiddle.net/Gui2tq
