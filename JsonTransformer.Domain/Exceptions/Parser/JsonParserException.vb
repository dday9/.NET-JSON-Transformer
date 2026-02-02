Namespace Exceptions.Lexer

    Public Class JsonParserException
        Inherits Exception

        Public Property Position As Integer

        Sub New(positionAt As Integer)
            MyBase.New()
            Position = positionAt
        End Sub

        Sub New(positionAt As Integer, message As String)
            MyBase.New(message)
            Position = positionAt
        End Sub

        Sub New(positionAt As Integer, message As String, innerException As Exception)
            MyBase.New(message, innerException)
            Position = positionAt
        End Sub

    End Class

End Namespace
