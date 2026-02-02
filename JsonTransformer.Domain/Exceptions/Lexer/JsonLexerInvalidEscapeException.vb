Namespace Exceptions.Lexer

    Public Class JsonLexerInvalidEscapeException
        Inherits JsonLexerException

        Sub New()
            MyBase.New()
        End Sub

        Sub New(message As String)
            MyBase.New(message)
        End Sub

        Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

    End Class

End Namespace
