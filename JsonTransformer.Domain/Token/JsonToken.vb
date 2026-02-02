Namespace Token

    Public Class JsonToken
        Public ReadOnly Type As JsonTokenType
        Public ReadOnly Value As String
        Public ReadOnly Position As Integer

        Public Sub New(type As JsonTokenType, value As String, position As Integer)
            Me.Type = type
            Me.Value = value
            Me.Position = position
        End Sub
    End Class

End Namespace
