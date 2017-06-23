
Namespace System.DirectoryServices.ActiveDirectory
    Class DirectoryContext

        Private _p1 As Object
        Private _properties As Object

        Sub New(p1 As Object)
            ' TODO: Complete member initialization 
            _p1 = p1
        End Sub

        Property SchemaClassName As Object

        Property Properties(p1 As String) As Object
            Get
                Return _properties
            End Get
            Set(value As Object)
                _properties = value
            End Set
        End Property

    End Class
End Namespace
