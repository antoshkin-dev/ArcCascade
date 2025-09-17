Public Class Params
    Private _IsSet As Boolean
    Private _Value As String
    Private _Name As String
    Private _Reqired As Boolean
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property
    Public Property Reqired As Boolean
        Get
            Return _Reqired
        End Get
        Set(value As Boolean)
            _Reqired = value
        End Set
    End Property
    Public ReadOnly Property IsSet As Boolean
        Get
            Return _IsSet
        End Get
    End Property
    Public Property Value As String
        Get
            Return _Value
        End Get
        Set(value As String)
            _Value = value
            _IsSet = True
        End Set
    End Property
    Public Sub New(Name As String, Reqired As Boolean)
        _Name = Name
        _Reqired = Reqired
        _IsSet = False
    End Sub
    Public Sub New(Name As String, Reqired As Boolean, Value As Object)
        _Name = Name
        _Reqired = Reqired
        _IsSet = False
        _Value = Value
    End Sub
End Class
