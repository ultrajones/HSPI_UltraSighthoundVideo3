Public Class SighthoundRule

  Protected _id As Integer = 0
  Protected _name As String = String.Empty
  Protected _desc As String = String.Empty
  Protected _status As String = String.Empty
  Protected _action As String = String.Empty
  Protected _enabled As Boolean = False

#Region "SighthoundRule Object"

  Public Property Id() As Integer
    Get
      Return Me._id
    End Get
    Set(value As Integer)
      Me._id = value
    End Set
  End Property

  Public Property Name() As String
    Get
      Return Me._name
    End Get
    Set(value As String)
      Me._name = value
    End Set
  End Property

  Public Property Description() As String
    Get
      Return Me._desc
    End Get
    Set(value As String)
      Me._desc = value
    End Set
  End Property

  Public Property Status() As String
    Get
      Return Me._status
    End Get
    Set(value As String)
      Me._status = value
    End Set
  End Property

  Public Property Action() As String
    Get
      Return Me._action
    End Get
    Set(value As String)
      Me._action = value
    End Set
  End Property

  Public Property Enabled() As Boolean
    Get
      Return Me._enabled
    End Get
    Set(value As Boolean)
      Me._enabled = value
    End Set
  End Property

#End Region

End Class
