Public Class SighthoundCameraStatus

  Private _status As String = "failed"
  Private _enabled As Boolean = False

  Public Property Status() As String
    Get
      Return Me._status
    End Get
    Set(value As String)
      Me._status = value
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

End Class
