Imports System.Text

Public Class NetCamDevice

  Protected _id As Integer
  Protected _device_id As String = String.Empty
  Protected _name As String = String.Empty
  Protected _uri As String = String.Empty
  Protected _status As String = String.Empty
  Protected _enabled As Boolean = False

  Protected SighthoundRules As New List(Of SighthoundRule)

#Region "NetCam Object"

  Public Property Id() As Integer
    Get
      Return Me._id
    End Get
    Set(value As Integer)
      Me._id = value
      Me._device_id = Id.ToString.PadLeft(3, "0")
    End Set
  End Property

  Public ReadOnly Property DeviceId() As String
    Get
      Return Me._device_id
    End Get
  End Property

  Public Property Name() As String
    Get
      Return Me._name
    End Get
    Set(value As String)
      Me._name = value
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

  Public Property Enabled() As Boolean
    Get
      Return Me._enabled
    End Get
    Set(value As Boolean)
      Me._enabled = value
    End Set
  End Property

  Public Property Uri() As String
    Get
      Return Me._uri
    End Get
    Set(value As String)
      Me._uri = value
    End Set
  End Property

  Public ReadOnly Property dv_addr() As String
    Get
      Return String.Format("{0}{1}-Camera", "Sighthound", _device_id)
    End Get
  End Property

  Public ReadOnly Property dv_addr_status() As String
    Get
      Return String.Format("{0}{1}-Status", "Sighthound", _device_id)
    End Get
  End Property

  Public Property Rules() As List(Of SighthoundRule)
    Get
      Return Me.SighthoundRules
    End Get
    Set(value As List(Of SighthoundRule))
      Me.SighthoundRules = value
    End Set
  End Property

  Public ReadOnly Property RulesCount() As Integer
    Get
      Return Me.SighthoundRules.Count
    End Get
  End Property

  Public ReadOnly Property RulesEnabled() As Integer
    Get
      Dim Enabled As Integer = 0
      For Each SighthoundRule In Me.SighthoundRules
        If SighthoundRule.Enabled = True Then Enabled += 1
      Next
      Return Enabled
    End Get
  End Property

  Public ReadOnly Property RulesSummary() As String
    Get
      Dim sb As New StringBuilder
      Dim index As Integer = 1
      For Each SighthoundRule In Me.SighthoundRules
        sb.AppendFormat("{0})  Rule='{1}', Action='{2}', Enabled='{3}', Status='{4}'", index.ToString, SighthoundRule.Name, SighthoundRule.Action, SighthoundRule.Enabled, SighthoundRule.Status)
        sb.AppendLine("")
        index += 1
      Next
      Return sb.ToString
    End Get
  End Property

#End Region

  Public Sub New(ByVal Id As Integer, ByVal Name As String, ByVal Uri As String, ByVal Enabled As Boolean, ByVal Status As String)

    Me.Id = Id
    _name = Name
    _uri = Uri
    _enabled = Enabled
    _status = Status

  End Sub

End Class
