Imports System.Net
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Text
Imports System.Xml
Imports System.Threading
Imports System.Drawing
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Public Class hspi_sighthound_api

  Private CommandQueue As New Queue
  Private NetCamDevices As New List(Of NetCamDevice)

  Private _pingResponse As String = String.Empty
  Private _pingFlag As Long = 0

  Private _querySuccess As ULong = 0
  Private _queryFailure As ULong = 0

  Private dv_connection As String = String.Concat(IFACE_NAME, "-Connection")
  Private HTTPLock As New Object

  Public Sub New()

    Try

      ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CertificateValidationCallBack)

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Returns the Last Ping Response Status
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPingResponse() As String

    If _pingResponse.Length > 0 Then
      Return "OK"
    Else
      Return "Error"
    End If

  End Function

  ''' <summary>
  ''' Poll Sighthound Video
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub PollSighthoundVideoThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gMonitoring = True Then
          RefreshNetCamList()

          '
          ' Upddate the device connection status
          '
          SetDeviceValue(dv_connection, _pingFlag)
        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Options", "PollSighthoundVideoThread", "5", gINIFile))
        Thread.Sleep(1000 * (iCheckInterval))

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("PollSighthoundVideoThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "PollSighthoundVideoThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Refresh the Network Camera List
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshNetCamList()

    Try

      Dim ping As String = SendPing()

      If ping.Length > 0 Then
        Dim camerasEnabled As Integer = 0

        ' remoteGetCameraNames
        ' remoteGetLiveCameras
        ' remoteGetRulesForCamera

        '
        ' Refresh the Network Cameras
        '
        Dim CameraNames As ArrayList = remoteGetCameraNames()
        For Each cameraName As String In CameraNames
          Try
            Dim SighthoundCameraStatus As SighthoundCameraStatus = getCameraStatusAndEnabled(cameraName)
            Dim cameraUri As String = remoteGetCameraUri(cameraName)
            Dim cameraId As Integer = 0

            If NetCamDevices.Any(Function(c) c.Name = cameraName) = False Then
              '
              ' Add New Camera
              '
              cameraId = GetNetCamDeviceId(cameraName)
              If cameraId = 0 Then
                cameraId = InsertNetCamDevice(cameraName, cameraUri)
              End If

              If cameraId > 0 Then
                Dim NetCamDevice As New NetCamDevice(cameraId, cameraName, cameraUri, SighthoundCameraStatus.Enabled, SighthoundCameraStatus.Status)
                NetCamDevice.Rules = GetNetworkCameraRules(NetCamDevice)
                UpdateCameraEnabledStatus(NetCamDevice)
                UpdateCameraConnectionStatus(NetCamDevice)
                NetCamDevices.Add(NetCamDevice)
              End If

            Else
              '
              ' Update Existing Camera
              '
              Dim NetCamDevice As NetCamDevice = NetCamDevices.Find(Function(s) s.Name = cameraName)

              cameraId = NetCamDevice.Id

              If cameraId > 0 Then

                NetCamDevice.Name = cameraName
                NetCamDevice.Uri = cameraUri
                NetCamDevice.Enabled = SighthoundCameraStatus.Enabled
                NetCamDevice.Status = SighthoundCameraStatus.Status

                NetCamDevice.Rules = GetNetworkCameraRules(NetCamDevice)

                Dim dv_addr As String = NetCamDevice.dv_addr
                Dim dv_value As Integer = IIf(SighthoundCameraStatus.Enabled = True, 1, 0)
                SetDeviceValue(dv_addr, dv_value)

                Dim dv_addr_status As String = NetCamDevice.dv_addr_status
                Dim dv_value_status As Integer = IIf(NetCamDevice.Status = "failed", 2, 1)
                SetDeviceValue(dv_addr_status, dv_value_status)

              End If

            End If

            '
            ' Begin Processing Snapshots
            '
            If cameraId > 0 Then

              Dim strIdentifier As String = String.Format("{0}{1}-Camera", "Sighthound", cameraId.ToString.PadLeft(3, "0"))
              Dim strSnapshotFilename As String = FixPath(String.Format("{0}\{1}_snapshot.jpg", gSnapshotDirectory, strIdentifier))
              Dim strThumbnailFilename As String = FixPath(String.Format("{0}\{1}_thumbnail.jpg", gSnapshotDirectory, strIdentifier))

              If SighthoundCameraStatus.Enabled = False Then

                Try

                  Dim strCameraOff As String = FixPath(String.Format("{0}\html\images\hspi_ultrasighthoundvideo3\{1}", HSAppPath, "cameraOff.jpg"))
                  If System.IO.File.Exists(strCameraOff) = True Then
                    System.IO.File.Copy(strCameraOff, strSnapshotFilename, True)
                    System.IO.File.Copy(strCameraOff, strThumbnailFilename, True)
                  End If

                Catch ex As Exception

                End Try

              ElseIf SighthoundCameraStatus.Status = "failed" Then

                Try

                  Dim strCameraOff As String = FixPath(String.Format("{0}\html\images\hspi_ultrasighthoundvideo3\{1}", HSAppPath, "cameraFailed.jpg"))
                  If System.IO.File.Exists(strCameraOff) = True Then
                    System.IO.File.Copy(strCameraOff, strSnapshotFilename, True)
                    System.IO.File.Copy(strCameraOff, strThumbnailFilename, True)
                  End If

                Catch ex As Exception

                End Try

              Else

                camerasEnabled += 1

                If cameraId > 0 And cameraUri.Length > 0 Then
                  Dim WebExceptionStatus As WebExceptionStatus = GetCameraSnapshot(cameraId, cameraUri, strSnapshotFilename, strThumbnailFilename, 10)
                  Select Case WebExceptionStatus
                    Case WebExceptionStatus.Success
                    Case Else
                      Dim strErrorMessage As String = String.Format("The {0} snapshot request sent to {1} failed: {2}", cameraName, gSighthoundURL, WebExceptionStatus.ToString)
                      WriteMessage(strErrorMessage, MessageType.Warning)
                  End Select

                End If

              End If

            End If

          Catch pEx As Exception
            '
            ' Catch Program Exception
            '
            Call ProcessError(pEx, "RefreshNetCamList()")
          End Try

        Next

        '
        ' Update Cameras Enabled
        '
        UpdateSighthoundArming(camerasEnabled, CameraNames.Count)

      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "RefreshNetCamList()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets the Camera Name by Device Id
  ''' </summary>
  ''' <param name="SighthoundCamera"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraNameById(ByVal SighthoundCamera As String)

    Try
      Dim DeviceId As String = Regex.Match(SighthoundCamera, "\d\d\d").Value
      Dim NetCamDevice As NetCamDevice = NetCamDevices.Find(Function(s) s.DeviceId = DeviceId)
      Return NetCamDevice.Name

    Catch pEx As Exception
      Return SighthoundCamera
    End Try

  End Function

  ''' <summary>
  ''' Updates the Camera Enabled Status
  ''' </summary>
  ''' <param name="NetCamDevice"></param>
  ''' <remarks></remarks>
  Private Sub UpdateCameraEnabledStatus(ByRef NetCamDevice As NetCamDevice)

    Try

      Dim dv_root_addr As String = String.Format("{0}{1}-Root", "Sighthound", NetCamDevice.DeviceId)
      Dim dv_root_type As String = "Sighthound Camera"
      Dim dv_root_name As String = "Sighthound Camera"

      Dim dv_addr As String = NetCamDevice.dv_addr
      Dim dv_name As String = NetCamDevice.Name
      Dim dv_type As String = "Sighthound Camera"

      GetHomeSeerDevice(dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
      Dim dv_value As Integer = IIf(NetCamDevice.Enabled = True, 1, 0)
      SetDeviceValue(dv_addr, dv_value)

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Updates the Camera Connection Status
  ''' </summary>
  ''' <param name="NetCamDevice"></param>
  ''' <remarks></remarks>
  Private Sub UpdateCameraConnectionStatus(ByRef NetCamDevice As NetCamDevice)

    Try

      Dim dv_root_addr As String = String.Format("{0}{1}-Root", "Sighthound", NetCamDevice.DeviceId)
      Dim dv_root_type As String = "Sighthound Camera"
      Dim dv_root_name As String = "Sighthound Camera"

      Dim dv_addr As String = NetCamDevice.dv_addr_status
      Dim dv_name As String = NetCamDevice.Name
      Dim dv_type As String = "Sighthound Camera Connection"

      GetHomeSeerDevice(dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
      Dim dv_value As Integer = IIf(NetCamDevice.Status = "failed", 2, 1)
      SetDeviceValue(dv_addr, dv_value)

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Updates the Netowrk Camera Rules
  ''' </summary>
  ''' <param name="NetCamDevice"></param>
  ''' <remarks></remarks>
  Private Function GetNetworkCameraRules(ByRef NetCamDevice As NetCamDevice) As List(Of SighthoundRule)

    Dim SightHoundRules As New List(Of SighthoundRule)

    Try
      '
      ' Process Rules List
      '
      Dim RulesEnabled As Integer = 0
      Dim RulesList As ArrayList = remoteGetRulesForCamera(NetCamDevice.Name)
      For Each ruleName As String In RulesList
        Dim SighthoundRule As SighthoundRule = getRuleInfo(ruleName)
        If SighthoundRule.Name.Length > 0 Then
          Dim rule_id As Integer = GetNetCamRuleId(NetCamDevice.Id, SighthoundRule.Name)
          If rule_id > 0 Then
            SighthoundRule.Id = rule_id
            SightHoundRules.Add(SighthoundRule)

            If SighthoundRule.Enabled = True Then RulesEnabled += 1

            '
            ' Update the HS Device Rule
            '
            Dim NetCamDeviceId As String = NetCamDevice.Id.ToString.PadLeft(3, "0")
            Dim RuleId As String = rule_id.ToString.PadLeft(3, "0")

            Dim dv_root_addr As String = String.Format("{0}{1}-Root", "Sighthound", NetCamDeviceId)
            Dim dv_root_type As String = "Sighthound Camera"
            Dim dv_root_name As String = "Sighthound Camera"

            Dim dv_addr As String = String.Format("{0}{1}-Rule{2}", "Sighthound", NetCamDeviceId, RuleId)
            Dim dv_name As String = SighthoundRule.Name
            Dim dv_type As String = "Sighthound Camera Rule"

            GetHomeSeerDevice(dv_root_name, dv_root_type, dv_root_addr, dv_name, dv_type, dv_addr)
            Dim dv_value As Integer = IIf(SighthoundRule.Enabled = True, 1, 0)
            SetDeviceValue(dv_addr, dv_value)
          End If

        End If
      Next

    Catch pEx As Exception

    End Try

    Return SightHoundRules

  End Function

  ''' <summary>
  ''' Refresh the Network Camera List
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub EnableNetworkCameras(ByVal Enabled As Boolean)

    Try

      Dim ping As String = SendPing()

      If ping.Length > 0 Then

        Dim CameraNames As ArrayList = remoteGetCameraNames()
        For Each cameraName As String In CameraNames
          Try
            Dim SighthoundCameraStatus As SighthoundCameraStatus = getCameraStatusAndEnabled(cameraName)

            If SighthoundCameraStatus.Enabled <> Enabled Then

              Dim NetCamDevice As NetCamDevice = NetCamDevices.Find(Function(s) s.Name = cameraName)

              Dim cameraId As Integer = NetCamDevice.Id
              EnableNetworkCamera(cameraId, Enabled)
            End If

          Catch pEx As Exception
            '
            ' Catch Program Exception
            '
          End Try

        Next

      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EnableNetworkCameras()")
    End Try

  End Sub

  ''' <summary>
  ''' Sets the Enable State of a Network Camera
  ''' </summary>
  ''' <param name="Id"></param>
  ''' <param name="Enabled"></param>
  ''' <remarks></remarks>
  Public Sub EnableNetworkCamera(ByVal Id As Integer, ByVal Enabled As Boolean)

    Try

      Dim NetCamDevice As NetCamDevice = NetCamDevices.Find(Function(c) c.Id = Id)
      If Not IsNothing(NetCamDevice) Then
        EnableCamera(NetCamDevice.Name, Enabled)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EnableNetworkCamera()")
    End Try

  End Sub

  ''' <summary>
  ''' Updates the Sighthound Arming Status
  ''' </summary>
  ''' <param name="camerasEnabled"></param>
  ''' <param name="camerasTotal"></param>
  ''' <remarks></remarks>
  Public Sub UpdateSighthoundArming(camerasEnabled As Integer, camerasTotal As Integer)

    Try

      Dim dv_addr As String = String.Concat(IFACE_NAME, "-Arming")
      SetDeviceValue(dv_addr, camerasEnabled)

      Dim dv_string As String = String.Format("{0}/{1} ON", camerasEnabled, camerasTotal)
      SetDeviceString(dv_addr, dv_string)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "UpdateSighthoundArming()")
    End Try

  End Sub

  ''' <summary>
  ''' Returns the NetCamDevices List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevices() As List(Of NetCamDevice)

    Try

      Return NetCamDevices

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Returns the NetCamDevice
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevice(ByVal SighthoundCamera) As NetCamDevice

    Try

      Dim DeviceId As String = Regex.Match(SighthoundCamera, "\d\d\d").Value
      Dim NetCamDevice As NetCamDevice = NetCamDevices.Find(Function(s) s.DeviceId = DeviceId)
      Return NetCamDevice

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Get Camera Snapshot
  ''' </summary>
  ''' <param name="cameraId"></param>
  ''' <param name="cameraUri"></param>
  ''' <param name="iTimeoutSec"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Friend Function GetCameraSnapshot(ByVal cameraId As Integer,
                                              ByVal cameraUri As String,
                                              ByVal strSnapshotFilename As String,
                                              ByVal strThumbnailFilename As String,
                                              Optional ByVal iTimeoutSec As Integer = 60) As WebExceptionStatus

    Try
      '
      ' Format the URL
      '
      Dim strProtocol As String = IIf(gSighthoundVersion = "2", "http", "https")
      Dim netcam_url As String = String.Format("{0}://{1}:{2}{3}?height=800&width=800", strProtocol, gSighthoundURL, gSighthoundPort.ToString, cameraUri)
      WriteMessage(String.Format("GetCameraSnapshot is running command: {0}.", netcam_url), MessageType.Debug)

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(netcam_url), HttpWebRequest)
      lxRequest.Timeout = iTimeoutSec * 1000
      lxRequest.Credentials = New NetworkCredential(gSighthoundUsername, gSighthoundPassword)

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

        Dim lnBuffer As Byte()
        Dim lnFile As Byte()

        Using lxBR As New BinaryReader(lxResponse.GetResponseStream())

          Using lxMS As New MemoryStream()

            lnBuffer = lxBR.ReadBytes(1024)

            While lnBuffer.Length > 0
              lxMS.Write(lnBuffer, 0, lnBuffer.Length)
              lnBuffer = lxBR.ReadBytes(1024)
            End While

            lnFile = New Byte(CInt(lxMS.Length) - 1) {}
            lxMS.Position = 0
            lxMS.Read(lnFile, 0, lnFile.Length)

            Try

              Using image As Image = Image.FromStream(lxMS)

                If Not image Is Nothing Then

                  If strSnapshotFilename.EndsWith("_snapshot.jpg") Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                    End If
                  End If

                  If strThumbnailFilename.EndsWith("_thumbnail.jpg") Then
                    If File.Exists(strThumbnailFilename) = True Then
                      File.SetCreationTime(strThumbnailFilename, DateTime.Now)
                    End If
                  End If

                  image.Save(strSnapshotFilename, image.RawFormat)

                  Dim iWidth As Integer = image.Width
                  Dim iHeight As Integer = image.Height
                  For i = 1 To 10 Step 1
                    iWidth = image.Width / i
                    iHeight = image.Height / i
                    If iHeight <= 36 Then Exit For
                  Next

                  Dim imageThumb As Image = image.GetThumbnailImage(iWidth, iHeight, Nothing, New IntPtr())
                  imageThumb.Save(strThumbnailFilename, image.RawFormat)

                End If

                image.Dispose()

              End Using

            Catch pEx As ArgumentException
              '
              ' We got here because the data was not an image
              '
              Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} failed: {1}", gSighthoundURL, pEx.Message)
              WriteMessage(strErrorMessage, MessageType.Warning)

            End Try

            lxMS.Close()
            lxBR.Close()

          End Using

        End Using

        lxResponse.Close()

      End Using

      Return WebExceptionStatus.Success

    Catch pEx As System.Net.WebException
      '
      ' Process the WebException
      '
      Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} failed: {1}", gSighthoundURL, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return pEx.Status
    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} failed: {1}", gSighthoundURL, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Debug)

      Return WebExceptionStatus.UnknownError
    End Try

  End Function

  ''' <summary>
  ''' Sends the ping method to Sighthound
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SendPing() As String

    Dim methodResponse As String = String.Empty
    Dim pingResponse As String = String.Empty

    Try

      Dim xml = New XElement("methodCall", _
                          New XElement("methodName", "ping"), _
                          New XElement("params"))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          ' <string>&lt;type 'exceptions.Exception'&gt;:method "remoteGetCameraNames1" is not supported</string>
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        pingResponse = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/string").InnerText

      End Using

      _pingResponse = pingResponse

      _pingFlag = 1

      '<?xml version='1.0'?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <string>r14470yM_1000l12</string>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      _pingResponse = String.Empty

      _pingFlag = -1
    End Try

    Return pingResponse

  End Function

  ''' <summary>
  ''' Returns the list of the Sighthound Camera Names
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function remoteGetCameraNames() As ArrayList

    Dim methodResponse As String = String.Empty
    Dim cameraNames As New ArrayList

    Try

      Dim xml = New XElement("methodCall", _
                          New XElement("methodName", "remoteGetCameraNames"), _
                          New XElement("params"))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim responseNodes As XmlNodeList = xmlDoc.SelectNodes("//methodResponse/params/param/value/array/data/value/array/data/value")
        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        '
        ' Process the methodResponse
        '
        If Not IsNothing(responseNodes) = True Then

          For Each XmlNode As XmlNode In responseNodes
            Dim cameraName As String = XmlNode.InnerText
            If Regex.IsMatch(cameraName, "Any camera", RegexOptions.IgnoreCase) = False Then
              cameraNames.Add(cameraName)
            End If
          Next

        End If

      End Using

      '<?xml version='1.0'?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <boolean>1</boolean>
      '            </value>
      '            <value>
      '              <array>
      '                <data>
      '                  <value>
      '                    <string>Any camera</string>
      '                  </value>
      '                  <value>
      '                    <string>Basement</string>
      '                  </value>
      '                  <value>
      '                    <string>Basement-Utility</string>
      '                  </value>
      '                  <value>
      '                    <string>Den</string>
      '                  </value>
      '                  <value>
      '                    <string>Family-Room</string>
      '                  </value>
      '                  <value>
      '                    <string>Garage</string>
      '                  </value>
      '                </data>
      '              </array>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      '
      ' The XML is not valid
      '
    End Try

    Return cameraNames

  End Function

  ''' <summary>
  ''' Returns the Camera Uri
  ''' </summary>
  ''' <param name="cameraName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getCameraStatusAndEnabled(cameraName As String) As SighthoundCameraStatus

    Dim methodResponse As String = String.Empty
    Dim SighthoundCameraStatus As New SighthoundCameraStatus

    Try

      Dim xml = New XElement("methodCall", _
                              New XElement("methodName", "getCameraStatusAndEnabled"), _
                              New XElement("params",
                              New XElement("param", New XElement("value", New XElement("string", cameraName))) _
                          ))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        Dim status As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value/string").InnerText
        SighthoundCameraStatus.Status = status

        Dim enabled As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value/boolean").InnerText
        SighthoundCameraStatus.Enabled = IIf(enabled = "1", True, False)

      End Using

      '<?xml version='1.0'?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <string>off</string>
      '            </value>
      '            <value>
      '              <boolean>0</boolean>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return SighthoundCameraStatus

  End Function

  ''' <summary>
  ''' Returns the Camera Uri
  ''' </summary>
  ''' <param name="cameraName"></param>
  ''' <param name="imageWidth"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function remoteGetCameraUri(cameraName As String, Optional imageWidth As Integer = 520) As String

    Dim methodResponse As String = String.Empty
    Dim strCameraUri As String = String.Empty

    Try

      Dim xml = New XElement("methodCall", _
                              New XElement("methodName", "remoteGetCameraUri"), _
                              New XElement("params",
                              New XElement("param", New XElement("value", New XElement("string", cameraName))), _
                              New XElement("param", New XElement("value", New XElement("int", imageWidth))), _
                              New XElement("param", New XElement("value", New XElement("string", ""))), _
                              New XElement("param", New XElement("value", New XElement("string", "image/jpeg"))) _
                          ))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        strCameraUri = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value/string").InnerText
      End Using

      '<?xml version='1.0'?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <boolean>1</boolean>
      '            </value>
      '            <value>
      '              <string>/camera/49103/image.jpg</string>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return strCameraUri

  End Function

  ''' <summary>
  ''' Returns a list of Live Cameras
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function remoteGetLiveCameras() As ArrayList

    Dim methodResponse As String = String.Empty
    Dim cameraList As New ArrayList

    Try

      Dim xml = New XElement("methodCall", _
                          New XElement("methodName", "remoteGetLiveCameras"), _
                          New XElement("params"))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim responseNodes As XmlNodeList = xmlDoc.SelectNodes("//methodResponse/params/param/value/array/data/value/array/data/value/array/data/value/string")
        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        '
        ' Process the methodResponse
        '
        If Not IsNothing(responseNodes) = True Then

          For Each XmlNode As XmlNode In responseNodes
            Dim cameraName As String = XmlNode.InnerText
            If cameraName.Length > 0 Then
              cameraList.Add(cameraName)
            End If
          Next

        End If

      End Using

      '<?xml version='1.0'?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <boolean>1</boolean>
      '            </value>
      '            <value>
      '              <array>
      '                <data>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Basement</string>
      '                        </value>
      '                        <value>
      '                          <string></string>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Basement-Utility</string>
      '                        </value>
      '                        <value>
      '                          <string></string>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Den</string>
      '                        </value>
      '                        <value>
      '                          <string></string>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Family-Room</string>
      '                        </value>
      '                        <value>
      '                          <string></string>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Garage</string>
      '                        </value>
      '                        <value>
      '                          <string></string>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                </data>
      '              </array>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception

    End Try

    Return cameraList

  End Function

  ''' <summary>
  ''' Returns the Rules for a Camera
  ''' </summary>
  ''' <param name="cameraName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function remoteGetClipsForRule(cameraName As String, ruleName As String, ticks As Long, Optional clipcount As Integer = 25) As ArrayList

    Dim methodResponse As String = String.Empty
    Dim rulesList As New ArrayList

    Try

      Dim xml = New XElement("methodCall", _
                              New XElement("methodName", "remoteGetClipsForRule"), _
                              New XElement("params",
                              New XElement("param", New XElement("value", New XElement("string", cameraName))), _
                              New XElement("param", New XElement("value", New XElement("string", ruleName))), _
                              New XElement("param", New XElement("value", New XElement("double", ticks.ToString))), _
                              New XElement("param", New XElement("value", New XElement("int", clipcount.ToString))), _
                              New XElement("param", New XElement("value", New XElement("int", "0"))), _
                              New XElement("param", New XElement("value", New XElement("boolean", "0"))) _
                          ))

      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim responseNodes As XmlNodeList = xmlDoc.SelectNodes("//methodResponse/params/param/value/array/data/value/array/data/value/array/data")
        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        '
        ' Process the methodResponse
        '
        If Not IsNothing(responseNodes) = True Then

          For Each XmlNode As XmlNode In responseNodes
            '  If ruleName.Length > 0 Then
            '    rulesList.Add(ruleName)
            '  End If
          Next

        End If

      End Using

      '<?xml version="1.0"?>
      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value><boolean>1</boolean></value>
      '            <value>
      '              <array>
      '                <data>
      '                  <value>
      '                    <array>
      '                      <data>
      '                        <value>
      '                          <string>Den</string>
      '                        </value>
      '                        <value>
      '                          <array>
      '                            <data>
      '                              <value>
      '                                <int>1428282664</int>
      '                              </value>
      '                              <value>
      '                                <int>983</int>
      '                              </value>
      '                            </data>
      '                          </array>
      '                        </value>
      '                        <value>
      '                          <array>
      '                            <data>
      '                              <value>
      '                                <int>1428282682</int>
      '                              </value>
      '                              <value>
      '                                <int>983</int>
      '                              </value>
      '                            </data>
      '                          </array>
      '                        </value>
      '                        <value>
      '                          <array>
      '                            <data>
      '                              <value>
      '                                <int>1428282672</int>
      '                              </value>
      '                              <value>
      '                                <int>983</int>
      '                              </value>
      '                            </data>
      '                          </array>
      '                        </value>
      '                        <value>
      '                          <string>9:11 pm</string>
      '                        </value>
      '                        <value>
      '                          <array>
      '                            <data>
      '                              <value>
      '                                <int>121138</int>
      '                              </value>
      '                              <value>
      '                                <int>121139</int>
      '                              </value>
      '                            </data>
      '                          </array>
      '                        </value>
      '                      </data>
      '                    </array>
      '                  </value>
      '                </data>
      '              </array>
      '            </value>
      '            <value>
      '              <int>56</int>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return rulesList

  End Function

  ''' <summary>
  ''' Returns the Rules for a Camera
  ''' </summary>
  ''' <param name="cameraName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function remoteGetRulesForCamera(cameraName As String) As ArrayList

    Dim methodResponse As String = String.Empty
    Dim rulesList As New ArrayList

    Try

      Dim xml = New XElement("methodCall", _
                              New XElement("methodName", "remoteGetRulesForCamera"), _
                              New XElement("params",
                              New XElement("param", New XElement("value", New XElement("string", cameraName))) _
                          ))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim responseNodes As XmlNodeList = xmlDoc.SelectNodes("//methodResponse/params/param/value/array/data/value/array/data/value/string")
        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        '
        ' Process the methodResponse
        '
        If Not IsNothing(responseNodes) = True Then

          For Each XmlNode As XmlNode In responseNodes
            Dim ruleName As String = XmlNode.InnerText
            If ruleName.Length > 0 Then
              rulesList.Add(ruleName)
            End If
          Next

        End If

      End Using

      ' Results for "Any camera"
      '  <methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <boolean>1</boolean>
      '            </value>
      '            <value>
      '              <array>
      '                <data>
      '                  <value>
      '                    <string>All objects</string>
      '                  </value>
      '                  <value>
      '                    <string>People</string>
      '                  </value>
      '                  <value>
      '                    <string>Unknown objects</string>
      '                  </value>
      '                  <value>
      '                    <string>Any object in Basement-Utility</string>
      '                  </value>
      '                  <value>
      '                    <string>Any object in Basement</string>
      '                  </value>
      '                  <value>
      '                    <string>Any object in Den</string>
      '                  </value>
      '                  <value>
      '                    <string>Any object in Family-Room</string>
      '                  </value>
      '                  <value>
      '                    <string>Any object in Garage</string>
      '                  </value>
      '                  <value>
      '                    <string>People entering or leaving through a door in Den</string>
      '                  </value>
      '                  <value>
      '                    <string>People entering through a door in Basement</string>
      '                  </value>
      '                  <value>
      '                    <string>People leaving through a door in Basement</string>
      '                  </value>
      '                </data>
      '              </array>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return rulesList

  End Function

  ''' <summary>
  ''' Returns the Rule Information
  ''' </summary>
  ''' <param name="ruleName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getRuleInfo(ruleName As String) As SighthoundRule

    Dim methodResponse As String = String.Empty
    Dim SighthoundRule As New SighthoundRule

    Try

      Dim xml = New XElement("methodCall", _
                              New XElement("methodName", "getRuleInfo"), _
                              New XElement("params",
                              New XElement("param", New XElement("value", New XElement("string", ruleName))) _
                          ))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

        '
        ' Process the methodResponse
        '
        Try

          Dim ruleDesc1 As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value[1]/string").InnerText
          Dim ruleDesc2 As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value[2]/string").InnerText
          Dim ruleStatus As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value[3]/string").InnerText
          Dim ruleEnabled As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value/boolean").InnerText
          Dim ruleAction As String = xmlDoc.SelectSingleNode("//methodResponse/params/param/value/array/data/value/array/data/value/string").InnerText

          SighthoundRule.Name = ruleDesc1
          SighthoundRule.Description = ruleDesc2
          SighthoundRule.Status = ruleStatus
          SighthoundRule.Enabled = CBool(ruleEnabled)
          SighthoundRule.Action = ruleAction

        Catch pEx As Exception

        End Try

      End Using

      '<methodResponse>
      '  <params>
      '    <param>
      '      <value>
      '        <array>
      '          <data>
      '            <value>
      '              <string>Any object in Den</string>
      '            </value>
      '            <value>
      '              <string>Any object in Den</string>
      '            </value>
      '            <value>
      '              <string>Every day - 24 hours</string>
      '            </value>
      '            <value>
      '              <boolean>1</boolean>
      '            </value>
      '            <value>
      '              <array>
      '                <data>
      '                  <value>
      '                    <string>RecordResponse</string>
      '                  </value>
      '                </data>
      '              </array>
      '            </value>
      '          </data>
      '        </array>
      '      </value>
      '    </param>
      '  </params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return SighthoundRule

  End Function

  ''' <summary>
  ''' Enables or Disables a Camera Rule
  ''' </summary>
  ''' <param name="ruleName"></param>
  ''' <param name="ruleEnabled"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function remoteEnableRule(ByVal ruleName As String, ByVal ruleEnabled As Boolean) As Boolean

    Dim methodResponse As String = String.Empty

    Try

      Dim Enabled As Integer = IIf(ruleEnabled = True, 1, 0)

      Dim xml = New XElement("methodCall", _
                          New XElement("methodName", "remoteEnableRule"), _
                          New XElement("params",
                          New XElement("param", New XElement("value", New XElement("string", ruleName))), _
                          New XElement("param", New XElement("value", New XElement("boolean", Enabled.ToString)))))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

      End Using

      '<methodCall>
      '<methodName>remoteEnableRule</methodName>
      '<params>
      '<param>
      '<value>
      '<string>people entering through a door in den</string>
      '</value>
      '</param>
      '<param>
      '<value>
      '<boolean>0</boolean>
      '</value>
      '</param>
      '</params>
      '</methodCall>

      '<?xml version='1.0'?>
      '<methodResponse>
      '<params>
      '<param>
      '<value><boolean>1</boolean></value>
      '</param>
      '</params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Enable Camera Rule By Id
  ''' </summary>
  ''' <param name="ruleId"></param>
  ''' <param name="ruleEnabled"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function EnableSighthoundRuleById(ByVal ruleId As String, ByVal ruleEnabled As Boolean) As Boolean

    Try
      Dim SighthoundRuleName As String = GetNetCamRuleById(ruleId)

      Return remoteEnableRule(SighthoundRuleName, ruleEnabled)

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Enables or Disables a Network Camera
  ''' </summary>
  ''' <param name="cameraName"></param>
  ''' <param name="cameraEnabled"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function EnableCamera(ByVal cameraName As String, ByVal cameraEnabled As Boolean) As Boolean

    Dim methodResponse As String = String.Empty

    Try

      Dim Enabled As Integer = IIf(cameraEnabled = True, 1, 0)

      Dim xml = New XElement("methodCall", _
                          New XElement("methodName", "enableCamera"), _
                          New XElement("params",
                          New XElement("param", New XElement("value", New XElement("string", cameraName))), _
                          New XElement("param", New XElement("value", New XElement("boolean", Enabled.ToString)))))
      Dim data As Byte() = New ASCIIEncoding().GetBytes(xml.ToString)

      methodResponse = MethodCall(data)

      Using xmlResponse As New XmlTextReader(New StringReader(methodResponse))

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(xmlResponse)

        Dim faultNode As XmlNode = xmlDoc.SelectSingleNode("//methodResponse/fault")

        '
        ' Process the methodResponse fault
        '
        If Not IsNothing(faultNode) Then
          Dim errorString As String = faultNode.SelectSingleNode("value/struct/member/value/string").InnerText
          Throw New Exception(errorString)
        End If

      End Using

      '<methodCall>
      '<methodName>enableCamera</methodName>
      '<params>
      '<param>
      '<value>
      '<string>Basement</string>
      '</value>
      '</param>
      '<param>
      '<value>
      '<boolean>0</boolean>
      '</value>
      '</param>
      '</params>
      '</methodCall>

      '<?xml version='1.0'?>
      '<methodResponse>
      '<params>
      '<param>
      '<value><nil/></value></param>
      '</params>
      '</methodResponse>

    Catch pEx As Exception
      Throw New Exception(pEx.Message)
    End Try

    Return True

  End Function

  ''' <summary>
  ''' MethodCall to Sighthound Video
  ''' </summary>
  ''' <remarks></remarks>
  Private Function MethodCall(data As Byte(), Optional ByVal Timeout As Integer = 5) As String

    Dim responseText As String = String.Empty

    Try

      SyncLock HTTPLock

        Dim strProtocol As String = IIf(gSighthoundVersion = "2", "http", "https")
        Dim strURL As String = String.Format("{0}://{1}:{2}/xmlrpc/", strProtocol, gSighthoundURL, gSighthoundPort)

        '
        ' Build the HTTP Web Request
        '
        Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(strURL), HttpWebRequest)
        lxRequest.Timeout = Timeout * 1000
        lxRequest.Credentials = New NetworkCredential(gSighthoundUsername, gSighthoundPassword)
        lxRequest.Method = "POST"
        lxRequest.ContentType = "application/xml"
        'lxRequest.ContentLength = data.Length

        Using myStream As Stream = lxRequest.GetRequestStream
          myStream.Write(data, 0, data.Length)

          Using response As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

            Using reader = New StreamReader(response.GetResponseStream())

              responseText = reader.ReadToEnd()

              reader.Close()
            End Using

            response.Close()
          End Using

          myStream.Close()
        End Using

      End SyncLock

      _querySuccess += 1

    Catch pEx As Exception

      _queryFailure += 1

      Dim xml = New XElement("methodResponse", _
                              New XElement("fault",
                              New XElement("value",
                              New XElement("struct",
                              New XElement("member", New XElement("name", "faultCode"), New XElement("value", New XElement("int", "1"))), _
                              New XElement("member", New XElement("name", "faultString"), New XElement("value", New XElement("string", pEx.Message))) _
                              ))))
      Return xml.ToString

    End Try

    Return responseText

  End Function

  ''' <summary>
  ''' Checks to see if required credentials are available
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CheckCredentials() As Boolean

    Try

      Dim sbWarning As New StringBuilder

      If gSighthoundUsername.Length = 0 Then
        sbWarning.Append("Sighthound Username")
      End If
      If gSighthoundPassword.Length = 0 Then
        sbWarning.Append("Sighthound Password")
      End If
      If sbWarning.Length = 0 Then Return True

    Catch pEx As Exception

    End Try

    Return False

  End Function

  Function CertificateValidationCallBack(ByVal sender As Object,
                                         ByVal certificate As X509Certificate,
                                         ByVal chain As X509Chain,
                                         ByVal sslPolicyErrors As SslPolicyErrors) As Boolean
    Return True
  End Function

End Class
