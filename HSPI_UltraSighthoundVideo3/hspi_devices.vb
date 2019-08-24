Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service
Imports System.Text.RegularExpressions

Module hspi_devices

  Public DEV_INTERFACE As Byte = 1
  Public DEV_CONNECTION As Byte = 2
  Public DEV_CAMERAS As Byte = 3

  Dim bCreateRootDevice = True

  ''' <summary>
  ''' Update the list of monitored devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub FixDeviceTypes()

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration

    Dim strMessage As String = ""

    strMessage = "Entered FixDeviceTypes() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then

            If dv.Interface(Nothing) = IFACE_NAME Then
              '
              ' We found our device, so process based on device type
              '
              Dim dv_type As String = dv.Device_Type_String(hs)
              Dim dv_name As String = dv.Name(hs)
              Dim dv_location As String = dv.Location(hs)
              Dim dv_location2 As String = dv.Location2(hs)

              If dv_type.Contains("Sighhound") = True Or dv_location2.Contains("Sighhound") = True Then
                dv.Device_Type_String(hs) = dv_type.Replace("Sighhound", "Sighthound")
                dv.Name(hs) = dv_name.Replace("Sighhound", "Sighthound")
                dv.Location(hs) = dv_location.Replace("Sighhound", "Sighthound")
                dv.Location2(hs) = dv_location2.Replace("Sighhound", "Sighthound")

                hs.SaveEventsDevices()
              End If

            End If
          End If
        Loop
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "FixDeviceTypes()")
    End Try

    DevEnum = Nothing
    dv = Nothing

  End Sub

  ''' <summary>
  ''' Update the list of monitored devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateRootDevices()

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration

    Dim strMessage As String = ""

    strMessage = "Entered UpdateRootDevices() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then

            If dv.Interface(Nothing) = IFACE_NAME Then
              '
              ' We found our device, so process based on device type
              '
              Dim dv_ref As Integer = dv.Ref(hs)
              Dim dv_addr As String = dv.Address(hs)
              Dim dv_name As String = dv.Name(hs)
              Dim dv_type As String = dv.Device_Type_String(hs)

              If dv_addr = "NetCam-Root" Then

                Dim ChildDevices As Integer() = dv.AssociatedDevices(hs)
                For Each dv_child_ref As Integer In ChildDevices

                  Dim dv_child As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(dv_child_ref)
                  Dim bDeviceExists As Boolean = Not dv_child Is Nothing

                  If bDeviceExists = True Then
                    Dim dv_child_addr As String = dv_child.Address(hs)
                    Dim regexPattern As String = "^Sighthound-(?<id>(\d\d\d)$)"
                    If Regex.IsMatch(dv_child_addr, regexPattern) = True Then
                      '
                      ' Fix the Device Address
                      '
                      Dim cameraId As String = Regex.Match(dv_child_addr, regexPattern).Groups("id").ToString()
                      dv_child_addr = String.Format("Sighthound{0}-Camera", cameraId)
                      dv_child.Address(hs) = dv_child_addr

                      '
                      ' Remove the relationship
                      '
                      dv_child.AssociatedDevice_ClearAll(hs)
                      dv_child.Relationship(hs) = Enums.eRelationship.Not_Set

                      '
                      ' Save the changes
                      '
                      hs.SaveEventsDevices()
                    End If

                  End If

                Next

                hs.DeleteDevice(dv_ref)
                hs.SaveEventsDevices()
              End If

            End If
          End If
        Loop
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateRootDevices()")
    End Try

    DevEnum = Nothing
    dv = Nothing

  End Sub

  ''' <summary>
  ''' Function to initilize our plug-ins devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitPluginDevices() As String

    Dim strMessage As String = ""

    WriteMessage("Entered InitPluginDevices() function.", MessageType.Debug)

    Try
      Dim Devices As Byte() = {DEV_INTERFACE, DEV_CONNECTION, DEV_CAMERAS}
      For Each dev_cod As Byte In Devices
        Dim strResult As String = CreatePluginDevice(IFACE_NAME, dev_cod)
        If strResult.Length > 0 Then Return strResult
      Next

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitPluginDevices()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Subroutine to create a HomeSeer device
  ''' </summary>
  ''' <param name="base_code"></param>
  ''' <param name="dev_code"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePluginDevice(ByVal base_code As String, ByVal dev_code As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Try

      Select Case dev_code
        Case DEV_INTERFACE.ToString
          '
          ' Create the Sighthound State device
          '
          dv_name = "Sighthound State"
          dv_type = "Sighthound State"
          dv_addr = String.Concat(base_code, "-State")

        Case DEV_CONNECTION.ToString
          '
          ' Create the Sighthound Connectoin device
          '
          dv_name = "Sighthound Connection"
          dv_type = "Sighthound Connection"
          dv_addr = String.Concat(base_code, "-Connection")

        Case DEV_CAMERAS.ToString
          '
          ' Create the Sighthound Arming device
          '
          dv_name = "Sighthound Arming"
          dv_type = "Sighthound Arming"
          dv_addr = String.Concat(base_code, "-Arming")

        Case Else
          Throw New Exception(String.Format("Unable to create plug-in device for unsupported device name: {0}", dv_name))
      End Select

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Plug-ins"
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If bCreateRootDevice = True Then
        dv.AssociatedDevice_ClearAll(hs)
        Dim dvp_ref As Integer = CreateRootDevice("Plugin", IFACE_NAME, "Sighthound Plugin", IFACE_NAME, dv_ref)
        If dvp_ref > 0 Then
          dv.AssociatedDevice_Add(hs, dvp_ref)
        End If
        dv.Relationship(hs) = Enums.eRelationship.Child
      End If

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      Dim VSPair As VSPair
      Dim VGPair As VGPair

      Select Case dv_type
        Case "Sighthound State"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -3
          VSPair.Status = ""
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Disable"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Enable"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Disabled"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Enabled"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          Dim dev_status As Integer = IIf(gMonitoring = True, 1, 0)
          hs.SetDeviceValueByRef(dv_ref, dev_status, False)

        Case "Sighthound Connection"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Disconnected"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Unknown"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Connected"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VGPairs
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -1
          VGPair.RangeEnd = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "network_disconnected.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 1
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "network_connected.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          hs.SetDeviceValueByRef(dv_ref, 1, False)

        Case "Sighthound Arming"
          '
          ' Sighthound Camera Arming Device
          '
          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -3
          VSPair.Status = ""
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Disarm All"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Arm All"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 100
          VSPair.RangeStatusSuffix = "ON"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VGPairs
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -3
          VGPair.RangeEnd = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "arming_disabled.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 1
          VGPair.RangeEnd = 100
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "arming_enabled.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          'hs.SetDeviceValueByRef(dv_ref, 1, False)

        Case "Database"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Close"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Closed"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)
      End Select

      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)

      dv.Status_Support(hs) = True

      hs.SaveEventsDevices()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreatePluinDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="dev_root_id"></param>
  ''' <param name="dev_root_name"></param>
  ''' <param name="dev_root_type"></param>
  ''' <param name="dev_root_addr"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal dev_root_id As String, _
                                   ByVal dev_root_name As String, _
                                   ByVal dev_root_type As String, _
                                   ByVal dev_root_addr As String, _
                                   ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If dev_root_id = "Plugin" Then
        dv_name = "Sighthound Plugin"
        dv_addr = String.Format("{0}-Root", dev_root_name)
        dv_type = dev_root_type
      Else
        dv_name = dev_root_name
        dv_addr = dev_root_addr
        dv_type = dev_root_type
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(dev_root_id = "Plugin", "Plug-ins", dv_type)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Subroutine to create the HomeSeer device
  ''' </summary>
  ''' <param name="dv_root_name"></param>
  ''' <param name="dv_root_type"></param>
  ''' <param name="dv_root_addr"></param>
  ''' <param name="dv_name"></param>
  ''' <param name="dv_type"></param>
  ''' <param name="dv_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetHomeSeerDevice(ByVal dv_root_name As String, _
                                    ByVal dv_root_type As String, _
                                    ByVal dv_root_addr As String, _
                                    ByVal dv_name As String, _
                                    ByVal dv_type As String, _
                                    ByVal dv_addr As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim DevicePairs As New ArrayList

    Try
      '
      ' Define local variables
      '
      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        If bCreateRootDevice = True Then
          dv.AssociatedDevice_ClearAll(hs)
          Dim dvp_ref As Integer = CreateRootDevice("", dv_root_name, dv_root_type, dv_root_addr, dv_ref)
          If dvp_ref > 0 Then
            dv.AssociatedDevice_Add(hs, dvp_ref)
          End If
          dv.Relationship(hs) = Enums.eRelationship.Child
        End If

        hs.SaveEventsDevices()
      End If

      '
      ' Exit if our device exists
      '
      If bDeviceExists = True Then Return dv_addr

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Sighthound Camera"

        '
        ' Update the last change date
        ' 
        dv.Last_Change(hs) = DateTime.Now
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      Dim VSPair As VSPair
      Dim VGPair As VGPair
      Select Case dv_type
        Case "Sighthound Camera"
          '
          ' Set the image
          '
          Dim strSnapshotFilename As String = String.Format("{0}/{1}_snapshot.jpg", "images/hspi_ultrasighthoundvideo3/snapshots", dv_addr)
          Dim strThumbnailFilename As String = String.Format("{0}/{1}_thumbnail.jpg", "images/hspi_ultrasighthoundvideo3/snapshots", dv_addr)
          dv.Image(hs) = strThumbnailFilename
          dv.ImageLarge(hs) = strSnapshotFilename

          DevicePairs.Clear()
          DevicePairs.Add(New hspi_device_pairs(-3, "", "netcam_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-2, "Disable", "netcam_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Enable", "netcam_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Disabled", "netcam_disabled.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Enabled", "netcam_enabled.png", HomeSeerAPI.ePairStatusControl.Status))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

        Case "Sighthound Camera Connection"
          '
          ' Sighthound Camera Status
          '
          DevicePairs.Add(New hspi_device_pairs(0, "Offline", "camera_offline.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Online", "camera_online.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(2, "Failed", "camera_failed.png", HomeSeerAPI.ePairStatusControl.Status))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

        Case "Sighthound Camera Rule"
          '
          ' Sighthound Camera Rules
          '
          DevicePairs.Add(New hspi_device_pairs(-3, "", "camera_rule_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-2, "Disable", "camera_rule_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Enable", "camera_rule_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Disabled", "camera_rule_disabled.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Enabled", "camera_rule_enabled.png", HomeSeerAPI.ePairStatusControl.Status))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

      End Select

      hs.SaveEventsDevices()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "GetHomeSeerDevice()")
    End Try

    Return dv_addr

  End Function

  ''' <summary>
  ''' Returns the HomeSeer Device Address
  ''' </summary>
  ''' <param name="Address"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDeviceAddress(ByVal Address As String) As String

    Try

      Dim dev_ref As Integer = hs.DeviceExistsAddress(Address, False)
      If dev_ref > 0 Then
        Return Address
      Else
        Return ""
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "GetDevCode")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="srDeviceCode"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByCode(ByVal srDeviceCode As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.GetDeviceRef(srDeviceCode)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "LocateDeviceByCode")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String, _
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblNewValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblNewValue Then
            hs.SetDeviceValueByRef(dv_ref, dblNewValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String, _
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        Dim strDeviceString As String = hs.DeviceString(dv_ref)

        If strDeviceString <> dv_string Then
          hs.SetDeviceString(dv_ref, dv_string, True)
        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
