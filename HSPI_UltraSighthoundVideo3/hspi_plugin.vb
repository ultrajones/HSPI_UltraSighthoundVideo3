Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Net.Sockets
Imports System.Text
Imports System.Net
Imports System.Xml
Imports System.Data.Common
Imports System.Drawing
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel
Imports System.IO

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable

  Const Pagename = "Events"

  Public HSDevices As New SortedList

  Public SighthoundVideoAPI As hspi_sighthound_api

  Public gSighthoundUsername As String = String.Empty
  Public gSighthoundPassword As String = String.Empty
  Public gSighthoundVersion As String = "2"
  Public gSighthoundURL As String = String.Empty
  Public gSighthoundPort As Integer = 80

  Public Const IFACE_NAME As String = "UltraSighthoundVideo3"

  Public Const LINK_TARGET As String = "hspi_ultrasighthoundvideo3/hspi_ultrasighthoundvideo3.aspx"
  Public Const LINK_URL As String = "hspi_ultrasighthoundvideo3.html"
  Public Const LINK_TEXT As String = "UltraSighthoundVideo3"
  Public Const LINK_PAGE_TITLE As String = "UltraSighthoundVideo3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultrasighthoundvideo3/UltraSighthoundVideo3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = String.Empty
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultrasighthoundvideo3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gEventEmailNotification As Boolean = False     ' Indicates if events should generate an e-mail notification
  Public gSnapshotMaxWidth As String = "Auto"
  Public gStatusImageSizeWidth As Integer = 32
  Public gStatusImageSizeHeight As Integer = 32

  Public gMonitoring As Boolean = True

  Public HSAppPath As String = ""
  Public gSnapshotDirectory As String = ""

  Public EMAIL_SUBJECT As String = String.Format("{0} - {1} [{2}] Snapshot", IFACE_NAME, "$camera_name", "$camera_id")
  Public Const EMAIL_BODY_TEMPLATE As String = "The Sighthound Camera $camera_name [$camera_id] is $camera_enabled and $camera_status as of $date $time.~~" & _
                                               "Sighthound Camera Rule Summary:~" & _
                                               "$camera_rules_summary~~" & _
                                               "This camera has $camera_rules_enabled out of $camera_rules_count rules enabled."

#Region "UltraSighthoundVideo3 Public Functions"

#Region "HSPI - NetCam Devices"

  ''' <summary>
  ''' Gets the NetCam Id from the database
  ''' </summary>
  ''' <param name="netcam_name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDeviceId(ByVal netcam_name As String) As Integer

    Dim netcam_id As Integer = 0

    Try

      Dim strSQL As String = String.Format("SELECT netcam_id FROM tblNetCamDevices WHERE netcam_name='{0}'", netcam_name)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the results
          '
          If dtrResults.Read() Then
            netcam_id = dtrResults("netcam_id")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception

    End Try

    Return netcam_id

  End Function

  ''' <summary>
  ''' Gets the NetCam Devices from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevicesFromDB() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetNetCamDevicesFromDB() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblNetCamDevices")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim strDbProviderFactory As String = hs.GetINISetting("Database", "DbProviderFactory", "System.Data.SQLite", gINIFile)
      Dim MyProvider As DbProviderFactory = DbProviderFactories.GetFactory(strDbProviderFactory)

      Dim MyDA As System.Data.IDbDataAdapter = MyProvider.CreateDataAdapter
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetNetCamDevicesFromDB()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new NetCam Device into the database
  ''' </summary>
  ''' <param name="netcam_name"></param>
  ''' <param name="netcam_uri"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertNetCamDevice(ByVal netcam_name As String,
                                     ByVal netcam_uri As String) As Integer

    Dim strMessage As String = ""
    Dim netcam_id As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_name.Length = 0 Then
        Throw New Exception("The netcam_name field is empty.  Unable to insert new NetCam device into the database.")
      End If

      Dim strSQL As String = String.Format("INSERT INTO tblNetCamDevices (" _
                                           & " netcam_name, netcam_uri" _
                                           & ") VALUES (" _
                                           & "'{0}', '{1}' );",
                                           netcam_name, netcam_uri)
      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          netcam_id = dbcmd.ExecuteScalar()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertNetCamDevice()")
      Return 0
    End Try

    Return netcam_id

  End Function

  ''' <summary>
  ''' Updates existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <param name="netcam_name"></param>
  ''' <param name="netcam_uri"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateNetCamDevice(ByVal netcam_id As Integer, _
                                     ByVal netcam_name As String, _
                                     ByVal netcam_uri As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_name.Length = 0 Then
        Throw New Exception("The netcam_name field is empty.  Unable to update NetCam device in the database.")
      End If

      Dim strSql As String = String.Format("UPDATE tblNetCamDevices SET " _
                                          & " netcam_name='{0}', " _
                                          & " netcam_uri='{1}'" _
                                          & "WHERE netcam_id={3}", netcam_name, netcam_uri, netcam_id.ToString)

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSql


        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "UpdateNetCamDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateNetCamDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Removes existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamDevice(ByVal netcam_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = String.Format("DELETE FROM tblNetCamDevices WHERE netcam_id={0}", netcam_id.ToString)

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "DeleteNetCamDevice() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        '
        ' Remove the camera snapshot directory
        '
        DeleteNetCamSnapshotDir(netcam_id.ToString)

        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteNetCamDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Gets the Rule Id from the database
  ''' </summary>
  ''' <param name="natcam_id"></param>
  ''' <param name="rule_name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamRuleId(ByVal natcam_id As Integer, ByVal rule_name As String) As Integer

    Dim rule_id As Integer = 0

    Try
      Dim strSQL As String = String.Format("SELECT rule_id FROM tblNetCamRules WHERE netcam_id={0} AND rule_name='{1}'", natcam_id, rule_name)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the results
          '
          If dtrResults.Read() Then
            rule_id = dtrResults("rule_id")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      '
      ' Insert the rule if rule_name cannot be found
      '
      If rule_id = 0 Then

        strSQL = String.Format("INSERT INTO tblNetCamRules (netcam_id, rule_name) VALUES ({0}, '{1}' );", natcam_id, rule_name)
        strSQL &= "SELECT last_insert_rowid();"

        Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          dbcmd.Connection = DBConnectionMain
          dbcmd.CommandType = CommandType.Text
          dbcmd.CommandText = strSQL

          SyncLock SyncLockMain
            rule_id = dbcmd.ExecuteScalar()
          End SyncLock

          dbcmd.Dispose()

        End Using

      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "GetNetCamRuleId()")
      Return 0
    End Try

    Return rule_id

  End Function

  ''' <summary>
  ''' Gets the Rule Name from the database
  ''' </summary>
  ''' <param name="rule_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamRuleById(ByVal rule_id As Integer) As String

    Dim rule_name As String = String.Empty

    Try
      Dim strSQL As String = String.Format("SELECT rule_name FROM tblNetCamRules WHERE rule_id={0}", rule_id)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the results
          '
          If dtrResults.Read() Then
            rule_name = dtrResults("rule_name")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "GetNetCamRuleId()")
      Return String.Empty
    End Try

    Return rule_name

  End Function

#End Region

  ''' <summary>
  ''' Removes the snapshot directory
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <remarks></remarks>
  Public Sub DeleteNetCamSnapshotDir(ByVal netcam_id As String)

    Try
      '
      ' Format the NetCam Id
      '
      Dim strNetCamId As String = netcam_id
      If Regex.IsMatch(strNetCamId, "NetCam\d\d\d") = False Then
        strNetCamId = String.Format("NetCam{0}", netcam_id.PadLeft(3, "0"))
      End If

      '
      ' Remove the directory
      '
      Dim strDirectory As String = FixPath(String.Format("{0}/{1}", gSnapshotDirectory, strNetCamId))
      If Directory.Exists(strDirectory) = True Then
        WriteMessage(String.Format("Removing {0} because the camera is no longer defined.", strNetCamId), MessageType.Warning)
        Directory.Delete(strDirectory, True)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "DeleteNetCamSnapshotDir()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() function.", MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strSection = "Sighthound" And strKey = "Password" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Apply the Sighthound Username
      '
      If strSection = "Sighthound" And strKey = "Username" Then
        gSighthoundUsername = strValue
      End If

      '
      ' Apply the Sighthound Password
      '
      If strSection = "Sighthound" And strKey = "Password" Then
        If strValue.Length = 0 Then Exit Sub
        gSighthoundPassword = strValue
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      '
      ' Apply the Sighthound SnapshotsMaxWidth
      '
      If strSection = "Options" And strKey = "SnapshotsMaxWidth" Then
        gSnapshotMaxWidth = strValue
      End If

      If strSection = "EmailNotification" And strKey = "EmailEnabled" Then
        gEventEmailNotification = CBool(strValue)
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraSighthoundVideo3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, GetEnumDescription(SighthoundVideoTriggers.MotionTrigger))    ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      'Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber

      Case SighthoundVideoTriggers.MotionTrigger
        Dim triggerName As String = GetEnumName(SighthoundVideoTriggers.MotionTrigger)

        '
        ' Start SighthoundVideo Status Trigger
        '
        Dim ActionSelected As String = trigger.Item("RuleNumber")
        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "RuleNumber", UID, sUnique)

        Dim jqStatus As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStatus.autoPostBack = True

        jqStatus.AddItem("(Select Rule Number)", "", (ActionSelected = ""))
        For index As Integer = 0 To 99
          Dim strOptionValue As String = index.ToString
          Dim strOptionName As String = index.ToString
          jqStatus.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next
        stb.Append("Select the Rule Number:  ")
        stb.Append(jqStatus.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case SighthoundVideoTriggers.MotionTrigger
          Dim triggerName As String = GetEnumName(SighthoundVideoTriggers.MotionTrigger)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True

              Case InStr(sKey, triggerName & "RuleNumber_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("RuleNumber") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case SighthoundVideoTriggers.MotionTrigger
          If trigger.Item("RuleNumber") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case SighthoundVideoTriggers.MotionTrigger
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = "Motion Trigger"
          Dim strRuleNumber As String = trigger.Item("RuleNumber")

          Dim strLanIPAddress As String = hs.LANIP()
          strLanIPAddress = Regex.Match(strLanIPAddress, "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}").Value

          Dim iWebServerPort As String = hs.WebServerPort()
          Dim strProtocol As String = IIf(gSighthoundVersion = "2", "http", "https")
          Dim strURLPath As String = String.Format("{0}://{1}:{2}/UltraSighthoundVideo3?rule={3}", strProtocol, strLanIPAddress, iWebServerPort, strRuleNumber)

          stb.AppendFormat("<div>{0} rule number {1}</div>", strTriggerName, strRuleNumber)
          stb.AppendLine("<div>On your Sighthound Video Server, use the following command to trigger this event:</div>")
          stb.AppendFormat("<div>curl --request GET {0}</div>", strURLPath)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Public Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case SighthoundVideoTriggers.MotionTrigger
                Dim strEventType As String = "Motion Trigger"
                Dim strRuleNumber As String = trigger.Item("RuleNumber")

                Dim strTriggerCheck As String = String.Format("{0},{1}$", strEventType, strRuleNumber)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, GetEnumName(SighthoundVideoActions.EmailCameraSnapshot))   ' 1
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case SighthoundVideoActions.EmailCameraSnapshot
        Dim actionName As String = GetEnumName(SighthoundVideoActions.EmailCameraSnapshot)

        Dim ActionSelected As String = action.Item("SighthoundCamera")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "SighthoundCamera", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.toolTip = "Select the Sighthound Video Camera"
        jqNetCam.autoPostBack = False

        jqNetCam.AddItem("(Select Sighthound Camera)", "", (ActionSelected = ""))

        For Each NetCamDevice As NetCamDevice In SighthoundVideoAPI.GetNetCamDevices()

          Dim strOptionValue As String = String.Format("Sighthound{0}", NetCamDevice.DeviceId)
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("<table border='1'>")
        stb.Append("<tr>")
        stb.Append(" <td style='text-align:right'>")
        stb.Append("Sighthound Camera:")
        stb.Append(" </td>")
        stb.Append(" <td colspan='2'>")
        stb.Append(jqNetCam.Build)
        stb.Append(" </td>")
        stb.Append("</tr>")

        '
        ' Start E-mail Recipient
        '
        Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "gSMTPTo", "")
        Dim txtEmailRcptTo As String = GetSetting("EmailNotification", "EmailRcptTo", strEmailRcptTo)
        ActionSelected = IIf(action.Item("EmailRecipient").Length = 0, txtEmailRcptTo, action.Item("EmailRecipient"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailRecipient", UID, sUnique)

        Dim jqEmailRecipient As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 75, True)
        jqEmailRecipient.toolTip = "Enter the Email address of the intended recipient."
        jqEmailRecipient.editable = True

        stb.Append("<tr>")
        stb.Append(" <td style='text-align:right'>")
        stb.Append("Recipient:")
        stb.Append(" </td>")
        stb.Append(" <td colspan='2'>")
        stb.Append(jqEmailRecipient.Build)
        stb.Append(" </td>")
        stb.Append("</tr>")

        '
        ' Start E-mail From
        '
        Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "gSMTPFrom", "")
        Dim txtEmailFrom As String = GetSetting("EmailNotification", "EmailFrom", strEmailFromDefault)
        ActionSelected = IIf(action.Item("EmailFrom").Length = 0, txtEmailFrom, action.Item("EmailFrom"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailFrom", UID, sUnique)

        Dim jqEmailFrom As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 75, True)
        jqEmailFrom.toolTip = "Enter the Email address of the intended recipient."
        jqEmailFrom.editable = True

        stb.Append("<tr>")
        stb.Append(" <td style='text-align:right'>")
        stb.Append("Sender:")
        stb.Append(" </td>")
        stb.Append(" <td colspan='2'>")
        stb.Append(jqEmailFrom.Build)
        stb.Append(" </td>")
        stb.Append("</tr>")

        '
        ' Start E-mail Subject
        '
        Dim txtEmailSubject As String = GetSetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT)
        ActionSelected = IIf(action.Item("EmailSubject").Length = 0, txtEmailSubject, action.Item("EmailSubject"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailSubject", UID, sUnique)

        Dim jqEmailSubject As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 75, True)
        jqEmailSubject.toolTip = "Enter the subject of the message."
        jqEmailSubject.editable = True

        stb.Append("<tr>")
        stb.Append(" <td style='text-align:right'>")
        stb.Append("Subject:")
        stb.Append(" </td>")
        stb.Append(" <td colspan='2'>")
        stb.Append(jqEmailSubject.Build)
        stb.Append(" </td>")
        stb.Append("</tr>")

        '
        ' Start E-mail Body
        '
        Dim txtEmailBody As String = GetSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE).Replace("~", vbCrLf)
        ActionSelected = IIf(action.Item("EmailBody").Length = 0, txtEmailBody, action.Item("EmailBody"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailBody", UID, sUnique)

        Dim jqEmailBody As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 75, True)
        jqEmailBody.toolTip = "Enter the body of the message.  See the helpfile for a list of supported replacment variables."
        jqEmailBody.editable = True

        stb.AppendLine("<tr>")
        stb.AppendLine(" <td style='text-align:right'>")
        stb.AppendLine("Message:")
        stb.AppendLine(" </td>")
        stb.AppendLine(" <td>")
        stb.AppendFormat("<textarea style='display: table-cell;' cols='60' rows='10' id='txtMessageBody' class='txtDropTarget' name='{0}'>{1}</textarea>", actionId, ActionSelected)
        stb.AppendLine(" </td>")
        stb.AppendLine("</tr>")

        '
        ' Start Submit Button
        '
        Dim jqButton As New clsJQuery.jqButton("btnSave", "Save Message", Pagename, True)

        stb.Append("<tr>")
        stb.Append(" <td>")
        stb.Append(" </td>")
        stb.Append(" <td colspan='2'>")
        stb.Append(jqButton.Build)
        stb.Append(" </td>")
        stb.Append("</tr>")
        stb.Append("</table>")

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case SighthoundVideoActions.EmailCameraSnapshot
          Dim actionName As String = GetEnumName(SighthoundVideoActions.EmailCameraSnapshot)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "SighthoundCamera_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SighthoundCamera") = ActionValue

              Case InStr(sKey, actionName & "EmailRecipient_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailRecipient") = ActionValue

              Case InStr(sKey, actionName & "EmailFrom_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailFrom") = ActionValue

              Case InStr(sKey, actionName & "EmailSubject_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailSubject") = ActionValue

              Case InStr(sKey, actionName & "EmailBody_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailBody") = ActionValue


            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case SighthoundVideoActions.EmailCameraSnapshot
        If action.Item("SighthoundCamera") = "" Then Configured = False
        If action.Item("EmailRecipient") = "" Then Configured = False
        If action.Item("EmailFrom") = "" Then Configured = False
        If action.Item("EmailSubject") = "" Then Configured = False
        If action.Item("EmailBody") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case SighthoundVideoActions.EmailCameraSnapshot
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(SighthoundVideoActions.EmailCameraSnapshot)

          Dim SighthoundCamera As String = action.Item("SighthoundCamera")
          Dim SighthoundCameraName As String = SighthoundVideoAPI.GetCameraNameById(SighthoundCamera)

          Dim EmailRecipient As String = action.Item("EmailRecipient")
          Dim EmailFrom As String = action.Item("EmailFrom")
          Dim EmailSubject As String = action.Item("EmailSubject")

          stb.AppendFormat("{0} {1}: Send latest <font class='event_Txt_Selection'>{2} [{3}]</font> snapshot " & _
                  "from <font class='event_Txt_Selection'>{4}</font> " & _
                  "to <font class='event_Txt_Selection'>{5}</font> " & _
                  "with the subject <font class='event_Txt_Selection'>{6}</font>",
                  IFACE_NAME,
                  strActionName, _
                  SighthoundCameraName, _
                  SighthoundCamera, _
                  EmailFrom, _
                  EmailRecipient, _
                  EmailSubject)

        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber

        Case SighthoundVideoActions.EmailCameraSnapshot

          Dim SighthoundCamera As String = action.Item("SighthoundCamera")
          Dim NetCamDevice As NetCamDevice = SighthoundVideoAPI.GetNetCamDevice(SighthoundCamera)

          Dim EmailRecipient As String = action.Item("EmailRecipient")
          Dim EmailFrom As String = action.Item("EmailFrom")
          Dim EmailSubject As String = action.Item("EmailSubject")
          Dim EmailBody As String = action.Item("EmailBody")

          If EmailRecipient.Length = 0 Then
            Throw New Exception(String.Format("Unable to send e-mail notification for {0} because the e-mail recipient is emtpy.", SighthoundCamera))
          ElseIf EmailFrom.Length = 0 Then
            Throw New Exception(String.Format("Unable to send e-mail notification for {0} because the e-mail sender is emtpy.", SighthoundCamera))
          ElseIf EmailSubject.Length = 0 Then
            Throw New Exception(String.Format("Unable to send e-mail notification for {0} because the e-mail subject is emtpy.", SighthoundCamera))
          ElseIf EmailBody.Length = 0 Then
            Throw New Exception(String.Format("Unable to send e-mail notification for {0} because the e-mail body is emtpy.", SighthoundCamera))
          ElseIf NetCamDevice Is Nothing Then
            Throw New Exception(String.Format("Unable to send e-mail notification for {0} because it was not found in the active list of Sighthound cameras.", SighthoundCamera))
          End If

          EmailSubject = EmailSubject.Replace("$camera_id", SighthoundCamera)
          EmailSubject = EmailSubject.Replace("$camera_name", NetCamDevice.Name)

          EmailBody = EmailBody.Replace("$camera_id", SighthoundCamera)
          EmailBody = EmailBody.Replace("$camera_name", NetCamDevice.Name)
          EmailBody = EmailBody.Replace("$camera_enabled", IIf(NetCamDevice.Enabled = True, "armed", "disarmed"))
          EmailBody = EmailBody.Replace("$camera_status", IIf(NetCamDevice.Status = "connecting", "online", NetCamDevice.Status))
          EmailBody = EmailBody.Replace("$camera_rules_enabled", NetCamDevice.RulesEnabled)
          EmailBody = EmailBody.Replace("$camera_rules_count", NetCamDevice.RulesCount)
          EmailBody = EmailBody.Replace("$camera_rules_summary", NetCamDevice.RulesSummary)


          Dim Attachments As String() = {""}
          Dim strIdentifier As String = String.Format("{0}-Camera", SighthoundCamera)
          Dim strSnapshotFilename As String = FixPath(String.Format("{0}\{1}_snapshot.jpg", gSnapshotDirectory, strIdentifier))
          Attachments.SetValue(strSnapshotFilename, 0)

          Dim List() As String = hs.GetPluginsList()
          If List.Contains("UltraSMTP3:") = True Then
            '
            ' Send e-mail using UltraSMTP3
            '
            hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {EmailRecipient, EmailSubject, EmailBody, Attachments})
          Else
            '
            ' Send e-mail using HomeSeer
            '
            hs.SendEmail(EmailRecipient, EmailFrom, "", "", EmailSubject, EmailBody, Attachments(0))
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum SighthoundVideoTriggers
  <Description("Sighthound Video Motion Trigger")> _
  MotionTrigger = 1
End Enum

Public Enum SighthoundVideoActions
  <Description("Email Camera Snapshot")> _
  EmailCameraSnapshot = 1
End Enum
