Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Web.UI.WebControls
Imports System.IO
Imports System.Text.RegularExpressions

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder
      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)

        '
        ' Check Motion Triggers
        '
        If parts("rule") <> "" Then
          If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
            stb.AppendFormat("{0} is not authorized.", LoggedInUser)
          Else

            Dim rule As String = parts("rule")

            Dim strTrigger As String = String.Format("{0},{1}", "Motion Trigger", rule)
            hspi_plugin.CheckTrigger(IFACE_NAME, SighthoundVideoTriggers.MotionTrigger, -1, strTrigger)

            stb.AppendLine("OK")

          End If

          Return stb.ToString

        End If
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasighthoundvideo3/js/lightbox.min.js""></script>")
      Header.AppendLine("<link type=""text/css"" rel=""stylesheet"" href=""/hspi_ultrasighthoundvideo3/css/lightbox.css"" />")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Setup page timer
      '
      Me.RefreshIntervalMilliSeconds = 1000 * 3
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'>" & BuildTabOptions() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Cameras"
      tab.tabDIVID = "tabCameras"
      tab.tabContent = "<div id='divCameras'>" & BuildTabCameras() & "</div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Sighthound Video System Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", gSighthoundVersion)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Host:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", gSighthoundURL)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Port:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", gSighthoundPort)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", SighthoundVideoAPI.GetPingResponse)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' General Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Sighthound Server</td>")
      stb.AppendLine(" </tr>")

      Dim selSighthoundVersion As New clsJQuery.jqDropList("SighthoundVersion", Me.PageName, False)
      selSighthoundVersion.id = "SighthoundVersion"
      selSighthoundVersion.toolTip = "Select the Sighthound Version"

      selSighthoundVersion.AddItem("Version 2.x", "2", gSighthoundVersion = "2")
      selSighthoundVersion.AddItem("Version 3.x", "3", gSighthoundVersion = "3")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Sighthound Version</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSighthoundVersion.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim tbIPAddress As New clsJQuery.jqTextBox("SighthoundURL", "text", gSighthoundURL, PageName, 20, False)
      tbIPAddress.id = "SighthoundURL"
      tbIPAddress.promptText = "Enter your Sighthound IP Address"
      tbIPAddress.toolTip = tbIPAddress.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">IP Address</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbIPAddress.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim tbPort As New clsJQuery.jqTextBox("SighthoundPort", "text", gSighthoundPort, PageName, 5, False)
      tbPort.id = "SighthoundPort"
      tbPort.promptText = "Enter your Sighthound Port Number"
      tbPort.toolTip = tbPort.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Port</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbPort.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Sighthound Credentials</td>")
      stb.AppendLine(" </tr>")

      Dim tbUsername As New clsJQuery.jqTextBox("SighthoundUsername", "text", gSighthoundUsername, PageName, 25, False)
      tbUsername.id = "SighthoundUsername"
      tbUsername.promptText = "Enter your Sighthound Username"
      tbUsername.toolTip = tbUsername.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Username</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbUsername.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim tbPassword As New clsJQuery.jqTextBox("SighthoundPassword", "text", "", PageName, 25, False)
      tbPassword.id = "SighthoundPassword"
      tbPassword.promptText = "Enter your Sighthound Password"
      tbPassword.toolTip = tbPassword.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Password</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbPassword.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Sighthound Options</td>")
      stb.AppendLine(" </tr>")

      Dim selRefreshInterval As New clsJQuery.jqDropList("selRefreshInterval", Me.PageName, False)
      selRefreshInterval.id = "selRefreshInterval"
      selRefreshInterval.toolTip = "Specify how often the plug-in should refresh the snapshots."

      Dim txtRefreshInterval As String = GetSetting("Options", "PollSighthoundVideoThread", "5")
      For index As Integer = 2 To 60
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Seconds", index.ToString)
        selRefreshInterval.AddItem(desc, value, index.ToString = txtRefreshInterval)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Refresh&nbsp;Interval</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selRefreshInterval.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>E-Mail Notification Options</td>")
      stb.AppendLine(" </tr>")

      Dim bSendEmail As Boolean = True
      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "gSMTPFrom", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "gSMTPTo", "")

      '
      ' E-Mail Notification Options (Send Email)
      '
      'Dim selSendEmail As New clsJQuery.jqDropList("selSendEmail", Me.PageName, False)
      'selSendEmail.id = "selSendEmail"
      'selSendEmail.toolTip = "Enable sending snapshot event e-mail notifications using the plug-in."

      'Dim bSendEmail As Boolean = CBool(GetSetting("EmailNotification", "EmailEnabled", False))
      'Dim txtSendEmail As String = IIf(bSendEmail = True, "1", "0")

      'selSendEmail.AddItem("No", "0", txtSendEmail = "0")
      'selSendEmail.AddItem("Yes", "1", txtSendEmail = "1")

      'stb.AppendLine(" <tr>")
      'stb.AppendLine("  <td class='tablecell'>Email Notification</td>")
      'stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSendEmail.Build, vbCrLf)
      'stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email To)
      '
      Dim txtEmailRcptTo As String = GetSetting("EmailNotification", "EmailRcptTo", strEmailRcptTo)
      Dim tbEmailRcptTo As New clsJQuery.jqTextBox("txtEmailRcptTo", "text", txtEmailRcptTo, PageName, 50, False)
      tbEmailRcptTo.id = "txtEmailRcptTo"
      tbEmailRcptTo.promptText = "The default e-mail notification recipient address."
      tbEmailRcptTo.toolTip = tbEmailRcptTo.promptText
      tbEmailRcptTo.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email To</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailRcptTo.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email From)
      '
      Dim txtEmailFrom As String = GetSetting("EmailNotification", "EmailFrom", strEmailFromDefault)
      Dim tbEmailFrom As New clsJQuery.jqTextBox("txtEmailFrom", "text", txtEmailFrom, PageName, 50, False)
      tbEmailFrom.id = "txtEmailFrom"
      tbEmailFrom.promptText = "The default e-mail notification sender address."
      tbEmailFrom.toolTip = tbEmailFrom.promptText
      tbEmailFrom.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email From</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailFrom.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email Subject)
      '
      Dim txtEmailSubject As String = GetSetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT)
      Dim tbEmailSubject As New clsJQuery.jqTextBox("txtEmailSubject", "text", txtEmailSubject, PageName, 65, False)
      tbEmailSubject.id = "txtEmailSubject"
      tbEmailSubject.promptText = "The default e-mail notification subject."
      tbEmailSubject.toolTip = tbEmailSubject.promptText
      tbEmailSubject.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email Subject</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailSubject.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email Body Template)
      '
      Dim jqButton1 As New clsJQuery.jqButton("btnSaveEmailBody", "Save", Me.PageName, True)
      Dim chkResetEmailBody As New clsJQuery.jqCheckBox("chkResetEmailBody", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkResetEmailBody.checked = False
      chkResetEmailBody.enabled = bSendEmail

      Dim txtEmailBodyDisabled As String = IIf(bSendEmail = True, "", "disabled")
      Dim txtEmailBody As String = GetSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email Body Template</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea {0} rows='6' cols='70' name='txtEmailBody'>{1}</textarea>{2}{3}</td>{4}", txtEmailBodyDisabled,
                                                                                                                                    txtEmailBody.Trim.Replace("~", vbCrLf),
                                                                                                                                    jqButton1.Build(),
                                                                                                                                    chkResetEmailBody.Build,
                                                                                                                                    vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCameras(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder
      Dim cameras As New ArrayList

      stb.Append(clsPageBuilder.FormStart("frmCameras", "frmCameras", "Post"))

      Dim selSnapshotsMaxWidth As New clsJQuery.jqDropList("selSnapshotsMaxWidth", Me.PageName, False)

      selSnapshotsMaxWidth.id = "selSnapshotsMaxWidth"
      selSnapshotsMaxWidth.toolTip = "Specifies the maximum snapshot width."
      selSnapshotsMaxWidth.AddItem("Auto", "Auto", IIf(gSnapshotMaxWidth = "Auto", True, False))
      selSnapshotsMaxWidth.AddItem("160 px", "160px", IIf(gSnapshotMaxWidth = "160px", True, False))
      selSnapshotsMaxWidth.AddItem("320 px", "320px", IIf(gSnapshotMaxWidth = "320px", True, False))
      selSnapshotsMaxWidth.AddItem("640 px", "640px", IIf(gSnapshotMaxWidth = "640px", True, False))
      selSnapshotsMaxWidth.autoPostBack = True

      '
      ' General Options
      '
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Sighthound Cameras</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'><div id='lastRefresh'/></td>")
      stb.AppendLine("  <td class='tablecell' align='right'>Snapshot Width: " & selSnapshotsMaxWidth.Build & "</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' colspan='2'>")

      Dim NetCamDevices As List(Of NetCamDevice) = SighthoundVideoAPI.GetNetCamDevices()
      For Each NetCamDevice As NetCamDevice In NetCamDevices
        Dim strSnapshotFilename As String = String.Format("{0}_snapshot.jpg", NetCamDevice.dv_addr)
        cameras.Add(NetCamDevice.dv_addr)

        'Dim strFileName As String = String.Format("{0}/{1}", gSnapshotDirectory, strSnapshotFilename).Replace("/", "\")
        'Dim Date2 As DateTime = File.GetCreationTime(strFileName)
        'Dim Date1 As DateTime = DateTime.Now
        'Dim ts As TimeSpan = Date1.Subtract(Date2)

        stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", gSnapshotMaxWidth)
        stb.AppendFormat("   <a id=""lnk_{0}"" href=""#"" title=""{1}"" data-lightbox=""lightbox[0]"">", NetCamDevice.dv_addr, NetCamDevice.Name)
        stb.AppendFormat("    <img id=""img_{0}"" rel=""lightbox[0]"" style=""width:100%"" />", NetCamDevice.dv_addr)
        stb.AppendLine("   </a>")
        stb.AppendFormat("   <div>{0}</div><div>{1}</div>", NetCamDevice.Name, NetCamDevice.dv_addr)
        stb.AppendLine("  </div>")
      Next

      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      '
      ' Update the Refresh Interval
      '
      Dim iRefreshInterval As Integer = 1000
      Dim strRefreshInterval As String = GetSetting("Options", "PollSighthoundVideoThread", "5")
      If IsNumeric(strRefreshInterval) = True Then
        iRefreshInterval *= Integer.Parse(strRefreshInterval)
      End If

      stb.AppendLine("<script>")
      stb.AppendLine("function refreshSnapshots() {")
      stb.AppendLine("  var ticks = new Date().getTime();")
      stb.AppendLine("  var url = 'images/hspi_ultrasighthoundvideo3/snapshots/';")
      For Each dv_addr As String In cameras
        Dim strSnapshotFilename As String = String.Format("{0}_snapshot.jpg", dv_addr)
        stb.AppendLine("    $('#img_" & dv_addr & "').attr('src', url + '" & strSnapshotFilename & "?ticks=' + ticks);")
        stb.AppendLine("    $('#lnk_" & dv_addr & "').attr('href', url + '" & strSnapshotFilename & "?ticks=' + ticks);")
      Next
      stb.AppendLine("    $('#lastRefresh').html( new Date() + '' );")
      stb.AppendLine("};")

      stb.AppendLine("$(function() { refreshSnapshots(); setInterval(function() { refreshSnapshots(); }, " & iRefreshInterval.ToString & ");});")

      'stb.AppendLine("var options = {};")
      'stb.AppendLine("$( "".effect"" ).effect( ""highlight"", options, 3000, callback );")
      'stb.AppendLine("});")
      'stb.AppendLine("function callback() { };")
      stb.AppendLine("</script>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divCameras", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabCameras")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "PostMessage")
    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If
      For Each keyName As String In postData.AllKeys
        Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
      Next

      ''
      '' Process Sighthound Rules
      ''
      'Dim regexPattern As String = "^SighthoundRule-(?<ruleId>(\d+))_\d+$"
      'If Regex.IsMatch(postData("id"), regexPattern) = True Then
      '  Dim ruleId As Integer = Integer.Parse(Regex.Match(postData("id"), regexPattern).Groups("ruleId").ToString())
      '  Dim SighthoundRule As String = String.Format("SighthoundRule-{0}", ruleId.ToString)
      '  Dim enabled As Boolean = CBool(postData(SighthoundRule))
      '  SighthoundVideoAPI.EnableSighthoundRuleByName(ruleId, enabled)
      'End If

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          Me.divToUpdate.Add("divStatus", BuildTabStatus())

        Case "tabOptions"

        Case "tabCameras"
          Me.pageCommands.Add("starttimer", "")

        Case "selSnapshotsMaxWidth"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotsMaxWidth", value)

          Me.divToUpdate.Add("divCameras", BuildTabCameras())

        Case "SighthoundVersion"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sighthound", "Version", value)

          gSighthoundVersion = value

        Case "SighthoundURL"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sighthound", "URL", value)

          gSighthoundURL = value

        Case "SighthoundPort"
          Dim value As String = postData(postData("id"))
          If IsNumeric(value) = True Then
            SaveSetting("Sighthound", "Port", value)
          End If

          gSighthoundPort = value

        Case "SighthoundUsername"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sighthound", "Username", value)

          gSighthoundUsername = value

        Case "SighthoundPassword"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sighthound", "Password", value)

          gSighthoundPassword = value

        Case "selRefreshInterval"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "PollSighthoundVideoThread", value)

          PostMessage("The Sighthound Video refresh interval has been updated.")

        Case "selSendEmail"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailEnabled", strValue)
          BuildTabOptions(True)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailRcptTo"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailRcptTo", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailFrom"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailFrom", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailSubject"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailSubject", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "btnSaveEmailBody", "txtEmailBody"
          Dim strValue As String = postData("txtEmailBody").Trim.Replace(vbCrLf, "~")
          SaveSetting("EmailNotification", "EmailBody", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "chkResetEmailBody"
          SaveSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE)
          BuildTabOptions(True)

          PostMessage("The E-mail Notification option has been updated.")

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class