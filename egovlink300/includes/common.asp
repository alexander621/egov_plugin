<%
' to turn the system off and redirect to the outage.html page, set the 0 to 1 below
SystemOutage 0

'LEAVE THIS COMMENTED OUT, Chip - 4/17/02    'Option Explicit 
Response.Buffer = True
%>
<!-- #include file="adovbs.inc" //-->
<!-- #include file="../custom/includes/custom.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_upload.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_email.asp" //-->

<%
Dim sTmp, oAutoLoginCmd, sTimeOutPath

Const RootPath = "/ecTeamLink/"

' Code to catch known SQL Intrusion attempts
If InStr(UCase(request.servervariables("QUERY_STRING")),"DECLARE") > 0 And InStr(UCase(request.servervariables("QUERY_STRING")),"CHAR") > 0 And InStr(UCase(request.servervariables("QUERY_STRING")),"EXEC") > 0 Then 
	response.End 
End If 
If InStr(UCase(request.servervariables("QUERY_STRING")),"SELECT") > 0 And InStr(UCase(request.servervariables("QUERY_STRING")),"VARCHAR") > 0 And InStr(UCase(request.servervariables("QUERY_STRING")),"CAST") > 0 Then 
	response.End
End If

'Some constants used on almost every page
' Some are dupes from adovbs.inc and so are commented out
'Const adExecuteNoRecords = 128
'Const adCmdStoredProc = 4
Const adBit = 11
'Const adCmdText = 1
'Const adInteger = 3
'Const adVarChar = 200
'Const adLongVarChar = 201
Const adDateTime = 135
'Const adParamReturnValue = 4
'Const adParamInput = 1
'Const adParamOutput = 2
'Const adOpenStatic = 3
'Const adUseClient = 3
'Const adLockReadOnly = 1
'Const adStateOpen = 1

Const tabCount = 8
Const tabHome =			1
Const tabMessages =		2
Const tabDocuments =	3
Const tabCommittees =	4
Const tabDiscussions =	5
Const tabVoting =		6
Const tabMeetings =		7
Const tabAdmin =		8

Const ITEM_TYPE_DOCUMENT    = 1
Const ITEM_TYPE_VOTE        = 2
Const ITEM_TYPE_DISCUSSION  = 3
Const ITEM_TYPE_TEXT        = 4

Const MAILBOX_IN            = 1
Const MAILBOX_DRAFT         = 2
Const MAILBOX_SENT          = 3

Const COMPOSE_TYPE_REPLY    = 1
Const COMPOSE_TYPE_REPLYALL = 2
Const COMPOSE_TYPE_FOWARD   = 3

'If not logged in, check if user has stored cookie
sTmp = Request.Cookies("User")("UserID") & ""

'Check to see if user is logged in already 
If Session("UserID") = 0 Or Session("UserID") = "" Then

  If sTmp <> "" Then  'If they have a valid cookie
    Session("UserID") = CLng(sTmp)
    Session("OrgID") = Request.Cookies("User")("OrgID")
    Session("FullName") = Request.Cookies("User")("FullName") & ""
    Session("PageSize") = Request.Cookies("User")("PageSize")
    Session("ShowStockTicker") = Request.Cookies("User")("ShowStockTicker")
    Session("Permissions") = Request.Cookies("User")("Permissions") & ""
    Session("GroupImage") = Request.Cookies("User")("GroupImage")

    Set oAutoLoginCmd = Server.CreateObject("ADODB.Command")
    With oAutoLoginCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "NewAuditEvent"
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("UserID", adInteger, adParamInput, 4, sTmp)
      .Parameters.Append .CreateParameter("Type", adVarChar, adParamInput, 25, "Login (Auto)")
      .Parameters.Append .CreateParameter("Object", adVarChar, adParamInput, 50, Session.SessionID)
      .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 200, Request.ServerVariables("REMOTE_ADDR"))
      .Execute
    End With
    Set oAutoLoginCmd = Nothing
  Else
    'If no stored cookie then send to login page if required
    If Application("LogInRequired") = True Then
      If PageIsRequiredByLogin = False Then
        'keep initial page request in memory
        Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
        'redirect to login page
        Response.Redirect RootPath & "login.asp"
      End If
    End If
  End If
End If


'If User is not a guest and they dont log in automatically, write a script to prevent Session Timeout
If Session("UserID") > 0 And sTmp = "" Then
%>
  <script language="Javascript">
    function mySetTimeout() {
      window.setTimeout("onTimeout()", 60000*15);  //15 minutes til popup, then 5 more til logout if no response
    }
    function onTimeout() {
      window.open("<%=RootPath%>timeoutPopup.asp", "_timeout", "alwaysRaised=yes,height=100,width=400,location=no,menubar=no,resizeable=no,scrollbars=no,resizeable=no,status=no,left=" + (screen.width-400)/2 + ",top=" + (screen.height-100)/2)
    }
    window.onload = mySetTimeout;
  </script>
<%
End If


'----------------------------------------------------------------------------------------
'	void SystemOutage iOutageFlag
'----------------------------------------------------------------------------------------
Sub SystemOutage( ByVal iOutageFlag )
	' this will redirect to the outage page so entire system can be shut down.
	Dim sRedirect

	If clng(iOutageFlag) = clng(1) Then
		sRedirect = "http://" + request.ServerVariables("server_name") + "/" + GetVirtualDirectyName() + "/outage.html"
		response.redirect sRedirect
	End If 

End Sub 


'Check to see if user has permission to access a specific function
Function HasPermission( PrivelageName )
  If InStr(1, Session("Permissions"), PrivelageName) Then
    HasPermission = True
  Else
    HasPermission = False
  End If
End Function

'Draw Standard Tab Heading (Which tab to put in front is passed as a param)
'Also the directory level is passed, so we know how many "../" -s we need

' -------------------------------------------------------------------------------------------------
Sub DrawTabs( FrontTab, dirLevel )

  dim sMsg, i, sLink, sWords, bPrevInFront, sRelDir

  'sRelDir=left("../../../../../../",dirLevel*3)
  sRelDir = RootPath

  sMsg = "        <table border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & vbCRLF
  
  ' Display customer logo 
  sMsg = sMsg & "<td valign=top background=""" & sRelDir & "images/logo_background.jpg"" >"
  If Session("GroupImage") <> "" Then
    sMsg = sMsg & "<img src='" & sRelDir & Session("GroupImage") & "' height=33></td>" 
  Else
    sMsg = sMsg & "<img src='" & sRelDir & custGraphic & "home.jpg' height=33></td>" 
  End If
  
  ' Begin iterating thru the various tabs
  for i = 1 to tabCount
    if mid(custTabVisible,i,1)="Y" then  'Use logic like this to implement FGA, and skip sections
	    
      if NOT(session("UserID")=0 and (i=tabMessages Or i=tabAdmin)) then 'dont draw the messages tab if not logged in
      
        sMsg = sMsg & "<td><img src=""" & sRelDir 

        select case i
          case tabHome:
            sLink = "default.asp"
            sWords = langTabHome
          case tabMessages:
            sLink = "messages/"
            sWords = langTabMessages
          case tabDocuments:
            sLink = "docs/"
            sWords = langTabDocuments
          case tabCommittees:
            sLink = "dirs/"
            sWords = langTabCommittees
          case tabDiscussions:
            sLink = "discussions/"
            sWords = langTabDiscussions
          case tabVoting:
            sLink = "polls/"
            sWords = langTabVoting
          case tabMeetings:
            sLink = "meetings/"
            sWords = langTabMeetings
          case tabAdmin:
            sLink = "admin/"
            sWords = langTabAdmin
        end select

        if i=1 then  'First tab gets special "home" handling:  Graphic is slightly different
          sMsg = sMsg & "images/tabfront"
          if FrontTab = 1 then 
            sMsg = sMsg & "_selected"
            bPrevInFront = True
          end if
        else 
          if bPrevInFront then 
            sMsg = sMsg & "images/tabright_selected"
            bPrevInFront = False
          else 
            sMsg = sMsg & "images/lefttab"
            if FrontTab = i then 
              sMsg = sMsg & "_selected"
              bPrevInFront = True
            end if
          end if
        end if

        sMsg = sMsg & ".jpg""></td><td background=""" & sRelDir & "images/back"
        if FrontTab = i then sMsg = sMsg & "_selected"
        sMsg = sMsg & ".jpg""><a href=""" & sRelDir & sLink & """ class=""tab"" target=""_top"">" & sWords & "</a></td>" & vbCRLF

      end if
   end if  'end Skip tab logic
  next

  Response.Write "  <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""menu"">"
  Response.Write "    <tr>"
  Response.Write "      <td height=""15""></td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"
  Response.Write "      <td colspan=""2"" background=""" & sRelDir  & "images/back_main.jpg"">"
  Response.Write sMsg
  Response.Write "            <td><img src=""" & sRelDir & "images/tabend"
  
  if bPrevInFront  then
    Response.Write "_selected"
  end if
  
  Response.Write ".jpg""></td>"
  if session("UserID")=0 or session("UserID")=""  then	 
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "login.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogIn & "</a></td>"
  else
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "signoff.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogOut & "</a></td>"
  end if
  Response.Write "          </tr>"
  Response.Write "        </table>"
  Response.Write "      </td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"

  if FrontTab=tabDocuments then
    Response.Write "      <td class=""submenu""><img src=""images/spacer.gif"" width=""8"" height=""1""></td>"
    Response.Write "      <td class=""submenu"" width=""100%"" style=""padding-top:2px;"">"  
    Response.Write "        <span class=""button"" onclick=""toggleMenu();"" onmouseover=""this.className='buttona';status='" & langCollapseMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/collapse.gif"" id=""img_toggle"" width=""18"" height=""18"" border=""0"" alt=""" & langCollapseMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraToc.document.location.reload()"" onmouseover=""this.className='buttona';status='" & langRefreshMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/refresh.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langRefreshMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=main.asp'"" onmouseover=""this.className='buttona';status='Home';"" onmouseout=""this.className='button';status='';""><img src=""images/document_home.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langDocumentsHome & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.focus();parent.fraTopic.print();"" onmouseover=""this.className='buttona';status='" & langPrintCurrDoc & "';"" onmouseout=""this.className='button';status='';""><img src=""images/print.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langPrintCurrDoc & """></span>"
    Response.Write "        &nbsp;"

    If HasPermission("CanEditDocuments") Then
      Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addfolder.asp'"" onmouseover=""this.className='buttona';status='" & langAddFolder & "';"" onmouseout=""this.className='button';status='';""><img src=""images/folder_closed.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddFolder & """></span>"
      Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addarticle.asp'"" onmouseover=""this.className='buttona';status='" & langAddDocument & "';"" onmouseout=""this.className='button';status='';""><img src=""images/document.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddDoc & """></span>"
      Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../addhelp.asp'"" onmouseover=""this.className='buttona';status='" & langAddHelp & "';"" onmouseout=""this.className='button';status='';""><img src=""images/helpdocument.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddHelp & """></span>"
       Response.Write "          <span class=""button"" onclick=""parent.fraToc.ToggleContextMenu();"" onmouseover=""this.className='buttona';status='" & langAddToggle & "';"" onmouseout=""this.className='button';status='';""><img src=""images/menu.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddToggle & """></span>"
      Response.Write "          &nbsp;"
      'JS 3/28/02 Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=admin.asp'"" onmouseover=""this.className='buttona';status='" & langAdmin & "';"" onmouseout=""this.className='button';status='';""><img src=""images/admin.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAdmin & """></span>"
    End If 
    Response.Write "      </td>" 
  else
    Response.Write "      <td class=""submenu""><img src=""" & sRelDir & "images/spacer.gif"" height=""28"" width=""1"" border=""0""></td>"
  end if

  Response.Write "    </tr>"
  Response.Write "  </table>"
  Response.Write sTimeOutScript
End Sub

' -------------------------------------------------------------------------------------------------
Sub DrawQuicklinks(searchCaption, dirLevel)

  Dim sDirLevel, i, sPath
  sDirLevel = Left("../../../../../../", dirLevel*3)
  'sPath = Request.ServerVariables("SCRIPT_NAME")
%>
  <div style="padding-bottom:8px;"><b><%=langQuicklinks%></b></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/home_small.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>"><%=langTabHome%></a></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/document_home.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>docs"><%=langDocuments%></a></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newgroup.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>dirs"><%=langTabCommittees%></a></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newdisc.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>discussions"><%=langDiscussions%></a></div>
 <!-- <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newfav.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>favorites"><%=langFavorites%></a></div> //-->
        
  <% If searchCaption <> "" Then %>
    <form name="frmQlSearch" action="search.asp" method="get">
      <input type="hidden" name="p" value="1">
      <div style="padding-bottom:3px;"><%=searchCaption%>:</div>
      <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" name="s"><br>
      <div align="right"><a href="javascript:document.all.frmQlSearch.submit();"><img src="<%=sDirLevel%>images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
    </form>
  <% End If %>
<%
End Sub

' -------------------------------------------------------------------------------------------------
Function AsciiToHtml( AsciiString )
  Dim sTmp
  sTmp = AsciiString

  sTmp = Replace(sTmp, """", "&quot;")
  'sTmp = Replace(sTmp, "<", "&lt;")
  'sTmp = Replace(sTmp, ">", "&gt;")
  sTmp = Replace(sTmp, vbCrLf, "<br>")

  AsciiToHtml = sTmp
End Function

' -------------------------------------------------------------------------------------------------
Function SQLText( DbString )
  If DbString & "" = "" Then
    SQLText = NULL
  Else
    SQLText = Replace(DbString, "'", "''")
  End If
End Function


' GET INFORMATION FOR CURRENT INSTITUTION
Dim sTopGraphicLeftURL,sHomeWebsiteURL,sTopGraphicRighURL,sEgovWebsiteURL,iPaymentGatewayID 
Dim sWelcomeMessage,sTagline,sActionDescription,sOrgName,iHeaderSize,sPaymentDescription,sHomeWebsiteTag,sEgovWebsiteTag,bCustomButtonsOn
Dim blnOrgAction,blnOrgDocument,blnOrgPayment,blnOrgCalendar,blnOrgFaq
Dim sOrgActionName,sOrgPaymentName,sOrgCalendarName,sOrgDocumentName,sOrgFaqName
Dim sorgVirtualSiteName
Dim blnCalRequest,iCalForm
Dim sOrgRegistration
Dim iTimeOffset
Dim blnMenuOn,blnFooterOn,blnCustomMenu
Dim sDefaultPhone,sDefaultEmail,sWaiverText
Dim sDefaultCity,sDefaultState,sDefaultZip
Dim sGoogleAnalyticAccnt
Dim blnSeparateIndex


iOrgID = SetOrganizationParameters()  ' This is in ../include_top_functions.asp
session("OrgId") = iOrgID

' -------------------------------------------------------------------------------------------------
' FUNCTION READFILE(SNAME)
'--------------------------------------------------------------------------------------------------
Function ReadFile(sPath)

	Const ForReading = 1
	FormatASCII = 0

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If oFSO.FileExists(sPath) Then 
		Set oStream = oFSO.OpenTextFile(sPath, ForReading, false, FormatASCII) 
		If Not oStream.AtEndofStream Then
			sRetValue = oStream.ReadAll()
		End If
		Set oStream = Nothing
	Else
		' NO FILE - DON'T DISPLAY 
		' response.write sPath

	End If

	
	Set oFSO = Nothing

	ReadFile = sRetValue

End Function


'--------------------------------------------------------------------------------------------------
' void LogPageVisit iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP, iorgid 
'--------------------------------------------------------------------------------------------------
Sub LogPageVisit( ByVal iSectionID, ByVal sDocumentTitle, ByVal sURL, ByVal datDate, ByVal datDateTime, ByVal sVisitorIP, ByVal iorgid )
	Dim oCmd

	' Log Page Visits
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "LogPageVisit"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@SectionId", 3, 1, 4, iSectionID)
		.Parameters.Append oCmd.CreateParameter("@DocumentTitle", 200, 1, 255, sDocumentTitle)
		.Parameters.Append oCmd.CreateParameter("@CompleteURL", 200, 1, 1024, Left(sURL,1000))
		.Parameters.Append oCmd.CreateParameter("@OrgId", 3, 1, 4, iorgid)
		.Parameters.Append oCmd.CreateParameter("@VisitorIP", 200, 1, 100, sVisitorIP)
		.Execute
	End With

	Set oCmd = Nothing

End Sub


' -------------------------------------------------------------------------------------------------
Function Track_DBsafe( ByVal strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then Track_DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	Track_DBsafe = sNewString
End Function


' -------------------------------------------------------------------------------------------------
Function Decode(sIn)
    dim x, y, abfrom, abto
    Decode="": ABFrom = ""

    For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 

    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
    For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
        If y = 0 then
            Decode = Decode & Mid(sin, x, 1)
        Else
            Decode = Decode & Mid(abfrom, y, 1)
        End If
    Next
End Function


' -------------------------------------------------------------------------------------------------
Function Encode(sIn)
    dim x, y, abfrom, abto
    Encode="": ABFrom = ""

    For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 

    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
    For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
        If y = 0 Then
             Encode = Encode & Mid(sin, x, 1)
        Else
             Encode = Encode & Mid(abto, y, 1)
        End If
    Next
End Function 




'--------------------------------------------------------------------------------------------------
' string GetInternalDefaultEmail( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetInternalDefaultEmail( ByVal iOrgId )
	Dim sSql, oDefaultEmail

	'GetInternalDefaultEmail = "webmaster@eclink.com" 
 GetInternalDefaultEmail = "noreply@eclink.com"

	' This will yield one row per org with the default email for the internal contact
	sSql = "Select isnull(internal_default_email,'') as internal_default_email from organizations where orgid = " & iOrgId 

	Set oDefaultEmail = Server.CreateObject("ADODB.Recordset")
	oDefaultEmail.Open sSql, Application("DSN"), 3, 1

	If Not oDefaultEmail.EOF Then 
		GetInternalDefaultEmail = oDefaultEmail("internal_default_email")
	End If 

	oDefaultEmail.close
	Set oDefaultEmail = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function FormatForJavaScript( sString )
'--------------------------------------------------------------------------------------------------
Function FormatForJavaScript( ByVal sString )
	Dim sNewString 

	sNewString = Replace( sString, "'","\'" )

	FormatForJavaScript = sNewString
End Function 




'--------------------------------------------------------------------------------------------------
' void BreakOutAddress ByVal sAddress, ByRef sStreetNumber, ByRef sStreetName 
'--------------------------------------------------------------------------------------------------
Sub BreakOutAddress( ByVal sAddress, ByRef sStreetNumber, ByRef sStreetName )
	' Break out the number from the name, should be at first space from left
	Dim iPos

	iPos = InStr(sAddress, " ")
	If Not IsNull(iPos) And iPos > 0 Then
		sStreetNumber = Trim(Left(sAddress, (iPos - 1)))
		If IsNumeric(sStreetNumber) Then 
			sStreetName = Trim(Mid(sAddress,(iPos + 1)))
		Else
			' The first field is not a number, so this is just a street name or something
			sStreetNumber = ""
			sStreetName = sAddress
		End If 
	Else
		' no space so maybe this is a street name or something that will not work
		sStreetNumber = ""
		sStreetName = sAddress
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function IsValidAddress( sStreetNumber, sStreetName )
'--------------------------------------------------------------------------------------------------
Function IsValidAddress( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oAddress

'	sSql = "Select count(residentaddressid) as hits From egov_residentaddresses "
'	sSql = sSql & " Where residentstreetnumber = '" & track_dbsafe(sStreetNumber) & "' and residentstreetname = '" & track_dbsafe(sStreetName) & "' and orgid = " & iOrgID

'	New way to validate as of 4/3/2008
	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
	sSql = sSql & " WHERE residentstreetnumber = '" & track_dbsafe(sStreetNumber) & "' "
	sSql = sSql & " AND (ltrim(rtrim(residentstreetname)) = '" & track_dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & track_dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & track_dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) + ' ' + ltrim(rtrim(streetdirection)) = '" & track_dbsafe(sStreetName) & "' )"
	sSql = sSql & " AND orgid = " & iOrgID 

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSql, Application("DSN"), 0, 1

	If CLng(oAddress("hits")) > CLng(0) Then
		IsValidAddress = True 
	Else
		IsValidAddress =  False 
	End If 

	oAddress.Close
	Set oAddress = Nothing 

End Function 

'--------------------------------------------------------------------------------------------------
function IsValidAddress_byStreetName(p_orgid, sStreetName)
  dim sSQL, oAddress

    lcl_return = False

   'New way to validate as of 4/3/2008
	sSQL = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE (residentstreetname = '" & track_dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & track_dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & track_dbsafe(sStreetName) & "' "
	sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & track_dbsafe(sStreetName) & "' )"
	sSQL = sSQL & " AND orgid = " & p_orgid

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1

	if CLng(oAddress("hits")) > CLng(0) then
		lcl_return = True
	end if

	oAddress.close
	set oAddress = nothing 

    IsValidAddress_byStreetName = lcl_return

end function


'--------------------------------------------------------------------------------------------------
' boolean CartHasItems( iSessionID )
'--------------------------------------------------------------------------------------------------
Function CartHasItems( ByVal iSessionID )
	Dim sSql, oCart

	If iSessionID <> "" Then
		sSql = "SELECT COUNT(cartid) AS hits FROM egov_class_cart WHERE sessionid = " & iSessionID

		Set oCart = Server.CreateObject("ADODB.Recordset")
		oCart.Open sSql, Application("DSN"), 0, 1

		If Not oCart.EOF Then 
			If clng(oCart("hits")) > clng(0) Then
				CartHasItems = True 
			Else
				CartHasItems = False 
			End If 
		Else
			CartHasItems = False 
		End If 

		oCart.Close
		Set oCart = Nothing 
	Else
		CartHasItems = False
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean UserIsRootAdmin( iUserID )
'--------------------------------------------------------------------------------------------------
Function UserIsRootAdmin( ByVal iUserID )
	Dim sSql, oRs

	UserIsRootAdmin = False 
	sSql = "SELECT isnull(isrootadmin,0) as isrootadmin FROM users WHERE userid = " & iUserID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isrootadmin") Then 
			UserIsRootAdmin = True 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

function isFeatureOffline(p_feature)
 'If the user is ROOT ADMIN then bypass the check for any features that may be offline
'  if UserIsRootAdmin(session("userid")) then
'     lcl_feature_offline = "N"
'  else
     sSql = "SELECT distinct feature_offline "
     sSql = sSql & " FROM egov_organization_features "
     sSql = sSql & " WHERE feature_offline = 'Y' "
     sSql = sSql & " AND UPPER(feature) IN ('" & UCASE(REPLACE(p_feature,",","','")) & "') "

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSql, Application("DSN"),3,1

     if not rs.eof then
        lcl_feature_offline = "Y"
     else
        lcl_feature_offline = "N"
     end if
'  end if

	rs.Close
	Set rs = Nothing 

  isFeatureOffline = lcl_feature_offline

end function

'--------------------------------------------------------------------------------------------------
function changeBGColor(p_bgcolor,p_first_color,p_second_color)
 'Set up the row colors
  if p_first_color <> "" then
     lcl_first_color = p_first_color
  else
     lcl_first_color = "#eeeeee"
  end if

  if p_second_color <> "" then
     lcl_second_color = p_second_color
  else
     lcl_second_color = "#ffffff"
  end if

 'Determine which row color to display
  if p_bgcolor <> "" then
     lcl_bgcolor = p_bgcolor

     if UCASE(lcl_bgcolor) = UCASE(lcl_first_color) then
        lcl_bgcolor = lcl_second_color
     else
        lcl_bgcolor = lcl_first_color
     end if
  else
     lcl_bgcolor = lcl_first_color
  end if

  changeBGColor = lcl_bgcolor

end function


'--------------------------------------------------------------------------------------------------
' boolean FeatureIsOffline( p_feature )
'--------------------------------------------------------------------------------------------------
Function FeatureIsOffline( ByVal p_feature, ByVal iUserid )
	Dim sSql, oRs 

	'If the user is ROOT ADMIN then bypass the check for any features that may be offline
	If UserIsRootAdmin( iUserid ) Then 
		ParentFeatureIsOffline = False 
	Else 
		sSql = "SELECT UPPER(feature_offline) AS feature_offline FROM egov_organization_features "
		sSql = sSql & " WHERE parentfeatureid = 0 AND feature = '" & p_feature & "'"

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then
			If oRs("feature_offline") = "Y" Then 
				FeatureIsOffline = True 
			Else
				FeatureIsOffline = False 
			End If 
		Else 
			FeatureIsOffline = False 
		End  If 
		oRs.Close
		Set oRs = Nothing 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' void PageDisplayCheck( sFeature )
'--------------------------------------------------------------------------------------------------
Sub PageDisplayCheck( ByVal sFeature, ByVal sLevel, ByVal iUserid )
	' This sub handles feature offline functions from one point. Makes future changes easier

	If FeatureIsOffline( sFeature ) Then 
		response.redirect sLevel & "outage_feature_offline.asp"
	End If 

End Sub 

'--------------------------------------------------------------------------------------------------
function buildStreetAddress(sStreetNumber, sPrefix, sStreetName, sSuffix, sDirection)
  lcl_street_name = ""

  if trim(sStreetNumber) <> "" then
     lcl_street_name = trim(sStreetNumber)
  end if

  if trim(sPrefix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sPrefix)
     else
        lcl_street_name = trim(sPrefix)
     end if
  end if

  if trim(sStreetName) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sStreetName)
     else
        lcl_street_name = trim(sStreetName)
     end if
  end if

  if trim(sSuffix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sSuffix)
     else
        lcl_street_name = trim(sSuffix)
     end if
  end if

  if trim(sDirection) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sDirection)
     else
        lcl_street_name = trim(sDirection)
     end if
  end if

  buildStreetAddress = trim(lcl_street_name)

end Function

'------------------------------------------------------------------------------
sub GetAddressInfoNew(ByVal iOrgHasFeature_LargeAddressList, ByVal p_orgid, ByVal sStreetNumber, ByVal sStreetName, _
                      ByRef sNumber, ByRef sPrefix, ByRef sAddress, ByRef sSuffix, ByRef sDirection, _
                      ByRef sLatitude, ByRef sLongitude, ByRef sCity, ByRef sState, ByRef sZip, ByRef sCounty, _
                      ByRef sParcelID, ByRef sListedOwner, ByRef sLegalDescription, ByRef sResidentType, _
                      ByRef sRegisteredUserID, ByRef sValidStreet)

 'Set return variables
  sValidStreet      = "N"
  lcl_streetnumber  = dbsafe(sStreetNumber)
  lcl_streetname    = dbsafe(sStreetName)
		sNumber           = ""
  sPrefix           = ""
  sAddress          = ""
  sSuffix           = ""
  sDirection        = ""
	 sLatitude         = ""
  sLongitude        = ""
  sCity             = ""
  sState            = ""
  sZip              = ""
  sCounty           = ""
  sParcelID         = ""
  sListedOwner      = ""
  sLegalDescription = ""
  sResidentType     = ""
  sRegisteredUserID = ""

	 sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
  sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, residentcity, residentstate, residentzip, "
  sSQL = sSQL & " county, parcelidnumber, listedowner, legaldescription, residenttype, registereduserid "
  sSQL = sSQL & " FROM egov_residentaddresses "

  if iOrgHasFeature_LargeAddressList then
    'Format the streetname
     lcl_streetname = "'" & lcl_streetname & "'"

   	 sSQL = sSQL & " WHERE orgid = " & p_orgid
     sSQL = sSQL & " AND excludefromactionline = 0 "
     sSQL = sSQL & " AND UPPER(residentstreetnumber) = UPPER('" & lcl_streetnumber & "') "
     sSQL = sSQL & " AND (residentstreetname = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " )"
  else
    'Format the streetname
     if lcl_streetname = "" then
        lcl_streetname = 0
     end if

    	sSQL = sSQL & " WHERE residentaddressid = " & lcl_streetname
  end if

	 set oAddress = Server.CreateObject("ADODB.Recordset")
	 oAddress.Open sSQL, Application("DSN"), 3, 1
	
	 if not oAddress.eof then
   		sNumber           = trim(oAddress("residentstreetnumber"))
     sPrefix           = oAddress("residentstreetprefix")
   		sAddress          = oAddress("residentstreetname")
     sSuffix           = oAddress("streetsuffix")
     sDirection        = oAddress("streetdirection")
	 	  sLatitude         = oAddress("latitude")
   		sLongitude        = oAddress("longitude")
     sCity             = oAddress("residentcity")
     sState            = oAddress("residentstate")
     sZip              = oAddress("residentzip")
     sCounty           = oAddress("county")
     sParcelID         = oAddress("parcelidnumber")
     sListedOwner      = oAddress("listedowner")
     sLegalDescription = oAddress("legaldescription")
     sResidentType     = oAddress("residenttype")
     sRegisteredUserID = oAddress("registereduserid")
     sValidStreet      = "Y"
	 end if

	 oAddress.close
	 set oAddress = nothing

end sub

'----------------------------------------------------------------------------------------
' integer CleanAndCountForPayFlowPro( sParameter )
'----------------------------------------------------------------------------------------
Function CleanAndCountForPayFlowPro( ByRef sParameter )
	' cleans forbidden characters and returns the string length. Used by PayFlow Pro Functions

	sParameter = Replace(sParameter, Chr(34), "")
	sParameter = Replace(sParameter, Chr(13), "")
	sParameter = Replace(sParameter, Chr(10), "")
	sParameter = Replace(sParameter, "'", "")
	sParameter = Replace(sParameter, "&", "and")
	sParameter = Replace(sParameter, "=", "is")
	sParameter = Replace(sParameter, "</br>", ", ")
	sParameter = Replace(sParameter, "<br />", ", ")
	sParameter = Replace(sParameter, "<br>", ", ")
	sParameter = Replace(sParameter, ", ,", "")
	sParameter = Trim(sParameter)
	sParameter = Left(sParameter, 128)	' 128 characters is the PayPal limit on the Comment1 and Comment2 fields

	CleanAndCountForPayFlowPro = Len(sParameter)
End Function 


'--------------------------------------------------------------------------------------------------
' boolean OrgHasDisplay( iorgid, sDisplay )
'--------------------------------------------------------------------------------------------------
Function OrgHasDisplay( ByVal iOrgId, ByVal sDisplay )
	Dim sSql, oDisplay, blnReturnValue

	'SET DEFAULT
	blnReturnValue = False

	'LOOKUP passed display FOR the current ORGANIZATION 
	sSql = "SELECT COUNT(OD.displayid) AS display_count "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid "
	sSql = sSql & " AND orgid = " & iOrgId
	sSql = sSql & " AND D.display = '" & sDisplay & "' "

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If clng(oDisplay("display_count")) > 0 Then
 		'the ORGANIZATION HAS the Display
		  blnReturnValue = True
	End If
	
	oDisplay.Close 
	Set oDisplay = Nothing
	
	'set the RETURN  value
	OrgHasDisplay = blnReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' integer GetDisplayId( sDisplay )
'--------------------------------------------------------------------------------------------------
Function GetDisplayId( ByVal sDisplay )
	Dim sSql, oDisplay

	sSql = "SELECT displayid FROM egov_organization_displays WHERE display = '" & sDisplay & "' "

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If Not oDisplay.EOF Then
		GetDisplayId = clng(oDisplay("displayid"))
	Else
		GetDisplayId = clng(0)
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' string GetOrgDisplayWithId( iorgid, iDisplayId )
'--------------------------------------------------------------------------------------------------
Function GetOrgDisplayWithId( ByVal iOrgId, ByVal iDisplayId, ByVal bUsesDisplayName )
	Dim sSql, oRs, sField

	If bUsesDisplayName Then
		sField = "displayname"
	Else
		sField = "displaydescription"
	End If 

	' LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT ISNULL(OD." & sField & ", D." & sField & ") AS displayfield "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid AND orgid = " & iOrgId & " AND D.displayid = " & iDisplayId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		' the ORGANIZATION HAS the Display
		GetOrgDisplayWithId = oRs("displayfield")
	Else
		GetOrgDisplayWithId = GetDisplayName( iDisplayId )
	End If
	
	oRs.Close 
	Set oRs = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' string GetDisplayName( iDisplayId )
'--------------------------------------------------------------------------------------------------
Function GetDisplayName( ByVal iDisplayId )
	Dim sSql, oRs

	'LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT displayname FROM egov_organization_displays WHERE displayid = " & iDisplayId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
	 	'the ORGANIZATION HAS the Display
		  GetDisplayName = oRs("displayname")
	Else
		GetDisplayName = ""
	End If
	
	oRs.Close 
	Set oRs = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' string GetFeatureName( sFeature )
'------------------------------------------------------------------------------------------------------------
Function GetFeatureName( ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid AND FO.orgid = " & iorgid & " AND feature = '" & sFeature & "'" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
  		GetFeatureName = oRs("featurename")
	Else
		GetFeatureName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
function dbready_string( ByVal p_value, ByVal p_length )
	Dim lcl_return

  lcl_return = ""

  if p_value <> "" AND p_length <> "" then
     lcl_return = trim(p_value)
     lcl_return = replace(lcl_return,"<","&lt;")
     lcl_return = replace(lcl_return,">","&gt;")

    'Verify the length
     if len(lcl_return) > p_length then
        lcl_return = mid(lcl_return,1,p_length)
     end if

     lcl_return = replace(lcl_return,"'","''")

  end if

  dbready_string = lcl_return

end function


'------------------------------------------------------------------------------
function dbready_date(p_value)
	Dim lcl_return

  lcl_return = False

  if p_value <> "" then
     lcl_return = trim(p_value)

     if isDate(lcl_return) then
        lcl_return = True
     end if
  end if

  dbready_date = lcl_return

end function


'------------------------------------------------------------------------------
function dbready_number(p_value)
	Dim lcl_return

  lcl_return = False

  if p_value <> "" then
     lcl_return = trim(p_value)

     if isNumeric(lcl_return) then
        lcl_return = True
     end if
  end if

  dbready_number = lcl_return

end Function


'------------------------------------------------------------------------------------------------------------
' Function GetGoogleApiKey( iOrgId )
'------------------------------------------------------------------------------------------------------------
Function GetGoogleMapApiKey( iOrgId )
	Dim sSql, oRs

	sSql = "SELECT googlemapapikey FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetGoogleMapApiKey = oRs("googlemapapikey")
	Else
		GetGoogleMapApiKey = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' void RunSQLStatement sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQLStatement( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush
	session("RunSQLStatement") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

	session("RunSQLStatement") = ""

End Sub 


'-------------------------------------------------------------------------------------------------
' integer RunIdentityInsertStatement( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsertStatement( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush
	session("InsertStatement") = sInsertStatement

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")

	oInsert.Close
	Set oInsert = Nothing
	session("InsertStatement") = ""

	RunIdentityInsertStatement = iReturnValue

End Function
 

'------------------------------------------------------------------------------
' integer GetDefaultRelationShipId( iOrgid )
'------------------------------------------------------------------------------
Function GetDefaultRelationShipId( ByVal iOrgid )
	Dim sSql, oRs

	sSql = "SELECT relationshipid FROM egov_familymember_relationships WHERE orgid = " & iorgid & " AND isdefault = 1"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultRelationShipId = oRs("relationshipid") 
	Else
		GetDefaultRelationShipId = 0 
	End if
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' boolean UserIsMissingKeyData( iUserid )
'------------------------------------------------------------------------------
Function UserIsMissingKeyData( ByVal iUserid )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname, userhomephone FROM egov_users WHERE userid = " & iUserid 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If IsNull(oRs("userfname")) Or IsNull(oRs("userlname")) Or IsNull(oRs("userhomephone")) Then 
			UserIsMissingKeyData = True 
		Else
			UserIsMissingKeyData = False 
		End If 
	Else
		UserIsMissingKeyData = True  
	End if
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' string  FormatPhoneNumber(  Number )
'------------------------------------------------------------------------------
Function FormatPhoneNumber( ByVal Number )
	If Len(Number) = 10 Then
		FormatPhoneNumber = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhoneNumber = Number
	End If
End Function

'--------------------------------------------------------------------------------------------------
Function BuildHTMLMessage( ByVal sBody )
	'Build email message
 	Dim lcl_return

 	lcl_return = lcl_return & "<html>" & vbcrlf
 	lcl_return = lcl_return & "<head>" & vbcrlf
 	'lcl_return = lcl_return & "  <STYLE> td {font-family: arial,tahoma; font-size: 12px; color: #000000;} </STYLE>"
 	lcl_return = lcl_return & "</head>" & vbcrlf
 	lcl_return = lcl_return & "<body bgcolor=""#efefef"">" & vbcrlf
 	lcl_return = lcl_return & "<font face=""helvetica, arial"">" & vbcrlf
 	lcl_return = lcl_return & "<p style=""margin:0px""></p>" & vbcrlf
 	lcl_return = lcl_return & "<table bordercolor=""#4A9E9F"" bgcolor=""#ffffff"" cellspacing=""0"" cellpadding=""5"" width=""95%"" align=""center"" border=""2"" valign=""top"">" & vbcrlf
  lcl_return = lcl_return & "<tr>" & vbcrlf
  lcl_return = lcl_return & "<td style=""font-family:arial,tahoma; font-size:12px; color:#000000;"">" & vbcrlf
  lcl_return = lcl_return & sBody & vbcrlf
 	lcl_return = lcl_return & "<center>" & vbcrlf
 	lcl_return = lcl_return & "<br />" & vbcrlf
 	lcl_return = lcl_return & "<hr color=""black"" size=""1"" width=""95%"">" & vbcrlf
 	lcl_return = lcl_return & "<font size=""-2"">Copyright 2004 - " & year(now) & ". <i>Electronic Commerce</i> Link, Inc. dba <i>EC</i> Link.</font>" & vbcrlf
  lcl_return = lcl_return & "</center>" & vbcrlf
  lcl_return = lcl_return & "</td>" & vbcrlf
  lcl_return = lcl_return & "</tr>" & vbcrlf
  lcl_return = lcl_return & "</table>" & vbcrlf
  lcl_return = lcl_return & "</font>" & vbcrlf
 	lcl_return = lcl_return & "</body>" & vbcrlf
 	lcl_return = lcl_return & "</html>" & vbcrlf

 	BuildHTMLMessage = lcl_return

End Function

'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
function checkSendEmail(iDoNotSendSelfEmail,iDoNotSendAllEmail,iUserCompareEmail)

  lcl_return = "Y"

 'Check to see if the user does not want to send the email to self/all assigned
  lcl_doNotSendSelfEmail = iDoNotSendSelfEmail
  lcl_doNotSendAllEmail  = iDoNotSendAllEmail

 'Check to see if the "do not send to all" is clicked
  if lcl_doNotSendAllEmail = "on" then
     lcl_return = "N"
  end if

 'Check to see if the "do not send to self" is clicked
  if lcl_doNotSendSelfEmail = "on" then
     lcl_sessionuser_email = getUserEmail(iuserid)

     if UCASE(lcl_sessionuser_email) = UCASE(iUserCompareEmail) then
        lcl_return = "N"
     end if
  end if

  checkSendEmail = lcl_return

end function


'------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
Function GetTimeOffset(sOrgID)
	Dim sSql, oTime

	sSql = "SELECT T.gmtoffset FROM Organizations O, TimeZones T Where O.OrgTimeZoneID = T.TimeZoneID And O.orgid = " & sOrgID

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSql, Application("DSN"), 3, 1

	If Not oTime.EOF Then 
  		GetTimeOffset =  clng(oTime("gmtoffset"))
	Else
  		GetTimeOffset = clng(0)
	End If 

	oTime.Close
	Set oTime = Nothing 

End Function

'------------------------------------------------------------------------------
function ConvertDateTimetoTimeZone( ByVal sOrgID )
  Dim sSql, oTimeOffset, lcl_localdate, datCurrentDate

  datCurrentDate = Now()

  'Get the local date/time for the timezone
  sSql = "SELECT dbo.GetLocalDate(" & sOrgID & ", '" & datCurrentDate & "') AS localDate "
  sSql = sSql & " FROM organizations "
  sSql = sSql & " WHERE orgid = " & sOrgID

  set oTimeOffset = Server.CreateObject("ADODB.Recordset")
  oTimeOffset.Open sSql, Application("DSN"), 3, 1

  if not oTimeOffset.eof then
    lcl_localdate = oTimeOffset("localDate")
  else
   lcl_localdate = "NULL"
  end if

  oTimeOffset.close
  set oTimeOffset = nothing

  ConvertDateTimetoTimeZone = lcl_localdate

end function

'------------------------------------------------------------------------------
sub displayAddThisButton(p_orgid)

 '-----------------------------------------------------------------------------
 'Must have the following .js file in your code.
  '<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
  '<script type="text/javascript">var addthis_pub="cschappacher";</script>
 '-----------------------------------------------------------------------------

 'Determine if the org has the feature turned on.
  sSql = "SELECT count(FO.featureid) AS feature_count "
  sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
  sSql = sSql & " WHERE FO.featureid = F.featureid "
  sSql = sSql & " AND FO.orgid = " & p_orgid
  sSql = sSql & " AND UPPER(F.feature) = 'BUTTON_ADDTHIS' "

 	set oHasFeature = Server.CreateObject("ADODB.Recordset")
	 oHasFeature.Open  sSql, Application("DSN"), 3, 1

 'If the org has the feature then display the ADD THIS button.
  if clng(oHasFeature("feature_count")) > 0 then
     response.write "<a href=""http://www.addthis.com/bookmark.php?v=20"" onmouseover=""return addthis_open(this, '', '[URL]', '[TITLE]')"" onmouseout=""addthis_close()"" onclick=""return addthis_sendto()"">" & vbcrlf
     response.write "<img src=""https://s7.addthis.com/static/btn/lg-addthis-en.gif"" width=""125"" height=""16"" alt=""Bookmark and Share"" style=""border:0""/>" & vbcrlf
     response.write "</a>" & vbcrlf
  end if

  oHasFeature.close
  set oHasFeature = nothing

end sub

'------------------------------------------------------------------------------
sub displayAddThisButtonNew(p_orgid)

 '-----------------------------------------------------------------------------
 'Must have the following .js file in your code.
  '<script type="text/javascript">var addthis_config = {"data_track_clickback":true};</script>
  '<script type="text/javascript" src="https://s7.addthis.com/js/250/addthis_widget.js#pubid=egovlink"></script>
 '-----------------------------------------------------------------------------

 'Determine if the org has the feature turned on.
  sSql = "SELECT count(FO.featureid) AS feature_count "
  sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
  sSql = sSql & " WHERE FO.featureid = F.featureid "
  sSql = sSql & " AND FO.orgid = " & p_orgid
  sSql = sSql & " AND UPPER(F.feature) = 'BUTTON_ADDTHIS' "

 	set oHasFeature = Server.CreateObject("ADODB.Recordset")
	 oHasFeature.Open  sSql, Application("DSN"), 3, 1

 'If the org has the feature then display the ADD THIS button.
  if clng(oHasFeature("feature_count")) > 0 then
     response.write "<div id=""addthis"" class=""addthis_toolbox addthis_default_style"">" & vbcrlf
     response.write "<a class=""addthis_button_facebook""></a>" & vbcrlf
     response.write "<a class=""addthis_button_twitter""></a>" & vbcrlf
     response.write "<a class=""addthis_button_email""></a>" & vbcrlf
     response.write "<a class=""addthis_button_print""></a>" & vbcrlf
     response.write "<a class=""addthis_button_compact""></a>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  oHasFeature.close
  set oHasFeature = nothing

end sub


'------------------------------------------------------------------------------
sub displayYahooBuzzButton(p_orgid)

'NOTE: Yahoo Buzz has been shut down.


 'Determine if the org has the feature turned on.
'  sSql = "SELECT count(FO.featureid) AS feature_count "
'  sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
'  sSql = sSql & " WHERE FO.featureid = F.featureid "
'  sSql = sSql & " AND FO.orgid = " & p_orgid
'  sSql = sSql & " AND UPPER(F.feature) = 'BUTTON_YAHOOBUZZ' "

' 	set oHasFeature = Server.CreateObject("ADODB.Recordset")
'	 oHasFeature.Open  sSql, Application("DSN"), 3, 1

 'If the org has the feature then display the YAHOO BUZZ button.
'  if clng(oHasFeature("feature_count")) > 0 then
'     response.write "<script type=""text/javascript"" src=""http://d.yimg.com/ds/badge2.js"" badgetype=""small"">ARTICLEURL</script>" & vbcrlf
'  end if

'  oHasFeature.close
'  set oHasFeature = nothing

response.write ""

end sub

'------------------------------------------------------------------------------
sub checkForRSSFeed(p_orgid, p_rss_feedid, p_featureid, p_feedname, p_egovwebsiteurl)
  lcl_feedname  = ""
  lcl_featureid = 0

  if p_egovwebsiteurl <> "" then
     lcl_egovwebsiteurl = p_egovwebsiteurl & "/rssfeeds.asp"
  else
     lcl_egovwebsiteurl = "#"
  end if

  if p_rss_feedid <> "" OR p_featureid <> "" OR p_feedname <> "" then

    'Determine if this feed exists and is active.
     lcl_exists = "N"

    'First get the feature name of the RSS Feed (egov_rssfeed.feature) using either the rss_feedid or feedname passed in.
     if p_rss_feedid <> "" OR p_feedname <> "" then
        lcl_rss_featureid = 0

        if p_rss_feedid <> "" then
           lcl_rss_feature = getRSSFeature(p_rss_feedid, "")
        else

           lcl_feedname = ""

           if p_feedname <> "" then
              lcl_feedname = UCASE(p_feedname)
           end if

           lcl_rss_feature = getRSSFeature("",lcl_feedname)
        end if

        if lcl_rss_feature <> "" then
           lcl_featureid = getFeatureID(lcl_rss_feature)
        end if
     else
        lcl_featureid = p_featureid
     end if

     sSql = "SELECT distinct 'Y' as feedexists "
     sSql = sSql & " FROM egov_rssfeeds rf, egov_organization_features f, egov_organizations_to_features FO "
     sSql = sSql & " WHERE UPPER(rf.feature) = UPPER(f.feature) "
     sSql = sSql & " AND f.featureid = fo.featureid "
     sSql = sSql & " AND FO.orgid = " & p_orgid
     sSql = sSql & " AND upper(rf.feature) = (select upper(of2.feature) "
     sSql = sSql &                          " from egov_organization_features of2 "
     sSql = sSql &                          " where of2.featureid = " & lcl_featureid & ") "


'     if p_featureid <> "" then
'        sSql = sSql & " AND upper(rf.feature) = (select upper(of2.feature) "
'        sSql = sSql &                          " from egov_organization_features of2 "
'        sSql = sSql &                          " where of2.featureid = " & p_featureid & ") "
'     elseif p_rss_feedid <> "" then
'        sSql = sSql & " AND rf.feedid = " & p_rss_feedid
'     else
'        sSql = sSql & " AND UPPER(rf.feedname) = '" & lcl_feedname & "'"
'     end if

     sSql = sSql & " AND (select count(rss.rssid) "
     sSql = sSql &      " from egov_rss rss "
     sSql = sSql &      " where rss.orgid = " & p_orgid
     sSql = sSql &      " and rss.feedid = rf.feedid) > 0 "
'response.write sSQL
     set oFeedExists = Server.CreateObject("ADODB.Recordset")
     oFeedExists.Open sSql, Application("DSN"), 3, 1

     if not oFeedExists.eof then
        lcl_exists = oFeedExists("feedexists")
     end if

     'oFeedExists.close
     set oFeedExists = nothing

    'If the feed exists then display the image/link.
     if lcl_exists = "Y" then
        response.write "<a href="""  & lcl_egovwebsiteurl & """ target=""_top"">" & vbcrlf
        response.write "<img src=""" & replace(p_egovwebsiteurl,"http:","https:") & "/images/communitylink/socialsites/icon_rss.png"" border=""0"" alt=""Subscribe to this RSS Feed"" />" & vbcrlf
        response.write "</a>" & vbcrlf
     end if

  end if

end sub

'------------------------------------------------------------------------------
function getRSSFeature(iFeedID, iFeedName)
  lcl_return = ""

  if iFeedID <> "" OR iFeedName <> "" then
     sSQL = "SELECT feature "
     sSQL = sSQL & " FROM egov_rssfeeds "

     if iFeedID <> "" then
        sSQL = sSQL & " WHERE feedid = " & iFeedID
     else
        lcl_feedname = iFeedName
        lcl_feedname = UCASE(lcl_feedname)
        lcl_feedname = "'" & track_dbsafe(lcl_feedname) & "'"

        sSQL = sSQL & " WHERE UPPER(feedname) = " & lcl_feedname
     end if

    	set oGetRSSFeature = Server.CreateObject("ADODB.Recordset")
   	 oGetRSSFeature.Open sSQL, Application("DSN"), 3, 1

     if not oGetRSSFeature.eof then
        lcl_return = oGetRSSFeature("feature")
     end if

     oGetRSSFeature.close
     set oGetRSSFeature = nothing
  end if

  getRSSFeature = lcl_return

end function

'------------------------------------------------------------------------------
function getFeatureID(p_feature)
  lcl_return = 0

  if p_feature <> "" then
     sSql = "SELECT featureid "
     sSql = sSql & " FROM egov_organization_features "
     sSql = sSql & " WHERE UPPER(feature) = '" & UCASE(p_feature) & "' "

    	set oGetFeatureID = Server.CreateObject("ADODB.Recordset")
   	 oGetFeatureID.Open sSql, Application("DSN"), 3, 1

     if not oGetFeatureID.eof then
        lcl_return = oGetFeatureID("featureid")
     end if

     oGetFeatureID.close
     set oGetFeatureID = nothing
  end if

  getFeatureID = lcl_return
end function

'------------------------------------------------------------------------------
function getCommentsFormID(p_orgid, p_featureid, p_feature)
  lcl_return    = 0
  lcl_featureid = 0

  if p_featureid <> "" then
     lcl_featureid = p_featureid
  else
     if p_feature <> "" then
        lcl_featureid = getFeatureID(p_feature)
     end if
  end if

  sSql = "SELECT CL_postcomments_formid "
  sSql = sSql & " FROM egov_organizations_to_features "
  sSql = sSql & " WHERE orgid = "   & p_orgid
  sSql = sSql & " AND featureid = " & lcl_featureid

  set oCLFormID = Server.CreateObject("ADODB.Recordset")
  oCLFormID.Open sSql, Application("DSN"), 3, 1

  if not oCLFormID.eof then
     lcl_return = oCLFormID("CL_postcomments_formid")
  end if

  oCLFormID.close
  set oCLFormID = nothing

  getCommentsFormID = lcl_return

end function

'------------------------------------------------------------------------------
function getCommentsLabel(p_orgid, p_featureid, p_feature)
  lcl_return    = ""
  lcl_featureid = 0

  if p_featureid <> "" then
     lcl_featureid = p_featureid
  else
     if p_feature <> "" then
        lcl_featureid = getFeatureID(p_feature)
     end if
  end if

  sSQL = "SELECT CL_postcomments_label "
  sSQL = sSQL & " FROM egov_organizations_to_features "
  sSQL = sSQL & " WHERE orgid = "   & p_orgid
  sSQL = sSQL & " AND featureid = " & lcl_featureid

  set oCLLabel = Server.CreateObject("ADODB.Recordset")
  oCLLabel.Open sSQL, Application("DSN"), 3, 1

  if not oCLLabel.eof then
     lcl_return = oCLLabel("CL_postcomments_label")
  end if

  oCLLabel.close
  set oCLLabel = nothing

  getCommentsLabel = lcl_return

end function

'------------------------------------------------------------------------------
function getUserEmail(iUserID)
  lcl_return = ""

  if iUserID <> "" then
     sSql = "SELECT email "
     sSql = sSql & " FROM users "
     sSql = sSql & " WHERE userid = " & iUserID

     set oUserEmail = Server.CreateObject("ADODB.Recordset")
   	 oUserEmail.Open sSql, Application("DSN"), 3, 1

     if not oUserEmail.eof then
        lcl_return = oUserEmail("email")
     end if

     oUserEmail.close
     set oUserEmail = nothing
  end if

  getUserEmail = lcl_return

end function

'------------------------------------------------------------------------------
function checkAccessMethod(iHttpUserAgent)

  lcl_return = ""

 'Determine if the user is accessing the site via a desktop or mobile device.
  if iHttpUserAgent <> "" then

     iHttpUserAgent = UCASE(iHttpUserAgent)

     if instr(iHttpUserAgent,"BLACKBERRY") > 0 then
        lcl_return = "BLACKBERRY"
     elseif instr(iHttpUserAgent,"IPHONE") > 0 then
        lcl_return = "IPHONE"
     elseif instr(iHttpUserAgent,"ANDROID") > 0 then
        lcl_return = "ANDROID"
     end if

  end if

  checkAccessMethod = lcl_return

end function

'------------------------------------------------------------------------------
sub displaySwitchViewModeLink(p_orgname, p_viewMode)

 'Determine which view we are going to SWITCH TO.
  if p_viewMode = "M" then
     lcl_label   = "Standard"
     lcl_newMode = "S"
  else
     lcl_label   = "Mobile"
     lcl_newMode = "M"
  end if

  response.write "View " & p_orgname & " in:<br />" & vbcrlf
  response.write "<table border=""1"" bordercolor=""#000000"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td bgcolor=""#efefef"">" & vbcrlf
  response.write "          <a href=""communitylink.asp?setDeviceViewMode=" & lcl_newMode & """ style=""color:#0000ff"">" & lcl_label & "</a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSql = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
  set oDTBDebug = Server.CreateObject("ADODB.Recordset")
  oDTBDebug.Open sSql, Application("DSN"), 3, 1

  set oDTBDebug = nothing
end Sub


'--------------------------------------------------------------------------------------------------
' Function generateRequestID( tmpLength )
'--------------------------------------------------------------------------------------------------
'Create the Unique ID - This is for PayPal, so do not change
Function generateRequestID( tmpLength )

	Randomize Timer
  	Dim tmpCounter, tmpGUID
  	Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  	For tmpCounter = 1 To tmpLength
    	tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  	Next
  	generateRequestID = tmpGUID

End Function

'------------------------------------------------------------------------------
function formatArticle(iArticle)

  lcl_return = ""

  if iArticle <> "" then
     lcl_return = iArticle
     lcl_return = replace(lcl_return,chr(13),"<br />")
     lcl_return = replace(lcl_return,"’","'")
  end if

  formatArticle = lcl_return

end function

'------------------------------------------------------------------------------
sub getDelegateInfo(ByVal iUserID, ByRef lcl_delegateid, ByRef lcl_delegate_username, ByRef lcl_delegate_useremail)
  lcl_delegateid         = 0
  lcl_delegate_username  = ""
  lcl_delegate_useremail = ""

  if iUserID <> "" then
     sSql = "SELECT userid, firstname, lastname, email "
     sSql = sSql & " FROM users "
     sSql = sSql & " WHERE userid = (select u2.delegateid "
     sSql = sSql &                 " from users u2 "
     sSql = sSql &                 " where userid = " & iUserID & ") "

     set oDelegateInfo = Server.CreateObject("ADODB.Recordset")
     oDelegateInfo.Open sSql, Application("DSN"), 1, 3

     if not oDelegateInfo.eof then
        lcl_delegateid         = oDelegateInfo("userid")
        lcl_delegate_username  = oDelegateInfo("firstname") & " " & oDelegateInfo("lastname")
        lcl_delegate_useremail = oDelegateInfo("email")
     end if

     oDelegateInfo.close
     set oDelegateInfo = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub setupSendToAndDelegateEmails(ByVal p_sendTo_email, ByVal p_delegate_email, ByRef lcl_email_sendto, ByRef lcl_email_cc)

  lcl_emailaddress_sendto   = ""
  lcl_emailaddress_delegate = ""
  lcl_sendEmailToUser       = "N"
  lcl_sendEmailToDelegate   = "N"
  lcl_email_sendto          = ""
  lcl_email_cc              = ""
  sEmailTo                  = ""
  sEmailDelegate            = ""

 'Setup the SENDTO, CC, and check for a DELEGATE
  if p_sendTo_email <> "" then
     lcl_emailaddress_sendto = formatSendToEmail(p_sendTo_email)

     if isValidEmail(lcl_emailaddress_sendto) then
        sEmailTo            = lcl_emailaddress_sendto
        lcl_sendEmailToUser = "Y"
     end if
  end if

  if p_delegate_email <> "" then
     lcl_emailaddress_delegate = formatSendToEmail(p_delegate_email)

     if isValidEmail(lcl_emailaddress_delegate) then
        sEmailDelegate          = lcl_emailaddress_delegate
        lcl_sendEmailToDelegate = "Y"
     end if
  end if

 'Determine who we are sending the email to.
  if lcl_sendEmailToUser = "Y" AND lcl_sendEmailToDelegate = "Y" then
     lcl_email_sendto = sEmailDelegate
     lcl_email_cc     = sEmailTo
  else
    'If there is no delegate or the delegate's email is invalid
     if lcl_sendEmailToDelegate <> "Y" then

       'The user and cc has a valid email
        if lcl_sendEmailToUser = "Y" then
           lcl_email_sendto = sEmailTo
        end if
     else
        lcl_email_sendto = sEmailDelegate

        if lcl_sendEmailToUser = "Y" then
           lcl_email_cc = sEmailTo
        end if
     end if
  end if
end sub

'------------------------------------------------------------------------------
function getStartingFolder( ByVal iStartFolder )
 lcl_return = "/published_documents"

'Determine which folder to start in
 if trim(iStartFolder) <> "" then
    if ucase(trim(iStartFolder)) = "CITY_ROOT" then
       lcl_return = ""
    else
       lcl_return = "/" & iStartFolder
    end if
 end if

 getStartingFolder = lcl_return

end function

'------------------------------------------------------------------------------
function getActionLineTrackingNumber(p_action_autoid)
  lcl_return = ""

  if p_action_autoid <> "" then
     sSql = "SELECT [Tracking Number] AS tracking_number"
     sSql = sSql & " FROM egov_rpt_actionline "
     sSql = sSql & " WHERE action_autoid = " & p_action_autoid

     set oGetTrackNum = Server.CreateObject("ADODB.Recordset")
     oGetTrackNum.Open sSql, Application("DSN"), 1, 3

     if not oGetTrackNum.eof then
        lcl_return = oGetTrackNum("tracking_number")
     end if

     oGetTrackNum.close
     set oGetTrackNum = nothing

  end if

  getActionLineTrackingNumber = lcl_return

end function

'------------------------------------------------------------------------------
sub formatEventDateTime(ByVal p_date, ByVal p_end_date, ByRef sDate1, ByRef sDate2)

 'Used to trim seconds from dates displayed on calendar pages.
		sDate1 = cStr(p_date)
		sDate2 = cStr(p_end_date)

		iTrimDate1 = clng(InStrRev(sDate1,":"))
		iTrimDate2 = clng(InStrRev(sDate2,":"))

	'Retrieves AM/PM, trims final :00 and builds string
		if iTrimDate1 > 0 then
  			sTemp  = Right(sDate1, 2)
		  	sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
				 sTemp  = ""
  end if

		if iTrimDate2 > 0 then
     sTemp  = Right(sDate2, 2)
     sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
	    sTemp  = ""
  end if

end sub

'------------------------------------------------------------------------------
function containsApostrophe(p_value)

  lcl_return = False

  if p_value <> "" then
     lcl_value = p_value

     if isnumeric(lcl_value) then
        lcl_value = CStr(lcl_value)
     end if

     if instr(lcl_value,"'") > 0 then
        lcl_return = True
     end if
  end if

  containsApostrophe = lcl_return

end Function


'--------------------------------------------------------------------------------------------------
' boolean PaymentGatewayRequiresFeeCheck( iOrgId )
'--------------------------------------------------------------------------------------------------
Function PaymentGatewayRequiresFeeCheck( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT requiresfeecheck FROM egov_payment_gateways G, Organizations O "
	sSql = sSql & "WHERE G.paymentgatewayid = O.OrgPaymentGateway AND O.orgid = " & iOrgId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("requiresfeecheck") Then
			PaymentGatewayRequiresFeeCheck = True 
		Else
			PaymentGatewayRequiresFeeCheck = False 
		End If 
	Else
		PaymentGatewayRequiresFeeCheck = False 
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function


'--------------------------------------------------------------------------------------------------
' boolean CitizenPaysFee( iOrgId )
'--------------------------------------------------------------------------------------------------
Function CitizenPaysFee( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT citizenpaysfee FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("citizenpaysfee") Then
			CitizenPaysFee = True 
		Else
			CitizenPaysFee = False 
		End If 
	Else
		CitizenPaysFee = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPNPFee( dPurchaseAmount, dFeeAmount, sErrorMsg )
'--------------------------------------------------------------------------------------------------
Function GetPNPFee( ByVal dPurchaseAmount, ByRef dFeeAmount, ByRef sErrorMsg )
	Dim parmList, objWinHttp, transResponse, sStatus

	'dFeeAmount = CDbl("2.95")
	'sErrorMsg = "none"
	'GetPNPFee = True 

	parmList = "chkamount=" & FormatNumber(CDbl(dPurchaseAmount),2,,,0)

	Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

	objWinHttp.Open "GET", sEgovWebsiteURL & "/payment_processors/pnpfeecheck.aspx?" & parmList, False

	objWinHttp.setRequestHeader "Content-Type", "text/namevalue"

	objWinHttp.Send parmList

	If objWinHttp.Status = 200 Then 
		' Get the text of the response.
		transResponse = objWinHttp.ResponseText
	End If 

	' Trash our object now that we are finished with it.
	Set objWinHttp = Nothing

	sStatus = GetPNPResponseValue(transResponse, "status")
	sErrorMsg = GetPNPResponseValue(transResponse, "errors")

	If sStatus <> "success" Then
		GetPNPFee = False 
	Else
		dFeeAmount = GetPNPResponseValue(transResponse, "fee")
		dFeeAmount = CDbl(dFeeAmount)
		GetPNPFee = True 
	End If 

End Function


'------------------------------------------------------------------------------
' string GetPNPResponseValue( sResponse, sParamName )
'------------------------------------------------------------------------------
Function GetPNPResponseValue( ByVal sResponse, ByVal sParamName )
	Dim curString, name, value, varString, MyValue

	curString = sResponse
	MyValue = ""

	Do While Len(curString) <> 0

		If InStr(curString,"&") Then
  			varString = Left(curString, InStr(curString , "&" ) -1)
 		Else 
  			varString = curString
 		End If 
 
 		name = Left(varString, InStr(varString, "=" ) -1)
 		value = Right(varString, Len(varString) - (Len(name)+1))

  		If UCase(name) = UCase(sParamName) Then 
  			MyValue = value
 			Exit Do
  		End If 
  	
  		If Len(curString) <> Len(varString) Then 
  			curString = Right(curString, Len(curString) - (Len(varString)+1))
  		Else 
  			curString = ""
  		End If 
 
	Loop

	GetPNPResponseValue = MyValue

End Function


'------------------------------------------------------------------------------
' integer CreatePaymentsControlRow( sLogEntry, sAppSide, sFeature )
'------------------------------------------------------------------------------
Function CreatePaymentsControlRow( ByVal sLogEntry, ByVal sAppSide, ByVal sFeature )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iOrgID & ", " & sAppSide & ", " & sFeature & ", '" & dbready_string(sLogEntry,500) & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunIdentityInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentsControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' void AddToPaymentsLog iPaymentControlNumber, sLogEntry, sAppSide, sFeature 
'------------------------------------------------------------------------------
Sub AddToPaymentsLog( ByVal iPaymentControlNumber, ByVal sLogEntry, ByVal sAppSide, ByVal sFeature  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgID & ", " & sAppSide & ", " & sFeature & ", '" & dbready_string(sLogEntry,500) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub 


'------------------------------------------------------------------------------
' integer StoreGatewayError( iOrgId, sPaymentGateway, sFeature, sError, sFeeAmount )
'------------------------------------------------------------------------------
Function StoreGatewayError( ByVal iOrgId, ByVal sPaymentGateway, ByVal sFeature, ByVal sError, ByVal sAmount )
	Dim sSql, iGatewayErrorId

	If LCase(sAmount) <> "null" Then
		sAmount = "'" & sAmount & "'"
	End If 

	sSql = "INSERT INTO egov_paymentgatewayerrors ( orgid, paymentgateway, feature, errormessage, amount) VALUES ( "
	sSql = sSql & iOrgId & ", '" & dbready_string( sPaymentGateway, 50) & "', '" & dbready_string( sFeature, 50) & "', '"
	sSql = sSql & dbready_string( sError, 1000) & "', " & sAmount & " )"

	iGatewayErrorId = RunIdentityInsertStatement( sSql )

	StoreGatewayError = iGatewayErrorId

End Function 


'------------------------------------------------------------------------------
' string GetGatewayErrorMsg( iGatewayErrorId )
'------------------------------------------------------------------------------
Function GetGatewayErrorMsg( ByVal iGatewayErrorId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(errormessage, '') AS errormessage FROM egov_paymentgatewayerrors "
	sSql = sSql & "WHERE gatewayerrorid = " & iGatewayErrorId & " AND orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetGatewayErrorMsg = oRs("errormessage")
	Else
		GetGatewayErrorMsg = "No error message."
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'------------------------------------------------------------------------------
' double GetProcessingFee( iPaymentId )
'------------------------------------------------------------------------------
Function GetProcessingFee( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(processingfee,0.00) AS processingfee FROM egov_verisign_payment_information "
	sSql = sSql & "WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetProcessingFee = CDbl(oRs("processingfee"))
	Else
		GetProcessingFee = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
function checkAccessToList( ByVal iUserID, ByVal iOrgID, ByVal iListType )
	Dim sSql, oCheckListAccess, lcl_return

	lcl_return = False

	'Determine the list type
	if iListType <> "" then

		sSql = "SELECT isDoNotKnockVendor_" & iListType & " as 'ListAccess' "
		sSql = sSql & " FROM egov_users "
		sSql = sSql & " WHERE orgid = " & CLng(iOrgID)
		sSql = sSql & " AND userid = " & CLng(iUserID)

		set oCheckListAccess = Server.CreateObject("ADODB.Recordset")
		oCheckListAccess.Open sSql, Application("DSN"), 3, 1

		if not oCheckListAccess.eof then
			lcl_return = oCheckListAccess("ListAccess")
		end if

		oCheckListAccess.close
		set oCheckListAccess = nothing

	end if

	checkAccessToList = lcl_return

end function


'--------------------------------------------------------------------------------------------------
' string StripTags( source )
'--------------------------------------------------------------------------------------------------
Function StripTags( ByVal sSource )
	Dim x, sReturn, iSourceSize, sChar, bInside

	iSourceSize = Len(sSource)
	sReturn = ""
	bInside = False 

	For x = 1 To iSourceSize
		sChar = Mid(sSource, x, 1)

		If sChar = "<" Then 
			bInside = True
		ElseIf sChar = ">" Then 
			bInside = False 
		Else 
			If Not bInside Then 
				sReturn = sReturn & sChar
			End If 
		End If 
	Next 

	StripTags = sReturn

End Function 


'------------------------------------------------------------------------------
' void DisplayGenderPicks sGender 
'------------------------------------------------------------------------------
Sub DisplayGenderPicks( ByVal sElement, ByVal sGenderMatch )
	Dim sSql, oRs

	sSql = "SELECT gender, genderdescription FROM egov_user_genders ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""" & sElement & """ name=""" & sElement & """>"
		response.write vbcrlf & "<option value=""N"">Select a gender...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("gender") & """"
			If sGenderMatch = oRs("gender") Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("genderdescription") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		' this should never happen
		response.write vbcrlf & "<input type=""hidden"" id=""" & sElement & """ name=""" & sElement & """ value=""N"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' boolean OrgFeatureHasAlternateAccount( sOrgId, sFeature )
'------------------------------------------------------------------------------
Function OrgFeatureHasAlternateAccount( ByVal sOrgId, ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT V.orgid, V.feature FROM egov_verisign_options V, Organizations O "
	sSql = sSql & "WHERE V.orgid = O.orgid AND O.hasAlternatePaymentAccounts = 1 AND "
	sSql = sSql & "V.orgid = " & sOrgId & " AND V.feature = '" & sFeature & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		OrgFeatureHasAlternateAccount = True 
	Else
		OrgFeatureHasAlternateAccount = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------
' AdjustCitizenAccountBalance iUserID, sEntryType, sAmount 
'------------------------------------------------------------------------------
Sub AdjustCitizenAccountBalance( ByVal iUserID, ByVal sEntryType, ByVal sAmount )
	Dim sNewBalance, cPriorBalance, sSql

	cPriorBalance = GetCitizenAccountAmount( iUserId )

	If sEntryType = "credit" Then
		sNewBalance = CDbl(cPriorBalance) + CDbl(sAmount)
	Else  ' debit
		sNewBalance = CDbl(cPriorBalance) - CDbl(sAmount)
	End If 

	sSql = "UPDATE egov_users SET accountbalance = " & sNewBalance & " WHERE userid = " & iUserID

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' double GetCitizenAccountAmount( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetCitizenAccountAmount( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountbalance,0.000) AS accountbalance FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCitizenAccountAmount = CDbl( oRs("accountbalance") )
	Else
		GetCitizenAccountAmount = CDbl( 0.0000 )
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFeaturePublicURL( iOrgId, sFeature )
'--------------------------------------------------------------------------------------------------
Function GetFeaturePublicURL( ByVal iOrgId, ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(O.publicurl,F.publicurl) AS publicurl "
	sSql = sSql & "FROM egov_organizations_to_features O, egov_organization_features F "
	sSql = sSql & "WHERE O.featureid = F.featureid AND F.feature = '" & sFeature & "' AND O.orgid = " & iOrgId
	'response.write sSql & "<br >"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFeaturePublicURL =  oRs("publicurl")
	Else
		GetFeaturePublicURL = ""
	End If 

End Function 


'------------------------------------------------------------------------------
function getGoogleSearchID(iOrgID, iDBColumn)
	 dim sSQL, sDBColumn, sOrgID, lcl_return

  sOrgID     = 0
  sDBColumn  = ""
  lcl_return = ""

  if not containsApostrophe(iOrgID) then
     sOrgID = clng(iOrgID)
  end if

  if not containsApostrophe(iDBColumn) then
     sDBColumn = iDBColumn
  end if

  sSQL = "SELECT " & sDBColumn & " as googleSearchID "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid = " & sOrgID

 	set oRs = Server.CreateObject("ADODB.Recordset")
 	oRs.Open sSQL, Application("DSN"), 3, 1

 	if not oRs.eof then
	   	lcl_return = oRs("googleSearchID")
  end if

 	oRs.Close
  set oRs = nothing

  getGoogleSearchID = lcl_return

end function


'------------------------------------------------------------------------------
Function getOrganization_WP_URL( byVal iOrgId, byVal sColumnName )
  ' This pulls the Word Press URL for the passed org and column
  Dim sSql, oRs, sURL
  
  sURL = ""
  
  sSql = "SELECT " & sColumnName & " AS wp_url,OrgPublicWebsiteURL FROM organizations WHERE wpLive = 1 and orgid = " & CLng(iOrgId)
  'response.write sSQL & "<br /><br />"
  
  Set oRs = Server.CreateObject("ADODB.Recordset")
  oRs.Open sSql, Application("DSN"), 0, 1
  
  'response.write oRs("wp_url") & "<br /><br />"

  If Not oRs.EOF Then
    sURL = oRs("wp_url")
    if sURL = "" or isnull(sURL) then
    	if sColumnName = "wp_actionline_url" then
	    sURL = oRs("OrgPublicWebsiteURL") & "/citizen-action-line/"
    	elseif sColumnName = "wp_subscriptions_url" then
	    sURL = oRs("OrgPublicWebsiteURL") & "/subscriptions/"
    	end if
    end if
  End If
  
  oRs.Close
  Set oRs = Nothing  
  
  getOrganization_WP_URL = sURL
   
End Function

'------------------------------------------------------------------------------
function displayPrivacyPolicyLink(iIsDefaultPage, iOrgID)

  dim oRs, sSQL, lcl_return, sOrgID, lcl_current_value, lcl_link_class, sIsDefaultPage, lcl_egov_url
  dim lcl_url_start, lcl_url_end, lcl_url_length, lcl_text_start, lcl_text_end, lcl_text_length

  sOrgID            = 0
  sIsDefaultPage    = false
  lcl_return        = ""
  lcl_current_value = ""
  lcl_website_url   = ""
  lcl_website_text  = ""
  lcl_display_text  = ""
  lcl_display_url   = ""
  lcl_url_value     = ""
  lcl_link_class    = "afooter"
  lcl_egov_url      = ""
  lcl_url_start     = 0
  lcl_url_end       = 0
  lcl_url_length    = 0
  lcl_text_start    = 0
  lcl_text_end      = 0
  lcl_text_length   = 0

  if not containsApostrophe(iOrgID) then
     sOrgID = clng(iOrgID)
  end if

  if iIsDefaultPage then
     sIsDefaultPage = iIsDefaultPage
  end if

  if sIsDefaultPage then
     lcl_link_class = "adefaultfooter"
  end if

  sSQL = "SELECT privacypolicy_egov "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid = " & sOrgID

 	set oRs = Server.CreateObject("ADODB.Recordset")
 	oRs.Open sSQL, Application("DSN"), 3, 1

 	if not oRs.eof then
	   	lcl_current_value = oRs("privacypolicy_egov")

     if trim(lcl_current_value) <> "" then
        lcl_url_start    = instr(lcl_current_value,"[")
        lcl_url_end      = instr(lcl_current_value,"]")
        lcl_url_length   = lcl_url_end - lcl_url_start

        lcl_text_start   = instr(lcl_current_value,"<")
        lcl_text_end     = instr(lcl_current_value,">")
        lcl_text_length  = lcl_text_end - lcl_text_start

        if lcl_url_start > -1 AND lcl_url_length > 0 then
           lcl_website_url  = mid(lcl_current_value,lcl_url_start,lcl_url_length)
           lcl_website_url  = replace(lcl_website_url,"[","")
           lcl_website_url  = replace(lcl_website_url,"]","")
        end if

        if lcl_text_start > -1 AND lcl_text_length > 0 then
           lcl_website_text = mid(lcl_current_value,lcl_text_start,lcl_text_length)
           lcl_website_text = replace(lcl_website_text,"<","")
           lcl_website_text = replace(lcl_website_text,">","")
        end if

        if lcl_website_text <> "" then
           lcl_display_text = lcl_website_text
        else
           lcl_display_text = lcl_website_url
        end if

        lcl_url_value = replace(lcl_url_value,lcl_current_value,"")

       'Build the "display_url"
        if lcl_website_url <> "" then
           if instr(lcl_website_url,"http://") = 0 AND instr(lcl_website_url,"https://") = 0 then
              'lcl_display_url = lcl_display_url & "http://"
              lcl_egov_url = getOrgEgovWebsiteURL(iOrgID) & "/"
           end if

           lcl_display_url = " | "
           lcl_display_url = lcl_display_url & "<a href="""
           lcl_display_url = lcl_display_url & lcl_egov_url
           lcl_display_url = lcl_display_url & lcl_website_url
           lcl_display_url = lcl_display_url & """ class=""" & lcl_link_class & """>"
           lcl_display_url = lcl_display_url & lcl_display_text
           lcl_display_url = lcl_display_url & "</a>"

           lcl_return = lcl_display_url
        end if
     end if
  end if

 	oRs.Close
  set oRs = nothing

  displayPrivacyPolicyLink = lcl_return

end function

'------------------------------------------------------------------------------
function getOrgEgovWebsiteURL(iOrgID)

		Dim sSql, oRs, lcl_return

  lcl_return = ""

		sSql = "SELECT ISNULL(OrgEgovWebsiteURL,'') AS OrgEgovWebsiteURL FROM organizations WHERE orgid = " & iOrgID

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			lcl_return = oRs("OrgEgovWebsiteURL")
		End If
			
		oRs.Close
		Set oRs = Nothing

  getOrgEgovWebsiteURL = lcl_return

	End Function 
'------------------------------------------------------------------------------
' string sRefundName = GetRefundName()
'------------------------------------------------------------------------------
Function GetRefundName( )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypename FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE T.isrefundmethod = 1 AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & iOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundName = oRs("paymenttypename")
	Else
		GetRefundName = "Refund Voucher"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 
'------------------------------------------------------------------------------
' integer iUserId = GetFamilyId( iUserId )
'------------------------------------------------------------------------------
Function GetFamilyId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT familyid FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyId = oRs("familyid")
	Else
		GetFamilyId = iUserId
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 
'------------------------------------------------------------------------------
'  integer iPaymentTypeId = GetRefundPaymentTypeId()
'------------------------------------------------------------------------------
Function GetRefundPaymentTypeId( )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypeid FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE T.isrefundmethod = 1 AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & iOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundPaymentTypeId = oRs("paymenttypeid")
	Else
		GetRefundPaymentTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 
'------------------------------------------------------------------------------
' integer GetPaymentAccountId( iOrgId, iPaymentTypeId )
'------------------------------------------------------------------------------
Function GetPaymentAccountId( ByVal iOrgId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_organizations_to_paymenttypes "
	sSql = sSql & "WHERE orgid = " & iOrgId & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPaymentAccountId = CLng(oRs("accountid"))
	Else
		GetPaymentAccountId = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 
'------------------------------------------------------------------------------
'  Function GetRefundDebitAccountId( )
'------------------------------------------------------------------------------
Function GetRefundDebitAccountId( )
	Dim sSql, oRefund

	sSql = "select isnull(O.accountid,0) as accountid from egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " where T.isrefunddebit = 1 and T.paymenttypeid = O.paymenttypeid and O.orgid = " & iOrgID

	Set oRefund = Server.CreateObject("ADODB.Recordset")
	oRefund.Open sSql, Application("DSN"), 0, 1

	If Not oRefund.EOF Then
		GetRefundDebitAccountId = oRefund("accountid")
	Else
		GetRefundDebitAccountId = 0
	End If 

	oRefund.Close
	Set oRefund = Nothing 
End Function 
%>
