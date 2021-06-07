<!--BEGIN: BOTTOM MENU AND COPYRIGHT INFORMATION-->

<% Set oFooterOrg = New classOrganization %>

<div class="footerbox">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr><td height="5" bgcolor="#93bee1" style="border-bottom: solid 1px #000000;">&nbsp; </td></tr>
		<tr>
			<td valign="top" align="center">
				<font style="font-size:10px;font-weight:bold;">Copyright &copy;2004-<%=Year(Date())%>. All Rights Reserved. <% =oFooterOrg.GetOrgDisplayName( "admin footer brand link" )%>&nbsp;
<%
			Dim iLoadTime
			iLoadTime = CDbl(0.00)
			If iPageLogStartSecs <> "" Then 
				iLoadTime = timer - iPageLogStartSecs
				If UserIsRootAdmin( session("UserID") ) Then 
					response.write FormatNumber(iLoadTime,3) & " seconds"
				End If 
			End If 
%>
				</font>
			</td>
		</tr>
	</table>
</div>

 <script>
 function onElementHeightChange(elm, callback){
    var lastHeight = elm.clientHeight, newHeight;
    (function run(){
        newHeight = elm.clientHeight;
        if( lastHeight != newHeight )
            callback();
        lastHeight = newHeight;

        if( elm.onElementHeightChangeTimer )
            clearTimeout(elm.onElementHeightChangeTimer);

        elm.onElementHeightChangeTimer = setTimeout(run, 200);
    })();
}


	if (window.top!=window.self)
	{
		onElementHeightChange(document.body, function(){
			//alert("HERE");
			var height = document.body.scrollHeight;
 			parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: false }},"*")
		});
		var height = document.body.scrollHeight;
		parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: true }},"*")

	}
 </script>

 <% if session("orgid") = "211" then%>
 	<script>
	function getCookie(cname) {
  		var name = cname + "=";
  		var decodedCookie = decodeURIComponent(document.cookie);
  		var ca = decodedCookie.split(';');
  		for(var i = 0; i <ca.length; i++) {
    			var c = ca[i];
    			while (c.charAt(0) == ' ') {
      				c = c.substring(1);
    			}
    			if (c.indexOf(name) == 0) {
      				return c.substring(name.length, c.length);
    			}
  		}
  		return "";
	}
	function nagUser()
	{
		var nagResposne = prompt("Your account is overdue.  Please call 513-591-7379 to make payment arrangements or to have this warning removed.  Please enter 'overdue' below to acknowledge.");
		if (nagResposne != "overdue")
		{
			nagUser();
		}
		else
		{
			var now = new Date();
               		now.setTime(now.getTime() + 1 * 3600 * 1000);
			document.cookie = "nag=done; expires=" + now.toUTCString() + "; path=/";
		}
	}
	nagState = getCookie("nag");
	if (nagState != "done")
	{
		nagUser();
	}
	</script>
  <%end if %>

<% Set oFooterOrg = Nothing 

LogThePage iLoadTime

%>
<!--END: BOTTOM MENU AND COPYRIGHT INFORMATION-->

<%
'------------------------------------------------------------------------------------------------------------
' FUNCTION AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void LogThePage iLoadTime 
'------------------------------------------------------------------------------------------------------------
Sub LogThePage( ByVal iLoadTime )
	Dim sSql, oCmd, sScriptName, sVirtualDirectory, aVirtualDirectory, sPage, arr, sUserAgent, sUserAgentGroup

	sScriptName = Request.ServerVariables("SCRIPT_NAME")

	If request.servervariables("http_user_agent") <> "" Then 
		sUserAgent = "'" & Track_DBsafe(Trim(Left(request.servervariables("http_user_agent"),480))) & "'"
	Else
		sUserAgent = "NULL"
	End If 

	If Len(Trim(request.servervariables("http_user_agent"))) > 0 Then 
		sUserAgentGroup = "'" & GetUserAgentGroup( LCase(request.servervariables("http_user_agent")) ) & "'"
	Else
		sUserAgentGroup = "'" & GetUntrackedUserAgentGroup( ) & "'"
	End If 

	' Get the virtual directory
	aVirtualDirectory = Split(sScriptName, "/", -1, 1) 
	sVirtualDirectory = "/" & aVirtualDirectory(1) 
	sVirtualDirectory = "'" & Replace(sVirtualDirectory,"/","") & "'"

	' Get the page
	For Each arr in aVirtualDirectory 
		sPage = arr 
	Next 

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, "
	sSql = sSql & " querystring, servername, remoteaddress, requestmethod, orgid, userid, username, useragent, useragentgroup,requestformcollection, cookiescollection, sessioncollection, sessionid ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'admin', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3,,,0) & ", "
	sSql = sSql & "'" & sScriptName & "', "

	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		sSql = sSql & "'" & DBsafe(Left(Request.ServerVariables("QUERY_STRING"),500)) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 
	' our server name
	sSql = sSql & "'" & Request.ServerVariables("SERVER_NAME") & "', "

	' remote address
	sSql = sSql & "'" & Request.ServerVariables("REMOTE_ADDR") & "', "

	' request method - GET or POST
	sSql = sSql & "'" & Request.ServerVariables("REQUEST_METHOD") & "', "

	' orgid
	If session("orgid") <> "" Then 
		sSql = sSql & session("orgid") & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Userid
	If session("userid") <> "" Then
		sSql = sSql & session("userid") & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Get username
	If session("fullname") <> "" Then
		sSql = sSql & "'" & dbsafe(session("fullname")) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup & ", "
         'requestformcollection, cookiescollection, sessioncollection
	sSql = sSql & "'" & GetRequestFormCollection() & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql
	session("PageLogSQL") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
	session("PageLogSQL") = ""

End Sub 

Function GetRequestFormCollection()
	sPostLog = ""
	on error resume next
	For each item in Request.Form
		sPostLog = sPostLog & item & ":  " &	 request.form(item) & vbcrlf
	Next
	on error goto 0

	GetRequestFormCollection = dbsafe(sPostLog)
End Function
Function GetCookiesCollection()
	Collection = ""
	on error resume next
	For Each Item in Request.Cookies
		Collection = Collection & Item & ":  " & request.cookies(Item) & vbcrlf
	Next
	on error goto 0
	GetCookiesCollectionCollection = dbsafe(Collection)
End Function
Function GetSessionCollection()
	sSessionLog = ""
	on error resume next
	For each session_name in Session.Contents
		sSessionLog = sSessionLog & session_name & ":  " & session(session_name) & vbcrlf
	Next
	on error goto 0

	GetSessionCollection = dbsafe(sSessionLog)
End Function


'------------------------------------------------------------------------------------------------------------
' string GetUserAgentGroup( sUserAgent )
'------------------------------------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		'If clng(InStr( sUserAgent, oRs("useragentgroup") )) > clng(0) Then
		If clng(InStr( 1, sUserAgent, LCase(oRs("useragentgroup")), 1 )) > clng(0) Then
			sUserAgentGroup = oRs("useragentgroup")
			Exit Do 
		End If 
		oRs.MoveNext
	Loop 
	
	oRs.Close
	Set oRs = Nothing 
	
	GetUserAgentGroup = sUserAgentGroup

End Function 


'------------------------------------------------------------------------------------------------------------
' string GetUntrackedUserAgentGroup()
'------------------------------------------------------------------------------------------------------------
Function GetUntrackedUserAgentGroup( )
	Dim sSql, oRs

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetUntrackedUserAgentGroup = oRs("useragentgroup")
	Else
		GetUntrackedUserAgentGroup = "untracked"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>

