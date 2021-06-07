<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: dl_mananage_subscribers.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects registerd users to be subscribers
'
' MODIFICATION HISTORY
' 1.?	 11/10/08	 Steve Loar - Changed to handle session timeout and memory leak from first recordset not being destroyed
' 2.0  12/08/08  David Boyer - Added "Export Subscribers"
' 2.1  01/26/09  David Boyer - Fixed bug raised when clicking the "Edit" button and trying to pass in the list name with an apostrophe.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iMaillistid, iMaillistname, sDistributionListName

iMaillistid   = CLng(request("idlid"))
'iMaillistname = request("iname")

sDistributionListName = GetDistributionListName( iMaillistid )

'Check for org permissions
 'lcl_orghasfeature_customreports                = orghasfeature("customreports")
 lcl_orghasfeature_customreports_subscriberlist = orghasfeature("customreports_subscriberlist")

'Check for user permissions
 'lcl_userhaspermission_customreports                = userhaspermission(session("userid"),"customreports")
 lcl_userhaspermission_customreports_subscriberlist = userhaspermission(session("userid"),"customreports_subscriberlist")
%>
<html>
<head>

	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="classes.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script src="../scripts/tooltip_new.js"></script>

</head>
<body  bgcolor="#c9def0">

	<table border="0" cellpadding="10" cellspacing="0" width="100%" bgcolor="#c9def0">
	  <tr>
			 <td colspan="2" valign="top">
	<%
	 'Display list of subscribed mailing lists
			response.write "<center>" & vbcrlf
			response.write "<strong>Distribution List: "& sDistributionListName &" </strong><br />" & vbcrlf
			'response.write "<a href='javascript:self.close();'><font size=""2"">Close Window</font></a></font>" & vbcrlf

   response.write "<p>" & vbcrlf
			response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
			response.write "  <tr valign=""top"">" & vbcrlf
   response.write "      <td>" & vbcrlf
                          			subDisplaySubscribedUsers
			response.write "      </td>" & vbcrlf

		'Display Arrows
			response.write "      <td align=""center"" valign=""middle"">" & vbcrlf
			'response.write "          &nbsp;<a href='javascript:document.sl.submit();'><img src=""../images/ieforward.gif"" align=""absmiddle"" border=""0""></a>" & vbcrlf
			'response.write "          <a href='javascript:document.al.submit();'><img src=""../images/ieback.gif"" align=""absmiddle"" border=""0""></a>" & vbcrlf
			response.write "          <p>" & vbcrlf
   response.write "          &nbsp;<img src=""../images/ieforward.gif"" align=""absmiddle"" border=""0"" onclick=""document.sl.submit();"" class=""hotspot"" onmouseover=""tooltip.show('Click to UNSUBSCRIBE User(s)');"" onmouseout=""tooltip.hide();"" />" & vbcrlf
			response.write "          <p/>" & vbcrlf
   response.write "          <img src=""../images/ieback.gif"" align=""absmiddle"" border=""0"" onclick=""document.al.submit();"" class=""hotspot"" onmouseover=""tooltip.show('Click to SUBSCRIBE User(s)');"" onmouseout=""tooltip.hide();"" />" & vbcrlf
			response.write "      </td>" & vbcrlf
			response.write "      <td>" & vbcrlf
                          			subDisplayAvailableUsers
			response.write "      </td>" & vbcrlf
			response.write "  </tr>" & vbcrlf
			response.write "</table>" & vbcrlf
   response.write "</p>" & vbcrlf

   response.write "<input type=""button"" name=""sCloseWin"" id=""sCloseWin"" value=""Close Window"" class=""button"" onclick=""self.close();"" />" & vbcrlf

			response.write "</center>" & vbcrlf
	%>
		  </td>
			</tr>
	</table>
</body>
</html>
<%
'------------------------------------------------------------------------------
function GetDistributionListName( iMaillistid )
 	Dim sSql, oList

 	sSQL = "SELECT ISNULL(distributionlistname,'') AS distributionlistname "
 	sSQL = sSQL & " FROM EGOV_CLASS_DISTRIBUTIONLIST "
 	sSQL = sSQL & " WHERE orgid = " & SESSION("orgid")
 	sSQL = sSQL & " AND distributionlistid = " & iMaillistid

 	set oList = Server.CreateObject("ADODB.Recordset")
 	oList.Open sSQL, Application("DSN"), 0, 1

 	if not oList.eof then
   		GetDistributionListName = oList("distributionlistname")
	 else
	 	  GetDistributionListName = ""
	 end if

	 oList.Close
	 set oList = nothing 

end function

'------------------------------------------------------------------------------
sub subDisplaySubscribedUsers()
 	Dim sSql, oList

 	response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
 	response.write "  <form name=""sl"" method=""post"" action=""dl_deletemember.asp"">" & vbcrlf
 	response.write "  <tr><td height=""20"" align=""center""><strong>SUBSCRIBED</strong></td></tr>" & vbcrlf
 	response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
 	response.write "          <input type=""hidden"" name=""maillistid"" value="""   & iMaillistid   & """ />" & vbcrlf
 	response.write "          <input type=""hidden"" name=""maillistname"" value=""" & iMaillistname & """ />" & vbcrlf
 	response.write "          <select size=""15"" border=""0"" style=""width:365px"" name=""subscribedlist"" multiple=""multiple"">" & vbcrlf

 	sSQL = "SELECT u.*, '" & replace(sDistributionListName,"'","''") & "' AS categoryname "
 	sSQL = sSQL & " FROM egov_users u "
 	sSQL = sSQL & " INNER JOIN egov_class_distributionlist_to_user ug ON u.userid = ug.userid "
 	sSQL = sSQL & " WHERE (ug.distributionlistid = '" & iMaillistid & "') "
  sSQL = sSQL & " AND u.isdeleted = 0 "
 	sSQL = sSQL & " ORDER BY u.userlname, userfname, useremail "

  session("CR_SUBSCRIBERLIST") = sSQL

 	set oList = Server.CreateObject("ADODB.Recordset")
 	oList.Open sSQL, Application("DSN"), 0, 1

 'Loop thru Available Lists
 	if not oList.eof then
     iSubscriberCount = 0

	   	while not oList.eof
        iSubscriberCount = iSubscriberCount + 1

		 	    response.write "  <option value=""" & oList("userid") & """>" & vbcrlf

    			 if  (trim(oList("userlname")) = "" OR isnull(oList("userlname"))) _
        AND (trim(oList("userfname")) = "" OR isnull(oList("userfname"))) then
        			response.write oList("useremail")
    	 		else
       				response.write oList("userlname") & ", " & oList("userfname")
    			 end if

        response.write "</option>" & vbcrlf

     			oList.movenext
   		wend
 	end if

 	oList.close
	 set oList = nothing 

 	response.write "          </select><br />" & vbcrlf
  response.write "          <div align=""center""><strong>Total Subscribers: [" & iSubscriberCount & "]</strong></div><br />"

 'If the user has the following permissions then show the "Export Subscribers" button.
  'if  lcl_orghasfeature_customreports _
  'AND lcl_userhaspermission_customreports _
  if lcl_orghasfeature_customreports_subscriberlist AND lcl_userhaspermission_customreports_subscriberlist then
     response.write "          <div align=""center""><input type=""button"" value=""Export Subscribers"" onclick=""location.href='../customreports/customreports.asp?cr=SUBSCRIBERLIST&export=Y'"" /></div>" & vbcrlf
  end if

 	response.write "      </td>" & vbcrlf
 	response.write "  </tr>" & vbcrlf
 	response.write "  </form>" & vbcrlf
 	response.write "</table>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub SubDisplayAvailableUsers()
	 Dim sSql, oList

 	response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
 	response.write "  <form name=""al"" method=""post"" action=""dl_addmember.asp"">" & vbcrlf
 	response.write "  <tr><td height=""20"" align=""center""><strong>AVAILABLE</strong></td></tr>" & vbcrlf
 	response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
 	response.write "          <input type=""hidden"" name=""maillistid"" value="""   & iMaillistid   & """ />" & vbcrlf
 	response.write "          <input type=""hidden"" name=""maillistname"" value=""" & iMaillistname & """ />" & vbcrlf
 	response.write "          <select size=""15"" border=""0"" style=""width:365px"" name=""availablelist"" multiple=""multiple"">" & vbcrlf

 	sSQL = "SELECT u.*, '" & replace(sDistributionListName,"'","''") & "' AS categoryname "
	 sSQL = sSQL & " FROM egov_users AS u "
	 sSQL = sSQL & " WHERE userregistered = 1 and useremail IS NOT NULL"
	 sSQL = sSQL & " AND (useremail is not NULL OR userlname is not NULL and userfname is not NULL) "
	 sSQL = sSQL & " AND (userid NOT IN (SELECT userid "
	 sSQL = sSQL &                     " FROM egov_class_distributionlist_to_user AS ug "
	 sSql = sSql &                     " WHERE (distributionlistid = '" & iMaillistid & "'))) "
	 sSQL = sSQL & " AND (orgid = '" & SESSION("ORGID") & "') AND u.isdeleted = 0 "
	 sSQL = sSQL & " ORDER BY userlname, userfname, useremail "

 	set oList = Server.CreateObject("ADODB.Recordset")
 	oList.Open sSQL, Application("DSN"), 0, 1

 'Loop thru Available Lists
 	if not oList.eof then
  		 while not oList.eof
		      response.write "  <option value=""" & oList("userid") & """>" & vbcrlf

    			 if  (trim(oList("userlname")) = "" OR isnull(oList("userlname"))) _
        AND (trim(oList("userfname")) = "" OR isnull(oList("userfname"))) then
				       response.write oList("useremail")
    			 else
				       response.write oList("userlname") & ", " & oList("userfname")
    			 end if

        response.write "</option>" & vbcrlf

    			 oList.movenext
     wend
  end if

 	oList.close
	 set oList = nothing

 	response.write "          </select>" & vbcrlf
 	response.write "      </td>" & vbcrlf
 	response.write "  </tr>" & vbcrlf
 	response.write "  </form>" & vbcrlf
 	response.write "</table>" & vbcrlf

end sub
%>


