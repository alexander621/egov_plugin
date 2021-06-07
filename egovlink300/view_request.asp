<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 dim sError

'If users supplied comments then update them
 if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	   sCitizenMsg = request("sMsg")
   	iFormID     = request("iFormID")
   	iUserID     = request("iUSerID")
   	sStatus     = request("sStatus")
   	iOrgID      = iorgid 
   	AddCommentTaskComment sStatus,sCitizenMsg,iFormID,iUserID,iOrgID 
 end if
%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	<script src="scripts/modules.js"></script>
 <script src="../scripts/layers.js"></script>

<script language="javascript">
<!--
		function openWin2(url, name) {
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
 	}
//-->
</script>
</head>

<!--#Include file="include_top.asp"-->

<% RegisteredUserDisplay("") %>

<div id="content">
  <div id="centercontent">
<%
 'Get information for this request
  iTrackID = request("request_id")
  'iTime    = right(iTrackID,4)
  'iID      = replace(iTrackID,iTime,"")

  if iTrackID = "" then
    	iTrackID = "000000"
  end if

  sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid = " & CLng(iTrackID)
  set oRequest = Server.CreateObject("ADODB.Recordset")
  oRequest.Open sSQL, Application("DSN"), 3, 1

  if not oRequest.eof then
    	blnFound      = True
    	sTitle        = oRequest("category_title")
   		sStatus       = oRequest("status")
   		datSubmitDate = oRequest("submit_date")
   		sComment      = oRequest("comment")
   		iFormID       = oRequest("action_autoid")
   		iUserID       = oRequest("userid")

				 if sTitle = "" then
   					'response.write "</strong><br /><font color=red>!No action line category name provided!</font>"
        response.write "<br /><font color=""#ff0000"">!No action line category name provided!</font>" & vbcrlf
				 end if
%>
	<div style="margin-left:20px;" class="box_header4">Action Line Item: <%=sTitle%></div>
  	<div class="group" style="margin-left:20px;">
<p>
    <strong>Request Form Responses:</strong>
 			<%
	  			if sComment <> "" then
    					response.write "<br />" & sComment & vbcrlf
				  else
    					response.write "<br/><i><font color=""red"">!No comment/description provided!</i></font>" & vbcrlf
      end if
 			%>
</p>
			
<!--BEGIN: ONLINE DIALOG RESPONSE-->
<form name="frmPost" action="#" method="POST">
  <input type="hidden" name="iFormID" value="<%=iFormID%>" />
		<input type="hidden" name="iUserID" value="<%=iUserID%>" />
		<input type="hidden" name="sStatus" value="<%=sStatus%>" />
		<input type="hidden" name="REQUEST_ID" value="<%=iTrackID%>" />
<div id="post_form" style="padding:5px;margin-top:5px;border:solid 1px #000000;background-color:#E0E0E0;">
<table>
  <tr>
      <td>
          <strong>Post a response/question:</strong><br />
          <textarea onMouseOut="this.style.backgroundColor='#ffffff';" onMouseOver="this.style.backgroundColor='#FFFFCC';" name="sMsg" rows="5" cols="80"></textarea>
      </td>
  </tr>
		<tr>
      <td>
          <input type="submit" value="POST MESSAGE" />
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" value="CLEAR MESSAGE" />
      </td>
  </tr>
</table>
</form>
</div>
<!--END: ONLINE DIALOG RESPONSE-->

<p>
<strong>Action Request Activity:</strong>
<div style="margin-top:5px;border-top:solid 1px #000000;border-bottom:solid 1px #000000;background-color:#FFFFFF">
<table>
  <tr><td><strong>Status - Date of Activity</strong></td></tr>
</table>
</div>
<%
 'List History
		List_Comments(iTrackID)

  response.write "</p>" & vbcrlf

  if not OrgHasFeature( iOrgId, "hide email actionline" ) then
     response.write "<p>" & vbcrlf
     response.write "<strong>Email Contact:</strong>" & vbcrlf
     response.write "<div style=""padding: 5px; margin-top:5px;border:solid 1px #000000;background-color:#FFFFFF"">" & vbcrlf

 				sSQLa = "SELECT assigned_email FROM egov_action_request_view where action_autoid = " & iTrackID
	 			set oAssigned = Server.CreateObject("ADODB.Recordset")
		 		oAssigned.Open sSQLa, Application("DSN"), 3, 1

     response.write "<strong>" & oAssigned("assigned_email") & "</strong> has been assigned to this request." & vbcrlf
     response.write "Please contact via email - <a href=""mailto:" & oAssigned("assigned_email") & """>" & oAssigned("assigned_email") & "</a> - for further information regarding this request." & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf

 				oAssigned.close
	 			set oAssigned = nothing 
  end if

  response.write "</div>" & vbcrlf
  response.write "</div>" & vbcrlf

else
	
	'request not found
 	blnFound = False
 	response.write "<div style=""margin-left:20px;"" class=""box_header2"">Action Line Request Lookup</div>" & vbcrlf
  response.write "<div class=""groupsmall"" style=""margin-left:20px;"">" & vbcrlf
  response.write "<p>We could not locate an action line request using <strong>TRACKING NUMBER (" & iTrackID & ")</strong>.</p>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "   Please press <strong>BACK</strong> on your browser, check the <strong>TRACKING NUMBER</strong>, and try again. " & vbcrlf
  'response.write "   If you continue to receive this message please contact <a href=""mailto:jstullenberger@eclink.com"">jstullenberger@eclink.com</a> for further assistance with this request." & vbcrlf
  response.write "   If you continue to receive this message please contact <a href=""mailto:egovsupport@eclink.com"">egovsupport@eclink.com</a> for further assistance with this request." & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<p>Thank you for using " & sorgname & " E-Gov website.</p>" & vbcrlf
  response.write "</div>" & vbcrlf
end if

set oRequest = nothing 
%>

  </div>
</div>

<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
function List_Comments(iTrackID)

	 sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_action_responses "
  sSQL = sSQL &   " LEFT OUTER JOIN egov_users ON egov_action_responses.action_userid = egov_users.userid "
  sSQL = sSQL & " WHERE action_autoid = " & iTrackID
  sSQL = sSQL & " ORDER BY action_editdate DESC"

	 set oCommentList = Server.CreateObject("ADODB.Recordset")
	 oCommentList.Open sSQL, Application("DSN") , 3, 1

  sBGColor = "#ffffff"

 	if not oCommentList.eof then
   		while not oCommentList.eof
        sBGColor = changeBGColor(sBGColor,"#ffffff","#e0e0e0")

     			response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
        response.write "<table>" & vbcrlf
        response.write "  <tr><td>" & UCASE(oCommentList("action_status")) & " - " &  oCommentList("action_editdate") & "</td></tr>" & vbcrlf
			
        if oCommentList("action_externalcomment") <> "" then
       				response.write "  <tr><td>&nbsp;&nbsp;&nbsp;<strong>" & sOrgName & ": </strong><i>" & oCommentList("action_externalcomment")  & "</i></td></tr>" & vbcrlf
        elseif oCommentList("action_citizen") <> "" then
           response.write "  <tr><td>&nbsp;&nbsp;&nbsp;<strong>" & oCommentList("userfname")  & " " & oCommentList("userlname") & " : </strong><i>" & oCommentList("action_citizen")  & "</i></td></tr>" & vbcrlf
        else
       				response.write "  <tr><td>&nbsp;&nbsp;&nbsp;<strong>" & sOrgName & ": </strong><i>Your request was reviewed and/or status was updated.</i></td></tr>" & vbcrlf
        end if

        response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf

        oCommentList.movenext
     wend
  else
     response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
     response.write "<table>" & vbcrlf
     response.write "  <tr><td><font color=""#ff0000""><i>No activity</i></td></tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

end function

'------------------------------------------------------------------------------
function CheckSelected(sValue,sValue2)
	 sReturnValue = ""

	 if sValue = sValue2 then
	   	sReturnValue = "SELECTED"
	 end if

 	CheckSelected = sReturnValue

end function

'------------------------------------------------------------------------------
function AddCommentTaskComment(sStatus,sCitizenMsg,iFormID,iUserID,iOrgID)

		sSQL = "INSERT egov_action_responses ("
  sSQL = sSQL & " action_status,"
  sSQL = sSQL & " action_citizen,"
  sSQL = sSQL & " action_userid,"
  sSQL = sSQL & " action_orgid,"
  sSQL = sSQL & " action_autoid"
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL & "'" & sStatus             & "', "
  sSQL = sSQL & "'" & DBsafe(sCitizenMsg) & "', "
  sSQL = sSQL & "'" & iUserID             & "',"
  sSQL = sSQL & "'" & iOrgID              & "',"
  sSQL = sSQL & "'" & iFormID             & "'"
  sSQL = sSQL & ")"

		set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN") , 3, 1
		set oComment = nothing

end function

'------------------------------------------------------------------------------
function DBsafe( strDB )
	 dim sNewString

  if not VarType( strDB ) = vbString then
     DBsafe = strDB : Exit Function
  end if

 	sNewString = replace( strDB, "'", "''" )
	 sNewString = replace( sNewString, "<", "&lt;" )

 	DBsafe = sNewString

end function
%>