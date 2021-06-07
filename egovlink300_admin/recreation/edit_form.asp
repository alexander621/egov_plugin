<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: edit_form.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "rec payment alerts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' IF UPDATE PROCESS ITEMS
If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN
	assignUserID = request.form("assignUserID")
	dash = instr(assignUserID,"-")
	UserID = left(assignUserID,dash-1)
	EmailLen = len(assignUserID)
	EmailLen = EmailLen - dash
	Email = right(assignUserID,EmailLen)
	
	if Email = "" then 
		errorEmailMsg = "<p><b><font color=#cc0000>No email on file for user " & UserID & ". Update not completed.</font></b>"
		blnUpdate = True

	else 
		' UPDATE DATABASE
		sSQL = "SELECT * FROM egov_paymentservices where paymentserviceid=" & request("FORM_ID")
		Set oUpdate = Server.CreateObject("ADODB.Recordset")
		oUpdate.CursorLocation = 3
		oUpdate.Open sSQL, Application("DSN") , 1, 3
		oUpdate("assigned_email") = Email
		oUpdate("assigned_userID") = UserID
		oUpdate.Update
		Set oUpdate = Nothing
		blnUpdate = True
	
		session("updatedAssignedUserID") = UserID	
	end if
End If
%>


<html>
<head>
  <title><%=langBSPayments%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="facility.css" />

  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs tabRecreation,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="../images/icon_home.jpg"></td>-->
      <td><font size="+1"><b>Set Recreation Payment Alerts</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;
	  
	  <% If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN %>
	     <a href="manage_recreation_alerts.asp"><%=langBackToStart%></a>
	  <% else %>
	     <a href="javascript:history.go(-1)"><%=langBackToStart%></a>
	  <% end if %>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <!--<td valign="top" nowrap>-->

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
'        sLinks = "<div style=""padding-bottom:8px;""><b>" & langEventLinks & "</b></div>"

'        If bCanEdit Then
'          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
'          bShown = True
'        End If
        
'        If bShown Then
'          Response.Write sLinks & "<br>"
'        End If
        %>

        <% 'Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      <!--</td>-->
        
      <td colspan="2" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
	    <%
		iID = request("control")
		If iID = "" Then
			iID = "0"
		End If

		' CHECK IF POSTING FROM UPDATE
	    	If blnUpdate Then 
			iID = request("FORM_ID")
			if errorEmailMsg <> "" then 
				response.write errorEmailMsg
			else
				response.write "<b><font color=blue> << Record updated >></font></b>"
			end if		
		End If
      	
		sSQL = "SELECT * FROM egov_paymentservices where paymentserviceid=" & iID
		
		Set oForm = Server.CreateObject("ADODB.Recordset")
		oForm.Open sSQL, Application("DSN") , 3, 1

		' CHECK FOR INFORMATION
		If NOT oForm.EOF Then
			sFormName = oForm("paymentservicename")
			sUserID = oForm("assigned_UserID")

			If sUserID = "" or IsNull(sUserID) Then
				sUserID = session("updatedAssignedUserID")
			End If
		Else	
			response.write "FORM NOT FOUND"
		End If

		Set oForm = Nothing 
		%>

<!--BEGIN EDIT FORM -->
<form name=frmUpdate action="edit_form.asp" method="post">
<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="manage_action_forms.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdate.submit();">Update</a></div>

	  <input name=form_id type=hidden value="<%=iID%>" />
	  <div class="shadow">
	  <table class="tablelist" cellpadding=5 cellspacing=0>
		<tr><th align=left colspan=2>Form: <%=sFormName%></th></tr>
		<tr><td>Notification</td><td><select name=assignUserID>
		
		<%
		eSQL = "SELECT userID,email,FirstName,LastName FROM Users where orgId=" & session("orgId")
		Set oUsers = Server.CreateObject("ADODB.Recordset")
		oUsers.Open eSQL, Application("DSN"), 1, 3
		
		do while not oUsers.EOF
			if sUserID = oUsers("userID") then selected = " selected" else selected = ""
			response.write "<option value=" & oUsers("userID") & "-" & oUsers("email") & selected & ">" & oUsers("FirstName") & " " & oUsers("LastName") & "</option>"
		oUsers.MoveNext
		Loop
		
		%>
		
		</select></td></tr>
	</table>
	</div>

<div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="manage_recreation_alerts.asp">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmUpdate.submit();">Update</a></div>
</form>
<!--END EDIT FORM -->

      </td>
       
    </tr>
  </table>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


