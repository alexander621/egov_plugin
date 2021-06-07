<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: edit_form.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the payments list
'
' MODIFICATION HISTORY
' 1.0   ???			???? - INITIAL VERSION
' 1.1	10/12/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, selected

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "payment notifications" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' IF UPDATE PROCESS ITEMS
If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN
	assignUserID = request("assignUserID")
	dash = InStr(assignUserID,"-")
	UserID = Left(assignUserID,dash-1)
	EmailLen = Len(assignUserID)
	EmailLen = EmailLen - dash
	Email = Right(assignUserID,EmailLen)
	
	If Email = "" Then 
		errorEmailMsg = "<p><b><font color=#cc0000>No email on file for user " & UserID & ". Update not completed.</font></b>"
		blnUpdate = True
	Else 
		' UPDATE DATABASE
'		sSql = "SELECT * FROM egov_paymentservices where paymentserviceid=" & request("FORM_ID")
'		Set oUpdate = Server.CreateObject("ADODB.Recordset")
'		oUpdate.CursorLocation = 3
'		oUpdate.Open sSql, Application("DSN") , 1, 3
'		oUpdate("assigned_email") = Email
'		oUpdate("assigned_userID") = UserID
'		oUpdate.Update
'		Set oUpdate = Nothing

		sSql = "UPDATE egov_paymentservices "
		sSql = sSql & "SET assigned_userID = " & UserID
		sSql = sSql & ", assigned_email = '" & Email & "' "
		sSql = sSql & "WHERE paymentserviceid = " & request("FORM_ID")

		RunSQLStatement sSql

		blnUpdate = True
	
		session("updatedAssignedUserID") = UserID	
	End If 
End If
%>


<html>
<head>
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

	<script language="javascript">
	<!--

		function UpdateForm()
		{
			document.frmUpdate.submit();
		}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs tabPayments,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="../images/icon_home.jpg"></td>-->
      <td><font size="+1"><b>Edit Online Payment Form Notifications</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;
	  
	  <% If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN %>
	     <a href="manage_action_forms.asp?useSessions=1"><%=langBackToStart%></a>
	  <% else %>
	     <a href="javascript:history.go(-1)"><%=langBackToStart%></a>
	  <% end if %>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
        
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
			If errorEmailMsg <> "" Then 
				response.write errorEmailMsg
			Else 
				response.write "<b><font color=""red""> << Record updated >></font></b>"
			End if		 
		End If
      	
		sSql = "SELECT paymentservicename, assigned_UserID FROM egov_paymentservices WHERE paymentserviceid = " & iID
		
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		' CHECK FOR INFORMATION
		If NOT oRs.EOF Then
			sFormName = oRs("paymentservicename")
			sUserID = oRs("assigned_UserID")

			If sUserID = "" or IsNull(sUserID) Then
				sUserID = session("updatedAssignedUserID")
			End If
		Else	
			response.write "FORM NOT FOUND"
		End If

		oRs.Close 
		Set oRs = Nothing 
		%>

<!--BEGIN EDIT FORM -->
<form name="frmUpdate" action="edit_form.asp" method="post">
	<div style="font-size:10px; padding-bottom:5px;">
		<img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="manage_action_forms.asp">Cancel</a>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="JavaScript:document.frmUpdate.submit();">Update</a>
	</div>

	  <input name="form_id" type="hidden" value="<%=iID%>">

	  <div class="shadow">
	  <table class="tablelist" cellpadding="5" cellspacing="0">
		<tr><th align="left" colspan="2">Form: <%=sFormName%></th></tr>
		<tr><td>Notification</td>
		<td>
			<select name="assignUserID">
			<%
			sSql = "SELECT userid, email, FirstName, LastName FROM users "
			sSql = sSql & "WHERE orgid = " & session("orgId")
			sSql = sSql & " ORDER BY lastname, firstname"

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 0, 1
			
			Do While Not oRs.EOF
				If sUserID = oRs("userid") Then
					selected = " selected=""selected"" " 
				Else
					selected = ""
				End If 
				response.write "<option value=""" & oRs("userID") & "-" & oRs("email") & """" & selected & ">" & oRs("FirstName") & " " & oRs("LastName") & "</option>"
				oRs.MoveNext
			Loop

			oRs.Close 
			Set oRs = Nothing 
			
			%>
			</select>

		</td></tr>
	</table>
	</div>

	<div style="font-size:10px; padding-top:5px;">
		<img src="../images/cancel.gif" align="absmiddle" />&nbsp;<a href="manage_action_forms.asp">Cancel</a>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<img src="../images/go.gif" align="absmiddle" />&nbsp;<a href="JavaScript:document.frmUpdate.submit();">Update</a>
	</div>
</form>
<!--END EDIT FORM -->

      </td>
       
    </tr>
  </table>
  
  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


