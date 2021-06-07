<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

	<script language="Javascript">
	  <!--

		function UpdateFamily( sUserId )
		{
			location.href='../dirs/family_members.asp?userid=' + sUserId;
		}
	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">

			<% ShowHeader sLevel %>
			<!--#Include file="../menu/menu.asp"--> 

      </td>
    </tr>

  </table>

<!-- #include file="dir_constants.asp"-->

<div id="content">
	<div id="centercontent">
<!------------------------------------- -->

<%
Dim conn, cmd, ResultID,strSuccess

%>

<table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center">&nbsp;<!--<img src="../images/icon_directory.jpg">--></td>
      <td><font size="+1"><b><%=langCommittees%></b></font><br>
	  <div id="goback" name="goback">
	  <img src='../images/arrow_back.gif' align='absmiddle'><a href='javascript:history.go(-1)'><%=langGoBack%></a>
	  
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink_citizen.asp'-->      
      </td>
      <td colspan="2" valign="top">
        
			<%
			if trim(request.querystring("userid"))="" then
				response.end
			end if

			set conn = Server.CreateObject("ADODB.Connection")
			set cmd=Server.CreateObject("ADODB.Command")
			conn.Open Application("DSN")
			Set cmd.ActiveConnection=conn
			cmd.commandtext="DeleteCitizen"
			cmd.commandtype=&H0004
			cmd.Parameters.Refresh
			cmd.parameters(1) = CLng(request.querystring("userid"))
			cmd.execute
			ResultID=cmd.parameters(0)
			conn.close
			set conn=nothing
			set cmd=nothing

			'response.write "<br>ResultID="&ResultID
			select case ResultID
			case -1
			case 0
			case 1
			response.write "<br>"&langSucessfulDeleted&"<br>"
			strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_citizen.asp'>"&langBackToUserDisplay&"</a>"
			response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"
			end select
			%>
		</td>
	</tr>
</table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->

