<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitalertedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/18/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Sets the permit alert for a permit
'
' MODIFICATION HISTORY
' 1.0   08/18/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sAlertMsg, sAlertSetByUser, dAlertDate

iPermitId = CLng(request("permitid"))

GetPermitAlertDetails iPermitId, sAlertMsg, sAlertSetByUser, dAlertDate ' in permitcommonfunctions.asp

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="Javascript">
		<!--

		function setAlert()
		{
			if (document.frmPermit.alertmsg.value == "")
			{
				alert( "Please enter some text for the alert message.\nThen try setting the alert again." );
				document.frmPermit.alertmsg.focus();
				return;
			}
			var sParameter = 'permitid=' + encodeURIComponent(document.frmPermit.permitid.value);
			sParameter += '&type=' + encodeURIComponent(document.frmPermit.actiontype.value);
			sParameter += '&alertmsg=' + encodeURIComponent(document.frmPermit.alertmsg.value);
			//alert(sParameter);
			doAjax('permitalertupdate.asp', sParameter, '', 'post', '0');
			parent.document.getElementById("permitalert").innerHTML = document.frmPermit.alertmsg.value;
			doClose();
			//document.frmPermit.submit();
		}

		function clearAlert()
		{
			doAjax('permitalertupdate.asp', 'permitid=<%=iPermitId%>&type=clear', '', 'get', '0');
			parent.document.getElementById("permitalert").innerHTML = ' ';
			doClose();
		}

		function CloseThisSaved( sResult )
		{
			alert(sResult);
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		//-->
		</script>

	</head>
	<body onload="setMaxLength();">
		<div id="content">
			<div id="centercontent">
				<form name="frmPermit" action="permitalertupdate.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<input type="hidden" name="actiontype" value="set" />
					<table>
						<tr><td>
<%							If sAlertSetByUser <> "" Then %>
								<strong>Set By: <%=sAlertSetByUser%> on <%=dAlertDate%></strong><br />
<%							End If		%>
								<textarea id="alertmsg" name="alertmsg" rows="5" cols="80" maxlength="1000"><%=sAlertMsg%></textarea>
							</td>
						</tr>
					</table>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Set Alert" onclick="setAlert();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Clear Alert" onclick="clearAlert();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


%>
