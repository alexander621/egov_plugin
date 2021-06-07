<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: expirationdateedit.asp
' AUTHOR: Steve Loar
' CREATED: 05/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits a manually entered fee for a permit
'
' MODIFICATION HISTORY
' 1.0   05/20/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, sExpirationDate

iPermitId = CLng(request("permitid"))

sSql = "SELECT expirationdate FROM egov_permits WHERE permitid = " & iPermitId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	sExpirationDate = FormatDateTime(oRs("expirationdate"),2)
End If 

oRs.Close
Set oRs = Nothing 

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="Javascript">
		<!--

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmDate", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doUpdate()
		{
			// Check the expirationdate
			if (document.frmDate.expirationdate.value == "")
			{
				alert("Please enter an expiration date");
				document.frmDate.expirationdate.focus();
				return;
			}
			else
			{
				if (! isValidDate(document.frmDate.expirationdate.value))
				{
					alert("The expiration date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmDate.expirationdate.focus();
					return;
				}
			}

			// Fire off AJAX call to update the date
			doAjax('expirationdateupdate.asp', 'permitid=<%=iPermitId%>&expirationdate=' + document.frmDate.expirationdate.value, '', 'get', '0');

			// Update the parent window
			window.opener.document.getElementById("expirationdate").innerHTML = document.frmDate.expirationdate.value;

			// Close yourself
			doClose();
		}
		
		function doClose()
		{
			window.close();
			window.opener.focus();
		}

		function init()
		{
			document.getElementById("expirationdate").focus();
		}

		window.onload = init; 

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<font size="+1"><strong>Edit Expiration Date</strong></font><br /><br />
				<form name="frmDate" action="expirationdateedit.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p> 
						Expiration Date: &nbsp; <input type="input" name="expirationdate" id="expirationdate" value="<%=sExpirationDate%>" size="10" maxlength="10" />
						&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('expirationdate');" /></span>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" value="Save Changes" onclick="doUpdate();" />
						 &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
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
