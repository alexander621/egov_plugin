<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feemultiplieredit.asp
' AUTHOR: Steve Loar
' CREATED: 12/18/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/18/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iFeeMultiplierTypeid, sFeeMultiplier, sFeeMultiplierRate

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit permits" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If isFeatureOffline("fee multiplier rates") = "Y" Then 
    response.redirect "../admin/outage_feature_offline.asp"
End If 

iFeeMultiplierTypeid = CLng(request("feemultipliertypeid") )

If CLng(iFeeMultiplierTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetFeeMultiplier iFeeMultiplierTypeid
Else
	sTitle = "New"
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function Another()
		{
			location.href="feemultiplieredit.asp?feemultipliertypeid=0";
		}

		function Validate()
		{
			var rege;
			var Ok; 

			if (document.frmRates.feemultiplier.value == '')
			{
				alert("Please provide a name for the fee multiplier, then try saving again.");
				document.frmRates.feemultiplier.focus();
				return;
			}

			if (document.frmRates.feemultiplierrate.value == '')
			{
				alert("Please provide a rate for the fee multilpier, then try saving again.");
				document.frmRates.feemultiplierrate.focus();
				return;
			}
			else
			{
				// Validate the rate format
				rege = /^\d{0,3}\.{0,1}\d{0,4}$/;
				Ok = rege.test(document.frmRates.feemultiplierrate.value);
				if (! Ok)
				{
					alert("The rate should be a numberic value in the format '###.####'.\nPlease correct this and try saving again.");
					document.frmRates.feemultiplierrate.focus();
					return;
				}
			}

			// All is OK so submit
			document.frmRates.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this fee multiplier rate?"))
			{
				location.href="feemultiplierdelete.asp?feemultipliertypeid=<%=iFeeMultiplierTypeid%>";
			}
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function commonIFrameUpdateFunction()
		{
			UpdateParentMultipliers('feemultipliers','feemultiplierDD')
		}
		function UpdateParentMultipliers(poptype, classname)
		{

			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type='+poptype+'&value=<%=request("id")%>', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Update the unselected values
				var unSelVals = parent.document.getElementById('newfeemultipliertypeid');
				unSelVals.innerHTML = newDDVals;
				unSelVals.value = unSelVals.options[0].value;

				//Update the already selected values
				var selVals = parent.document.getElementById('feemultipliertypeid').options;
				for (var i = 0; i < selVals.length; i++) {
					var unSelOptions = unSelVals.options;
					for (var j = 0; j < unSelOptions.length; j++) {

						//Find Matching option in unselected values
						if (selVals[i].value == unSelOptions[j].value)
						{
							//Update Name in Selected Values
							selVals[i].innerHTML = unSelOptions[j].innerHTML;
	
							//Purge from unselected values
							unSelVals.removeChild(unSelOptions[j]);
						}
					}
				}

			}

		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>


	//-->
	</script>

</head>

<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Fee Multiplier Rate</strong></font><br /><br />
				<a href="feemultiplierlist.asp?id=<%=request("id")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

		<div id="functionlinks">
<%		If CLng(iFeeMultiplierTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Update" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>

		<form name="frmRates" action="feemultiplierupdate.asp" method="post">
		<input type="hidden" name="id" value="<%=request("id")%>" />
		<input type="hidden" name="feemultipliertypeid" value="<%=iFeeMultiplierTypeid%>" />
		<div class="shadow">
			<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
				<tr>
					<td>
						<table cellpadding="5" cellspacing="0" border="0">
							<tr>
								<td align="right">Name:</td><td><input type="text" name="feemultiplier" value="<%=sFeeMultiplier%>" size="50" maxlength="50" /></td>
							</tr>
							<tr>
								<td align="right">Rate:</td><td><input type="text" name="feemultiplierrate" value="<%=sFeeMultiplierRate%>" size="8" maxlength="8" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
		</form>
<%		If CLng(iFeeMultiplierTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Update" /> &nbsp; &nbsp;
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		<!--END: EDIT FORM-->

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetFeeMultiplier( iFeeMultiplierTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetFeeMultiplier( iFeeMultiplierTypeid )
	Dim sSql, oRs

	sSql = "SELECT feemultiplier, feemultiplierrate FROM egov_feemultipliertypes WHERE feemultipliertypeid = " & iFeeMultiplierTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sFeeMultiplier = oRs("feemultiplier")
		sFeeMultiplierRate = oRs("feemultiplierrate")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
