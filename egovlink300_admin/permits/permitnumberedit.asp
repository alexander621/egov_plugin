<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitnumberedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/25/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits a permit number. This allows them ot override the generated number so that 
'				numbers in their old system will match with this system.
'
' MODIFICATION HISTORY
' 1.0   08/25/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, iPermitYear, iPermitNumber, iCharacters

iPermitId = CLng(request("permitid"))

sSql = "SELECT permitnumberyear, permitnumber FROM egov_permits WHERE permitid = " & iPermitId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	iPermitYear = oRs("permitnumberyear")
	iPermitNumber = oRs("permitnumber")
Else
	iPermitYear = Year(Date)
	iPermitNumber = 1
End If 

oRs.Close
Set oRs = Nothing 

iCharacters = GetPermitNumberSequenceSize()

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/formatnumber.js"></script>
		<script language="JavaScript" src="../scripts/removespaces.js"></script>
		<script language="JavaScript" src="../scripts/removecommas.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--
		function init()
		{
			$("permitnumber").focus();
		}

		function doUpdate() 
		{
			// validate the permit year
			if ($("permitnumberyear").value != '')
			{
				// Remove any extra spaces
				$("permitnumberyear").value = removeSpaces($("permitnumberyear").value);
				//Remove commas that would cause problems in validation
				$("permitnumberyear").value = removeCommas($("permitnumberyear").value);

				rege = /^\d{4}$/;
				Ok = rege.test($("permitnumberyear").value);
				if ( ! Ok )
				{
					alert("The year must be a number composed of 4 digits only.\nPlease correct this and try saving again.");
					$("permitnumberyear").focus();
					return;
				}
			}
			else
			{
				alert("The year is required.\nPlease correct this and try saving again.");
				$("permitnumberyear").focus();
				return;
			}

			// validate the permit number
			if ($("permitnumber").value != '')
			{
				// Remove any extra spaces
				$("permitnumber").value = removeSpaces($("permitnumber").value);
				//Remove commas that would cause problems in validation
				$("permitnumber").value = removeCommas($("permitnumber").value);

				rege = /^\d*$/;
				Ok = rege.test($("permitnumber").value);
				if ( ! Ok )
				{
					alert("The permit number must be numeric only.\nPlease correct this and try saving again.");
					$("permitnumber").focus();
					return;
				}
			}
			else
			{
				alert("The permit number is required.\nPlease correct this and try saving again.");
				$("permitnumber").focus();
				return;
			}
			//alert("OK");
			//document.frmPermitNumber.submit();

			//Do Ajax save call
			doAjax('permitnumberupdate.asp', 'permitid=' + $("permitid").value + '&permitnumberyear=' + $("permitnumberyear").value + '&permitnumber=' + $("permitnumber").value, 'UpdatePermitNumberDisplay', 'get', '0');
		}

		function UpdatePermitNumberDisplay( sReturn )
		{
			if ( sReturn != "Failed" )
			{
				parent.document.getElementById("permitnumberdisplay").innerHTML = sReturn;
			}
			doClose();
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		window.onload = init; 
		
		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmPermitNumber" action="permitnumberupdate.asp" method="post">
					<input type="hidden" id="permitid" name="permitid" value="<%=iPermitId%>" />
					<p> 
						<table cellpadding="5" cellspacing="0" border="0" id="permitnumbertable">
							<tr><td align="right" class="rowlabel">Year:</td><td><input type="input" name="permitnumberyear" id="permitnumberyear" value="<%=iPermitYear%>" size="4" maxlength="4" /></td></tr>
							<tr><td align="right" class="rowlabel" nowrap="nowrap">Permit Number:</td><td><input type="input" name="permitnumber" id="permitnumber" value="<%=iPermitNumber%>" size="<%=iCharacters%>" maxlength="<%=iCharacters%>" /></td></tr>
						</table>
						<br />
							<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" value="Save Changes" onclick="doUpdate();" />
							 &nbsp; &nbsp; 
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

'--------------------------------------------------------------------------------------------------
' Function GetPermitNumberSequenceSize()
'--------------------------------------------------------------------------------------------------
Function GetPermitNumberSequenceSize()
	Dim sSql, oRs

	sSql = "SELECT characters FROM egov_permitnumberformat WHERE isforbuildingpermits = 1 AND element = 'sequence' AND orgid =  " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitNumberSequenceSize = CLng(oRs("characters"))
	Else
		GetPermitNumberSequenceSize = 1
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Function 



%>
