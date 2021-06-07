<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manualfeeedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits a manually entered fee for a permit
'
' MODIFICATION HISTORY
' 1.0   04/10/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, sMethod, sFeeName, iPermitId, bPermitIsCompleted, bIsOnHold

iPermitFeeId = CLng(request("permitfeeid"))

sSql = "SELECT F.permitfeeprefix, F.permitfee, ISNULL(F.feeamount,0.00) AS feeamount, M.permitfeemethod, F.permitid "
sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitfeeid = " & iPermitFeeId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	If CDbl(oRs("feeamount")) > CDbl(0.00) Then 
		sFeeAmount = FormatNumber(oRs("feeamount"),2,,,0)
	Else
		sFeeAmount = ""
	End If 
	If oRs("permitfeeprefix") <> "" Then
		sFeeName = oRs("permitfeeprefix") & " "
	End If 
	sFeeName = sFeeName & oRs("permitfee")
	sMethod = oRs("permitfeemethod")
	iPermitId = oRs("permitid")
Else
	iPermitId = 0
End If 

oRs.Close
Set oRs = Nothing 

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

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

		<script language="Javascript">
		<!--
		function init()
		{
			document.getElementById("feeamount").focus();
		}

		function doUpdate() 
		{
			// validate the fee amount
			if (document.getElementById("feeamount").value != '')
			{
				// Remove any extra spaces
				document.getElementById("feeamount").value = removeSpaces(document.getElementById("feeamount").value);
				//Remove commas that would cause problems in validation
				document.getElementById("feeamount").value = removeCommas(document.getElementById("feeamount").value);

				rege = /^-?\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("feeamount").value);
				if ( ! Ok )
				{
					alert("The fee amount must be numeric with up to two decimal places.\nPlease correct this and try saving again.");
					document.getElementById("feeamount").focus();
					return;
				}
				else
				{
					document.getElementById("feeamount").value = format_number(Number(document.getElementById("feeamount").value),2);
				}
			}
			else
			{
				document.getElementById("feeamount").value = format_number(Number(0.00),2);
			}
			//alert("OK");

			//Do Ajax save call
			doAjax('manualfeeupdate.asp', 'permitfeeid=' + document.frmFee.permitfeeid.value + '&feeamount=' + document.frmFee.feeamount.value, 'UpdateFees', 'get', '0');
		}

		function UpdateFees( sReturn )
		{
			// Update the fee amount
			parent.document.getElementById("fee" + document.frmFee.permitfeeid.value).innerHTML = document.frmFee.feeamount.value;
			// Get the new fee total for the permit
			doAjax('getpermitfeetotal.asp', 'permitid=' + document.frmFee.permitid.value, 'UpdateTotalFees', 'get', '0');
		}

		function UpdateTotalFees( sReturn )
		{
			if ( sReturn != "Failed" )
			{
				parent.document.getElementById("feetabfeetotal").innerHTML = sReturn;
				parent.document.getElementById("invoicetabfeetotal").innerHTML = sReturn;
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
				<form name="frmFee" action="manualfeeedit.asp" method="post">
					<input type="hidden" name="permitfeeid" value="<%=iPermitFeeId%>" />
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p> 
						<table cellpadding="5" cellspacing="0" border="0" id="manualfeetable">
							<tr><td align="right" class="manualfeelabel">Fee:</td><td><strong><%=sFeeName%></strong></td></tr>
							<!--<tr><td align="right" class="manualfeelabel">Method:</td><td><%=sMethod%></td></tr>-->
							<tr><td align="right" class="manualfeelabel">Fee Amount:</td><td><input type="input" name="feeamount" id="feeamount" value="<%=sFeeAmount%>" size="10" maxlength="10" /></td></tr>
						</table>
						<br />
<%					
					tooltipclass=""
					tooltip = ""
					disabled = ""
					If bPermitIsCompleted or bIsOnHold Then
						tooltipclass="tooltip"
						disabled = " disabled "
						tooltip = "<span class=""tooltiptext"">You cannot edit fixtures because:<br />The permit is complete or on hold.</span>"
					end if
					%>
						<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="doUpdate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp; 
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

%>
