<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: hourlyratefeeedit.asp
' AUTHOR: Steve Loar
' CREATED: 05/07/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views an hourly rate fee for a permit
'
' MODIFICATION HISTORY
' 1.0   05/07/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, sMethod, sFeeName, iPermitId, sHours, sRate, sMultipliers
Dim sBase, sUnitQty, sMinFeeAmount, sCalcFeeAmount

iPermitFeeId = CLng(request("permitfeeid"))

sSql = "SELECT F.permitid, F.permitfeeprefix, F.permitfee, F.feeamount, M.permitfeemethod, F.permitid, ISNULL(F.minimumamount,0.00) AS minimumamount, "
sSql = sSql & " ISNULL(F.unitqty,1) AS unitqty, ISNULL(F.unitamount,0.00) AS unitamount, ISNULL(F.baseamount,0.00) AS baseamount "
sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitfeeid = " & iPermitFeeId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	sFeeAmount = FormatNumber(oRs("feeamount"),2,,,0)
	sHours = GetPermitExamHours( oRs("permitid") )
	sRate = FormatNumber(oRs("unitamount"),2,,,0)
	sMultipliers = GetFeeMultipliersForDisplay( iPermitFeeId )
	sBase = FormatNumber(oRs("baseamount"),2,,,0)
	sUnitQty = oRs("unitqty")
	If oRs("permitfeeprefix") <> "" Then
		sFeeName = oRs("permitfeeprefix") & " "
	End If 
	sFeeName = sFeeName & oRs("permitfee")
	sMethod = oRs("permitfeemethod")
	iPermitId = oRs("permitid")
	sMinFeeAmount = FormatNumber(oRs("minimumamount"),2,,,0)

	' Get the calculated fee amount
	sCalcFeeAmount = CDbl(sHours)
	sCalcFeeAmount = sCalcFeeAmount * GetFeeMultipliers( iPermitFeeId )
	If CLng(oRs("unitqty")) > CLng(1) Then 
		sCalcFeeAmount = sCalcFeeAmount / oRs("unitqty")
	End If 
	If CDbl(oRs("unitamount")) > CDbl(0) Then 
		sCalcFeeAmount = sCalcFeeAmount * CDbl(oRs("unitamount"))
	End If 
	sCalcFeeAmount = sCalcFeeAmount + CDbl(oRs("baseamount"))
	sCalcFeeAmount = FormatNumber(sCalcFeeAmount,2,,,0)
End If 

oRs.Close
Set oRs = Nothing 

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

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmFee" action="hourlyratefeeedit.asp" method="post">
					<input type="hidden" name="permitfeeid" value="<%=iPermitFeeId%>" />
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p><strong>Fee: <%=sFeeName%></strong></p>
					<p> 
						Minimum Fee Amount: <%=sMinFeeAmount%>  <br /><br />
						OR <br /><br />
						<table cellpadding="2" cellspacing="0" border="0" id="hourlyratefeetable">
							<tr>
								<th>Hours</th>
<%								If CDbl(sRate) > CDbl(1.00) Then %>
									<th>* Rate</th>
<%								End If %>
<%								If sMultipliers <> "" Then %>
									<th>* Multipliers</th>
<%								End If %>
<%								If CLng(sUnitQty) > CLng(1) Then %>
									<th>/ Unit Qty</th>
<%								End If %>
<%								If CDbl(sBase) > CDbl(0.00) Then %>
									<th>+ Base Amount</th>
<%								End If %>
								<th>=</th><th>Fee Amount</th>
							</tr>
							<tr>
								<td align="center"><%=sHours%></td>
<%								If CDbl(sRate) > CDbl(1.00) Then %>
									<td align="center">* <%=sRate%></td>
<%								End If %>
<%								If sMultipliers <> "" Then %>
									<td align="center"><%=sMultipliers%></td>
<%								End If %>
<%								If CLng(sUnitQty) > CLng(1) Then %>
									<th>/ <%=sUnitQty%></th>
<%								End If %>
<%								If CDbl(sBase) > CDbl(0.00) Then %>
									<th>+ <%=sBase%></th>
<%								End If %>
								<td align="center">=</td><td align="center"><%=sCalcFeeAmount%></td>
							</tr>
						</table>
					</p>
					<p>
						<strong>Total of Fees Applied: <%=sFeeAmount%></strong>
					</p>
					<p>
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
