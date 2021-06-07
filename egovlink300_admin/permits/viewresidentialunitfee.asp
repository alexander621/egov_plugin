<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewresidentialunitfee.asp
' AUTHOR: Steve Loar
' CREATED: 11/05/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views the calculation for a residential unit fee of a permit
'
' MODIFICATION HISTORY
' 1.0   11/05/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, iResidentialUnits, sFeeFormula, iPermitId, sMinFeeAmount, sAppliedFee
Dim sFeeName

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId )

iResidentialUnits = GetResidentialUnits( iPermitId )

sMinFeeAmount = GetMiniumFeeAmount( iPermitFeeId )

sFeeFormula = GetResidentialUnitFeeFormula( iPermitFeeId, iResidentialUnits )

sAppliedFee = GetAppliedFeeAmount( iPermitFeeId )

sFeeName = GetPermitFeeName( iPermitFeeId )

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

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
				<form name="frmFee" action="viewvaluationfee.asp" method="post">
					<p>
						Fee: <%=sFeeName%>
					</p>
					<p> 
						Minimum Fee Amount: <%=sMinFeeAmount%>  <br /><br />
						OR <br /><br />
						<%=sFeeFormula%>
					</p>
					<p>
						<strong>Total of Fees Applied: <%=sAppliedFee%></strong>
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

'-------------------------------------------------------------------------------------------------
' Function GetResidentialUnits( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetResidentialUnits( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(residentialunits,0) AS residentialunits FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetResidentialUnits = oRs("residentialunits")
	Else
		GetResidentialUnits = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetResidentialUnitFeeFormula( iPermitFeeId, iResidentialUnits )
'-------------------------------------------------------------------------------------------------
Function GetResidentialUnitFeeFormula( iPermitFeeId, iResidentialUnits )
	Dim sSql, oRs, sFormula, sFeeAmount

	sSql = "SELECT F.permitfeeid, ISNULL(S.baseamount,0.00) AS baseamount, S.atleastqty, S.notmorethanqty, "
	sSql = sSql & " ISNULL(S.unitamount,0.00) AS unitamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitresidentialunitstepfees S "
	sSql = sSql & " WHERE F.isresidentialunittypefee = 1 AND F.permitfeeid = S.permitfeeid "
	sSql = sSql & " AND " & iResidentialUnits & " >= S.atleastqty AND " & iResidentialUnits & " < S.notmorethanqty "
	sSql = sSql & " AND F.permitfeeid = " & iPermitFeeId

	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFeeAmount = CLng(iResidentialUnits)
		sFeeAmount = CLng(sFeeAmount) - CLng(oRs("atleastqty"))
		sFeeAmount = sFeeAmount * oRs("unitamount")
		sFeeAmount = sFeeAmount + oRs("baseamount")

		sFormula = "( " & CLng(iResidentialUnits)
		sFormula = sFormula & " - " & oRs("atleastqty") & " )"
		sFormula = sFormula & " * " & FormatNumber(oRs("unitamount"),2,,,0) & ") + " & FormatNumber(oRs("baseamount"),2,,,0) & " = " 
		sFormula = sFormula & FormatNumber(sFeeAmount,2,,,0)
	Else
		sFormula = "No Fee Formula Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetResidentialUnitFeeFormula = sFormula
End Function 



%>
