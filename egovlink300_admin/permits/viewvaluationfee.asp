<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewvaluationfee.asp
' AUTHOR: Steve Loar
' CREATED: 04/16/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views the calculation for a valuation fee of a permit
'
' MODIFICATION HISTORY
' 1.0   04/16/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, sJobValue, sFeeFormula, iPermitId, sMinFeeAmount, sAppliedFee
Dim sFeeName

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId )

sJobValue = GetPermitJobValue( iPermitId )

sMinFeeAmount = GetMiniumFeeAmount( iPermitFeeId )

sFeeFormula = GetValuationFormula( iPermitFeeId, sJobValue )

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
' Function GetPermitJobValue( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitJobValue( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(jobvalue,0.00) AS jobvalue FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitJobValue = FormatNumber(oRs("jobvalue"),2,,,0)
	Else
		GetPermitJobValue = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetValuationFormula( iPermitFeeId, dJobValue )
'-------------------------------------------------------------------------------------------------
Function GetValuationFormula( iPermitFeeId, dJobValue )
	Dim sSql, oRs, sFormula, sFeeAmount

	sSql = "SELECT F.permitfeeid, ISNULL(S.baseamount,0.00) AS baseamount, S.atleastvalue, S.notmorethanvalue, "
	sSql = sSql & " ISNULL(S.unitqty,0) AS unitqty, ISNULL(S.unitamount,0.00) AS unitamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitvaluationstepfees S "
	sSql = sSql & " WHERE F.isvaluationtypefee = 1 AND F.permitfeeid = S.permitfeeid "
	sSql = sSql & " AND " & dJobValue & " >= S.atleastvalue AND " & dJobValue & " < S.notmorethanvalue "
	sSql = sSql & " AND F.permitfeeid = " & iPermitFeeId

	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("unitqty")) > CLng(0) Then 
			sFeeAmount = Ceiling(CDbl(dJobValue / oRs("unitqty"))) * oRs("unitqty")
			sFeeAmount = CDbl(sFeeAmount) - CDbl(oRs("atleastvalue"))
			sFeeAmount = sFeeAmount / oRs("unitqty")
			sFeeAmount = sFeeAmount * oRs("unitamount")
			sFeeAmount = sFeeAmount + oRs("baseamount")

			sFormula = "((((Ceiling(" & dJobValue & " / " & oRs("unitqty") & ") * " & oRs("unitqty") & ") - "
			sFormula = sFormula & FormatNumber(oRs("atleastvalue"),2,,,0)
			sFormula = sFormula & ") / " & oRs("unitqty") & ") * "
			sFormula = sFormula & FormatNumber(oRs("unitamount"),2,,,0) & ") + " & FormatNumber(oRs("baseamount"),2,,,0) & " = " 
			sFormula = sFormula & FormatNumber(sFeeAmount,2,,,0)
			sFormula = sFormula & "<br />* The Ceiling() function rounds a number up to an integer."
		Else
			sFormula = "Base Amount = " & FormatNumber(oRs("baseamount"),2,,,0)
		End If 
	Else
		sFormula = "No Fee Formula Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetValuationFormula = sFormula
End Function 



%>
