<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewsqfootagefee.asp
' AUTHOR: Steve Loar
' CREATED: 08/18/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views the calculation for a Sq Footage fee of a permit
'
' MODIFICATION HISTORY
' 1.0   08/18/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, dblSqFootage, sFeeFormula, iPermitId, sMinFeeAmount, sAppliedFee
Dim sFeeName, sSqFootageType

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId ) ' Get the permit id

sSqFootageType = GetSqFootageType( iPermitFeeId ) ' Get which sq footage to use

dblSqFootage = GetPermitSqFootage( iPermitId, sSqFootageType ) ' Pull the correct sq footage

sMinFeeAmount = GetMiniumFeeAmount( iPermitFeeId ) ' Pull out for display

sFeeFormula = GetFeeFormula( iPermitFeeId, dblSqFootage ) ' The formula to display

sAppliedFee = GetAppliedFeeAmount( iPermitFeeId ) ' The fee amount applied

sFeeName = GetPermitFeeName( iPermitFeeId )  ' Get the fee name to display

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
			//window.close();
			//window.opener.focus();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmFee" action="viewsqfootagefee.asp" method="post">
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
' Function GetPermitSqFootage( iPermitId, sSqFootageType )
'-------------------------------------------------------------------------------------------------
Function GetPermitSqFootage( ByVal iPermitId, ByVal sSqFootageType )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(" & sSqFootageType & ",0.00) AS sqfootage FROM egov_permits WHERE permitid = " & iPermitId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitSqFootage = FormatNumber(oRs("sqfootage"),2,,,0)
	Else
		GetPermitSqFootage = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetFeeFormula( iPermitFeeId, dblSqFootage )
'-------------------------------------------------------------------------------------------------
Function GetFeeFormula( iPermitFeeId, dblSqFootage )
	Dim sSql, oRs, sFormula, sFeeAmount, sShowFees

	sFormula = ""
	sShowFees = ""

	sSql = "SELECT permitfeeid, ISNULL(baseamount,0.00) AS baseamount,  "
	sSql = sSql & " ISNULL(unitqty,0) AS unitqty, ISNULL(unitamount,0.0000) AS unitamount "
	sSql = sSql & " FROM egov_permitfees "
	sSql = sSql & " WHERE permitfeeid = " & iPermitFeeId

	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFormula = "( " & FormatNumber(dblSqFootage,2,,,0)
		sFeeAmount = CDbl(dblSqFootage)

		' sq footage * unit amount / unit qty
		If CLng(oRs("unitqty")) > CLng(0) Then 
			sFormula = sFormula & " / " & oRs("unitqty")
			sFormula = sFormula & " * " & FormatNumber(oRs("unitamount"),4,,,0)
			sFeeAmount = ( sFeeAmount / CLng(oRs("unitqty")) ) * CDbl(oRs("unitamount"))
		End If	

		' Multipliers 
		sFeeAmount = sFeeAmount * GetFeeMultipliers( iPermitFeeId, sShowFees )
		If sShowFees <> "" Then
			sFormula = sFormula & " * " & sShowFees
		End If

		sFormula = sFormula & " )"

		' Base Amounts added 
		If CDbl(oRs("baseamount")) > CDbl(0.00) Then 
			sFormula = sFormula & " + " & FormatNumber(oRs("baseamount"),2,,,0) 
			sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
		End If 
		sFormula = sFormula & " = " & FormatNumber(sFeeAmount,2,,,0)
	Else
		sFormula = "No Fee Formula Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetFeeFormula = sFormula
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetSqFootageType( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetSqFootageType( ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = " SELECT F.permitfeeid, M.istotalsqft, M.isfinishedsqft, M.isunfinishedsqft, M.isothersqft "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("istotalsqft") Then
			GetSqFootageType = "totalsqft"
		End If 
		If oRs("isfinishedsqft") Then
			GetSqFootageType = "finishedsqft"
		End If 
		If oRs("isunfinishedsqft") Then
			GetSqFootageType = "unfinishedsqft"
		End If 
		If oRs("isothersqft") Then
			GetSqFootageType = "othersqft"
		End If 
	Else
		GetSqFootageType = "totalsqft"
	End If
	
	oRs.Close
	Set oRs = Nothing 


End Function 


'-------------------------------------------------------------------------------------------------
' Function GetFeeMultipliers( iPermitFeeId, sShowFees )
'-------------------------------------------------------------------------------------------------
Function GetFeeMultipliers( ByVal iPermitFeeId, ByRef sShowFees )
	Dim sSql, oRs, sRate

	sRate = CDbl(1.0)
	sShowFees = ""
	sSql = "SELECT feemultiplierrate FROM egov_permitfeemultipliers WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' multiply the rates together
		sRate = CDbl(sRate) * CDbl(oRs("feemultiplierrate"))
		If sShowFees <> "" Then 
			sShowFees = sShowFees & " * "
		End If 
		sShowFees = sShowFees & FormatNumber(oRs("feemultiplierrate"),4,,,0)
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetFeeMultipliers = sRate
End Function 


%>
