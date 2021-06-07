<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewconstructiontypefee.asp
' AUTHOR: Steve Loar
' CREATED: 04/25/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views the calculation for a construction type fee of a permit
'
' MODIFICATION HISTORY
' 1.0   04/25/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sAppliedFee, sSqFtField, sFeeFormula, iPermitId, sMinFeeAmount
Dim sSqFtValue, iConstructiontyperate, sFeeName

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId )

sMinFeeAmount = GetMiniumFeeAmount( iPermitFeeId )
sAppliedFee = GetAppliedFeeAmount( iPermitFeeId ) ' In permitcommonfunctions.asp
sSqFtField = GetSqFtField( iPermitFeeId )
sSqFtValue = GetSqFtValue( iPermitId, sSqFtField )
iConstructiontyperate = GetConstructionTypeRate( iPermitId )
sFeeFormula = GetConstructionTypeFormula( iPermitFeeId, sSqFtValue, iConstructiontyperate )
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
				<form name="frmFee" action="viewconstructiontypefee.asp" method="post">
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
' Function GetSqFtField( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetSqFtField( iPermitFeeId )
	Dim sSql, oRs
		
	sSql = "SELECT M.isconstructiontypefinished, M.isconstructiontypegross, M.isconstructiontypeunfinished, M.isconstructiontypeother "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE M.permitfeemethodid = F.permitfeemethodid "
	sSql = sSql & " AND F.permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isconstructiontypegross") Then
			GetSqFtField = "totalsqft"
		ElseIf oRs("isconstructiontypefinished") Then
			GetSqFtField = "finishedsqft"
		ElseIf oRs("isconstructiontypeunfinished") Then
			GetSqFtField = "unfinishedsqft"
		ElseIf oRs("isconstructiontypeother") Then
			GetSqFtField = "othersqft"
		Else
			' DO not know what this is
			GetSqFtField = "totalsqft"
		End If 
	Else
		GetSqFtField = "totalsqft"  ' Should never have this condition
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetSqFtField( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetSqFtValue( iPermitId, sSqFtField )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(" & sSqFtField & ",0.00) AS sqfootage FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetSqFtValue = CDbl(oRs("sqfootage"))
	Else
		GetSqFtValue = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetConstructionTypeRate( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetConstructionTypeRate( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(constructiontyperate,0.00) AS constructiontyperate FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetConstructionTypeRate = CDbl(oRs("constructiontyperate"))
	Else
		GetConstructionTypeRate = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetFeeMultipliers( iPermitFeeId )
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



'-------------------------------------------------------------------------------------------------
' Function GetConstructionTypeFormula( iPermitFeeId, sSqFtValue, iConstructiontyperate )
'-------------------------------------------------------------------------------------------------
Function GetConstructionTypeFormula( iPermitFeeId, sSqFtValue, iConstructiontyperate )
	Dim sSql, oRs, sShowFees, sFormula

	sShowFees = ""

	sFeeAmount = CDbl(sSqFtValue) * CDbl(iConstructiontyperate)
	sFeeAmount = sFeeAmount * GetFeeMultipliers( iPermitFeeId, sShowFees )

	sFormula = FormatNumber(sSqFtValue,2,,,0) & " * " & FormatNumber(iConstructiontyperate,4,,,0)
	If sShowFees <> "" Then
		sFormula = sFormula & " * " & sShowFees
	End If 
	sFormula = sFormula & " = " & FormatNumber(sFeeAmount,2,,,0)

	GetConstructionTypeFormula = sFormula
End Function 



%>
