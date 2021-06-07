<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpercentagefee.asp
' AUTHOR: Steve Loar
' CREATED: 09/08/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Views the percentage fee type fees of a permit
'
' MODIFICATION HISTORY
' 1.0   09/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, sFeeAmount, iPermitId, sMinFeeAmount, sAppliedFee
Dim sFeeName

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId )

sMinFeeAmount = GetMiniumFeeAmount( iPermitFeeId )

sFeeFormula = GetPercentageFormula( iPermitFeeId, iPermitId )

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
				<form name="frmFee" action="viewpercentagefee.asp" method="post">
					<p>
						Fee: <%=sFeeName%>
					</p>
					<p> 
						Minimum Fee Amount: <%=sMinFeeAmount%>  
						
						<br /><br /> OR <br /><br />

						<%=sFeeFormula%>
					</p>
					<p>
						<strong>Applied Amount: <%=sAppliedFee%></strong>
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
' Function GetPercentageFormula( iPermitFeeId, dJobValue )
'-------------------------------------------------------------------------------------------------
Function GetPercentageFormula( iPermitFeeId, iPermitId )
	Dim sSql, oRs, sFormula, sFeeAmount

	sFormula = "Fee Formula not found."

	sSql = "SELECT percentage, permitfeecategorytypeid FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("permitfeecategorytypeid")) <> CLng(-1) Then 
			sFeeAmount = GetCategoryFeesTotalForPercentage( iPermitId, oRs("permitfeecategorytypeid") )	 ' in permitcommonfunctions.asp
		Else
			sFeeAmount = GetAllFeesTotalForPercentage( iPermitId )	 ' in permitcommonfunctions.asp
		End If 
		sFormula = FormatNumber(sFeeAmount,2) & " * " & FormatNumber(oRs("percentage"),4) & " = " 
		sFeeAmount = sFeeAmount * CDbl(oRs("percentage"))
		sFormula = sFormula & FormatNumber(sFeeAmount,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPercentageFormula = sFormula
End Function 



%>
