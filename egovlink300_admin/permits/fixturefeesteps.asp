<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: fixturefeesteps.asp
' AUTHOR: Steve Loar
' CREATED: 04/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Displays the fixture fees and allows input of quantities and to select for the permit.
'
' MODIFICATION HISTORY
' 1.0   04/30/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFixtureId, sFixtureName

iPermitFixtureId = CLng(request("permitfixtureid"))

sFixtureName = GetFixtureName( iPermitFixtureId )  ' In permitcommonfunctions.asp


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
				<form name="frmFee" action="fixturefeeupdate.asp" method="post">
				<input type="hidden" name="permitfeeid" value="<%=iPermitFixtureId%>" />
					<p>
						Fixture: <%=sFixtureName%>
					</p>
					<p>
						<% ShowStepTable iPermitFixtureId %>
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
' Sub ShowStepTable( iPermitFixtureId )
'-------------------------------------------------------------------------------------------------
Sub ShowStepTable( iPermitFixtureId )
	Dim sSql, oRs, iFixtureCount

	iFixtureCount = 0

	sSql = "SELECT atleastqty, notmorethanqty, unitqty, unitamount, baseamount "
	sSql = sSql & " FROM egov_permitfixturestepfees WHERE permitfixtureid = " & iPermitFixtureId
	sSql = sSql & " ORDER BY atleastqty"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vrcrlf & "<table cellpadding=""2"" cellspacing=""0"" border=""0"" class=""tableadmin"">"
		response.write vbcrlf & "<tr><th>At Least Qty</th><th>Not More Than Qty</th><th>Unit Qty</th><th>Unit Amount</th><th>Base Amount</th></tr>"
		Do While Not oRs.EOF
			iFixtureCount = iFixtureCount + 1
			response.write vbcrlf & "<tr"
			If iFixtureCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write vbcrlf & "<td align=""center"">" & FormatNumber(oRs("atleastqty"),0) & "</td>"
			response.write vbcrlf & "<td align=""center"">" & FormatNumber(oRs("notmorethanqty"),0) & "</td>"
			response.write vbcrlf & "<td align=""center"">" & FormatNumber(oRs("unitqty"),0) & "</td>"
			response.write vbcrlf & "<td align=""center"">" & FormatNumber(oRs("unitamount"),2) & "</td>"
			response.write vbcrlf & "<td align=""center"">" & FormatNumber(oRs("baseamount"),2)& "</td>"
			response.write vbcrlf & "</tr>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
