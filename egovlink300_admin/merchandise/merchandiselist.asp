<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiselist.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	4/24/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sMerchandise

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "merchandise setup", sLevel	' In common.asp

If request("merchandise") <> "" Then
	sMerchandise = request("merchandise")
Else
	sMerchandise = ""
End If 

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--
	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Merchandise Management</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
	 	<fieldset class="filterselection">
			<legend class="filterselection">Search Options</legend>
			<p>
				<form name="MerchandiseList" method="post" action="merchandiselist.asp">
					<table border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td>Merchandise Item Like: </td>
							<td colspan="2">
								<input type="text" name="merchandise" value="<%=sMerchandise%>" size="100" maxlength="100" />
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td colspan="2"><input class="button" type="submit" value="Refresh Results" /></td>
						</tr>
					</table>
				</form>
			</p>
 		</fieldset>
	</div>
	<!--END: FILTER SELECTION-->

		<p>
			<input type="button" class="button" value="New Merchandise Item" onclick="location.href='merchandiseedit.asp?merchandiseid=0';" /> &nbsp;
		</p>

	<!--BEGIN: Merchandise LIST-->

	<% 
		DisplayMerchandise sMerchandise
	%>

	<!--END: Merchandise LIST-->
	</div>
</div>

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub DisplayMerchandise( sMerchandise )
'--------------------------------------------------------------------------------------------------
Sub DisplayMerchandise( sMerchandise )
	Dim sSql, sWhere, oRs

	If sMerchandise <> "" Then 
		sWhere = ""
	Else
		sWhere = ""
	End If 

	sSql = "SELECT merchandiseid, merchandise, price FROM egov_merchandise WHERE orgid = " & session("orgid") & " ORDER BY merchandise"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'DRAW TABLE WITH MERCHANDISE LISTED
		response.write vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""merchandiselist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		
		'HEADER ROW
		response.write vbcrlf & "<tr><th>Merchandise Items</th><th>Price</th></tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY The EVENTS
		Do While Not oRs.EOF
  			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 

			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" 

			response.write "<td align=""left"" onClick=""location.href='merchandiseedit.asp?merchandiseid=" & oRs("merchandiseid") & "';"">" 
			response.write oRs("merchandise")
			response.write "</td>"
			response.write "<td align=""center"" onClick=""location.href='merchandiseedit.asp?merchandiseid=" & oRs("merchandiseid") & "';"">" 
			response.write FormatNumber(oRs("price"),2)
			response.write "</td>"
			response.write "</tr>"

  			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 

	Else
		response.write "<p><font color=""red""><b>No merchandise could be could be found.</b></font></p>"
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>

