<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feemultiplierlist.asp
' AUTHOR: Steve Loar
' CREATED: 12/18/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/18/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit permits", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function commonIFrameUpdateFunction()
		{
			UpdateParentMultipliers('feemultipliers','feemultiplierDD')
		}
		function UpdateParentMultipliers(poptype, classname)
		{

			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type='+poptype+'&value=<%=request("id")%>', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Update the unselected values
				var unSelVals = parent.document.getElementById('newfeemultipliertypeid');
				unSelVals.innerHTML = newDDVals;
				unSelVals.value = unSelVals.options[0].value;

				//Update the already selected values
				var selVals = parent.document.getElementById('feemultipliertypeid').options;
				for (var i = 0; i < selVals.length; i++) {
					var unSelOptions = unSelVals.options;
					for (var j = 0; j < unSelOptions.length; j++) {

						//Find Matching option in unselected values
						if (selVals[i].value == unSelOptions[j].value)
						{
							//Update Name in Selected Values
							selVals[i].innerHTML = unSelOptions[j].innerHTML;
	
							//Purge from unselected values
							unSelVals.removeChild(unSelOptions[j]);
						}
					}
				}

			}

		}
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Fee Multiplier Rates</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmFixtureSearch" method="post" action="feemultiplierlist.asp">
			<div>
				<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
				<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
				&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Multiplier" onclick="location.href='feemultiplieredit.asp?feemultipliertypeid=0';" />
				&nbsp; &nbsp; <input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
				<br /><br />
			</div>

			<div class="shadow" id="feemultipliershadow">
			<table id="feemultipliertypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Multiplier</th><th>Rate</th></tr>
				<%	
					ShowFeeMultiplierTypeRates session("orgid"), sSearch 
				%>
			</table>
			</div>
		</div>
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
' Sub ShowFeeMultiplierTypeRates( iOrgid )
'--------------------------------------------------------------------------------------------------
Sub ShowFeeMultiplierTypeRates( iOrgid, sSearch )
	Dim sSql, oRates, iRowCount, iFeeMultiplierTypeid

	iRowCount = 0
	iFeeMultiplierTypeid = CLng(0)
	sSql = "SELECT feemultipliertypeid, feemultiplier, feemultiplierrate "
	sSql = sSql & " FROM egov_feemultipliertypes "
	sSql = sSql & " WHERE orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND feemultiplier LIKE '%" & dbsafe(sSearch) & "%' "
	End If 
	sSql = sSql & " ORDER BY feemultiplier, feemultiplierrate"

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSQL, Application("DSN"), 3, 1

	If Not oRates.EOF Then
		Do While Not oRates.EOF
			If iOccupancytypeid <> CLng(oRates("feemultipliertypeid")) Then 
				If iFeeMultiplierTypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iFeeMultiplierTypeid = CLng(oRates("feemultipliertypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				response.write "<td class=""leftcol"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='feemultiplieredit.asp?feemultipliertypeid=" & oRates("feemultipliertypeid") & "&id=" & request("id") & "';"">&nbsp;" & oRates("feemultiplier") & "</td>"
			End If 
			' output the rates
			response.write "<td align=""center"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='feemultiplieredit.asp?feemultipliertypeid=" & oRates("feemultipliertypeid") & "';"">" & oRates("feemultiplierrate") & "</td>"
			oRates.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Multipliers could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;Click on the New Multiplier button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRates.Close
	Set oRates = Nothing 

End Sub 



%>
