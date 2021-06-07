<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturetypelist.asp
' AUTHOR: Steve Loar
' CREATED: 12/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit fixture types
'
' MODIFICATION HISTORY
' 1.0   12/19/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch

sLevel = "../" ' Override of value from common.asp

' Check page availability and user access rights in one call
PageDisplayCheck "permit fixture types", sLevel	' In common.asp

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
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/tablednd.js"></script>

	<script language="Javascript">
	<!--

		function Init()
		{
			var table = $('categorytypes');
			var tableDnD = new TableDnD();
			tableDnD.init(table);

			// Redefine the onDrop so that we can update things
			tableDnD.onDrop = function(table, row) 
			{
				var iRowNo = -1;
				var rows = this.table.tBodies[0].rows;
				var debugStr = 'rows now: ';
				for (var i=0; i<rows.length; i++) 
				{
					iRowNo += 1;
					// skip the header row
					if (iRowNo > 0)
					{
						debugStr += iRowNo + ' = ' + rows[i].id + '\n';
						rows[i].className = iRowNo & 1? '':'altrow';  // set the row background class
						//alert($("oldrow" + rows[i].id).value + ": " + $("permitinspectionid" + rows[i].id).value + " now " + iRowNo);
						// Fire off ajax routine here to reorder the rows to this order
						doAjax('changefixturetypeorder.asp', 'permitfixturetypeid=' + $("permitfixturetypeid" + rows[i].id).value + '&displayorder=' + iRowNo, '', 'get', '0');
						$("oldrow" + rows[i].id).value = iRowNo;
					}
				}
			}
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function commonIFrameUpdateFunction()
		{
			UpdateParentFixtures()
		}
		function UpdateParentFixtures()
		{

			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type=fixtures', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Update the unselected values
				var unSelVals = parent.document.getElementById('newpermitfixturetypeid');
				unSelVals.innerHTML = newDDVals;
				unSelVals.value = unSelVals.options[0].value;

				//Update the already selected values
				var selVals = parent.document.getElementById('permitfixturetypeid').options;
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

<body onload="Init();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Fixture Types</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmFixtureSearch" method="post" action="permitfixturetypelist.asp">
				<div>
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
					&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Fixture Type" onclick="location.href='permitfixturetypeedit.asp?permitfixturetypeid=0';" />
					&nbsp; &nbsp; <input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
					<br /><br />
				</div>
			</form>

			<div class="shadow">
			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>&nbsp;</th><th>Fixture Types</th></tr>
				<%	
					ShowPermitFixtureTypes session("orgid"), sSearch 
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
' Sub ShowPermitFixtureTypes( iOrgid )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFixtureTypes( iOrgid, sSearch )
	Dim sSql, oRs, iRowCount, iPermitFixtureTypeid

	iRowCount = 0
	iPermitFixtureTypeid = CLng(0)
	'sSql = "SELECT permitfixturetypeid, permitfixture, unitqty, unitfeeamount, usesteptable "
	sSql = "SELECT permitfixturetypeid, permitfixture "
	sSql = sSql & " FROM egov_permitfixturetypes "
	sSql = sSql & " WHERE orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND permitfixture LIKE '%" & dbsafe(sSearch) & "%' "
	End If 
	sSql = sSql & " ORDER BY displayorder, permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iPermitFixtureTypeid <> CLng(oRs("permitfixturetypeid")) Then 
				If iPermitFixtureTypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iPermitFixtureTypeid = CLng(oRs("permitfixturetypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

				' Drag and Drop Icon
				response.write "<td align=""center""><img src=""..\images\up_down_arrow.gif"" class=""DRAGIMG"" width=""13"" height=""19"" border=""0"" alt=""drag and drop"" />"
				response.write "<input type=""hidden"" id=""permitfixturetypeid" & iRowCount & """ name=""permitfixturetypeid" & iRowCount & """ value=""" & oRs("permitfixturetypeid") & """ />"
				response.write "<input type=""hidden"" id=""oldrow" & iRowCount & """ name=""oldrow" & iRowCount & """ value=""" & iRowCount & """ />"
				response.write "</td>"

				' Fixture Type
				response.write "<td class=""leftcol"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='permitfixturetypeedit.asp?permitfixturetypeid=" & oRs("permitfixturetypeid") & "';"">&nbsp;" & oRs("permitfixture") & "</td>"
			End If 
			oRs.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td>&nbsp;No Fixture Types could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td>&nbsp;No Fixture Types could be found. Click on the New Fixture Type button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
