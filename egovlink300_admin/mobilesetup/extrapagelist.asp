<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: extrapagelist.asp
' AUTHOR: Steve Loar
' CREATED: 08/15/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creation and editing of ad hoc mobile pages
'
' MODIFICATION HISTORY
' 1.0   04/15/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />'
sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "mobileextrapages", sLevel	' In common.asp

If request("s") = "d" Then
	sLoadMsg = "displayScreenMsg('The Page Was Successfully Deleted');"
End If 

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	
	<title>E-GovLink Administration Console</title>

	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="mobilesetupstyles.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<script src="../scripts/ajaxLib.js"></script>
	<script src="../scripts/tablednd.js"></script>
	<script src="../scripts/modules.js"></script>

	<script>
	<!--

		function displayScreenMsg( iMsg ) 
		{
			if( iMsg != "" ) 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>

			var tableElement = document.getElementById('mobilelist');
			var tableDnD = new TableDnD();
			tableDnD.init(tableElement);

			// Redefine the onDrop so that we can update things
			tableDnD.onDrop = function(table, row) 
			{
				
				var tableBody = tableElement.getElementsByTagName("tbody").item(0);
				var rowCount = tableBody.rows.length;

				for (var i=1; i < rowCount; i++) 
				{
					$("tr#" + tableBody.rows[i].id).removeClass("altrow");
					if ( i % 2 === 0) {
						$("tr#" + tableBody.rows[i].id).addClass("altrow");
					}
					doAjax('extrapageorderupdate.asp', 'pageid=' + tableBody.rows[i].id + '&displayorder=' + i , '', 'get', '0');
				}
			}
		}

		function makeNewPage()
		{
			location.href = "extrapageedit.asp?pageid=0";
		}

	//-->
	</script>

</head>
<body onload="SetUpPage();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<h3>Extra Mobile Pages</h3>

		<div id="topbtnsholder">
			<span id="screenMsg"></span>
			<input type="button" class="button" id="newpagebtn" name="newpagebtn" value="Create a New Page" onclick="makeNewPage()" /> &nbsp; 
		</div>

<%		ShowExtraPages		%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'------------------------------------------------------------------------------
' ShowExtraPages
'------------------------------------------------------------------------------
Sub ShowExtraPages()
	Dim sSql, oRs, iRowCount

	iRowCount = CLng(0)

	sSql = "SELECT pageid, ISNULL(pagetitle,'') AS pagetitle, displayorder, displaypage "
	sSql = sSql & "FROM egov_extramobilepages WHERE orgid = " & SESSION("orgid")
	sSql = sSQl & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""instructions"">"
		response.write vbcrlf & "Click a row to edit a page. Drag and Drop the arrows to change the display order of the pages."
		response.write vbcrlf & "</div>"

		response.write vbcrlf & "<table id=""mobilelist"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr noDrop=""true"" noDrag=""true""><th id=""changeordercol"">Display<br />Order</th><th>Page Title</th><th>Is<br />Displayed</tr>"
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			'response.write vbcrlf & "<tr id=""" & iRowCount & """"
			response.write vbcrlf & "<tr id=""" & oRs("pageid") & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			' Display Order
			response.write "<td align=""center""><img src=""..\images\up_down_arrow.gif"" class=""DRAGIMG"" width=""13"" height=""19"" border=""0"" alt=""drag and drop"" />"
			'response.write "<input type=""hidden"" id=""pageid" & iRowCount & """ name=""pageid" & iRowCount & """ value=""" & oRs("pageid") & """ />"
			'response.write "<input type=""hidden"" id=""oldrow" & iRowCount & """ name=""oldrow" & iRowCount & """ value=""" & iRowCount & """ />"
			response.write "</td>"

			' Page Title
			response.write "<td class=""pagetitle"" align=""left"" title=""click to edit"" onClick=""location.href='extrapageedit.asp?pageid=" & oRs("pageid") & "';"" nowrap=""nowrap"">"
			response.write oRs("pagetitle") & "</td>"

			' Is displayed
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='extrapageedit.asp?pageid=" & oRs("pageid") & "';"" nowrap=""nowrap"">"
			If oRs("displaypage") Then 
				response.write "Yes"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</table>"
	Else
		response.write "<p>There are no pages to display.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Sub 


%>