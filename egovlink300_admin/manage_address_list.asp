<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError, sLatitude, sLongitude

sLevel = "" ' Override of value from common.asp

PageDisplayCheck "address list", sLevel	' In common.asp

' Set Timezone information into session
Session("iUserOffset") = request.cookies("tz")

if request.servervariables("REQUEST_METHOD") = "POST" then
	set oCmd = server.createobject("ADODB.Connection")
	oCmd.Open Application("DSN")
	if request.querystring("postAddressUpdate") = "true" then
			'Process Address Edit

			sLatitude = request("latitude")
			If sLatitude = "" Then
				sLatitude = "NULL"
			End If 
			sLongitude = request("longitude")
			If sLongitude = "" Then
				sLongitude = "NULL"
			End If 
			
			'Insert New address
			if request.form("ID") = 0 then
				sSQL = "INSERT INTO egov_residentaddresses (orgid,residentstreetnumber,residentstreetname,residentcity,residentstate,residenttype,latitude,longitude,sortstreetname,residentzip) VALUES(" & session("OrgID") & ",'" & DBSafe(UCase(request.form("Number"))) & "','" & DBSafe(UCase(request.form("name"))) & "','" & DBSafe(UCase(request.form("city"))) & "','" & DBSafe(UCase(request.form("state"))) & "','" & request.form("Type") & "'," & sLatitude & "," & sLongitude & ",'" & DBSafe(UCase(request.form("name"))) & "', '" & DBSafe(request.form("zip")) & "')"
				oCmd.Execute(sSQL)
				'response.write sSQL
				'response.end
			'Update Address
			else
				sSQL = "UPDATE egov_residentaddresses SET residentstreetnumber = '" & DBSafe(UCase(request.form("number"))) & "', residentstreetname = '" & DBSafe(UCase(request.form("name"))) & "', residentcity = '" & DBSafe(UCase(request.form("city"))) & "', residentstate = '" & DBSafe(UCase(request.form("state"))) & "', residentzip = '" & DBSafe(request.form("zip")) & "', residenttype = '" & DBSafe(request.form("type")) & "', latitude = " & sLatitude & ", longitude = " & sLongitude & ", sortstreetname='" & DBSafe(UCase(request.form("name"))) & "' WHERE residentaddressID = " & request.form("ID")
				'response.write sSQL
				'response.end
				oCmd.Execute(sSQL)
			end if

			'Reload parent window and close this window
			%>
			<script>
			window.opener.doPageReload();
			opener.window.focus();
			self.close();
			</script>
			<%response.end
	else
		for each fieldname in request.form
			'Process Address Delete
			'response.write fieldname & "=" & request.form(fieldname) & "<br>"
			arrIDs = split(request.form(fieldname),", ")
			For x=0 to UBOUND(arrIDs)
				'response.write "'" & arrIDs(x) & "'<br>"
				sSQL = "DELETE FROM egov_residentaddresses WHERE residentaddressid = " & arrIDs(x)
				'response.write sSQL & "<br>"
				oCmd.Execute(sSQL)
			Next
			
		next%>
			<script>
	    		window.location = 'manage_address_list.asp?<%=request.servervariables("QUERY_STRING")%>'
			</script>
		<%
	end if
	oCmd.Close
	Set oCmd = Nothing
end If

if request.querystring("editaddresspop")="true" Then
	' This is code to display the New Address page
	sSQL = "SELECT * FROM egov_residentaddresses WHERE residentaddressid = " & CLng(request("ID"))

    Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	%>
	
<html>
<head>
	<link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />

	<script language="JavaScript" src="scripts/removespaces.js"></script>
	<script language="JavaScript" src="scripts/removecommas.js"></script>

	<script language="Javascript" > 
	<!--
		function Validate()
		{
			var rege;
			var Ok; 

			// Remove any extra spaces
			document.editaddress.number.value = removeSpaces(document.editaddress.number.value);
			//Remove commas that would cause problems in validation
			document.editaddress.number.value = removeCommas(document.editaddress.number.value);

			// Check that the streen number is numeric
			if (document.editaddress.number.value != '')
			{
				rege = /^\d+$/;
				Ok = rege.test(document.editaddress.number.value);

				if (! Ok)
				{
					alert("The street number must be a whole number, or blank.");
					document.editaddress.number.focus();
					return;
				}
			}

			// check the Latitude
			if (document.editaddress.latitude.value.length > 0)
			{
				rege = /^-?\d{1,3}\.\d+$/;
				Ok = rege.test(document.editaddress.latitude.value);

				if (! Ok)
				{
					alert("The latitude must be a number, or blank\n and in the range 90 to -90.");
					document.editaddress.latitude.focus();
					return;
				}
				else
				{
					if (document.editaddress.latitude.value > 90 || document.editaddress.latitude.value < -90)
					{
						alert("The latitude must be a number, or blank\n and in the range 90 to -90.");
						document.editaddress.latitude.focus();
						return;
					}
				}
			}
			// check the Longitude
			if (document.editaddress.longitude.value.length > 0)
			{
				rege = /^-?\d{1,3}\.\d+$/;
				Ok = rege.test(document.editaddress.longitude.value);

				if (! Ok)
				{
					alert("The longitude must be a number, or blank\n and in the range 180 to -180");
					document.editaddress.longitude.focus();
					return;
				}
				else
				{
					if (document.editaddress.longitude.value > 180 || document.editaddress.longitude.value < -180)
					{
						alert("The longitude must be a number, or blank\n and in the range 180 to -180.");
						document.editaddress.longitude.focus();
						return;
					}
				}
			}
			document.editaddress.submit();
		}
	//-->
	</script>

</head>
<body>
<div id="content">
	<div id="centercontent">

	<form name="editaddress" method="post" action="manage_address_list.asp?postAddressUpdate=true">
		<input type="hidden" name="id" value="<% if not oRs.EOF then%><%=oRs("residentaddressid")%><%else%>0<%end if%>">
		<table id="newaddress">
			<tr>
				<th align="right">Street Number</th>
				<td><input type="text" name="number" maxlength="10" value="<%if not oRs.EOF then response.write oRs("residentstreetnumber")%>" /></td>
			</tr>
			<tr>
				<th align="right">Street Name</th>
				<td><input type=text name=name maxlength=50 value="<%if not oRs.EOF then response.write oRs("residentstreetname")%>"></td>
			</tr>
			<tr>
				<th align="right">City</th>
				<td><input type=text name=city maxlength=50 value="<%if not oRs.EOF then response.write oRs("residentcity")%>"></td>
			</tr>
			<tr>
				<th align="right">State</th>
				<td><input type=text name=state maxlength=50 value="<%if not oRs.EOF then response.write oRs("residentstate")%>"><br></td>
			</tr>
			<tr>
				<th align="right">Zip</th>
				<td><input type="text" name="zip" size="10" maxlength="10" value="<%if not oRs.EOF then response.write oRs("residentzip")%>" /><br></td>
			</tr>
			<tr>
				<th align="right">Address Type</th>
				<td>
					<input type="radio" name="type" value=R<% if not oRs.EOF then%><% if oRs("residenttype") = "R" then response.write " checked=""checked"" "%><% end if %>>Residential<br />
					<input type="radio" name="type" value=B<% if not oRs.EOF then%><% if oRs("residenttype") = "B" then response.write " checked=""checked"" "%><% end if %>>Business
				</td>
			</tr>
			<tr>
				<th colspan="2">For Mapping enter the Map Coordinates below.<br />You can find them <a href="http://www.batchgeocode.com/lookup/" target="_blank">here.</a></th>
			</tr>
			<tr>
				<th align="right">Latitude</th>
				<td><input type="text" name="latitude" maxlength="11" value="<%if not oRs.EOF then response.write oRs("latitude")%>" /><br /></td>
			</tr>
			<tr>
				<th align="right">Longitude</th>
				<td><input type="text" name="longitude" maxlength="11" value="<%if not oRs.EOF then response.write oRs("longitude")%>" /><br /></td>
			</tr>
			<tr>
				<td align="right"><input type="button" value="Save" class="button" onClick="Validate();" /></td>
				<td align="left"><input type="button" class="button" value="Cancel" onClick="window.opener.focus();self.close();" /></td>
			</tr>
		</table>
	</form>
</div>
</div>
</body>
</html>
	<%
	oRs.Close
	Set oRs = Nothing
	response.end
end if

%>

<html>
<head>
  <title><%=langBSHome%></title>

	<link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />

	<script language="Javascript" src="scripts/modules.js"></script>
	<script language="Javascript" src="scripts/selectall.js"></script>

	<script language="Javascript" > 
	<!--
		var w = (screen.width - 640)/2;
		var h = (screen.height - 450)/2;
		//Set timezone in cookie to retrieve later
		var d=new Date()
		if (d.getTimezoneOffset)
		{
			var iMinutes = d.getTimezoneOffset();
			document.cookie = "tz=" + iMinutes;
		}

		function doPicker(sFormField) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=435,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPageReload() {
			window.location = 'manage_address_list.asp?<%=request.servervariables("QUERY_STRING")%>';
		}

		function DeleteConfirm()
		{
			if (confirm('Are you sure you want to delete the selected addresses?'))
			{
				document.DeleteAddress.submit();
			}
		}

		function ShowAll()
		{
			document.Search.keyword.value = '';
			document.Search.submit();
		}


	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<% ShowHeader sLevel %>
	<!--#Include file="menu/menu.asp"--> 

	<%
		'pagesize=Session("PageSize")
		pagesize = GetUserPageSize( Session("UserId") ) ' Steve Loar 2/20/2007

		if request("Keyword") <> "" Then
			sKeyword = "&keyword=" & keyword
			keyword = dbsafe(request("Keyword"))
			'sSQL = "SELECT * FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " AND (residentstreetname LIKE '%" & keyword & "%' OR residentstreetnumber LIKE '%" & keyword & "%' OR parcelidnumber = '" & keyword & "')  ORDER BY sortStreetName, Cast(residentstreetnumber as INT),ResidentCity"
			sSQL = "SELECT * FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " AND (residentstreetname LIKE '%" & keyword & "%' OR residentstreetnumber LIKE '%" & keyword & "%' OR parcelidnumber = '" & keyword & "')  ORDER BY sortStreetName, Cast(residentstreetnumber as decimal),ResidentCity"
		Else
			sKeyword = ""
			'sSQL = "SELECT * FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " ORDER BY sortStreetName, Cast(residentstreetnumber as INT),ResidentCity"
			'sSQL = "SELECT * FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " ORDER BY sortStreetName, Cast(replace(residentstreetnumber,' ','') as decimal),ResidentCity"
			sSQL = "SELECT * FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " ORDER BY sortStreetName, Cast(Replace(LEFT(SUBSTRING(residentstreetnumber, PATINDEX('%[0-9.-]%', residentstreetnumber), 8000), PATINDEX('%[^0-9.-]%', SUBSTRING(residentstreetnumber, PATINDEX('%[0-9.-]%', residentstreetnumber), 8000) + 'X') -1),'-','') as decimal),ResidentCity"
		end if

   		Set oRs = Server.CreateObject("ADODB.Recordset")

		oRs.PageSize = pagesize
		oRs.CacheSize = pagesize
		oRs.CursorLocation = 3
		oRs.Open sSQL, Application("DSN"), 3,1

		if request("pagenum") <> "" then
			pagenum = request("pagenum")
			sPageNum = "&pagenum=" & pagenum
		else
			pagenum = 0
			sPageNum = ""
		end if

		If (len(pagenum) = 0 or clng(pagenum) < 1) and not oRs.EOF then
			oRs.AbsolutePage = 1
		elseif not oRs.EOF then
			if clng(Request("pagenum")) <= oRs.PageCount then
				oRs.AbsolutePage = Request("pageNum")
			else
				oRs.AbsolutePage = 1
			end if
		end if

		Dim abspage, pagecnt
		abspage = oRs.AbsolutePage
		pagecnt = oRs.PageCount
	%>
<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="images/icon_home.jpg"></td>-->
      <td colspan="2">
          <font style="font-size:14px;"><b>Resident Address Administration</b></font><br>
        <div style="padding-top:2px;">Number of Addresses:<%=oRs.RecordCount%>, Number of Pages:<%=pagecnt%></div>
	    </td>
    </tr>
    <tr>
      <!--<td valign="top" width="151">
		<font size="1" face="Verdana,Tahoma,Arial"><b><%'=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="images/poweredby.jpg" align="center" border="0"></A></div>
	  -->
        <!-- END: QUICK LINKS MODULE //-->

      <!--</td>-->
      <td valign="top">
      		<%if request("keyword") <> "" then 
			limit="keyword=" & request("keyword") & "&"
		end if %>
		<form name="Search" method="submit" action="manage_address_list.asp">
            <input type="text" name="keyword" style="width:200px;" maxlength="50" value="<%=Request("keyword")%>" />&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value="Search" onclick="javascript:if(document.Search.keyword.value != '') {document.Search.submit();} else {alert('Please enter a keyword for your search.');}" />	&nbsp; &nbsp; 
			<input type="button" class="button" value="Display All" onclick="ShowAll();" />
          	<div style="padding-bottom:5px;">
				<span style="font-size:10px;">(ex: "10522" OR "ACREWOOD")</span>
			</div>
		</form>


		<form name="DeleteAddress" method="post" action=#>
		<div style='font-size:10px; padding-bottom:10px;'>
			<% if abspage > 1 then%><a href="manage_address_list.asp?<%=limit%>pagenum=<%=abspage-1%>"><% end if %><img src='images/arrow_back.gif' align='absmiddle' border=0>&nbsp;<%=langPrev%>&nbsp;<%=pagesize%><% if abspage > 1 then%></a><% end if %>&nbsp;&nbsp;
			<% if abspage < pagecnt then%><a href="manage_address_list.asp?<%=limit%>pagenum=<%=abspage+1%>"><% end if %><%=langNext%>&nbsp;<%=pagesize%><img src='images/arrow_forward.gif' align='absmiddle' border=0><% if abspage < pagecnt then%></a><% end if %>&nbsp;&nbsp;&nbsp;&nbsp;

			<input type="button" class="button" value="New Address" onClick="location.href='addressedit.asp?residentaddressid=0';" />&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value="Delete Selected Addresses" onclick="DeleteConfirm();" />
		</div>

		<div class="shadow">
		<table border="0" cellpadding="5" cellspacing="0" class="tablelist">
			<tr>
				<th align="left"><input type="checkbox" onClick="selectAll('DeleteAddress', this.checked, 'delete');"></th>
				<th align="center">Street Number</th>
				<th align="center">Street Name</th>
				<th align="center">Unit/<br />Suite</th>
				<th align="center">City</th>
				<th align="center">State</th>
				<th align="center">Zip</th>
				<th align="center">Type</th>
				<th align="center">Latitude</th>
				<th align="center">Longitude</th>
			</tr>
		<%
		iRowCount = 0
		For intRec=1 To oRs.PageSize
			If Not oRs.EOF Then 
				iRowCount = iRowCount + 1
				If iRowCount Mod 2 = 0 Then
					sClass = " class=""altrow"" "
				Else
					sClass = ""
				End If 
				%>
				<tr id="<%=iRowCount%>" <%=sClass%> onMouseOver="mouseOverRow(this);" onMouseOut="mouseOutRow(this);">
					<td>
						<input type="checkbox" name="delete" value="<%=oRs("residentaddressid")%>" />
					</td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("residentstreetnumber")%></td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';">
					<%
					If oRs("residentstreetprefix") <> "" Then 
						response.write oRs("residentstreetprefix") & " " 
					End If 
					response.write oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If
					%>	
					</td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("residentunit")%></td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("residentcity")%></td>
					<td title="click to edit" align="center" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("residentstate")%></td>
					<td title="click to edit" align="center" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("residentzip")%></td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%if oRs("residenttype") = "R" then response.write "Residental"%><%if oRs("residenttype") = "B" then response.write "Business"%></td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("latitude")%></td>
					<td title="click to edit" onClick="location.href='addressedit.asp?residentaddressid=<%=oRs("residentaddressid")%><%=sKeyword%><%=sPageNum%>';"><%=oRs("longitude")%></td>
				</tr>
				<%oRs.MoveNext
			End If 
		Next

		If iRowCount = 0 And request("keyword") <> "" Then
			response.write vbcrlf & "<td colspan=""10""> &nbsp; No address could be found that matches your search criteria.</td>"
		End If 
		%>
		</table>
		</div>
		</form>
      </td>
    </tr>
  </table>
  
   </div>
 </div>

<!--#Include file="admin_footer.asp"-->  

</body>
</html>

<script language=javascript>
<!--
	function openWin2(url, name) 
	{
		popupWin = window.open(url, name,"resizable,width=480,height=300");
	}
//-->
</script>

