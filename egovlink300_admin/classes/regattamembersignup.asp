<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattamembersignup.asp
' AUTHOR: Steve Loar
' CREATED: 03/10/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	3/10/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassId, iCartId, sReturnTo, iUserId, iMaxTeamMembers, iTeamId, sClassName, sDetails
Dim sStartDate, sRegistrationStart, sRegistrationEnd, iItemTypeId, iRegattaSignupTypeId
Dim iClassSeasonId, iTeamClassId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

iClassId = CLng(request("classid"))
iMaxTeamMembers = CLng(0)

GetClassDetails iClassId, sClassName, sDetails, sStartDate, sRegistrationStart, sRegistrationEnd, iRegattaSignupTypeId, iClassSeasonId

' Get the classid of the team signup for this
iTeamClassId = GetRegattaClassId( iClassSeasonId, "isteamsignup" )

iItemTypeId = GetItemTypeIdBySignupTypeId( iRegattaSignupTypeId )		' In class_global_functions.asp

If request("cartid") <> "" Then
	iCartId = CLng(request("cartid"))
	iUserId = GetCartValue( iCartId, "userid" )
	iTeamId = GetCartValue( iCartId, "regattateamid" )
Else
	iCartId = 0
	iTeamId = 0
	iUserId = 0
End If 

If CLng(iCartId) = CLng(0) Then
	sReturnTo = "regattalist"
	sSaveButton = "Add To Cart"
Else
	sReturnTo = "class_cart"
	sSaveButton = "Save Changes"
End If 

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--
		var tabView;
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function autoTab(input,len, e, bSetFlag) 
		{
			var keyCode = (isNN) ? e.which : e.keyCode; 
			var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

			if(input.value.length >= len && !containsElement(filter,keyCode)) {
				input.value = input.value.slice(0, len);
			var addNdx = 1;

			while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
			{
				addNdx++;
				//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
			}

			input.form[(getIndex(input)+addNdx) % input.form.length].focus();
		}

		function containsElement(arr, ele) 
		{
			var found = false, index = 0;

			while(!found && index < arr.length)
				if(arr[index] == ele)
					found = true;
				else
					index++;
			return found;
		}

		function getIndex(input) 
		{
			var index = -1, i = 0, found = false;

			while (i < input.form.length && index == -1)
				if (input.form[i] == input)index = i;
				else i++;
					return index;
		}
			return true;
		}

		function upperCase(x)
		{
			var y = document.getElementById(x).value;
			document.getElementById(x).value = y.toUpperCase();
		}

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.AddTeam.searchname.value);
			var searchtext = document.AddTeam.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.AddTeam.egovuserid.length ; x++)
			{
				optiontext = document.AddTeam.egovuserid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.AddTeam.egovuserid.selectedIndex = x;
					document.AddTeam.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.AddTeam.searchstart.value = x;
					//document.AddTeam.submit();
					return;
				}
			}
			document.AddTeam.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.AddTeam.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.AddTeam.searchstart.value = -1;
		}

		function UserPick()
		{
			document.AddTeam.searchname.value = '';
			document.AddTeam.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.AddTeam.searchstart.value = -1;
			//document.AddTeam.submit();
		}

		function EditUser()
		{
			var iUserId = document.AddTeam.egovuserid.options[document.AddTeam.egovuserid.selectedIndex].value;
			location.href='../dirs/update_citizen.asp?userid=' + iUserId;
		}

		function NewUser()
		{
			location.href='../dirs/register_citizen.asp';
		}

		function SearchTeams( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			var searchtext = document.AddTeam.searchteam.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.AddTeam.teamid.length ; x++)
			{
				optiontext = document.AddTeam.teamid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.AddTeam.teamid.selectedIndex = x;
					document.AddTeam.teamresults.value = 'Possible Match Found.';
					document.getElementById('searchteamresults').innerHTML = 'Possible Match Found.';
					document.AddTeam.searchteamstart.value = x;
					return;
				}
			}
			document.AddTeam.teamresults.value = 'No Match Found.';
			document.getElementById('searchteamresults').innerHTML = 'No Match Found.';
			document.AddTeam.searchteamstart.value = -1;
		}

		function ClearTeamSearch()
		{
			document.AddTeam.searchteamstart.value = -1;
		}

		function TeamPick()
		{
			document.AddTeam.searchteam.value = '';
			document.AddTeam.teamresults.value = '';
			document.getElementById('searchteamresults').innerHTML = '';
			document.AddTeam.searchteamstart.value = -1;
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
		}

		function addTeamMemberLine()
		{
			document.getElementById("maxteammembers").value = parseInt(document.getElementById("maxteammembers").value) + 1;
			var tbl = document.getElementById("teammemberentry");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.getElementById("maxteammembers").value);
			var row = tbl.insertRow(lastRow);

			// Class text
			cellOne = row.insertCell(0);
			cellOne.align = 'left';
			e1 = document.createElement('input');
			e1.type = 'text';
			e1.name = 'regattateammember' + newRow;
			e1.id = 'regattateammember' + newRow;
			e1.size = '50';
			e1.maxLength = '50';
			cellOne.appendChild(e1);
			//alert(document.getElementById("maxteammembers").value);
			e1.focus();
		}

		function ValidatePrice( oPrice )
		{
			clearMsg('unitprice');

			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					tabView.set('activeIndex',3);
					//oPrice.value = format_number(0,2);
					oPrice.value = '';
					document.getElementById(oPrice.id).focus();
					inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices must be numbers in currency format.',5,oPrice.id);
					return false;
				}
			}
			else
			{
				tabView.set('activeIndex',3);
				//oPrice.value = format_number(0,2);
				oPrice.value = '';
				document.getElementById(oPrice.id).focus();
				inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices cannot be blank.',5,oPrice.id);
				return false;
			}
			return true;
		}

		function validate()
		{
			// Make sure a registered user is picked
			if (document.AddTeam.egovuserid.options[document.AddTeam.egovuserid.selectedIndex].value == 0 )
			{
				tabView.set('activeIndex',0);
				$("egovuserid").focus();
				inlineMsg(document.getElementById("egovuserid").id,'<strong>Invalid Selection: </strong>Please select a registered user.',5,document.getElementById("egovuserid").id);
				return;
			}

			// A team is always picked so we do not have to worry about that

			// Make sure that at least one name is entered
			var bHasMembers = false;
			var MaxMembers = parseInt(document.getElementById("maxteammembers").value);
			for (var x = 1; x <= MaxMembers ; x++ )
			{
				if ($("regattateammember" + x).value != '')
				{
					bHasMembers = true;
					break;
				}
			}

			if (bHasMembers == false)
			{
				tabView.set('activeIndex',2);
				$("regattateammember1").focus();
				//alert("Please enter at least one person to add to the team.");
				inlineMsg(document.getElementById("regattateammember1").id,'<strong>Missing Information: </strong>Please enter at least one person to add to the team.',5,document.getElementById("regattateammember1").id);
				return;
			}
			
			// check the price again due to how Javascript works the submit will happen even on bad prices.
			bIsValid = ValidatePrice( $("unitprice") );
			
			if (bIsValid)
			{
				//alert("Valid");
				document.AddTeam.submit();
			}
		}

		function ViewTeam( )
		{
			var teamid = document.AddTeam.teamid.options[document.AddTeam.teamid.selectedIndex].value;
			winHandle = eval('window.open("viewteamdetails.asp?regattateamid=' + teamid + '", "_details", "width=900,height=500,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=' + h + '")');
		}

	//-->
	</script>
</head>
<body class="yui-skin-sam">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>

	<p>
		<input type="button" class="button" value="<< Back" onclick="javascript:location.href='<%=sReturnTo%>.asp';" />
	</p>

	<h3><%=sClassName%></h3>
	<fieldset>
		<legend><strong> Details </strong></legend>
		<p>
			<%=sDetails%>
		</p>
		<p>
			<strong>Event Date:</strong> <%=sStartDate%>
		</p>
		<p>
			<strong>Registration Starts:</strong> <%=sRegistrationStart%>
		</p>
		<p>
			<strong>Registration Ends:</strong> <%=sRegistrationEnd%>
		</p>
		<p>
			<strong>Waivers:</strong> <% ShowClassWaiverLinks iClassId  %>
		</p>
	</fieldset>

	<form name="AddTeam" action="regattamembertocart.asp" method="post">
	
	<input type="hidden" name="classid" value="<%=iClassId%>" />
	<input type="hidden" name="cartid" value="<%=iCartId%>" />
	<input type="hidden" name="itemtypeid" value="<%=iItemTypeId%>" />

	<div id="demo" class="yui-navset">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Purchaser</em></a></li>
			<li><a href="#tab2"><em>Team Selection</em></a></li>
			<li><a href="#tab3"><em>New Team Members</em></a></li>
			<li><a href="#tab4"><em>Price Selection</em></a></li>
		</ul>            
		<div class="yui-content">

			<div id="tab1"> <!-- Purchaser Information -->
				<p><br />
					Select the registered user who is making the purchase.<br /><br />
<%					' Show pick of registered users and their detail info.
					ShowRegisteredUsers iUserId
%>
				</p>
			</div>
			<div id="tab2"> <!-- Team Information -->
				<p><br />
					Select the team that is having members added to it.<br /><br />
<%					' Show pick of regatta teams and their captains.
					ShowRegattaTeams iTeamClassId, iTeamId
%>
				</p>
			</div>
			<div id="tab3"> <!-- Team Members -->
				<p><br />
					Enter Team Member Names in the lines below. Add extra team member lines as needed. Extra lines can be left blank.
				</p>
				<p>
					<input type="button" class="button" value="Add A Team Member Line" onclick="addTeamMemberLine();" />
				</p>
				<table id="teammemberentry" cellpadding="0" cellspacing="0" border="0">
					<tr><th colspan="2">Team Member Name</th></tr>
<%					' Get team members		
					iMaxTeamMembers = ShowRegattaTeamMembers( CLng(iCartId), iTeamId )
%>
				</table>
				<input type="hidden" name="maxteammembers" id="maxteammembers" value="<%=iMaxTeamMembers%>" />
			</div>
			<div id="tab4"> <!-- Price Selection -->
				<p><br />
					Select the price to be applied per team member.
				</p>
				<p>
<%					ShowRegattaPriceOptions iClassId, iCartId	%>
				</p>
			</div>
		</div>
	</div>
	
	<p>
		<input type="button" class="button" value="<%=sSaveButton%>" onclick="validate();" />
	</p>

	</form>

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
' Sub ShowRegattaPriceOptions( iClassId, iCartId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegattaPriceOptions( iClassId, iCartId )
	Dim sSql, oRs

	If CLng(iCartId) > CLng(0) Then
		sSql = "SELECT P.pricetypeid, P.pricetypename, A.unitprice AS unitprice, C.amount AS baseprice "
		sSql = sSql & " FROM egov_price_types P, egov_class_pricetype_price C, egov_class_cart_price A "
		sSql = sSql & " WHERE P.isregattaprice = 1 AND P.orgid = " & session("orgid")
		sSql = sSql & " AND A.pricetypeid = P.pricetypeid AND A.cartid = " & iCartId
		sSql = sSql & " AND P.pricetypeid = C.pricetypeid AND C.classid = " & iClassId
	Else
		sSql = "SELECT P.pricetypeid, P.pricetypename, C.amount AS unitprice, C.amount AS baseprice "
		sSql = sSql & " FROM egov_price_types P, egov_class_pricetype_price C "
		sSql = sSql & " WHERE P.isregattaprice = 1 AND P.orgid = " & session("orgid")
		sSql = sSql & " AND P.pricetypeid = C.pricetypeid AND C.classid = " & iClassId
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""regattapricetable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		response.write vbcrlf & "<tr><td nowrap=""nowrap"">"
		response.write "<input type=""radio"" checked=""checked"" id=""pricetypeidpick"" name=""pricetypeidpick"" value=""" & oRs("pricetypeid") & """ /> "
		response.write "<input type=""hidden"" id=""pricetypeid"" name=""pricetypeid"" value=""" & oRs("pricetypeid") & """ /> "
		response.write " &nbsp; " & oRs("pricetypename") & "</td>"
		response.write "<td>"
		response.write "<input type=""text"" id=""unitprice"" name=""unitprice"" value=""" & FormatNumber(CDbl(oRs("unitprice")),2,,,0) & """ size=""10"" maxlength=""9"" onchange=""return ValidatePrice(this);"" />"
		response.write "</td>"
		response.write "<td nowrap=""nowrap"">&nbsp;" & FormatCurrency(oRs("baseprice")) & " Per Team Member</td>"
		response.write "</tr></table>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegisteredUsers( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUsers( iUserId )

	response.write vbcrlf & "<p>Name Search: <input type=""text"" name=""searchname"" value="""" size=""25"" maxlength=""50"" onchange=""javascript:ClearSearch();"" />"
	response.write vbcrlf & "<input type=""button"" class=""button"" value=""Search"" onclick=""javascript:SearchCitizens(document.AddTeam.searchstart.value);"" /> &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:NewUser();"" value=""New User"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""results"" value="""" />"
	response.write vbcrlf & "<input type=""hidden"" name=""searchstart"" value="""" />"
	response.write vbcrlf & "<span id=""searchresults""> </span>"
	response.write vbcrlf & "<br /><div id=""searchtip"">(last name, first name)</div>"
	response.write vbcrlf & "</p>"
	response.write vbcrlf & "<p>"
	response.write vbcrlf & "Select Name: <select id=""egovuserid"" name=""egovuserid"" onchange=""javascript:UserPick();"">"
	
	' Create the user pick dropdown
	ShowUserDropDown iUserId 
	
	response.write vbcrlf & "</select>"
	response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:EditUser();"" value=""Edit User Profile"" />"
	response.write vbcrlf & "</p>" 
	response.write vbcrlf & "<div id=""userinfo""> </div>"
End Sub 


'------------------------------------------------------------------------------
' Sub ShowUserDropDown(iUserId)
'------------------------------------------------------------------------------
Sub ShowUserDropDown( iUserId )
	Dim oCmd, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oRs = .Execute
	End With

	response.write vbcrlf & "<option value=""0"">Select a Registered User...</option>"
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"
		If CLng(iUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("userlname") & ", " & oRs("userfname") & " &ndash; " & oRs("useraddress") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegattaTeams( iTeamClassId, iTeamId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegattaTeams( iTeamClassId, iTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateamid, regattateam, captainname FROM egov_regattateams WHERE classid = " & iTeamClassId
	sSql = sSql & " ORDER BY regattateam, captainname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<p>Team Search: <input type=""text"" id=""searchteam"" name=""searchteam"" value="""" size=""60"" maxlength=""100"" onchange=""javascript:ClearTeamSearch();"" />"
		response.write vbcrlf & "<input type=""button"" class=""button"" value=""Search"" onclick=""javascript:SearchTeams(document.AddTeam.searchteamstart.value);"" />"
		response.write vbcrlf & "<input type=""hidden"" name=""teamresults"" value="""" />"
		response.write vbcrlf & "<input type=""hidden"" name=""searchteamstart"" value="""" />"
		response.write vbcrlf & "<span id=""searchteamresults""> </span>"
		response.write vbcrlf & "<br /><div id=""searchteamtip"">(team name or captain name)</div>"
		response.write vbcrlf & "</p>"
		response.write vbcrlf & "<p>"
		response.write vbcrlf & "Select Team: <select id=""teamid"" name=""teamid"" onchange=""javascript:TeamPick();"">"
		
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("regattateamid") & """"
			If CLng(iTeamId) = CLng(oRs("regattateamid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("regattateam") & " (" & oRs("captainname") & ")</option>"
			oRs.MoveNext
		Loop 
		
		response.write vbcrlf & "</select>"
		response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:ViewTeam();"" value=""View Team Details"" />"
		response.write vbcrlf & "</p>" 
		response.write vbcrlf & "<div id=""userinfo""> </div>"
	Else
		response.write vbcrlf & "<p>No teams have been setup.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' Function ShowRegattaTeamMembers( iCartId, iTeamId )
'------------------------------------------------------------------------------
Function ShowRegattaTeamMembers( iCartId, iTeamId )
	Dim sSql, oRs, iRowCount, x

	iRowCount = CLng(0) 
	sSql = "SELECT regattateammember FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId
	sSql = sSql & " AND cartteamid = " & iTeamId & " ORDER BY regattateammember"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount = iRowCount + CLng(1)
			response.write vbcrlf & "<tr><td><input type=""text"" id=""regattateammember" & iRowCount & """ name=""regattateammember" & iRowCount & """ value=""" & oRs("regattateammember") & """ size=""50"" maxlength=""50"" /></td></tr>"
			oRs.MoveNext 
		Loop
	Else
		' Give them 5 rows to start with
		iRowCount = CLng(5)
		For x = 1 To 5
			response.write vbcrlf & "<tr><td><input type=""text"" id=""regattateammember" & x & """ name=""regattateammember" & x & """ value="""" size=""50"" maxlength=""50"" /></td></tr>"
		Next 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowRegattaTeamMembers = iRowCount

End Function 


'------------------------------------------------------------------------------
' Sub GetClassDetails( iClassId, sClassName, sDetails, sStartDate, sRegistrationStart, sRegistrationEnd, iRegattaSignupTypeId )
'------------------------------------------------------------------------------
Sub GetClassDetails( ByVal iClassId, ByRef sClassName, ByRef sDetails, ByRef sStartDate, ByRef sRegistrationStart, ByRef sRegistrationEnd, ByRef iRegattaSignupTypeId, ByRef iClassSeasonId )
	Dim sSql, oRs

	iRowCount = CLng(0) 
	sSql = "SELECT classname, classdescription, startdate, registrationstartdate, registrationenddate, regattasignuptypeid, classseasonid "
	sSql = sSql & " FROM egov_class WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sClassName = oRs("classname")
		sDetails = oRs("classdescription")
		sStartDate = oRs("startdate")
		sRegistrationStart = oRs("registrationstartdate")
		sRegistrationEnd = oRs("registrationenddate")
		iRegattaSignupTypeId = oRs("regattasignuptypeid")
		iClassSeasonId = oRs("classseasonid")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' Sub GetCartTeamValues( iCartId, iTeamId, sTeamName, sCaptainName, sCaptainAddress, sCaptainCity, sCaptainState, sCaptainZip, sCaptainPhone )
'------------------------------------------------------------------------------
Sub GetCartTeamValues( ByVal iCartId, ByRef iTeamId, ByRef sTeamName, ByRef sCaptainName, ByRef sCaptainAddress, ByRef sCaptainCity, ByRef sCaptainState, ByRef sCaptainZip, ByRef sCaptainPhone )
	Dim sSql, oRs

	iRowCount = CLng(0) 
	sSql = "SELECT cartteamid, regattateam, captainname, captainaddress, captaincity, captainstate, captainzip, captainphone "
	sSql = sSql & " FROM egov_class_cart_regattateams WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iTeamId = oRs("cartteamid")
		sTeamName = oRs("regattateam")
		sCaptainName = oRs("captainname")
		sCaptainAddress = oRs("captainaddress")
		sCaptainCity = oRs("captaincity")
		sCaptainState = oRs("captainstate")
		sCaptainZip = oRs("captainzip")
		sCaptainPhone = oRs("captainphone")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 




%>
