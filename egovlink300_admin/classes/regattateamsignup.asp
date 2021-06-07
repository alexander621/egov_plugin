<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamsignup.asp
' AUTHOR: Steve Loar
' CREATED: 02/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	2/26/2009	Steve Loar	-	Initial version
' 1.1	4/7/2010	Steve Loar - Removing team member code, adding team group pick
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassId, iCartId, sReturnTo, iUserId, sTeamName, sCaptainFirstName, sCaptainLastName, sCaptainAddress, sCaptainCity
Dim sCaptainState, sCaptainZip, sCaptainPhone, iMaxTeamMembers, iCartTeamId, sClassName, sDetails
Dim sStartDate, sRegistrationStart, sRegistrationEnd, iItemTypeId, iRegattaSignupTypeId, iRegattaTeamGroupId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

iClassId = CLng(request("classid"))
iMaxTeamMembers = CLng(0)

GetClassDetails iClassId, sClassName, sDetails, sStartDate, sRegistrationStart, sRegistrationEnd, iRegattaSignupTypeId

iItemTypeId = GetItemTypeIdBySignupTypeId( iRegattaSignupTypeId )		' In class_global_functions.asp

If request("cartid") <> "" Then
	iCartId = CLng(request("cartid"))
	iUserId = GetCartValue( iCartId, "userid" )
	GetCartTeamValues iCartId, iCartTeamId, sTeamName, sCaptainFirstName, sCaptainLastName, sCaptainAddress, sCaptainCity, sCaptainState, sCaptainZip, sCaptainPhone, iRegattaTeamGroupId
Else
	iCartId = 0
	iCartTeamId = 0
	iUserId = 0
	sTeamName = ""
	sCaptainFirstName = ""
	sCaptainLastName = ""
	sCaptainAddress = ""
	sCaptainCity = ""
	sCaptainState = ""
	sCaptainZip = ""
	sCaptainPhone = ""
	iRegattaTeamGroupId = 0
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

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--
		var tabView;

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
					doAjax('getcaptaininfo.asp', 'userid=' + document.AddTeam.egovuserid.options[document.AddTeam.egovuserid.selectedIndex].value, 'UpdateCaptainInfo', 'get', '0');

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
			doAjax('getcaptaininfo.asp', 'userid=' + document.AddTeam.egovuserid.options[document.AddTeam.egovuserid.selectedIndex].value, 'UpdateCaptainInfo', 'get', '0');
		}

		function UpdateCaptainInfo( sReturnJSON )
		{
			var json = sReturnJSON.evalJSON(true); 
			if (document.AddTeam.egovuserid.options[document.AddTeam.egovuserid.selectedIndex].value > 0)
			{
				if (json.flag == 'success')
				{
					$("captainfirstname").value = json.captainfirstname;
					$("captainlastname").value = json.captainlastname;
					$("captainaddress").value = json.captainaddress;
					$("captaincity").value = json.captaincity;
					$("captainstate").value = json.captainstate;
					$("captainzip").value = json.captainzip;
					$("areacode").value = json.areacode;
					$("exchange").value = json.exchange;
					$("line").value = json.line;
				}
			}
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
					tabView.set('activeIndex',2);
					//oPrice.value = format_number(0,2);
					oPrice.value = '';
					document.getElementById(oPrice.id).focus();
					inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices must be numbers in currency format.',5,oPrice.id);
					return false;
				}
			}
			else
			{
				tabView.set('activeIndex',2);
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

			// Make sure all fields are entered for the team info
			if ($F("regattateam") == '')
			{
				tabView.set('activeIndex',1);
				$("regattateam").focus();
				inlineMsg(document.getElementById("regattateam").id,'<strong>Invalid Team Name: </strong>Please enter a team name.',5,document.getElementById("regattateam").id);
				return;
			}
			if ($F("captainfirstname") == '')
			{
				tabView.set('activeIndex',1);
				$("captainfirstname").focus();
				inlineMsg(document.getElementById("captainfirstname").id,'<strong>Invalid Name: </strong>Please enter a first name.',5,document.getElementById("captainfirstname").id);
				return;
			}
			if ($F("captainlastname") == '')
			{
				tabView.set('activeIndex',1);
				$("captainlastname").focus();
				inlineMsg(document.getElementById("captainlastname").id,'<strong>Invalid Name: </strong>Please enter a last name.',5,document.getElementById("captainlastname").id);
				return;
			}
			if ($F("captainaddress") == '')
			{
				tabView.set('activeIndex',1);
				$("captainaddress").focus();
				inlineMsg(document.getElementById("captainaddress").id,'<strong>Invalid Address: </strong>Please enter an address.',5,document.getElementById("captainaddress").id);
				return;
			}
			if ($F("captaincity") == '')
			{
				tabView.set('activeIndex',1);
				$("captaincity").focus();
				inlineMsg(document.getElementById("captaincity").id,'<strong>Invalid City: </strong>Please enter a city.',5,document.getElementById("captaincity").id);
				return;
			}
			if ($F("captainstate") == '')
			{
				tabView.set('activeIndex',1);
				$("captainstate").focus();
				inlineMsg(document.getElementById("captainstate").id,'<strong>Invalid State: </strong>Please a state.',5,document.getElementById("captainstate").id);
				return;
			}
			if ($F("captainzip") == '')
			{
				tabView.set('activeIndex',1);
				$("captainzip").focus();
				inlineMsg(document.getElementById("captainzip").id,'<strong>Invalid Zip: </strong>Please enter a zip.',5,document.getElementById("captainzip").id);
				return;
			}
			// Check the phone
			var captainphone = document.getElementById("areacode").value;
            captainphone = captainphone + document.getElementById("exchange").value;
            captainphone = captainphone + document.getElementById("line").value;
			$("captainphone").value = captainphone;
			//alert(captainphone);
			if ( captainphone == '' || captainphone.length < 10 )
			{
				tabView.set('activeIndex',1);
				$("line").focus();
				inlineMsg(document.getElementById("line").id,'<strong>Invalid phone: </strong>Please enter a complete phone number.',5,document.getElementById("line").id);
				return;
			}
			else
			{
				var rege = /^\d+$/;
				var Ok = rege.exec(captainphone);
				if ( ! Ok )
				{
					tabView.set('activeIndex',1);
					$("line").focus();
					inlineMsg(document.getElementById("line").id,'<strong>Invalid phone: </strong>Please enter a valid, numeric phone number.',5,document.getElementById("line").id);
					return;
				}
			}

			// check the price again due to how Javascript works the submit will happen even on bad prices.
			bIsValid = ValidatePrice( $("unitprice") );
			
			if (bIsValid)
			{
				// Do not care about teammembers, so submit
				//alert("Valid");
				document.AddTeam.submit();
			}
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

	<form name="AddTeam" action="regattateamtocart.asp" method="post">
	
	<input type="hidden" name="classid" value="<%=iClassId%>" />
	<input type="hidden" name="cartid" value="<%=iCartId%>" />
	<input type="hidden" name="itemtypeid" value="<%=iItemTypeId%>" />

	<div id="demo" class="yui-navset">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Purchaser</em></a></li>
			<li><a href="#tab2"><em>Team Information</em></a></li>
			<li><a href="#tab3"><em>Price Selection</em></a></li>
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
					Enter the Team Name and Captain Information.<br /><br />
					Team Name: <input type="text" id="regattateam" name="regattateam" value="<%=sTeamName%>" size="100" maxlength="100" />
				</p>
				<p>
					Team Group: <% ShowTeamGroups iRegattaTeamGroupId %>
				</p>
				<table id="captaininfoentry" cellpadding="0" cellspacing="0" border="0">
					<tr><th colspan="2">Captain Information</th></tr>
					<tr><td align="right">Name:&nbsp;</td>
						<td><input type="text" id="captainfirstname" name="captainfirstname" value="<%=sCaptainFirstName%>" size="25" maxlength="25" />
							&nbsp; <input type="text" id="captainlastname" name="captainlastname" value="<%=sCaptainLastName%>" size="25" maxlength="25" />
						</td>
					</tr>
					<tr><td align="right">Address:&nbsp;</td><td><input type="text" id="captainaddress" name="captainaddress" value="<%=sCaptainAddress%>" size="50" maxlength="50" /></td></tr>
					<tr><td align="right">City:&nbsp;</td><td><input type="text" id="captaincity" name="captaincity" value="<%=sCaptainCity%>" size="50" maxlength="50" /></td></tr>
					<tr><td align="right">State:&nbsp;</td><td><input type="text" id="captainstate" name="captainstate" value="<%=sCaptainState%>" size="2" maxlength="2" onchange="upperCase(this.id);" /></td></tr>
					<tr><td align="right">Zip:&nbsp;</td><td><input type="text" id="captainzip" name="captainzip" value="<%=sCaptainZip%>" size="10" maxlength="10" /></td></tr>
					<tr><td align="right">Phone:&nbsp;</td><td nowrap="nowrap">
						<input type="hidden" value="<%=sCaptainPhone%>" id="captainphone" name="captainphone" />
						(<input type="text" value="<%=Left(sCaptainPhone,3)%>" id="areacode" name="areacode" onKeyUp="return autoTab(this, 3, event, true);" size="3" maxlength="3" />)&nbsp;
						<input type="text" value="<%=Mid(sCaptainPhone,4,3)%>" id="exchange" name="exchange" onKeyUp="return autoTab(this, 3, event, true);" size="3" maxlength="3" />&ndash;
						<input type="text" value="<%=Right(sCaptainPhone,4)%>" id="line" name="line" size="4" maxlength="4" />					
					</td></tr>
				</table>
			</div>
			<div id="tab3"> <!-- Price Selection -->
				<p><br />
					Select the price to be applied.
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
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowRegattaPriceOptions( iClassId, iCartId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegattaPriceOptions( ByVal iClassId, ByVal iCartId )
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
		response.write "<td nowrap=""nowrap"">&nbsp;" & FormatCurrency(oRs("baseprice")) & " " & oRs("pricetypename") & " </td>"
		response.write "</tr></table>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRegisteredUsers( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUsers( ByVal iUserId )

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
	response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:EditUser();"" value=""Edit the Selected User's Profile"" />"
	response.write vbcrlf & "</p>" 
	response.write vbcrlf & "<div id=""userinfo""> </div>"
End Sub 


'------------------------------------------------------------------------------
' void ShowUserDropDown(iUserId)
'------------------------------------------------------------------------------
Sub ShowUserDropDown( ByVal iUserId )
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


'------------------------------------------------------------------------------
' integer ShowRegattaTeamMembers( iCartId, iCartTeamId )
'------------------------------------------------------------------------------
Function ShowRegattaTeamMembers( ByVal iCartId, ByVal iCartTeamId )
	Dim sSql, oRs, iRowCount, x

	iRowCount = CLng(0) 
	sSql = "SELECT regattateammember FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId
	sSql = sSql & " AND cartteamid = " & iCartTeamId & " ORDER BY regattateammember"

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
' void GetClassDetails( iClassId, sClassName, sDetails, sStartDate, sRegistrationStart, sRegistrationEnd, iRegattaSignupTypeId )
'------------------------------------------------------------------------------
Sub GetClassDetails( ByVal iClassId, ByRef sClassName, ByRef sDetails, ByRef sStartDate, ByRef sRegistrationStart, ByRef sRegistrationEnd, ByRef iRegattaSignupTypeId )
	Dim sSql, oRs

	iRowCount = CLng(0) 
	sSql = "SELECT classname, classdescription, startdate, registrationstartdate, registrationenddate, regattasignuptypeid "
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
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void GetCartTeamValues( iCartId, iCartTeamId, sTeamName, sCaptainFirstName, sCaptainLastName, sCaptainAddress, sCaptainCity, sCaptainState, sCaptainZip, sCaptainPhone, iRegattaTeamGroupId )
'------------------------------------------------------------------------------
Sub GetCartTeamValues( ByVal iCartId, ByRef iCartTeamId, ByRef sTeamName, ByRef sCaptainFirstName, ByRef sCaptainLastName, ByRef sCaptainAddress, ByRef sCaptainCity, ByRef sCaptainState, ByRef sCaptainZip, ByRef sCaptainPhone, ByRef iRegattaTeamGroupId )
	Dim sSql, oRs

	iRowCount = CLng(0) 
	sSql = "SELECT cartteamid, regattateam, captainfirstname, captainlastname, captainaddress, captaincity, captainstate, "
	sSql = sSql & "captainzip, captainphone, regattateamgroupid "
	sSql = sSql & " FROM egov_class_cart_regattateams WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iCartTeamId = oRs("cartteamid")
		sTeamName = oRs("regattateam")
		sCaptainFirstName = oRs("captainfirstname")
		sCaptainLastName = oRs("captainlastname")
		sCaptainAddress = oRs("captainaddress")
		sCaptainCity = oRs("captaincity")
		sCaptainState = oRs("captainstate")
		sCaptainZip = oRs("captainzip")
		sCaptainPhone = oRs("captainphone")
		iRegattaTeamGroupId = oRs("regattateamgroupid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 




%>
