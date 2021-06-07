<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamlist.asp
' AUTHOR: Steve Loar
' CREATED: 02/25/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	2/25/2009	Steve Loar	-	Initial version
' 1.1	4/7/2010	Steve Loar - Modified to remove team member list
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamId, sTeamName, sCaptainFirstname, sCaptainLastname, sCaptainaddress, sCaptaincity
Dim sCaptainstate, sCaptainzip, sCaptainphone, sRegattaTeamGroupName, iRegattaTeamGroupId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

iRegattaTeamId = CLng(request("regattateamid"))

GetTeamInformation iRegattaTeamId

If request("u") <> "" Then
	If request("u") = "s" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If 
End If 


%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		function deleteMember( iMemberId ) 
		{
			if (confirm('Are you certain you want to delete this team member?'))
			{
				location.href = 'regattamemberdelete.asp?regattateammemberid=' + iMemberId + '&regattateamid=<%=iRegattaTeamId%>';
			}
		}

		function editTeamName()
		{
			winHandle = eval('window.open("regattateamnameedit.asp?regattateamid=<%=iRegattaTeamId%>", "_team", "width=700,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

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

		function validate()
		{

			// Make sure all fields are entered for the team info
			if ($F("regattateam") == '')
			{
				$("regattateam").focus();
				inlineMsg(document.getElementById("regattateam").id,'<strong>Invalid Team Name: </strong>Please enter a team name.',5,document.getElementById("regattateam").id);
				return;
			}

			if ($F("captainfirstname") == '')
			{
				$("captainfirstname").focus();
				inlineMsg(document.getElementById("captainfirstname").id,'<strong>Invalid Name: </strong>Please enter a firstname.',5,document.getElementById("captainfirstname").id);
				return;
			}

			if ($F("captainlastname") == '')
			{
				$("captainlastname").focus();
				inlineMsg(document.getElementById("captainlastname").id,'<strong>Invalid Name: </strong>Please enter a lastname.',5,document.getElementById("captainlastname").id);
				return;
			}

			if ($F("captainaddress") == '')
			{
				$("captainaddress").focus();
				inlineMsg(document.getElementById("captainaddress").id,'<strong>Invalid Address: </strong>Please enter an address.',5,document.getElementById("captainaddress").id);
				return;
			}

			if ($F("captaincity") == '')
			{
				$("captaincity").focus();
				inlineMsg(document.getElementById("captaincity").id,'<strong>Invalid City: </strong>Please enter a city.',5,document.getElementById("captaincity").id);
				return;
			}

			if ($F("captainstate") == '')
			{
				$("captainstate").focus();
				inlineMsg(document.getElementById("captainstate").id,'<strong>Invalid State: </strong>Please a state.',5,document.getElementById("captainstate").id);
				return;
			}

			if ($F("captainzip") == '')
			{
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
					$("line").focus();
					inlineMsg(document.getElementById("line").id,'<strong>Invalid phone: </strong>Please enter a valid, numeric phone number.',5,document.getElementById("line").id);
					return;
				}
			}

			//alert("Valid");
			document.frmEditTeam.submit();
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
			$("regattateam").focus();
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

	<span id="screenMsg"></span>

<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>

	<p>
		<input type="button" class="button" value="<< Back" onclick="javascript:location.href='regattalist.asp'" />
	</p>

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>River Regatta Team Roster</strong></font><br /><br />
	</p>
	<!--END: PAGE TITLE-->

	<form name="frmEditTeam" action="regattateamupdate.asp" method="post">
		<input type="hidden" id="regattateamid" name="regattateamid" value="<%=iRegattaTeamId%>" />
		<p>
			<span class="teamnamelabel">Team Name:</span>
			<input type="text" id="regattateam" name="regattateam" value="<%=sTeamName%>" size="100" maxlength="100" />
			<!--<input type="button" class="button" name="editteamname" id="editteamname" value="Edit Team Name" onclick="editTeamName();" /> -->
		</p>

		<p><span class="teamnamelabel">Team Group:</span><% ShowTeamGroups iRegattaTeamGroupId %>
		</p>

		<table id="captaindata" cellpadding="3" cellspacing="0" border="0">
			<tr>
				<td>
					<table id="captaininfoentry" cellpadding="0" cellspacing="0" border="0">
						<tr><th colspan="2" align="center"><span class="teamnamelabel">Captain Information</span></th></tr>
						<tr><td align="right">Name:&nbsp;</td>
							<td nowrap="nowrap"><input type="text" id="captainfirstname" name="captainfirstname" value="<%=sCaptainFirstName%>" size="25" maxlength="25" />
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

				</td>
			</tr>
		</table>

		<p id="teameditsave">
			<input type="button" class="button" value="Save Changes" onclick="validate();" />
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
' void GetTeamInformation iRegattaTeamId 
'--------------------------------------------------------------------------------------------------
Sub GetTeamInformation( ByVal iRegattaTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateam, captainfirstname, captainlastname, captainaddress, captaincity, captainstate, "
	sSql = sSql & "captainzip, captainphone, ISNULL(regattateamgroupid,0) AS regattateamgroupid "
	sSql = sSql & " FROM egov_regattateams WHERE orgid = " & session("orgid") & " AND regattateamid = " & iRegattaTeamId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sTeamName = oRs("regattateam")
		sCaptainFirstname = oRs("captainfirstname")
		sCaptainLastname = oRs("captainlastname")
		sCaptainaddress = oRs("captainaddress")
		sCaptaincity = oRs("captaincity")
		sCaptainstate = oRs("captainstate")
		sCaptainzip = oRs("captainzip")
		sCaptainphone = oRs("captainphone")
		sRegattaTeamGroupName = GetTeamGroupName( oRs("regattateamgroupid") )
		iRegattaTeamGroupId = oRs("regattateamgroupid")
	Else
		sTeamName = ""
		sCaptainname = ""
		sCaptainaddress = ""
		sCaptaincity = ""
		sCaptainstate = ""
		sCaptainzip = ""
		sCaptainphone = ""
		sRegattaTeamGroupName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowTeamMembers iRegattaTeamId 
'--------------------------------------------------------------------------------------------------
Sub ShowTeamMembers( ByVal iRegattaTeamId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT R.regattateammemberid, R.regattateammember, R.isteamcaptain, C.paymentid "
	sSql = sSql & " FROM egov_regattateammembers R, egov_class_list C "
	sSql = sSql & " WHERE R.regattateamid = " & iRegattaTeamId 
	sSql = sSql & " AND R.orgid = " & session("orgid") & " AND C.classlistid = R.classlistid ORDER BY R.regattateammember"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write  vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""regattateamlist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		response.write vbcrlf & "<tr><th>Team Members</th><th>Receipt</th><th>Transfer</th><th>Delete</th></tr>"
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 
			response.write "><td>" & oRs("regattateammember")
			If oRs("isteamcaptain") Then
				response.write " &nbsp; (Captain)"
			End If 
			response.write "</td>"
			response.write "<td align=""center"">"
			'response.write "<a href='view_receipt.asp?iPaymentId=" & oRs("paymentid") & "'>" & oRs("paymentid") & "</a> "
			response.write "<input type=""button"" value=""View Receipt #" & oRs("paymentid") & """ class=""button"" onclick=""location.href='view_receipt.asp?iPaymentId=" & oRs("paymentid") & "'"" />"
			response.write "</td>"
			response.write "<td align=""center"">"
			If oRs("isteamcaptain") Then
				response.write "&nbsp;"
			Else
				response.write "<input type=""button"" value=""Transfer"" class=""button"" onclick=""location.href='regattamembertransfer.asp?regattateammemberid=" & oRs("regattateammemberid") & "&regattateamid=" & iRegattaTeamId & "';"" />"
			End If 
			response.write "</td>"
			response.write "<td align=""center"">"
			If oRs("isteamcaptain") Then
				response.write "&nbsp;"
			Else
				response.write "<input type=""button"" value=""Delete"" class=""button"" onclick=""deleteMember( " & oRs("regattateammemberid") & " );"" />"
			End If 
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 
	Else
		response.write "<p>No members could be found for this team.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>