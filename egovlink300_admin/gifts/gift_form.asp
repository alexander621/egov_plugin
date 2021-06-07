<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sResults, sSearchStart

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "purchase gifts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' these Get us back To this page from the edit user pages - Steve Loar 6/6/06
Session("RedirectPage") = "../gifts/gift_form.asp?userid=" & request("userid") & "&searchname=" & request("searchname") & "&results=" & request("results") & "&g=" & request("G") & "&searchstart=" & request("searchstart")
Session("RedirectLang") = "Return to Gift Purchase"

' GIFT FORM SELECTION
igiftformid = request("G")

If igiftformid = "" Then
	igiftformid = GetFirstGiftId()
End If

sFirstName = ""
sLastName = ""
sAddress1 = ""
sCity = ""
sState = ""
sZip = ""
sEmail = ""
sHomePhone = ""
sWorkPhone = ""
sFax = ""
sBusinessName = ""

If request("userid") <> "" Then
	iUserId = request("userid")
Else
	iUserId = GetFirstUserId()
End If

SetUserInformation( iUserid )

' See if a search term was passed
If request("searchname") <> "" Then 
	sSearchName = request("searchname")
Else
	sSearchName = ""
End If 

If request("results") <> "" Then
	sResults = request("results")
Else
	sResults = ""
End If 

If request("searchstart") <> "" Then
	sSearchStart = request("searchstart")
Else
	sSearchStart = -1
End If

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/reservation.css" />

	<script language="Javascript" src="../scripts/easyform.js"></script>

	<script language="Javascript">
	<!--

	function SearchCitizens( iSearchStart )
	{
		var optiontext;
		var optionchanged;
		//alert(document.frmpayment.searchname.value);
		var searchtext = document.frmpayment.searchname.value;
		var searchchanged = searchtext.toLowerCase();

		iSearchStart = parseInt(iSearchStart) + 1;

		for (x=iSearchStart; x < document.frmpayment.userid.length ; x++)
		{
			//alert(document.frmpayment.userid.options[x].text);
			optiontext = document.frmpayment.userid.options[x].text;
			optionchanged = optiontext.toLowerCase();
			if (optionchanged.indexOf(searchchanged) != -1)
			{
				document.frmpayment.userid.selectedIndex = x;
				document.frmpayment.results.value = 'Possible Match Found.';
				document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
				document.frmpayment.searchstart.value = x;
				location.href='gift_form.asp?userid=' + document.frmpayment.userid.options[document.frmpayment.userid.selectedIndex].value + '&searchname=' + document.frmpayment.searchname.value + '&results=' + document.frmpayment.results.value + '&g=' + document.frmpayment.gift.options[document.frmpayment.gift.selectedIndex].value + '&searchstart=' + document.frmpayment.searchstart.value;
				return;
			}
		}
		document.frmpayment.results.value = 'No Match Found.';
		document.getElementById('searchresults').innerHTML = 'No Match Found.';
		document.frmpayment.searchstart.value = -1;
	}

	function ClearSearch()
	{
		document.frmpayment.searchstart.value = -1;
	}

	function UserPick()
	{
		document.frmpayment.searchname.value = '';
		document.frmpayment.results.value = '';
		document.getElementById('searchresults').innerHTML = ''; 
		document.frmpayment.searchstart.value = -1;
		location.href='gift_form.asp?userid=' + document.frmpayment.userid.options[document.frmpayment.userid.selectedIndex].value + '&searchname=' + document.frmpayment.searchname.value + '&results=' + document.frmpayment.results.value + '&g=' + document.frmpayment.gift.options[document.frmpayment.gift.selectedIndex].value + '&searchstart=' + document.frmpayment.searchstart.value;
	}

	function getinfo()
	{
		if (document.frmpayment.chkSameAs.checked) 
		{
			// CHECK USE ABOVE
			document.frmpayment.txtack_name.value = document.frmpayment.txtfirstname.value + ' '  + document.frmpayment.txtlastname.value;
			document.frmpayment.txtack_address1.value = document.frmpayment.txthome_address1.value;
			document.frmpayment.txtack_city.value = document.frmpayment.txthome_city.value;
			document.frmpayment.txtack_state.value = document.frmpayment.cbohome_state.value;
			document.frmpayment.txtAcknoledgeZip.value = document.frmpayment.txthome_zip.value;
		}
		else 
		{
			// UNCHECKED CLEAR VALUES
			document.frmpayment.txtack_name.value = '';
			document.frmpayment.txtack_address1.value = '';
			document.frmpayment.txtack_city.value = '';
			document.frmpayment.txtack_state.value = '';
			document.frmpayment.txtAcknoledgeZip.value = '';
		}

	}
//-->
</script>

</head>


<body>

 <%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<form  name="frmpayment" action="facility_cashcheck_receipt.asp" method="post">

<!--BEGIN PAGE CONTENT-->
<div style="padding:20px;">


<!--BEGIN: GIFT SELECTION-->
<div class="reserveformtitle">Commemorative Gift Selection</div>
<div class="reserveforminputarea">
	<% DrawGiftSelection( igiftformid ) %>
</div>
<!--END: GIFT SELECTION-->


<!--BEGIN: USER MENU-->
<div class="reserveformtitle">Contact Information</div>
<div class="reserveforminputarea">
	<P><font class="reserveforminstructions"><b>Instructions:</b> 
		Select a registed Citizen from the drop down list.  If their name is not on the list then select New User. 
		Or, select view/edit to review their current contact information.</font>
	</p>
	<p>
		Name Search: <input type="text" name="searchname" value="<%=sSearchName%>" size="25" maxlength="50" onchange="javascript:ClearSearch();" />
		<input type="button" class="button" value="Search" onclick="javascript:SearchCitizens(document.frmpayment.searchstart.value);" />
		<input type="hidden" name="results" value="" />
		<input type="hidden" name="searchstart" value="<%=sSearchStart%>" />
		<span id="searchresults"> <%=sResults%></span>
		<br /><div id="searchtip">(last name, first name)</div>
	</p>
	<P><font class=reservationformlabel>Select Name: </font> &nbsp; 
		<select name="userid" onchange="javascript:UserPick();"> <!--<%'=request.servervariables("QUERY_STRING")%>-->
			<%=ShowUserDropDown(iUserId)%>
		</select> &nbsp; <input class=reserveformbuttonsmall type=button value="Edit/View" onClick="location.href='../dirs/update_citizen.asp?userid=' + document.frmpayment.userid.options[document.frmpayment.userid.selectedIndex].value;"> &nbsp; <input onClick="location.href='../dirs/register_citizen.asp';" class=reserveformbuttonsmall type=button value="New User">
	</p>

</div>
<!--END: USER MENU-->


<!--BEGIN PAGE CONTENT-->
<P>
<div class=reserveformtitle>
<% 
Dim sGiftName, sGiftAmount
GetGiftInformation(igiftformid) 
response.write sGiftName
%>
</div>
</P>


<!--BEGIN:  VERISIGN PAYMENT FORM DETAILS-->
<%
response.write "<input type=hidden name=""GIFTID"" value=""" & igiftformid & """>"
response.write "<input type=hidden name=""ITEM_NUMBER"" value=""G" & igiftformid & "00"">"
response.write "<input type=hidden name=""ITEM_NAME"" value=""" & sGiftName & """>"
response.write "<input type=hidden name=""amount"" value=""" & sGiftAmount & """>"
response.write "<input type=hidden name=""iPAYMENT_MODULE"" value=""1"">"
response.write "<input type=""hidden"" name=""iuserid"" value=""" & iuserid &""" />"
%>
<input type=hidden name="ef:iuserid-text/req" value="Logging In or Registering ">


<!--PERSON(S) INITIATING GIFT-->
<% DrawSection1() %>


<!-- ACKNOWLEDGEMENT -->
<div class=reserveformtitle>Acknowledgement  </div>
<div class=reserveforminputarea>
<P><% DrawSection2() %></p>
</div>


<!-- GIFT SPECIFIC INFORMATION -->
<p><% DrawGiftOptions(igiftformid)%></p>
<!--END: PAGE CONTENT-->



<!--BEGIN: MAKE PURCHASE-->
<div class=reserveformtitle>Make Purchase</div>
<div class=reserveforminputarea>
	<table border="0" cellpadding="5" cellspacing="0" >
			<tr><td> 
			Payment Type: <select name="paymenttype" size="1">
				<option value="1">CreditCard</option>
				<option value="2">Check</option>
				<option value="3">Cash</option>
			</select>
			</td>
			<td>
				Payment Location: <select name="paymentlocation" size="1">
					<option value="1">Walk In</option>
					<option value="2">Phone Call</option>
				</select>
			</td>
			<td width="200">
					<input type="button" class="reserveformbutton" style="width:200px;text-align:center;" name="continue" value="Continue with Purchase" onclick="if (validateForm('frmpayment')) {document.frmpayment.submit()}" />
			</td></tr>
	</table>
</div>
<!--END: MAKE PURCHASE-->

</form>

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
' FUNCTION SHOWUSERDROPDOWN(IUSERID)
'--------------------------------------------------------------------------------------------------
Function ShowUserDropDown( iUserId )
	Dim sSQL

	sSQL = "Select userid, userfname, userlname, useraddress FROM egov_users where orgid = " & Session("OrgID") & " and userregistered = 1 and useremail is not NULL order by userlname, userfname, userid"

	Set oResident = Server.CreateObject("ADODB.Recordset")
	oResident.Open sSQL, Application("DSN"), 0, 1

	Do While Not oResident.eof 
		response.write vbcrlf & "<option value=""" & oResident("userid") & """" 
		If CLng(iUserId) = CLng(oResident("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oResident("userlname") & ", " & oResident("userfname") & " &ndash; " & oResident("useraddress") & "</option>"
		oResident.movenext
	Loop 
		
	oResident.close
	Set oResident = Nothing
End Function 




'--------------------------------------------------------------------------------------------------
' SUB DRAWSECTION1()
'--------------------------------------------------------------------------------------------------
Sub DrawSection1()
%>
<input type=hidden name="txtfirstname" maxLength="15"  style="font-family:Arial; font-size:8pt; width:80px" title="Enter first name here" onFocus="this.select()"  value="<%=sFirstName%>" ID="txtfirstname">
<input name="txtMI" type=hidden maxLength="1" style="font-family:Arial; font-size:8pt; width:20px"  value="" title="Enter middle initial here" onFocus="this.select()" ID="txtMI">
<input type=hidden name="txtlastname" maxLength="15" style="font-family:Arial; font-size:8pt; width:110px" title="Enter last name here" onFocus="this.select()" ID="txtlastname" value="<%=sLastName%>">
<INPUT type="hidden" title="Enter your home street address here.  Please spell out words such as Street, Road, Avenue, etc. in the Address line." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=sAddress1%>" maxLength="50" name="txthome_address1" ID="txthome_address1">
<INPUT type="hidden" title="Enter additional home address information here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txthome_address2" ID="txthome_address2"  value="<%=sAddress2%>" >
<INPUT title="Enter home city here" type="hidden" style="FONT-SIZE: 8pt; WIDTH: 150px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="25" size="25" name="txthome_city" ID="txthome_city"   value="<%=sCity%>" >
<input type=hidden name="cbohome_state" ID="cbohome_state" value="<%=sState%>">
<INPUT type="hidden" title="Enter home postal code here" style="FONT-SIZE: 8pt; WIDTH: 100px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="10" name="txthome_zip" ID="txthome_zip" value="<%=sZip%>">
<INPUT type="hidden" title="Enter work area code here" style="FONT-SIZE: 8pt; WIDTH: 25px; FONT-FAMILY: Arial"  value="<%=Left(sHomePhone,3)%>" onfocus="this.select()" maxLength="3" name="txtPhone1" ID="txtPhone1">
<INPUT title="Enter work exchange here" type="hidden" style="FONT-SIZE: 8pt; WIDTH: 25px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=Mid(sHomePhone,4,3)%>" maxLength="3" name="txtPhone2" ID="txtPhone2">
<INPUT title="Enter work number here" type="hidden" style="FONT-SIZE: 8pt; WIDTH: 35px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=Mid(sHomePhone,7,4)%>" maxLength="4" name="txtPhone3" ID="txtPhone3">
<INPUT type="hidden" title="Enter your email address here." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50"  value="<%=sEmail%>" name="txtEmail" ID="txtEmail">
<!--BEGIN: VALIDATION-->
<!--input type=hidden name="ef:txtfirstname-text/req" value="First Name">
<input type=hidden name="ef:txtlastname-text/req" value="Last Name">
<input type=hidden name="ef:txthome_address1-text/req" value="Home Address Line 1">
<input type=hidden name="ef:txthome_city-text/req" value="City">
<input type=hidden name="ef:cbohome_state-text/req" value="State">
<input type=hidden name="ef:txthome_zip-text/req" value="Zipcode">
<input type=hidden name="ef:txtEmail-text/req" value="Email"-->
<!--END: VALIDATION-->
							<%
End Sub


'--------------------------------------------------------------------------------------------------
' SUB DRAWSECTION2()
'--------------------------------------------------------------------------------------------------
Sub DrawSection2()
%>
<!-- Start Acknowledgement -->
<TABLE cellSpacing="0" cellPadding="0" width="625"  border="0" ID="Table1">

							<TR bgColor="#e0e0e0"> <!-------------------------> <!-- Sign IN ID --> <!------------------------->
								<TD align="right">
									<INPUT onClick="getinfo();" title="Same as above" type="checkbox" maxLength="30" name="chkSameAs" value="TRUE" ID="chkSameAs">
								</TD>
								<TD style="font-family:Arial; font-size:8pt; color:#000000" align="left">Address is the same as the citizen's</TD>
							</TR>
							<TR bgColor="#e0e0e0">
								<TD style="font-family:Arial; font-size:8pt; color:#000000" align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Name(s):</TD>
								<TD style="font-family:Arial; font-size:8pt; color:#000000" align="left">
									<INPUT type="text" title="Name Here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="15" name="txtack_name" ID="txtack_name"></TD>
							</TR> <!-------------------------> <!-- HOME ADDRESS LINE 1 --> <!------------------------->
							<TR bgColor="#e0e0e0">
								<TD style="font-family:Arial; font-size:8pt; color:#000000" vAlign="center" align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Address:</TD>
								<TD align="left"><INPUT type="text" title="Enter your home street address here.  Please spell out words such as Street, Road, Avenue, etc. in the Address line." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txtack_address1" ID="txtack_address1"></TD>
							</TR> <!-------------------------> <!-- HOME ADDRESS LINE 2 --> <!------------------------->
							<TR bgColor="#e0e0e0">
								<TD style="font-family:Arial; font-size:8pt; color:#000000" vAlign="center" align="right">&nbsp;</TD>
								<TD align="left"><INPUT type="text" title="Enter additional home address information here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txtack_address2" ID="txtack_address2"></TD>
							</TR> <!---------------------------> <!-- HOME CITY, STATE, ZIP --> <!--------------------------->
							<TR bgColor="#e0e0e0"> <!-- CITY -->
								<TD style="font-family:Arial; font-size:8pt; color:#000000" vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>City:</TD>
								<TD style="font-family:Arial; font-size:8pt; color:#000000" noWrap>
									<INPUT title="Enter home city here" type="text" style="FONT-SIZE: 8pt; WIDTH: 150px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="25" size="25" name="txtack_city" ID="txtack_city">&nbsp;  State:<!--STATE-->
									<SELECT style="FONT-SIZE: 8pt; WIDTH: 50px; FONT-FAMILY: Arial" name="txtack_state" ID="txtack_state">
										<!--Loop through each state-->
										
										<OPTION value='AK'>AK</OPTION>
										
										<OPTION value='AL'>AL</OPTION>
										
										<OPTION value='AR'>AR</OPTION>
										
										<OPTION value='AZ'>AZ</OPTION>
										
										<OPTION value='CA'>CA</OPTION>
										
										<OPTION value='CO'>CO</OPTION>
										
										<OPTION value='CT'>CT</OPTION>
										
										<OPTION value='DE'>DE</OPTION>
										
										<OPTION value='FL'>FL</OPTION>
										
										<OPTION value='GA'>GA</OPTION>
										
										<OPTION value='HI'>HI</OPTION>
										
										<OPTION value='IA'>IA</OPTION>
										
										<OPTION value='ID'>ID</OPTION>
										
										<OPTION value='IL'>IL</OPTION>
										
										<OPTION value='IN'>IN</OPTION>
										
										<OPTION value='KS'>KS</OPTION>
										
										<OPTION value='KY'>KY</OPTION>
										
										<OPTION value='LA'>LA</OPTION>
										
										<OPTION value='MA'>MA</OPTION>
										
										<OPTION value='MD'>MD</OPTION>
										
										<OPTION value='ME'>ME</OPTION>
										
										<OPTION value='MI'>MI</OPTION>
										
										<OPTION value='MN'>MN</OPTION>
										
										<OPTION value='MO'>MO</OPTION>
										
										<OPTION value='MS'>MS</OPTION>
										
										<OPTION value='MT'>MT</OPTION>
										
										<OPTION value='NC'>NC</OPTION>
										
										<OPTION value='ND'>ND</OPTION>
										
										<OPTION value='NE'>NE</OPTION>
										
										<OPTION value='NH'>NH</OPTION>
										
										<OPTION value='NJ'>NJ</OPTION>
										
										<OPTION value='NM'>NM</OPTION>
										
										<OPTION value='NV'>NV</OPTION>
										
										<OPTION value='NY'>NY</OPTION>
										
										<OPTION <%If UCASE(LEFT(sState,2)) = "OH" Then response.write "SELECTED" End If%> value='OH'>OH</OPTION>
										
										<OPTION value='OK'>OK</OPTION>
										
										<OPTION value='OR'>OR</OPTION>
										
										<OPTION value='PA'>PA</OPTION>
										
										<OPTION value='RI'>RI</OPTION>
										
										<OPTION value='RI'>RI</OPTION>
										
										<OPTION value='SC'>SC</OPTION>
										
										<OPTION value='SD'>SD</OPTION>
										
										<OPTION value='TN'>TN</OPTION>
										
										<OPTION value='TX'>TX</OPTION>
										
										<OPTION value='UT'>UT</OPTION>
										
										<OPTION value='VA'>VA</OPTION>
										
										<OPTION value='VT'>VT</OPTION>
										
										<OPTION value='WA'>WA</OPTION>
										
										<OPTION value='WI'>WI</OPTION>
										
										<OPTION value='WV'>WV</OPTION>
										
										<OPTION value='WY'>WY</OPTION>
										
									</SELECT>
								</TD>
							</TR>
							<TR bgColor="#e0e0e0"> <!-- ZIP -->
								<TD style="font-family:Arial; font-size:8pt; color:#000000" vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Zip/Postal 
									Code:</TD>
								<TD style="font-family:Arial; font-size:8pt; color:#000000" noWrap><INPUT type="text" title="Enter home postal code here" style="FONT-SIZE: 8pt; WIDTH: 100px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="10" name="txtAcknoledgeZip" ID="txtAcknoledgeZip">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								</TD>
							</TR>
							</table>
							<!--BEGIN: VALIDATION-->
<input type=hidden name="ef:txtack_name-text/req" value="Acknowledgement Name">
<input type=hidden name="ef:txtack_address1-text/req" value="Acknowledgement Address Line 1">
<input type=hidden name="ef:txtack_city-text/req" value="Acknowledgement City">
<input type=hidden name="ef:txtack_state-text/req" value="Acknowledgement State">
<input type=hidden name="ef:txtAcknoledgeZip-text/req" value="Acknowledgement Zipcode">
<!--END: VALIDATION-->
<%
End Sub


'--------------------------------------------------------------------------------------------------
' SUB DRAWGIFTOPTIONS(IGIFTID)
'--------------------------------------------------------------------------------------------------
Sub DrawGiftOptions(iGiftID)

	sSQL = "SELECT  * FROM egov_gift_group where orgid='" & session("orgid") & "' and giftid='" & iGiftID & "'"

	Set oGiftGroups = Server.CreateObject("ADODB.Recordset")
	oGiftGroups.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oGiftGroups.EOF Then
		
		Do While NOT oGiftGroups.EOF
			' DRAW PAYMENT FIELD GROUP TITLE
			response.write "<div class=reserveformtitle>" & oGiftGroups("giftgroupname") & "</div>"
			response.write "<div class=reserveforminputarea>"
			DrawGiftFields igiftid,oGiftGroups("giftgroupid")
			response.write "</div>"
			oGiftGroups.MoveNext
		Loop

	End If

	Set oGiftGroups = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DRAWGIFTOPTIONS(IGIFTID)
'--------------------------------------------------------------------------------------------------
Sub DrawGiftFields(iGiftID,iGroupID)

	sSQL = "SELECT * FROM egov_gift_fields where giftid='" & iGiftID & "' and groupid='" & iGroupID & "'"

	Set oGiftFields = Server.CreateObject("ADODB.Recordset")
	oGiftFields.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oGiftFields.EOF Then
		
		response.write "<P>"
		response.write "<TABLE cellSpacing=""0"" cellPadding=""0"" width=""625""  border=""0"" ID=""Table1"">"

		Do While NOT oGiftFields.EOF

			' CHECK IF IT IS REQUIRED
			If  oGiftFields("isrequired") Then
				sRequired = "<SPAN style=""COLOR: #ff0000"">*&nbsp;</SPAN>"
				' ERROR CHECKING FIELD
				response.write "<input type=hidden name=""ef:custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & "-text/" & oGiftFields("validation") & "/req"" value=""" & oGiftFields("fieldprompt") & """>"
			Else
				sRequired = " "
			End If

			' SET HEIGHT FOR INPUT BOX BASED ON FIELD TYPE, 1=STANDARD, 2=SIMULATED TEXT AREA
			If  oGiftFields("fieldtype") = 2 Then
				sHeight = "HEIGHT: 100px;"
			Else
				sHeight = " " 
			End If
		
			response.write "<TR bgColor=""#e0e0e0"">"
			response.write "<TD vAlign=""top"" style=""font-family:Arial; font-size:8pt; color:#000000"" align=""right"">" & sRequired
			response.write oGiftFields("fieldprompt")
			response.write ": </TD>"

			response.write "<TD style=""font-family:Arial; font-size:8pt; color:#000000"" align=""left"">"
			
			Select Case oGiftFields("fieldtype")

				Case "1"
					' TEXT BOX
					response.write "<INPUT name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ type=""text"" style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" >"
				Case "2"
					' PSEUDO TEXT AREA
					response.write "<INPUT name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ type=""text"" style=""FONT-SIZE: 8pt; WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" >"
				Case "3"
					' SELECT BOX
					arrAnswers = split(oGiftFields("fieldchoices"),"@@")
			
					response.write "<select name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ >"
					For alist = 0 to ubound(arrAnswers)
						response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
					Next
					response.write "</select>"

				Case Else
					' UNKNOWN TYPE DONT PROCESS
					response.write "INPUT TYPE ERROR. PLEASE CONSULT SETUP."
			
			End Select

			response.write "</TD></TR>"

			oGiftFields.MoveNext
		Loop

		response.write "</TABLE>"
		response.write "</P>"

	End If

	Set oGiftFields = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' GETGIFTINFORMATION(IGIFTID)
'--------------------------------------------------------------------------------------------------
Sub GetGiftInformation(iGiftID)

	sSQL = "SELECT giftname,amount from egov_gift where giftid='" & iGiftID & "'"

	Set oGift = Server.CreateObject("ADODB.Recordset")
	oGift.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oGift.EOF Then
		sGiftName = oGift("giftname")
		sGiftAmount = oGift("Amount")
	Else
		sGiftName = "Unknown Gift"
		sGiftAmount = "0.00"
	End If
	
	Set oGift = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SETUSERINFORMATION()
'--------------------------------------------------------------------------------------------------
Sub SetUserInformation( iUserid )
	If iUserid <> "" and iUserid <> "-1" Then
	'response.write "HERE3"
		
		'iUserID = request.cookies("userid")
	
		sSQL = "SELECT * FROM egov_users WHERE userid=" & iUserID
		Set oInfo = Server.CreateObject("ADODB.Recordset")
		oInfo.Open sSQL, Application("DSN"), 3, 1

		If NOT oInfo.EOF Then
			'response.write "HERE"
			' USER FOUND SET VALUES
			sFirstName = oInfo("userfname")
			sLastName = oInfo("userlname")
			sAddress1 = oInfo("useraddress")
			sCity = oInfo("usercity")
			sState = oInfo("userstate")
			sZip = oInfo("userzip")
			sEmail = oInfo("useremail")
			sHomePhone = oInfo("userhomephone")
			sWorkPhone = oInfo("userworkphone")
			sBusinessName = oInfo("userbusinessname")
			sFax = oInfo("userfax")

		Else
			' USER NOT FOUND SET VALUES TO EMPTY
			sFirstName = ""
			sLastName = ""
			sAddress1 = ""
			sCity = ""
			sState = ""
			sZip = ""
			sEmail = ""
			sHomePhone = ""
			sWorkPhone = ""
			sFax = ""
			sBusinessName = ""

		End If

		Set oInfo = Nothing

	End If
End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION DRAWGIFTSELECTION(IGIFTFORMID)
'--------------------------------------------------------------------------------------------------
Function DrawGiftSelection(igiftformid)

	' GET LIST OF GIFTS
	sSQL = "SELECT * FROM egov_gift WHERE orgid='" & Session("OrgID") & "'" & " Order by giftname"

	' OPEN RECORDSET
	Set oGift = Server.CreateObject("ADODB.Recordset")
	oGift.Open sSQL, Application("DSN"), 3, 1
	

	' IF NOT RECORDSET NOT EMPTY DISPLAY SELECT BOX
	If NOT oGift.EOF Then
	
		' BEGIN SELECT BOX
		response.write "<select name=""gift"" onChange=""location.href='gift_form.asp?g=' + document.frmpayment.gift.options[document.frmpayment.gift.selectedIndex].value;"" >"

		' LIST OPTIONS
		Do While Not oGift.eof 
			
			response.write vbcrlf & "<option value=""" & oGift("giftid") & """"
			
			' IF SELECTED MARK AS SELECTED OPTION
			If clng(igiftformid) = clng(oGift("giftid")) Then
				response.write " selected=""selected"" "
			End If 

			response.write ">" & oGift("giftname") & "</option>"

			oGift.movenext
		Loop 

		' END SELECT BOX
		response.write "</select>"

	End If 


	' CLEAN UP OBJECTS
	oGift.close
	Set oGift  = Nothing

End Function


Function GetFirstUserId()
	Dim sSQl

	sSQL = "SELECT  TOP 1 userid FROM egov_users WHERE orgid = " & Session("OrgID") 
	sSQL = sSQL & " ORDER BY userlname, userfname, userid"
	
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN") , 3, 1

	GetFirstUserId = oUser("userid")

	oUser.close
	Set oUser = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetFirstGiftId()
'--------------------------------------------------------------------------------------------------
Function GetFirstGiftId()
	Dim sSql, oGifts

	sSQL = "SELECT giftid FROM egov_gift WHERE orgid = " & Session("OrgID") & " Order by giftname"
	Set oGifts = Server.CreateObject("ADODB.Recordset")
	oGifts.Open sSQL, Application("DSN"), 3, 1

	If Not oGifts.EOF Then 
		GetFirstGiftId = oGifts("giftid")
	Else
		GetFirstGiftId = 0
	End If 

	oGifts.close
	Set oGifts = Nothing
End Function 

%>


