<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: GIFT_FORM.ASP
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

' CAPTURE CURRENT PATH - USED FOR LOGIN REDIRECTS
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Gift Purchase"
Dim sFirstName,sLastName,sAddress1,sCity,sState,sZip,sEmail,sHomePhone,sWorkPhone,sFax,sBusinessName
SetUserInformation()

Dim sGiftName, sGiftAmount

' Handle SQL Intrusions gracefully
If Not IsNumeric(request("G")) Then 
	response.redirect "gift_list.asp"
End If 

GetGiftInformation CLng(request("G"))

%>

<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>


	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/easyform.js"></script>

	<script language="Javascript">
	<!--

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

<!--#Include file="../include_top.asp"-->

<!--BEGIN:  USER REGISTRATION - USER MENU-->
<% If sOrgRegistration Then %>
		<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then %>
				<%	RegisteredUserDisplay( "../" ) %>
			<%	Else 
					' Added this to make the login work like classes and events - Steve Loar 5/19/2006
					session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
					response.redirect "../user_login.asp"
			%>
					<!--REGISTRATION OR LOGIN-->
					<!--<P><div class=reserveformtitle>Contact Information</div>
					<div class=reserveforminputarea>
					<P><font class=reserveforminstructions>You need to sign in or register now to complete your purchase. As a result of changes to the City of Montgomery website you will need to re-register.  We apologize for the inconvenience, however this is a one-time process and once you register, you will not need to register again.</font></p>
					<P>
					<input onClick="location.href='../user_login.asp';" value="Login" class=reserveformbutton style="width:75px;text-align:center;" type=button>  or <input value="Register Now!" class=reserveformbutton style="width:150px;text-align:center;" type=button onClick="location.href='../register.asp';">
					</p>
					</div></p> -->
		<% End If %>
<% Else %>
	<!--REGISTRATION OR LOGIN-->
	<p><br />
		<div class="reserveformtitle">Contact Information</div>
		<div class="reserveforminputarea">
			<p>
				<font class="reserveforminstructions">You need to sign in or register now to complete your purchase. As a result of changes to the City of Montgomery website you will need to re-register.  We apologize for the inconvenience, however this is a one-time process and once you register, you will not need to register again.</font>
			</p>
			<p>
				<input onClick="location.href='../user_login.asp';" value="Login" class="reserveformbutton" style="width:75px;text-align:center;" type="button" />  or <input value="Register Now!" class="reserveformbutton" style="width:150px;text-align:center;" type=button onClick="location.href='../register.asp';">
			</p>
		</div>
	</p>
<% End If%>
<!--END:  USER REGISTRATION - USER MENU-->


<!--BEGIN PAGE CONTENT-->
<p>
	<div class="reserveformtitle">
	<% 
		response.write sGiftName
	%>
	</div>
</p>


<!--BEGIN:  VERISIGN PAYMENT FORM DETAILS-->
<%
response.write  "<FORM  name=""frmpayment"" ACTION=""" &  Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/recreation_payments/VERISIGN_FORM.ASP"" METHOD=""POST"">"


response.write "<input type=hidden name=""GIFTID"" value=""" & request("G") & """>"
response.write "<input type=hidden name=""ITEM_NUMBER"" value=""G" & request("G") & "00"">"
response.write "<input type=hidden name=""ITEM_NAME"" value=""" & sGiftName & """>"
if request("G") = "5" then%>
<div class=reserveformtitle>Gift </div>
<div class=reserveforminputarea>
	Donation Amount: <input type=text name="amount" value="<%=sGiftAmount%>" />
	<input type=hidden name="ef:amount-text/number/req" value="Donation Amount ">
</div>
<%
else
	response.write "<input type=hidden name=""amount"" value=""" & sGiftAmount & """>"
end if
response.write "<input type=hidden name=""iPAYMENT_MODULE"" value=""1"">"
response.write "<input type=""hidden"" name=""iuserid"" value=""" & request.cookies("userid") &""" />"
%>
<input type=hidden name="ef:iuserid-text/req" value="Logging In or Registering ">


<!--PERSON(S) INITIATING GIFT-->
<div class=reserveformtitle>Person(s) Initiating Gift </div>
<div class="reserveforminputarea giftfields">
<P><% DrawSection1 %></p>
</div>


<!-- ACKNOWLEDGEMENT -->
<div class="reserveformtitle">Acknowledgement</div>
<div class="reserveforminputarea giftfields">
	<p>
		<% DrawSection2 %>
	</p>
</DIV>


<!-- GIFT SPECIFIC INFORMATION -->
<P><% DrawGiftOptions CLng(request("G")) %></P>
<!--END: PAGE CONTENT-->


<!--MAKE PURCHASE-->
<div class="reserveformtitle">Continue to Payment Information</div>
<div class="reserveforminputarea giftfields">
	<center>
		<input value="Click Here To Continue To Payment Information" class="actionbtn" type="button" onclick="if (validateForm('frmpayment')) {document.frmpayment.submit();}" >
	</center>
</div>
</form>

<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB DRAWSECTION1()
'--------------------------------------------------------------------------------------------------
Sub DrawSection1()
%>
	<TABLE cellSpacing="0" cellPadding="0" width="625"  border="0" ID="Table1">

		<TR>
			<TD align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>First Name:</TD>
			<td align="left"><input type=text name="txtfirstname" maxLength="15"  style="font-family:Arial; font-size:8pt; width:80px" title="Enter first name here" onFocus="this.select()"  value="<%=sFirstName%>" ID="txtfirstname">&nbsp;&nbsp;MI:<input name="txtMI" type=text maxLength="1" style="font-family:Arial; font-size:8pt; width:20px"  value="" title="Enter middle initial here" onFocus="this.select()" ID="txtMI">&nbsp;&nbsp;
				<span style="color:#ff0000">*&nbsp;</span>Last Name:<input type=text name="txtlastname" maxLength="15" style="font-family:Arial; font-size:8pt; width:110px" title="Enter last name here" onFocus="this.select()" ID="txtlastname" value="<%=sLastName%>">
			</td>

		</TR> <!-------------------------> <!-- HOME ADDRESS LINE 1 --> <!------------------------->
		<TR>
			<TD vAlign="center" align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Address:</TD>
			<TD align="left"><INPUT type="text" title="Enter your home street address here.  Please spell out words such as Street, Road, Avenue, etc. in the Address line." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=sAddress1%>" maxLength="50" name="txthome_address1" ID="txthome_address1"></TD>
		</TR> <!-------------------------> <!-- HOME ADDRESS LINE 2 --> <!------------------------->
		<TR>
			<TD vAlign="center" align="right">&nbsp;</TD>
			<TD align="left"><INPUT type="text" title="Enter additional home address information here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txthome_address2" ID="txthome_address2"  value="<%=sAddress2%>" ></TD>
		</TR> <!---------------------------> <!-- HOME CITY, STATE, ZIP --> <!--------------------------->
		<TR  <!-- CITY -->
			<TD vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>City:</TD>
			<TD noWrap><INPUT title="Enter home city here" type="text" style="FONT-SIZE: 8pt; WIDTH: 150px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="25" size="25" name="txthome_city" ID="txthome_city"   value="<%=sCity%>" >&nbsp; 
				<span class="respCol">State:</span> <span class="respCol"><!-- STATE --><SELECT style="FONT-SIZE: 8pt; WIDTH: 50px; FONT-FAMILY: Arial" name="cbohome_state" ID="cbohome_state">
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
						
						<OPTION  <%If UCASE(LEFT(sState,2)) = "OH" Then response.write "SELECTED" End If%> value='OH'>OH</OPTION>
						
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
						
				</SELECT></span>
			</TD>
		</TR>
		<TR  <!-- ZIP & Phone -->
			<TD vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Zip/Postal 
				Code:</TD>
			<TD noWrap><INPUT type="text" title="Enter home postal code here" style="FONT-SIZE: 8pt; WIDTH: 100px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="10" name="txthome_zip" ID="txthome_zip" value="<%=sZip%>"> 
				<span class="respCol">Phone: </span>
				<span class="respCol"><input class="phonenum" type="text" title="Enter work area code here" size="3" style="FONT-SIZE: 8pt; WIDTH: 25px; FONT-FAMILY: Arial"  value="<%=Left(sHomePhone,3)%>" onfocus="this.select()" maxLength="3" name="txtPhone1" ID="txtPhone1" />-
				<input class="phonenum" title="Enter work exchange here" type="text" size="3" style="FONT-SIZE: 8pt; WIDTH: 25px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=Mid(sHomePhone,4,3)%>" maxLength="3" name="txtPhone2" ID="txtPhone2" />-
				<input class="phonenum" title="Enter work number here" type="text" size="3" style="FONT-SIZE: 8pt; WIDTH: 35px; FONT-FAMILY: Arial" onfocus="this.select()"  value="<%=Mid(sHomePhone,7,4)%>" maxLength="4" name="txtPhone3" ID="txtPhone3" /></span></TD>
		</TR> <!---------------> <!-- separator --> <!--------------->
		<TR>
			<TD vAlign="center" align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Email:</TD>
			<TD align="left"><INPUT type="text" title="Enter your email address here." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50"  value="<%=sEmail%>" name="txtEmail" ID="txtEmail"></TD>
		</TR> <!---------------------------> <!-- HOME CITY, STATE --> <!--------------------------->
		<!-- End Persons Initiating Gift -->
	</table>

	<!--BEGIN: VALIDATION-->
	<input type=hidden name="ef:txtfirstname-text/req" value="First Name">
	<input type=hidden name="ef:txtlastname-text/req" value="Last Name">
	<input type=hidden name="ef:txthome_address1-text/req" value="Home Address Line 1">
	<input type=hidden name="ef:txthome_city-text/req" value="City">
	<input type=hidden name="ef:cbohome_state-text/req" value="State">
	<input type=hidden name="ef:txthome_zip-text/req" value="Zipcode">
	<input type=hidden name="ef:txtEmail-text/req" value="Email">
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

	<TR> <!-------------------------> <!-- Sign IN ID --> <!------------------------->
		<TD align="right">
			<INPUT onClick="getinfo();" title="Same as above" type="checkbox" maxLength="30" name="chkSameAs" value="TRUE" ID="chkSameAs">
		</TD>
		<TD align="left">Address is the same as above</TD>
	</TR>
	<TR>
		<TD align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Name(s):</TD>
		<TD align="left">
			<INPUT type="text" title="Name Here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="15" name="txtack_name" ID="txtack_name"></TD>
	</TR> <!-------------------------> <!-- HOME ADDRESS LINE 1 --> <!------------------------->
	<TR>
		<TD vAlign="center" align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Address:</TD>
		<TD align="left"><INPUT type="text" title="Enter your home street address here.  Please spell out words such as Street, Road, Avenue, etc. in the Address line." style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txtack_address1" ID="txtack_address1"></TD>
	</TR> <!-------------------------> <!-- HOME ADDRESS LINE 2 --> <!------------------------->
	<TR>
		<TD vAlign="center" align="right">&nbsp;</TD>
		<TD align="left"><INPUT type="text" title="Enter additional home address information here" style="FONT-SIZE: 8pt; WIDTH: 300px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="50" name="txtack_address2" ID="txtack_address2"></TD>
	</TR> <!---------------------------> <!-- HOME CITY, STATE, ZIP --> <!--------------------------->
	<TR> <!-- CITY -->
		<TD vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>City:</TD>
		<TD noWrap>
			<INPUT title="Enter home city here" type="text" style="FONT-SIZE: 8pt; WIDTH: 150px; FONT-FAMILY: Arial" onfocus="this.select()" maxLength="25" size="25" name="txtack_city" ID="txtack_city">&nbsp;  
			<span class="respCol">State:<!--STATE--></span>
			
			<span class="respCol"><SELECT style="FONT-SIZE: 8pt; WIDTH: 50px; FONT-FAMILY: Arial" name="txtack_state" ID="txtack_state">
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
				
			</SELECT></span>
		</TD>
	</TR>
	<TR> <!-- ZIP -->
		<TD vAlign="center" noWrap align="right"><SPAN style="COLOR: #ff0000">*&nbsp;</SPAN>Zip/Postal 
			Code:</TD>
		<TD noWrap><INPUT type="text" title="Enter home postal code here" onfocus="this.select()" maxLength="10" name="txtAcknoledgeZip" ID="txtAcknoledgeZip">
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

	sSQL = "SELECT  * FROM egov_gift_group where orgid='" & iorgid & "' and giftid='" & iGiftID & "'"

	Set oGiftGroups = Server.CreateObject("ADODB.Recordset")
	oGiftGroups.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oGiftGroups.EOF Then
		
		Do While NOT oGiftGroups.EOF
			' DRAW PAYMENT FIELD GROUP TITLE
			response.write "<div class=reserveformtitle>" & oGiftGroups("giftgroupname") & "</div>"
			response.write "<div class=""reserveforminputarea giftfields"">"
			DrawGiftFields igiftid,oGiftGroups("giftgroupid")
			response.write "</div>"
			oGiftGroups.MoveNext
		Loop

	End If

	Set oGiftGroups = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DrawGiftFields(iGiftID,iGroupID)
'--------------------------------------------------------------------------------------------------
Sub DrawGiftFields(iGiftID,iGroupID)
	Dim sSQL, oGiftFields

	sSQL = "SELECT * FROM egov_gift_fields where giftid='" & iGiftID & "' and groupid='" & iGroupID & "'"

	Set oGiftFields = Server.CreateObject("ADODB.Recordset")
	oGiftFields.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oGiftFields.EOF Then
		
		response.write "<p>"
		response.write "<table cellSpacing=""0"" cellPadding=""0"" width=""625""  border=""0"" id=""Table1"">"

		Do While NOT oGiftFields.EOF

			' CHECK IF IT IS REQUIRED
			If  oGiftFields("isrequired") Then
				sRequired = "<SPAN style=""COLOR: #ff0000"">*&nbsp;</SPAN>"
				' ERROR CHECKING FIELD
				response.write "<input type=""hidden"" name=""ef:custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & "-text/" & oGiftFields("validation") & """ value=""" & oGiftFields("fieldprompt") & """ />"
			Else
				sRequired = " "
			End If

			' SET HEIGHT FOR INPUT BOX BASED ON FIELD TYPE, 1=STANDARD, 2=SIMULATED TEXT AREA
			If  oGiftFields("fieldtype") = 2 Then
				sHeight = "HEIGHT: 100px;"
			Else
				sHeight = " " 
			End If
		
			response.write "<tr>"
			response.write "<td valign=""top"" align=""right"">" & sRequired
			response.write oGiftFields("fieldprompt")
			response.write ": </td>"

			response.write "<td align=""left"">"
			
			Select Case oGiftFields("fieldtype")

				Case "1"
					' TEXT BOX
					response.write "<input name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ type=""text"" style=""WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" MAXLENGTH=""" & oGiftFields("maxfieldsize") & """> <b>"  & oGiftFields("helptext") & "</b>"
				Case "2"
					' PSEUDO TEXT AREA
					response.write "<input name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ type=""text"" style=""WIDTH: 300px; " & sHeight & " FONT-FAMILY: Arial"" MAXLENGTH=""" & oGiftFields("maxfieldsize") & """ ><b>"  & oGiftFields("helptext") & "</b>"
				Case "3"
					' SELECT BOX
					arrAnswers = split(oGiftFields("fieldchoices"),"@@")
			
					response.write "<select name=""custom_" & oGiftFields("fieldid") & "_" & oGiftFields("groupid") & """ >"
					For alist = 0 to ubound(arrAnswers)
						response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
					Next
					response.write "</select><b>"  & oGiftFields("helptext") & "</b>"

				Case Else
					' UNKNOWN TYPE DONT PROCESS
					response.write "INPUT TYPE ERROR. PLEASE CONSULT SETUP."
			
			End Select

			response.write "</td></tr>"

			oGiftFields.MoveNext
		Loop

		response.write "</table>"
		response.write "</p>"

	End If

	oGiftFields.close 
	Set oGiftFields = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' GETGIFTINFORMATION(IGIFTID)
'--------------------------------------------------------------------------------------------------
Sub GetGiftInformation(iGiftID)
	Dim sSQL, oGift

	sSQL = "SELECT giftname,amount from egov_gift where giftid='" & iGiftID & "'"

	Set oGift = Server.CreateObject("ADODB.Recordset")
	oGift.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oGift.EOF Then
		sGiftName = oGift("giftname")
		sGiftAmount = oGift("Amount")
	Else
		sGiftName = "Unknown Gift"
		sGiftAmount = "0.00"
	End If
	
	oGift.close
	Set oGift = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SETUSERINFORMATION()
'--------------------------------------------------------------------------------------------------
Sub SetUserInformation()
	If sOrgRegistration Then 
		If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
			
			iUserID = request.cookies("userid")
		
			sSQL = "SELECT * FROM egov_users WHERE userid=" & iUserID
			Set oInfo = Server.CreateObject("ADODB.Recordset")
			oInfo.Open sSQL, Application("DSN") , 3, 1

			If NOT oInfo.EOF Then
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
	End If
End Sub


%>

