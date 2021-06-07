<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DISCOUNT_EDIT.ASP
' AUTHOR: TERRY FOSTSER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/26/06   TERRY FOSTER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "discounts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' INITIALIZE VARIABLES
Dim sName, sQtyRequired, sAmount, sDescription, blnIsShared, iDiscountTypeId
Dim iPriceDiscountID 
Dim iClassID 

' GET DISCOUNT ID
If request("PriceDiscountid") = "" OR NOT IsNumeric(request("PriceDiscountid")) OR clng(request("PriceDiscountid")) = 0 Then
	' CREATE NEW DISCOUNT
	iPriceDiscountID = 0
	sTitle = "Add New Discount"
	sLinkText = "Create Discount"
Else
	' EDIT EXISTING DISCOUNT
	iPriceDiscountID = request("PriceDiscountid")
	sTitle = "Edit Discount"
	sLinkText = "Save Changes"
End If

' GET DISCOUNT INFORMATION
GetDiscountInfo iPriceDiscountID 

If blnIsShared Then
	sChecked= " checked=""checked"" "
End If

%>


<html>
<head>
	 <meta charset="UTF-8">

 	<title>E-Gov Administration Console</title>

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="../recreation/facility.css" />
 	<link rel="stylesheet" href="classes.css" />

 	<script src="tablesort.js"></script>

	<script>
	<!--

	function IsNumeric(sText) 
	{
		var ValidChars = "01234567890.";
		var IsNumber=true;
		var Char;
		
		for(i=0;i<sText.length && IsNumber == true;i++) {
			Char = sText.charAt(i);
			if(ValidChars.indexOf(Char) == -1) {
				IsNumber = false;
			}
		}
		return IsNumber;
	}

	function save() 
	{
		theForm = document.frmdiscount

		var msg = "";

		if(!IsNumeric(theForm.sQTYRequired.value)) {
			msg += "You must enter a numeric value for Qty. Required.\n";
		}
		
		if(!IsNumeric(theForm.sAmount.value)) {
			msg += "You must enter a numeric value for Amount.\n";
		}

		if(msg != "") {
			alert("Your form cannot be submitted for the following reasons:\n\n" + msg);
		}
		else {
			theForm.submit();
		}
	}

	//-->
	</script>

</head>

<body>
 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: <%=sTitle%></strong></font><br />
	<!--<a href="discount_mgmt.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: FUNCTION LINKS-->
<div id="functionlinks">
		<a href="discount_mgmt.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">Return to Discount Management</a>&nbsp;&nbsp;
		<a href="javascript:save();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;<%=sLinkText%></a>&nbsp;&nbsp;
</div>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="frmdiscount" action="discount_save.asp" method="post">
<input type="hidden" name="iPriceDiscountid" value="<%=iPriceDiscountID%>" />

<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="instructortable">
		<tr>
			<th>Price Discount Information</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td align="right">Name:</td>
						<td><input type="text" name="sName" size="50" maxlength="50" value="<%=sName%>" /></td>
					</tr>
					<tr>
						<td align="right">Discount Type:</td>
						<td><% ShowDiscountPicks iDiscountTypeId  %></td>
					</tr>
					<tr>
						<td align="right">Qty. Required:</td>
						<td><input type="text" name="sQTYRequired" value="<%=sQTYRequired%>" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td align="right">Discounted Price:</td>
						<td><input type="text" name="sAmount" value="<%= FormatNumber(sAmount,2)%>" size="6" maxlength="6" /></td>
					</tr>
					<tr>
						<td align="right">Description:</td>
						<td><input type="text" name="sDescription"  size="100" maxlength="255" value="<%=sDescription%>" /></td>
					</tr>
					<tr>
						<td align="right"><input <%=sChecked%> name="bIsShared" Type="checkbox" /></td>
						<td>Is shared amoung all activities with this discount that require registration.</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

</form>
<!--END: EDIT FORM-->

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
' Sub GetDiscountInfo( icategoryID )
'--------------------------------------------------------------------------------------------------
Sub GetDiscountInfo( icategoryID )
	Dim sSql, oValues 

	sSQL = "SELECT * FROM egov_price_discount WHERE pricediscountid = '" & iPriceDiscountID & "'"
	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If NOT oValues.EOF Then
		sName = oValues("discountname")
		iClassID = oValues("ClassID")
		sQtyRequired = oValues("qtyrequired")
		sAmount = oValues("discountamount")
		sDescription = oValues("discountdescription")
		blnIsShared = oValues("isshared")
		iDiscountTypeId = oValues("discounttypeid")
	End If

	oValues.close
	Set oValues = nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowDiscountPicks( iDiscountTypeId )
'--------------------------------------------------------------------------------------------------
Sub ShowDiscountPicks( iDiscountTypeId )
	Dim sSql, oValues 

	sSQL = "SELECT discounttypeid, discounttype, discounttypedescription FROM egov_price_discount_types Order by discounttypeid"
	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oValues.EOF Then 
		response.write vbcrlf & "<select name=""discounttypeid"">"
		Do While Not oValues.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oValues("discounttypeid") & """ "
			If clng(iDiscountTypeId) = clng(oValues("discounttypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oValues("discounttype") & " &ndash; " & oValues("discounttypedescription") & "</option>"
			oValues.movenext
		Loop
		response.write vbcrlf & "</select>"
	End If 

	oValues.close
	Set oValues = Nothing 

End Sub 
%>


