<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_offerings
' AUTHOR: Steve Loar
' CREATED: 03/19/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   03/19/2007	Steve Loar - Initial version
' 1.1	10/10/2011	Steve Loar - Added Gender Restriction
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassId, iTimeId, iClassListId

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iClassId = request("classid")
iTimeId = request("timeid")
iClassListId = request("classlistid")
	
Session("RedirectPage") = GetCurrentURL()
Session("RedirectLang") = "Return To Class Roster"

%>


<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	
	<script language="Javascript" src="tablesort.js"></script>
	<script src="../scripts/layers.js"></script>

	<script language="JavaScript">
	<!--

		function GoToPrint( )
		{
			document.frmprint.submit();
		}

		function confirm_drop(iclasslistid, srostername)
		{
			if (confirm("Are you sure you want to drop (" + srostername + ")?"))
				{ 
					// DELETE HAS BEEN VERIFIED
					location.href='drop_registrant.asp?classid=<%=request("classid")%>&timeid=<%=request("timeid")%>&iclasslistid=' + iclasslistid;
				}
		}

		function confirm_move()
		{
			var sclassname = document.frmrosterlist.moveclassid.options[document.frmrosterlist.moveclassid.selectedIndex].text;
			if (confirm("Are you sure you want to move the selected registrants to  (" + sclassname + ")?"))
				{ 
					document.frmrosterlist.submit();
				}
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
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
<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Class Offerings</strong></font><br /><br />
	<a href="roster_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
<form id="FormClassOffering" name="FormClassOffering" action="view_roster.asp" method="post">
	<input type="hidden" name="classid" value="<%=iClassId%>" />

	<fieldset><legend><strong>Class Details</strong></legend>
		<% DisplayClassDetails request("classid") %> 
	</fieldset>

	<fieldset><legend><strong>Pricing</strong></legend>
		<% DisplayClassPricing request("classid") %> 
	</fieldset>

	<fieldset><legend><strong>Available Activities</strong></legend>
		<% DisplayClassActivities request("classid"), 0, True  ' In class_global_functions.asp %>
	</fieldset>

</form>

<!--END: CLASS LIST-->

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
' void DisplayClassPricing iClassId 
'--------------------------------------------------------------------------------------------------
Sub DisplayClassPricing( ByVal iClassId )
	Dim sSql, oRs, dFees

	dFees = GetFeeTotal( iClassId )

	sSql = "SELECT PT.pricetypename, PT.basepricetypeid, ISNULL(CPP.amount,0.00) AS amount, PT.isdropin "
	sSql = sSql & "FROM egov_class_pricetype_price CPP, egov_price_types PT "
	sSql = sSql & "WHERE CPP.pricetypeid = PT.pricetypeid AND PT.isactiveforclasses = 1 AND PT.isfee = 0 "
	sSql = sSql & "AND CPP.classid = " & iClassId & " ORDER BY PT.displayorder"


	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF then
		response.write vbcrlf & "<table id=""offeringprices""cellpadding=""0"" cellspacing=""0"" border=""0"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr><td><strong>" & oRs("pricetypename") & ":</strong></td><td>"
			If oRs("isdropin") Then
				' Drop in is a fixed amount
				response.write FormatCurrency(oRs("amount"),2)
			Else
				If IsNull(oRs("basepricetypeid")) Then
					' IF they are a base price just add any fees
					response.write FormatCurrency((dFees + CDbl(oRs("amount"))),2)
				Else
					' add the fees, this price and its base price
					response.write FormatCurrency((dFees + CDbl(oRs("amount")) + GetBasePrice(oRs("basepricetypeid"),iClassId)),2)
				End If 
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' double GetBasePrice( iBasePriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetBasePrice( ByVal iBasePriceTypeId, ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(amount,0.00) AS amount FROM egov_class_pricetype_price "
	sSql = sSql & "WHERE pricetypeid = " & iBasePriceTypeId & " AND classid = " & iClassId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetBasePrice = CDbl(oRs("amount"))
	Else
		GetBasePrice = CDbl(0.0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetFeeTotal( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetFeeTotal( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT SUM(amount) AS amount FROM egov_class_pricetype_price CPP, egov_price_types PT "
	sSql = sSql & "WHERE CPP.pricetypeid = PT.pricetypeid AND isactiveforclasses = 1 AND isfee = 1 AND classid = " & iClassId & " GROUP BY classid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If Not IsNull(oRs("amount")) Then 
			GetFeeTotal = CDbl(oRs("amount"))
		Else
			GetFeeTotal = CDbl(0.0)
		End If 
	Else
		GetFeeTotal = CDbl(0.0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void DisplayClassDetails iClassId 
'--------------------------------------------------------------------------------------------------
Sub DisplayClassDetails( ByVal iClassId )
	Dim sSql, oRs, bHasGenderRestrictions, iGenderNotRequiredId

	bHasGenderRestrictions = orgHasFeature("gender restriction") 
	iGenderNotRequiredId = GetGenderNotRequiredId( )

	sSql = "SELECT classname, classdescription, startdate, enddate, minage, maxage, ISNULL(genderrestrictionid," & iGenderNotRequiredId & ") AS genderrestrictionid, "
	sSql = sSql & "ISNULL(locationid,0) AS locationid, ISNULL(classseasonid,0) AS classseasonid "
	sSql = sSql & "FROM egov_class WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<p class=""offeringdetails""><strong>Name:</strong>&nbsp; " & oRs("classname") & " &nbsp; &nbsp; &nbsp; <strong>Location:</strong>&nbsp; " & GetLocationName( oRs("locationid") ) & " &nbsp; &nbsp; &nbsp; <strong>Season:</strong>&nbsp; " & GetSeasonName( oRs("classseasonid") ) & "</p>"
		response.write vbcrlf & "<p class=""offeringdetails""><strong>Start Date:</strong>&nbsp; " & oRs("startdate") & " &nbsp; &nbsp; &nbsp; <strong>End Date:</strong>&nbsp; " & oRs("enddate") & "</p>"
		response.write vbcrlf & "<p class=""offeringdetails""><strong>Min Age:</strong>&nbsp; " & oRs("minage") & " &nbsp; &nbsp; &nbsp; <strong>Max Age:</strong>&nbsp; " & oRs("maxage") & "</p>"

		If bHasGenderRestrictions Then 
			response.write vbcrlf & "<p class=""offeringdetails""><strong>Gender Restriction:</strong>&nbsp; " 
			response.write GetGenderRestrictionText( oRs("genderrestrictionid") )		' in class_global_functions.asp
			response.write "</p>"
		End If 
		
		' Display Waiver Links
		response.write vbcrlf & "<p class=""offeringdetails""><strong>Waivers:</strong>&nbsp; " 
		ShowClassWaiverLinks iClassId 
		response.write "</p>"

		response.write vbcrlf & "<p class=""offeringdetails""><strong>Description:</strong>&nbsp; " & oRs("classdescription") & "</p>"
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Sub 



%>