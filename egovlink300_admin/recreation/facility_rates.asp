<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_RATES.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/06	STEVE LOAR - CODE ADDED
' 2.0	01/22/07	JOHN STULLENBERGER - NEW VERSION WITH DIFFERENT PRICING LEVELS
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' INITIALIZE AND DECLARE VARIABLES
Dim oFacilities, iRowCount, iFacilityId, sFacilityName

sLevel = "../" ' OVERRIDE OF VALUE FROM COMMON.ASP

%>


<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Facility Rate Management</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="facility.css" />

	<script language="Javascript">
		<!--
		function ConfirmDelete( sRate, iRateId ) 
		{
			var msg = "Do you wish to delete " + sRate + "?"
			if (confirm(msg))
			{
				location.href='rate_delete.asp?iRateId='+ iRateId;
			}
		}

		function SaveRate( passForm )
		{

			if (passForm.ratedescription.value == "") {
				alert("Please enter a description.");
				passForm.ratedescription.focus();
				return false;
			}

			passForm.submit();
		}

		//-->
	</script>
</head>


<body>

 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	
	<p>
		<font size="+1"><strong>Recreation: Facility Rate Management </strong></font><br />

		<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>

	</p>

	<!-- <div class="shadow"> -->
	
		<% DisplayFacilityRates %>

	<!-- </div> -->
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->    

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUB DISPLAYFACILITYRATES()
'--------------------------------------------------------------------------------------------------
Sub DisplayFacilityRates()
	Dim sSql, oRs, iPriceTypeCount
	
	' DIPSLAY RATES FOR ORG
	sSql = "SELECT pricetypeid, pricetypename FROM egov_price_types WHERE orgid = " & session("orgid")
	sSql = sSql & " AND isactiveforfacility = 1 ORDER BY pricetypegroupid, displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
		
	' IF NOT EMPTY DISPLAY RATES
	If Not oRs.EOF Then

		' BEGIN TABLE
		response.write vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""facilityrates"">"
		
		' BEGIN HEADER ROW
		response.write vbcrlf & "<tr>"

		response.write "<th>Rate Description</th>"

		Do While Not oRs.EOF 
			response.write "<th>" & oRs("pricetypename") & "</th>"
			oRs.MoveNext
		Loop

		response.write "<th>Action</th>"

		response.write "</tr>"
		' END HEADER ROW


		' BEGIN NEW\BLANK ROW
		response.write vbcrlf & "<form name=""rateform0"" method=""post"" action=""rate_save.asp"">"
		response.write vbcrlf & "<tr>"
		response.write "<td>"
		response.write "<input type=""hidden"" name=""iRateId"" value=""0"">"
		response.write "<input type=""text"" name=""ratedescription"" value="""" size=""80"" maxlength=""250""></td>"
		
		' LOOP THRU PRICE TYPES ADDING BLANK
		oRs.MoveFirst
		iPriceTypeCount = 0
		Do While Not oRs.EOF 
			iPriceTypeCount = iPriceTypeCount + 1
			response.write "<td>"
			response.write "<input type=""text"" name=""amount" & iPriceTypeCount & """ value="""" size=""10"" maxlength=""10"" />"
			response.write "<input type=""hidden"" name=""pricetypeid" & iPriceTypeCount & """ value=""" & oRs("pricetypeid") & """ />"
			response.write "</td>"
			oRs.MoveNext
		Loop

		response.write "<td class=""action"" valign=""bottom"">"
		response.write "<input type=""hidden"" name=""pricetypecount"" value=""" & iPriceTypeCount & """ />"
		response.write "<a href=""javascript:SaveRate(document.rateform0);"">Add New</a>"
		response.write "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</form>"
		' END NEW\BLANK ROW

		DisplayFacilityRateData oRs
	
		' END TABLE
		response.write vbcrlf & "</table>"
	
	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYFACILITYRATEDATA()
'--------------------------------------------------------------------------------------------------
Sub DisplayFacilityRateData( ByRef oPriceTypes )
	Dim sSql, oPrice, iRowCount, iPriceTypeCount

	sSql = "SELECT rateid, ratedescription FROM egov_facility_rates WHERE orgid = " & session("orgid")

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSql, Application("DSN"), 3, 1

	
	' IF NOT EMPTY DISPLAY RATES
	If Not oPrice.EOF Then
		iRowCount = 0

		Do While Not oPrice.EOF
			iPriceTypeCount = 0
			iRowCount = iRowCount + 1

			response.write "<form name=""rateform" & iRowCount & """ method=""POST"" action=""rate_save.asp"">"
			
			' SET ROW STYLING
			If iRowCOunt Mod 2 = 1 Then
				response.write "<tr class=""alt_row"">"
			Else
				response.write "<tr>"
			End If

			response.write "<td>"
			response.write "<input type=""hidden"" name=""iRateId"" value=""" & oPrice("rateid") & """ />"
			response.write "<input type=""text"" name=""ratedescription"" value=""" & oPrice("ratedescription") & """ size=""80"" maxlength=""250"" /></td>"

			oPriceTypes.MoveFirst
			Do While Not oPriceTypes.EOF 
				iPriceTypeCount = iPriceTypeCount + 1
				curPrice = GetRatePrice( oPriceTypes("pricetypeid"), oPrice("rateid") )
				If curPrice <> ""  Then
					curPrice = FormatNumber(curPrice,2,,,0)
				End If

				response.write "<td><input type=""text"" name=""amount" & iPriceTypeCount & """ value=""" & curPrice & """ size=""10"" maxlength=""10"" />"
				response.write "<input type=""hidden"" name=""pricetypeid" & iPriceTypeCount & """ value=""" & oPriceTypes("pricetypeid") & """ />"
				response.write "</td>"
				
				oPriceTypes.MoveNext
			Loop
			response.write "<td class=""action"" valign=""bottom"">"
			response.write "<input type=""hidden"" name=""pricetypecount"" value=""" & iPriceTypeCount & """ />"
			response.write "	<a href=""javascript:SaveRate(document.rateform" &  iRowCount & ");"">Save</a>&nbsp;&nbsp;"
			response.write "	<a href=""javascript:ConfirmDelete('" & oPrice("ratedescription") & "'," & oPrice("rateid") & ");"">Delete</a>"
			response.write "</td>"
			response.write "</tr>"
			response.write "</form>"

			oPrice.MoveNext

		Loop

	End If

	oPrice.Close
	Set oPrice = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' FUNCTION GETRATEPRICE(IRATEID,IPRICETYPEID)
'--------------------------------------------------------------------------------------------------
Function GetRatePrice( ByVal iPriceTypeId, ByVal iRateId )
	Dim sSql, oRs, curReturnValue
	
	curReturnValue = 0.00

	sSql = "SELECT amount FROM egov_facility_rate_to_pricetype WHERE rateid = " & CLng(iRateID) & " AND pricetypeid = " & CLng(iPriceTypeID)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		curReturnValue = oRs("Amount")
	End If

	oRs.Close
	Set oRs = Nothing

	GetRatePrice = curReturnValue

End Function


%>

