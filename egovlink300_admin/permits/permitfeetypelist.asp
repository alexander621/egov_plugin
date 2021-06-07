<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfeetypelist.asp
' AUTHOR: Steve Loar
' CREATED: 1/07/08
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit fee types
'
' MODIFICATION HISTORY
' 1.0   01/07/08	Steve Loar - INITIAL VERSION
' 1.1	04/14/2008	Steve Loar - Valuation Fees Added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit fee types", sLevel	' In common.asp
PageDisplayCheck "permit types", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="permitfeedd.js"></script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Fee Types</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmFeeSearch" method="post" action="permitfeetypelist.asp">
				<div id="functionlinks">
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="40" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
     					<div class="dropdown">
  						<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-plus" aria-hidden="true"></i> New</button>
  						<div class="dropdown-content">
							<a href="permitformulafeetypeedit.asp?permitfeetypeid=0">Formula Fee</a>
							<a href="permitfixturefeetypeedit.asp?permitfeetypeid=0">Fixture Fee</a>
							<a href="permitvaluationfeetypeedit.asp?permitfeetypeid=0">Valuation Fee</a>
							<a href="permitconstructiontypefeeedit.asp?permitfeetypeid=0">Construction Type Fee</a>
							<a href="permitpercentagetypefeeedit.asp?permitfeetypeid=0">Percentage Fee</a>
							<a href="permitresidentalunitfeetypeedit.asp?permitfeetypeid=0">Resident Unit Type Fee</a>
						</div>
					</div> &nbsp; &nbsp;
					<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
				</div>
			</form>

			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Fee Types</th><th>Calculation<br />Method</th></tr>
				<%	
					ShowPermitFeeTypes session("orgid"), sSearch 
				%>
			</table>
		</div>
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
' void ShowPermitFeeTypes( iOrgid, sSearch )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFeeTypes( ByVal iOrgid, ByVal sSearch )
	Dim sSql, oRates, iRowCount, iPermitFeeTypeid, sUrl

	iRowCount = 0
	iPermitFixtureTypeid = CLng(0)

	sSql = "SELECT F.permitfeetypeid, F.permitfee, F.isfixturetypefee, F.isvaluationtypefee, M.permitfeemethod, "
	sSql = sSql & " F.isbuildingpermitfee, F.isconstructiontypefee, F.ispercentagetypefee, F.isresidentialunittypefee "
	sSql = sSql & " FROM egov_permitfeetypes F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND F.permitfee LIKE '%" & dbsafe(sSearch) & "%' "
	End If 
	sSql = sSql & " ORDER BY F.permitfee, F.permitfeetypeid"

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSQL, Application("DSN"), 3, 1

	If Not oRates.EOF Then
		Do While Not oRates.EOF
			If iPermitFeeTypeid <> CLng(oRates("permitfeetypeid")) Then 
				If iPermitFeeTypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iPermitFeeTypeid = CLng(oRates("permitfeetypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			End If 

			If oRates("isfixturetypefee") Then
				sUrl = "permitfixturefeetypeedit"
			Else
				If oRates("isvaluationtypefee") Then
					sUrl = "permitvaluationfeetypeedit"
				Else
					If oRates("isconstructiontypefee") Then
						sUrl = "permitconstructiontypefeeedit"
					Else
						If oRates("ispercentagetypefee") Then
							sUrl = "permitpercentagetypefeeedit"
						Else
							If oRates("isresidentialunittypefee") Then
								sUrl = "permitresidentalunitfeetypeedit"
							Else
								sUrl = "permitformulafeetypeedit"
							End If 
						End If 
					End If 
				End If 
			End If 

			response.write "<td class=""leftcol"" title=""click to edit"" onClick=""location.href='" & sUrl & ".asp?permitfeetypeid=" & oRates("permitfeetypeid") & "';"">&nbsp;" & oRates("permitfee") & "</td>"
'			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='" & sUrl & ".asp?permitfeetypeid=" & oRates("permitfeetypeid") & "';"">"
'			If oRates("isbuildingpermitfee") Then
'				response.write "yes"
'			Else 
'				response.write "&nbsp;"
'			End If 
'			response.write "</td>"
			response.write "<td class=""feemethod"" title=""click to edit"" onClick=""location.href='" & sUrl & ".asp?permitfeetypeid=" & oRates("permitfeetypeid") & "';"">&nbsp;" & Left(oRates("permitfeemethod"),23) & "</td>"

			oRates.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Fee Types could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Fee Types could be found. Click on one of the New Fee buttons to start entering data.</td></tr>"
		End If 
	End If  
	
	oRates.Close
	Set oRates = Nothing 

End Sub 



%>
