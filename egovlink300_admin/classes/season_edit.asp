<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: season_edit.asp
' AUTHOR: Steve Loar
' CREATED: 2/1/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the Season Edit page
'
' MODIFICATION HISTORY
' 1.0   2/1/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassSeasonId, sSeasonName, iSeasonId, iSeasonYear, sRegistrationStartDate, sPublicationEndDate
Dim sPublicationStartDate, iIsClosed, iShowPublic, sRegistrationEndDate

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "seasons" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("sid") <> "0" Then 
	iClassSeasonId = clng(request("sid"))
	GetSeasonProperties iClassSeasonId
Else
	iClassSeasonId = 0 ' New Season
End If 


%>

<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

<script language="Javascript">
<!--

	function Validate()
	{
		var rege;
		var Ok;

		// check the name
		if (document.FormSeason.seasonname.value == "")
		{
			alert('Please enter a name.');
			document.FormSeason.seasonname.focus();
			return;
		}
		if (document.FormSeason.seasonname.value.length > 50)
		{
			alert('The name is limited to 50 characters.\nPlease shorten this name.');
			document.FormSeason.seasonname.focus();
			return;
		}
		// check the year
		if (document.FormSeason.seasonyear.value == "")
		{
			alert('Please enter a year.');
			document.FormSeason.seasonyear.focus();
			return;
		}
		if (document.FormSeason.seasonyear.value.length > 0)
		{
			rege = /^\d+$/;
			Ok = rege.test(document.FormSeason.seasonyear.value);

			if (! Ok)
			{
				alert("The year must be a number");
				document.FormSeason.seasonyear.focus();
				return;
			}
		}
		if (document.FormSeason.seasonyear.value.length < 4)
		{
			alert("The year must be four numbers");
			document.FormSeason.seasonyear.focus();
			return;
		}
		// Check Reg start date
		if (document.FormSeason.registrationstartdate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.FormSeason.registrationstartdate.value);
			if (! Ok)
			{
				alert("Registration start date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.FormSeason.registrationstartdate.focus();
				return;
			}
		}
		// Check Reg end date
		if (document.FormSeason.registrationenddate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.FormSeason.registrationenddate.value);
			if (! Ok)
			{
				alert("Registration end date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.FormSeason.registrationenddate.focus();
				return;
			}
		}
		// Check Publication end date
		if (document.FormSeason.publicationenddate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.FormSeason.publicationenddate.value);
			if (! Ok)
			{
				alert("Publication end date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.FormSeason.publicationenddate.focus();
				return;
			}
		}
		// Check Publication end date
		if (document.FormSeason.publicationstartdate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.FormSeason.publicationstartdate.value);
			if (! Ok)
			{
				alert("Publication start date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.FormSeason.publicationstartdate.focus();
				return;
			}
		}

		for (var p = parseInt(document.FormSeason.minpricetypeid.value); p <= parseInt(document.FormSeason.maxpricetypeid.value); p++)
		{
			// Does it exist
			if (document.getElementById("registrationstartdate" + p))
			{
				//alert(document.getElementById("registrationstartdate" + p).value);
				if (document.getElementById("registrationstartdate" + p).value != "")
				{
					rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
					Ok = rege.test(document.getElementById("registrationstartdate" + p).value);
					if (! Ok)
					{
						alert("Registration start dates cannot be blank and should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
						document.getElementById("registrationstartdate" + p).focus();
						return;
					}
//					else
//					{
//						alert(document.getElementById("registrationstartdate" + p).value);
//					}
				}
			}
		}

		//alert("OK");
		document.FormSeason.submit();
	}

	function doCalendar( sField )
	{
      var w = (screen.width - 350)/2;
      var h = (screen.height - 350)/2;
      eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=FormSeason", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
	
		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>
		<%		If clng(iClassSeasonId) = clng(0) Then %>
					New  
		<%		End If %>
				Season Properties</strong></font><br />
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: FUNCTION LINKS-->
		<div id="functionlinks">
				<a href="season_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to Seasons</a>&nbsp;&nbsp;
				<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;
		<%		If clng(iClassSeasonId) = clng(0) Then %>
					Create 
		<%		Else %>
					Update 
		<%		End If %>
				Season</a>&nbsp;&nbsp;
		</div>
		<!--END: FUNCTION LINKS-->


		<!--BEGIN: EDIT FORM-->
		<form name="FormSeason" action="season_update.asp" method="post">
			<input type="hidden" name="classseasonid" value="<%=iClassSeasonId%>" />
			<div class="shadow">
				<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
					<tr><th>Season Properties</th></tr>
					<tr>
						<td>
							<table>
								<tr>
									<td wrap="nowrap" align="right">Name:</td><td><input type="text" name="seasonname" value="<%=sSeasonName%>" size="55" maxlength="50" /></td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">Year:</td><td><input type="text" name="seasonyear" value="<%=iSeasonYear%>" size="6" maxlength="4" /></td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">Season:</td>
									<td><select name="seasonid">
<%											ShowSeasonPicks iSeasonId	%>
										</select>	
									</td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">Registration Start (display):</td><td><input type="text" name="registrationstartdate" value="<%=sRegistrationStartDate%>" class="datefield" />
									&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationstartdate');" /></span>
									</td>
								</tr>

<%									ShowRegistrationStartsByPriceTypes iClassSeasonId %>

								<tr>
									<td wrap="nowrap" align="right">Registration Ends:</td><td><input type="text" name="registrationenddate" value="<%=sRegistrationEndDate%>" class="datefield" />
									&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationenddate');" /></span>
									</td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">Publication Start:</td><td><input type="text" name="publicationstartdate" value="<%=sPublicationStartDate%>" class="datefield" />
									&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationstartdate');" /></span>
									</td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">Publication End:</td><td><input type="text" name="publicationenddate" value="<%=sPublicationEndDate%>" class="datefield" />
									&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationenddate');" /></span>
									</td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">&nbsp;</td><td><input type="checkbox" name="isclosed" <% If iIsClosed Then response.write " checked=""checked"" "%> /> Closed</td>
								</tr>
								<tr>
									<td wrap="nowrap" align="right">&nbsp;</td><td><input type="checkbox" name="showpublic" <% If iShowPublic Then response.write " checked=""checked"" "%> /> Show Public</td>
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

'------------------------------------------------------------------------------------------------------------
' Sub GetSeasonProperties( iClassSeasonId )
'------------------------------------------------------------------------------------------------------------
Sub GetSeasonProperties( iClassSeasonId )
	Dim sSql, oSeason

	sSql = "SELECT seasonname, seasonid, seasonyear, registrationstartdate, registrationenddate, publicationstartdate, publicationenddate, isclosed, showpublic "
	sSql = sSql & " From egov_class_seasons Where classseasonid = " & iClassSeasonId

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), 3, 1

	If Not oSeason.EOF Then 
		sSeasonName = oSeason("seasonname")
		iSeasonId = oSeason("seasonid")
		iSeasonYear = oSeason("seasonyear")
		sRegistrationStartDate = oSeason("registrationstartdate")
		sRegistrationEndDate = oSeason("registrationenddate")
		sPublicationStartDate = oSeason("publicationstartdate")
		sPublicationEndDate = oSeason("publicationenddate")
		iIsClosed = oSeason("isclosed")
		iShowPublic = oSeason("showpublic")
	End If
		
	oSeason.close
	Set oSeason = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowSeasonPicks( iSeasonId )
'------------------------------------------------------------------------------------------------------------
Sub ShowSeasonPicks( iSeasonId )
	Dim sSql, oSeasons

	sSql = "Select seasonid, season from egov_seasons order by displayorder"

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 3, 1

	Do While Not oSeasons.EOF  
		response.write vbcrlf & "<option value=""" & oSeasons("seasonid") & """"
		If clng(iSeasonId) = clng(oSeasons("seasonid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oSeasons("season") & "</option>"
		oSeasons.MoveNext
	Loop 
		
	oSeasons.close
	Set oSeasons = Nothing
End Sub


'------------------------------------------------------------------------------------------------------------
' Sub ShowRegistrationStartsByPriceTypes( iClassSeasonId )
'------------------------------------------------------------------------------------------------------------
Sub ShowRegistrationStartsByPriceTypes( iClassSeasonId )
	Dim sSql, oSeasons, bFirst, iMax, iMin

	bFirst = True 
	iMax = 0
	iMin = 10000

	sSql = "select P.pricetypeid, P.pricetypename + ' Registration Start' as pricetypename, S.registrationstartdate from egov_price_types P "
	sSql = sSql & " left outer join egov_class_seasons_to_pricetypes_dates S on P.pricetypeid = S.pricetypeid and S.classseasonid = " & iClassSeasonId
	sSql = sSql & " where P.orgid = " & SESSION("ORGID") & " and P.needsregistrationstartdate = 1 order by P.displayorder"

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 3, 1

	If Not oSeasons.EOF Then 
		Do While Not oSeasons.EOF  
			If clng(oSeasons("pricetypeid")) > clng(iMax) Then 
				iMax = oSeasons("pricetypeid")
			End If 
			If clng(oSeasons("pricetypeid")) < clng(iMin) Then 
				iMin = oSeasons("pricetypeid")
			End If 
			If Not bFirst Then
				response.write "</td>"
				response.write "</tr>"
			Else
				bFirst = False 
			End If 
			response.write vbcrlf & "<tr>"
			response.write "<td wrap=""nowrap"" align=""right"">" & oSeasons("pricetypename") & ":</td><td><input type=""text"" name=""registrationstartdate" & oSeasons("pricetypeid") & """ id=""registrationstartdate" & oSeasons("pricetypeid") & """ value=""" & oSeasons("registrationstartdate") & """ class=""datefield"" />"
			response.write vbcrlf & "&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('registrationstartdate" & oSeasons("pricetypeid") & "');"" /></span>"
			oSeasons.MoveNext
		Loop

		response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMin & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMax & """ />"
		response.write "</td>"
		response.write "</tr>"
	End If 

	oSeasons.close
	Set oSeasons = Nothing
End Sub 
%>
