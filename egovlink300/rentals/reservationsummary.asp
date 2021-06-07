<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationsummary.asp
' AUTHOR: Steve Loar
' CREATED: 02/03/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Summary page of reservation details before going off to the payment page or receipt page
'
' MODIFICATION HISTORY
' 1.0   02/03/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Force the page to be re-loaded on back button
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button after purchase problems

Dim iReservationTempId, sTitle, sMessage, sLoadMsg, iRentalId, bHasData, bHasHours
Dim bIsAllDayOnly, sStartTimeLabel, sEndTimeLabel, sSelectedDate
Dim iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, sDepartureAmPm
Dim iArrivalHour, iArrivalMinute, sArrivalAmPm, iDepartureHour, iDepartureMinute
Dim iFeeTotal, sStartDateTime, sEndDateTime, iIncludePriceTypeId, iCitizenUserId
Dim sResidentType, sUserType, sSource

If request("rti") = "" Then
	response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
Else 
	If Not IsNumeric(request("rti")) Then
		response.redirect sEgovWebsiteURL & "/rentals/rentalcategories.asp"
	Else 
		iReservationTempId = CLng(request("rti"))
	End If 
End If 

sSource = request("src")

iIncludePriceTypeId = 0

' still need to confirm that the data is there, and if not take them away from this page.
bHasData = SetPageVariables( iReservationTempId, iOrgId )

If bHasData = False  Then 
	' Take them somewhere safe, as their data is gone.
	response.redirect "rentalcategories.asp"
End If 

sMessage = request("msg")
If sMessage = "nt" Then
	sLoadMsg = "displayScreenMsg('You must agree to the terms/conditions before continuing.');"
End If

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

If bIsAllDayOnly Then 
	sStartLabel = "Arrival"
	sEndLabel = "Departure"
	sShowStartHour = iArrivalHour
	sShowStartMinute = iArrivalMinute
	sShowStartAmPm = sArrivalAmPm
	sShowEndHour = iDepartureHour
	sShowEndMinute = iDepartureMinute
	sShowEndAmPm = sDepartureAmPm
Else
	sStartLabel = "Start"
	sEndLabel = "End"
	sShowStartHour = iStartHour
	sShowStartMinute = iStartMinute
	sShowStartAmPm = sStartAmPm
	sShowEndHour = iEndHour
	sShowEndMinute = iEndMinute
	sShowEndAmPm = sEndAmPm
End If 

iFeeTotal = CDbl(0.00)

%>

<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js" ></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function goBack()
		{
			document.frmBack.submit();
		}

		function continueReservation()
		{
			document.frmRentalSummary.submit();
		}


<%		If sLoadMsg <> "" Then	%>		
			YAHOO.util.Event.onDOMReady(initMsg);
			
			function initMsg() {
				<%=sLoadMsg%>
			}

			function displayScreenMsg( iMsg ) 
			{
				if(iMsg!="") 
				{
					$("screenMsg").innerHTML = iMsg;
					window.setTimeout("clearScreenMsg()", (10 * 1000));
				}
			}

			function clearScreenMsg() 
			{
				$("screenMsg").innerHTML = "";
			}

<%		End If %>

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<span id="screenMsg">&nbsp;</span>

<p>
	<input type="button" class="button" value="<< Back" onclick="goBack();" />
</p>

<form name="frmBack" method="post" action="rentaltimeselection.asp">
	<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
</form>

<%
'CHECK TO SEE IF USER/ADDRESS IS BLOCKED, ONLY BLOCK FOR MONTGOMERY (ORG 26)
if iOrgId = 26 or iOrgId = 37 then
	sSQL = "SELECT u.userid " _
			& " FROM egov_Users u " _
			& " INNER JOIN egov_users u2 ON ((u2.useraddress = u.useraddress AND u2.usercity = u.usercity AND u2.userstate = u.userstate " _
										& " AND u2.userzip = u.userzip) OR u2.useremail = u.useremail)  " _
										& " AND u2.orgid = u.orgid " _
			& " WHERE u2.userID = '" & Track_DBSafe(request.cookies("userid")) & "' AND u.FacilityABUSE = 1 "
	Set oBlock = Server.CreateObject("ADODB.RecordSet")
	oBlock.Open sSQL, Application("DSN"), 3, 1
	if not oBlock.EOF then
		response.write "<p>We are not able to make this reservation at this time.  Please call 513-891-2424 for more information.</p>" 
		response.end
	end if
	oBlock.Close
	Set oBlock = Nothing
end if
%>

<p id="summarytitle">Please review your reservation information for accuracy and, if required, agree to any terms before continuing.
</p>

<span class="summarysubtitle">You have selected: </span>
<p class="selectedsummarydetails">
	<% ShowRentalNameAndLocation iRentalId  %><br /><br />

	Reservation Date: <% response.write WeekDayName(Weekday(CDate(sSelectedDate))) & ", " & sSelectedDate%><br /><br />

<%
	If bIsAllDayOnly Then
		response.write "The reservation period is all day.<br />"
	Else
		response.write "Reservation period:<br />"
	End If 
	response.write "<span class=""summarytime"">" & sStartLabel & " Time: " & sShowStartHour & ":" & sShowStartMinute & " " & sShowStartAmPm & "</span><br />"
	response.write "<span class=""summarytime"">" & sEndLabel & " Time: " & sShowEndHour & ":" & sShowEndMinute & " " & sShowEndAmPm & "</span><br />"
%>
</p>

<hr class="summarydivider" />

<span class="summarysubtitle">Fees:</span>
<div id="summaryfees">
	<table id="summaryfeetable" cellpadding="0" cellspacing="0" border="0">
<%'		<tr><th align="center">Charges</th><th align="center">Amount</th></tr>		%>
<%		iFeeTotal = ShowReservationFees( iRentalId, bIsAllDayOnly, sStartDateTime, sEndDateTime, iIncludePriceTypeId )
		response.write vbcrlf & "<tr><td class=""firstcol feetotal"" colspan=""2"">Total Fees</td><td class=""feetotal lastcol"" align=""right"">" & FormatNumber(iFeeTotal,2,,,0) & "</td></tr>"
		' put the fee total in the temp table for the payment form to access
		UpdateReservationTotalFees iReservationTempId, iFeeTotal
%>
	</table>
</div>

<form method="post" name="frmRentalSummary" action="rentalcontrol.asp">
	<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
	<input type="hidden" id="src" name="src" value="sp" />
	<p>

<%		ShowRentalTerms iRentalId		%>

	<br /><br />
	<input type="button" class="button" onclick="continueReservation();" value="Continue With Reservation" />
	</p>
</form>


<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' boolean SetPageVariables( iReservationTempId, iOrgId )
'--------------------------------------------------------------------------------------------------
Function SetPageVariables( ByVal iReservationTempId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT rentalid, selecteddate, ISNULL(starthour,1) AS starthour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(startminute,0),2) AS startminute, "
	sSql = sSql & " ISNULL(startampm,'PM') AS startampm, ISNULL(endhour,2) AS endhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(endminute,0),2) AS endminute, "
	sSql = sSql & " ISNULL(endampm,'PM') AS endampm, ISNULL(arrivalhour,1) AS arrivalhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(arrivalminute,0),2) AS arrivalminute, "
	sSql = sSql & " ISNULL(arrivalampm,'PM') AS arrivalampm, ISNULL(departurehour,2) AS departurehour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(departureminute,0),2) AS departureminute, "
	sSql = sSql & " ISNULL(departureampm,'PM') AS departureampm, isallday, "
	sSql = sSql & " ISNULL(includepricetypeid,0) AS includepricetypeid, ISNULL(citizenuserid,0) AS citizenuserid "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iRentalId = CLng(oRs("rentalid"))
		sSelectedDate = oRs("selecteddate")
		iStartHour = oRs("starthour")
		iStartMinute = oRs("startminute")
		sStartAmPm = oRs("startampm")
		iEndHour = oRs("endhour")
		iEndMinute = oRs("endminute")
		sEndAmPm = oRs("endampm")
		iArrivalHour = oRs("arrivalhour")
		iArrivalMinute = oRs("arrivalminute")
		sArrivalAmPm = oRs("arrivalampm")
		iDepartureHour = oRs("departurehour")
		iDepartureMinute = oRs("departureminute")
		sDepartureAmPm = oRs("departureampm")
		If oRs("isallday") Then
			bIsAllDayOnly = True 
		Else
			bIsAllDayOnly = False 
		End If 
		sStartDateTime = CDate(sSelectedDate & " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm)
		
		sEndDateTime = CDate(sSelectedDate & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm)
		If sEndDateTime < sStartDateTime Then 
			sEndDateTime = DateAdd("d", 1, sEndDateTime)
		End If 

		iIncludePriceTypeId = oRs("includepricetypeid")
		iCitizenUserId = oRs("citizenuserid")

		sResidentType = GetUserResidentType( iCitizenUserId )
		'If they are not one of these (R, N), we have to figure which they are
		If sUserType <> "R" And sUserType <> "N" Then 
			'This leaves E and B - See if they are a resident, also
			sUserType = GetResidentTypeByAddress( iCitizenUserId, iOrgId )
		End If 


		SetPageVariables = True 
	Else
		SetPageVariables = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' double ShowReservationFees( iRentalId, bIsAllDayOnly, sStartDateTime, sEndDateTime, iIncludePriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ShowReservationFees( ByVal iRentalId, ByVal bIsAllDayOnly, ByVal sStartDateTime, ByVal sEndDateTime, ByVal iIncludePriceTypeId )
	Dim sTime, dHours, bOffSeasonFlag, iWeekday

	dHours = CalculateDurationInHours( sStartDateTime, sEndDateTime )

	bOffSeasonFlag = GetOffSeasonFlag( iRentalId, CDate(sStartDateTime) )

	iWeekday = Weekday(CDate(sStartDateTime))

	' find out if this is all day only
	If bIsAllDayOnly Then 
		sTime = "All Day"
	Else 
		' Get the reservation time
		sTime = FormatNumber(dHours,2,,,0) & " hours"
	End If 

	' Find out if this is a no charge rental
	If RentalHasNoCosts( iRentalId ) Then 
		' this is all there is for phase 1
		response.write "<tr><td class=""firstcol"">" & sTime & "</td><td>" & sTime & " at No Charge</td><td align=""right"" class=""lastcol"">0.00</td></tr>"
		ShowReservationFees = CDbl(0.00)
	Else 
		' Get the applicable hourly fees - only want the total like NonResident 3 hrs $75.00 
		' Get and apply any weekend surcharges
		dRateCharges = ShowRates( iRentalid, bOffSeasonFlag, iWeekday, sResidentType, dHours, bIsAllDayOnly, sStartDateTime, sEndDateTime )
		' Get the deposit amount
		' Get the alcohol surcharge if they checked that
		dRentalChargeTotal = ShowRentalCharges( iRentalid, iIncludePriceTypeId )
		ShowReservationFees = dRateCharges + dRentalChargeTotal

	
		if iOrgId = 228 then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""firstcol"" colspan=""2"">"
			response.write "Processing Fee</td>"
			response.write "<td align=""right"" class=""lastcol"">" & FormatNumber(ShowReservationFees * .035,2,,,0) 
			response.write "</td>"
			response.write "</tr>"
			ShowReservationFees = ShowReservationFees * 1.035
		end if
	End If 

End Function



'--------------------------------------------------------------------------------------------------
' double ShowRentalCharges( iRentalid, iIncludePriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ShowRentalCharges( ByVal iRentalid, ByVal iIncludePriceTypeId )
	Dim sSql, oRs, dTotal, bInclude

	dTotal = CDbl(0.00)

	sSql = "SELECT P.pricetypeid, P.pricetypename, ISNULL(F.amount,0.00) AS amount, P.needsprompt, "
	sSql = sSql & " ISNULL(F.prompt,'') AS prompt "
	sSql = sSql & " FROM egov_price_types P, egov_rentalfees F "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.rentalid = " & iRentalid
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("needsprompt") Then
			If CLng(iIncludePriceTypeId) = CLng(oRs("pricetypeid")) Then
				bInclude = True 
			Else 
				bInclude = False 
			End If 
		Else
			bInclude = True 
		End If 

		If bInclude Then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""firstcol"" colspan=""2"">"
			response.write oRs("pricetypename") & "</td>"
			response.write "<td align=""right"" class=""lastcol"">" & FormatNumber(oRs("amount"),2,,,0) 
			dTotal = dTotal + CDbl(oRs("amount"))
			response.write "</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowRentalCharges = dTotal

End Function 



'--------------------------------------------------------------------------------------------------
' void ShowRentalTerms iRentalId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalTerms( ByVal iRentalId )
	Dim sTerms

	sTerms = GetRentalTerms( iRentalId )

	If sTerms <> "" Then
		response.write vbcrlf & "<hr class=""summarydivider"" />"
		response.write vbcrlf & "<span class=""summarysubtitle"">Important! You must read and agree to each of the terms/conditions below.</span>"
		response.write vbcrlf & "<table id=""summaryterms"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td valign=""top"" align=""center"" class=""checkcol"">"
		response.write "<input type=""checkbox"" name=""agreetoterms"" />"
		response.write "</td>"
		response.write "<td class=""termcol"">" & sTerms
		response.write "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
	Else
		response.write "<input type=""hidden"" name=""agreetoterms"" value=""noterms"" />"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' double ShowRates( iRentalid, bOffSeasonFlag, iWeekday, sResidentType, dHours, bIsAllDayOnly, sStartDateTime, sEndDateTime )
'--------------------------------------------------------------------------------------------------
Function ShowRates( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal sResidentType, ByVal dHours, ByVal bIsAllDayOnly, ByVal sStartDateTime, ByVal sEndDateTime )
	Dim sSql, oRs, sPriceTypeName, dBaseAmount, dAmount, sRateType, iBasePriceTypeId, dTotal, dCharge
	Dim dSurchargeStart, iDuration, sTimeUnit, sDuration

	dTotal = CDbl(0.00)

	sSql = "SELECT R.pricetypeid, P.pricetypename, R.ratetypeid, ISNULL(amount,0.00) AS amount, T.ratetype, "
	sSql = sSql & "ISNULL(R.starthour,0) AS starthour, dbo.AddLeadingZeros(ISNULL(R.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(R.startampm,'AM') AS startampm, P.pricetype, P.isbaseprice, P.isfee, P.isweekendsurcharge, "
	sSql = sSql & "ISNULL(P.basepricetypeid,0) AS basepricetypeid, P.checkresidency, P.isresident, T.datediffstring "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & "AND T.ratetypeid = R.ratetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	sSql = sSql & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sPriceTypeName = oRs("pricetypename")
		sRateType = oRs("ratetype")
		If oRs("isbaseprice") Then 
			dBaseAmount = CDbl(oRs("amount"))
			dAmount = dBaseAmount
			iBasePriceTypeId = oRs("pricetypeid")
			'response.write "<!--ISBase: " & dBaseAmount & "-->"
			'response.write "<!--BASEAmount: " & dAmount & "-->"
		Else
			If oRs("isfee") Then
				dAmount = CDbl(oRs("amount"))
				bShow = True
			Else 
				If CLng(iBasePriceTypeId) = CLng(oRs("basepricetypeid")) Then 
					dAmount = dBaseAmount + CDbl(oRs("amount"))	
				Else
					dAmount = CDbl(oRs("amount"))
				End If 
			End If 
			'response.write "<!--Base: " & dBaseAmount & "-->"
			'response.write "<!--Amount: " & dAmount & "-->"
		End If 
		If oRs("checkresidency") Then
			If oRs("pricetype") = sResidentType Then
				bShow = True 
			Else
				bShow = False 
			End If
		Else
			bShow = True 
		End If 

		If bShow Then 
			If Not oRs("isweekendsurcharge") Then 
				response.write vbcrlf & "<tr>"
				response.write "<td class=""firstcol"">" & sPriceTypeName & "</td>"
				response.write "<td align=""center"">"
				If bIsAllDayOnly Then 
					dCharge = CDbl(dAmount)
					sDuration = " all day"
					sTimeUnit = "day"
				Else 
					' see if the rate is in $ per hour
					If oRs("datediffstring") = "h" Then
						dCharge = CDbl(dAmount * dHours)
						sDuration = dHours & " hours"
						sTimeUnit = "hour"
					Else
						' this would leave minutes
						dCharge = CDbl(dAmount * dHours * 60)
						sDuration = (dHours * 60) & " minutes"
						sTimeUnit = "minute"
					End If 
				End If 
				'response.write "</td>"
				response.write sDuration & " at " & FormatNumber(dAmount,2,,,0) & " per " & sTimeUnit & "</td>"

				dTotal = dTotal + dCharge
				response.write "<td align=""right"" class=""lastcol"">" & FormatNumber(dCharge,2,,,0)
				response.write "</td>"
				response.write "</tr>"
			Else 
				If clng(oRs("starthour")) > clng(0) Then
					' Need the hours for time after the surcharge starts
					'response.write " added for any time after " & DateValue(sStartDateTime) & " " & oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm")
					dSurchargeStart = CDate(DateValue(sStartDateTime) & " " & oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm"))
					If DateDiff( "n", CDate(sEndDateTime), CDate(sSurchargeStart)) < 0 Then
						If DateDiff( "n", CDate(sStartDateTime), CDate(dSurchargeStart)) < 0 Then
							iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
							If oRs("datediffstring") = "h" Then
								iDuration = CDbl(iDuration / 60)
							End If 
							dCharge = CDbl(iDuration) * dAmount
							sDuration = iDuration & " hours"
							sTimeUnit = "hour"
						Else
							iDuration = DateDiff("n", dSurchargeStart, CDate(sEndDateTime))
							If oRs("datediffstring") = "h" Then
								iDuration = CDbl(iDuration / 60)
								sTimeUnit = "hour"
								sDuration = iDuration & " hours"
							ElseIf oRs("datediffstring") = "d" Then
								iDuration = CDbl(1.00)
								sDuration = " all day"
								sTimeUnit = "day"
							End If 
							dCharge = CDbl(iDuration) * dAmount
						End If 
						If dCharge > CDbl(0.00) Then 
							response.write vbcrlf & "<tr>"
							response.write "<td class=""firstcol"">" & sPriceTypeName & "</td>"
							response.write "<td align=""center"">"
							response.write sDuration & " at " & FormatNumber(dAmount,2,,,0) & " per " & sTimeUnit & "</td>"
							response.write "</td>"
							dTotal = dTotal + dCharge
							response.write "<td align=""right"" class=""lastcol"">" & FormatNumber(dCharge,2,,,0)
							response.write "</td>"
							response.write "</tr>"
						End If 
					End If 
				End If 
			End If 
			
		End If 
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

	ShowRates = dTotal

End Function 


'--------------------------------------------------------------------------------------------------
' void UpdateReservationTotalFees iReservationTempId, iFeeTotal
'--------------------------------------------------------------------------------------------------
Sub UpdateReservationTotalFees( ByVal iReservationTempId, ByVal iFeeTotal )
	Dim sSql

	sSql = "UPDATE egov_rentalreservationstemppublic "
	sSql = sSql & "SET feetotal = " & iFeeTotal
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

End Sub 



%>
