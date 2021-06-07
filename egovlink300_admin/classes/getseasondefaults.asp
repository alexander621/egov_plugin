<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getseasondefaults.asp
' AUTHOR: Steve Loar
' CREATED: 03/02/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets season defaults, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   03/02/07	Steve Loar - INITIAL VERSION
' 1.1	08/13/2007	Steve Loar - Added Registrationenddate
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sDefaultString, oSeason, oPriceTypeSeason

	sSql = "Select registrationstartdate, registrationenddate, publicationstartdate, publicationenddate "
	sSql = sSql & " from egov_class_seasons where classseasonid = " & CLng(request("csi"))

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), 0, 1

	If IsNull(oSeason("registrationstartdate")) Or Trim(oSeason("registrationstartdate")) = "" Then 
		sDefaultString = "NULL"
	Else 
		sDefaultString = oSeason("registrationstartdate")
	End If 

	If IsNull(oSeason("publicationstartdate")) Or Trim(oSeason("publicationstartdate")) = "" Then 
		sDefaultString = sDefaultString & "|NULL"
	Else 
		sDefaultString = sDefaultString & "|" & oSeason("publicationstartdate")
	End If 

	If IsNull(oSeason("publicationenddate")) Or Trim(oSeason("publicationenddate")) = "" Then 
		sDefaultString = sDefaultString & "|NULL"
	Else 
		sDefaultString = sDefaultString & "|" & oSeason("publicationenddate")
	End If 

	If IsNull(oSeason("registrationenddate")) Or Trim(oSeason("registrationenddate")) = "" Then 
		sDefaultString = sDefaultString & "|NULL"
	Else 
		sDefaultString = sDefaultString & "|" & oSeason("registrationenddate")
	End If 

	oSeason.close
	Set oSeason = Nothing

	sSql = "Select pricetypeid, registrationstartdate "
	sSql = sSql & " from egov_class_seasons_to_pricetypes_dates where classseasonid = " & CLng(request("csi"))

	Set oPriceTypeSeason = Server.CreateObject("ADODB.Recordset")
	oPriceTypeSeason.Open sSQL, Application("DSN"), 0, 1

	If Not oPriceTypeSeason.EOF Then 
		Do While Not oPriceTypeSeason.EOF
			If IsNull(oPriceTypeSeason("registrationstartdate")) Or Trim(oPriceTypeSeason("registrationstartdate")) = "" Then 
				sDefaultString = sDefaultString & "|" & oPriceTypeSeason("pricetypeid") & ";NULL"
			Else 
				sDefaultString = sDefaultString & "|" & oPriceTypeSeason("pricetypeid") & ";" & oPriceTypeSeason("registrationstartdate")
			End If 
			oPriceTypeSeason.MoveNext 
		Loop 
	End If
	
	oPriceTypeSeason.Close
	Set oPriceTypeSeason = Nothing 

	response.write sDefaultString

%>