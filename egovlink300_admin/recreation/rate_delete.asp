<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: RATE_SAVE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE FACILITY RATE AND ADD NEW RATES
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/06	STEVE LOAR - CODE ADDED
' 2.0	01/22/07	JOHN STULLENBERGER - CLEANEDD UP CODE FORMATTING AND DOCUMENTATION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRateId, iFacilityId, sSql, oCmd

iRateId = CLng(request("iRateId"))
iFacilityId = CLng(request("iFacilityId"))


' DELETE FROM THE RATE TABLE
sSql = "DELETE FROM egov_facility_rates WHERE rateid = " & iRateId  
RunSQLStatement sSql


' DELETE FROM PRICE TYPES TO RATE TABLE
sSql = "DELETE FROM egov_facility_rate_to_pricetype WHERE rateid = " & iRateId 
RunSQLStatement sSql


' REDIRECT TO FACILITY RATE PAGE
response.redirect( "facility_rates.asp?facilityid=" & iFacilityId )

%>


