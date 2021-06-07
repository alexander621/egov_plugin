<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
' 
' AUTHOR: Steve Loar
' CREATED: 04/03/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
'Retrieve the search criteria
 orderBy       = session("orderBy")
 subTotals     = session("subTotals")
 showDetail    = session("showDetail")
 fromDate      = session("fromDate")
 toDate        = session("toDate")
 sUserlname    = session("userlname")
 iMembershipId = session("imembershipid")
 iPeriodId     = session("iperiodid")

 lcl_search_criteria = "?"
 lcl_search_criteria = lcl_search_criteria & "orderBy="       & orderBy
 lcl_search_criteria = lcl_search_criteria & "&subTotals="    & subTotals
 lcl_search_criteria = lcl_search_criteria & "&showDetail="   & showDetail
 lcl_search_criteria = lcl_search_criteria & "&fromDate="     & fromDate
 lcl_search_criteria = lcl_search_criteria & "&toDate="       & toDate
 lcl_search_criteria = lcl_search_criteria & "&userlname="    & sUserlname
 lcl_search_criteria = lcl_search_criteria & "&membershipid=" & iMembershipId
 lcl_search_criteria = lcl_search_criteria & "&periodid="     & iperiodid
 lcl_search_criteria = lcl_search_criteria & "&success=SD"

'Delete from egov_poolpassmembers
	sSQL = "DELETE FROM egov_poolpassmembers WHERE poolpassid = " & request("passid")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSQL, Application("DSN"), 3, 1

'Delete from egov_poolpasspurchases
 sSQL2 = "DELETE FROM egov_poolpasspurchases where poolpassid = " & request("passid")
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.Open sSQL2, Application("DSN"), 3, 1

 set rs  = nothing
 set rs2 = nothing

	response.redirect( "poolpass_list.asp" & lcl_search_criteria )
%>