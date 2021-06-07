<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: facility_global_functions.asp
' AUTHOR: David Boyer ?
' CREATED: ???
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module processes payments.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	??? ??? - Initial Version
' 1.2	9/25/2009	Steve Loar - Added the session timeout check.
'
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
Function getSuccessMessage( ByVal p_success )
  Dim lcl_return
  lcl_return = ""

  If p_success <> "" then
     If UCase(p_success) = "NA" Then 
        lcl_return = "<strong style=""color:#FF0000"">*** This facility is no longer available to reserve ***</strong>"
     End If 
	 If UCase(p_success) = "TIMEOUT" Then 
        lcl_return = "<div style=""border:1px solid red; text-align:center;""><span style=""color:#FF0000; font-size:10pt;"">*** We're sorry but, your session has timed out. You must start your reservation again. ***</span><br /><span style=""color:#FF0000; font-size:11pt; font-weight:bold;"">*** You have not been charged for this transaction. ***</span></div>"
     End If 
  End If 

  getSuccessMessage = lcl_return

End Function 
'------------------------------------------------------------------------------
function isFacilityAvail(lcl_facilityscheduleid, lcl_checkindate, lcl_checkoutdate, lcl_timepartid, lcl_facilityid, lcl_current_date)
  lcl_return     = False
  lcl_total_cnt1 = 0
  lcl_total_cnt2 = 0
  lcl_userid     = 0
  lcl_sessionid  = ""

 'Determine if the facility is/isn't available.
 '1. If a facilityscheduleid is passed in then pull all of the data for the facility.
 '2. If no facilityscheduleid is passed in then validate the data passed in.
 '3. Now check to see if a record exists for the facility, org, and date AND the "result" and "status" columns are NULL.
 '4. EXTRA: if userid is passed in and it matches the userid on an existing facility record then:
           '- if all of the fields match
           '- AND
           '- if a sessionid is passed in and it matches the sessionid on the facility record then the facility is available
           '- OR
           '- allow the facility to be accessed
           '- update the facility record with the new sessionid.

  if lcl_facilityscheduleid <> "" then
     lcl_checkindate  = ""
     lcl_checkoutdate = ""
     lcl_timepartid   = 0
     lcl_facilityid   = 0
     lcl_result       = ""
     lcl_status       = ""

     if dbready_number(lcl_facilityscheduleid) then
        sSQL1 = "SELECT checkindate, checkoutdate, facilitytimepartid, facilityid, result, status, lesseeid, sessionid "
        sSQL1 = sSQL1 & " FROM egov_facilityschedule "
        sSQL1 = sSQL1 & " WHERE facilityscheduleid = " & lcl_facilityscheduleid

      	set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open sSQL1, Application("DSN"), 3, 1

        if not rs1.eof then
           lcl_checkindate  = rs1("checkindate")
           lcl_checkoutdate = rs1("checkoutdate")
           lcl_timepartid   = rs1("facilitytimepartid")
           lcl_facilityid   = rs1("facilityid")
           lcl_result       = rs1("result")
           lcl_status       = rs1("status")
           lcl_userid       = rs1("lesseeid")
           lcl_sessionid    = rs1("sessionid")
        end if

        set rs1 = nothing

     end if
  else
    'Validate the data
     lcl_checkindate  = dbready_string(lcl_checkindate,50)
     lcl_checkoutdate = dbready_string(lcl_checkoutdate,50)

     if NOT dbready_number(lcl_timepartid) then
        lcl_timepartid = 0
     end if

     if NOT dbready_number(lcl_facilityid) then
        lcl_facilityid = 0
     end if

  end if

 'Check for an existing facility record
  sSQL = "SELECT count(facilityscheduleid) AS total_cnt1, 0 AS total_cnt2 "
  sSQL = sSQL & " FROM egov_facilityschedule "
  sSQL = sSQL & " WHERE checkindate = '"     & lcl_checkindate  & "' "
  sSQL = sSQL & " AND checkoutdate = '"      & lcl_checkoutdate & "' "
  sSQL = sSQL & " AND facilitytimepartid = " & lcl_timepartid
  sSQL = sSQL & " AND facilityid = "         & lcl_facilityid
  sSQL = sSQL & " AND orgid = "              & iorgid
  sSQL = sSQL & " AND checkindate = '"       & lcl_current_date & "' "
  sSQL = sSQL & " AND (result = '' OR result is null) "
  sSQL = sSQL & " AND (status = '' OR status is null) "

  sSQL = sSQL & " UNION ALL "

  sSQL = sSQL & " SELECT 0 AS total_cnt1, count(facilityscheduleid) AS total_cnt2 "
  sSQL = sSQL & " FROM egov_facilityschedule "
  sSQL = sSQL & " WHERE checkindate = '"     & lcl_checkindate  & "' "
  sSQL = sSQL & " AND checkoutdate = '"      & lcl_checkoutdate & "' "
  sSQL = sSQL & " AND facilitytimepartid = " & lcl_timepartid
  sSQL = sSQL & " AND facilityid = "         & lcl_facilityid
  sSQL = sSQL & " AND orgid = "              & iorgid
  sSQL = sSQL & " AND checkindate = '"       & lcl_current_date & "' "
  sSQL = sSQL & " AND result = 'APPROVED' "
  sSQL = sSQL & " AND status IN ('CANCELLED') "

  set oCheck = Server.CreateObject("ADODB.Recordset")
  oCheck.Open sSQL, Application("DSN"), 3, 1

  if oCheck("total_cnt1") = 0 AND oCheck("total_cnt2") >= 0 then
     lcl_return = True
  else
    'If a record exists we then need to retrieve all of the data into local variables.
    'This ONLY has to be done if a "lcl_facilityscheduleid" has NOT been passed in.
     if lcl_facilityscheduleid = "" then
        sSQL1 = "SELECT checkindate, checkoutdate, facilitytimepartid, facilityid, result, status, lesseeid, sessionid "
        sSQL1 = sSQL1 & " FROM egov_facilityschedule "
        sSQL1 = sSQL1 & " WHERE checkindate = '"     & lcl_checkindate  & "' "
        sSQL1 = sSQL1 & " AND checkoutdate = '"      & lcl_checkoutdate & "' "
        sSQL1 = sSQL1 & " AND facilitytimepartid = " & lcl_timepartid
        sSQL1 = sSQL1 & " AND facilityid = "         & lcl_facilityid
        sSQL1 = sSQL1 & " AND orgid = "              & iorgid
        sSQL1 = sSQL1 & " AND checkindate = '"       & lcl_current_date & "' "
        sSQL1 = sSQL1 & " AND (result = '' OR result is null) "
        sSQL1 = sSQL1 & " AND (status = '' OR status is null) "

      	 set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open sSQL1, Application("DSN"), 3, 1

        if not rs1.eof then
           lcl_checkindate  = rs1("checkindate")
           lcl_checkoutdate = rs1("checkoutdate")
           lcl_timepartid   = rs1("facilitytimepartid")
           lcl_facilityid   = rs1("facilityid")
           lcl_result       = rs1("result")
           lcl_status       = rs1("status")
           lcl_userid       = rs1("lesseeid")
           lcl_sessionid    = rs1("sessionid")
        end if

        set rs1 = nothing
     end if

    'Try to find the userid from one of the many variables it could be stored in.
     if request("lesseeid") <> "" then
        lcl_lesseeid = request("lesseeid")
     else
        if request("iuserid") <> "" then
           lcl_lesseeid = request("iuserid")
        else
           lcl_lesseeid = request.cookies("userid")
        end if
     end if

    'Now check to see if the userid matches the record.
    'if so then allow access AND update the record with the new sessionid
     'if request.cookies("userid") <> "" then
     if lcl_lesseeid <> "" then
        'if CLng(lcl_userid) = CLng(request.cookies("userid")) then
        if CLng(lcl_userid) = CLng(lcl_lesseeid) then
           lcl_return = True

           if lcl_sessionid <> "" AND lcl_sessionid <> session.sessionid AND lcl_current_date <> "" then
              lcl_facilityscheduleid = updateFacilitySessionID(lcl_facilityscheduleid, lcl_checkindate, lcl_checkoutdate, lcl_timepartid, lcl_facilityid, lcl_userid, lcl_sessionid, session.sessionid, lcl_current_date)
           end if
        end if
     end if

  end if

  set oCheck = nothing

  isFacilityAvail = lcl_return

end function

'------------------------------------------------------------------------------
function updateFacilitySessionID(lcl_facilityscheduleid, lcl_checkindate, lcl_checkoutdate, lcl_timepartid, lcl_facilityid, lcl_userid, lcl_sessionid, lcl_new_sessionid, lcl_currentdate)
  lcl_return       = 0
  lcl_where_clause = ""

  if NOT dbready_date(lcl_currentdate) then
     lcl_currentdate = ""
  end if

  if NOT dbready_number(lcl_userid) then
     iuserid = 0
  else
     iuserid = lcl_userid
  end if

  if dbready_number(lcl_new_sessionid) then
     sSQLu = "UPDATE egov_facilityschedule SET sessionid = " & lcl_new_sessionid

     if lcl_facilityscheduleid <> "" then
        lcl_where_clause = lcl_where_clause & " WHERE facilityscheduleid = " & lcl_facilityscheduleid
     else
        lcl_where_clause = lcl_where_clause & " WHERE orgid = "            & iorgid
        lcl_where_clause = lcl_where_clause & " AND checkindate = '"       & lcl_checkindate  & "' "
        lcl_where_clause = lcl_where_clause & " AND checkoutdate = '"      & lcl_checkoutdate & "' "
        lcl_where_clause = lcl_where_clause & " AND facilitytimepartid = " & lcl_timepartid
        lcl_where_clause = lcl_where_clause & " AND facilityid = "         & lcl_facilityid
        lcl_where_clause = lcl_where_clause & " AND lesseeid = "           & iuserid
        lcl_where_clause = lcl_where_clause & " AND sessionid = "          & lcl_sessionid
     end if 

     sSQLu = sSQLu & lcl_where_clause
     sSQLu = sSQLu & " AND checkindate = '" & lcl_currentdate & "'"
 
     set oUpdate = Server.CreateObject("ADODB.Recordset")
     oUpdate.Open sSQLu, Application("DSN"), 3, 1

    'Retreive the facilityscheduleid for the record just updated.
     sSQL2 = "SELECT isnull(max(facilityscheduleid),0) as max_id "
     sSQL2 = sSQL2 & " FROM egov_facilityschedule "
     sSQL2 = sSQL2 & lcl_where_clause

     set rs2 = Server.CreateObject("ADODB.Recordset")
     rs2.Open sSQL2, Application("DSN"), 3, 1

     if not rs2.eof then
        lcl_return = rs2("max_id")
     end if

     set oUpdate = nothing
     set rs2     = nothing
  end if

  updateFacilitySessionID = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  if p_value <> "" then
     sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
   	 set rsi = Server.CreateObject("ADODB.Recordset")
     rsi.Open sSQLi, Application("DSN"), 3, 1
  end if
end Sub


'--------------------------------------------------------------------------------------------------
' ReservationDataIsGone( iFacilityPaymentID )
'--------------------------------------------------------------------------------------------------
Function ReservationDataIsGone( iFacilityPaymentID )
	Dim oRs, sSql

	' See if the row has been deleted from the system due to a session timeout (global.asa will remove the row in this situation)
	sSql = "SELECT COUNT(facilityscheduleid) AS hits FROM egov_facilityschedule WHERE facilityscheduleid = " & iFacilityPaymentID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			ReservationDataIsGone = False 
		Else
			ReservationDataIsGone = True 
		End If 
	Else
		ReservationDataIsGone = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' getFacilityName iFacilityId
'--------------------------------------------------------------------------------------------------
Function getFacilityName( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT facilityname FROM egov_facility WHERE facilityid = " & iFacilityId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		getFacilityName = oRs("facilityname")
	Else
		getFacilityName = ""
	End If

	oRs.Close
	Set oRs = Nothing

End Function  




%>
