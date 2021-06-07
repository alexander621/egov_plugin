<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../action_line/action_line_global_functions.asp" //-->
<%

updateMobileGeoLoc request("orgid"), _
                   request("userid"), _
                   request("requestid"), _
                   request("latitude"), _
                   request("longitude")

'------------------------------------------------------------------------------
sub updateMobileGeoLoc(iOrgID, _
                       iUserID, _
                       iRequestID, _
                       iLatitude, _
                       iLongitude)

  dim lcl_success, sOrgID, sUserID, sRequestID, sLatitude, sLongitude
	
  lcl_success = ""
  sOrgID      = 0
  sUserID     = 0
  sRequestID  = 0
  sLatitude   = "NULL"
  sLongitude  = "NULL"

 	if iOrgID <> "" then
 		  sOrgID = CLng(iOrgID)
  end if

 	if iUserID <> "" then
 		  sUserID = CLng(iUserID)
  end if

 	if iRequestID <> "" then
 		  sRequestID = CLng(iRequestID)
  end if

  if iLatitude <> "" then
   		sLatitude = dbsafe(iLatitude)
     sLatitude = "'" & sLatitude  & "'"
  end if

  if iLongitude <> "" then
   		sLongitude = dbsafe(iLongitude)
     sLongitude = "'" & sLongitude & "'"
  end if

  if sRequestID > 0 then
     sSQL = "UPDATE egov_actionline_requests SET "
     sSQL = sSQL & "   mobileoption_latitude = "  & sLatitude
     sSQL = sSQL & " , mobileoption_longitude = " & sLongitude
     sSQL = sSQL & " WHERE action_autoid = " & sRequestID

   		set oUpdateMobileGeoLoc = Server.CreateObject("ADODB.Recordset")
	   	oUpdateMobileGeoLoc.Open sSQL, Application("DSN"), 3, 1

     set oUpdateMobileGeoLoc = nothing

     lcl_success = "Changes saved successfully..."

  end if

  response.write lcl_success

end sub
%>