<!-- #include file="../includes/common.asp" //-->
<%
 lcl_orgid                  = 0
 lcl_eventScrollerDateLimit = 14
 lcl_isAjax                 = "Y"

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("eventScrollerDateLimit") <> "" then
    if isnumeric(request("eventScrollerDateLimit")) then
       lcl_eventScrollerDateLimit = request("eventScrollerDateLimit")
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = ucase(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_orgid > 0 AND lcl_eventScrollerDateLimit <> "" then
    updateEventScrollerDateLimit lcl_orgid, _
                                 lcl_eventScrollerDateLimit
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update section order - Error in AJAX Routine"
    else
       response.write "default.asp?success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub updateEventScrollerDateLimit(iOrgID, iEventScrollerDateLimit)
  dim sSQL, lcl_success
  dim sOrgID, sEventScrollerDateLimit

  sSQL                    = ""
  lcl_success             = ""
  sOrgID                  = 0
  sEventScrollerDateLimit = 14

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = iOrgID
     end if
  end if

  if iEventScrollerDateLimit <> "" then
     if isnumeric(iEventScrollerDateLimit) then
        sEventScrollerDateLimit = iEventScrollerDateLimit
     end if
  end if

  if sOrgID > 0 then
     sSQL = "UPDATE organizations SET "
     sSQL = sSQL & "eventScrollerDateLimit = " & sEventScrollerDateLimit
     sSQL = sSQL & " WHERE orgid = " & sOrgID

     lcl_success = "Successfully Updated..."
  end if

  if sSQL <> "" then
    	set oUpdateEventScrollerDateLimit = Server.CreateObject("ADODB.Recordset")
   	 oUpdateEventScrollerDateLimit.Open sSQL, Application("DSN"), 3, 1

     set oUpdateEventScrollerDateLimit = nothing
  end if

  response.write lcl_success

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>