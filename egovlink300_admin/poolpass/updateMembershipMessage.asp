<!-- #include file="../includes/common.asp" //-->
<%
  lcl_orgid        = 0
  lcl_membershipid = 0
  lcl_showMessage  = "0"
  lcl_isAjax       = "Y"

  if request("orgid") <> "" then
     if isnumeric(request("orgid")) then
        lcl_orgid = clng(request("orgid"))
     end if
  end if

  if request("membershipid") <> "" then
     if isnumeric(request("membershipid")) then
        lcl_membershipid = clng(request("membershipid"))
     end if
  end if

  if request("action") <> "" then
     if not containsApostrophe(request("action")) then
        lcl_action = ucase(request("action"))
     end if
  end if

  if request("showMessage") <> "" then
     if not containsApostrophe(request("showMessage")) then
        if ucase(request("showMessage")) = "ON" then
           lcl_showMessage = "1"
        end if
     end if
  end if

  if request("isAjax") <> "" then
     if not containsApostrophe(request("isAjax")) then
        lcl_isAjax = ucase(request("isAjax"))
     end if
  end if

  if lcl_membershipid > 0 then
     sSQL = "UPDATE egov_memberships SET "
     sSQL = sSQL & " showMessage = " & lcl_showMessage
     sSQL = sSQL & " WHERE orgid = "      & lcl_orgid
     sSQL = sSQL & " AND membershipid = " & lcl_membershipid

     set oUpdateMembershipMessage = Server.CreateObject("ADODB.Recordset")
     oUpdateMembershipMessage.Open sSQL, Application("DSN"), 3, 1

     set oUpdateMembershipMessage = nothing

     if lcl_isAjax = "Y" then
        response.write "Successfully Updated"
     else
        response.redirect "poolpass_rates.asp?success=SU"
     end if

  else
     if lcl_isAjax = "Y" then
        response.write "Failed to update section order - Error in AJAX Routine"
     else
        response.redirect "poolpass_rates.asp?success=AJAX_ERROR"
     end if
  end if

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>