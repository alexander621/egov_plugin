<!-- #include file="../includes/common.asp" //-->
<%
  lcl_orgid           = 0
  lcl_membershipid    = 0
  lcl_showTerms       = "0"
  lcl_membershipTerms = "NULL"
  lcl_action          = "EDIT_TERMS"
  lcl_isAjax          = "Y"

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

  if lcl_action = "SHOW_TERMS" then
     if request("showTerms") <> "" then
        if not containsApostrophe(request("showTerms")) then
           if ucase(request("showTerms")) = "ON" then
              lcl_showTerms = "1"
           end if
        end if
     end if
  else
     if request("membershipTerms") <> "" then
        lcl_membershipTerms = request("membershipTerms")
        lcl_membershipTerms = dbsafe(lcl_membershipTerms)
        lcl_membershipTerms = "'" & lcl_membershipTerms & "'"
     end if
  end if

  if request("isAjax") <> "" then
     if not containsApostrophe(request("isAjax")) then
        lcl_isAjax = ucase(request("isAjax"))
     end if
  end if

  if lcl_membershipid > 0 then
     sSQL = "UPDATE egov_memberships SET "

     if lcl_action = "SHOW_TERMS" then
        sSQL = sSQL & " showTerms = " & lcl_showTerms
     else
        sSQL = sSQL & " membershipterms = " & lcl_membershipTerms
     end if

     sSQL = sSQL & " WHERE orgid = "      & lcl_orgid
     sSQL = sSQL & " AND membershipid = " & lcl_membershipid

     set oUpdateMembershipTerms = Server.CreateObject("ADODB.Recordset")
     oUpdateMembershipTerms.Open sSQL, Application("DSN"), 3, 1

     set oUpdateMembershipTerms = nothing

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