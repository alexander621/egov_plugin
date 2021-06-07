<%
'------------------------------------------------------------------------------
function getEmailFormatDesc(iEmailFormat)

  lcl_return = ""

  if iEmailFormat <> "" then
     lcl_emailformat = CStr(iEmailFormat)

     if lcl_emailformat = "1" then
        lcl_return = "Plain Text Only"
     elseif lcl_emailformat = "2" then
        lcl_return = "HTML And Plain Text"
     elseif lcl_emailformat = "3" then
        lcl_return = "HTML Only"
     end if
  end if

  getEmailFormatDesc = lcl_return

end function

'------------------------------------------------------------------------------
function getDistributionListNames(iOrgID, iDLListIDs)

  lcl_return    = ""
  lcl_listnames = ""

  if iDLListIDs <> "" then
     sSQL = "SELECT distributionlistname "
     sSQL = sSQL & " FROM egov_class_distributionlist "
     sSQL = sSQL & " WHERE orgid = " & iOrgID
     sSQL = sSQL & " AND distributionlistid IN (" & iDLListIDs & ") "

    	set oGetDLNames = Server.CreateObject("ADODB.Recordset")
     oGetDLNames.Open sSQL, Application("DSN"), 3, 1

     if not oGetDLNames.eof then
        do while not oGetDLNames.eof

           if lcl_listnames <> "" then
              lcl_listnames = lcl_listnames & ", " & oGetDLNames("distributionlistname")
           else
              lcl_listnames = oGetDLNames("distributionlistname")
           end if

           oGetDLNames.movenext
        loop
     end if

     oGetDLNames.close
     set oGetDLNames = nothing

     lcl_return = lcl_listnames

  end if

  getDistributionListNames = lcl_return

end function
%>