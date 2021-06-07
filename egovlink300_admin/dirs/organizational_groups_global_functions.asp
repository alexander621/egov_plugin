<%
'-----------------------------------------------------------------------------
function check_for_sub_org_groups(p_parent_org_group_id)

  if p_parent_org_group_id <> 0 then
     sSQL = "SELECT distinct 'Y' AS lcl_exists " 
     sSQL = sSQL & " FROM egov_staff_directory_groups "
     sSQL = sSQL & " WHERE parent_org_group_id = " & p_parent_org_group_id

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSQL, Application("DSN"), 3, 1

     if not rs.eof then
        lcl_exists = rs("lcl_exists")
     else
        lcl_exists = "N"
     end if
  else
     lcl_exists = "N"
  end if

  rs.close
  set rs = nothing

  check_for_sub_org_groups = lcl_exists

end function

'------------------------------------------------------------------------------
sub update_sub_org_levels(p_parent_org_group_id, p_org_level)
  lcl_org_level = p_org_level + 1

 'Update the org_levels for the org_groups with this parent_org_group_id
  sSQL2 = "UPDATE egov_staff_directory_groups SET "
  sSQL2 = sSQL2 & " org_level = " & lcl_org_level
  sSQL2 = sSQL2 & " WHERE parent_org_group_id = " & p_parent_org_group_id

  set rs2 = Server.CreateObject("ADODB.Recordset")
  rs2.Open sSQL2, Application("DSN"), 3, 1

  sSQL = "SELECT org_group_id, "
  sSQL = sSQL & " org_name, "
  sSQL = sSQL & " org_level "
  sSQL = sSQL & " FROM egov_staff_directory_groups "
  sSQL = sSQL & " WHERE parent_org_group_id = " & p_parent_org_group_id
  sSQL = sSQL & " ORDER BY org_level, UPPER(org_name) "

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     do while not rs.eof

       if check_for_sub_org_groups(rs("org_group_id")) = "Y" then
          update_sub_org_levels rs("org_group_id"), _
                                rs("org_level")
       end if

       rs.movenext
    loop
  end if

  rs.close

  set rs  = nothing
  set rs2 = nothing

end sub
%>