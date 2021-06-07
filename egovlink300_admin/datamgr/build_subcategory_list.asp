<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
 lcl_userid           = 0
 lcl_orgid            = 0
 lcl_dm_typeid        = 0
 lcl_dmid             = 0
 lcl_categoryid       = 0
 lcl_sub_categoryid   = ""
 lcl_sub_mergeCatInto = ""
 lcl_sub_categoryname = ""
 lcl_sub_delete       = ""
 lcl_sub_isActive     = "Y"
 lcl_useraction       = "DISPLAY"

 if request("userid") <> "" then
    if isnumeric(request("userid")) then
       lcl_userid = request("userid")
    end if
 end if

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("dm_typeid") <> "" then
    if isnumeric(request("dm_typeid")) then
       lcl_dm_typeid = request("dm_typeid")
    end if
 end if

 if request("dmid") <> "" then
    if isnumeric(request("dmid")) then
       lcl_dmid = request("dmid")
    end if
 end if

 if request("categoryid") <> "" then
    if isnumeric(request("categoryid")) then
       lcl_categoryid = request("categoryid")
    end if
 end if

 if request("sub_categoryid") <> "" then
    if not containsApostrophe(request("sub_categoryid")) then
       lcl_sub_categoryid = request("sub_categoryid")
    else
       sLevel = "../"  'Override of value from common.asp
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 if request("sub_categoryname") <> "" then
    lcl_sub_categoryname = request("sub_categoryname")
 end if

 if request("useraction") <> "" then
    if not containsApostrophe(request("useraction")) then
       lcl_useraction = ucase(request("useraction"))
    else
       sLevel = "../"  'Override of value from common.asp
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

'Determine if we are:
'  1. Displaying a list of sub-categories.
'  2. Adding sub-categories.
'  3. Maintain sub-category assignments.
 if lcl_useraction <> "DISPLAY" then

    if lcl_useraction = "ADD" then
       lcl_assignSubCategory = true

       lcl_sub_categoryid = maintainSubCategory(lcl_orgid, lcl_dm_typeid, lcl_dmid, lcl_userid, _
                                                lcl_sub_delete, lcl_sub_mergeCatInto, lcl_sub_categoryid, _
                                                lcl_sub_categoryname, lcl_sub_isActive, lcl_categoryid, _
                                                lcl_assignSubCategory)

       if lcl_sub_categoryid <> "already exists" then
          if clng(lcl_sub_categoryid) > 0 then
             if lcl_isAjax = "Y" then
                response.write lcl_sub_categoryid
             end if
          end if
       end if
    else
      'Determine if we need to assign or unassign sub-category ids from a specific DM Data record (dmid).
      'lcl_useraction values: ASSIGN or UNASSIGN
       if lcl_sub_categoryid <> "" then
          if lcl_useraction = "ASSIGN" then

            'Remove all of the sub-category assignments from the dmid
             deleteSubCategoryAssignments lcl_dmid, lcl_sub_categoryid

            'Loop through the sub-category ids, even if there's only one, to create the assignments
             sSQL = "SELECT distinct categoryid "
             sSQL = sSQL & " FROM egov_dm_categories "
             sSQL = sSQL & " WHERE orgid = " & lcl_orgid
             sSQL = sSQL & " AND categoryid IN (" & lcl_sub_categoryid & ") "

             set oGetSubCategories = Server.CreateObject("ADODB.Recordset")
             oGetSubCategories.Open sSQL, Application("DSN"), 3, 1

             if not oGetSubCategories.eof then
                do while not oGetSubCategories.eof

                   lcl_dm_importid = ""

                   addSubCategoryAssignment lcl_orgid, _
                                            lcl_dm_typeid, _
                                            lcl_dmid, _
                                            oGetSubCategories("categoryid"), _
                                            lcl_dm_importid

                   oGetSubCategories.movenext
                loop
             end if

             oGetSubCategories.close
             set oGetSubCategories = nothing
          else
             deleteSubCategoryAssignments lcl_dmid, lcl_sub_categoryid
          end if
       end if
    end if
 else  'useraction = "DISPLAY"
    if lcl_categoryid > 0 then
       displaySubCategories lcl_dmid, lcl_categoryid
    end if
 end if

 'else
 '   if lcl_isAjax = "Y" then
 '      response.write "Failed to display list - Error in AJAX Routine"
 '   else
 '      response.write "datamgr_maint.asp?dmid=" & lcl_dmid & "&success=AJAX_ERROR"
 '   end if
 'end if

'------------------------------------------------------------------------------
sub displaySubCategories(iDMID, iParentCategoryID)

  sDMID             = 0
  sParentCategoryID = 0

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iParentCategoryID <> "" then
     sParentCategoryID = clng(iParentCategoryID)
  end if

  response.write "<input type=""hidden"" name=""subcategory_parent_categoryid"" id=""subcategory_parent_categoryid"" value=""" & sParentCategoryID & """ />" & vbcrlf
  response.write "<table id=""subCategoriesList"">" & vbcrlf

  sSQL = "SELECT c.categoryid, "
  sSQL = sSQL & " c.categoryname, "
  sSQL = sSQL & " dtc.dmid_categoryid "
  sSQL = sSQL & " FROM egov_dm_categories c "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dmdata_to_dmcategories dtc "
  sSQL = sSQL &                      " ON c.categoryid = dtc.categoryid "
  sSQL = sSQL &                      " AND dtc.dmid = " & sDMID
  sSQL = sSQL & " WHERE c.parent_categoryid = " & sParentCategoryID
  sSQL = sSQL & " AND c.isActive = 1 "
  sSQL = sSQL & " AND c.isApproved = 1 "
  sSQL = sSQL & " ORDER BY UPPER(c.categoryname) "

  set oDisplaySubCategoryOptions = Server.CreateObject("ADODB.Recordset")
  oDisplaySubCategoryOptions.Open sSQL, Application("DSN"), 3, 1

  if not oDisplaySubCategoryOptions.eof then
     iRowCount    = 0
     iColumnCount = 0
     iFieldCount  = 0

     do while not oDisplaySubCategoryOptions.eof
        iRowCount               = iRowCount    + 1
        iColumnCount            = iColumnCount + 1
        iFieldCount             = iFieldCount  + 1
        lcl_checked_subcategory = ""

        if iColumnCount > 5 then
           iRowCount    = 1
           iColumnCount = 1
           response.write "  </tr>" & vbcrlf
        end if

        if iRowCount = 1 then
           response.write "  <tr id=""subcategoryrow" & iFieldCount & """>" & vbcrlf
        end if

        if oDisplaySubCategoryOptions("dmid_categoryid") <> "" then
           lcl_checked_subcategory = " checked=""checked"""
        end if

        response.write "      <td id=""subcategorycell" & iFieldCount & """ class=""subCategoryCell"">" & vbcrlf
        response.write "          <input type=""checkbox"" name=""subcategoryid" & iFieldCount & """ id=""subcategoryid" & iFieldCount & """ value=""" & oDisplaySubCategoryOptions("categoryid") & """" & lcl_checked_subcategory & " /><span id=""subcategoryname" & iFieldCount & """>" & oDisplaySubCategoryOptions("categoryname") & "</span>" & vbcrlf
        response.write "          <input type=""hidden"" name=""isNewSubCategory" & iFieldCount & """ id=""isNewSubCategory" & iFieldCount & """ value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
        response.write "      </td>" & vbcrlf

        oDisplaySubCategoryOptions.movenext
     loop

     if iColumnCount > 1 then
        lcl_colspan = (3 - iColumnCount)
        response.write "    <td colspan=""" & lcl_colspan & """>&nbsp;</td>" & vbcrlf
     end if

     response.write "  </tr>" & vbcrlf
     'response.write "<div id=""subCategoryAddRow"">" & vbcrlf
     'response.write "  Other: <input type=""text"" name=""subcategory_add"" id=""subcategory_add"" value="""" size=""20"" maxlength=""100"" onchange=""clearMsg('subCategoryAddButton');"" />" & vbcrlf
     'response.write "  <input type=""button"" name=""subCategoryAddButton"" id=""subCategoryAddButton"" class=""button"" value=""Add"" onclick=""addSubCategory();"" />" & vbcrlf
     'response.write "  <img src=""../images/help.jpg"" name=""helpFeature_addSubCategory"" id=""helpFeature_addSubCategory"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
     'response.write "  <div name=""helpFeature_addSubCategory_text"" id=""helpFeature_addSubCategory_text"" class=""helpOptionText"">" & vbcrlf
     'response.write "    <p><strong>E-GOV TIP:</strong><br />Clicking on the ""add"" button will add the sub-category but NOT automatically assign it.</p>" & vbcrlf
     'response.write "  </div>" & vbcrlf
     'response.write "</div>" & vbcrlf

  end if

  response.write "</table>" & vbcrlf
  response.write "<input type=""hidden"" name=""total_subcategories"" id=""total_subcategories"" value=""" & iFieldCount & """ />" & vbcrlf
  response.write "<input type=""hidden"" name=""total_subcategories_new"" id=""total_subcategories_new"" value=""0"" />" & vbcrlf

  oDisplaySubCategoryOptions.close
  set oDisplaySubCategoryOptions = nothing

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