<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!--#include file="../include_top_functions.asp"-->
<%
 dim sSQLd

 lcl_userid           = 0
 lcl_orgid            = 0
 lcl_dm_typeid        = 0
 lcl_dmid             = 0
 lcl_categoryid       = 0
 lcl_sub_categoryid   = ""
 lcl_subcategoryids   = ""
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

 if request("subcategoryids") <> "" then
    if not containsApostrophe(request("subcategoryids")) then
       lcl_subcategoryids = request("subcategoryids")
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
 if lcl_useraction = "DISPLAY" then
    if lcl_categoryid > 0 then
       displaySubCategories lcl_dmid, lcl_categoryid
    end if
 elseif lcl_useraction = "SEARCH" then
    displaySubCategoriesSearchOptions lcl_orgid, lcl_dm_typeid, lcl_categoryid, lcl_subcategoryids
 else 'ADD/ASSIGN/UNASSIGN
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
             sSQLd = "SELECT distinct categoryid "
             sSQLd = sSQLd & " FROM egov_dm_categories "
             sSQLd = sSQLd & " WHERE orgid = " & lcl_orgid
             sSQLd = sSQLd & " AND categoryid IN (" & lcl_sub_categoryid & ") "

             set oGetSubCategories = Server.CreateObject("ADODB.Recordset")
             oGetSubCategories.Open sSQLd, Application("DSN"), 3, 1

             if not oGetSubCategories.eof then
                do while not oGetSubCategories.eof

                   addSubCategoryAssignment lcl_orgid, lcl_dm_typeid, lcl_dmid, oGetSubCategories("categoryid")

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
 end if

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
  sSQL = sSQL & " dtc.dmid_categoryid, "
  sSQL = sSQL & " c.isApproved, "
  sSQL = sSQL & " isnull(c.approvedeniedbydate,'') as approvedeniedbydate "
  sSQL = sSQL & " FROM egov_dm_categories c "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dmdata_to_dmcategories dtc "
  sSQL = sSQL &                      " ON c.categoryid = dtc.categoryid "
  sSQL = sSQL &                      " AND dtc.dmid = " & sDMID
  sSQL = sSQL & " WHERE c.parent_categoryid = " & sParentCategoryID
  sSQL = sSQL & " AND c.isActive = 1 "
  sSQL = sSQL & " AND c.isApproved = 1 "
  sSQL = sSQL & " UNION ALL "
  sSQL = sSQL & " SELECT c.categoryid, "
  sSQL = sSQL & " c.categoryname, "
  sSQL = sSQL & " dtc.dmid_categoryid, "
  sSQL = sSQL & " c.isApproved, "
  sSQL = sSQL & " isnull(c.approvedeniedbydate,'') as approvedeniedbydate "
  sSQL = sSQL & " FROM egov_dm_categories c "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dmdata_to_dmcategories dtc "
  sSQL = sSQL &                      " ON c.categoryid = dtc.categoryid "
  sSQL = sSQL &                      " AND dtc.dmid = " & sDMID
  sSQL = sSQL & " WHERE c.parent_categoryid = " & sParentCategoryID
  sSQL = sSQL & " AND c.isActive = 1 "
  sSQL = sSQL & " AND c.isApproved = 0 "
  sSQL = sSQL & " AND c.approvedeniedbydate is null "
  sSQL = sSQL & " ORDER BY c.categoryname "

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

        if iColumnCount > 3 then
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

       'Determine if the sub-category is still waiting for approval
        lcl_subcat_waitingapproval = ""
        lcl_isApproved             = oDisplaySubCategoryOptions("isApproved")
        lcl_approvedeniedbydate    = oDisplaySubCategoryOptions("approvedeniedbydate")

       '*** NOTE: NULL values for dates in SQL Server are set to '1/1/1900'
        if not lcl_isApproved AND lcl_approvedeniedbydate = "1/1/1900" then
           lcl_subcat_waitingapproval = "<span class=""redText"">[waiting for approval]</span>"
        end if

        response.write "      <td id=""subcategorycell" & iFieldCount & """ class=""subCategoryCell"">" & vbcrlf
        response.write "          <input type=""checkbox"" name=""subcategoryid" & iFieldCount & """ id=""subcategoryid" & iFieldCount & """ value=""" & oDisplaySubCategoryOptions("categoryid") & """" & lcl_checked_subcategory & " /><span id=""subcategoryname" & iFieldCount & """>" & oDisplaySubCategoryOptions("categoryname") & "</span>" & lcl_subcat_waitingapproval & vbcrlf
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
sub displaySubCategoriesSearchOptions(iOrgID, iDMTypeID, iParentCategoryID, iSelectedSubCategoryIDs)

  sOrgID                  = 0
  sDMTypID                = 0
  sParentCategoryID       = 0
  sSelectedSubCategoryIDs = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iParentCategoryID <> "" then
     sParentCategoryID = clng(iParentCategoryID)
  end if

  if iSelectedSubCategoryIDs <> "" then
     sSelectedSubCategoryIDs = iSelectedSubCategoryIDs
  end if

  sSQL = "SELECT distinct "
  sSQL = sSQL & " c.categoryid, "
  sSQL = sSQL & " c.categoryname, "
  sSQL = sSQL & " c.isApproved, "
  sSQL = sSQL & " isnull(c.approvedeniedbydate,'') as approvedeniedbydate " 
  'sSQL = sSQL & " dtc.dmid_categoryid "
  sSQL = sSQL & " FROM egov_dm_categories c "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dmdata_to_dmcategories dtc "
  sSQL = sSQL &                      " ON c.categoryid = dtc.categoryid "
  sSQL = sSQL &                   " AND dtc.dmid IN (select distinct dmid "
  sSQL = sSQL &                                    " from egov_dm_data "
  sSQL = sSQL &                                    " where dm_typeid = " & sDMTypeID
  sSQL = sSQL &                                    " and isActive = 1 "
  sSQL = sSQL &                                    " and isApproved = 1 ) "
  sSQL = sSQL & " WHERE c.isActive = 1 "
'  sSQL = sSQL & " AND (c.isApproved = 1 OR c.isApprovedByDate = '') "
  sSQL = ssQL & " AND c.parent_categoryid > 0 "
  sSQL = sSQL & " AND c.dm_typeid = " & sDMTypeID

  if sParentCategoryID > 0 then
     sSQL = sSQL & " AND c.parent_categoryid = " & sParentCategoryID
  end if

  sSQL = sSQL & " ORDER BY c.categoryname, c.categoryid "
'dtb_debug(sSQL)
  set oDisplaySubCatSearchOpts = Server.CreateObject("ADODB.Recordset")
  oDisplaySubCatSearchOpts.Open sSQL, Application("DSN"), 3, 1

  if not oDisplaySubCatSearchOpts.eof then
     iRowCount    = 0
     iColumnCount = 0
     iFieldCount  = 0

     response.write "<input type=""hidden"" name=""subcategory_parent_categoryid"" id=""subcategory_parent_categoryid"" value=""" & iParentCategoryID & """ />" & vbcrlf
     response.write "<table id=""subCategoriesList"">" & vbcrlf

     do while not oDisplaySubCatSearchOpts.eof
        iRowCount               = iRowCount    + 1
        iColumnCount            = iColumnCount + 1
        iFieldCount             = iFieldCount  + 1
        lcl_checked_subcategory = ""
        lcl_isApproved          = oDisplaySubCatSearchOpts("isApproved")
        lcl_approvedeniedbydate = oDisplaySubCatSearchOpts("approvedeniedbydate")
        lcl_showOption          = true

        if not lcl_isApproved AND lcl_approvedeniedbydate = "1/1/1900" then
           lcl_showOption = false
        end if

        if lcl_showOption then
           if iColumnCount > 5 then
              iRowCount    = 1
              iColumnCount = 1
              response.write "  </tr>" & vbcrlf
           end if

           if iRowCount = 1 then
              response.write "  <tr id=""subcategoryrow" & iFieldCount & """>" & vbcrlf
           end if

           'if oDisplaySubCatSearchOpts("dmid_categoryid") <> "" then
           '   lcl_checked_subcategory = " checked=""checked"""
           'end if
           if isSubCategorySelected(sOrgID, oDisplaySubCatSearchOpts("categoryid"), sSelectedSubCategoryIDs) then
              lcl_checked_subcategory = " checked=""checked"""
           end if

           response.write "      <td id=""subcategorycell" & iFieldCount & """ class=""subCategoryCell"">" & vbcrlf
           response.write "          <input type=""checkbox"" name=""subcategoryid" & iFieldCount & """ id=""subcategoryid"   & iFieldCount & """ value=""" & oDisplaySubCatSearchOpts("categoryid")   & """" & lcl_checked_subcategory & " /><span id=""subcategoryname" & iFieldCount & """>" & oDisplaySubCatSearchOpts("categoryname") & "</span>" & vbcrlf
           response.write "          <input type=""hidden"" name=""subcategoryname"   & iFieldCount & """ id=""subcategoryname" & iFieldCount & """ value=""" & oDisplaySubCatSearchOpts("categoryname") & """ />" & vbcrlf
           response.write "          <input type=""hidden"" name=""isNewSubCategory" & iFieldCount & """ id=""isNewSubCategory" & iFieldCount & """ value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
        end if

        oDisplaySubCatSearchOpts.movenext
     loop

     if iColumnCount > 1 then
        lcl_colspan = (3 - iColumnCount)
        response.write "    <td colspan=""" & lcl_colspan & """>&nbsp;</td>" & vbcrlf
     end if

     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "<input type=""hidden"" name=""total_subcategories"" id=""total_subcategories"" value=""" & iFieldCount & """ />" & vbcrlf
     response.write "<input type=""hidden"" name=""total_subcategories_new"" id=""total_subcategories_new"" value=""0"" />" & vbcrlf
     'response.write "<div id=""subCategoryAddRow"">" & vbcrlf
     'response.write "  Other: <input type=""text"" name=""subcategory_add"" id=""subcategory_add"" value="""" size=""20"" maxlength=""100"" onchange=""clearMsg('subCategoryAddButton');"" />" & vbcrlf
     'response.write "  <input type=""button"" name=""subCategoryAddButton"" id=""subCategoryAddButton"" class=""button"" value=""Add"" onclick=""addSubCategory();"" />" & vbcrlf
     'response.write "  <img src=""../images/help.jpg"" name=""helpFeature_addSubCategory"" id=""helpFeature_addSubCategory"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
     'response.write "  <div name=""helpFeature_addSubCategory_text"" id=""helpFeature_addSubCategory_text"" class=""helpOptionText"">" & vbcrlf
     'response.write "    <p><strong>E-GOV TIP:</strong><br />Clicking on the ""add"" button will add the sub-category but NOT automatically assign it.</p>" & vbcrlf
     'response.write "  </div>" & vbcrlf
     'response.write "</div>" & vbcrlf

  end if

  oDisplaySubCatSearchOpts.close
  set oDisplaySubCatSearchOpts = nothing

end sub

'------------------------------------------------------------------------------
function isSubCategorySelected(iOrgID, iSubCategoryID, iSelectedSubCategories)
  lcl_return = false

  sOrgID                 = 0
  sSubCategoryID         = 0
  sSelectedSubCategories = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iSubCategoryID <> "" then
     sSubCategoryID = clng(iSubCategoryID)
  end if

  if iSelectedSubCategories <> "" then
     if not containsApostrophe(iSelectedSubCategories) then
        sSelectedSubCategories = iSelectedSubCategories
     end if
  end if

  if sSubCategoryID > 0 AND sSelectedSubCategories <> "" then
     sSQL = "SELECT distinct categoryid "
     sSQL = sSQL & " FROM egov_dm_categories "
     sSQL = sSQL & " WHERE orgid = " & sOrgID
     sSQL = sSQL & " AND categoryid IN (" & sSelectedSubCategories & ") "
     sSQL = ssQL & " AND categoryid = " & sSubCategoryID

     set oIsSubCatSelected = Server.CreateObject("ADODB.Recordset")
     oIsSubCatSelected.Open sSQL, Application("DSN"), 3, 1

     if not oIsSubCatSelected.eof then
        lcl_return = true
     end if

     oIsSubCatSelected.close
     set oIsSubCatSelected = nothing
  end if

  isSubCategorySelected = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>