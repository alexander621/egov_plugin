<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'Retrieve the categoryid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 dim lcl_categoryid

 if request("categoryid") <> "" then
    lcl_categoryid = request("categoryid")

    if not isnumeric(lcl_categoryid) then
       response.redirect "datamgr_categories_list.asp"
    end if
 else
    lcl_categoryid = 0
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature     = "datamgr_types_maint"
 lcl_dm_typeid   = 0

 if request("f") <> "" then
    lcl_feature = request("f")
 end if

'Retrieve the DM_TypeID
 if request("dm_typeid") <> "" then
    lcl_dm_typeid = request("dm_typeid")
 else
    lcl_dm_typeid = getDMTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

    if lcl_dm_typeid = 0 then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_categoryid = clng(lcl_categoryid)
 lcl_dm_typeid  = clng(lcl_dm_typeid)
 lcl_success    = request("success")

'Retrieve the search options
 lcl_sc_categoryname = ""

 if request("sc_categoryname") <> "" then
    lcl_sc_categoryname = request("sc_categoryname")
 end if

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid",       lcl_dm_typeid)
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_categoryname", lcl_sc_categoryname)

 if lcl_feature <> "datamgr_types_maint" then
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

 if request("user_action") = "" then
    response.redirect "datamgr_category_list.asp" & lcl_url_parameters
 end if

 'Setup variables
  lcl_useraction        = ""
  lcl_categoryname      = "NULL"
  lcl_orgid             = request("orgid")
  lcl_isActive          = 0
  lcl_parent_categoryid = 0
  lcl_isApproved        = 1
  lcl_mappointcolor     = "NULL"
  lcl_userid            = session("userid")
  lcl_current_date      = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  lcl_redirect_url      = "datamgr_list.asp"

  if request("user_action") <> "" then
     lcl_useraction = UCASE(request("user_action"))
  end if

 'Execute the user's action
  if lcl_useraction = "DELETE" then
    'First: delete all sub-categories associated to this Category
     if lcl_categoryid > 0 then
        sSQL1 = "DELETE FROM egov_dm_categories WHERE parent_categoryid = " & lcl_categoryid

      		set oDeleteCategory1 = Server.CreateObject("ADODB.Recordset")
       	oDeleteCategory1.Open sSQL1, Application("DSN"), 3, 1

        set oDeleteCategory1 = nothing
     end if

    'Second: delete the Category
     sSQL2 = "DELETE FROM egov_dm_categories WHERE categoryid = " & lcl_categoryid

   		set oDeleteCategory2 = Server.CreateObject("ADODB.Recordset")
    	oDeleteCategory2.Open sSQL2, Application("DSN"), 3, 1

     set oDeleteCategory2 = nothing

     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SD")
     lcl_redirect_url   = "datamgr_categories_list.asp" & lcl_url_parameters
  else

     if request("categoryname") <> "" then
        lcl_categoryname = "'" & dbsafe(request("categoryname")) & "'"
     end if

     if request("isActive") = "Y" then
        lcl_isActive = 1
     end if

     if request("parent_categoryid") <> "" then
        lcl_parentcategoryid = request("parent_categoryid")
     end if

     if request("isApproved") <> "" then
        if request("isApproved") then
           lcl_isApproved = 1
        else
           lcl_isApproved = 0
        end if
     end if

     if request("mappointcolor") <> "" then
        lcl_mappointcolor = "'" & dbsafe(request("mappointcolor")) & "'"
     end if

     if lcl_useraction = "UPDATE" then

      		sSQL = "UPDATE egov_dm_categories SET "
        sSQL = sSQL & "categoryname = "       & lcl_categoryname      & ", "
        sSQL = sSQL & "isActive = "           & lcl_isActive          & ", "
        sSQL = sSQL & "lastmodifiedbyid = "   & lcl_userid            & ", "
        sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date      & ", "
        sSQL = sSQL & "parent_categoryid = "  & lcl_parent_categoryid & ", "
        sSQL = sSQL & "isApproved = "         & lcl_isApproved        & ", "
        sSQL = sSQL & "mappointcolor= "       & lcl_mappointcolor
        sSQL = sSQL & " WHERE categoryid = " & lcl_categoryid

      		set oUpdateCategory = Server.CreateObject("ADODB.Recordset")
	      	oUpdateCategory.Open sSQL, Application("DSN"), 3, 1

        set oUpdateCategory = nothing

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "categoryid", lcl_categoryid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success",    "SU")
        lcl_redirect_url   = "datamgr_categories_maint.asp" & lcl_url_parameters

    '---------------------------------------------------------------------------
     else  'New Category
    '---------------------------------------------------------------------------
        sCreatedByID    = lcl_userid
        sCreatedByDate  = lcl_current_date
        sApprovedByID   = lcl_userid
        sApprovedByDate = lcl_current_date

       'Since we are creating the category from the admin-side we want to automatically
       '"approve" the category when it's created.
        lcl_isApproved = 1

     		'Insert the new Category
   	   	sSQL = "INSERT INTO egov_dm_categories ("
        sSQL = sSQL & "categoryname, "
        sSQL = sSQL & "orgid, "
        sSQL = sSQL & "dm_typeid, "
        sSQL = sSQL & "isActive, "
        sSQL = sSQL & "createdbyid, "
        sSQL = sSQL & "createdbydate, "
        sSQL = sSQL & "lastmodifiedbyid, "
        sSQL = sSQL & "lastmodifiedbydate, "
        sSQL = sSQL & "parent_categoryid, "
        sSQL = sSQL & "isApproved, "
        sSQL = sSQL & "approvedeniedbyid, "
        sSQL = sSQL & "approvedeniedbydate, "
        sSQL = sSQL & "mappointcolor"
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL & lcl_categoryname      & ", "
        sSQL = sSQL & lcl_orgid             & ", "
        sSQL = sSQL & lcl_dm_typeid         & ", "
        sSQL = sSQL & lcl_isActive          & ", "
        sSQL = sSQL & sCreatedByID          & ", "
        sSQL = sSQL & sCreatedByDate        & ", "
        sSQL = sSQL & "NULL,NULL"           & ", "
        sSQL = sSQL & lcl_parent_categoryid & ", "
        sSQL = sSQL & lcl_isApproved        & ", "
        sSQL = sSQL & sApprovedByID         & ", "
        sSQL = sSQL & sApprovedByDate       & ", "
        sSQL = sSQL & lcl_mappointcolor
        sSQL = sSQL & ")"

     		'Get the DMID
    	  	lcl_categoryid = RunIdentityInsert(sSQL)

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "categoryid", lcl_categoryid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success",    "SA")
        lcl_redirect_url   = "datamgr_categories_maint.asp" & lcl_url_parameters

     end if

    'Add/Update the Sub-Categories
     dim lcl_totalsubcategories

     lcl_totalsubcategories = 0

     if request("totalsubcategories") <> "" then
        if isnumeric(lcl_totalsubcategories) then
           lcl_totalsubcategories = clng(request("totalsubcategories"))
        end if
     end if

     if lcl_totalsubcategories > 0 then
        for i = 1 to lcl_totalsubcategories
           dim lcl_sub_categoryid, lcl_sub_categoryname, lcl_sub_isActive, lcl_sub_delete
           dim lcl_mergeIntoCategory, lcl_parent_categoryid, lcl_dmid, lcl_assign_subcategory

           lcl_sub_categoryid     = request("sub_categoryid" & i)
           lcl_sub_categoryname   = request("sub_categoryname" & i)
           lcl_sub_isActive       = "Y"
           lcl_sub_delete         = request("sub_delete" & i)
           lcl_mergeIntoCategory  = request("mergeIntoCategory" & i)
           lcl_parent_categoryid  = lcl_categoryid
           lcl_dmid               = 0
           lcl_assign_subcategory = false

           lcl_sub_categoryid = maintainSubCategory(lcl_orgid, lcl_dm_typeid, lcl_dmid, lcl_userid, lcl_sub_delete, _
                                                    lcl_mergeIntoCategory, lcl_sub_categoryid, lcl_sub_categoryname, _
                                                    lcl_sub_isActive, lcl_parent_categoryid, lcl_assign_subcategory)
        next
     end if

  end if

  response.redirect lcl_redirect_url
%>