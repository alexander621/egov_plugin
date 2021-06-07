<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
 dim lcl_action, lcl_orgid, lcl_categoryid, lcl_categoryname

 lcl_action       = ""
 lcl_orgid        = 0
 lcl_categoryid   = 0
 lcl_categoryname = ""

 if request("action") <> "" then
    if not containsApostrophe(request("action")) then
       lcl_action = ucase(request("action"))
    end if
 end if

 if request("orgid") <> "" then
    lcl_orgid = clng(request("orgid"))
 end if

 if lcl_action = "DELETE" then
    if request("categoryid") <> "" then
       lcl_categoryid = clng(request("categoryid"))
    end if

    sSQLd = "DELETE FROM egov_form_categories WHERE form_category_id = " & lcl_categoryid

    set oDeleteCategory = Server.CreateObject("ADODB.Recordset")
    oDeleteCategory.Open sSQLd, lcl_dsn, 1, 3

    lcl_success = "SD"

    set oDeleteCategory = nothing

   'Renumber the remaining categories
    renumberFormCategories lcl_orgid
 elseif lcl_action = "ADD" then
    if request("newCategoryName") <> "" then
       lcl_categoryname = request("newCategoryName")
       lcl_categoryname = dbsafe(lcl_categoryname)
       lcl_categoryname = "'" & lcl_categoryname & "'"

       lcl_max_sequence = getMaxCategorySequence(lcl_orgid)

       sSQLi = "INSERT egov_form_categories ("
       sSQLi = sSQLi & " form_category_name, "
       sSQLi = sSQLi & " form_category_sequence, "
       sSQLi = sSQLi & " orgid"
       sSQLi = sSQLi & " ) VALUES ("
       sSQLi = sSQLi & lcl_categoryname & ", "
       sSQLi = sSQLi & lcl_max_sequence & ", "
       sSQLi = sSQLi & lcl_orgid
       sSQLi = sSQLi & ") "

       'response.write sSQLi
       'response.end

       set oNewCategory = Server.CreateObject("ADODB.Recordset")
       oNewCategory.Open sSQLi, Application("DSN"), 3, 1
       set oNewCategory = nothing 

       renumberFormCategories lcl_orgid

       lcl_success = "SA"
    end if

 elseif lcl_action = "EDIT" then
   for each Item In request.form
     	if left(Item,12) = "editcategory" then
    	  		if request.form(Item) <> "" then
            sCatID = replace(Item,"editcategory","")

            sCatName = request("editcategory" & sCatID)
            sCatName = dbsafe(sCatName)
            sCatName = "'" & sCatName & "'"

      	  			sSQLa = "UPDATE egov_form_categories SET "
            sSQLa = sSQLa & " form_category_name = " & sCatName
            sSQLa = sSQLa & " WHERE form_category_id=" & sCatID
            sSQLa = sSQLa & " AND orgid = " & lcl_orgid

        				set oUpdateCategory = Server.CreateObject("ADODB.Recordset")
        				oUpdateCategory.Open sSQLa, Application("DSN"), 3, 1
        				set oUpdateCategory = nothing 
    	  		end if
     	end if
   next

   lcl_success = "SU"

 end if

 response.redirect "actioncategories.asp?success=" & lcl_success

'------------------------------------------------------------------------------
sub renumberFormCategories(iOrgID)
	dim sSQL, sSQLu, lcl_lineCount, sOrgID

 sOrgID        = 0
 lcl_lineCount = 0

 if iOrgID <> "" then
    sOrgID = clng(iOrgID)
 end if

	sSQL = "SELECT form_category_id "
 sSQL = sSQL & " FROM egov_form_categories "
 sSQL = sSQL & " WHERE orgid = " & sOrgID
 sSQL = sSQL & " ORDER BY form_category_sequence"

	set oRenumberCategories = Server.CreateObject("ADODB.Recordset")
	oRenumberCategories.Open sSQL, Application("DSN"), 3, 1

	do while not oRenumberCategories.eof
  		lcl_lineCount = lcl_lineCount + 1

    sSQLu = "UPDATE egov_form_categories SET "
    sSQLu = sSQLu & " form_category_sequence = " & lcl_lineCount
    sSQLu = sSQLu & " WHERE form_category_id = " & oRenumberCategories("form_category_id")

   	set oUpdateCategorySequence = Server.CreateObject("ADODB.Recordset")
   	oUpdateCategorySequence.Open sSQLu, Application("DSN"), 3, 1

    set oUpdateCategorySequence = nothing

  		oRenumberCategories.movenext
 loop
	
	oRenumberCategories.close
	set oRenumberCategories = nothing

end sub

'------------------------------------------------------------------------------
function getMaxCategorySequence(iOrgID)

  dim sOrgID, lcl_return

  sOrgID     = 0
  lcl_return = 1

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  sSQLs = "SELECT max(form_category_sequence)+1 as new_category_sequence "
  sSQLs = sSQLs & " FROM egov_form_categories "
  sSQLs = sSQLs & " WHERE orgid = " & sOrgID

 	set oGetNextCategorySequence = Server.CreateObject("ADODB.Recordset")
 	oGetNextCategorySequence.Open sSQLs, Application("DSN"), 3, 1

  if not oGetNextCategorySequence.eof then
     lcl_return = oGetNextCategorySequence("new_category_sequence")
  end if

  if lcl_return = "" or isnull(lcl_return) then lcl_return = 1

  set oGetNextCategorySequence = nothing

  getMaxCategorySequence = lcl_return

end function
%>
