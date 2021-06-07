<!-- #include file="../includes/common.asp" //-->
<%
  dim lcl_task, lcl_cal, lcl_success, lcl_total_categories, i

  lcl_task             = ""
  lcl_cal              = ""
  lcl_success          = ""
  lcl_total_categories = 0

  if request("_task") <> "" then
     if not containsApostrophe(request("_task")) then
        lcl_task = request("_task")
     end if
  end if

  if request("cal") <> "" then
     if not containsApostrophe(request("cal")) then
        lcl_cal = request("cal")
     end if
  end if

  if lcl_task = "editcategories" then
     if request("totalCategories") <> "" then
        if isnumeric(request("totalCategories")) then
           lcl_total_categories = clng(request("totalCategories"))
        end if
     end if

     if lcl_total_categories > 0 then
        i = 1

        for i = 1 to lcl_total_categories
           if request("deleteCategory_" & i) <> "" then
              lcl_categoryid = 0
              sSQLa          = ""

              if request("deleteCategory_" & i) <> "" then
                 if isnumeric(request("deleteCategory_" & i)) then
                    lcl_categoryid = clng(request("deleteCategory_" & i))

                    sSQLa = "DELETE FROM eventcategories WHERE categoryid = " & lcl_categoryid
                 end if
              end if
           else
              if request("CustomCategory_" & i) <> "" OR request("CustomColor_" & i) <> "" then
                 lcl_categoryid       = 0
                 lcl_newcategoryname  = ""
                 lcl_newcategorycolor = ""

                 if request("categoryid_" & i) <> "" then
                    if isnumeric(request("categoryid_" & i)) then
                       lcl_categoryid = clng(request("categoryid_" & i))
                    end if
                 end if

                 if request("CustomCategory_" & i) <> "" then
                    lcl_newcategoryname = request("CustomCategory_" & i)
                    lcl_newcategoryname = dbsafe(lcl_newcategoryname)
                    lcl_newcategoryname = "'" & lcl_newcategoryname & "'"
                 end if

                 if request("CustomColor_" & i) <> "" then
                    lcl_newcategorycolor = request("CustomColor_" & i)
                    lcl_newcategorycolor = dbsafe(lcl_newcategorycolor)
                    lcl_newcategorycolor = "'" & lcl_newcategorycolor & "'"
                 end if

                 sSQLa = "UPDATE eventcategories SET "

                 if lcl_newcategoryname <> "" then
                    sSQLa = sSQLa & " categoryname = " & lcl_newcategoryname
                 end if

                 if lcl_newcategorycolor <> "" then
                    if lcl_newcategoryname <> "" then
                       sSQLa = sSQLa & ","
                    end if

                    sSQLa = sSQLa & " color = " & lcl_newcategorycolor
                 end if

                 sSQLa = sSQLa & " WHERE categoryid = " & lcl_categoryid
              end if
           end if

           if sSQLa <> "" then
              set oUpdate = Server.CreateObject("ADODB.Recordset")
          				oUpdate.Open sSQLa, Application("DSN"), 3, 1
         	 			set oUpdate = nothing
           end if

        next
     end if

     lcl_success = "SU"

  end if

  response.redirect "eventcategories.asp?success=" & lcl_success & "&cal=" & lcl_cal

%>