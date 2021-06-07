<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  dim lcl_rowCount, lcl_fieldtype, lcl_fieldvalue
  dim lcl_button_label, lcl_field_label, lcl_help_url_displayvalue_mouseover, lcl_help_url_displayvalue_mouseout
  dim lcl_url_value, lcl_current_value, lcl_display_url, lcl_display_text, lcl_total_urls
  dim lcl_comma_position, lcl_url_start, lcl_url_end, lcl_url_length, lcl_text_start, lcl_text_end
  dim lcl_text_length, lcl_website_url, lcl_website_text

  lcl_rowCount   = 0
  lcl_fieldtype  = ""
  lcl_fieldvalue = ""

  if request("rowCount") <> "" then
     lcl_rowCount = clng(request("rowCount"))
  end if

  if request("fieldtype") <> "" then
     if not containsApostrophe(request("fieldtype")) then
        lcl_fieldtype = request("fieldtype")
        lcl_fieldtype = ucase(lcl_fieldtype)
     end if
  end if

  if request("fieldvalue") <> "" then
     lcl_fieldvalue = request("fieldvalue")
  end if

  'lcl_fieldvalue = getFieldValue_by_DMValueID(lcl_dm_valueid)

  if instr(lcl_fieldtype,"WEBSITE") > 0 then
     lcl_button_label = "Website"
     lcl_field_label  = "URL"
  else
     lcl_button_label = "Email"
     lcl_field_label  = "Email"
  end if

  lcl_help_url_displayvalue_mouseover = " onMouseOver=""tooltip.show('If a DISPLAY VALUE is not entered then the URL will be used as the clickable link.');"""
  lcl_help_url_displayvalue_mouseout  = " onMouseOut=""tooltip.hide();"""

  'response.write "<div id=""maintain_url" & lcl_rowCount & """ class=""maintain_url"">" & vbcrlf
  response.write "  <p>" & vbcrlf
  response.write "    <input type=""button"" name=""addURLButton"" id=""addURLButton"" value=""Add " & lcl_button_label & """ onclick=""editURL('" & lcl_rowCount & "','','ADD');"" />" & vbcrlf
  response.write "  </p>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" name=""website_table" & lcl_rowCount & """ id=""website_table" & lcl_rowCount & """>" & vbcrlf

 'Break out the values.  Websites and Emails are stored in the following format:
 '1. URL/Email address
 '2. Display Text (clickable link - value CAN be NULL)
 '3. URLs will be surrounded by [].
 '4. Display Text will be surrounded by <>.
 '5. Multiple websites and emails will be seperated by a comma.
 '6. If Display Text is NULL then the URL/Email will be used as the "clickable link".
 '   i.e. [www.mywebsite.com]<My Website>,[www.anotherwebsite.com]<>
  lcl_url_value     = lcl_fieldvalue
  lcl_current_value = ""
  'lcl_display_value = ""
  'lcl_display_url   = ""
  lcl_display_text  = ""
  lcl_total_urls    = 0

  do until lcl_url_value = ""
      lcl_comma_position = 0
      lcl_url_start      = 0
      lcl_url_end        = 0
      lcl_url_length     = 0
      lcl_text_start     = 0
      lcl_text_end       = 0
      lcl_text_length    = 0
      lcl_website_url    = ""
      lcl_website_text   = ""

      if lcl_url_value <> "" then
         lcl_total_urls     = lcl_total_urls + 1
         lcl_comma_position = instr(lcl_url_value,">,[")

         if lcl_comma_position > 0 then
            lcl_current_value = mid(lcl_url_value,1,lcl_comma_position+1)
         else
            lcl_current_value = lcl_url_value
         end if

         lcl_url_start      = instr(lcl_current_value,"[")
         lcl_url_end        = instr(lcl_current_value,"]")
         lcl_url_length     = lcl_url_end - lcl_url_start

         lcl_text_start     = instr(lcl_current_value,"<")
         lcl_text_end       = instr(lcl_current_value,">")
         lcl_text_length    = lcl_text_end - lcl_text_start

         lcl_website_url  = mid(lcl_current_value,lcl_url_start,lcl_url_length)
         lcl_website_url  = replace(lcl_website_url,"[","")
         lcl_website_url  = replace(lcl_website_url,"]","")

         lcl_website_text = mid(lcl_current_value,lcl_text_start,lcl_text_length)
         lcl_website_text = replace(lcl_website_text,"<","")
         lcl_website_text = replace(lcl_website_text,">","")

         'lcl_display_url  = ""
         lcl_display_text = lcl_website_url

         lcl_url_value = replace(lcl_url_value,lcl_current_value,"")

         if lcl_website_text <> "" then
            lcl_display_text = lcl_website_text
         end if

        'Build the "display_url"
         'lcl_display_url = "<a href="""

         'if instr(lcl_fieldtype,"EMAIL") > 0 then
         '   lcl_display_url = lcl_display_url & "mailto:"
         'end if

         'lcl_display_url = lcl_display_url & lcl_website_url & """ target=""_blank"">" & lcl_display_text & "</a>"

         'if lcl_display_value <> "" then
         '   lcl_display_value = lcl_display_value & "<br />" & lcl_display_url
         'else
         '   lcl_display_value = lcl_display_value & "<a href=""" & lcl_website_url & """ target=""_blank"">" & lcl_display_text & "</a>"
         'end if
      end if

      response.write "    <tr>" & vbcrlf
      response.write "        <td>" & vbcrlf
      response.write              lcl_field_label & ": " & vbcrlf
      response.write "            <input type=""text"" name=""website_url"            & lcl_rowCount & "_" & lcl_total_urls & """ id=""website_url"          & lcl_rowCount & "_" & lcl_total_urls & """ value=""" & lcl_website_url & """ size=""40"" maxlength=""50"" onchange=""clearMsg('website_url" & lcl_rowCount & "_" & lcl_total_urls & "');"" />" & vbcrlf
      response.write "            <input type=""hidden"" name=""original_website_url" & lcl_rowCount & "_" & lcl_total_urls & """ id=""original_website_url" & lcl_rowCount & "_" & lcl_total_urls & """ value=""" & lcl_website_url & """ size=""40"" maxlength=""50"" />" & vbcrlf
      response.write "        </td>" & vbcrlf
      response.write "        <td>" & vbcrlf
      response.write "            Display Value: " & vbcrlf
      response.write "            <input type=""text"" name=""website_text"            & lcl_rowCount & "_" & lcl_total_urls & """ id=""website_text"          & lcl_rowCount & "_" & lcl_total_urls & """ value=""" & lcl_website_text & """ size=""30"" maxlength=""50"" />" & vbcrlf
      response.write "            <input type=""hidden"" name=""original_website_text" & lcl_rowCount & "_" & lcl_total_urls & """ id=""original_website_text" & lcl_rowCount & "_" & lcl_total_urls & """ value=""" & lcl_website_text & """ size=""30"" maxlength=""50"" />" & vbcrlf
      response.write "            <img src=""../images/help_graybg.jpg"" name=""helpFeature_url" & lcl_rowCount & "_" & lcl_total_urls & """ id=""helpFeature_url" & lcl_rowCount & "_" & lcl_total_urls & """ class=""helpOption""" & lcl_help_url_displayvalue_mouseover & lcl_help_url_displayvalue_mouseout & " />" & vbcrlf
      response.write "        </td>" & vbcrlf
      response.write "        <td>" & vbcrlf
      response.write "            <input type=""checkbox"" name=""removeURL" & lcl_rowCount & "_" & lcl_total_urls & """ id=""removeURL" & lcl_rowCount & "_" & lcl_total_urls & """ value=""Y"" onclick=""editURL('" & lcl_rowCount & "','" & lcl_total_urls & "','DELETE');"" /> Remove" & vbcrlf
      response.write "        </td>" & vbcrlf
      response.write "    </tr>" & vbcrlf
  loop

  response.write "  </table>" & vbcrlf
  response.write "  <p>" & vbcrlf
  response.write "    <input type=""hidden"" name=""total_urls_" & lcl_rowCount & """ id=""total_urls_" & lcl_rowCount & """ value=""" & lcl_total_urls & """ size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""button"" name=""cancelURLButton"" id=""cancelURLButton"" value=""Cancel"" class=""button"" onclick=""editURL('" & lcl_rowCount & "','" & lcl_total_urls & "','CANCEL');"" />" & vbcrlf
  response.write "    <input type=""button"" name=""saveURLButton"" id=""saveURLButton"" value=""Finished " & lcl_button_label & " Changes"" class=""button"" onclick=""editURL('" & lcl_rowCount & "','" & lcl_total_urls & "','SAVE');"" />" & vbcrlf
  response.write "  </p>" & vbcrlf
  'response.write "</div>" & vbcrlf

             'Determine if there are any scripts to run
  'if lcl_display_value <> "" then
  '   response.write "<script type=""text/javascript"">" & vbcrlf
  '   response.write "$('#dm_fieldvalue" & lcl_rowCount & "_display').html('" & lcl_display_value & "');" & vbcrlf
  '   response.write "</script>" & vbcrlf
  'end if
%>