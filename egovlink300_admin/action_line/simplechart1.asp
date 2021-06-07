<!-- #include file="../includes/common.asp" //-->
<%
sLevel = "../"     'Override of value from common.asp

'Get the search parameters
 lcl_sc_form_name  = request("sc_form_name")
 lcl_sc_start_date = request("sc_start_date")
 lcl_sc_end_date   = request("sc_end_date")
 lcl_sc_chart_type = request("sc_chart_type")

 if lcl_sc_chart_type = "" OR isnull(lcl_sc_chart_type) then
    lcl_sc_chart_type = "p"
 end if

 if lcl_sc_start_date = "" OR isnull(lcl_sc_start_date) then
    lcl_sc_start_date = "01/01/2009"
 end If

 If lcl_sc_chart_type = "p" Then 
	sChartSrc = "simplepie1.asp"
 End If 

 iMax = CLng(0)
 
'Build CHART Query
 sSQL = "SELECT distinct action_formtitle, action_formid, count(action_formid) as total_count "
 sSQL = sSQL & " FROM egov_action_request_view "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND action_formtitle NOT LIKE ('Old System Requests%') "
 sSQL = sSQL & " AND submit_date >= '" & lcl_sc_start_date & "' "

 if lcl_sc_end_date <> "" AND NOT isnull(lcl_sc_end_date) then
    sSQL = sSQL & " AND submit_date < '" & lcl_sc_end_date & "' "
 end if
 if lcl_sc_form_name <> "" AND NOT isnull(lcl_sc_form_name) then
    sSQL = sSQL & " AND UPPER(action_formtitle) LIKE ('%" & UCASE(lcl_sc_form_name) & "%') "
 end If
 sSql = sSql & " group by action_formtitle, action_formid order by action_formtitle"

session("chartsql") = sSql

%>
<html>
<body>

<form name="display_graphs" method="post" action="test.asp">

<fieldset>
  <legend>Search Criteria </legend>
  <table border="0" cellspacing="0" cellpadding="2">
    <tr>
        <td>Action Line Form Name: </td>
        <td><input type="text" name="sc_form_name" value="<%=lcl_sc_form_name%>" size="30" maxlength="512"></td>
        <td>Chart Type: </td>
        <td><select name="sc_chart_type">
              <%
                lcl_selected_bhs = ""
                lcl_selected_bvs = ""
                lcl_selected_p   = ""
                lcl_selected_p3  = ""

                if lcl_sc_chart_type = "bhs" then lcl_selected_bhs = " selected" end if
                if lcl_sc_chart_type = "bvs" then lcl_selected_bvs = " selected" end if
                if lcl_sc_chart_type = "p"   then lcl_selected_p   = " selected" end if
                if lcl_sc_chart_type = "p3"  then lcl_selected_p3  = " selected" end if
              %>

              <option value="bhs"<%=lcl_selected_bhs%>>Bar Chart (Horizontal)</option>
              <option value="bvs"<%=lcl_selected_bvs%>>Bar Chart (Vertical)</option>
              <option value="p"<%=lcl_selected_p%>>Pie Graph (2-Dimensional)</option>
              <option value="p3"<%=lcl_selected_p3%>>Pie Graph (3-Dimensional)</option>
            </select>
        </td>
    </tr>
    <tr>
        <td>Start Date: </td>
        <td colspan="3"><input type="text" name="sc_start_date" value="<%=lcl_sc_start_date%>" size="10" maxlength="10"></td>
    </tr>
    <tr>
        <td>End Date: </td>
        <td colspan="3"><input type="text" name="sc_end_date" value="<%=lcl_sc_end_date%>" size="10" maxlength="10"></td>
    </tr>
    <tr>
        <td colspan="4">
            <input type="submit" name="sAction" value="View Chart">
        </td>
    </tr>
  </table>
</fieldset>
<p>
<center>
	<img src="<%=sChartSrc%>" width="1000" height="800" />
</center>
<p>



</form>
</body>
</html>
