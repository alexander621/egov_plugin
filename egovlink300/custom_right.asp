<%

sPath = Server.MapPath("custom_html/custom_right_" & iOrgID & ".asp")

 response.write ReadFile(sPath)

%>