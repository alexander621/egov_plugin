<div style="padding-bottom:8px;"><b>Registration Links</b></div>

<% If UserHasPermission( Session("UserId"), "groups" ) Then %>
	<div class="quicklink">&nbsp;&nbsp;<img src="<%=RootPath%>images/newgroup.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="<%=RootPath%>dirs/display_citizen_groups.asp">All Citizen Groups</a></div>
<% End If %>

<% If UserHasPermission( Session("UserId"), "edit citizens" ) Then %>
	<div class="quicklink">&nbsp;&nbsp;<img src="<%=RootPath%>images/newuser.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="<%=RootPath%>dirs/display_citizen.asp">All Citizens</a></div>
<% End If %>

