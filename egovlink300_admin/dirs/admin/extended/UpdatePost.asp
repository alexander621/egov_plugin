<FORM METHOD=POST ACTION="<%=thisname%>?iofaction=<%=ActUpdatePostSave%>&userid=<%=request.querystring("userid")%>">
<table border="0"  class='tablelist' width="400">
	<% 
	intcolumns=UBOUND(conTABLEFIELDS)
	for j=0 to intcolumns
	if l_modify(j) then 
	  if l_fields(j)>=100 then	   %>          
	    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><textarea rows="2" cols="50" name="<% =conTABLEFIELDS(j) %>"><%=rs(j)%></textarea></td>
       </tr>    
	<% 	elseif instr(TableDescription(j)," bit") then 
	if rs(j)=true then string1="checked" else string1="" 
	if rs(j)=false then string2="checked" else string2="" 
	
	%>
	<tr>
    <td width="200"><% =fields_description(j) %></td>
    <td width="351"><input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string1%> value='true'>Yes&nbsp;&nbsp;
	                <input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string2%> value='false'>No
	</td>
	</tr>


	<% elseif instr(conTABLEFIELDS(j),"tatus")>0 then 
	' want to add some special to status modify
	%>
    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351">
	<% if rs(j)=1 then 
	   response.write " closed"
	  else
	  if rs(j)=0 then string1="checked" else string1="" 
      if rs(j)=-1 then string2="checked" else string2="" 
	  %>
        <input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string1%> value=0>Want Pickup&nbsp;&nbsp;
	    <input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string2%> value=-1>Solved by self
	<%  end if   %>
	</td>
    </tr>

	
	 <% else %>
    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><input type=text name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>" size=<%=l_fields(j)%> maxlength=<%=l_fields(j)%>></td>
    </tr>
	  <% end if %>
	<% else %>

	<input type=hidden name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>">

	<%
	end if 
	next
	%>
</table>
<INPUT TYPE="submit" value='Update Now'>
</FORM>

