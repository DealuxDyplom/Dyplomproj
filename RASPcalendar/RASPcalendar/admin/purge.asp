<% @EnableSessionState=True %>
<% Option Explicit %>
<%
'--------------------------------------------------
':::: RASPcalendar ::::
'--------------------------------------------------
'ASP & Microsoft Access online calendar application.
'Includes an administration interface to allow for easy updates and maintenance.
'Copyright (C) 2010  Robert Temple
'--------------------------------------------------
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'You can contact the author at rttucson@gmail.com
'Please do not remove the author, copyright, or license info.
%><%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%> 
<% If Request.Cookies("AppLoggedIn") <> "True" Then Response.Redirect("default.asp")
If Request.Cookies("Admin") <> "True" Then Response.Redirect("default.asp") %>
<!--#include file="../Connections/connRASPcalendarAdmin.asp" -->
<html>
<head>
<title>Calendar Admin</title>
<link href="../assets/RASPcalendar.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<table width="760" border="0" cellpadding="1" cellspacing="0" bgcolor="#003399">
<tr> 
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
        <tr>
          <td bgcolor="#FFFFFF">
 <div class = "links">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2" align="left" valign="middle" bgcolor="#FFFFFF" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td colspan="2" align="left" valign="middle" bgcolor="#000000" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
      <td height="18" align="left" valign="middle" bgcolor="#FFE5B2" class="textBold">&nbsp;
        <a href="events.asp">EVENT LISTING</a> | <a href="submit.asp">ADD NEW EVENT</a> | <a href="purge.asp">PURGE OLD EVENTS</a> | <a href="category.asp">EDIT CATEGORIES</a></td>
      <td align="right" valign="middle" bgcolor="#FFE5B2" class="textBold"><A HREF="logout.asp">LOGOUT</A>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" align="left" valign="middle" bgcolor="#333333"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
</div>
</td></tr>
<tr>
<td bgcolor="#FFFFFF">&nbsp;</td></tr>
<tr>
<td bgcolor="#FFFFFF">
   <%
Dim EventsRS, Query, Flag
If Request("Flag") = "True" Then Flag = True

' If the form hasn't been submitted, get all records older that 182 days (approx. six months):		
If Not Flag Then
		Dim CountRS, CountQuery, cNum
		Set CountRS = Server.CreateObject("ADODB.Recordset")
		CountQuery = "SELECT EventDated FROM Events WHERE EventDated <= dateAdd('d',now(),-182)"	
		CountRS.Open CountQuery, Connect, 3, 3
		
'set our variable count equal to the number of records
cNum=CountRS.recordcount

' Alert if there are no records older than six months:	
	 If cNum="0" then Response.Write "<p align=""center"" class=""textBold"">There are no records older than six months in the database.</p><p>&nbsp;</p>" 	
%>
     <%
	 If (cNum > 0) Then
' If there are records older than six months, display delete form:
	 %>
  <p align="center" class="textBold">Do you wish to delete ALL events older than 
			  <% Dim vble
			vble= dateAdd("d",now(),-182)
	 Response.Write vble
	 %>
     ? (This cannot be undone.)</p>
     <p align="center" class="textBold">Total records to be deleted = <% =cNum  %></p></p>
            <form name="form_delete" method="post" action="<%= Request.ServerVariables("URL") %>">
          <table width="200" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
            <tr> 
              <td><img src="../images/spacer.gif" width="1" height="10"></td>
              </tr>        
            <tr> 
                <td><input type="hidden" name="Flag" value="True"></td>
              </tr>
            <tr> 
              <td align="center"><input type="submit" name="Submit" value="Delete Old Events" class="text"></td>
              </tr>
            <tr> 
              <td>&nbsp;</td>
              </tr>
            <tr> 
              <td align="center"><a href="events.asp">Do not delete</a></td>
              </tr>
          </table>
      </form>
      <%
	  End If
	  %>
</td>
</tr>
</table>      
    </td>
  </tr>
</table>
<% 
' If the form was submitted, delete the old records:
Else
		Set EventsRS = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT * FROM Events WHERE EventDated <= dateAdd('d',now(),-182)"	
		EventsRS.Open Query, Connect, 2, 3		
		
If Not EventsRS.EOF Then
	Do While Not EventsRS.EOF
		EventsRS.Delete	
		EventsRS.MoveNext
	Loop
		Response.Write "<p align=""center"" class=""textBold"">Old events were successfully deleted.</p><p>&nbsp;</p>"
End If

End If

%>      
</body>
</html>