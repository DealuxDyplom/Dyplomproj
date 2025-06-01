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
<table width="760" border="0" cellspacing="2" cellpadding="0">
<tr> 
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
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
<%
Dim Flag, Query
' If the Category Name field is empty, and the form has been submitted:			
If Request("TypeName")="" Then
	If Flag Then
		Response.Write "<p class='subtitle'>You must enter a cateogry name.</p>"
		Else
' If Flag isn't true, show the edit form:
Dim CategoryRS
Set CategoryRS = Server.CreateObject("ADODB.Recordset")
			CategoryRS.Open "SELECT * FROM Type WHERE TypeId =" & Request("iType"), Connect, 2, 3
			End If
			%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr>      
      <td align="left" valign="top"> 
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td align="left" valign="middle"><p>&nbsp;</p><form method="POST" action="<%= Request.ServerVariables("URL") %>" name="ADD">
              <table align="center" class="textBold">
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">EDIT CATEGORY NAME:</td>
                  <td> <input name="TypeName" type="text" class="form" value="<%=CategoryRS("TypeName")%>" size="50" maxlength="50"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">CATEGORY LINK COLOR:</td>
                  <td> <select name="LinkColor">
                  <%
				  Dim LinkColorRS
				  Set LinkColorRS = Server.CreateObject("ADODB.Recordset")
LinkColorRS.Open "SELECT *  FROM LinkColor", Connect, 3, 3
  While (Not LinkColorRS.EOF)
    Response.Write "<option STYLE='color: " & LinkColorRS("ColorCode") & " !important' value='" & LinkColorRS("ColorCode") & "'"
	If (CategoryRS("TypeColor") = LinkColorRS("ColorCode")) Then Response.Write "selected"
	Response.Write">" & LinkColorRS("ColorName") & "</option>"
    LinkColorRS.MoveNext
	Wend
		LinkColorRS.Close()
			CategoryRS.Close()
%>
                  </select> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right">&nbsp;</td>
                  <td> <input name="Submit" type="submit" class="form" id="Submit" value="Submit"> 
                  </td>
                </tr>
              </table>
              <input type="hidden" name="Flag" value="True">
              <input type="hidden" name="TypeId" value="<%=Request("iType")%>">
            </form>
             <%
' If the form has been submitted (Flag = True), and the Category Name field isn't empty, update the record:
Else
Dim TypeRS
	Query = "SELECT * FROM Type WHERE TypeId=" & Request("TypeId")
	Set TypeRS = Server.CreateObject("ADODB.Recordset")
	TypeRS.Open Query, Connect, 2, 3
	
	TypeRS("TypeName") = Request("TypeName")
	TypeRS("TypeColor") = Request("LinkColor")

	TypeRS.Update
	TypeRS.Close()
	Response.Redirect("category.asp")
	
End If
%></td>
        </tr>       
        <tr> 
          <td align="center" valign="top">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
		  </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
      <!--#include file="inc_footer.asp" -->
    </td>
  </tr>  
</table>
</body>
</html>
