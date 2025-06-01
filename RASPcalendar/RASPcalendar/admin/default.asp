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
<!--#include file="../Connections/connRASPcalendarAdmin.asp" -->
<%
Dim Flag, Message, LoggedIn

If Request("Flag") = "True" Then Flag = True Else Flag = False

If Flag Then

	Dim UsersRS, Query
	
	Query = "SELECT * FROM Users WHERE Username='" & Request("Username") & "' AND Password='" & Request("Password") & "'"
	Set UsersRS = Connect.Execute(Query)		
			
	If Not UsersRS.EOF Then
	
'	==================================================================================
' If the user is in the database and has administration privileges, write admin cookies and
' redirect to the events listing page:		
		If UsersRS("Administration") = "Y" Then Response.Cookies("Admin") = "True"
		Response.Cookies("AppUName") = UsersRS("Username")
		Response.Cookies("AppLoggedIn") = "True"
		Response.Redirect("events.asp")
		
'Close ADO database connection and free DB variables
UsersRS.Close 
Set UsersRS = Nothing

				
	Else
		Message = "<p class='required'>Invalid log in. Please try again.</p>"
		Response.Cookies("AppLoggedIn") = "False"
	End If
	
Else

	
End If
%>
<html>
<head>
<title>RASPcalendar</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../assets/RASPcalendar.css" rel="stylesheet" type="text/css">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bg OnLoad="document.form1.id.focus();">
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="../../calendar/assets/bground.gif">
  <tr>
    <td align="center" valign="middle"><table width="100%" height="300" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="2" align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td align="center" valign="middle" bgcolor="#006699"><form name="form1" method="POST" action="<%= Request.ServerVariables("URL") %>">
              <table border="0" cellspacing="4" cellpadding="4">
                <tr> 
                  <td colspan="2" class="textBold"><font color="#FFFFFF" size="4"><strong>Calendar ADMIN PANEL</strong></font></td>
                </tr>
                <tr> 
                  <td align="right" valign="middle" class="textBold"><font color="#FFFFFF">Username: </font></td>
                  <td><font color="#FFFFFF"> 
                    <input name="Username" type="text" class="form" value="" size="20">
                    </font></td>
                </tr>
                <tr> 
                  <td align="right" valign="middle" class="textBold"><font color="#FFFFFF">Password: 
                    </font></td>
                  <td><font color="#FFFFFF"> 
                    <input name="Password" type="Password" class="form" id="Password" size="20">
                    </font></td>
                </tr>
                <tr> 
                  <td><input type="hidden" name="Flag" value="True">
</td>
                  <td> <input name="Submit" type="submit" class="form"  value="Login"></td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr> 
          <td height="2" align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>