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
' *** Edit Operations: declare variables
Dim MM_editAction, MM_abortEdit, MM_editQuery, MM_recordCount
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>

<%
' *** Delete Record: declare variables
Dim MM_editTable, MM_editColumn, MM_recordId, MM_editRedirectUrl

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editTable = "Type"
  MM_editColumn = "TypeId"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "category.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>

<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
  
  	' delete links of this type
	If(Request("MM_recordCount") <> 0) Then
Dim evntDeleteRS
	set evntDeleteRS = Server.CreateObject("ADODB.Recordset")
	evntDeleteRS.Open "SELECT * FROM Events WHERE EventType = " & Request("MM_recordId"), Connect, 2, 3
	  While (Not evntDeleteRS.EOF)
	evntDeleteRS.Delete	
    evntDeleteRS.MoveNext
  Wend
End If

    ' execute the delete
	Dim CategoryDeleteRS			
	Set CategoryDeleteRS = Server.CreateObject("ADODB.Recordset")
			CategoryDeleteRS.Open "SELECT * FROM Type WHERE TypeId =" & Request("MM_recordId"), Connect, 2, 3
			CategoryDeleteRS.Delete	

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
' Display current categories along with the number of events in each category:
Dim rsTypes, rsTypes_numRows
set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.Open "SELECT *, (SELECT COUNT(*) FROM Events WHERE EventType = TypeId) AS EVE_COUNT FROM Type ORDER BY TypeName ASC", Connect, 3, 3
rsTypes_numRows = 0

Dim HLooper1__numRows
HLooper1__numRows = -2
Dim HLooper1__index
HLooper1__index = 0
rsTypes_numRows = rsTypes_numRows + HLooper1__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
      <td align="left" valign="top"> 
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td align="left" valign="middle"><p>&nbsp;</p><form method="POST" action="category_add.asp" name="ADD">
              <table align="center" class="textBold">
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">ADD NEW CATEGORY:</td>
                  <td> <input name="TypeName" type="text" class="form" value="" size="50" maxlength="50"> 
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
    Response.Write "<option STYLE='color: " & LinkColorRS("ColorCode") & " !important' value='" & LinkColorRS("ColorCode") & "'>" & LinkColorRS("ColorName") & "</option>"
    LinkColorRS.MoveNext
	Wend
%>
                  </select> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right">&nbsp;</td>
                  <td> <input name="Add" type="submit" class="form" id="Add" value="Add"> 
                  </td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr> 
          <td align="left" valign="middle"> <hr> </td>
        </tr>
        <tr> 
          <td align="left" valign="top" class="textBold"><font color="#FF0000">Warning: deleting a category will also delete all of its events.</font></td>
        </tr>
        <tr> 
          <td align="center" valign="top"> <table width="98%">
              <%
Dim startrw, endrw, numberColumns, numrows
startrw = 0
endrw = HLooper1__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsTypes.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
              <tr align="center" valign="top"> 
<%
While ((startrw <= endrw) AND (Not rsTypes.EOF))
%>
                <form ACTION="<%=MM_editAction%>" METHOD="POST" name="DELETE">
                  <td> <table width="100%" border="0" cellpadding="2" cellspacing="2">
                      <tr align="left" valign="middle"> 
                        <td width="22"> <input name="Submit" type="submit" class="form" value="DELETE"> 
                        </td>
                        <td class="textBold"><a href="category_edit.asp?iType=<%=(rsTypes.Fields.Item("TypeId").Value)%>"><%=(rsTypes.Fields.Item("TypeName").Value)%></a> (<%=(rsTypes.Fields.Item("EVE_COUNT").Value)%>) 
                       </td> 
						<input type="hidden" name="MM_delete" value="DELETE">
                        <input type="hidden" name="MM_recordId" value="<%= rsTypes.Fields.Item("TypeId").Value %>">
                        <input type="hidden" name="MM_recordCount" value="<%=(rsTypes.Fields.Item("EVE_COUNT").Value)%>">
                     
                    </table></td>
                </form>
                <%
	startrw = startrw + 1
	rsTypes.MoveNext()
	Wend
	%>
              </tr>
              <%
 numrows=numrows-1
 Wend
 %>
            </table></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
<%
rsTypes.Close()
%>
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
