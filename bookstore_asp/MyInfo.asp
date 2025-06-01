<!-- #INCLUDE FILE="Common.asp" -->

<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->
<%
'
'   Filename: MyInfo.asp
'   Generated with CodeCharge 1.1.19
'   ASP.ccp build 5/9/2001
'

sFileName = "MyInfo.asp"



CheckSecurity(1)


sFormErr = ""

sAction = GetParam("FormAction")
sForm = GetParam("FormName")
Select Case sForm
  Case "Form"
    FormAction(sAction)
end select


%><html>
<head>
<title>Book Store</title>
<meta name="GENERATOR" content="YesSoftware CodeCharge v.1.1.19 using 'ASP.ccp'">
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="expires" content="0"> 
<meta http-equiv="cache-control" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
</head>
<body style="background-color: #FFFFFF; color: #000000; font-family: Arial, Tahoma, Verdana, Helveticabackground-color: #FFFFFF; color: #000000; font-family: Arial, Tahoma, Verdana, Helvetica">
<center>
 <table>
  <tr>
   <td valign="top">
 <% Menu_Show %>
   
   </td>
  </tr>
 </table>
</center>
<hr>
 <table>
  <tr>
   
   <td valign="top">
<% Form_Show %>
    
   </td>
  </tr>
 </table>

<center>
<hr size=1 width=60%>
 <table>
  <tr>
   <td valign="top">
<% Footer_Show %>
    </td>
   
  </tr>
 </table>
 
<center><font face="Arial"><small>This dynamic site was generated with <a href="http://www.codecharge.com">CodeCharge</a></small></font></center>
</body>
</html>
<% 



'********************************************************************************



Sub FormAction(sAction)
  
  sActionFileName = "ShoppingCart.asp"
  
  
  if sAction = "cancel" then response.redirect sActionFileName

  sWhere = ""
  bErr = False  


  ' Create WHERE statement
  if sAction = "update" or sAction = "delete" then
    pPKmember_id = GetParam("PK_member_id")
    if IsEmpty(pPKmember_id) then exit sub
    sWhere = "member_id=" & ToSQL(pPKmember_id, "Number")
  end if


  ' Load all form fields into variables

  fldUserID = Session("UserID")
  fldmember_password = GetParam("member_password")
  fldname = GetParam("name")
  fldlast_name = GetParam("last_name")
  fldemail = GetParam("email")
  fldaddress = GetParam("address")
  fldphone = GetParam("phone")
  fldnotes = GetParam("notes")
  fldcard_type_id = GetParam("card_type_id")
  fldcard_number = GetParam("card_number")
  ' Validate fields
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldmember_password) then
      sFormErr = sFormErr & "The value in field Password* is required.<br>"
    end if
    if IsEmpty(fldname) then
      sFormErr = sFormErr & "The value in field First Name* is required.<br>"
    end if
    if IsEmpty(fldlast_name) then
      sFormErr = sFormErr & "The value in field Last Name* is required.<br>"
    end if
    if IsEmpty(fldemail) then
      sFormErr = sFormErr & "The value in field Email* is required.<br>"
    end if
    if not isNumeric(fldcard_type_id) then
      sFormErr = sFormErr & "The value in field Credit Card Type is incorrect.<br>"
    end if
    If len(sFormErr) > 0 then
      exit sub
    end if
  end if


  sSQL = ""
  ' Create SQL statement

  select case sAction
    case "update"
      sSQL = "update members set " & _
        "member_password=" & ToSQL(fldmember_password, "Text") & _
        ",first_name=" & ToSQL(fldname, "Text") & _
        ",last_name=" & ToSQL(fldlast_name, "Text") & _
        ",email=" & ToSQL(fldemail, "Text") & _
        ",address=" & ToSQL(fldaddress, "Text") & _
        ",phone=" & ToSQL(fldphone, "Text") & _
        ",notes=" & ToSQL(fldnotes, "Text") & _
        ",card_type_id=" & ToSQL(fldcard_type_id, "Number") & _
        ",card_number=" & ToSQL(fldcard_number, "Text")
      sSQL = sSQL & " where " & sWhere
  end select

  ' Execute SQL statement
  
  if len(sFormErr) > 0 then Exit Sub
  on error resume next
  cn.execute sSQL
  sFormErr = ProceedError
  if len(sFormErr) > 0 then Exit Sub
  on error goto 0
  response.redirect sActionFileName
end sub


Sub Form_Show()
  
  sWhere = ""
  
  bPK = true

%>
   
   <table style="">
   <form method="POST" action="<%= sFileName %>" name="Form">
   <tr><td style="background-color: #336699; text-align: Center; border-style: outset; border-width: 1" colspan="2"><font style="font-size: 12pt; color: #FFFFFF; font-weight: bold">MyInfo</font></td></tr>
   <% if not (sFormErr = "") then %>
		<tr><td style="background-color: #FFFFFF; border-width: 1" colspan="2"><font style="font-size: 10pt; color: #000000"><%= sFormErr %></font></td></tr>
	 <% end if %>
<% 

  if sFormErr = "" then
    ' Load primary key and form parameters
  else
    ' Load primary key, form parameters and form fields
    fldmember_id = GetParam("member_id")
    fldmember_password = GetParam("member_password")
    fldname = GetParam("name")
    fldlast_name = GetParam("last_name")
    fldemail = GetParam("email")
    fldaddress = GetParam("address")
    fldphone = GetParam("phone")
    fldnotes = GetParam("notes")
    fldcard_type_id = GetParam("card_type_id")
    fldcard_number = GetParam("card_number")
  end if

  
  pmember_id = Session("UserID")
  if IsEmpty(pmember_id) then bPK = False
  sWhere = sWhere & "member_id=" & ToSQL(pmember_id, "Number")
  PK_member_id = pmember_id

  sSQL = "select * from members where " & sWhere


  if bPK and not(sAction = "insert" and sForm = "Form") then 
    ' Open recordset
    openrs rs, sSQL
    fldmember_id = GetValue(rs, "member_id")
    fldmember_login = GetValue(rs, "member_login")
    if sFormErr = "" then
      ' Load data from recordset when form displayed first time
      fldmember_password = GetValue(rs, "member_password")
      fldname = GetValue(rs, "first_name")
      fldlast_name = GetValue(rs, "last_name")
      fldemail = GetValue(rs, "email")
      fldaddress = GetValue(rs, "address")
      fldphone = GetValue(rs, "phone")
      fldnotes = GetValue(rs, "notes")
      fldcard_type_id = GetValue(rs, "card_type_id")
      fldcard_number = GetValue(rs, "card_number")
    end if
  else
    if sFormErr = "" then
      fldmember_id = ToHTML(Session("UserID"))
    end if
  end if

  

  ' Show form field
    %>
    <tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Login</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><%=ToHTML(fldmember_login)%>&nbsp;</font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Password*</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="password" name="member_password" maxlength="20" value="<%= ToHTML(fldmember_password) %>" size="20" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">First Name*</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="name" maxlength="50" value="<%= ToHTML(fldname) %>" size="50" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Last Name*</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="last_name" maxlength="50" value="<%= ToHTML(fldlast_name) %>" size="50" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Email*</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="email" maxlength="50" value="<%= ToHTML(fldemail) %>" size="50" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Address</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="address" maxlength="50" value="<%= ToHTML(fldaddress) %>" size="50" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Phone</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="phone" maxlength="50" value="<%= ToHTML(fldphone) %>" size="50" ></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Notes</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><textarea name="notes" cols="50" rows="5"><%=ToHTML(fldnotes)%></textarea></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Credit Card Type</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><select name="card_type_id"><%= get_options("select card_type_id, name from card_types order by 2",false,true,fldcard_type_id)%></select></font>
       </td>
     </tr>
<tr>
       <td style="background-color: #FFEAC5; border-style: inset; border-width: 0">
         <font style="font-size: 10pt; color: #000000">Credit Card Number</font>
       </td>
       <td style="background-color: #FFFFFF; border-width: 1">
         <font style="font-size: 10pt; color: #000000"><input type="text" name="card_number" maxlength="50" value="<%= ToHTML(fldcard_number) %>" size="50" ></font>
       </td>
     </tr>

    <tr><td colspan="2" align="right">
    

<% if not (bPK and (not (sAction="insert" and sForm="Form"))) then %>
	
<%else%>
 <input type="hidden" value="update" name="FormAction"/>
 
 <% if bPK then %>
   <input type="submit" value="Update" onclick="document.Form.FormAction.value = 'update';">
 <% end if %>
 
<%end if%>

 <input type="submit" value="Cancel" onclick="document.Form.FormAction.value = 'cancel';">
   <input type="hidden" name="FormName" value="Form">
  
  <input type="hidden" name="PK_member_id" value="<%= pmember_id %>">  
  <input type="hidden" name="member_id" value="<%= ToHTML(fldmember_id)%>">
  </td></tr>
  </form>
  </table>
<%
  

End Sub

%>