<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Session.Contents("strUserName") & "的日记本—资料修改" %>
<%
If Not(Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName"))) Then
	Response.Redirect("index.asp")
	Response.End()
End If

strPassword = Trim(Request.Form("Password"))
strMail = Trim(Request.Form("Mail"))

If strPassword <> "" And strMail <> "" Then
	ConnectionDatabase
	strSQL = "Update [UserList] Set strPassword = '"& strPassword &"', strMail = '"& strMail &"' Where strUserName = '" & Session.Contents("strUserName") & "'"
	Conn.Execute(strSQL)
	Set strSQL = Nothing
	Set Rs = Nothing
	CloseDatabase

	Title = Session.Contents("strUserName") & "的日记本—修改成功"
	strMsg = "您的资料修改成功！<br><a href=""diary.asp?username=" & Session.Contents("strUserName") & """ onFocus=""this.blur()"">查看日记本</a>"
End If
%>
<!--#include file="inc_header.asp"-->
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;"> 
<tr> 
  <td width="600" height="30" align="center" valign="middle" background="img/td_bg.gif"><%= Title %></td>
</tr>
<tr> 
  <td align="center" valign="middle"> <br> <table width="500" border="0" align="center" cellpadding="0" cellspacing="0" style="Word-Break:break-all; BORDER: 1px solid #E2E2E2;"> 
    <tr> 
      <td> <div align="center"> 
<% If strMsg <> "" Then %><%= strMsg %><% Else %><%
ConnectionDatabase
strSQL = "Select strUserName, strPassword, strMail From [UserList] Where strUserName = '" & Session.Contents("strUserName") & "'"
Set Rs = Conn.Execute(strSQL)
%><form name="modify" method="post" action="modify.asp"> 
        昵称： <input name="UserName" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strUserName") %>" size="15" ReadOnly>
                <br>
                密码： 
                <input name="Password" type="password" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strPassword") %>" size="15">
                <br>
                信箱：
                <input name="Mail" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strMail") %>" size="15">
                <br>
                <input name="Submit" type="submit" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="提交">　<input name="Reset" type="reset" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="重置">
            </form>
<%
Set Rs = Nothing
Set strSQL = Nothing
CloseDatabase
%>
<% End If %>
              </div>
          </td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->