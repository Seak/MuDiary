<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "���������" %>
<%
strUserName = Trim(Request.Form("UserName"))
strPassword = Trim(Request.Form("Password"))
strMail = Trim(Request.Form("Mail"))

If strUserName <> "" And strPassword <> "" And strMail <> "" Then
	ConnectionDatabase
	strSQL = "Select strUserName, strPassword, strMail, strRegIP, dtmDateTime From [UserList] Where strUserName='"& strUserName &"'"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open strSQL, Conn, 3, 2
	If Rs.BOF Or Rs.EOF Then
		Rs.AddNew
		Rs("strUserName") = strUserName
		Rs("strPassword") = strPassword
		Rs("strMail") = strMail
		Rs("strRegIP") = Request.ServerVariables("Remote_Addr")
		Rs("dtmDateTime") = Now()
		Rs.Update

		Title = Pro_Name & "������ɹ�"
		strMsg = "��ӭ��" & strUserName & "���������ǣ�����[<a href=""login.asp"" onFocus=""this.blur()"">��½</a>]��"
		Session.Contents("strUserName") = strUserName
	Else
		Title = Pro_Name & "��������Ϣ"
		strMsg = "��" & strUserName & "����ע�ᣡ"
	End If
	Rs.Close
	Set strSQL = Nothing
	Set Rs = Nothing
	CloseDatabase
End If
%>
<!--#include file="inc_header.asp"-->
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;">
  <tr>
    <td width="600" height="30" align="center" valign="middle" background="img/td_bg.gif"><%= Title %></td>
  </tr>
  <tr> 
    <td align="center" valign="middle">
	 <br>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0" style="Word-Break:break-all; BORDER: 1px solid #E2E2E2;">
        <tr>
          <td>
              <div align="center">
<% If Session.Contents("strUserName") = "" Then %>
              <form name="join" method="post" action="join.asp">
                <%= strMsg %><br>
                �ǳƣ� 
                <input name="UserName" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <br>
                ���룺 
                <input name="Password" type="password" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <br>
                ���䣺
                <input name="Mail" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <br>
                <input name="Submit" type="submit" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="�ύ">��<input name="Reset" type="reset" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="����">
            </form>
<% Else %>
<%
Session.Contents("strUserName") = ""
Response.Write(strMsg)
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