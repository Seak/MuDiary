<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "����½ϵͳ" %>
<%
strUserName = Trim(Request.Form("UserName"))
strPassword = Trim(Request.Form("Password"))

If strUserName <> "" And strPassword <> "" Then
	ConnectionDatabase
	strSQL = "Select strUserName, strPassword From [UserList] Where strUserName='"& strUserName &"' And strPassword = '" & strPassword & "'"
	Set Rs = Conn.Execute(strSQL)
	If Rs.BOF Or Rs.EOF Then
		Title = Pro_Name & "��������Ϣ"
		strMsg = "�ǳơ�" & strUserName & "��δע�ᣬ���������"
	Else
		Session.Contents("strUserName") = strUserName
		Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName")) = True
		Title = Pro_Name & "����½�ɹ�"
		strMsg = "��ӭ��" & strUserName & "���ĵ�����<br>�����Լ����ռǱ���[<a href=""diary.asp?username=" & strUserName & """ onFocus=""this.blur()"">����</a>]��"
	End If
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
<% If Not(Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName"))) Then %><form name="login" method="post" action="login.asp"><%= strMsg %><br>
                �ǳƣ� 
                <input name="UserName" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <br>
                ���룺 
                <input name="Password" type="password" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <br>
                <input name="Submit" type="submit" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="�ύ">��<input name="Reset" type="reset" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="����">
            </form><% Else %><%= strMsg %><% End If %>
              </div>
          </td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->