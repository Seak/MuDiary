<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Session.Contents("strUserName") & "���ռǱ�����д�ռ�" %>
<%
If Not(Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName"))) Then
	Response.Redirect("index.asp")
	Response.End()
End If

strUserName = Trim(Session.Contents("strUserName"))
dtmDateTime = Trim(Request.Form("Year") & "-" & Request.Form("Month") & "-" & Request.Form("Day"))
strWeather = Trim(Request.Form("Weather"))
strTitle = Trim(Request.Form("Title"))
strContent = Trim(Request.Form("Content"))
blnSecrecy = Request.Form("Secrecy")
strIP = Request.ServerVariables("Remote_Addr")

If strUserName <> "" And dtmDateTime <> "" And strWeather <> "" And strTitle <> "" And strContent <> "" Then
	ConnectionDatabase
	strSQL = "Insert Into [Diary] (strUserName, dtmDateTime, strWeather, strTitle, strContent, blnSecrecy, strIP) Values ('"& strUserName &"', '"& dtmDateTime &"', '"& strWeather &"', '"& strTitle &"', '"& strContent &"', '"& blnSecrecy &"', '"& strIP &"')"
	Conn.Execute(strSQL)
	Set strSQL = Nothing
	CloseDatabase
	Title = Session.Contents("strUserName") & "���ռǱ�����д�ɹ�"
	strMsg = "<div align=""center""><a href=""diary.asp?username=" & Session.Contents("strUserName") & """ onFocus=""this.blur()"">�鿴�ռǱ�</a></div>"
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
          <td><% If strMsg <> "" Then %><%= strMsg %><% Else %><form name="write" method="post" action="write.asp">
              ���ڣ� 
                <select name="Year" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt">
                  <% For dtmYear = Year(Now) - 1 To Year(Now) %>
                  <option value="<%= dtmYear %>"<% If dtmYear = Year(Now) Then Response.Write(" selected") %>><%= dtmYear %></option>
                  <% Next 'dtmYear %>
                </select>
                �� 
                <select name="Month" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt">
                  <% For dtmMonth = 1 To 12 %>
                  <option value="<%= dtmMonth %>"<% If dtmMonth = Month(Now) Then Response.Write(" selected") %>><%= dtmMonth %></option>
                  <% Next 'dtmMonth %>
                </select>
                �� 
                <select name="Day" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt">
                  <% For dtmDay = 1 To 31 %>
                  <option value="<%= dtmDay %>"<% If dtmDay = Day(Now) Then Response.Write(" selected") %>><%= dtmDay %></option>
                  <% Next 'dtmDay %>
                </select>
              �ա������������� 
              <input name="Secrecy" type="radio" value="0" checked>
              �����ܣ� 
              <input type="radio" name="Secrecy" value="1">
              <br>
                ������ 
                <input name="Weather" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="15">
                <font color="#FF0000">*</font>�硢�ơ������硢�ꡢѩ���ס�������ȵ�<br>
                ���⣺ 
                <input name="Title" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" size="30">
                <font color="#FF0000">*</font>��Ҫ����15���ַ�<br>
                ���ݣ� 
                <textarea name="Content" cols="70" rows="8" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt"></textarea>
                <br>
              <div align="center">
                <input type="submit" name="Submit" value="�ύ" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt">��
                <input type="reset" name="Reset" value="����" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt">
              </div>
            </form><% End If %></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->