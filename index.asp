<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "����ҳ" %>
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
          <td>��������������ֻ��������������������дЩʲô������ͿѻҲ�ã���������Ҳ�ã�����Ҳ�ã�����Ҳ�ã�д���ˣ��������߹��ĵ��εΣ�Ҳ��û��̫�����壬ֻ��������һЩʲô�������ͷ׵����籣��һ�������밲�顣<br>
            ���������������еľ����ϸ��Լ�����һ˿�ռ䣬���ٸе��޼ҿɹ飬�������Ƿ����塣��ʱ����Ҫ��˵����ʱ����Ҫ�����Լ�����ʱ��ֻ��Ҫ֤���Լ����ǻ��ߡ����ˣ���Щ���㹻�ˣ�<% If Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName")) Then %><br>
            ���������Լ����ռǱ���[<a href="diary.asp?username=<%= Session.Contents("strUserName") %>" onFocus="this.blur()">����</a>]��<% End If %></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->