<%
'=============================================================
'=             Copyright (c) 2010 è��(QQ:77068320)          =
'=                  All rights reserverd.                    =
'=============================================================
'=                  URL ת��ϵͳ v_1.10.614                  =
'=        ��ʾ��ַ��http://url.myw3.cn                       =
'=        �ٷ���վ��http://www.myw3.cn/myDevise/Url/         =
'=        ���߲��ͣ�http://www.miaoqiyuan.cn                 =
'=        ������Ʒ��http://www.myw3.cn/myDevise/             =
'=============================================================
'=   �����飺��������һ��ASP+Access������С�ɵ�URLת��ϵ   =
'= ͳ�����ذ���22KB���ҡ����ܺ����������ҵ�񣬵����Զ��ر�  =
'= ���趨����ת���ȹ��ܡ�                                    =
'=   ���޸�#/inc.asp��masterWebΪ���Ĺ����ַ��              =
'=============================================================
'=  �ļ���#/admin/url.Class.asp                              =
'=  ���ܣ��û�ת���б�                                       =
'=============================================================


  Class urlClass
    Dim PageSize
    Public Sub init(byval idx)
      Url = "?act="& Req("act") & "&fun=" & Req("fun") & "&page="
      PageSize = 100
      if idx="dme" then
        id = Req("id")
        if id = "" or not isnumeric(id) then id=0
        id = CLng(id)
        Rs.Open "select d_id from [domain] where d_id="&id,Conn,1,1
        if Rs.Eof then WebErr "û���ҵ���ҵ��",1,-1
          DmeWhere = "d_id="&id
          Url = "?act="& Req("act") & "&fun=" & Req("fun") & "&id=" & Req("fun") & "&page="
        Rs.Close
      end if
      Call Sub_Lst(DmeWhere,"u_id desc",Url)
    End Sub
    
    Private Sub Sub_Lst(byval where,byval order,byval url)
%>  <table cellspacing="0" cellpadding="0" class="list" width="100%">
      <tr class="tit">
        <td>��������</td>
        <td>ת����ַ</td>
        <td>ת������</td>
        <td>��ѡ����</td>
      </tr>
<%
      Dim Sql
      Sql = "Select * from [url],[domain] where d_id=u_did"
      if where<>"" then Sql = Sql & " and "&Where
      if order<>"" then Sql = Sql & " order by "&Order
      Rs.Open Sql,Conn,1,1
      if Rs.Eof then
%>      <tr class="tit">
        <td colspan="4">��ô�ҵ���ؼ�¼</td>
      </tr>
<%
      else
        rs.PageSize = PageSize
        pgCount = rs.PageCount
        rsCount = rs.RecordCount
        iPage = Req("page")
        if not isnumeric(iPage) or iPage = "" then iPage = 1
        iPage = cLng(iPage)
        if iPage > pgCount then iPage = pgCount
        rs.absolutepage = iPage
        do while not rs.eof and i<PageSize
          i = i + 1
%>      <tr>
        <td><%=Rs("u_res")%></td>
        <td><%=Rs("u_url")%></td>
        <td align="center"><%if Rs("u_type")=1 then%>����<%else%>��ת<%end if%></td>
        <td><a href="?act=lst&fun=edt&id=<%=Rs("d_id")%>">��������</a></td>
      </tr>
<%
          rs.movenext
        loop
        if pgCount>1 then
%>      <tr class="tit">
        <td colspan="4"><%=csPage(iPage,pgCount,url,"")%> ��<%=iPage%>ҳ/��<%=pgCount%>ҳ  ÿҳ<%=PageSize%>����¼/��<%=rsCount%>����¼</td>
      </tr>
<%
        end if
      end if
      rs.close
      set rs = nothing
%>    </table><%
    End Sub
  End Class
%>