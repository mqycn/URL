<%
'=============================================================
'=             Copyright (c) 2010 猫七(QQ:77068320)          =
'=                  All rights reserverd.                    =
'=============================================================
'=                  URL 转发系统 v_1.10.614                  =
'=        演示地址：http://url.myw3.cn                       =
'=        官方网站：http://www.myw3.cn/myDevise/Url/         =
'=        作者博客：http://www.miaoqiyuan.cn                 =
'=        其他作品：http://www.myw3.cn/myDevise/             =
'=============================================================
'=   程序简介：本程序是一款ASP+Access开发的小巧的URL转发系   =
'= 统，下载包仅22KB左右。功能涵盖了添加新业务，到期自动关闭  =
'= ，设定域名转发等功能。                                    =
'=   请修改#/inc.asp中masterWeb为您的管理地址。              =
'=============================================================
'=  文件：#/admin/url.Class.asp                              =
'=  功能：用户转发列表                                       =
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
        if Rs.Eof then WebErr "没有找到给业务",1,-1
          DmeWhere = "d_id="&id
          Url = "?act="& Req("act") & "&fun=" & Req("fun") & "&id=" & Req("fun") & "&page="
        Rs.Close
      end if
      Call Sub_Lst(DmeWhere,"u_id desc",Url)
    End Sub
    
    Private Sub Sub_Lst(byval where,byval order,byval url)
%>  <table cellspacing="0" cellpadding="0" class="list" width="100%">
      <tr class="tit">
        <td>所属域名</td>
        <td>转发地址</td>
        <td>转发类型</td>
        <td>可选操作</td>
      </tr>
<%
      Dim Sql
      Sql = "Select * from [url],[domain] where d_id=u_did"
      if where<>"" then Sql = Sql & " and "&Where
      if order<>"" then Sql = Sql & " order by "&Order
      Rs.Open Sql,Conn,1,1
      if Rs.Eof then
%>      <tr class="tit">
        <td colspan="4">怎么找到相关记录</td>
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
        <td align="center"><%if Rs("u_type")=1 then%>隐藏<%else%>跳转<%end if%></td>
        <td><a href="?act=lst&fun=edt&id=<%=Rs("d_id")%>">域名密码</a></td>
      </tr>
<%
          rs.movenext
        loop
        if pgCount>1 then
%>      <tr class="tit">
        <td colspan="4"><%=csPage(iPage,pgCount,url,"")%> 第<%=iPage%>页/共<%=pgCount%>页  每页<%=PageSize%>条记录/共<%=rsCount%>条记录</td>
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