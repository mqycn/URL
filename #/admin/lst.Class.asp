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
'=  文件：#/admin/lst.Class.asp                              =
'=  功能：转发域名管理类                                     =
'=============================================================

  Class lstClass
    Dim PageSize
    Public Sub init(byval idx)
      Url = "?act="& Req("act") & "&fun=" & Req("fun") & "&page="
      PageSize = 10
      Select Case idx
        Case "expiresmonth"
          Call Sub_Lst("d_expries<#" & date()+31 & "# and d_expries>#" & date() & "#","d_expries desc",Url)
        Case "expiresweek"
          Call Sub_Lst("d_expries<#" & date()+7 & "# and d_expries>#" & date() & "#","d_expries desc",Url)
        Case "buymonth"
          Call Sub_Lst("d_statime>#" & date()-31 & "# and d_statime<#" & date() & "#","d_statime desc",Url)
        Case "buyweek"
          Call Sub_Lst("d_statime>#" & date()-7 & "# and d_statime<#" & date() & "#","d_statime desc",Url)
        Case "bad"
          Call Sub_Lst("d_expries<#" & date() & "#","d_statime desc",Url)
        Case "add"
          Call Sub_Edit(0)
        Case "edt"
          id = Req("id")
          if isnumeric(id) and id<>"" then
            Call Sub_Edit(id)
          Else
            Response.Redirect "?act=" & Req("act")
          end if
        Case "del"
          id = Req("id")
          if isnumeric(id) and id<>"" and id<>"0" then
            Call Sub_Del(id)
          Else
            Response.Redirect "?act=" & Req("act")
          end if
        Case "sav"
          Call Sub_Sav()
        Case Else
          Call Sub_Lst("","d_statime desc",Url)
      End Select
    End Sub
    
    Private Sub Sub_Del(id)
      Rs.Open "select d_dme from [domain] where d_id="&id,Conn,1,1
      if Rs.Eof then WebErr "要操作的业务不存在",1,-1
      d_dme = rs("d_dme")
      Rs.Close
      Rs.Open "delete from [domain] where d_id="&id,Conn,3,2
      WebErr "删除" & d_dme & "成功",-1,"?act=lst&fun=edt&id=" & id
    End Sub
    
    Private Sub Sub_Sav()
      d_dme = Req("d_dme")
      id = Req("id")
      if not isnumeric(id) or id="" then Response.Redirect "?act=lst"
      
      d_pass = Req("d_pass")
      if len(d_pass)<6 or len(d_pass)>16 then WebErr "域名密码必须为6-16位，您输入的为" & len(d_pass) & "位。",1,-1
      
      d_max = Req("d_max")
      d_count = Req("d_count")
      if not isnumeric(d_max) or d_max="" then WebErr "购买条数必须是一个数字",1,-1
      d_max = cLng(d_max)
      if d_max <1 then WebErr "购买条数必须大于1",1,-1
      d_statime = Req("d_statime")
      d_expries = Req("d_expries")
      if (not isdate(d_expries)) or (not isdate(d_statime)) then WebErr "开通日期、到期日期必须都为一个日期",1,-1
      
      id = cLng(id)
      Rs.Open "select d_id from [domain] where d_dme='"& d_dme &"' and d_id<>"&id,Conn,1,1
      if not Rs.Eof then WebErr d_dme & "已经购买过我们的服务",1,-1
      Rs.Close
      
      if id=0 then
        Rs.Open "select * from [domain]",Conn,3,2
        Rs.AddNew
        rs("d_dme") = Req("d_dme")
        id = Rs("d_id")
      else
        Rs.Open "select * from [domain] where d_id="&id,Conn,3,2
      end if
      Rs("d_pass") = d_pass
      Rs("d_max") = d_max
      Rs("d_statime") = d_statime
      Rs("d_expries") = d_expries
      Rs("d_info") = Req("d_info")
      rs.Update
      rs.Close
      WebErr "操作成功",-1,"?act=lst&fun=edt&id=" & id
    End Sub
    
    Private Sub Sub_Edit(byval id)
    if id=0 then
      d_count = 0
      d_max = 2
      d_statime = formatdatetime(now,0)
      d_expries = formatdatetime(now + 365,0)
    else
      Rs.Open "select * from [domain] where d_id="&id,Conn,1,1
      if Rs.Eof then Response.Redirect "?act=lst"
      d_dme = Rs("d_dme")
      d_pass = Rs("d_pass")
      d_count = Rs("d_count")
      d_max = Rs("d_max")
      d_statime = Rs("d_statime")
      d_expries = Rs("d_expries")
      d_info = Rs("d_info")
      rs.Close
    end if
%>  <table cellspacing="0" cellpadding="0" class="list" width="100%">
      <form action="?act=lst&fun=sav&id=<%=id%>" method="post">
      <tr>
        <td class="tit">域名：</td>
        <td><input name="d_dme" value="<%=d_dme%>"<%if d_dme<>"" then response.write "readonly=""readonly""" end if%> class="inp"></td>
      </tr>
      <tr>
        <td class="tit">密码：</td>
        <td><input name="d_pass" value="<%=d_pass%>" class="inp"></td>
      </tr>
      <tr>
        <td class="tit">已用条数：</td>
        <td><input value="<%=d_count%>" readonly="readonly" class="inp"></td>
      </tr>
      <tr>
        <td class="tit">购买条数：</td>
        <td><input name="d_max" value="<%=d_max%>" class="inp"></td>
      </tr>
      <tr>
        <td class="tit">开通日期：</td>
        <td><input name="d_statime" value="<%=d_statime%>" class="inp"></td>
      </tr>
      <tr>
        <td class="tit">到期日期：</td>
        <td><input name="d_expries" value="<%=d_expries%>" class="inp"></td>
      </tr>
      <tr>
        <td class="tit">备注信息：</td>
        <td><textarea name="d_info" class="info"><%=d_info%></textarea></td>
      </tr>
      <tr class="tit" >
        <td colspan="2"><input value="保存" type="submit">&nbsp;<input value="重填" type="reset"></td>
      </tr>
      </form>
    </table><%
    End Sub
    
    Private Sub Sub_Lst(byval where,byval order,byval url)
%>  <table cellspacing="0" cellpadding="0" class="list" width="100%">
      <tr class="tit">
        <td>域名</td>
        <td>使用条数</td>
        <td>购买条数</td>
        <td>开通时间</td>
        <td>到期时间</td>
        <td>可选操作</td>
      </tr>
<%
      Dim Sql
      Sql = "Select * from [domain]"
      if where<>"" then Sql = Sql & " where "&Where
      if order<>"" then Sql = Sql & " order by "&Order
      Rs.Open Sql,Conn,1,1
      if Rs.Eof then
%>      <tr class="tit">
        <td colspan="6">怎么找到相关记录</td>
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
        <td><%=Rs("d_dme")%></td>
        <td><%=Rs("d_count")%></td>
        <td><%=Rs("d_max")%></td>
        <td><%=Rs("d_statime")%></td>
        <td><%=Rs("d_expries")%></td>
        <td><a href="?act=lst&fun=edt&id=<%=Rs("d_id")%>">修改</a> <a href="?act=lst&fun=del&id=<%=Rs("d_id")%>">删除</a> <a href="?act=url&fun=dme&id=<%=Rs("d_id")%>">URL</a></td>
      </tr>
<%
          rs.movenext
        loop
        if pgCount>1 then
%>      <tr class="tit">
        <td colspan="6"><%=csPage(iPage,pgCount,url,"")%> 第<%=iPage%>页/共<%=pgCount%>页  每页<%=PageSize%>条记录/共<%=rsCount%>条记录</td>
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