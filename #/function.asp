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
'=  文件：#/function.asp                                     =
'=  功能：公共函数，定义了一些常用的函数。                   =
'=============================================================

  Function Req(byval q)
    Dim Query
    Query=Request(q)
    Query=replace(Query,"'","")
    Query=nohtml(Query)
    if not isdate(Query) and instr(q,"_info")=0 then
      Query=Replace(Query," ","")
    end if
    Req=Trim(Query)
  End Function
  
  function nohtml(byval str)
    dim re
    Set re=new RegExp
    re.IgnoreCase =true
    re.Global=True
    re.Pattern="(<.[^<]*>)"
    str=re.replace(str," ")
    nohtml=str
    set re=nothing
  end function
  
  Sub WebErr(byval s,byval t,byval k)
    dim myUrl
    if t=1 then
      myUrl="history.go(" & k & ")"
    else
      myUrl="location.href='" & k & "'"
    end if
%>
<script type="text/javascript">
alert('<%=replace(s,"'","\'")%>');
<%=myUrl%>;
</script>
<%
    Response.end
  End Sub
%>