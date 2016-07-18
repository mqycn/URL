<!--#include file="#/inc.asp"-->
<!--#include file="#/function.asp"-->
<!--#include file="#/admin/url.Class.asp"-->
<!--#include file="#/admin/lst.Class.asp"-->
<!--#include file="#/admin/sys.Class.asp"-->
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
'=  功能：管理员管理页面                                     =
'=============================================================

  if Session("LoginDomain")<>"local" then Response.Redirect "NameServer.asp"

  function csPage(id,all,url1,url2)
    if id<>1 then tmp="<a href="""&url1&"1"&url2&""">首页</a> <a href="""&url1&(id-1)&url2&""">上一页</a> "
    if id-10>0 then tmp=tmp&" <a href="""&url1&id-10&url2&""">前10页</a> "
    istart=((id-1)\10)*10+1
    if(all-id)>9 and i>10 then
      iend=((id-1)\10)*10+10
    else
      istart=all-(all mod 10)+1
      iend=all
    end if
    for i=istart to iend
      if i=id then
        tmp=tmp&"<span>"&i&"</span> "
      else
        tmp=tmp&"<a href="""&url1&i&url2&""">"&i&"</a> "
      end if
    next
    if all-id>10 then tmp=tmp&" <a href="""&url1&id+10&url2&""">后10页</a> "
    if id<>all then tmp=tmp&"<a href="""&url1&(id+1)&url2&""">下一页</a> <a href="""&url1&all&url2&""">尾页</a>"
    csPage=tmp
  end function

  Sub getTopMenu(byval Act)
    'm = "服务设置:sys;URL转发域名:lst;URL转发详情:url;退出系统:sys&fun=loginout;返回前台:sys&fun=user;"
    m = "URL转发域名:lst;URL转发详情:url;退出系统:sys&fun=loginout;返回前台:sys&fun=user;"
    if instr(m,":" & Act & ";") = 0 and Act<>"sys" then Response.Redirect "?act=lst"
    for each i in split(m,";")
      if trim(i)<>"" and instr(i,":")>0 then
        t = split(i,":")
        if Act = t(1) then
          Response.write "<li class=""hover""><a href=""?act=" & t(1) & """>" & t(0) & "</a></li>"
        else
          Response.write "<li><a href=""?act=" & t(1) & """>" & t(0) & "</a></li>"
        end if
      end if
    next
  End Sub

  Sub getLeftMenu(byval Act,byref Fun)
    Select Case Act
      Case "sys"
        if Fun = "loginout" then
        	Session.contents.RemoveAll()
        	response.redirect "?act=sys"
        end if
        if Fun = "user" then
        	response.redirect "NameServer.asp"
        end if
        m = "服务器设置:control;"
      Case "lst"
        m = "URL转发域名:lst;本月到期:expiresmonth;本周到期:expiresweek;已过期:bad;本月新购:buymonth;本周新购:buyweek;添加新URL转发域名:add;"
      Case "url"
        m = "URL转发详情列表:lst;"
    End Select
    for each i in split(m,";")
      if trim(i)<>"" and instr(i,":")>0 then
        t = split(i,":")
        if Fun="" then Fun = t(1)
        Response.write "<li><a href=""?act=" & Act & "&fun=" & t(1) & """>" & t(0) & "</a></li>"
      end if
    next
  End Sub
  
  Sub getBody(byval Act,byval Fun)
    Execute("set Web = new "&Act&"Class")
    Execute("Web.init(Fun)")
  End Sub

  conn.open constr
  set rs=server.createobject("ADODB.recordset")
  
  Act = Lcase(Req("act"))
  Fun = Lcase(Req("fun"))
  
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="x-ua-compatible" content="ie=emulateie7" />
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title>URL转发控制面板</title>
<style type="text/css">
html{background:#666}
body{margin:50px auto;width:720px;border:5px #0A7BAA solid;padding:0px;font-size:12px;}
ul,li,p{margin:0px;padding:0px;list-style:none;}
a{text-decoration:none;color:#2AADE4;}
.menu{background:#FFF;height:30px;line-height:30px;}
  .menu li{float:left;}
  .menu a{color:#2AADE4;padding:0px 10px;display:block;font-weight:800;font-size:16px;}
  .menu li.hover a{background:#2AADE4;color:#FFF;}
  .menu.sub{background:#2AADE4;height:25px;line-height:25px;}
    .menu.sub a{color:#FFF;font-weight:400;font-size:12px;}
    .menu.sub li.hover a{background:#FFF;color:#C00;}
.list{background:#FFF;border:1px #666 solid;border-bottom:0px;border-right:0px;}
  .list td{line-height:20px;background:#EEE;border:1px #666 solid;border-top:0px;border-left:0px;color:#666;padding:0px 2px;}
  .list tr.tit td{text-align:center;background:#CCC;line-height:30px;font-weight:600;color:#333;font-size:13px;}
  .list td.tit{width:100px;height:30px;background:#CCC;text-align:right;font-weight:600}
  .list td .inp{border:solid 1px #666;background:#FFF;margin-left:10px;width:200px;}
  .list td .info{border:solid 1px #666;background:#FFF;margin-left:10px;width:500px;height:50px;}
</style>
</div>
</head>
<body>
  <div class="header">
    <ul class="menu"><%getTopMenu Act%></ul>
    <ul class="sub menu"><%getLeftMenu Act,Fun%></ul>
  </div>
  <div class="Main">
  <%getBody Act,Fun%>
  <div>
</body>
</html>