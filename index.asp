<!--#include file="#/inc.asp"--><%
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
'=  文件：index.asp                                          =
'=  功能：URL控制页面。                                      =
'=============================================================

	Dim hostName
	hostName=request.servervariables("SERVER_NAME")
	
	hostName=Lcase(hostName)
	if hostName=masterWeb then response.redirect "/NameServer.asp"
	hostName=replace(hostName,"'","")
	
	conn.open constr
	set rs=server.createobject("ADODB.recordset")
	rs.open "select url,utype,udme,title from rewrite where hostname='"&hostName&"' and expries>#"&date()&"#",conn,3,1
	if not rs.eof then
		u_url=rs(0)
		u_tit=rs(2)
		u_type=rs(1)
		title=rs(3)
		if u_type=1 then
%>
<!--#include file="#/Template/frame.html"-->
<%		else%>
<!--#include file="#/Template/redirect.html"-->
<%		end if
	else
%><!--#include file="#/Template/404.html"-->
<%
	end if
%>