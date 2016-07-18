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
'=  文件：#/inc.asp                                          =
'=  功能：配置页面，配置管理地址和数据库。                   =
'=============================================================

  Dim conStr
  Dim masterWeb
  
  masterWeb = "localhost"    '管理地址
  conStr="provider=microsoft.jet.oledb.4.0;data source="&server.mappath("#/url.mdb")
  set conn=server.createobject("adodb.connection")
%>