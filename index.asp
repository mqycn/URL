<!--#include file="#/inc.asp"--><%
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
'=  �ļ���index.asp                                          =
'=  ���ܣ�URL����ҳ�档                                      =
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