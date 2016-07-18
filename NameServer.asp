<!--#include file="#/inc.asp"-->
<!--#include file="#/function.asp"--><%
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
'=  功能：URL转发用户管理页面                                =
'=============================================================

	Dim hostName
	hostName=request.servervariables("SERVER_NAME")
	
	hostName=Lcase(hostName)
	if hostName<> masterWeb then
		server.transfer("404.html")
		response.end
	end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<title>域名转发系统</title>
<style type="text/css">
Body{background:#C0C0C0;}
.Box{
	position:absolute;left:50%;top:100px;
	width:600px;margin-left:-300px;
	background:#FFFFFF;
	border:solid 2px #1F7208;
}
.Box.wid300{
	width:300px;margin-left:-150px;
}
.Box h2{
	margin:0px;padding:0px;
	height:30px;line-height:30px;
	color:#FFFFFF;font-size:16px;
	text-indent:10px;
	background:#1F7208;
}
.Body form{margin:0px;padding:0px;}
.Body input{vertical-align:middle;}
.Body input.input{
	border:solid 1px #CCC;
	background:#EEE;
	color:#1F7208;
	width:151px;
	height:20px;
	line-height:20px;
}
.Body table td{
	height:25px;line-height:30px;
	font-size:12px;
	background:#EFEFEF;
}
.Body table.info td{
	height:18px;line-height:20px;
	text-indent:20px;
}
.Body table td.title{
	text-align:right;
}
.Body table tr.title td{
	text-align:center;
	background:#1F7208;
	color:#FFF;
}
.right{float:right}
.right a{color:#FFF;text-decoration:none;padding-right:10px;}
</style>
<head>
</head>
<body>
<%
function getRes(byval d)
	Res = Lcase(replace(d,Session("LoginDomain"),""))
	if Len(Res)=0 then
		getRes="@"
	else
		getRes=Left(Res,Len(Res)-1)
	end if
end function

function checkUrl(byval Url,byref ErrTxt)
	if LCase(left(Url,7))<>"http://" then
		ErrTxt="转发到的地址必须以Http://开头"
	else
		'预留转发黑名单
		checkUrl=1
		Exit Function
	end if
	checkUrl=-1
end function

Sub LoginUI
%>
	<div class="Box wid300">
		<h2>登录系统</h2>
		<div class="Body">
			<form action="?act=Login" method="post">
				<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td class="title">域名：</td>
						<td><input class="input" name="domain" /></td>
					</tr>
					<tr>
						<td class="title">密码：</td>
						<td><input class="input" name="password" type="password"/></td>
					</tr>
					<tr>
						<td colspan="2" align="center"/>
							<input class="button" type="submit" value="登录" />
						</td>
					</tr>
				</table>
			</form>
		</div>
	</div>
<%
End Sub

Sub Main
	Domain=Session("LoginDomain")
	set Rs=Conn.Execute("select * from [domain] where [d_dme]='"&Domain&"'")
%>
	<div class="Box">
		<h2><span class="right"><%if Session("LoginDomain")="local" then%><a href="NameServer_Admin.asp">[后台管理]</a><%end if%><a href="?act=LoginOut">[退出]</a></span>域名转发系统<%=Session("Domain")%></h2>
		<div class="Body">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="title">开通日期：</td>
					<td><span><%=Rs("d_statime")%></span></td>
					<td class="title">购买条数：</td>
					<td><span><%=Rs("d_max")%></span>条</td>
				</tr>
				<tr>
					<td class="title">到期日期：</td>
					<td><span><%=Rs("d_expries")%></span></td>
					<td class="title">已经使用：</td>
					<td><span><%=Rs("d_count")%></span>条</td>
				</tr>
				<tr>
					<td class="title">修改密码：</td>
					<td colspan="3"><form action="?act=ChgPwd" method="Post"> <input name="pwd" class="input" style="width:100px;"/> <input type="submit" value="修改" /></form></td>
				</tr>
			</table>
		</div>
		<h2><%=Domain%> URL转发记录</h2>
		<div class="Body">
			<table width="100%">
				<tr class="title">
					<td>主机名</td>
					<td>转发地址</td>
					<td>显示标题</td>
					<td>转发类型</td>
					<td>&nbsp;</td>
				</tr>
<%
		DomainID=Rs("d_id")
		DomainMx=Rs("d_max")
		DomainCt=Rs("d_count")
	Rs.Close
	Set Rs=nothing
	set Rs=Conn.Execute("select * from [url] where [u_did]="&DomainID)
	do while not Rs.eof
%>				<form action="?act=Save" method="post" onsubmit="if(this.res.value=='')this.res.value='@'">
				<tr align="center">
					<td><input name="res" value="<%=getRes(Rs("u_res"))%>" class="input" style="width:60px;text-align:right"/>.<%=Domain%><input name="oldres" type="hidden" value="<%=Rs("u_res")%>" /></td>
					<td><input name="url" value="<%=Rs("u_url")%>"  class="input" style="width:100px;"/></td>
					<td><input name="title" value="<%=Rs("u_title")%>" class="input" style="width:100px;"/></td>
					<td>
						<select name="type"/>
							<option value="1">隐藏URL</option>
							<option value="0" <%if Rs("u_type")=0 then response.write "selected=""selected""" end if%>>不隐藏</option>
						</select>
					</td>
					<td>
						<input type="submit" value="修改"/>
						<input type="button" value="删除" onclick="if(confirm('您真的要删除主机名为\'<%=getRes(Rs("u_res"))%>\'记录吗？'))location.href='?act=Del&oldres=<%=Rs("u_res")%>'"/>
					</td>
				</tr>
				</form>
<%
		rs.movenext
	loop
	Rs.Close
	Set Rs=nothing
	if DomainMx>DomainCt then
%>				<form action="?act=Save" method="post" onsubmit="if(this.res.value=='')this.res.value='@'">
				<tr align="center">
					<td><input name="res" value="" class="input" style="width:60px;text-align:right"/>.<%=Domain%></td>
					<td><input name="url" value=""  class="input" style="width:100px;"/></td>
					<td><input name="title" value="" class="input" style="width:100px;"/></td>
					<td>
						<select name="type"/>
							<option value="1">隐藏URL</option>
							<option value="0">不隐藏</option>
						</select>
					</td>
					<td>
						<input type="submit" value="添加"/>
					</td>
				</tr>
				</form>
<%
	end if
%>
			</table>
		</div>
		<h2>URL转发说明</h2>
		<div class="Body">
			<table width="100%" class="info">
				<tr>
					<td>第一步：将您要做转发的域名(<%=Domain%>)解析到 <span style="color:#C00"><%=hostName%></span>，比如www.<%=Domain%>，则需要添加www.<%=Domain%>的cname记录，值为<span style="color:#C00"><%=hostName%></span></td>
				</tr>
				<tr>
					<td>第二步：登录本系统，添加相应的记录。比如www.<%=Domain%>转发到http://www.baidu.com，主机名填：www，转发地址填：http://www.baidu.com，显示标题：百度网，转发类型：隐藏</td>
				</tr>
				<tr>
					<td style="color:#666">备注：如果是主域名，一般不支持cname，必须做A记录。比如<%=Domain%>转发到http://www.baidu.com，需要做<%=Domain%>的A记录到<%=hostName%>的IP</td>
				</tr>
			</table>
		</div>
	</div>
<%
End Sub

Sub Delete
	Domain=Session("LoginDomain")
	oldRes=Req("oldres")
	Conn.Execute("delete from [url] where [u_res]='"&oldRes&"' and u_did="&Conn.Execute("select [d_id],[d_count],[d_max] from [domain] where [d_dme]='"&Domain&"'")(0))
	Call updateDomainCount()
	WebErr "删除成功。",-1,"?act=Main"
End Sub

Sub ChgPwd
	Domain=Session("LoginDomain")
	pass=Req("pwd")
	if pass="" then WebErr "密码不能为空。",-1,"?act=Main"
	Conn.Execute("update [domain] set [d_pass]='"&pass&"' where [d_dme]='"&Domain&"'")
	WebErr "密码修改成功。",-1,"?act=Main"
End Sub

Sub Save
	u_Url=Req("url")
	if CheckUrl(u_Url,ErrTxt)=-1 then  WebErr ErrTxt,1,-1

	Domain=Session("LoginDomain")
	set Rs=Conn.Execute("select [d_id],[d_count],[d_max] from [domain] where [d_dme]='"&Domain&"'")
	d_id=Rs(0)
	d_Count=Rs(1)
	d_Max=Rs(2)
	Rs.Close
	
	Set Rs=Server.CreateObject("ADODB.Recordset")
	u_Res=Req("res")
	if len(u_Res)=0 then WebErr "主机名不能为空。",1,-1
	oldRes=Req("oldres")
	if u_Res="@" then
		u_Res= Domain
	else
		u_Res= u_Res & "." & Domain
	end if
	
	if Conn.Execute("select count([u_id]) from [url] where [u_res]='"&u_Res&"' and [u_res]<>'"&oldRes&"'")(0)>0 then WebErr "URL('"&u_Res&"')重复，请重试。",1,-1
	Set Rs=Server.CreateObject("ADODB.Recordset")
	if oldRes="" then
		if d_Max<d_Count then WebErr "您的URL转发记录已经超过最大记录，请购买URL转发条数。",1,-1
		Rs.Open "select * from [url]",Conn,3,2
		Rs.AddNew
		Rs("u_did")=d_id
	else
		Rs.Open "select * from [url] where [u_res]='"&oldRes&"' and u_did="&d_id,Conn,3,2
		if Rs.Eof then WebErr "您没有权限操作该域名。",1,-1
	end if
	Rs("u_title")=Req("title")
	Rs("u_res")=u_Res
	Rs("u_url")=u_Url
	if Req("type")="1" then
		Rs("u_type")=1
	else
		Rs("u_type")=0
	end if
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Call updateDomainCount()
	WebErr "操作成功。",-1,"?act=Main"
End Sub

Sub updateDomainCount()
	Conn.Execute("update [domain] set d_count="&Conn.Execute("SELECT Count([u_id]) FROM [url],[domain] WHERE [d_id]=[u_did] and [d_dme]='"&Session("LoginDomain")&"'")(0)&" where [d_dme]='"&Session("LoginDomain")&"'")
End Sub

Conn.Open Constr
if Session("LoginDomain")="" then
	if Req("domain")="" or Req("password")="" then
		Call LoginUI
	Else
		Set Rs=Conn.Execute("select [d_pass] from [domain] where d_dme='"&Req("domain")&"'")
		if Rs.eof then WebErr "该业务不存在",1,-1
		if Rs(0)<>Req("password") then WebErr "域名密码错误",1,-1
		Rs.Close
		Session("LoginDomain")=Req("domain")
		Response.Redirect "?act=Main"
	End if
else
	Select Case Req("act")
		Case "Main"
			Call Main
		Case "Save"
			Call Save
		Case "ChgPwd"
			Call ChgPwd
		Case "Del"
			Call Delete
		Case "LoginOut"
			Session.Contents.Removeall()
			Response.Redirect "?act=Main"
		Case Else
			Response.Redirect "?act=Main"
	End Select
end if

Conn.Close
%>
</body>
</html>