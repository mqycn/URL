<!--#include file="#/inc.asp"-->
<!--#include file="#/function.asp"--><%
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
'=  ���ܣ�URLת���û�����ҳ��                                =
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
<title>����ת��ϵͳ</title>
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
		ErrTxt="ת�����ĵ�ַ������Http://��ͷ"
	else
		'Ԥ��ת��������
		checkUrl=1
		Exit Function
	end if
	checkUrl=-1
end function

Sub LoginUI
%>
	<div class="Box wid300">
		<h2>��¼ϵͳ</h2>
		<div class="Body">
			<form action="?act=Login" method="post">
				<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td class="title">������</td>
						<td><input class="input" name="domain" /></td>
					</tr>
					<tr>
						<td class="title">���룺</td>
						<td><input class="input" name="password" type="password"/></td>
					</tr>
					<tr>
						<td colspan="2" align="center"/>
							<input class="button" type="submit" value="��¼" />
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
		<h2><span class="right"><%if Session("LoginDomain")="local" then%><a href="NameServer_Admin.asp">[��̨����]</a><%end if%><a href="?act=LoginOut">[�˳�]</a></span>����ת��ϵͳ<%=Session("Domain")%></h2>
		<div class="Body">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td class="title">��ͨ���ڣ�</td>
					<td><span><%=Rs("d_statime")%></span></td>
					<td class="title">����������</td>
					<td><span><%=Rs("d_max")%></span>��</td>
				</tr>
				<tr>
					<td class="title">�������ڣ�</td>
					<td><span><%=Rs("d_expries")%></span></td>
					<td class="title">�Ѿ�ʹ�ã�</td>
					<td><span><%=Rs("d_count")%></span>��</td>
				</tr>
				<tr>
					<td class="title">�޸����룺</td>
					<td colspan="3"><form action="?act=ChgPwd" method="Post"> <input name="pwd" class="input" style="width:100px;"/> <input type="submit" value="�޸�" /></form></td>
				</tr>
			</table>
		</div>
		<h2><%=Domain%> URLת����¼</h2>
		<div class="Body">
			<table width="100%">
				<tr class="title">
					<td>������</td>
					<td>ת����ַ</td>
					<td>��ʾ����</td>
					<td>ת������</td>
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
							<option value="1">����URL</option>
							<option value="0" <%if Rs("u_type")=0 then response.write "selected=""selected""" end if%>>������</option>
						</select>
					</td>
					<td>
						<input type="submit" value="�޸�"/>
						<input type="button" value="ɾ��" onclick="if(confirm('�����Ҫɾ��������Ϊ\'<%=getRes(Rs("u_res"))%>\'��¼��'))location.href='?act=Del&oldres=<%=Rs("u_res")%>'"/>
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
							<option value="1">����URL</option>
							<option value="0">������</option>
						</select>
					</td>
					<td>
						<input type="submit" value="���"/>
					</td>
				</tr>
				</form>
<%
	end if
%>
			</table>
		</div>
		<h2>URLת��˵��</h2>
		<div class="Body">
			<table width="100%" class="info">
				<tr>
					<td>��һ��������Ҫ��ת��������(<%=Domain%>)������ <span style="color:#C00"><%=hostName%></span>������www.<%=Domain%>������Ҫ���www.<%=Domain%>��cname��¼��ֵΪ<span style="color:#C00"><%=hostName%></span></td>
				</tr>
				<tr>
					<td>�ڶ�������¼��ϵͳ�������Ӧ�ļ�¼������www.<%=Domain%>ת����http://www.baidu.com���������www��ת����ַ�http://www.baidu.com����ʾ���⣺�ٶ�����ת�����ͣ�����</td>
				</tr>
				<tr>
					<td style="color:#666">��ע���������������һ�㲻֧��cname��������A��¼������<%=Domain%>ת����http://www.baidu.com����Ҫ��<%=Domain%>��A��¼��<%=hostName%>��IP</td>
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
	WebErr "ɾ���ɹ���",-1,"?act=Main"
End Sub

Sub ChgPwd
	Domain=Session("LoginDomain")
	pass=Req("pwd")
	if pass="" then WebErr "���벻��Ϊ�ա�",-1,"?act=Main"
	Conn.Execute("update [domain] set [d_pass]='"&pass&"' where [d_dme]='"&Domain&"'")
	WebErr "�����޸ĳɹ���",-1,"?act=Main"
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
	if len(u_Res)=0 then WebErr "����������Ϊ�ա�",1,-1
	oldRes=Req("oldres")
	if u_Res="@" then
		u_Res= Domain
	else
		u_Res= u_Res & "." & Domain
	end if
	
	if Conn.Execute("select count([u_id]) from [url] where [u_res]='"&u_Res&"' and [u_res]<>'"&oldRes&"'")(0)>0 then WebErr "URL('"&u_Res&"')�ظ��������ԡ�",1,-1
	Set Rs=Server.CreateObject("ADODB.Recordset")
	if oldRes="" then
		if d_Max<d_Count then WebErr "����URLת����¼�Ѿ���������¼���빺��URLת��������",1,-1
		Rs.Open "select * from [url]",Conn,3,2
		Rs.AddNew
		Rs("u_did")=d_id
	else
		Rs.Open "select * from [url] where [u_res]='"&oldRes&"' and u_did="&d_id,Conn,3,2
		if Rs.Eof then WebErr "��û��Ȩ�޲�����������",1,-1
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
	WebErr "�����ɹ���",-1,"?act=Main"
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
		if Rs.eof then WebErr "��ҵ�񲻴���",1,-1
		if Rs(0)<>Req("password") then WebErr "�����������",1,-1
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