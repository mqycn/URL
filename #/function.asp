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
'=  �ļ���#/function.asp                                     =
'=  ���ܣ�����������������һЩ���õĺ�����                   =
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