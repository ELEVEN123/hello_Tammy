<!--#Include File="../messageBoard/function.asp"-->

<html>
<head>
<title>�ռǲ���</title>
</head>
</html>
<%
'�������ݿ�
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs,sql,signature
signature=myDangerEncode(request.form("signature"))
set rs=server.createobject("adodb.recordset")
sql="select * from �ռǱ�"
rs.open sql,conn,1,2
rs.addnew
rs("�ռ���")=session("userID")
rs("�ռ�ʱ��")=date()
rs("�ռ�����")=signature
rs.update
rs.close
set rs=nothing
response.redirect"../homepageFrameset/personalPage.asp"

%>