<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 
%>

		
<html>
<head>
	<title>ɾ������</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
	
	<%
	
	dim txtID,rs,sql
	txtID=Request.querystring("txtID")
	set rs=Server.CreateObject("ADODB.recordset")
	sql="Select ������ From ���԰��¼�� where ID=" &txtID
	rs.open sql,conn,1,2
	
	If rs("������")=session("userID") Then
		Dim strSql
		strSql="Delete From ���԰��¼�� where ID=" &txtID 
		conn.Execute(strSql)
		Response.Redirect("index.asp")
	else response.write"�Բ����㲻��ɾ�����˵�����!"
	End If
	%>
