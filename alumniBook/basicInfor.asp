<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>������Ϣ</title>
</head>
<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs,sql,ID
ID=request.querystring("ID")
set rs=server.CreateObject("adodb.recordset")
sql="select * from �Źܰ��Ա�� where ѧ��='"&ID&"'"
rs.open sql,conn,1,2
response.write"<table width='290px' height='200px'><tr><td>������</td><td>"&rs("����")&"</td></tr><tr><td>�Ա�</td><td>"&rs("�Ա�")&"</td></tr><tr><td>���գ�</td><td>"&rs("����")&"</td></tr><tr><td>QQ�ţ�</td><td>"&rs("QQ��")&"</td></tr><tr><td>�ֻ����룺</td><td>"&rs("�ֻ�����")&"</td></tr><tr><td>��ϲ������ɫ��</td><td>"&rs("��ϲ������ɫ")&"</td></tr><tr><td>������������</td><td>"&rs("����������")&"</td></tr></table>"
%>
</body>
</html>

