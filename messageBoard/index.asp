<!--#Include File="function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 
%>
<html>
<head> 
	<title>�Źܰ༶���԰�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="../class.css">
	<script language="JavaScript">
		//�ú����������пͻ�����֤
		function check_Null(){
			
			if (document.form1.txtBody.value==""){
				alert("�ף�������û��ûɫŶ��");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<h1 align="center">13�Źܣ���ɫ����...</h1>


	<form name="form1" method="POST" action="insert.asp" onSubmit="JavaScript:return check_Null();">
		<table border="0" width="80%" bgcolor="white" align="center">
		  
			<tr>
				<td><font>�����ɫ��</font></td>
				<td><textarea name="txtBody" cols="80" rows="4" id="txtBody"></textarea></td>
			</tr><tr>
				<td></td>
				<td><input type="submit" name="Submit" value="�� ��"></td>
			</tr>
		</table>
	</form>
	<%
	Dim rs,strSql,sql,rs1
	strSql ="Select * From ���԰��¼�� Order By ����ʱ�� Desc"	
    Set rs=conn.Execute(strSql)
	Do While Not rs.Eof
	  set rs1=Server.CreateObject("ADODB.recordset")
	  sql="select ���� from �Źܰ��Ա�� where ѧ��='"&rs("������")&"'"
      rs1.open sql,conn,1,2
	  %>
		<table border="0" width="80%" align="center">
			<tr><td colspan="2"><hr></td></tr>
			<tr><td>����</td><td><%=myHTMLEncode(rs("��������"))%></td></tr>
			<tr><td>������</td><td><%=rs1("����")%></td></tr>
			<tr><td>ʱ��</td><td><%=rs("����ʱ��")%></td></tr>
			<tr><td></td><td><a href="delete.asp?txtID=<%=rs("ID")%>">ɾ��</a></td></tr>
		</table>
		<%
		rs1.close
		rs.MoveNext
	Loop 
	
	rs.Close
	Set rs=Nothing
	conn.Close
	Set conn=Nothing
	%>
</body>
</html>