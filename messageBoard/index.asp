<!--#Include File="function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 
%>
<html>
<head> 
	<title>信管班级留言板</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<link rel="stylesheet" href="../class.css">
	<script language="JavaScript">
		//该函数用来进行客户端验证
		function check_Null(){
			
			if (document.form1.txtBody.value==""){
				alert("亲，不可以没声没色哦！");
				return false;
			}
			return true;
		}
	</script>
</head>
<body >
	<h1 align="center">13信管，声色继续...</h1>


	<form name="form1" method="POST" action="insert.asp" onSubmit="JavaScript:return check_Null();">
		<table border="0" width="80%" bgcolor="white" align="center">
		  
			<tr>
				<td><font>你的声色：</font></td>
				<td><textarea name="txtBody" cols="80" rows="4" id="txtBody"></textarea></td>
			</tr><tr>
				<td></td>
				<td><input type="submit" name="Submit" value="提 交"></td>
			</tr>
		</table>
	</form>
	<%
	Dim rs,strSql,sql,rs1
	strSql ="Select * From 留言板记录表 Order By 留言时间 Desc"	
    Set rs=conn.Execute(strSql)
	Do While Not rs.Eof
	  set rs1=Server.CreateObject("ADODB.recordset")
	  sql="select 姓名 from 信管班成员表 where 学号='"&rs("留言人")&"'"
      rs1.open sql,conn,1,2
	  %>
		<table border="0" width="80%" align="center">
			<tr><td colspan="2"><hr></td></tr>
			<tr><td>内容</td><td><%=myHTMLEncode(rs("留言内容"))%></td></tr>
			<tr><td>留言人</td><td><%=rs1("姓名")%></td></tr>
			<tr><td>时间</td><td><%=rs("留言时间")%></td></tr>
			<tr><td></td><td><a href="delete.asp?txtID=<%=rs("ID")%>">删除</a></td></tr>
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