<!--#Include File="function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 
%>


<%
Dim strBody,rs,sql
strBody=myDangerEncode(request.form("txtBody"))
set rs=Server.CreateObject("ADODB.recordset")
	sql="Select * From ���԰��¼��"
	rs.open sql,conn,1,2
	rs.addnew
	rs("������")=session("userID")
	rs("��������")=strBody
	rs("����ʱ��")=now()
	rs.update
rs.close
Set conn=Nothing
Response.Redirect "index.asp"									
%>
