<!--#Include File="function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../ÐÅ¹Ü°à¼ÍÄî²á.mdb")
conn.Open strConn 
%>


<%
Dim strBody,rs,sql
strBody=myDangerEncode(request.form("txtBody"))
set rs=Server.CreateObject("ADODB.recordset")
	sql="Select * From ÁôÑÔ°å¼ÇÂ¼±í"
	rs.open sql,conn,1,2
	rs.addnew
	rs("ÁôÑÔÈË")=session("userID")
	rs("ÁôÑÔÄÚÈÝ")=strBody
	rs("ÁôÑÔÊ±¼ä")=now()
	rs.update
rs.close
Set conn=Nothing
Response.Redirect "index.asp"									
%>
