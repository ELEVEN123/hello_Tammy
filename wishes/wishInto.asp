<!--#Include File="../messageBoard/function.asp"-->
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql,wishTo,wishes
wishTo=request.form("wishTo")
wishes=myDangerEncode(request.form("wishes"))
set rs=server.createobject("adodb.recordset")
sql="select * from 祝福表"
rs.open sql,conn,1,2
rs.addnew
rs("祝福人")=session("userID")
rs("被祝福人")=wishTo
rs("祝福时间")=now()
rs("祝福内容")=wishes
rs.update
rs.close
set rs=nothing
response.redirect"../homepageFrameset/personalPage.asp"
%>
