<html>

<head>
<link rel="stylesheet" type="text/css" href="../styles/slideshow.css">
</head>

<body>
<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open (Server.MapPath("alumni_registration.mdb"))
Dim t1,t2,t3,t4,t5,t6,t7,t8,t9,t10
t1=Request.Form("t1")
t2=Request.Form("t2")
t3=Request.Form("t3")
t4=Request.Form("t4")
t5=Request.Form("t5")
t6=Request.Form("t6")
t7=Request.Form("t7")
t8=Request.Form("t8")
t9=Request.Form("t9")
t10=Request.Form("t10")
IF t1 <> "" AND t2 <> "" AND t3 <> "" AND t4 <> "" AND t5 <> "" AND t6 <> "" AND t7 <> "" AND t7 <> "" AND t8 <> "" AND t9 <> "" AND t10 <> "" THEN
sql="INSERT INTO Results (nam,fnam,dob,yar,cno,ca,pa,stg,wek,cp)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("t1") & "',"
sql=sql & "'" & Request.Form("t2") & "',"
sql=sql & "'" & Request.Form("t3") & "',"
sql=sql & "'" & Request.Form("t4") & "',"
sql=sql & "'" & Request.Form("t5") & "',"
sql=sql & "'" & Request.Form("t6") & "',"
sql=sql & "'" & Request.Form("t7") & "',"
sql=sql & "'" & Request.Form("t8") & "',"
sql=sql & "'" & Request.Form("t9") & "',"
sql=sql & "'" & Request.Form("t10") & "')"

on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("Due to technical reason, your data can not be added!")
else
  Response.Redirect "alumni_submitted.htm"
end if
conn.close
Session.Abandon
ELSE
	Response.Write "<p>Please click back on your browser and complete the following fields:</p>"
	IF t1 <> "" THEN
	ELSE
	Response.Write "<b>. Your name</b><br>"
	END IF
	IF t2 <> "" THEN
	ELSE
	Response.Write "<b>. Father's name</b><br>"
	END IF
	IF t3 <> "" THEN
	ELSE
	Response.Write "<b>. Date of birth</b><br>"
	END IF
	IF t4 <> "" THEN
	ELSE
	Response.Write "<b>. Year of passingl</b><br>"
	END IF
	IF t5 <> "" THEN
	ELSE
	Response.Write "<b>. Contact number</b><br>"
	END IF
	IF t6 <> "" THEN
	ELSE
	Response.Write "<b>. Current address</b><br>"
	END IF
	IF t7 <> "" THEN
	ELSE
	Response.Write "<b>. Permanent address</b><br>"
	END IF
	IF t8 <> "" THEN
	ELSE
	Response.Write "<b>. Strength</b><br>"
	END IF
	IF t9 <> "" THEN
	ELSE
	Response.Write "<b>. Weaknesses</b><br>"
	END IF
	IF t10 <> "" THEN
	ELSE
	Response.Write "<b>. Career plans</b><br>"
	END IF
END IF
%>
</body>
</html>
