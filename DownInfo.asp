<!-- #include file="DownInfo.inc" -->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>�s�W����1</title>
</head>

<body bgcolor="#FFFFFF">


<blockquote>
  <blockquote>
    <blockquote>
      <blockquote>


<p style="margin-top:10; margin-bottom:10; line-height:150%" align="left">
<font size="4" color="#800000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��Ƥw�e�X �A �� �� �N �� �� 
�P �z �p ��! �νШӹq (06)208-8808  (04)2260-0001</font></p>


<%
   dim buf 
	On error resume next

	Dim mySmartMail
	Set mySmartMail = Server.CreateObject("aspSmartMail.SmartMail")

'	Mail Server
'	***********
	mySmartMail.Server = "dyit.com.tw"

'	From
'	****
	mySmartMail.SenderName = request("company") & "-" &  request("name") & " TEL:" & request("phone")
	mySmartMail.SenderAddress = "dyit01@dyit.com.tw"

'	To
'	**
	mySmartMail.Recipients.Add "dyit01@dyit.com.tw", "�F�ɬ�ަ������q"

'	Message
'	*******
	mySmartMail.Subject = "�M�D�ե�" & request("product") & request("address")
'	buf =  "���q:" & Company & "�m�W:" & Name  & "�a�}:" & Address & "�q��:" & phone & "��~:" & Item & "�����n��:" & product
 	mySmartMail.Body = "132sss" 
 	
'	Send the message
'	****************
	mySmartMail.SendMail

	if Err.Number<>0 then

		Response.write "Error: " & Err.description

	else

		
	end if

%>


</p>


      </blockquote>
    </blockquote>
  </blockquote>
</blockquote>


</body>
</body>

</html>