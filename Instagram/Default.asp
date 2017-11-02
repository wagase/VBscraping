<%@ CODEPAGE=65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<%
' ******************************************************
' Instagramを覗きに行く
' ******************************************************
	On Error Resume Next
	Session.CodePage = 65001
	response.charset = "UTF-8"
	Dim URL,ID
	ID ="InstagramID"
	URL = "https://www.instagram.com/" & ID & "/"
	' 正規表現で対象しぼる
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = "https://scontent-nrt1-1\.cdninstagram\.com"
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索
	Dim tempDic
	Set tempDic = Server.CreateObject("Scripting.Dictionary")
	Dim array1
	Dim xmlhttp
	Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET",URL, False
	xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.Send ""
	resText = xmlhttp.responseText
	Set xmlhttp = Nothing
	array1 = split(resText,vblf)
	for each key in array1
		Set matches = regEx.Execute(key)
		If matches.Count <> 0 Then
			Do Until InStr(key,".jpg") <= 0
				num1 = InStr(key,"https://scontent-nrt1-1")
				num2 = InStr(key,".jpg") + 4
				str2 = Mid(key,num1,num2-num1)
				Call tempDic.Add(tempDic.Count,"<img src="""&str2&"""")
				key = Right(key,Len(key)-num2)
			Loop
		End If
	Next
	Set regEx = Nothing
	response.write Join(tempDic.Items(),">")
	If Err.Number <> 0 Then
		Response.write Err.Number & Err.Description & Err.Erl
		Err.Clear()
	End If
	On Error GoTo 0
%>
</body>
</html>