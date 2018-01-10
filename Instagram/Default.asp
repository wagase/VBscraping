<%@ CODEPAGE=65001 %>
<!DOCTYPE HTML>
<html>
<head>
	<title>Instagram</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<%
' *************************** ***************************
' Instagramスクレイピング 仕様 201801
' *************************** ***************************

	Session.CodePage = 65001
	response.charset = "UTF-8"

	' 観に行きたい人のInstagramのID
	Dim instagramId : instagramId = "****"

	' 正規表現1 URLcode取得用
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = """code"":.*?"", ""date"""
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索

	' 正規表現2 画像用
	Set regEx2 = CreateObject("VBScript.RegExp")
	regEx2.Pattern = "https://scontent-nrt1-1\.cdninstagram\.com.*?\.jpg"
	regEx2.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx2.Global = True ' 文字列全体を検索

	' 正規表現3 動画用
	Set regEx3 = CreateObject("VBScript.RegExp")
	regEx3.Pattern = "https://scontent-nrt1-1\.cdninstagram\.com/vp.*?\.mp4"
	regEx3.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx3.Global = True ' 文字列全体を検索

	Dim tempDic : Set tempDic = Server.CreateObject("Scripting.Dictionary")
	Dim baseurl
	Dim resText1,resText2
	Dim codeId

	baseurl = "https://www.instagram.com/"&instagramId&"/"
	resText1 = getXMLHTTP(baseurl)
	Set matches = regEx.Execute(resText1)
	If matches.Count <> 0 Then
		for each code in matches
			codeId = Mid(code,10,11)
			resText2 = getXMLHTTP("https://www.instagram.com/p/"&codeId)
			Set matches2 = regEx2.Execute(resText2)
			If matches2.Count <> 0 Then
				Call tempDic.Add(tempDic.Count,"<img src="""&matches2(0)&""">")
			End If
			Set matches3 = regEx3.Execute(resText2)
			If matches3.Count <> 0 Then
				Call tempDic.Add(tempDic.Count,"<video autoplay loop muted controls><source src="""&matches3(0)&""" type=""video/mp4"" /></video>")
			End If
		NEXT
	End If
	Set regEx = Nothing

	response.write Join(tempDic.Items(),"")


	' //=================================================================
	' // 関数: getXMLHTTP(byval url)
	' // 概要: サイトへアクセスして結果を返す
	' //=================================================================
	Private Function getXMLHTTP(byval url)

		Dim xmlhttp,resText
		Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
		xmlhttp.Open "GET",url, False
		xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlhttp.Send ""
		resText = xmlhttp.responseText
		Set xmlhttp = Nothing

		getXMLHTTP = resText
	End Function
%>
</body>
</html>