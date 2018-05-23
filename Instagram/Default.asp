<%@ CODEPAGE=65001 LANGUAGE="VBScript"%>
<!DOCTYPE HTML>
<html>
<head>
	<title>InstagramGet(201803仕様)</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<%
	' 観に行きたい人のInstagramのID 初期値
	Dim instagramId : instagramId = ""

	If Not IsEmpty(Request.QueryString("g_instagramId")) Then instagramId = Request.QueryString("g_instagramId")
	If Not IsEmpty(Request.QueryString("g_maximage")) Then maximage = Request.QueryString("g_maximage")
 %>
<form action="Default.asp" method="GET">
	<div><div>id</div><input type="text" name="g_instagramId" value="<%=instagramId%>"></div>
	<input type="submit" value="取得">
</form>
<%
' *************************** ***************************
' Instagramスクレイピング 仕様 20180314
' *************************** ***************************

	If instagramId="" Then Response.End

	Session.CodePage = 65001
	response.charset = "UTF-8"
	Server.ScriptTimeout = 3000

	' 正規表現1 URLcode取得用
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = """shortcode"":"".*?"",""edge_"
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索

	' 正規表現2 画像用
	Set regEx2 = CreateObject("VBScript.RegExp")
	regEx2.Pattern = "https://scontent-nrt1-1\.cdninstagram\.com.*?\.jpg"
	regEx2.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx2.Global = True ' 文字列全体を検索

	' 正規表現3 動画用
	Set regEx3 = CreateObject("VBScript.RegExp")
	regEx3.Pattern = "video_url"":""https://scontent-nrt1-1\.cdninstagram\.com/vp.*?\.mp4"
	regEx3.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx3.Global = True ' 文字列全体を検索

	Dim tempDicjpg : Set tempDicjpg = Server.CreateObject("Scripting.Dictionary")
	Dim tempDicmp4 : Set tempDicmp4 = Server.CreateObject("Scripting.Dictionary")

	baseurl = "https://www.instagram.com/"&instagramId&"/"

	response.Write "<div><a href='"&baseurl&"' target='_blank'>Instagramへ</a></div>"
	response.Flush

	Dim test
	Dim resText1
	Dim i

	' なぜか最初の数回は失敗するのでキャッシング用？に2回アクセス
	For i=0 To 2
		test = getXMLHTTP(baseurl)
	Next
	resText1 = getXMLHTTP(baseurl)
	Set matches = regEx.Execute(resText1)
	If matches.Count <> 0 Then
		For i= 0 To matches.Count - 1
			codeId = Mid(matches(i),14,11)
			Call addDicHtmlSimple(codeId)
		Next
	End If

	response.Flush

	Set regEx = Nothing
	Set regEx2 = Nothing
	Set regEx3 = Nothing

	Set matches = Nothing
	Set matches2 = Nothing
	Set matches3 = Nothing

	Set tempDicjpg = Nothing
	Set tempDicmp4 = Nothing



	' //=================================================================
	' // 関数: addDicHtmlSimple(byval codeId)
	' // 概要: サイトへアクセスして結果を辞書に入れる
	' //=================================================================
	Private Sub addDicHtmlSimple(byval codeId)
		Dim resText2
		Dim i,j
		Dim strjpg,strmp4
		resText2 = getXMLHTTP("https://www.instagram.com/p/"&codeId)
		Set matches2 = regEx2.Execute(resText2)
		If matches2.Count <> 0 Then
			For i = 0 To matches2.Count - 1
				strjpg = matches2(i).Value
				If Not Len(strjpg)>170 And Instr(strjpg,"/e35/") > 0 And Instr(strjpg,"640x640") = 0 And Instr(strjpg,"750x750") = 0 Then
					If Not tempDicjpg.Exists(strjpg) Then
						Call tempDicjpg.Add(strjpg,"")
						response.Write "<img src="""&strjpg&""">"
						response.Flush
					End If
				End If
			Next
		End If
		Set matches3 = regEx3.Execute(resText2)
		If matches3.Count <> 0 Then
			For i = 0 To matches3.Count - 1
				strmp4 = Right(matches3(i).Value,Len(matches3(i).Value)-12)
				If Not tempDicmp4.Exists(strmp4) Then
					Call tempDicmp4.Add(strmp4,"")
					response.Write "<video autoplay loop muted controls><source src="""&strmp4&""" type=""video/mp4"" /></video>"
    				response.Flush
				End If
			Next
		End If
	End Sub

	' //=================================================================
	' // 関数: getXMLHTTP(byval url)
	' // 概要: サイトへアクセスして結果を返す
	' //=================================================================
	Private Function getXMLHTTP(byval url)

		Dim xmlhttp,resText
		Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
		xmlhttp.Open "GET",url, False
		xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		On Error Resume Next
		xmlhttp.Send ""
		On Error Goto 0
		resText = xmlhttp.responseText
		Set xmlhttp = Nothing

		getXMLHTTP = resText
	End Function
%>
</body>
</html>