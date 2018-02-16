<%@ CODEPAGE=65001 LANGUAGE="VBScript"%>
<!DOCTYPE HTML>
<html>
<head>
	<title>InstagramGet(201801仕様)</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<%
	' 観に行きたい人のInstagramのID 初期値
	Dim instagramId : instagramId = ""
	' 取得する投稿の数 初期値 最新50枚取得したい時はmaximage=50にする 数が多いと処理に時間がかかります
	Dim maximage : maximage = 10

	If Not IsEmpty(Request.QueryString("g_instagramId")) Then instagramId = Request.QueryString("g_instagramId")
	If Not IsEmpty(Request.QueryString("g_maximage")) Then maximage = Request.QueryString("g_maximage")
 %>
<form action="Default.asp" method="GET">
	<div><div>id</div><input type="text" name="g_instagramId" value="<%=instagramId%>"></div>
	<div><div>取得する投稿数</div><input type="text" name="g_maximage" maxlength="4" value="<%=maximage%>"></div>
	<input type="submit" value="取得">
</form>
<%
' *************************** ***************************
' Instagramスクレイピング 仕様 201801
' *************************** ***************************

	If instagramId="" Then Response.End
	If IsNumeric(maximage) Then 
		maximage=Cdbl(maximage)
	Else
		maximage=1
	End If

	Session.CodePage = 65001
	response.charset = "UTF-8"
	Server.ScriptTimeout = 3000

	' 正規表現1 URLcode取得用
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = """code"":.*?"",""date"""
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索

	' 正規表現2 画像用
	Set regEx2 = CreateObject("VBScript.RegExp")
	regEx2.Pattern = "https://scontent-nrt1-1\.cdninstagram\.com.*?\.jpg"
	regEx2.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx2.Global = True ' 文字列全体を検索

	' 正規表現3 動画用
	Set regEx3 = CreateObject("VBScript.RegExp")
	regEx3.Pattern = "video_url"":""https://scontent-nrt1-1\.cdninstagram\.com/vp.*\.mp4"
	regEx3.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx3.Global = True ' 文字列全体を検索

	' 正規表現4 ページ送り
	Set regEx4 = CreateObject("VBScript.RegExp")
	regEx4.Pattern = """id"":""\d{19}"","
	regEx4.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx4.Global = True ' 文字列全体を検索

	Dim tempDic : Set tempDic = Server.CreateObject("Scripting.Dictionary")
	Dim tempDicjpg : Set tempDicjpg = Server.CreateObject("Scripting.Dictionary")
	Dim tempDicmp4 : Set tempDicmp4 = Server.CreateObject("Scripting.Dictionary")

	baseurl = "https://www.instagram.com/"&instagramId&"/?max_id="

	response.Write "<div><a href='"&baseurl&"' target='_blank'>Instagramへ</a></div>"

	Dim wurl,id
	Dim resText
	Dim i,j,k
	' 最新１件だけ先に取得
	Call addDicHtmlSimple(baseurl)
	If maximage > 1 Then
		resText = getXMLHTTP(baseurl)
		k=1
		For i=0 To maximage \ 12
			j = 0
			Set matches4 = regEx4.Execute(resText)
			If matches4.Count > 0 Then 
				For j=0 To matches4.Count - 1
					id = Mid(matches4(j),7,19)
					wurl = baseurl & id
					Call addDicHtmlSimple(wurl)
					' 取得した１２番目を次のURLに指定
					If j =11 ANd Not IsNull(matches4(11)) Then
						resText = getXMLHTTP(wurl)
					End If
					k=k+1
					If k >= maximage Then Exit For
				Next
			End If
			If k >= maximage Then Exit For
		Next
	End If

	response.Write Join(tempDic.Items(),"")
	response.Flush

	Set regEx = Nothing
	Set regEx2 = Nothing
	Set regEx4 = Nothing
	Set regEx4 = Nothing
	Set matches = Nothing
	Set matches2 = Nothing
	Set matches3 = Nothing
	Set matches4 = Nothing
	Set tempDic = Nothing



	' //=================================================================
	' // 関数: addDicHtmlSimple(byval url)
	' // 概要: サイトへアクセスして結果を辞書に入れる
	' //=================================================================
	Private Sub addDicHtmlSimple(byval url)
		Dim resText1,resText2
		Dim codeId
		Dim i,j
		Dim strjpg,strmp4
		resText1 = getXMLHTTP(url)
		Set matches = regEx.Execute(resText1)

		If matches.Count <> 0 Then
			codeId = Mid(matches(0),9,11)
			resText2 = getXMLHTTP("https://www.instagram.com/p/"&codeId)
			Set matches2 = regEx2.Execute(resText2)
			If matches2.Count <> 0 Then
				For i = 0 To matches2.Count - 1
					strjpg = matches2(i).Value
					If Not Len(strjpg)>170 And Instr(strjpg,"/e35/") > 0 And Instr(strjpg,"640x640") = 0 And Instr(strjpg,"750x750") = 0 Then
						If Not tempDicjpg.Exists(strjpg) Then
							Call tempDicjpg.Add(strjpg,"")
							Call tempDic.Add(tempDic.Count,"<img src="""&strjpg&""">")
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
						Call tempDic.Add(tempDic.Count,"<video autoplay loop muted controls><source src="""&strmp4&""" type=""video/mp4"" /></video>")
					End If
				Next
			End If
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