dim URL, TITLE, INTERVAL
'URL = "http://stocks.finance.yahoo.co.jp/stocks/detail/?code=9501.T"
'TITLE = "東京電力(株)【9501】：株式/株価 - Yahoo!ファイナンス"
URL = "http://stocks.finance.yahoo.co.jp/stocks/detail/?code=6701.T"
TITLE = "ＮＥＣ【6701】：株式/株価 - Yahoo!ファイナンス"
INTERVAL = 1000

'===============================================================================
' 監視 & 表示
'===============================================================================
dim ie,  cur, pre, elm
Set ie = AttachIE()

cur = 0
pre = 0

While True
	IfMarketOpen
	waitIE ie

	do
		if ie.Document.getElementsByTagName("TD").length = 0 then
			WScript.StdErr.WriteLine vbLf & Date & vbTab & Time & vbTab & _
				"[D] getElementsByTagName().length = 0" & vbCrLF
		else
			' Just Debug
			WScript.StdErr.WriteLine vbLf & Date & vbTab & Time & vbTab & _
				"[D] getElementsByTagName(TD).length = [" & _
				ie.Document.getElementsByTagName("TD").length & "]"
			exit do
		end if
	loop

	for each elm in ie.Document.getElementsByTagName("TD")

		On Error Resume Next
		do while IsNull(elm.outerText)
			WScript.Sleep(100)
		loop
		On Error Goto 0

		WScript.StdErr.WriteLine Date & vbTab & Time & vbTab & "[D]     [" & elm.outerText & "]"
		if elm.getattribute("className") = "stoksPrice" then
			cur = elm.outerText
			exit for
		end if
	next

	WScript.Echo Now & "," & cur
	if pre <> cur then
		WScript.StdErr.Write Date & vbTab & Time & vbTab & cur & vbCrLf
	else 
		WScript.StdErr.Write Date & vbTab & Time & vbTab & cur & vbCr
	end if
	pre = cur
	WScript.Sleep INTERVAL
Wend
WScript.Quit

'===============================================================================
' IE Window があれば attach なければ open
'===============================================================================
Function AttachIE ()
	dim flag, sa, objIE
	flag = 0
	set sa = CreateObject("Shell.Application")
	For Each objIE In sa.Windows
		if TypeName(objIE.document) = "HTMLDocument" then
			if objIE.document.Title = TITLE then
				flag = 1
				exit for
			end if
		end if
	Next
	if flag = 0 then
		set objIE = CreateObject("InternetExplorer.Application")
		objIE.Visible = True
		objIE.Navigate2 URL
		'objIE.Quit
		set objIE = nothing
		set objIE = sa.Windows.Item(sa.Windows.Count - 1)
	end if
	set AttachIE = objIE
End Function

'===============================================================================
' Market Close なら寝る
'===============================================================================
Sub IfMarketOpen()
	dim cur, flag
	flag = 0
	do
		cur = TimeValue(Now)
		if ( cur > #09:00:00# and cur < #11:30:00# ) or ( cur > #12:30:00# and cur < #15:00:00# ) then
			if flag = 1 Then
				flag = 0
				WScript.StdErr.WriteLine vbLf & Date & vbTab & Time & vbTab & "Wake up!!"
			end if
			exit do
		else
			if flag = 0 then
				flag = 1
			end if
			WScript.StdErr.Write Date & vbTab & Time & vbTab & "Sleeping!!" & vbCr
			WScript.Sleep INTERVAL
		end if
	loop
End Sub

'===============================================================================
' IE がビジー状態の間待つ
'===============================================================================
Sub waitIE(objIE)
	While objIE.Busy
		WScript.Sleep 100
	Wend
End Sub

'===============================================================================
' TODO
'===============================================================================
'以下のような出力が来たら、ブラウザをリロードする
' 2014/11/06 9:16:18      Internet Explorer ではこのページは表示できません
' ---------
' 2014.11.21
' Z:\Project\finance\l.vbs(22, 3) Microsoft VBScript 実行時エラー: 書き込みできません。: 'getattribute'
' ---------
' 2014.11.28
' Z:\Project\finance>cscript /nologo l.vbs >> 20141128.tsv
' 2014/11/28      9:21:07 369
' Z:\Project\finance\l.vbs(21, 2) Microsoft VBScript 実行時エラー: 書き込みできません。
' ---------
' 2014.12.01
' Z:\Project\finance\l.vbs(29, 4) Microsoft VBScript 実行時エラー: 書き込みできません。: 'outerText'
' ---------
' 2014.12.03
' Z:\Project\finance\m.vbs(31, 4) Microsoft VBScript 実行時エラー: 書き込みできません。: 'getattribute'
' ----------
' 2014.12.12
' Z:\Project\finance\n.vbs(35, 3) Microsoft VBScript 実行時エラー: 書き込みできません。: 'outerText'

