<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Class amazon_s3_plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME
	
	Private strAccessKeyID, strSecretAccessKey, awsStatus, awsStoreFile, s3_strBinaryData
	Private s3_strLocalFile, s3_strLocalFileRaw, s3_strRemoteFile, s3_strBucket, s3_strOutFileName
	Private STORABLE_FILES
	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"

		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' ' Check And Create Table
		' '------------------------------
		Dim PluginTableName
			PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME

    	If TableExist(PluginTableName) = False Then
    		Conn.Execute("SET NAMES utf8mb4;") 
    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
    		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

    		q=""
    		q=q+"CREATE TABLE `"& PluginTableName &"` ( "
    		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    		q=q+"  `LOCAL_NAME` varchar(255) DEFAULT NULL, "
    		q=q+"  `REMOTE_NAME` varchar(255) DEFAULT NULL, "
    		q=q+"  `UPLOAD_DATE` datetime DEFAULT NULL, "
    		q=q+"  `HASH` varchar(32) DEFAULT NULL, "
    		q=q+"  `ETAG` varchar(150) DEFAULT NULL, "
    		q=q+"  `SILINDI` int(1) DEFAULT 0, "
    		q=q+"  PRIMARY KEY (`ID`), "
    		q=q+"  UNIQUE KEY `IND2` (`ETAG`,`SILINDI`) USING BTREE, "
    		q=q+"  KEY `IND1` (`LOCAL_NAME`), "
    		q=q+"  KEY `IND3` (`REMOTE_NAME`,`SILINDI`) "
    		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
			Conn.Execute(q)

    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

			' Create Log
			'------------------------------
    		Call PanelLog(""& PLUGIN_CODE &" için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

			' Register Settings
			'------------------------------
			DebugTimer ""& PLUGIN_CODE &" class_register() End"
    	End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "amazon_s3_plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "11")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		a=GetSettings(""&PLUGIN_CODE&"_ACCESS_ID", "")
		a=GetSettings(""&PLUGIN_CODE&"_SECRET_KEY", "")
		a=GetSettings(""&PLUGIN_CODE&"_BUCKET", "")
		a=GetSettings(""&PLUGIN_CODE&"_LOCAL_TEMP_FOLDER", "/content/trash/")
		a=GetSettings(""&PLUGIN_CODE&"_STORE_FILES", "PRODUCT,BLOG,USERFILE")

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "AJAX:GetFile" Then
			AWS_FILE = Query.Data("FileName")

			Set Amazon = New amazon_s3_plugin
				Amazon.s3_RemoteFile                = AWS_FILE
				Amazon.s3_OutFileName               = AWS_FILE
				Amazon.s3_StreamToBrowser()
			Set Amazon = Nothing

			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:PluginLogs" Then
			Call PluginPage("Header")

			With Response 
				.Write "<div class=""table-responsive"">"
				.Write "	<table class=""table table-striped table-bordered"">"
				.Write "		<thead>"
				.Write "			<tr>"
				.Write "				<th>Obje</th>"
				.Write "				<th>Yükleme Tarihi</th>"
				.Write "				<th>ETag</th>"
				.Write "				<th></th>"
				.Write "			</tr>"
				.Write "		</thead>"
				.Write "		<tbody>"
				Set Siteler = Conn.Execute("SELECT * FROM tbl_plugin_aws_log WHERE SILINDI=0 ORDER BY ID DESC")
				If Siteler.Eof Then 
				    .Write "<tr>"
				    .Write "<td colspan=""4"" align=""center"">İşlem Geçmişi Bulunamadı</td>"
				    .Write "</tr>"
				End If
				Do While Not Siteler.Eof
				.Write "			<tr>"
				.Write "				<td><strong>"& Siteler("REMOTE_NAME") &"</strong><br /><small><code>"& Siteler("LOCAL_NAME") &"</code></small></td>"
				.Write "				<td>"& Siteler("UPLOAD_DATE") &"</td>"
				.Write "				<td>"& Siteler("ETAG") &"</td>"
				.Write "				<td align=""right""><a href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=AJAX:GetFile&FileName="& Siteler("REMOTE_NAME") &""" class=""btn btn-primary btn-sm"" download>Download</a></td>"
				.Write "			</tr>"
				Siteler.MoveNext : Loop
				Siteler.Close : Set Siteler = Nothing
				.Write "				"				
				.Write "				"				
				.Write "				"				
				.Write "		</tbody>"
				.Write "	</table>"
				.Write "</div>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_ACCESS_ID", "Access Key", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_SECRET_KEY", "Secret Key", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_BUCKET", "Bucket Name", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("file-explorer", ""& PLUGIN_CODE &"_LOCAL_TEMP_FOLDER", "Local Temp Folder", "/content/trash/", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write  		QuickSettings("multiselect", ""& PLUGIN_CODE &"_STORE_FILES", "Only Store This Files", STORABLE_FILES , TO_DB)
			.Write "    </div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""https://aws.amazon.com/tr/s3/"" class=""btn btn-info btn-sm"">"
			.Write "        	AWS S3 Hakkında"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"_&Page=SHOW:PluginLogs"" class=""btn btn-primary btn-sm"">"
			.Write "        	Kayıt Geçmişi"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	private sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "Amazon AWS S3 Bucket Plugin"
    	PLUGIN_CODE 			= "AWS_S3"
    	PLUGIN_DB_NAME 			= "aws_log"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_CREDITS 			= "Coded by @cavebring [https://github.com/cavebring/class-classic-ASP-aws-S3] Redevelopment @badursun"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Amazon-S3-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-amazon"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

    	STORABLE_FILES 			= Array("PRODUCT[Orj]", "PRODUCT[M]", "PRODUCT[T]", "PRODUCT[S]", "PRODUCT[Cms]","BLOG","USERFILE", "FILES")
    	awsStatus 				= Cint( GetSettings(""& PLUGIN_CODE &"_ACTIVE", "0") )
		strAccessKeyID 			= GetSettings(""& PLUGIN_CODE &"_ACCESS_ID", "")
		strSecretAccessKey 		= GetSettings(""& PLUGIN_CODE &"_SECRET_KEY", "")
		s3_strBucket 			= GetSettings(""& PLUGIN_CODE &"_BUCKET", "")
    	
    	class_register()
	end sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable )
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get strLocalTempDir()
		strLocalTempDir = Server.MapPath( GetSettings(""& PLUGIN_CODE &"_LOCAL_TEMP_FOLDER", "/content/trash/") ) & "\"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get s3_LocalFile()
		s3_LocalFile = s3_strLocalFile 
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get s3_RemoteFile()
		s3_RemoteFile = s3_strRemoteFile
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get s3_Bucket()
		s3_Bucket = s3_strBucket
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get s3_OutFileName()
		s3_OutFileName = s3_strOutFileName
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get s3_BinaryData()
		s3_BinaryData = s3_strBinaryData
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let s3_LocalFile(ByVal NewValue)
	   s3_strLocalFileRaw = NewValue
	   s3_strLocalFile = Server.MapPath( NewValue )
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let s3_RemoteFile(ByVal NewValue)
	   s3_strRemoteFile = NewValue
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let s3_Bucket(ByVal NewValue)
	   s3_strBucket = NewValue
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let s3_OutFileName(ByVal NewValue)
	   s3_strOutFileName = NewValue
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let s3_BinaryData(ByVal NewValue)
	   s3_strBinaryData = NewValue
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function NowInGMT()
		Dim sh: Set sh = Server.CreateObject("WScript.Shell")
		Dim iOffset: iOffset = sh.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
		Dim dtNowGMT: dtNowGMT = DateAdd("n", iOffset, Now())
		Dim strDay: strDay = "NA"
	
		Select Case Weekday(dtNowGMT)
			Case 1 strDay = "Sun"
			Case 2 strDay = "Mon"
			Case 3 strDay = "Tue"
			Case 4 strDay = "Wed"
			Case 5 strDay = "Thu"
			Case 6 strDay = "Fri"
			Case 7 strDay = "Sat"
			Case Else strDay = "Error"
		End Select

		Dim strMonth: strMonth = "NA"
		Select Case Month(dtNowGMT)
			Case 1 strMonth = "Jan"
			Case 2 strMonth = "Feb"
			Case 3 strMonth = "Mar"
			Case 4 strMonth = "Apr"
			Case 5 strMonth = "May"
			Case 6 strMonth = "Jun"
			Case 7 strMonth = "Jul"
			Case 8 strMonth = "Aug"
			Case 9 strMonth = "Sep"
			Case 10 strMonth = "Oct"
			Case 11 strMonth = "Nov"
			Case 12 strMonth = "Dec"
			Case Else strMonth = "Error"
		End Select

		Dim strHour: strHour = CStr(Hour(dtNowGMT))
		If Len(strHour) = 1 Then strHour = "0" & strHour
		
		Dim strMinute: strMinute = CStr(Minute(dtNowGMT))
		If Len(strMinute) = 1 Then strMinute = "0" & strMinute
		
		Dim strSecond: strSecond = CStr(Second(dtNowGMT))
		If Len(strSecond) = 1 Then strSecond = "0" & strSecond
		
		Dim strNowInGMT: strNowInGMT = _
		strDay & _
		", " & _
		Day(dtNowGMT) & _
		" " & _
		strMonth & _
		" " & _
		Year(dtNowGMT) & _
		" " & _
		strHour & _
		":" & _
		strMinute & _
		":" & _
		strSecond & _
		" +0000"
		NowInGMT = strNowInGMT
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function GetBytesFromString(strValue)
		Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
		stm.Open
		stm.Type = 2
		stm.Charset = "ascii"
		stm.WriteText strValue
		stm.Position = 0
		stm.Type = 1
		GetBytesFromString = stm.Read
		Set stm = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Function HMACSHA1(strKey, strValue)
		Dim sha1: Set sha1 = Server.CreateObject("System.Security.Cryptography.HMACSHA1")
		sha1.key = GetBytesFromString(strKey)
		HMACSHA1 = sha1.ComputeHash_2(GetBytesFromString(strValue))
		Set sha1 = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Function ConvertBytesToBase64(byteValue)
		Dim dom: Set dom = Server.CreateObject("MSXML2.DomDocument")
		Dim elm: Set elm = dom.CreateElement("b64")
		elm.dataType = "bin.base64"
		elm.nodeTypedValue = byteValue
		ConvertBytesToBase64 = elm.Text
		Set elm = Nothing
		Set dom = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Function GetBytesFromFile(strFileName)
		Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
		stm.Type = 1
		stm.Open
		stm.LoadFromFile strFileName
		stm.Position = 0
		GetBytesFromFile = stm.Read
		stm.Close
		Set stm = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Function GetBytesFromStream(strBinary)
		Dim stm: Set stm = Server.CreateObject("ADODB.Stream")
		stm.Type = 1
		stm.Open
		stm.Write strBinary
		stm.Position = 0
		GetBytesFromStream = stm.Read
		stm.Close
		Set stm = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_Delete()
		Dim strNowInGMT: strNowInGMT = NowInGMT()
		Dim strStringToSign: strStringToSign = _
		  "DELETE" & vbLf & _
		  "" & vbLf & _
		  "text/xml" & vbLf & _
		  strNowInGMT & vbLf & _
		  "/" & s3_strBucket + "/" & s3_strRemoteFile
		Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
		Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature

		' Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Dim xhttp: Set xhttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			xhttp.open "DELETE", "https://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
            xhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
			xhttp.setRequestHeader "Content-Type", "text/xml"
			xhttp.setRequestHeader "Date", strNowInGMT
			xhttp.setRequestHeader "Authorization", strAuthorization
			xhttp.send
			If xhttp.status = 204 Then '204 = delete ok'
				s3_Delete = "1"
				' ETAG = xhttp.getAllResponseHeaders()
				' Conn.Execute("UPDATE tbl_plugin_aws_log SET SILINDI=1 WHERE REMOTE_NAME='"& s3_strRemoteFile &"' AND SILINDI=0")
				Conn.Execute("DELETE FROM tbl_plugin_aws_log WHERE REMOTE_NAME='"& s3_strRemoteFile &"'")
			Else
				s3_Delete = "0:" & xhttp.responseText
			End If
		Set xhttp = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_UploadBinary()
		Dim FileGUID : FileGUID = Uguid()
		Dim strNowInGMT: strNowInGMT = NowInGMT()
		Dim strStringToSign: strStringToSign = _
		  "PUT" & vbLf & _
		  "" & vbLf & _
		  "text/xml" & vbLf & _
		  strNowInGMT & vbLf & _
		  "/" & s3_strBucket + "/" & s3_strRemoteFile
		Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
		Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature

		' Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Dim xhttp: Set xhttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			xhttp.open "PUT", "https://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
            xhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
			xhttp.setRequestHeader "Content-Type", "text/xml"
			xhttp.setRequestHeader "Date", strNowInGMT
			xhttp.setRequestHeader "Authorization", strAuthorization
			xhttp.send GetBytesFromStream(s3_strBinaryData)
			If xhttp.status = "200" Then
				ETAG = Replace( xhttp.getResponseHeader("ETag"), """", "")
				s3_UploadBinary = "1"
				Conn.Execute("REPLACE INTO tbl_plugin_aws_log(LOCAL_NAME, REMOTE_NAME, ETAG, UPLOAD_DATE) values('Binary-"& FileGUID &"', '"& s3_strRemoteFile &"', '"& ETAG &"', NOW())")
			Else
				s3_UploadBinary = "0:" & xhttp.responseText
			End If
		Set xhttp = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_Upload()
		Dim FileGUID : FileGUID = Uguid()
		Dim FileContentType : FileContentType = FileMimeType(s3_strLocalFile)
		Dim strNowInGMT: strNowInGMT = NowInGMT()
		Dim FileLocalPath : FileLocalPath = s3_strLocalFile

		Dim strStringToSign: strStringToSign = _
		  "PUT" & vbLf & _
		  "" & vbLf & _
		  ""& FileContentType &"" & vbLf & _
		  strNowInGMT & vbLf & _
		  "/" & s3_strBucket + "/" & s3_strRemoteFile
		Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
		Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature

		' Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Dim xhttp: Set xhttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			xhttp.open "PUT", "https://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
            xhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
			xhttp.setRequestHeader "Content-Type", ""& FileContentType &""
			xhttp.setRequestHeader "Date", strNowInGMT
			xhttp.setRequestHeader "Authorization", strAuthorization
			xhttp.send GetBytesFromFile(s3_strLocalFile)
			If xhttp.status = "200" Then
				ETAG = Replace( xhttp.getResponseHeader("ETag"), """", "")
				s3_Upload = "1"
				Conn.Execute("REPLACE INTO tbl_plugin_aws_log(LOCAL_NAME, REMOTE_NAME, ETAG, UPLOAD_DATE) values('"& s3_strLocalFileRaw &"', '"& s3_strRemoteFile &"', '"& ETAG &"', NOW() )")
			Else
				s3_Upload = "0:" & xhttp.responseText
			End If
		Set xhttp = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_Download()
		Dim FileGUID : FileGUID = Uguid()
		Dim FileContentType : FileContentType = FileMimeType(s3_strRemoteFile)
		Dim strNowInGMT: strNowInGMT = NowInGMT()
		Dim strStringToSign: strStringToSign = _
		  "GET" & vbLf & _
		  "" & vbLf & _
		  ""&FileContentType&"" & vbLf & _
		  strNowInGMT & vbLf & _
		  "/" & s3_strBucket + "/" & s3_strRemoteFile
		Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
		Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature
		
		Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
			xhttp.open "GET", "https://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
            xhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
			xhttp.setRequestHeader "Content-Type", ""&FileContentType&""
			xhttp.setRequestHeader "Date", strNowInGMT
			xhttp.setRequestHeader "Authorization", strAuthorization
			xhttp.send
			If xhttp.status = "200" Then
				Set oStream = Server.CreateObject("ADODB.Stream")
				oStream.Open
				oStream.Type = 1
				oStream.Write xhttp.responseBody
				oStream.SaveToFile s3_strLocalFile, 2
				oStream.Close
			  s3_Download = "1"
			Else
			  s3_Download = "0:" & xhttp.responseText
			End If
		Set xhttp = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_StreamToBrowser()
		Dim FileGUID : FileGUID = Uguid()
		Dim FileContentType : FileContentType = FileMimeType(s3_strLocalFile)
		Dim strNowInGMT: strNowInGMT = NowInGMT()
		Dim strStringToSign: strStringToSign = _
		  "GET" & vbLf & _
		  "" & vbLf & _
		  ""& FileContentType &"" & vbLf & _
		  strNowInGMT & vbLf & _
		  "/" & s3_strBucket + "/" & s3_strRemoteFile
		Dim strSignature: strSignature = ConvertBytesToBase64(HMACSHA1(strSecretAccessKey, strStringToSign))
		Dim strAuthorization: strAuthorization = "AWS " & strAccessKeyID & ":" & strSignature

		' Dim xhttp: Set xhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Dim xhttp: Set xhttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
			xhttp.open "GET", "https://" & s3_strBucket & ".s3.amazonaws.com/" & s3_strRemoteFile, False
            xhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
			xhttp.setRequestHeader "Content-Type", ""& FileContentType &""
			xhttp.setRequestHeader "Date", strNowInGMT
			xhttp.setRequestHeader "Authorization", strAuthorization
			xhttp.send

			If xhttp.status = "200" Then
				Set oStream = Server.CreateObject("ADODB.Stream")
				oStream.Open
				oStream.Type = 1
				oStream.Write xhttp.responseBody
				TempFile = strLocalTempDir & Uguid()
				oStream.SaveToFile TempFile & fname, 2

				Response.ContentType = FileMimeType(s3_strOutFileName)
				' select case lcase(right(s3_strOutFileName,3))
				' 	case "pdf"
				' 		Response.ContentType = "application/pdf"
				' 	case "htm","tml"
				' 		Response.ContentType = "text/HTML"
				' 	case "gif"
				' 		Response.ContentType = "image/GIF"
				' 	case "jpg","peg"
				' 		Response.ContentType = "image/JPEG"
				' 	case "txt"
				' 		Response.ContentType = "text/plain"
				' 	case "zip"
				' 		Response.ContentType = "application/zip"
				' 	case Else
				' 		Response.ContentType = "application/octet-stream"
				' end select
				Response.Charset = "UTF-8"
				Response.AddHeader "Content-Disposition", "attachment; filename="& s3_strOutFileName
				
				oStream.LoadFromFile(TempFile)
				
				do while not oStream.EOS
					response.binaryWrite oStream.read(3670016) 
					response.flush
				loop
				
				oStream.Close
			        Set objFSO = Createobject("Scripting.FileSystemObject")
			        If objFSO.Fileexists(TempFile) Then objFSO.DeleteFile TempFile
			        Set objFSO = Nothing

			    Set oStream = Nothing
			End If
		Set xhttp = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function s3_DeleteLocalFile()
        Set objFSO = Createobject("Scripting.FileSystemObject")
	        If objFSO.Fileexists(s3_strLocalFile) Then 
	        	objFSO.DeleteFile s3_strLocalFile
	        End If
        Set objFSO = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
End Class



' ' upload file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_LocalFile					= "/content/files/file.zip"
' 	Amazon.s3_RemoteFile				= "path/file.zip"
' 	result = Amazon.s3_Upload()
' 	Amazon.s3_DeleteLocalFile()
' 	response.write result
' Set Amazon = Nothing

' ' upload stream file'
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_LocalFile					= "/content/files/file.zip"
' 	Amazon.s3_RemoteFile				= "path/file.zip"
' 	result = Amazon.s3_Upload()
' 	response.write result
' Set Amazon = Nothing

' upload stream file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_UploadBinary				= binary
' 	Amazon.s3_RemoteFile				= "path/test.zip"
' 	result = Amazon.s3_UploadBinary()
' 	response.write result
' Set Amazon = Nothing

' ' download file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_LocalFile					= "/content/files/file.pdf"
' 	Amazon.s3_RemoteFile				= "test.pdf"
' 	result = Amazon.s3_Download()
' 	response.write result
' Set Amazon = Nothing

' ' stream file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_RemoteFile				= "test.pdf"
' 	Amazon.s3_OutFileName				= "test.pdf"
' 	Amazon.s3_StreamToBrowser()
' Set Amazon = Nothing

' ' delete local file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_LocalFile					= "test.pdf"
' 	Amazon.s3_DeleteLocalFile()
' Set Amazon = Nothing

' ' delete remote file
' Set Amazon = New amazon_s3_plugin
' 	Amazon.s3_RemoteFile				= "test.pdf"
' 	Amazon.s3_Delete()
' Set Amazon = Nothing
%>