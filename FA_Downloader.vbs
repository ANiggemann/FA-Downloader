'---------------------------------------------------------------------
' Downloader for Toshiba FlashAir Card
' Author: Andreas Niggemann, Speyer
' Web Site: http://www.lichtbildner.net
'
' License: CreativeCommons Zero (CC0)
'
' Versions:
'   21.05.2015 Successfull tested with FlashAir 3
'   26.12.2014 Ping function
'   26.09.2014 File date setting for FlashAir 2
'   05.02.2014 Support for Toshiba FlashAir 2 added
'   02.02.2014 Automatic detection of card type
'   01.02.2014 Card IP/Name instead of URL
'   22.01.2014 Support for Transcend WIFI added
'   25.03.2013 Remember first start, show manual
'   21.03.2013 Logging
'   24.02.2013 Start external CMD after each download
'   19.02.2013 Error corrected in command line parsing
'   19.02.2013 Messages via command line parameters
'   18.02.2013 Start-/Shutdown messages
'   16.02.2013
'---------------------------------------------------------------------

'- Configuration -----------------------------------------------------
' Name/IP of WiFi-Card
' FlashAir:
CARD = "flashair"
' Transcend WIFI:
' CARD = "192.168.169.3"

' Destination folder (local disc)
localfolder = "C:\TRANSFER\FLASHAIR\"

' File types to transfer (comma separated list)
filetypes = "JPG"

' Logging: 0 = no, 1 = standard, 2 = verbose
' Log file in TEMP directory
loglevel = 1
'---------------------------------------------------------------------

cardtype = ""
F2DateTimes = ""
                                     
' Main program
cl_param1 = ""
cl_param2 = ""
cl_param3 = ""
If WScript.Arguments.Count > 0 Then cl_param1 = WScript.Arguments(0)
If WScript.Arguments.Count > 1 Then cl_param2 = WScript.Arguments(1)
If WScript.Arguments.Count > 2 Then cl_param3 = WScript.Arguments(2)  
Set objFSO = CreateObject("Scripting.FileSystemObject")
Select Case cl_param1 ' URL or command M or G
  Case "M" ' Message output
    Mess cl_param2, CInt(cl_param3)
  Case "G" ' WGET File
    s = GetURL(cl_param2,cl_param3)
  Case "P" ' Ping URL
    s = Ping(cl_param2)
	If s = true Then wscript.echo "alive"
  Case else ' Standard processing
    StdProc cl_param1, cl_param2, cl_param3 
End Select
Set objFSO = Nothing
Wscript.Quit(0)

' Standard processing
Sub StdProc(clp1,clp2,clp3)
  ' FirstTime
  ExtPrg = UCASE(WScript.ScriptFullName)
  ExtPrg = Replace(ExtPrg,".VBS",".CMD") ' Build external cmd file name
  If objFSO.Fileexists(ExtPrg) Then 
    ExternalPrg = ExtPrg ' with external cmd
  Else 
    ExternalPrg = "" ' no external cmd
  End If
  tempfile = MakeTempFile()
  MessExt "FA_Downloader starting...", "2"
  LogWrite "FA_Downloader started",2
  If clp1 <> "" And clp2 <> "" Then ' with command line parameters
    CARD = clp1
    localfolder = clp2                        
    filetypes = clp3
  End If
  Do ' Loop until temp file is deleted
    If cardtype = "" Then ' Detect Card
      FLASHAIR_URL = "http://" & CARD & "/DCIM"
      TRANSCEND_URL = "http://" & CARD & "/sd/DCIM/"   
      DetectCard CARD, FLASHAIR_URL, TRANSCEND_URL, cardtype
      If cardtype <> "" Then
        If cardtype = "W" Then DCIMURL = TRANSCEND_URL 
        If cardtype = "F1" Or cardtype = "F2" Then DCIMURL = FLASHAIR_URL 
        ftypes = UCASE(filetypes) & ","
		Select Case cardtype ' W or F1 or F2
        Case "W" 
		  cardname = "Transcend WIFI"
		Case "F1"
          cardname = "Toshiba FlashAir 1"		
		Case "F2"
          cardname = "Toshiba FlashAir 2 or 3"	
		End Select
        MessExt cardname & " Card detected", "2"
        LogWrite "Parameter URL: " & DCIMURL,2
        LogWrite "Parameter target folder: " & localfolder,2
        LogWrite "Parameter file types: " & filetypes,2
      End If
    Else ' Card is detected
      counter = Download(DCIMURL,localfolder,ftypes,ExternalPrg,cardtype)
    End If
  Loop While objFSO.Fileexists(tempfile)  
  MessExt "FA_Downloader is shutting down", "5"
  LogWrite "FA_Downloader stopped",2
End Sub

' Loop thru all dirs and files
Function Download(durl,lclfolder,fitypes,extprg,sdtype)
  sc = 0
  str = GetURL(durl,"")
  LogWrite str,9
  If str <> "ERROR" Then
    LogWrite "Get filelist",2
    folderlist = GetList(str,"",sdtype)
    LogWrite folderlist,3
    folders = split(folderlist,",")
    filelist = ""
    F2DateTimes = ""
    For Each flashfolder In folders
      furl = durl & "/" & flashfolder 
      str = GetURL(furl,"")
      If str <> "ERROR" Then
        filelist = filelist & GetList(str,furl & "/",sdtype) & ","
      Else ' If error stop all
        filelist = ""
        Exit For  
      End If  
    Next
    If len(filelist) > 1 Then ' Filelist filled 
      sc = CopyIt(filelist,lclfolder,fitypes,extprg) ' Get all files
    End If  
  End If
  Download = sc
End Function

' Detect cardtype 
Sub DetectCard(ByRef c_addr, ByRef f_url, ByRef t_url, ByRef c_type)
  resu = ""
  cardname = ""
  card_address = c_addr
  field4 = 0
  Do
    If Ping(card_address) = true Then
      getu = UCASE(GetURL(f_url,""))
      If instr(getu,"404 NOT FOUND") = 0 And instr(getu,"ERROR") = 0 Then
        If instr(getu,"/DCIM,") > 0 Then
          resu = "F1"
          cardname = "Toshiba FlashAir 1"
        Else
          resu = "F2"
          cardname = "Toshiba FlashAir 2 or 3"
        End If
      Else
        getu = UCASE(GetURL(t_url,""))
        If instr(getu,"404 NOT FOUND") = 0 And instr(getu,"ERROR") = 0 Then 
          resu = "W"
          cardname = "Transcend WIFI"
        End If
      End If
      If resu <> "" Then 
        LogWrite "Card Type '" & cardname & "' detected at " & card_address,2
		c_addr = card_address
      End If
	Else
	  If instr(card_address,".") <> 0 Then ' IP
        ip_parts = Split(card_address,".")
	    If Ubound(ip_parts) = 3 Then 
	      field4 = CInt(ip_parts(3))
		  field4 = field4 + 1
		  new_card_address = ip_parts(0) & "." & ip_parts(1) & "." & ip_parts(2) & "." & field4
		  f_url = Replace(f_url,card_address,new_card_address)
		  t_url = Replace(t_url,card_address,new_card_address)
		  card_address = new_card_address
	    End If
      End If		
	End If
  Loop Until resu <> "" Or instr(card_address,".") = 0 Or field4 = 254
  c_type = resu 
End Sub

' Get HTML source or save binary file (photo)
Function GetURL(URL,saveTo)
  Dim Http
  On Error Resume Next 
  Err.Clear
  URL = RemoveDoubleSlashes(URL)
  LogWrite "URL: " & URL,3
  Set Http = CreateObject("MSXML2.XMLHTTP")
  Http.Open "GET", URL, False
  Http.Send
  If Err.Number = 0 Then ' Connection OK
    If saveTo <> "" Then ' Filename to save 
      LogWrite "File: " & saveto,3
      If Http.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 ' adTypeBinary
        objADOStream.Write Http.ResponseBody
        objADOStream.Position = 0 ' Set the stream position to the start
        If objFSO.Fileexists(saveTo) Then objFSO.DeleteFile saveTo
        objADOStream.SaveToFile saveTo
        objADOStream.Close
        SetFileDateTime(saveTo)
        Set objADOStream = Nothing    
        resu = "OK"
      End If
    Else ' HTML-Source
      resu = Http.ResponseText
    End If
  Else ' Error 
    Err.Clear
    resu = "ERROR"   
  End If
  On Error Goto 0  
  GetURL = resu 
  Set http = Nothing
End Function

' Get list of folders or files
Function GetList(htmls,prefix,sdcardtype)
  temps = Replace(htmls,vbCrLf,vbCr)
  temps = Replace(temps,vbLf,vbCr)
  lines = Split(temps,vbCr)
  fl = ""
  Select Case sdcardtype ' W or F1 or F2
  Case "W" ' Transcend WIFI
    findstr = "href=" & Chr(34)
    excludestr = "Parent Directory"
    splitchar = Chr(34)
    elemnumber = 3
  Case "F1" ' FlashAir1
    findstr = "/DCIM"
    excludestr = "__TSB"
    splitchar = ","
    elemnumber = 1
  Case "F2" ' FlashAir2
    findstr = "/DCIM"
    excludestr = "__TSB"
    splitchar = Chr(34)
    elemnumber = 7
  End Select
  For Each l In lines
    If instr(l,findstr) > 1 And instr(l,excludestr) = 0 Then ' Photo folder or photo file nanme
      lparts = split(l,splitchar)
      element = trim(lparts(elemnumber))
      If sdcardtype = "F2" And prefix <> "" Then ' Erstelldatum und Uhrzeit bei FlashAir2
        f2Date = GetFDateTime("D",makeNum(lparts(14)))
        f2Time = GetFDateTime("T",makeNum(lparts(16)))
        F2DateTimes = F2DateTimes & "||"  & localfolder & element & "|" & f2Date & " " & F2Time
      End If
      If element <> "" Then fl = fl & prefix & element & ","
    End If    
  Next
  If Len(fl) > 1 Then fl = Left(fl,Len(fl)-1) ' Delete last comma
  GetList = fl
End Function

' Copy remote files to local folder
Function CopyIt(flist,lcldest,fty,external)
  copycount = 0 
  flist = RemoveDoubleSlashes(flist)
  fl = split(flist,",")
  For Each urlfname In fl ' For all files
    If len(urlfname) > 1 Then
      parts = split(urlfname,"/") 
      destfname = parts(ubound(parts)) ' Get filename
      parts = split(destfname,".") 
      ext = parts(ubound(parts)) ' Get File extension
      If Len(fty) < 2 Or instr(fty,UCASE(ext) & ",") > 0 Then ' File extension ok
        fname = lcldest & destfname 
        If Not objFSO.Fileexists(fname) Then
          LogWrite "Download " & urlfname & " to " & fname,1 
          success = GetURL(urlfname,fname)
          If success = "ERROR" Then Exit For ' Error getting File = stop all
          If success = "OK" Then ' Count transfered files
            copycount = copycount + 1
            If external <> "" Then StartExternal external, urlfname, fname 
          End If  
        End If 
      End If    
    End If    
  Next
  CopyIt = copycount ' Number of files copied
End Function

' Remove all double slashes except at http://
Function RemoveDoubleSlashes(path)
  result = path
  result = Replace(result,"//","/")
  result = Replace(result,"http:/","http://")
  RemoveDoubleSlashes = result
End Function

' Make a temp file
Function MakeTempFile()
  temppath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
  tempf = temppath & "\FA_Downloader_" & DatePart("h", Now) & DatePart("n", Now) & DatePart("s", Now) & ".tmp"
  If objFSO.Fileexists(tempf) Then objFSO.DeleteFile tempf
  Set objOutputFile = objFSO.CreateTextFile(tempf, TRUE)
  objOutputFile.WriteLine("FA_Downloader")
  objOutputFile.Close
  MakeTempFile = tempf
End Function

' Show message with delay
Sub Mess(text,delayt)
  Set objShell = CreateObject("Wscript.Shell")
  button = objShell.Popup(text, delayt, "FA_Downloader", wshInformationMark + 4096)
  Set objShell = Nothing  
End Sub

' Show Message via FA_Downloader special command line parameter
Sub MessExt(text,delayt)
  prg = WScript.ScriptFullName
  pr = Chr(34) & prg & Chr(34) & " M " & Chr(34) & text & Chr(34) & " " & delayt
  ExecPrg pr, 1
End Sub

' Set File Modified  
Sub SetFileDateTime(filename)
  If cardtype = "F2" Then ' for FlashAir2
    pos = Instr(F2DateTimes,"||" & filename)
    If pos > 0 Then ' File found in list
      dtString = Mid(F2DateTimes,pos+Len(filename)+3,19)
      If Len(dtString) = 19 Then ' Process DateTime
        parts = split(filename,"\") 
        destfname = parts(ubound(parts)) ' Get filename only
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.NameSpace(localfolder)
        Set objFolderItem = objFolder.ParseName(destfname)
        objFolderItem.ModifyDate = dtString
        Set objShell = Nothing
      End If
    End if
  End If 
End Sub

' Convert FlashAir2 date or time into Strings 
Function GetFDateTime(typ, num)
  resu = ""
  binnum = Right("0000000000000000" & toBin2(num),16)
  If typ = "D" Then ' Date
    yy = CStr(BinToDec(Left(binnum,7)) + 1980)
    mm = CStr(BinToDec(Mid(binnum,8,4)))
    dd = CStr(BinToDec(Right(binnum,5)))
    resu = Right("00" & mm,2) & "/" & Right("00" & dd,2) & "/" & Right("0000" & yy,4)
  Else ' Time
    hh = CStr(BinToDec(Left(binnum,5)))
    nn = CStr(BinToDec(Mid(binnum,6,6)))
    se = CStr(BinToDec(Right(binnum,5)) * 2)
    resu = Right("00" & hh,2) & ":" & Right("00" & nn,2) & ":" & Right("00" & se,2)
  End IF
  GetFDateTime = resu
End Function

' Strip all non numeric characters
Function makeNum(text)
  resu = ""
  For i = 1 to Len(text)
    c = Mid(text,i,1)
    If Instr("0123456789",c) > 0 Then resu = resu & c
  Next
  makeNum = CLng(resu)
End Function

' Convert long to bin as string
Function toBin2(aInt)
  If aInt > 0 Then toBin2 = toBin2(aInt\2) & (aInt Mod 2)
End Function

' Convert bin string to long int
Function BinToDec(strBin)
  dim lngResult
  dim intIndex
  lngResult = 0
  for intIndex = len(strBin) to 1 step -1
    strDigit = mid(strBin, intIndex, 1)
    select case strDigit
      case "0"
        ' do nothing
      case "1"
        lngResult = lngResult + (2 ^ (len(strBin)-intIndex))
      case else
        ' invalid binary digit, so the whole thing is invalid
        lngResult = 0
        intIndex = 0 ' stop the loop
    end select
  next
  BinToDec = lngResult
End Function

' Hidden start external CMD with file URL and local file name as command line parameters
Sub StartExternal(prg,urlfna,fna)
  If objFSO.Fileexists(prg) Then
    pr = Chr(34) & prg & Chr(34) & " " & Chr(34) & urlfna & Chr(34) & " " & Chr(34) & fna & Chr(34)
    ExecPrg pr, 0
  End If
End Sub

' Start external program
Sub ExecPrg(cmdline,mode)
  If trim(cmdline) <> "" Then
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmdline, mode, false
    Set shell = Nothing
  End If
End Sub

' Write program log
Sub LogWrite(line,llevel)
  If llevel <= loglevel Then
    temppath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
    logfile =  temppath &"\FA_Downloader.log"
    dt = Now()
    timestamp = Year(dt) & "." & Right("0" & Month(dt),2) & "." & Right("0" & Day(dt),2) & " " _
               & Right("0" & Hour(dt),2) & ":" & Right("0" & Minute(dt),2) & ":" & Right("0" & Second(dt),2)    
    If Not objFSO.Fileexists(logfile) Then
      Set objLogFile = objFSO.CreateTextFile(logfile, TRUE)
    Else    
      Set objLogFile = objFSO.OpenTextFile(logfile, 8, True) 
    End If
    objLogFile.WriteLine(timestamp & "  " & line)  
    objLogFile.Close
  End If
End Sub

' Remember first start in registry, show manual at first start
Sub FirstTime
  On Error Resume Next
  Dim objWshShell, strkey
  strkey = "HKLM\SOFTWARE\Niggemann Software\FA_Downloader\Installed"
  Set objWshShell = CreateObject("WScript.Shell")
  w = objWshShell.RegRead(strkey)
  If Err.Number <> 0 Then ' Registry key missing
    str = GetURL("http://www.lichtbildner.net","")
    If str <> "ERROR" Then ExecPrg "http://www.lichtbildner.net/fa_manual.php", 1  ' Show manual
    objWshShell.RegWrite strkey, "1"
  End If
  Set objWshShell = nothing
  On Error Goto 0
End Sub

Function Ping(strHost)
	Dim objPing			'WMI Ping Object
	Dim objStatus			'PingStatus Object
	Ping = false
	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set objPing = objWMI.ExecQuery("SELECT * FROM Win32_PingStatus where address = '" & strHost & "'")
	'// Query Win32_PingStatus.StatusCode -> 0 = reachable
	For each objStatus in objPing
		If objStatus.StatusCode = 0 then
			Ping = true
			Exit Function
		End If
	Next
End Function
