<<<<<<< HEAD

'On Error Resume Next  

Dim fso, folder, fld, files, NewsFile
'Dim Folders

' SETUP PARAMETERS FROM ARGUMENTS AND DEFAULTS

	Limit=-1 'No limit
	
	set A=wscript.arguments
	
	NoAppend=True
	NoWarnAtEnd=True
	NoDebug=True
	
'	for i=0 to A.count-1
'		if A(i)="-Append" then 
'			NoAppend=False
'			A(i).delete
'		elseif A(i)="-WarnAtEnd" then 
'			NoWarnAtEnd=False
'			'Msgbox "HEY"
'			'A(i).delete
'		elseif A(i)="-Debug" then 
'			NoDebug=False
'			A(i).delete
'		end if
'	Next

	if A.count=4 then 
		If A(3)="-WarnAtEnd" then 
			NoWarnAtEnd=False
		end if 
	end if

	if A.count>=3 then 
		If Not IsNumeric(A(2)) then 
			msgbox "Syntax is: "+vbCrlf+vbCrlf+"acltest.vbs Mailboxname Outputfilename [Limit or -1] [-WarnAtEnd]"
			wscript.quit
		else 
			Limit=Cint(A(2))
		end if 
	end if
	if A.count<2 then 
		msgbox "Syntax is: "+vbCrlf+vbCrlf+"acltest.vbs Mailboxname Outputfilename [Limit or -1] [-WarnAtEnd]"
		wscript.quit
	end if
	mailboxname=A(0)
	OutputFileName=A(1)

'	Msgbox "Append:"+Cstr(Not NoAppend )+VbCrlf + _
'	"WarnAtEnd:"+Cstr(Not NoWarnAtEnd )+VbCrlf + _ 
'	"Debug:"+Cstr(Not NoDebug )+VbCrlf + _ 
'	"Mailbox Name:"+MailBoxName+VbCrlf + _ 
'	"Output FIle Name:"+OutputFileName+VbCrlf + _ 
'	"Output Path:"+FolderPath+VbCrlf + _
'	"Limit transactions to "+Cstr(limit)+ vbcrlf
	

'  Now basic parameters have been setup check for existence of output file ...

	Set fso = CreateObject("Scripting.FileSystemObject")  
'	If Instr(FolderPath,"\")>0 then 
'		FolderPath =StrReverse(Split(StrReverse(OutputFileName),"\",0))+"\"
'		Set folder = fso.GetFolder(FolderPath)  
'	else
'		Folderpath=""
'		Set folder = fso.GetFolder(".")  
'	End if
	
	
'	If sOutputFile = "" Then      
'		sOutputFile=OutputFileName
'	End If
	
	Set NewFile = fso.CreateTextFile(OutputFileName, True)
	
	Set olapp = CreateObject("Outlook.Application")
	Set fld = olapp.GetNamespace("MAPI").folders(MailboxName)
	s=""
	idx=0
	lvl=0
	Sout=""
'	Limit=50
	GetFolderItems idx, fld, lvl,nfld,nmsgs,newfile,Sout,Limit
'	NewFile.Write(Sout)  
	If Not NoWarnAtEnd then Msgbox "Finito!"

WScript.quit
'--------------------------------------------------------------------------------
'
' NDIGIT(X,N) - RETURNS INTEGER N AS A STRING 
' PADDED WITH ZEROS OF UP TO N CHARACTERS
'
' Utility
'
Function ndigit(x, n)
Dim s, i
Dim p10, k
        p10 = Array(1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000, 1000000000)
        s = ""
        k = x
        For i = 1 To n
                If IsNumeric(k) Then
                        s = CStr(k Mod 10) + s
                Else
                        s = "  " + s
                End If
                If IsNumeric(k) Then k = k \ 10
        Next
        ndigit = s
End Function
'--------------------------------------------------------------------------------
'
' GETFOLDERITEMS RECURSE THROUGH FOLDER STRUCTURE TO EXTRACT 
' FOLDERS AND ITEMS FORM FOLDERS.
'
' Utility
'
Sub GetFolderItems (ByRef idx, fld,  ByRef lvl,  Byref nfld,  nmsgs, NewFile, Sout, Limit)
Dim i,k
    If lvl = 0 Then
        s = ndigit(idx, 6) + "-" + FixedWidth(Nz(fld.FolderPath, ""), 64, "left", " ") + "-Fldr-" + ndigit(0, 3) + "-" + ndigit(lvl, 3) + "-"
        s = s + ndigit(fld.Folders.Count, 6) + "-" + ndigit(fld.Items.Count, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
        Sout = Sout + s + vbCrLf
	NewFile.writeline(s)
    End If
    
    If  Limit=-1 or idx < Limit Then
        For i = 1 To fld.Items.Count
            idx = idx + 1
            z = fld.Items(i).Subject
            z = Cleanup(z,"_")
            Dt = 0
            cls = fld.Items(i).Class
            s = ndigit(idx, 6) + "-" + FixedWidth(fld.FolderPath, 64, "left", " ") + "-Item-" + ndigit(fld.Items(i).Class, 3) + "-" + ndigit(lvl, 3) + "-"
            
            If cls = 43 Then 'email
                Dt = fld.Items(i).ReceivedTime
                sdt = ndigit(Year(Dt), 4) + ndigit(Month(Dt), 2) + ndigit(Day(Dt), 2) + "-" + ndigit(Hour(Dt), 2) + ndigit(Minute(Dt), 2) + ndigit(Second(Dt), 2)
                Dt = fld.Items(i).SentOn
                ssendt = ndigit(Year(Dt), 4) + ndigit(Month(Dt), 2) + ndigit(Day(Dt), 2) + "-" + ndigit(Hour(Dt), 2) + ndigit(Minute(Dt), 2) + ndigit(Second(Dt), 2)
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + sdt + "-" + ssendt + "-" + FixedWidth(fld.Items(i).SenderName, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
                If Err.number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo 0
            
            Else
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
            End If
            If  (idx = Limit) Then
                s = ndigit(idx, 6) + "-" + FixedWidth(fld.FolderPath, 64, "left", " ") + "-****-" + ndigit(fld.Items(i).Class, 3) + "-" + ndigit(lvl, 3) + "-"
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth("", 32, "left", "*") + "-" + FixedWidth("*** REACHED LIMIT: " + CStr(Limit) + " ", 128, "left", "*")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
                Exit For
            End If
        Next
    End If
    
    k = fld.Folders.Count
    If k = 0 Then Exit Sub
    lvl = lvl + 1
    For i = 1 To k
        idx = idx + 1
        nfld = nfld + 1
        s = ndigit(idx, 6) + "-" + FixedWidth(fld.Folders(i).FolderPath, 64, "left", " ") + "-Fldr-" + ndigit(0, 3) + "-" + ndigit(lvl, 3) + "-"
        z = nz(fld.Folders(i).Name,"")
        s = s + ndigit(fld.Folders(i).Folders.Count, 6) + "-" + ndigit(fld.Folders(i).Items.Count, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
        Sout = Sout + s + vbCrLf
	NewFile.writeline(s)
'	Msgbox s
	GetFolderItems idx, fld.Folders(i), lvl, nfld, nmsgs, NewFile, Sout, Limit
    Next
    lvl = lvl - 1
End Sub
'--------------------------------------------------------------------------------
'
' CLEANUP(X) - REPLACE UNICODE AND ? * / \  CHARACTERS WITH PAD
'
' Utility
'
Function Cleanup(x,PAD)
	PAD=Left(PAD,1)
	'Msgbox z
	y=""
	for i=1 to len(x)
		k=mid(x,i,1)
		if ascw(k)>255 or ascw(k)<32 or k="?" or k="/" or k="\" or k="*" then 
			'Msgbox cstr(asc(mid(x,i,1)))+" - "+x
			'y=left(y,i-1)+"_"+mid(y,i+1)
			y=y+PAD
		else
			y=y+mid(x,i,1)
		end if
		
	next
	cleanup=y
	'Msgbox y
	
End function
'--------------------------------------------------------------------------------
'
' FIXEDWIDTH(S,W,JUSTIF,PAD) - RETURNS STRING AS A STRING 
' PADDED WITH "PAD" CHARACTERS  UP TO W CHARACTERS WIDE
' JUSTIFIED (LEFT< CENTER OR RIGHT). TRUNCATE TO W CHARACTERS
'IF STRING LARGER AND REPLACE LAST CHARACTER WITH A STAR (*)
'
' Utility
'

Function FixedWidth(s, w, justif, pad)
    PAD=Left(PAD,1)
    If Len(s) = w Then
        FixedWidth = s
        Exit Function
    ElseIf Len(s) > w Then
        FixedWidth = Mid(s, 1, w - 1) + "*"
        Exit Function
    End If
    If justif = "center" Then
        FixedWidth = String(w - (w - Len(s)) \ 2, pad) + s + String((w - Len(s)) \ 2, pad)
    ElseIf justif = "Right" Then
        FixedWidth = String(w - Len(s), pad) + s
    Else
        FixedWidth = s + String(w - Len(s), pad)
    End If
End Function
'--------------------------------------------------------------------------------
'
' NZ(X,Y) - RETURNS Y IF X IS NULL OTHERWISE RETURNS X
'
' Utility
'

Function nz(x,y)
	if isnull(x) then 
		nz=y
	else
		nz=x
	end if
end function
=======

'On Error Resume Next  

Dim fso, folder, fld, files, NewsFile
'Dim Folders

' SETUP PARAMETERS FROM ARGUMENTS AND DEFAULTS

	Limit=-1 'No limit
	
	set A=wscript.arguments
	
	NoAppend=True
	NoWarnAtEnd=True
	NoDebug=True
	
'	for i=0 to A.count-1
'		if A(i)="-Append" then 
'			NoAppend=False
'			A(i).delete
'		elseif A(i)="-WarnAtEnd" then 
'			NoWarnAtEnd=False
'			'Msgbox "HEY"
'			'A(i).delete
'		elseif A(i)="-Debug" then 
'			NoDebug=False
'			A(i).delete
'		end if
'	Next

	if A.count=4 then 
		If A(3)="-WarnAtEnd" then 
			NoWarnAtEnd=False
		end if 
	end if

	if A.count>=3 then 
		If Not IsNumeric(A(2)) then 
			msgbox "Syntax is: "+vbCrlf+vbCrlf+"acltest.vbs Mailboxname Outputfilename [Limit or -1] [-WarnAtEnd]"
			wscript.quit
		else 
			Limit=Cint(A(2))
		end if 
	end if
	if A.count<2 then 
		msgbox "Syntax is: "+vbCrlf+vbCrlf+"acltest.vbs Mailboxname Outputfilename [Limit or -1] [-WarnAtEnd]"
		wscript.quit
	end if
	mailboxname=A(0)
	OutputFileName=A(1)

'	Msgbox "Append:"+Cstr(Not NoAppend )+VbCrlf + _
'	"WarnAtEnd:"+Cstr(Not NoWarnAtEnd )+VbCrlf + _ 
'	"Debug:"+Cstr(Not NoDebug )+VbCrlf + _ 
'	"Mailbox Name:"+MailBoxName+VbCrlf + _ 
'	"Output FIle Name:"+OutputFileName+VbCrlf + _ 
'	"Output Path:"+FolderPath+VbCrlf + _
'	"Limit transactions to "+Cstr(limit)+ vbcrlf
	

'  Now basic parameters have been setup check for existence of output file ...

	Set fso = CreateObject("Scripting.FileSystemObject")  
'	If Instr(FolderPath,"\")>0 then 
'		FolderPath =StrReverse(Split(StrReverse(OutputFileName),"\",0))+"\"
'		Set folder = fso.GetFolder(FolderPath)  
'	else
'		Folderpath=""
'		Set folder = fso.GetFolder(".")  
'	End if
	
	
'	If sOutputFile = "" Then      
'		sOutputFile=OutputFileName
'	End If
	
	Set NewFile = fso.CreateTextFile(OutputFileName, True)
	
	Set olapp = CreateObject("Outlook.Application")
	Set fld = olapp.GetNamespace("MAPI").folders(MailboxName)
	s=""
	idx=0
	lvl=0
	Sout=""
'	Limit=50
	GetFolderItems idx, fld, lvl,nfld,nmsgs,newfile,Sout,Limit
'	NewFile.Write(Sout)  
	If Not NoWarnAtEnd then Msgbox "Finito!"

WScript.quit
'--------------------------------------------------------------------------------
'
' NDIGIT(X,N) - RETURNS INTEGER N AS A STRING 
' PADDED WITH ZEROS OF UP TO N CHARACTERS
'
' Utility
'
Function ndigit(x, n)
Dim s, i
Dim p10, k
        p10 = Array(1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000, 1000000000)
        s = ""
        k = x
        For i = 1 To n
                If IsNumeric(k) Then
                        s = CStr(k Mod 10) + s
                Else
                        s = "  " + s
                End If
                If IsNumeric(k) Then k = k \ 10
        Next
        ndigit = s
End Function
'--------------------------------------------------------------------------------
'
' GETFOLDERITEMS RECURSE THROUGH FOLDER STRUCTURE TO EXTRACT 
' FOLDERS AND ITEMS FORM FOLDERS.
'
' Utility
'
Sub GetFolderItems (ByRef idx, fld,  ByRef lvl,  Byref nfld,  nmsgs, NewFile, Sout, Limit)
Dim i,k
    If lvl = 0 Then
        s = ndigit(idx, 6) + "-" + FixedWidth(Nz(fld.FolderPath, ""), 64, "left", " ") + "-Fldr-" + ndigit(0, 3) + "-" + ndigit(lvl, 3) + "-"
        s = s + ndigit(fld.Folders.Count, 6) + "-" + ndigit(fld.Items.Count, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
        Sout = Sout + s + vbCrLf
	NewFile.writeline(s)
    End If
    
    If  Limit=-1 or idx < Limit Then
        For i = 1 To fld.Items.Count
            idx = idx + 1
            z = fld.Items(i).Subject
            z = Cleanup(z,"_")
            Dt = 0
            cls = fld.Items(i).Class
            s = ndigit(idx, 6) + "-" + FixedWidth(fld.FolderPath, 64, "left", " ") + "-Item-" + ndigit(fld.Items(i).Class, 3) + "-" + ndigit(lvl, 3) + "-"
            
            If cls = 43 Then 'email
                Dt = fld.Items(i).ReceivedTime
                sdt = ndigit(Year(Dt), 4) + ndigit(Month(Dt), 2) + ndigit(Day(Dt), 2) + "-" + ndigit(Hour(Dt), 2) + ndigit(Minute(Dt), 2) + ndigit(Second(Dt), 2)
                Dt = fld.Items(i).SentOn
                ssendt = ndigit(Year(Dt), 4) + ndigit(Month(Dt), 2) + ndigit(Day(Dt), 2) + "-" + ndigit(Hour(Dt), 2) + ndigit(Minute(Dt), 2) + ndigit(Second(Dt), 2)
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + sdt + "-" + ssendt + "-" + FixedWidth(fld.Items(i).SenderName, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
                If Err.number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo 0
            
            Else
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
            End If
            If  (idx = Limit) Then
                s = ndigit(idx, 6) + "-" + FixedWidth(fld.FolderPath, 64, "left", " ") + "-****-" + ndigit(fld.Items(i).Class, 3) + "-" + ndigit(lvl, 3) + "-"
                s = s + ndigit(0, 6) + "-" + ndigit(0, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth("", 32, "left", "*") + "-" + FixedWidth("*** REACHED LIMIT: " + CStr(Limit) + " ", 128, "left", "*")
                Sout = Sout + s + vbCrLf
		NewFile.writeline(s)
                Exit For
            End If
        Next
    End If
    
    k = fld.Folders.Count
    If k = 0 Then Exit Sub
    lvl = lvl + 1
    For i = 1 To k
        idx = idx + 1
        nfld = nfld + 1
        s = ndigit(idx, 6) + "-" + FixedWidth(fld.Folders(i).FolderPath, 64, "left", " ") + "-Fldr-" + ndigit(0, 3) + "-" + ndigit(lvl, 3) + "-"
        z = nz(fld.Folders(i).Name,"")
        s = s + ndigit(fld.Folders(i).Folders.Count, 6) + "-" + ndigit(fld.Folders(i).Items.Count, 6) + "-" + "00000000-000000" + "-" + "00000000-000000" + "-" + FixedWidth(z, 32, "left", " ") + "-" + FixedWidth(z, 128, "left", " ")
        Sout = Sout + s + vbCrLf
	NewFile.writeline(s)
'	Msgbox s
	GetFolderItems idx, fld.Folders(i), lvl, nfld, nmsgs, NewFile, Sout, Limit
    Next
    lvl = lvl - 1
End Sub
'--------------------------------------------------------------------------------
'
' CLEANUP(X) - REPLACE UNICODE AND ? * / \  CHARACTERS WITH PAD
'
' Utility
'
Function Cleanup(x,PAD)
	PAD=Left(PAD,1)
	'Msgbox z
	y=""
	for i=1 to len(x)
		k=mid(x,i,1)
		if ascw(k)>255 or ascw(k)<32 or k="?" or k="/" or k="\" or k="*" then 
			'Msgbox cstr(asc(mid(x,i,1)))+" - "+x
			'y=left(y,i-1)+"_"+mid(y,i+1)
			y=y+PAD
		else
			y=y+mid(x,i,1)
		end if
		
	next
	cleanup=y
	'Msgbox y
	
End function
'--------------------------------------------------------------------------------
'
' FIXEDWIDTH(S,W,JUSTIF,PAD) - RETURNS STRING AS A STRING 
' PADDED WITH "PAD" CHARACTERS  UP TO W CHARACTERS WIDE
' JUSTIFIED (LEFT< CENTER OR RIGHT). TRUNCATE TO W CHARACTERS
'IF STRING LARGER AND REPLACE LAST CHARACTER WITH A STAR (*)
'
' Utility
'

Function FixedWidth(s, w, justif, pad)
    PAD=Left(PAD,1)
    If Len(s) = w Then
        FixedWidth = s
        Exit Function
    ElseIf Len(s) > w Then
        FixedWidth = Mid(s, 1, w - 1) + "*"
        Exit Function
    End If
    If justif = "center" Then
        FixedWidth = String(w - (w - Len(s)) \ 2, pad) + s + String((w - Len(s)) \ 2, pad)
    ElseIf justif = "Right" Then
        FixedWidth = String(w - Len(s), pad) + s
    Else
        FixedWidth = s + String(w - Len(s), pad)
    End If
End Function
'--------------------------------------------------------------------------------
'
' NZ(X,Y) - RETURNS Y IF X IS NULL OTHERWISE RETURNS X
'
' Utility
'

Function nz(x,y)
	if isnull(x) then 
		nz=y
	else
		nz=x
	end if
end function
>>>>>>> ebbcfa05064a84aabeffec3478db35d7527dd227
