'Autoback.vbs [Version 2.1]
'Written by Ray Tung RayT@hwl.com.hk . All rights reserved.
'Last Updated 13/01/2014
'This version is designed to use with Task Scheduler
'
'The script provides the following functions
' - Auto backup an entire directory
' - Auto backup compression
' - Generates log file
'
' Dependencies
' - Robocopy
' - 7zip
' - CScript

' Parameters names:
' src        = source directory
' tgt        = target backup directory
' purge_dir  = Directory that needs to be checked for purging. 
' prefix     = Prefix of the backup files. i.e Viewpoint_ will produce Viewpoint_yyyy-mm-dd.zip files
' drive      = Network Drive i.e G:
' svr        = Disk name with server i.e \\Hilfile02\gisdpvcsshare 
' purge_days = no. of days before purging/delete
' logg       = Desire log location

' ------------- PARAMETERS BELOW THIS LINE --------------------
 src        = "H:\"
 tgt        = "C:\Users\test"
 purge_dir  = "C:\Users\test"
 prefix     = "H_"
 drive      = "H:"
 svr        = "\\HILFILEA01\user$\gisd\RT23389\"
 purge_days = 7
 logg       = "C:\Users\test" & "\backup.log"
 
 
 
 

 ' Parameters names:
 ' robocopy    = path of Robocopy.exe
 ' seven_z_exe = path of 7z.exe
 
 'N.B. Default parameters will be passed into command line
 '     Environment Variables are supported
 
 ' ------------- DEFAULT PARAMETERS ---------------------------
robocopy    = "%windir%\System32\Robocopy.exe"
seven_z_exe = "%programfiles%\7-Zip\7z.exe"


' ------------- DO NOT MODIFY BELOW THIS LINE ---------------------------------------------------------



' ------------- GLOBAL VARIABLES ------------------------------
tgt        = formatPath(tgt)
purge_dir  = formatPath(purge_dir)
src        = formatPath(src)
svr        = formatPath(svr)
logg       = formatPath(logg)
' Separate object for purging. Allows purging at different directory
SET obj    = CreateObject("Scripting.FileSystemObject")
SET purge_ = CreateObject("Scripting.FileSystemObject")

SET shell  = CreateObject("WScript.Shell")
SET folder = obj.GetFolder(tgt)
SET p_fold = purge_.GetFolder(purge_dir)
SET files  = folder.Files
SET p_files= p_fold.Files
SET regex  = CreateObject("VBScript.RegExp")
	today_r= Year(Date) & "-" & formatDate(Month(Date)) & "-" & Day(Date)
    today  = cdate(today_r)
SET logFile= obj.OpenTextFile(logg, 8, True)
SET Net_   = CreateObject("WScript.Network")
SET NetDrv = Net_.EnumNetworkDrives


' Set REGEX for prefix pattern matching
regex.Pattern = "^" & prefix & "(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01]).zip$" 

' ------------- FUNCTIONS -------------------------------------
FUNCTION formatDate(num)
	formatDate = num
	IF (Len(num) = 1) THEN
		formatDate = "0" & num
	END IF
END FUNCTION

FUNCTION formatPath(path)
	ch = Right(path, 1)
	IF (ch = "\") THEN
		formatPath = Mid(path, 1, Len(path)-1)
	ELSE
		formatPath = path
	END IF
END FUNCTION

' only purge files with prefix
FUNCTION Purge()
	FOR EACH f IN p_files
		' regex.test() returns -1 if pattern matches. Else returns 0
		IF (regex.Test(f.Name) = -1) THEN
			tmp = RmExt(f.Name)
			tmp = RmPrfx(tmp)
			IF today - cdate(tmp) >= purge_days THEN
			
				' Treat this as Exception handling. Err.Number = 0 if nothing is wrong. 
				ON ERROR RESUME NEXT
				tmp_name = f.Name
				purge_.DeleteFile(f)
				IF (Err.Number = 0) THEN
					CALL WriteLog(2, tmp_name)
					Err.Clear
				ELSE
					CALL WriteLog(3, tmp_name)
					Err.Clear
				END IF
				ON ERROR GOTO 0
				
			END IF
		END IF
	NEXT
END FUNCTION	
 
' Remove File Extention from file name. 
' E.g Viewpoint_2013-01-01.zip --> Viewpoint_2013-01-01
FUNCTION RmExt(fileName)
	pos   = InstrRev(fileName, ".")
	RmExt = Mid(fileName, 1, pos-1)
END FUNCTION

' Remove Prefix from file name
' E.g Viewpoint_2013-01-01.zip --> 2013-01-01.zip
FUNCTION RmPrfx(fileName)
	pos    = Len(prefix)
	RmPrfx = Mid(fileName, pos+1)
END FUNCTION
 
' NetDrv.Item() is a list that stores all net drive info
' Presumably the list stores drive & server information together
' Hence it is required to increment by 2 in every iteration
'   StrComp() is a function that compares 2 strings. Ignores Case
'             Returns 0 if equals, -1 otherwise
FUNCTION CheckNetDriveInfo()
	j = 1
	FOR i = 0 TO NetDrv.Count-1 STEP 2
		compSvr = StrComp(NetDrv.Item(i+1), svr, vbTextCompare)
		compDrv = StrComp(NetDrv.Item(i), drive, vbTextCompare)
		IF (compSvr = 0 AND compDrv = 0) THEN
			ConnectNetDrive()
			j = 0
			EXIT FOR
		END IF
	NEXT
	
	' If the given drive & server combination is no found
	IF ( j = 1 ) THEN
		MapDrive()
		ConnectNetDrive()
	END IF
END FUNCTION
 
FUNCTION ConnectNetDrive()
	CALL shell.Run("net use /p:yes", 0, true)
	CALL WriteLog(4, svr)
END FUNCTION

FUNCTION MapDrive()
	ON ERROR RESUME NEXT
		CALL Net_.MapNetworkDrive(drive, svr)
		IF (Err.Numbr = 0) THEN
			CALL WriteLog(5, drive&svr)
		ELSE
			CALL WriteLog(6, drive&svr)
			Finish()
			WScript.Quit
		END IF
	ON ERROR GOTO 0
END FUNCTION

FUNCTION Backup()
	loc      = tgt & "\" & prefix & today_r
	' Chr(34) & xxxx & Chr(34)
	' will result in "(value xxx represents)"
	' Reason being VBS does not enforce type checking.
	' Double quoting directly will result in the variable
	' being treated as String
	ssrc     = Chr(34) & src & Chr(34)
	target   = Chr(34) & loc & Chr(34)
	seven_z  = Chr(34) & seven_z_exe & Chr(34)
	
	'r_flags = flags for Robocopy.exe
	'z_flags = flags for 7z.exe
	r_flags  = " /E /log:"& Chr(34) & loc & "\" & today_r & ".log" & Chr(34)
	z_flags  = "a"
	
	IF (obj.FileExists(loc&".zip")) THEN
		CALL WriteLog(11, loc & ".zip")
	END IF	
	
	CALL CreateFolder(loc)
	CALL Copy(robocopy, ssrc, target, r_flags)
	CALL Compress(seven_z, z_flags, loc)
	CALL Remove_Uncompressed_Folder(loc)
END FUNCTION

FUNCTION CreateFolder(loc)
	CALL WriteLog(7, loc)
	ON ERROR RESUME NEXT
		CALL obj.CreateFolder(loc)
		ErrLog()
	ON ERROR GOTO 0
END FUNCTION

FUNCTION Copy(robocopy, ssrc, target, args)
	CALL WriteLog(8, NULL)
	ON ERROR RESUME NEXT
		CALL shell.Run(robocopy &" " & ssrc &" " & target & args, 0, true)
		ErrLog()
	ON ERROR GOTO 0
END FUNCTION

' Directory Compression
' Calls 7z command line interface
FUNCTION Compress(seven_z, z_flags, loc)
	CALL WriteLog(9, loc)
	ON ERROR RESUME NEXT
		CALL shell.Run(seven_z & " " & z_flags & " " & Chr(34) & loc & ".zip" & Chr(34) & " " & Chr(34) & loc & "\*" & Chr(34) , 0, true)
		ErrLog()
	ON ERROR GOTO 0
END FUNCTION

FUNCTION Remove_Uncompressed_Folder(loc)
	CALL obj.DeleteFolder(loc)
END FUNCTION

FUNCTION Cleanup()
	SET obj    = NOTHING
	SET shell  = NOTHING
	SET folder = NOTHING
	SET files  = NOTHING
	SET regex  = NOTHING
END FUNCTION

' ------------- LOG FUNCTIONS --------------------------------
FUNCTION Start()
	CALL WriteLog(1, NULL)
END FUNCTION

FUNCTION Finish()
	CALL WriteLog(10, NULL)
END FUNCTION

FUNCTION ErrLog()
	'Err.number = 0 --> no error
	IF (Err.Number = 0) THEN
			CALL WriteLog(12, NULL)
		ELSE
			CALL WriteLog(13, Err.Description)
			Finish()
			WScript.Quit
		END IF
	Err.Clear
END FUNCTION

FUNCTION WriteLog(cmd, param)
	str = ""
	IF (cmd = 1) THEN
		logFile.WriteLine("")
		str = "Backup started"
	ELSEIF (cmd = 2) THEN
		str = "Purged:          " & param
	ELSEIF (cmd = 3) THEN
		str = "Purge Failed:    " & param
	ELSEIF (cmd = 4) THEN
		str = "Connecting:      " & param
	ELSEIF (cmd = 5) THEN
		str = "Connected"
	ELSEIF (cmd = 6) THEN
		str = "ERROR:           Cannot connect to " & drive &  svr 
	ELSEIF (cmd = 7) THEN
		str = "Creating folder: " & param
	ELSEIF (cmd = 8) THEN
		str = "Copying..."
	ELSEIF (cmd = 9) THEN
		str = "Compressing:     " & param
	ELSEIF (cmd = 10) THEN
		str = "Finished"
	ELSEIF (cmd = 11) THEN
		str = param & " already exist. Overwriting"
	ELSEIF (cmd = 12) THEN
		str = "Done"
	ELSEIF (cmd = 13) THEN
		str = "Failed:          " & param
	END IF

	logFile.WriteLine("[" & today_r & " " & Time & "] " & str)
END FUNCTION

' ------------- RUN -------------------------------------------
Start()
Purge()
CheckNetDriveInfo()
Backup()
Finish()
Cleanup()

