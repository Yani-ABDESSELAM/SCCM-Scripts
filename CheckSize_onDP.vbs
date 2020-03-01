'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: 
'
' AUTHOR: Microsoft , Microsoft
' DATE  : 23/10/2014
'
' COMMENT: 
'
'==========================================================================
On Error Resume Next

Const sServerName="PARS000scp03"
Const sDataBase="SMS_FP3"
Const sDataBaseUser="sms_read"
Const sDataBasePassword="sms_read"


const adOpenStatic = 3
Const adLockOptimistic = 3

Dim strToReplace
strToReplace="MSWNET:[""SMS_SITE=" & Replace(sDataBase,"SMS_","") & """]"

Set objConnection = CreateObject("ADODB.Connection")		
objConnection.Open "Provider=SQLOLEDB;Data Source=" & strComputer & ";" & ";Initial Catalog=" & sDataBase & ";User ID=" & sDataBaseUser & ";Password=" & sDataBasePassword
Dim objFso, objFolderSrc, objFolderDest
Set objFso = CreateObject("Scripting.FileSystemObject")


Set objRecordSet = CreateObject("ADODB.Recordset")
sSqlRequest = "SELECT  v_Package.Description, v_Package.PkgSourcePath, v_PackageStatusDistPointsSumm.SourceNALPath " &_
	"FROM v_Package " &_
	"inner join v_PackageStatusDistPointsSumm on v_PackageStatusDistPointsSumm.PackageID = v_Package.PackageID " &_
	"WHERE Description like '%CustomUniqueID%'"

objRecordSet.Open sSqlRequest, objConnection, adOpenStatic, adLockOptimistic
Do Until objRecordSet.EOF
	'wscript.Echo objRecordSet.Fields.Item(0).Value '& " | " & objRecordSet.Fields.Item(1).Value & " | " & Replace(objRecordSet.Fields.Item(2).Value,strToReplace,"")
	If Not objFso.FolderExists(objRecordSet.Fields.Item(1).Value) Then WScript.Echo "Le dossier " & objRecordSet.Fields.Item(1).Value & " n'existe pas"
	If Not objFso.FolderExists(Replace(objRecordSet.Fields.Item(2).Value,strToReplace,"")) Then WScript.Echo "Le dossier " & Replace(objRecordSet.Fields.Item(2).Value,strToReplace,"") & " n'existe pas"

	If objFso.FolderExists(objRecordSet.Fields.Item(1).Value) And objFso.FolderExists(Replace(objRecordSet.Fields.Item(2).Value,strToReplace,"")) Then 
		set objFolderSrc = objFso.GetFolder(objRecordSet.Fields.Item(1).Value)
		Set objFolderDest = objFso.GetFolder(Replace(objRecordSet.Fields.Item(2).Value,strToReplace,""))
		If objFolderSrc.Size<>objFolderDest.Size Then 
			WScript.Echo "Erreur avec la package " & objRecordSet.Fields.Item(0).Value & " Source " & objFolderSrc.Size & " Destination " & objFolderDest.Size	
		End If	
	End If

	set objFolderSrc = Nothing
	Set objFolderDest = Nothing

	objRecordSet.MoveNext	
Loop