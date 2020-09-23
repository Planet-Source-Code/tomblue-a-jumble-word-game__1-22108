Attribute VB_Name = "moddatabase"
Public dbwords As Database
Public rswords As Recordset
' connecting to database
' DAO is used since it is optimised for jet engine
Public Function opendatabase(Path As String) As Boolean
Dim dbpath As String
Dim strconnect As String

On Error GoTo DBERRORS
dbpath = Path
strconnect = ";DATABASE=" & dbpath

Set dbwords = DBEngine.Workspaces(0).opendatabase("", False, False, strconnect)
Set rswords = dbwords.OpenRecordset("words", dbOpenTable)
opendatabase = True
'10march
rswords.MoveLast
fieldcount = rswords.RecordCount
rswords.MoveFirst
Exit Function
DBERRORS:
MsgBox err.Description
opendatabase = False

End Function
