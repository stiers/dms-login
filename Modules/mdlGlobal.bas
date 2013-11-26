Option Explicit

'*****************************************************************
'* Global Module
'* Author: Ephramar Telog
'* Created: September 18, 2013
'* E-mail: ephramar@outlook.com
'*
'* Credits:
'* - Michael Ciurescu (CVMichael from vbforums.com)
'*
'* Copyright � 2013 Guill-Bern Corporation. All rights reserved.
'*****************************************************************

Public rsData As ADODB.Recordset
Public DBConn As ADODB.Connection

Public LogInUserID As Long
Public LogInUserName As String
Public LogInUserType As Integer

Public Function LoadDatabase(ByVal DatabaseName As String, Optional ByVal UserID As String, Optional ByVal Password As String) As ADODB.Connection
    Dim conData As ADODB.Connection
    
    Set conData = New ADODB.Connection
    
    conData.Provider = "Microsoft.Jet.OLEDB.4.0"
    conData.ConnectionString = "Data Source = " & DatabaseName
    conData.CursorLocation = adUseClient
    conData.Open , UserID, Password
    
    Set LoadDatabase = conData
End Function

Public Function FormAllowed(ByVal FormName As String) As Boolean
    If LogInUserType = 2 Then
        FormAllowed = True 'If ADMIN allow
    ElseIf Not (DBConn Is Nothing) Then
        'If USER, check to see if form is in the group
        
        FormAllowed = DBConn.Execute("SELECT Count(*) FROM (tblUsers AS u " & _
            "INNER JOIN tblUserGroupPrivileges AS gp ON u.UserGroupID = gp.GroupID) " & _
            "INNER JOIN tblForms AS f ON gp.FormID = f.FormID " & _
            "WHERE u.[ID] = " & LogInUserID & " AND LCase(f.ObjectName) = LCase('" & FormName & "')").Fields(0).Value > 0
    End If
End Function
