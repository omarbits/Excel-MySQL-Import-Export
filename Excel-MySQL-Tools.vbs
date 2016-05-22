
'''''''''''''''''''''''''''''''''''''''
'     Database Connection Variables
'''''''''''''''''''''''''''''''''''''''

Dim DB_Host As String           ' Database Host ex: localhost
Dim DB_Name As String           ' Database Name
Dim DB_Username As String       ' Database Username
Dim DB_Password As String       ' Database Password
Dim ODBC_Driver As String       ' Database Connector Name
Dim Conn As ADODB.Connection    ' Database Connection Object
Dim rs As ADODB.Recordset       ' Database Recordset
Dim fld As ADODB.Field          ' Database Field

Const debug_mode = False        ' Debugging Variable

Sub Backup_Data()

    On Error GoTo ErrorHandler
    
    Call delete_all_data_sheets
    
    Call database_connect
    
    Call database_to_sheets(output_table_data:=True)
    
Done:
    
    Call database_disconnect
        
    Exit Sub
    
ErrorHandler:
    
    Call ado_errorhandler
    
    On Error Resume Next
    
    GoTo Done

End Sub

Sub Tables_Structure()

    On Error GoTo ErrorHandler
    
    Call delete_all_data_sheets
    
    Call database_connect
    
    Call database_to_sheets(output_table_data:=False)
    
Done:
    
    Call database_disconnect
        
    Exit Sub
    
ErrorHandler:
    
    Call ado_errorhandler
    
    On Error Resume Next
    
    GoTo Done

End Sub

Sub Append_Data()

    On Error GoTo ErrorHandler
    
    Call database_connect
    
    Call sheets_to_database(replace:=False)
    
Done:
    
    Call database_disconnect
        
    Exit Sub
    
ErrorHandler:
    
    Call ado_errorhandler
    
    On Error Resume Next
    
    GoTo Done

End Sub


Sub Restore_Data()

    On Error GoTo ErrorHandler
    
    Call database_connect
    
    Call sheets_to_database(replace:=True)
    
Done:
    
    Call database_disconnect
        
    Exit Sub
    
ErrorHandler:
    
    Call ado_errorhandler
    
    On Error Resume Next
    
    GoTo Done

End Sub

Function database_connect()

    ''''''''''''''''''''''''''''''''''''
    '   Form Data Input
    ''''''''''''''''''''''''''''''''''''
    
    DB_Host = Range("Database.Host").Value          ' Get Value from Excel Sheet
    DB_Name = Range("Database.Name").Value          ' Get Value from Excel Sheet
    DB_Username = Range("Database.Username").Value  ' Get Value from Excel Sheet
    DB_Password = Range("Database.Password").Value  ' Get Value from Excel Sheet
    ODBC_Driver = Range("Database.Driver").Value
    
    Set Conn = New ADODB.Connection                 'Connect to MySQL server using Connector/ODBC
    
    Select Case ODBC_Driver
    
        Case "MySQL ODBC 3.51 Driver"
        Conn.ConnectionString = "Driver={" + ODBC_Driver + "}; Server=" + DB_Host + "; Database=" + DB_Name + "; User=" + DB_Username + "; Password=" + DB_Password + "; Option=3; Port=3306; charset=UTF8;"
     
        Case "MySQL ODBC 5.1 Driver"
        Conn.ConnectionString = "Driver={" + ODBC_Driver + "}; Server=" + DB_Host + "; Database=" + DB_Name + "; Uid=" + DB_Username + "; Pwd=" + DB_Password + "; Option=3; Port=3306; charset=utf8;"
        
    End Select
        
    If debug_mode Then MsgBox Conn.ConnectionString
     
    Conn.Open
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseServer

End Function

Function database_disconnect()

   ' Close all open objects.
     If Conn.State = adStateOpen Then
        
        Conn.Close
      
      End If

   ' Destroy anything not destroyed yet.
     Set Conn = Nothing

End Function

Function sheets_to_database(Optional ByVal replace As Boolean = False)
    
    Dim Column_Fields() As String   ' Table column field names
    Dim Row_Values() As String      ' Row values to be inserted
    Dim cell_value As String        ' current cell value
    Dim i As Integer                ' number of columns iterators
    Dim j As Integer                ' number of rows iterators
    Dim m As Integer                ' total number of columns
    
    For Each ws In Worksheets
        
        If Left(ws.Name, 1) <> "." Then 'Ignore sheets with names starting with a dot ex: .config
            
            If replace = True Then
                
                Conn.Execute "DELETE FROM " & ws.Name ' Removes all existing data inside table before inserting new data
            
                Conn.Execute "ALTER TABLE " & ws.Name & " AUTO_INCREMENT = 1"
                 
            End If
            
            i = 0
              
            Do While ActiveWorkbook.Worksheets(ws.Name).Cells(1, i + 1).Value <> "" 'Loop through the first row for column names
                
                i = i + 1
                
                ReDim Preserve Column_Fields(0 To i - 1)
                                
                Column_Fields(i - 1) = ActiveWorkbook.Worksheets(ws.Name).Cells(1, i).Value
                
            Loop
            
            If debug_mode Then MsgBox ws.Name + " => " + Join(Column_Fields(), ", ")
            
            m = UBound(Column_Fields(), 1)
             
            ReDim Row_Values(0 To m) As String
             
            j = 0
             
            rs.Open "select * from " & ws.Name, Conn, adOpenDynamic, adLockOptimistic
             
            Do While Application.CountA(ActiveWorkbook.Worksheets(ws.Name).Cells(j + 2, 1).EntireRow) <> 0  'Check if row is not empty
                
                j = j + 1
                 
                rs.AddNew
                
                For i = 0 To m
                    
                    cell_value = CStr(ActiveWorkbook.Worksheets(ws.Name).Cells(j + 1, i + 1).Value)
                
                    rs.Fields(Column_Fields(i)) = cell_value
   
                Next i
                
                rs.Update

            Loop
            
            rs.Close
            
        End If
        
    Next ws

End Function


Function database_to_sheets(Optional ByVal output_table_data As Boolean = True)

    Dim TablesSchema As ADODB.Recordset
    Dim table_name As String
    Dim column_name As String

    Call delete_all_data_sheets
    
    'Get all database tables.
    Set TablesSchema = Conn.OpenSchema(adSchemaTables)
    
      '
      '   Loop Thrugh
      '
    
    Do While Not TablesSchema.EOF
      
        table_name = TablesSchema("TABLE_NAME")
        
        '
        '   Create Sheet with Table Name and Fill in the Column Names
        '
        
        Call create_sheet_if_not_exist(table_name)
        
        rs.Open "SELECT * FROM " & table_name, Conn
        
        rs.MoveFirst
        
        i = 1
        
        j = 1
        
        For Each fld In rs.Fields
        
            If debug_mode Then MsgBox table_name & "=> " & fld.Name
            
            ActiveWorkbook.Worksheets(table_name).Cells(i, j).Value = fld.Name
            
            j = j + 1
        
        Next
        
        i = i + 1
        
        '
        '   Get Table Data
        '
        
        If output_table_data Then
        
        Do Until rs.EOF
        
        j = 1
        
            For Each fld In rs.Fields
        
                If debug_mode Then MsgBox table_name & "=> " & fld.Name & "=> " & fld.Value
        
                ActiveWorkbook.Worksheets(table_name).Cells(i, j).Value = fld.Value
        
                j = j + 1
        
            Next
            
            i = i + 1
            
            rs.MoveNext
        
        Loop
        
        End If
        
        '
        '   End of Get Table Data
        '
        
        rs.Close
        
        TablesSchema.MoveNext
        
        Loop

End Function

Function create_sheet_if_not_exist(sheet_name As String)
    
    Dim ws As Worksheet
    
    On Error Resume Next
    
    Set ws = Worksheets(sheet_name)
    
    If Err.Number = 9 Then
        
        Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))
        
        ws.Name = sheet_name
    
    End If

End Function

Function delete_all_data_sheets()
    
    ' This function deletes all sheets except those
    ' with names starting with a dot ex: ".config"
    
    Application.DisplayAlerts = False
    
    Application.ScreenUpdating = False
    
    For Each ws In Worksheets
        
        If Left(ws.Name, 1) <> "." Then
        
        ws.Delete
        
        End If
    
    Next
    
    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True

End Function

Function ado_errorhandler()

    '''''''''''''''''''''''''''''''''''''''
    '   Error Handling
    '''''''''''''''''''''''''''''''''''''''
    '
    '   This is the error handling snippet for extracting
    '   error information from ActiveX Data Objects in VB
    '   used from the below source:
    '   https://support.microsoft.com/en-us/kb/167957
    '
    
    'Todo: add Microsoft ActiveX Data Objects 2.8 Library Programmatically
    

    Dim errLoop
    Dim strError
    Dim e
    e = 1

   ' Process
     StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(Err.Number)
     StrTmp = StrTmp & vbCrLf & "   Generated by " & Err.Source
     StrTmp = StrTmp & vbCrLf & "   Description  " & Err.Description

   ' Enumerate Errors collection and display properties of
   ' each Error object.
     Set Errs1 = Conn.Errors
     For Each errLoop In Errs1
        With errLoop
            StrTmp = StrTmp & vbCrLf & "Error #" & i & ":"
            StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
            StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
            StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
            e = e + 1
       End With
    Next

      MsgBox StrTmp

      ' Clean up Gracefully

End Function

'Todo: Make this non-technical user friendly
'
'       + Config in hidden and password protected sheet
'       + Get custom query results
'       + Error reporting show which sheet and row the error occured when inserting values
'
