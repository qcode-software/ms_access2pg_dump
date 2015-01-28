Sub MSAccess2PGDump(out_file As String)
  '-- Convert Microsoft Access database to a PostgreSQL dump file
  MSAccessTables2PGDump out_file, False
  MSAccessRecords2PGDump out_file, True
  MSAccessIndexes2PGDump out_file, True
  MSAccessForeignKeys2PGDump out_file, True
  MSAccessAutoNumbers2PGDump out_file, True
  
  MsgBox "PostgreSQL Dump Complete: " & out_file
End Sub

Function MSAccessField2PGDataType(field As Object)
  '-- Takes DAO (Data Access Object) field and returns the equvalent PostgreSQL data type
  Dim data_type As String
  
  Select Case field.Type
    Case dbInteger
      data_type = "smallint"
    Case dbLong
      data_type = "integer"
    Case dbBoolean
      data_type = "boolean"
    Case dbCurrency, dbNumeric, dbDecimal, dbSingle
      data_type = "numeric"
    Case dbDate
      data_type = "date"
    Case dbTime
      data_type = "time with time zone"
    Case dbTimeStamp
      data_type = "timestamp with time zone"
    Case dbDouble
      data_type = "double precision"
    Case dbMemo
      data_type = "text"
    Case dbText
      If field.size = 0 Then
        data_type = "text"
      Else
        data_type = "varchar(" & field.size & ")"
      End If
    Case Else
      MsgBox "Error unknown field type " & field.Type & " for table " & tdf.name & " field " & field.name
      Exit Function
  End Select
              
  MSAccessField2PGDataType = data_type
End Function

Function MSAccessField2PGDefaultValue(field As Object)
  '-- Takes DAO (Data Access Object) field and returns the equivalent PostgreSQL default value
  Dim default_value As String
  
  Select Case field.Type
    Case dbBoolean
      ' Cast boolean values to true and false
      If field.DefaultValue = "Yes" Then
        default_value = "true"
      Else
        default_value = "false"
      End If
    Case dbText, dbMemo
      ' quote text values with ' instead of "
      default_value = strQuote(strDeQuote(field.DefaultValue, Chr$(34)), Chr$(39))
    Case Else
      default_value = field.DefaultValue
  End Select
              
  MSAccessField2PGDefaultValue = default_value
End Function

Sub MSAccessTables2PGDump(out_file As String, Optional fileAppend As Boolean = False)
  '-- Write PostgreSQL dump file to create tables for Microsoft Access database
  Dim line As String
  Dim lines() As String
  Dim temp_lines() As String
  Dim cols() As String
  Dim field As Object
  Dim db As Database
   
  Set db = CurrentDb
  
  For Each tdf In db.TableDefs
    '-- No system tables, linked tables or temp tables
    If Left(tdf.name, 4) <> "Msys" And Left(tdf.name, 1) <> "~" And tdf.Connect = "" Then
      Erase temp_lines
      
      '-- Create table
      Call pushStr(lines, "create table " & strQuote(tdf.name, Chr$(34)) & "(")
                            
      '-- Column definitions
      For Each field In tdf.fields
        '-- Col name & data type
        line = Chr$(9) & strQuote(field.name, Chr$(34)) & " " & MSAccessField2PGDataType(field) & " "
           
        '-- Default Value
        If field.DefaultValue <> "" Then
          line = line & "default " & MSAccessField2PGDefaultValue(field) & " "
        End If
            
        '-- Column Validation Rules
        If field.ValidationRule <> "" Then
          line = line & "check(" & field.ValidationRule & ") "
        End If
            
        '-- Not Null constraints
        If field.Required Or (field.Attributes And dbAutoIncrField) <> 0 Or field.Type = dbBoolean Then
          line = line & "not null "
        End If
        
        Call pushStr(temp_lines, line)
      Next
      
      '-- Table validation rules
      If tdf.ValidationRule <> "" Then
        Call pushStr(temp_lines, Chr$(9) & "check(" & tdf.ValidationRule & ")")
      End If
      
      '-- Primary key
      Erase cols
      For Each idx In tdf.Indexes
        If idx.Primary Then
          For Each field In idx.fields
            Call pushStr(cols, strQuote(field.name, Chr$(34)))
          Next
          
          Call pushStr(temp_lines, Chr$(9) & "primary key " & "(" & Join(cols, ",") & ")")
        End If
      Next
      
      ' Add trailing commas and combine with lines array
      temp_lines = Split(Join(temp_lines, "," & vbCrLf), vbCrLf)
      For i = 0 To UBound(temp_lines)
        Call pushStr(lines, temp_lines(i))
      Next i
                            
      '-- End table def
      Call pushStr(lines, ");")
    End If
  Next
  
  '-- Write out to file
  If fileAppend Then
    Open out_file For Append As #1
  Else
    Open out_file For Output As #1
  End If
  Print #1, Join(lines, vbCrLf)
  Close #1
  
End Sub

Sub MSAccessIndexes2PGDump(out_file As String, Optional fileAppend As Boolean = False)
  '-- Write PostgreSQL dump file to create indexes for Microsoft Access database
  Dim line As String
  Dim lines() As String
  Dim cols() As String
  Dim idx_i As Integer
  Dim idx_name As String
  Dim db As Database
   
  Set db = CurrentDb
  
  For Each tdf In db.TableDefs
    '-- No system tables, linked tables or temp tables
    If Left(tdf.name, 4) <> "Msys" And Left(tdf.name, 1) <> "~" And tdf.Connect = "" Then
      idx_i = 0
      For Each idx In tdf.Indexes
        '-- Index name
        If idx.Primary Then
          idx_name = tdf.name & "_PK"
        ElseIf Left(idx.name, 9) = "REFERENCE" Then
          idx_name = tdf.name & "_FK" & Mid(idx.name, 10)
        Else
          idx_i = idx_i + 1
          idx_name = tdf.name & "_IX" & idx_i
        End If
        
        '-- Columns
        Erase cols
        For Each field In idx.fields
          line = strQuote(field.name, Chr$(34))
          If field.Attributes & dbDecending Then
            line = line & " desc"
          Else
            line = line & " asc"
          End If
          
          Call pushStr(cols, line)
        Next
        
        '-- Create index SQL
        line = "create "
        If idx.Unique Then
          line = line & "unique "
        End If
        If idx.Clustered Then
          line = line & "clustered "
        End If
        line = line & "index " & strQuote(idx_name, Chr$(34)) & " on " & strQuote(tdf.name, Chr$(34)) & " (" & Join(cols, ",") & ");"
     
        Call pushStr(lines, line)
      Next
    End If
  Next
  
  '-- Write out to file
  If fileAppend Then
    Open out_file For Append As #1
  Else
    Open out_file For Output As #1
  End If
  Print #1, Join(lines, vbCrLf)
  Close #1
    
End Sub

Sub MSAccessForeignKeys2PGDump(out_file As String, Optional fileAppend As Boolean = False)
  '-- Write PostgreSQL dump file to create foreign keys for Microsoft Access database
  Dim line As String
  Dim lines() As String
  Dim cols() As String
  Dim foreign_cols() As String
  Dim db As Database
   
  Set db = CurrentDb
  
  For Each Relation In db.Relations
    '-- No system tables, linked tables or temp tables
    If Left(Relation.Table, 4) <> "Msys" And Left(Relation.ForeignTable, 4) <> "Msys" And Left(Relation.Table, 1) <> "~" And Left(Relation.ForeignTable, 1) <> "~" Then
      '-- fk_cols
      Erase cols
      Erase foreign_cols
      For Each field In Relation.fields
        Call pushStr(cols, strQuote(field.name, Chr$(34)))
        Call pushStr(foreign_cols, strQuote(field.ForeignName, Chr$(34)))
      Next
           
      '-- SQL to create foreign key in PostgreSQL
      line = "alter table " & strQuote(Relation.ForeignTable, Chr$(34)) _
        & " add foreign key (" & Join(foreign_cols, ", ") & ")" _
        & " references " & strQuote(Relation.Table, Chr$(34)) & " (" & Join(cols, ", ") & ")"
      If (Relation.Attributes & dbRelationUpdateCascade) <> 0 Then
        line = line + " on update cascade"
      End If
      If (Relation.Attributes & dbRelationDeleteCascade) <> 0 Then
        line = line + " on delete cascade"
      End If
      line = line & ";"
              
      Call pushStr(lines, line)
    End If
  Next
  
  '-- Write out to file
  If fileAppend Then
    Open out_file For Append As #1
  Else
    Open out_file For Output As #1
  End If
  Print #1, Join(lines, vbCrLf)
  Close #1
    
End Sub

Sub MSAccessAutoNumbers2PGDump(out_file As String, Optional fileAppend As Boolean = False)
  '-- Write PostgreSQL dump file to create sequences for Microsoft Access database
  Dim sequence_name As String
  Dim lines() As String
  Dim db As Database
   
  Set db = CurrentDb
  
  For Each tdf In db.TableDefs
    '-- No system tables, linked tables or temp tables
    If Left(tdf.name, 4) <> "Msys" And Left(tdf.name, 1) <> "~" And tdf.Connect = "" Then
      For Each field In tdf.fields
        If field.Type = dbLong Then
          If (field.Attributes And dbAutoIncrField) <> 0 Then
            '-- Found auto number field
            sequence_name = LCase(tdf.name & "_" & field.name & "_seq")
            '-- Create a PostgreSQL sequence starting at the max value of the current column + 1
            Call pushStr(lines, "CREATE SEQUENCE " & strQuote(sequence_name, Chr$(34)) & " START " & (Nz(DMax(field.name, tdf.name), 0) + 1) & "; ")
            ' -- Set column default to next value of sequence
            Call pushStr(lines, "ALTER TABLE " & strQuote(tdf.name, Chr$(34)) & " ALTER COLUMN " & strQuote(field.name, Chr$(34)) & " SET DEFAULT nextval(" & strQuote(sequence_name, Chr$(39)) & "::regclass);")
          End If
        End If
      Next
    End If
  Next
  
  '-- Write out to file
  If fileAppend Then
    Open out_file For Append As #1
  Else
    Open out_file For Output As #1
  End If
  Print #1, Join(lines, vbCrLf)
  Close #1
  
End Sub

Sub MSAccessRecords2PGDump(out_file As String, Optional fileAppend As Boolean = False)
  '-- Write PostgreSQL dump file to load data for Microsoft Access database records
  Dim lines() As String
  Dim cols() As String
  Dim cols_quoted() As String
  Dim value As String
  Dim values() As String
  Dim db As Database
  Dim rec As Recordset
  
  Set db = CurrentDb
  
  '-- Write or append to file
  If fileAppend Then
    Open out_file For Append As #1
  Else
    Open out_file For Output As #1
  End If
      
  For Each tdf In db.TableDefs
    '-- No system tables, linked tables or temp tables
    If Left(tdf.name, 4) <> "Msys" And Left(tdf.name, 1) <> "~" And tdf.Connect = "" Then
      Set rec = db.OpenRecordset(tdf.name, dbOpenSnapshot)
      
      Print #1, "BEGIN;";
      Print #1, Chr$(10);
      Print #1, "DELETE FROM " & strQuote(tdf.name, Chr$(34)) & " ;";
      Print #1, Chr$(10);
              
      '-- Columns
      Erase cols
      Erase cols_quoted
      For Each field In tdf.fields
        Call pushStr(cols, field.name)
        Call pushStr(cols_quoted, strQuote(field.name, Chr$(34)))
      Next
      
      '-- PostgreSQL Copy statement
      Print #1, "COPY " & strQuote(tdf.name, Chr$(34)) & " (" & Join(cols_quoted, ",") & ")" & " FROM STDIN;";
      Print #1, Chr$(10);
                  
      '-- Copy data
      If rec.RecordCount > 0 Then
        Do
          Erase values
          For i = 0 To UBound(cols)
            If IsNull(rec(cols(i))) Then
              Call pushStr(values, "\N")
            Else
              value = rec(cols(i))
              value = replace(value, "\", "\\")
              value = replace(value, Chr$(8), "\b")
              value = replace(value, Chr$(12), "\f")
              value = replace(value, Chr$(13), "\r")
              value = replace(value, Chr$(10), "\n")
              value = replace(value, Chr$(9), "\t")
              value = replace(value, Chr$(11), "\v")
              Call pushStr(values, value)
            End If
          Next
          Print #1, Join(values, Chr$(9));
          Print #1, Chr$(10);
  
          rec.MoveNext
        Loop Until rec.EOF
      End If
      
      Print #1, "\.";
      Print #1, Chr$(10);
      Print #1, "COMMIT;";
      Print #1, Chr$(10);
    End If
  Next
  
  Close #1
End Sub




