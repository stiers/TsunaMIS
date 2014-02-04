VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Display(List As ListBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT post_title FROM tbl_posts WHERE post_type = 'wpa_vendors'"
    
    Records.Open query, Connect
    
    List.Clear
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        Do Until Records.EOF Or Records.BOF
            With List
                .AddItem Records!post_title
            End With
        Records.MoveNext
        Loop
    End If
End Function

Public Function Dropdown(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT post_title FROM tbl_posts WHERE post_type = 'wpa_vendors'"
    
    Records.Open query, Connect
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        
        If Not Records.EOF Then
            For ComboData = 0 To Records.RecordCount - 1
                Combo.AddItem Records!post_title
                Records.MoveNext
            Next ComboData
        End If
    End If
    
End Function

Public Function AddNew(Post As String)
    query = "INSERT INTO tbl_posts(ID, post_author, post_date, post_date_gmt, post_content, post_title, post_excerpt, post_status, comment_status, ping_status, post_password, post_name, to_ping, pinged, post_modified, post_modified_gmt, post_content_filtered, post_parent, guid, menu_order, post_type, post_mime_type, comment_count) " & _
            "VALUES ('', '1', '', '', '', '" & Post & "', '', 'published', 'closed', 'closed', '', '" & LCase(Post) & "', '', '', '', '', '', '0', 'http://localhost/tsuna/?post_type=wpa_vendors', '', 'wpa_vendors', '', '')"
            
    Connect.Execute query
End Function
