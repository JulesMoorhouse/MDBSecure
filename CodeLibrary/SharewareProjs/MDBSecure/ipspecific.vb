Imports ADOX
Friend Module ipspecific

    Friend Structure flamer
        Friend strSD As String
        Friend strLD As String
        Friend strED As String
        Friend strHS As String
        Friend strLTD As String
    End Structure

    Private Sub SetRealDates(ByRef pstrResource As flamer, Optional ByVal pbooSetHS As Boolean = False)
        'Local to UK
        With pstrResource
            'Console.WriteLine("SetRealDatesB4= strSD=" & .strSD & " strLD=" & .strLD & " strED=" & .strED & " strHS=" & .strHS & " strLTD=" & .strLTD)
            .strSD = CodeCoverDate("6" & UKDateStr((CDate(.strSD).AddDays(13))))
            .strLD = CodeCoverDate("3" & UKDateStr((CDate(.strLD).AddDays(-34))))
            .strED = CodeCoverDate("8" & UKDateStr((CDate(.strED).AddDays(-73))))
            If pbooSetHS = False Then
                .strHS = CodeCoverDate("9" & UKDateStr((CDate(.strHS).AddDays(25))))
            Else
                '.strHS = CodeCoverDate("9" & "2" & "6/0" & "1/20" & "00 00" & ":00:00")
                .strHS = CodeCoverDate("9" & "26/01/2000 00:00:00")
            End If
            .strLTD = CodeCoverDate("1" & UKDateStr((CDate(.strLTD).AddDays(-17))))
            'Console.WriteLine("SetRealDates= strSD=" & .strSD & " strLD=" & .strLD & " strED=" & .strED & " strHS=" & .strHS & " strLTD=" & .strLTD)
        End With

    End Sub
    Private Sub IdeasPadSpecific(ByVal ADCat As ADOX.Catalog)
        Dim tblADOXTopics As ADOX.Table = New ADOX.Table()
        Dim tblADOXTopicDetail As ADOX.Table = New ADOX.Table()
        Dim tblADOXTopicMatches As ADOX.Table = New ADOX.Table()

        '*' Add tables
        With tblADOXTopics
            .Name = "Topics"
            .Columns.Append("TopicCode", DataTypeEnum.adVarWChar, 10)
            .Columns.Append("Level", DataTypeEnum.adInteger)
            .Columns.Append("Title", DataTypeEnum.adVarWChar, 50)
            .Columns.Append("ParentTopicCode", DataTypeEnum.adVarWChar, 10)
            .Columns("ParentTopicCode").Attributes = ColumnAttributesEnum.adColNullable
            .Columns.Append("SeqNum", DataTypeEnum.adInteger)
            .Columns.Append("InUse", DataTypeEnum.adBoolean)
        End With
        ADCat.Tables.Append(tblADOXTopics)

        With tblADOXTopicDetail
            .Name = "TopicDetail"
            .Columns.Append("TopicCode", DataTypeEnum.adVarWChar, 10)
            .Columns.Append("Level", DataTypeEnum.adInteger)
            .Columns.Append("Detail", DataTypeEnum.adLongVarWChar)
        End With
        ADCat.Tables.Append(tblADOXTopicDetail)

        With tblADOXTopicMatches
            .Name = "TopicMatches"
            .Columns.Append("A", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("B", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("C", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("D", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("E", DataTypeEnum.adVarWChar, 25)
        End With
        ADCat.Tables.Append(tblADOXTopicMatches)


    End Sub
    Private Sub IdeasPadSpecific2(ByVal Con As ADODB.Connection)

        Dim le As String = System.IO.Path.GetDirectoryName(System.Environment.GetFolderPath(Environment.SpecialFolder.System)) & "\explorer.exe"

        Dim lstrNewFlamer As New flamer()
        Dim lstrSQL As String

        'set values
        lstrNewFlamer.strSD = Date.Now
        lstrNewFlamer.strLD = Date.Now
        lstrNewFlamer.strED = System.IO.File.GetLastAccessTime(le)
        lstrNewFlamer.strHS = "01/01/2000 00:01:00"
        lstrNewFlamer.strLTD = Date.Now

        SetRealDates(lstrNewFlamer, True)

        lstrSQL = "Insert into TopicMatches (A,B,C,D,E) values ('" & _
                  lstrNewFlamer.strSD & "', '" & lstrNewFlamer.strLD & "', '" & lstrNewFlamer.strED & "', '" & _
                  lstrNewFlamer.strHS & "', '" & lstrNewFlamer.strLTD & "');"

        Con.Execute(lstrSQL)
    End Sub
    Private Function CodeCoverDate(ByVal pstrString As String) As String
        Dim lintArrInc As Integer
        Dim lstrM As String

        For lintArrInc = 1 To (pstrString).Length
            Select Case MidGet(pstrString, lintArrInc, 1)
                Case "0" : lstrM = "X"
                Case "1" : lstrM = "S"
                Case "2" : lstrM = "G"
                Case "3" : lstrM = "H"
                Case "4" : lstrM = "R"
                Case "5" : lstrM = "W"
                Case "6" : lstrM = "P"
                Case "7" : lstrM = "A"
                Case "8" : lstrM = "K"
                Case "9" : lstrM = "L"
                Case " " : lstrM = "D"
                Case "/" : lstrM = "B"
                Case ":" : lstrM = "O"
                Case "." : lstrM = "S"
                Case "-" : lstrM = "J"
            End Select
            CodeCoverDate = CodeCoverDate & lstrM
        Next lintArrInc

    End Function
    Private Function UKDateStr(ByVal pstrString As Date) As String

        'UKDateStr = (Format(pstrString, "dd/MM/yyyy HH:mm:ss"))
        UKDateStr = pstrString.ToString("dd/MM/yyyy HH:mm:ss")
        'Console.WriteLine("UKDateStr=" & UKDateStr)
    End Function
End Module
