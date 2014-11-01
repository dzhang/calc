' create temp tables
Public Function MakeDataTBL()
    If DLookup("Is_SQL", "Version") = -1 Then
        MsgBox "This option is not available in the SQLServer Version of Omega_ME", vbInformation
        Exit Function
    End If
    On Error Resume Next
    DoCmd.SetWarnings False
    Dim tdf As TableDef, Fld As Field, dbs As Database
    Set dbs = CurrentDb
             
    DoCmd.DeleteObject acTable, "tmpRunID"
    DoCmd.RunSQL ("SELECT AnalRuns.RunID INTO tmpRunID FROM AnalRuns WHERE (((AnalRuns.RunID) Is Null));")
    DoCmd.RunSQL ("CREATE INDEX RunID ON tmpRunID (RunID) WITH PRIMARY DISALLOW NULL;")
    DoCmd.RunSQL ("Alter Table tmpRunID ADD COLUMN WSID Text(20);")
    DoCmd.RunSQL ("Alter Table tmpRunID ADD AnalDate Date;")
    DoCmd.RunSQL ("Alter Table tmpRunID ADD QAName Text(100);")
    DoCmd.RunSQL ("Alter Table tmpRunID ADD RunNo Long;")
    
    DoCmd.DeleteObject acTable, "tmpSampID"
    DoCmd.RunSQL ("SELECT Samples.SampID, Samples.ClientSampID INTO tmpSampID FROM Samples WHERE (((Samples.SampID) Is Null));")
    DoCmd.RunSQL ("Alter Table tmpSampID ADD COLUMN Fraction Text(1);")
    DoCmd.RunSQL ("Alter Table tmpSampID ADD COLUMN SampNum Text(50);")
    DoCmd.RunSQL ("CREATE INDEX SampID ON tmpSampID (SampID) WITH PRIMARY DISALLOW NULL;")
    
    ' create table tmpDataEntrySeq
    DoCmd.DeleteObject acTable, "tmpDataEntrySeq"
    DoCmd.RunSQL ("SELECT AnalRunSeq.* INTO tmpDataEntrySeq FROM AnalRunSeq WHERE AnalRunSeq.SeqNo = 0;")
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq DROP COLUMN SeqNo;"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN SeqNo Long;"
    DoCmd.RunSQL ("CREATE INDEX SeqNo ON tmpDataEntrySeq (SeqNO) WITH PRIMARY DISALLOW NULL;")
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN Dept Text(10);"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN WorkOrder Text(20);"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN WSID Text(30);"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN HoldingTime Single;"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN PrepHoldTime Single;"
    DoCmd.RunSQL "Alter Table tmpDataEntrySeq ADD COLUMN VTSR Single;"
    
    Set tdf = dbs.TableDefs("tmpDataEntrySeq")
    Set Fld = tdf.CreateField("WriteToNet", dbBoolean)
    Fld.DefaultValue = False
    tdf.Fields.Append Fld
    Set Fld = tdf.CreateField("NoMoistCorrect", dbBoolean)
    Fld.DefaultValue = False
    tdf.Fields.Append Fld
    Set Fld = tdf.CreateField("CalcSamp", dbBoolean)
    Fld.DefaultValue = False
    tdf.Fields.Append Fld
    Set Fld = tdf.CreateField("HTVTSR", dbBoolean)
    Fld.DefaultValue = False
    tdf.Fields.Append Fld
    dbs.TableDefs.Refresh
    Set dbs = Nothing
        
    DoCmd.DeleteObject acTable, "tmpToggleSeqNo"
    DoCmd.RunSQL ("SELECT tmpDataEntrySeq.SeqNo AS FormSeqNo, tmpDataEntrySeq.SampID, tmpDataEntrySeq.TestCode, tmpDataEntrySeq.SampType, tmpDataEntrySeq.CalcLoop, tmpDataEntrySeq.Validated, tmpDataEntrySeq.SampTestNo, tmpDataEntrySeq.WorkOrder INTO tmpToggleSeqNo FROM tmpDataEntrySeq;")
    DoCmd.RunSQL ("CREATE INDEX FormSeqNo ON tmpToggleSeqNo (FormSeqNO) WITH PRIMARY DISALLOW NULL;")
    DoCmd.RunSQL "ALTER TABLE tmpToggleSeqNo ADD WSID TExt(20);"
   
    DoCmd.DeleteObject acTable, "tmpDataInputHdrTbl"
    DoCmd.RunSQL "SELECT ImportSpec_HeaderNameQry.* INTO tmpDataInputHdrTbl FROM ImportSpec_HeaderNameQry;"
    DoCmd.RunSQL "ALTER TABLE tmpDataInputHdrTbl ADD SampNo counter;"
    
    DoCmd.DeleteObject acTable, "tmpDataInputRsltTbl"
    DoCmd.RunSQL "SELECT ImportSpec_FieldNameQry.* INTO tmpDataInputRsltTbl FROM ImportSpec_FieldNameQry;"
    DoCmd.RunSQL "ALTER TABLE tmpDataInputRsltTbl ADD COLUMN SampNo Long;"
    
    DoCmd.RunSQL ("INSERT INTO ImportSpecs ( ImportSpec, FileType, Made ) SELECT 'Omega Excel File' AS Expr1, 4 AS Expr2, Now() AS Expr3;")
    DoCmd.RunSQL ("INSERT INTO ImportSpecField ( ImportSpec, ImportFieldName ) SELECT tmpOmegaSpec.ImportSpec, tmpOmegaSpec.ImportFieldName FROM tmpOmegaSpec;")
    DoCmd.RunSQL ("INSERT INTO Labpersonnel ( Name, Initials ) SELECT 'ADMIN' AS Expr1, 'ADM' AS Expr2;")
    DoCmd.SetWarnings True
    MsgBox "Tables re-created", vbInformation
End Function


' form AnalRuns_subB
' populate tmpDataEntrySeq
Private Sub Data_Entry_Click()
    On Error GoTo DEerr
    If Not IsNull(Forms!AnalRuns!FormIndex) Then
        RunQryTable "AnalRunSeq", Forms!AnalRuns!FormIndex  
        DoCmd.OpenForm "AnalRunSeqData"
    End If
    If Forms!AnalRuns!IS_SQL = -1 Then
        Forms!AnalRunSeqData!FormIndex.RowSource = "SELECT tmpDataEntrySeq.SeqNo, tmpDataEntrySeq.SampID FROM tmpDataEntrySeq INNER JOIN WSID ON tmpDataEntrySeq.WSID = WSID.WSID ORDER BY tmpDataEntrySeq.SeqNo;"
    End If
    Forms!AnalRunSeqData!FormIndex.SetFocus
    Forms!AnalRunSeqData!FormIndex = Forms!AnalRunSeqData!FormIndex.ItemData(0)
    
DE_Exit:
    Exit Sub
    
DEerr:
    MsgBox Err.Description & " Check to see if MakeTables needs to be run", vbInformation
    Resume DE_Exit

End Sub

' form AnalRunSeqData
' left panel pick list data source is select seqno, sampid from tmpDataEntrySeq order by seqno
' calc samp button
Private Sub SaveSamp_Click()
    DoCmd.SetWarnings False
    If Forms!AnalRuns!IS_SQL = 0 Then
        DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.CalcSamp = 0;")
    Else
        DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN WSID ON tmpDataEntrySeq.WSID = WSID.WSID SET tmpDataEntrySeq.CalcSamp = 0;")
    End If
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.CalcSamp = -1 WHERE tmpDataEntrySeq.SeqNo = " & Me!FormIndex & ";")
'   Added to prevent prep/analysis dates before collection date
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE ((Len([collectiondate])<11) AND ((tmpDataEntrySeq.PrepDate)<[collectiondate]));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.PrepDate)<[collectiondate]) AND (Len([prepdate])>10));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.PrepDate)<([collectiondate]-0.75)) AND ((Len([prepdate]))<11));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((tmpDataEntrySeq.AnalDate)<[collectiondate]) AND ((Len([collectiondate]))<11));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.AnalDate)<[collectiondate]) AND (Len([Analdate])>10));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.AnalDate)<([collectiondate]-0.75)) AND ((Len([Analdate]))<11));")
    DoCmd.SetWarnings True
    If DLookup("Validated", "tmpDataEntrySeq", "SeqNo = " & Me!FormIndex) = -1 Then
        If Forms!AnalRuns!OKTOValidate = 0 Then
            MsgBox "Sample has been validated. Use QA Authority to recalculate."
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((tmpDataEntrySeq.SeqNo)= " & Me!FormIndex & "));")
            DoCmd.SetWarnings True
        Else
        RunSaveQueries  ' call queries to do calculation
        End If
    Else
    RunSaveQueries  ' call queries to do calculation
    End If
End Sub

' calc seq button
Private Sub SaveSeqButton_Click()
Dim rst As Recordset
    Set rst = CurrentDb.OpenRecordset("SELECT tmpDataEntrySeq.SampID, tmpDataEntrySeq.Validated FROM tmpDataEntrySeq WHERE (((tmpDataEntrySeq.Validated)=-1));")
    DoCmd.SetWarnings False
    If Forms!AnalRuns!IS_SQL = 0 Then
        DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.CalcSamp = -1;")
    Else
        DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN WSID ON tmpDataEntrySeq.WSID = WSID.WSID SET tmpDataEntrySeq.CalcSamp = -1;")
    End If
'   Added to prevent prep/analysis dates before collection date
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE ((Len([collectiondate])<11) AND ((tmpDataEntrySeq.PrepDate)<[collectiondate]));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.PrepDate)<[collectiondate]) AND (Len([prepdate])>10));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.PrepDate)<([collectiondate]-0.75)) AND ((Len([prepdate]))<11));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((tmpDataEntrySeq.AnalDate)<[collectiondate]) AND ((Len([collectiondate]))<11));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.AnalDate)<[collectiondate]) AND (Len([Analdate])>10));")
    DoCmd.RunSQL ("UPDATE tmpDataEntrySeq INNER JOIN Samples ON tmpDataEntrySeq.SampID = Samples.SampID SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((Len([collectiondate]))>10) AND ((tmpDataEntrySeq.AnalDate)<([collectiondate]-0.75)) AND ((Len([Analdate]))<11));")
    DoCmd.SetWarnings True
    If rst.RecordCount > 0 Then
        If Forms!AnalRuns!OKTOValidate = 0 Then
            If MsgBox("This run contains validated data.  Only non-validated data can be calculated without QA Authority. Continue?", vbYesNo) = vbYes Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.CalcSamp = 0 WHERE (((tmpDataEntrySeq.Validated)=-1));")
            DoCmd.SetWarnings True
            RunSaveQueries  ' call queries to do calculation
            End If
        Else
        RunSaveQueries  ' call queries to do calculation
        End If
    Else
    RunSaveQueries
    End If
    rst.Close
End Sub

Public Sub RunSaveQueries()
    On Error GoTo UpdateSeqErr
    If MsgBox("Run Sample(s) through CALC process?", vbQuestion + vbYesNo, "Run CALC Queue") = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("DELETE tmpMsgBox.* FROM tmpMsgBox;")
        DoCmd.RunSQL ("INSERT INTO tmpMsgBox (msg) Values ('Running SAMPLE(s) through calculation queries. The CALC Queue contains over 60 queries so please be patient.');")
        DoCmd.SetWarnings True
        DoCmd.OpenForm "MsgPopup", , , , , acDialog, "AnalRunResult"
    End If

Exitdata:
    Me.Refresh
    Exit Sub

UpdateSeqErr:
    Select Case Err
        Case 3197
            Resume
        Case 3186, 3260
            DoEvents
            DBEngine.Idle dbRefreshCache
            Resume
        Case Else
            MsgBox Err.Description
            Resume Exitdata
        End Select

End Sub

' Form MsgPopup, timer event
' 10/8/14, Dong zhang
' call stored procedure
Private Sub Form_Timer()
    If Me.OpenArgs = "AnalRunResult" Then
        RunCalc   ' call stored procedure to do calculation
    Else
        RunQryTable (Me.OpenArgs)
    End If
    DoCmd.Close acForm, "MsgPopUp"
End Sub



' module BaseRunFunction
Public Function RunQryTable(TBLName As String, Optional MyPMVal As Variant)
    On Error GoTo NeedPM
    Dim rst As Recordset, StartVal As Single, SQLstr As String, qdf As QueryDef
    Dim dbs As Database, tbl As TableDef
    Dim wrkODBC As Workspace, conPubs As Connection
    Set dbs = CurrentDb
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE tmpTime.* FROM tmpTime WHERE tmpTime.TBLName = '" & TBLName & "';")
    If DLookup("IS_SQL", "Version") = 0 Then	' Access
        SQLstr = "SELECT QRYTable_Qry.* FROM QRYTable_Qry WHERE ((Run_AQry = -1) AND (QrySet = '" & TBLName & "')) Order By RunOrder, QryName;"
        Set rst = dbs.OpenRecordset(SQLstr)
        'GoTo RunAsStr
        Do Until rst.EOF
            StartVal = Timer
            DoCmd.OpenQuery rst!QryName
            DoCmd.RunSQL ("INSERT INTO tmpTime (QryName, RunTime, TBLname) SELECT '" & rst!QryName & "' AS Expr1, " & Timer - StartVal & " AS Expr2, '" & TBLName & "' AS Expr3;")
            rst.MoveNext
        Loop
        GoTo RQT_Exit
        ' runs qrys as strings
RunAsStr:
        Set qdf = CurrentDb.CreateQueryDef("")
        With qdf
            Do Until rst.EOF
                StartVal = Timer
                qdf.SQL = rst!AccessStr
                qdf.Execute
                DoCmd.RunSQL ("INSERT INTO tmpTime (QryName, RunTime, TBLname) SELECT '" & rst!QryName & "' AS Expr1, " & Timer - StartVal & " AS Expr2, '" & TBLName & "' AS Expr3;")
                rst.MoveNext
            Loop
        End With
    Else	' SQL
        SQLstr = "SELECT QRYTable_Qry.* FROM QRYTable_Qry WHERE ((Run_SQry = -1) AND (QrySet = '" & TBLName & "')) Order By RunOrder, QryName;"
        Set rst = dbs.OpenRecordset(SQLstr)
        Dim MyDSN As String, MyServer As String, MySQLDB As String
        MyDSN = DLookup("SQL_DSN", "Version")
        MyServer = DLookup("SQL_Name", "Version")
        MySQLDB = DLookup("SQL_DB", "Version")
'        Set wrkODBC = CreateWorkspace("tmpWS", "sa", "", dbUseODBC)
        Set conPubs = wrkODBC.OpenConnection("tmpConnect", , , "ODBC;DSN=" & MyDSN & ";SERVER=" & MyServer & ";DATABASE=" & MySQLDB & ";QueryLogFile=Yes")
        Set qdf = conPubs.CreateQueryDef("")
        With qdf
            Do Until rst.EOF
                StartVal = Timer
                If rst!RunAsLocal = -1 Then
                    DoCmd.OpenQuery rst!QryName
                Else
                    qdf.SQL = rst!SQLstr
                    qdf.Execute
                End If
                DoCmd.RunSQL ("INSERT INTO tmpTime (QryName, RunTime, TBLname) SELECT '" & rst!QryName & "' AS Expr1, " & Timer - StartVal & " AS Expr2, '" & TBLName & "' AS Expr3;")
                rst.MoveNext
            Loop
        End With
    End If
    rst.Close
    
RQT_Exit:
    DoCmd.SetWarnings True
    Exit Function
    
NeedPM:
    qdf.Parameters(0) = MyPMVal
    qdf.Execute
    Resume Next
End Function

' 10/8/2014 Dong zhang
' call stored procedure to do calculation, replace qryTable_qry query set AnalRunResult
Public Function RunCalc()
  On Error GoTo CalcErr
  
  DoCmd.SetWarnings False
  Dim WSID As String
  WSID = GetComputer  ' get computer name
  DoCmd.RunSQL ("DELETE WSID.* FROM WSID;")
  DoCmd.RunSQL ("INSERT INTO WSID (WSID) Values ('" & WSID & "');")
  
  ' get seqno to be calculated, used in stored procedure
  DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET CalcLoop = -1, Updated = Now(), UpdateBy = CurrentUser(), WriteToNet = -1 WHERE (((tmpDataEntrySeq.CalcSamp)=-1));")
  DoCmd.RunSQL ("INSERT INTO DataEntrySeq (SeqNo, WorkOrder, SampID, SampType, SampTestNo, TestCode, AnalDate, UpdateBy, Pmoist, NoMoistCorrect, BackFracFac, DF, PrepFac, OriginalFac, SigFigsMDL, ConvFac, BCMethod, BLKref, BackRef, RPDref, SigFigs, SpkFac, Dept, SPKref, Col2Ref, LotNo, WSID) " & _
    "SELECT SeqNo, WorkOrder, SampID, SampType, SampTestNo, TestCode, AnalDate, UpdateBy, Pmoist, NoMoistCorrect, BackFracFac, DF, PrepFac, OriginalFac, SigFigsMDL, ConvFac, BCMethod, BLKref, BackRef, RPDref, SigFigs, SpkFac, Dept, SPKref, Col2Ref, LotNo, WSID.WSID FROM tmpDataEntrySeq, WSID " & _
    "WHERE tmpDataEntrySeq.CalcSamp=-1;")
  
  Dim db As DAO.Database
  Dim qdf As DAO.QueryDef
  Set db = CurrentDb
  Set qdf = db.QueryDefs("z_callsp")
  ' Set the connection string
  Dim MyDSN As String, MyServer As String, MySQLDB As String
  MyDSN = DLookup("SQL_DSN", "Version")
  MyServer = DLookup("SQL_Name", "Version")
  MySQLDB = DLookup("SQL_DB", "Version")
  'qdf.Connect = "ODBC;DSN=" & MyDSN & ";SERVER=" & MyServer & ";DATABASE=" & MySQLDB & ";QueryLogFile=Yes"
  qdf.Connect = "ODBC;Driver={SQL Server};Server=" & MyServer & ";Database=" & MySQLDB & ";Trusted_Connection=Yes"
  'qdf.Connect = "ODBC;Driver={SQL Native Client};" & "Server=" & MyServer & ";" & "Database=" & MySQLDB & ";" & "Trusted_Connection=Yes"

  qdf.ReturnsRecords = False
  qdf.SQL = "EXEC dbo.calcResult '" & WSID & "'"
  qdf.Execute
  
  DoCmd.RunSQL ("UPDATE tmpDataEntrySeq SET tmpDataEntrySeq.BCStat = -1 WHERE (((tmpDataEntrySeq.CalcSamp)=-1) AND ((tmpDataEntrySeq.BCMethod)<>0));")
  DoCmd.SetWarnings True
  Exit Function
CalcErr:
  DoCmd.SetWarnings True
  MsgBox Err.Description, vbInformation
End Function





