/*
AnalRuns, key RunNo, RunID
AnalRunSeq, key SeqNo, reference AnalRuns through RunID, reference Tests through TestCode
AnalRunSeqResult, key SeqNo and Analyte, reference AnalRunSeq through SeqNo
*/
IF OBJECT_ID('dbo.calcResult', 'P') IS NOT NULL
	DROP PROC dbo.calcResult;
GO

CREATE PROC dbo.calcResult
  @wsid varchar(50)
AS

IF @@TRANCOUNT<>0 ROLLBACK;
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY;
  BEGIN TRAN;   -- put everything in transaction, so can rollback if error
  
  -- Updates the average values for PCB type analytes.
  WITH tmpPCBavg AS
  (
    SELECT a.SeqNo, a.AnalyteType, Avg(a.RawVal) AS AvgOfRawVal, t.PeakRef, t.TestCode  
    FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo=a.SeqNo 
      INNER JOIN TestCodeLimit t ON (d.TestCode = t.TestCode) AND (a.Analyte = t.Analyte) 
    WHERE a.RawVal Is Not Null AND a.AnalyteType='P' AND t.PeakRef Is Not Null AND t.PeakRef<>0 AND
      d.WSID=@wsid
    GROUP BY a.SeqNo, a.AnalyteType, t.PeakRef, t.TestCode
  )
  UPDATE a
  SET a.RawVal = p.AvgOfRawVal  
  FROM tmpPCBavg p INNER JOIN AnalRunSeqResult a ON (a.SeqNo = p.SeqNo) 
    INNER JOIN TestCodeLimit t ON (p.TestCode = t.TestCode) 
      AND (p.PeakRef = t.PeakRef) 
      AND (a.Analyte = t.Analyte)
  WHERE a.RawVal Is Null AND a.AnalyteType='A' AND t.PeakRef<>0;

  -- Clears out the old quals and resets RPD and REC = 0
  -- Qual can't be null, it may cause Access report error later on
  UPDATE a
  SET Qual = NULL, CalcVal = 0, RPD = 0, REC = 0
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo 
  WHERE d.WSID=@wsid;

  -- Appends '-dry' to units when pmoist calc is done.
  UPDATE a
  SET Units = Units+'-dry'
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.Units Not Like '%dry' AND d.Pmoist Is Not Null And d.Pmoist<>0 AND d.NoMoistCorrect=0 AND
    d.WSID=@wsid;

  -- Updates PMoist = 0 when isNull(Pmoist); cannot be zero and do calcs
  UPDATE DataEntrySeq
  SET Pmoist = 0
  WHERE Pmoist Is Null AND WSID=@wsid;

  -- Updates MDL and PQL for factors and dilutions w/ Pmoist correction
  WITH calc AS
  (
    SELECT ((d.BackFracFac*a.rawMDL*d.DF*d.PrepFac/d.OriginalFac)/(100-d.Pmoist))*100/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END) AS MDL, 
      ((d.BackFracFac*a.rawPQL*d.DF*d.PrepFac/d.OriginalFac)/(100-d.Pmoist))*100/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END) AS PQL,
      a.SeqNo, a.Analyte, d.SigFigsMDL
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
    WHERE d.WSID=@wsid
  )
  UPDATE a
  SET MDL = dbo.sigfig(c.MDL, c.SigFigsMDL), PQL = dbo.sigfig(c.PQL, c.SigFigsMDL)
  FROM AnalRunSeqResult a INNER JOIN calc c ON a.SeqNo=c.SeqNo AND a.analyte = c.analyte;

  -- Calculates CalcVal, FinalVal
  WITH calc AS
  (
    SELECT d.BackFracFac*(a.rawVal*d.DF*d.PrepFac/(100-d.Pmoist)*100*d.ConvFac/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END)-(CASE WHEN d.BCMethod=2 OR (d.BCMethod=1 AND a.BlkRefVal>0) THEN a.BlkRefVal ELSE 0 END)+(CASE WHEN d.BackRef=0 THEN 0 ELSE a.BackRefVal END)) AS val, 
      a.SeqNo, a.Analyte, d.SigFigs
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
    WHERE a.RawVal Is Not Null AND d.WSID=@wsid
  )
  UPDATE a
  SET CalcVal = c.val, FinalVal = dbo.sigfig(c.val, c.SigFigs)
  FROM AnalRunSeqResult a INNER JOIN calc c ON a.SeqNo=c.SeqNo AND a.analyte = c.analyte;

  WITH calc AS
  (
    SELECT d.BackFracFac*(a.rawVal*d.DF*d.PrepFac*(al.molwt/24.45)/(100-d.Pmoist)*100*d.ConvFac/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END)-(CASE WHEN d.BCMethod=2 OR (d.BCMethod=1 AND a.BlkRefVal>0) THEN a.BlkRefVal ELSE 0 END)+(CASE WHEN d.BackRef=0 THEN 0 ELSE a.BackRefVal END)) AS val, 
      a.SeqNo, a.Analyte, d.SigFigs
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
      INNER JOIN Analytes al ON a.Analyte = al.Analyte 
    WHERE a.RawVal Is Not Null AND d.TestCode Like 'TO_15%G%' AND d.WSID=@wsid
  )
  UPDATE a
  SET CalcVal = c.val, FinalVal = dbo.sigfig(c.val, c.SigFigs)
  FROM AnalRunSeqResult a INNER JOIN calc c ON a.SeqNo=c.SeqNo AND a.analyte = c.analyte;

  UPDATE a
  SET FinalVal = RawVal
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.RawVal Is Null AND d.WSID=@wsid;

  UPDATE a
  SET CalcVal = FinalVal
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.RawVal Is Null AND d.WSID=@wsid;

  -- Updates CalcVal, FinalVal = 0 when CalcVal < MDL
  UPDATE a
  SET FinalVal = 0, CalcVal = 0
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo 
    INNER JOIN SampleTypes s ON d.SampType = s.SampType 
  WHERE a.CalcVal<a.MDL AND s.CAL=0 AND d.WSID=@wsid;

  -- Updates BLKrefVal and BackrefVal for analytes with associated Blank hit
  WITH tmpBLKRef AS
  (
    SELECT d.SeqNo, a.Analyte, a.CalcVal AS RefVal
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.BLKref
    WHERE a.AnalyteType IN ('A', 'C', 'T') AND d.WSID=@wsid
  )
  UPDATE a
  SET BLKrefval = t.RefVal
  FROM AnalRunSeqResult a INNER JOIN tmpBLKRef t ON a.SeqNo=t.SeqNo AND a.analyte = t.analyte;

  WITH tmpBACKRef AS
  (
    SELECT d.SeqNo, a.Analyte, a.CalcVal AS RefVal
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.Backref
    WHERE a.AnalyteType IN ('A', 'C', 'T') AND d.WSID=@wsid
  )
  UPDATE a
  SET Backrefval = t.RefVal
  FROM AnalRunSeqResult a INNER JOIN tmpBACKRef t ON a.SeqNo=t.SeqNo AND a.analyte = t.analyte;

  -- Recalculates for BLKref and BACKref
  WITH calc AS
  (
    SELECT d.BackFracFac*(a.rawVal*d.DF*d.PrepFac/(100-d.Pmoist)*100*d.ConvFac/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END)-(CASE WHEN d.BCMethod=2 OR (d.BCMethod=1 AND a.BlkRefVal>0) THEN a.BlkRefVal ELSE 0 END)+(CASE WHEN d.BackRef=0 THEN 0 ELSE a.BackRefVal END)) AS val, 
      a.SeqNo, a.Analyte, d.SigFigs
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
    WHERE a.RawVal Is Not Null AND d.WSID=@wsid
  )
  UPDATE a
  SET CalcVal = c.val, FinalVal = dbo.sigfig(c.val, c.SigFigs)
  FROM AnalRunSeqResult a INNER JOIN calc c ON a.SeqNo=c.SeqNo AND a.analyte = c.analyte;

  WITH calc AS
  (
    SELECT d.BackFracFac*(a.rawVal*d.DF*d.PrepFac*(al.molwt/24.45)/(100-d.Pmoist)*100*d.ConvFac/(CASE WHEN a.RCStat = 0 THEN 1 ELSE a.HistRec END)-(CASE WHEN d.BCMethod=2 OR (d.BCMethod=1 AND a.BlkRefVal>0) THEN a.BlkRefVal ELSE 0 END)+(CASE WHEN d.BackRef=0 THEN 0 ELSE a.BackRefVal END)) AS val,
      a.SeqNo, a.Analyte, d.SigFigs
    FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
      INNER JOIN Analytes al ON a.Analyte = al.Analyte 
    WHERE a.RawVal Is Not Null AND d.TestCode Like 'TO_15%G%' AND d.WSID=@wsid
  )
  UPDATE a
  SET CalcVal = c.val, FinalVal = dbo.sigfig(c.val, c.SigFigs)
  FROM AnalRunSeqResult a INNER JOIN calc c ON a.SeqNo=c.SeqNo AND a.analyte = c.analyte;

  UPDATE a
  SET FinalVal = 0, CalcVal = 0
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
    INNER JOIN SampleTypes s ON d.SampType = s.SampType 
  WHERE a.CalcVal<MDL AND s.CAL=0 AND d.WSID=@wsid;

  UPDATE a
  SET SpkVal = a.RawSpkVal*d.SpkFac/(100-d.Pmoist)*100
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.RawSpkVal<>0 AND d.SampType<>'LCS1' AND d.WSID=@wsid;

  UPDATE a
  SET SpkVal = RawSpkVal
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.RawSpkVal<>0 AND d.SampType='LCS1' AND d.WSID=@wsid;

  UPDATE a
  SET FinalVal = 0, CalcVal = 0
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.FinalVal Is Null AND d.WSID=@wsid;

  UPDATE a
  SET SpkVal = a.SpkVal*d.DF
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE (d.Dept='AT' OR d.Dept Like '%VOA') AND a.RawSpkVal<>0 AND d.WSID=@wsid;

  -- Updates SpkVal = [SpkVal]*DF for PDS Samples to handle sample dilutions
  UPDATE a
  SET SpkVal = a.SpkVal*d.DF
  FROM AnalRunSeqResult a INNER JOIN DataEntrySeq d ON a.SeqNo=d.SeqNo
  WHERE a.RawSpkVal<>0 AND d.SampType='PDS' AND d.Dept<>'ME' AND d.WSID=@wsid;

  -- Updates SPKrefVal for analytes linked by spkref to a sample or blank
  SELECT d.SeqNo, a.Analyte, a.CalcVal AS RefVal INTO #tmpSPKRef
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SPKref = a.SeqNo
  WHERE a.AnalyteType in ('A', 'C', 'M') AND d.WSID=@wsid;

  UPDATE t
  SET RefVal=a.CalcVal/2
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SPKref = a.SeqNo
    INNER JOIN #tmpSPKRef t ON d.SeqNo = t.SeqNo 
  WHERE d.TestCode Like '%COD%';

  --Updates RPDrefVal for all analytes in sample with RPDref link
  SELECT d.SeqNo, a.Analyte, a.CalcVal AS RefVal INTO #tmpRPDRef
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.RPDref = a.SeqNo
  WHERE a.AnalyteType IN ('A', 'M') AND d.WSID=@wsid;

  UPDATE a
  SET SPKrefval = t.RefVal
  FROM #tmpSPKRef t INNER JOIN AnalRunSeqResult a ON (t.Analyte = a.Analyte) AND (t.SeqNo = a.SeqNo);

  UPDATE a
  SET RPDrefval = t.RefVal
  FROM #tmpRPDRef t INNER JOIN AnalRunSeqResult a ON (t.Analyte = a.Analyte) AND (t.SeqNo = a.SeqNo);

  -- Calculates %Rec for all analytes having a SpkVal <> 0
  UPDATE a 
  SET REC = (a.CalcVal-a.SPKrefval)/a.SpkVal
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE a.SpkVal<>0 AND d.WSID=@wsid;

  -- Calculates the RPD's
  UPDATE a
  SET RPD = Abs(a.CalcVal-a.RPDrefval)/(Abs(a.CalcVal+a.RPDrefval)/2)
  FROM AnalRunSeqResult a INNER JOIN #tmpRPDRef t ON (a.Analyte = t.Analyte) AND (a.SeqNo = t.SeqNo) 
  WHERE Abs(a.CalcVal+a.RPDrefVal)>0 AND a.CalcVal>=a.PQL AND a.AnalyteType in ('A', 'M');

  -- Updates qual to B when BLKrefVal > PQL
  UPDATE a
  SET Qual = 'B'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE a.BLKrefval<>0 And a.BLKrefval>a.PQL AND a.CalcVal>a.PQL AND d.WSID=@wsid;
  
  UPDATE a
  SET Qual = 'B'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE a.BLKrefval<>0 And a.BLKrefval>a.rawPQL AND a.CalcVal>a.PQL AND d.RunID LIKE 'voa%' AND d.WSID=@wsid;

  -- Updates Qual to J and  removes 1 SigFig from FinalVal when MDL < CalcVal  < PQL
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'J', FinalVal = dbo.sigfig(a.CalcVal, d.SigFigs)
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo 
    INNER JOIN SampleTypes s ON d.SampType = s.SampType 
  WHERE a.FinalVal<>0 AND a.CalcVal>=a.MDL And a.CalcVal<a.PQL AND s.CAL=0 AND d.WSID=@wsid;

  -- Updates Qual to Qual + S when REC outside limits
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'S'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo 
  WHERE (a.REC*100<a.LowLimit Or a.REC*100>a.HighLimit) AND a.SpkVal<>0 AND a.HighLimit<>0 AND d.WSID=@wsid;

  -- Updates Qual to Qual + R when RPDlimit is exceeded
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'R'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo 
  WHERE a.RPD*100>a.RPDlimit AND a.RPDlimit<>0 AND d.WSID=@wsid;

  -- Updates Qual to Qual + E when RawVal > UQL and UQL <> 0
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'E'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo 
  WHERE a.UQL<>0 AND a.RawVal>a.UQL AND d.WSID=@wsid;

  -- Updates all TIC qualifiers to be Qual & J
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'J'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE a.AnalyteType='T' AND d.WSID=@wsid;

  -- Updates TIC qualifiers to be Qual & N for TIC's that have a CAS
  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + 'N'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
    INNER JOIN AnalRunSeqResultSub asub ON (a.Analyte = asub.Analyte) AND (a.SeqNo = asub.SeqNo) 
  WHERE a.AnalyteType='T' AND asub.CAS Is Not Null AND d.WSID=@wsid;

  -- Add GalpRecord to Galp table if finalVal has been changed
  INSERT INTO GALPrecord ( PreRawVal, PreFinalVal, PostRawVal, PostFinalVal, Analyte, SampID, WorkOrder, SeqNo, Updated, UpdateBy, GalpComment, TestCode )
  SELECT a.PreRawVal, a.PreFinalVal, a.RawVal, a.FinalVal, a.Analyte, d.SampID, d.WorkOrder, d.SeqNo, getdate() AS Expr2, d.UpdateBy AS Expr3, a.GalpComment, d.TestCode
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE a.FinalVal<>a.PreFinalVal AND d.WSID=@wsid; 

  -- Updates completed in SampleTest to the AnalDate
  UPDATE st
  SET Completed = d.AnalDate, CompletedBy = d.UpdateBy, PartComp = 0, Updated = GETDATE(), UpdateBy = d.UpdateBy
  FROM DataEntrySeq d INNER JOIN SampleTest st ON d.SampTestNo = st.SampTestNo AND d.WSID=@wsid;

  -- Updates completed in SampleTest to Null and PC to -1 for missing analytes
  UPDATE st
  SET Completed = Null, CompletedBy = Null, PartComp = -1, Updated = GETDATE(), UpdateBy = d.UpdateBy
  FROM DataEntrySeq d INNER JOIN SampleTest st ON d.SampTestNo = st.SampTestNo
    INNER JOIN MissingAnalytes m ON m.SampTestNo = d.SampTestNo 
  WHERE d.WSID=@wsid;

  -- Sets PreRaw and PreFinal for next save operation
  UPDATE a
  SET PreRawVal = a.RawVal, PreFinalVal = a.FinalVal
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE d.WSID=@wsid;

  -- Calculates the Column percent difference between recoveries on the front and rear columns.
  UPDATE a
  SET Col_Diff = CASE WHEN (CASE WHEN a.CalcVal<a1.CalcVal THEN a.CalcVal ELSE a1.CalcVal END)=0 THEN -1 ELSE (Abs(a.CalcVal-a1.CalcVal)/CASE WHEN a.CalcVal<a1.CalcVal THEN a.CalcVal ELSE a1.CalcVal END) END
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo 
    INNER JOIN AnalRunSeqResult a1 ON (d.Col2Ref = a1.SeqNo) AND (a.Analyte = a1.Analyte) 
  WHERE a.CalcVal<>0 AND d.WSID=@wsid;

  UPDATE a 
  SET Qual = ISNULL(a.Qual, '') + 'HT'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE d.LotNo='HT' AND d.WSID=@wsid;

  -- Update Qual to H for samples exceeding analytical holdtime
  UPDATE a 
  SET Qual = ISNULL(a.Qual, '') + 'H'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
  WHERE d.LotNo = 'H' AND d.WSID=@wsid;

  UPDATE a
  SET Qual = ISNULL(a.Qual, '') + '*'
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
    INNER JOIN Tests t ON d.TestCode = t.TestCode 
  WHERE a.AnalyteType='A' AND d.WSID=@wsid;

  UPDATE a
  SET Qual = CASE WHEN REPLACE(Qual, '*', '')='' THEN NULL ELSE REPLACE(Qual, '*', '') END
  FROM DataEntrySeq d INNER JOIN AnalRunSeqResult a ON d.SeqNo = a.SeqNo
    INNER JOIN Tests t ON d.TestCode = t.TestCode
    INNER JOIN Analytes al ON (t.TestNo = al.TestNo) AND (al.Analyte = a.Analyte) 
  WHERE a.AnalyteType='A' AND al.Accredited=1 AND d.AnalDate>=al.AccredDate AND d.WSID=@wsid;
  
  COMMIT TRAN
END TRY
BEGIN CATCH
  IF @@TRANCOUNT<>0 ROLLBACK;
  
  IF OBJECT_ID('tempdb..#tmpRPDRef', 'U') IS NOT NULL
    DROP TABLE #tmpRPDRef;

  IF OBJECT_ID('tempdb..#tmpSPKRef', 'U') IS NOT NULL
    DROP TABLE #tmpSPKRef;

  DELETE FROM DataEntrySeq WHERE WSID=@wsid;
  
  DECLARE @ErrorMessage NVARCHAR(4000);
  DECLARE @ErrorSeverity INT;
  DECLARE @ErrorState INT;

  SELECT @ErrorMessage = ERROR_MESSAGE(),
         @ErrorSeverity = ERROR_SEVERITY(),
         @ErrorState = ERROR_STATE();

  -- Use RAISERROR inside the CATCH block to return 
  -- error information about the original error that 
  -- caused execution to jump to the CATCH block.
  RAISERROR (@ErrorMessage, -- Message text.
             @ErrorSeverity, -- Severity.
             @ErrorState -- State.
             ) WITH LOG;
END CATCH

IF OBJECT_ID('tempdb..#tmpRPDRef', 'U') IS NOT NULL
  DROP TABLE #tmpRPDRef;

IF OBJECT_ID('tempdb..#tmpSPKRef', 'U') IS NOT NULL
  DROP TABLE #tmpSPKRef;

DELETE FROM DataEntrySeq WHERE WSID=@wsid;

GO

GRANT EXECUTE ON dbo.calcResult TO Admins;
GO
GRANT EXECUTE ON dbo.calcResult TO Analysts;
GO
GRANT EXECUTE ON dbo.calcResult TO Mgmt;
GO