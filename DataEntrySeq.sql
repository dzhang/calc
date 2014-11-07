/*
table to hold data passed in from Access
used for results calculation
notice the WSID filed, each Access front has to pass in this value to prevent conflict on server
*/
IF OBJECT_ID('dbo.DataEntrySeq') IS NOT NULL
  DROP TABLE dbo.DataEntrySeq
GO

CREATE TABLE [dbo].[DataEntrySeq](
	[SeqNo] [int] NOT NULL,
	[SampID] [varchar](50) NOT NULL,
  [RunID] [varchar](50) NOT NULL,
	[TestCode] [varchar](50) NOT NULL,
	[SampType] [varchar](50) NOT NULL,
  [PrepFac] [float] NOT NULL,
	[OriginalFac] [float] NOT NULL,
	[SpkFac] [float] NOT NULL,
	[AnalDate] [datetime] NULL,
	[Pmoist] [real] NULL,
	[DF] [float] NULL,
	[SPKref] [int] NULL,
	[ConvFac] [float] NULL,
	[UpdateBy] [varchar](50) NULL,
	[SigFigs] [tinyint] NOT NULL,
	[Col2Ref] [int] NULL,
	[SigFigsMDL] [tinyint] NOT NULL,
	[BCMethod] [tinyint] NULL,
  [BLKref] [int] NULL,
	[Backref] [int] NULL,
  [RPDref] [int] NULL,
	[BackFracFac] [float] NULL,
	[SampTestNo] [int] NULL,
  WorkOrder varchar(20) null,
  NoMoistCorrect bit null,
  Dept varchar(10) NULL,
  [LotNo] [varchar](50) NULL,
  WSID varchar(50) not null
 CONSTRAINT [DataEntrySeq$PrimaryKey] PRIMARY KEY CLUSTERED 
(
	[SeqNo] ASC,
  WSID ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

GRANT SELECT ON DataEntrySeq TO Admins;
GO
GRANT INSERT ON DataEntrySeq TO Admins;
GO
GRANT UPDATE ON DataEntrySeq TO Admins;
GO
GRANT DELETE ON DataEntrySeq TO Admins;
GO
GRANT SELECT ON DataEntrySeq TO Analysts;
GO
GRANT INSERT ON DataEntrySeq TO Analysts;
GO
GRANT UPDATE ON DataEntrySeq TO Analysts;
GO
GRANT DELETE ON DataEntrySeq TO Analysts;
GO
GRANT SELECT ON DataEntrySeq TO Mgmt;
GO
GRANT INSERT ON DataEntrySeq TO Mgmt;
GO
GRANT UPDATE ON DataEntrySeq TO Mgmt;
GO
GRANT SELECT ON DataEntrySeq TO public;
GO