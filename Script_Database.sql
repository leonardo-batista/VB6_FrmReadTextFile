USE [BUSINESS]
GO

/****** Object:  Table [dbo].[TB_Fichier]    Script Date: 6/15/2019 20:03:52 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TB_Fichier](
	[ID_FICHIER] [bigint] IDENTITY(1,1) NOT NULL,
	[NOM_FICHIER] [varchar](100) NOT NULL,
	[DT_MISE_A_JOUR] [datetime] NULL,
	[DT_CHARGE_FICHIER] [datetime] NOT NULL,
	[CHARGE_PAR] [varchar](50) NOT NULL,
	[TOTAL_LIGNES] [int] NOT NULL,
	[TOTAL_CLIENTS] [int] NOT NULL,
 CONSTRAINT [PK_TB_Fichier] PRIMARY KEY CLUSTERED 
(
	[ID_FICHIER] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TB_Fichier] ADD  CONSTRAINT [DF_TB_Fichier_DT_CHARGE_FICHIER]  DEFAULT (getdate()) FOR [DT_CHARGE_FICHIER]
GO


CREATE TABLE [dbo].[TB_Client](
	[ID_CLIENT] [bigint] IDENTITY(1,1) NOT NULL,
	[PRENOM] [varchar](30) NOT NULL,
	[NOM] [varchar](30) NOT NULL,
	[DATA_NAISSANCE] [date] NOT NULL,
	[EMAIL] [varchar](50) NULL,
	[NAS] [varchar](9) NOT NULL,
	[TELEPHONE1] [varchar](10) NOT NULL,
	[TELEPHONE2] [varchar](10) NULL,
	[CODE_POSTAL] [varchar](6) NOT NULL,
	[NUMERO] [varchar](5) NOT NULL,
	[COMPLEMENT] [varchar](20) NULL,
	[ADRESSE] [varchar](60) NOT NULL,
	[VILLE] [varchar](50) NOT NULL,
	[PROVINCE] [char](2) NOT NULL,
	[DT_CREATION] [datetime] NOT NULL,
	[CREATE_PAR] [varchar](10) NOT NULL,
	[DT_MISE_A_JOUR] [datetime] NULL,
	[MISE_A_JOUR_PAR] [varchar](10) NULL,
	[ACTIVE] [bit] NOT NULL,
	[ID_FICHIER] [bigint] NOT NULL,
 CONSTRAINT [PK_TB_Client] PRIMARY KEY CLUSTERED 
(
	[ID_CLIENT] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ_NAS] UNIQUE NONCLUSTERED 
(
	[NAS] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TB_Client] ADD  CONSTRAINT [DF_TB_Client_Insere]  DEFAULT (getdate()) FOR [DT_CREATION]
GO

ALTER TABLE [dbo].[TB_Client] ADD  CONSTRAINT [DF_TB_Client_Active]  DEFAULT ((1)) FOR [ACTIVE]
GO

ALTER TABLE [dbo].[TB_Client]  WITH CHECK ADD  CONSTRAINT [FK_TB_Fichier] FOREIGN KEY([ID_FICHIER])
REFERENCES [dbo].[TB_Fichier] ([ID_FICHIER])
GO

ALTER TABLE [dbo].[TB_Client] CHECK CONSTRAINT [FK_TB_Fichier]
GO

