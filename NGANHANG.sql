USE [NGANHANG]
GO
/****** Object:  User [HTKN]    Script Date: 5/29/2022 9:47:50 PM ******/
CREATE USER [HTKN] FOR LOGIN [HTKN] WITH DEFAULT_SCHEMA=[dbo]
GO
/****** Object:  DatabaseRole [MSmerge_E1F62E7E7DB44631BC4C4561B3800760]    Script Date: 5/29/2022 9:47:50 PM ******/
CREATE ROLE [MSmerge_E1F62E7E7DB44631BC4C4561B3800760]
GO
/****** Object:  DatabaseRole [MSmerge_F227579622944553A50EA803F997A67D]    Script Date: 5/29/2022 9:47:50 PM ******/
CREATE ROLE [MSmerge_F227579622944553A50EA803F997A67D]
GO
/****** Object:  DatabaseRole [MSmerge_PAL_role]    Script Date: 5/29/2022 9:47:50 PM ******/
CREATE ROLE [MSmerge_PAL_role]
GO
ALTER ROLE [db_owner] ADD MEMBER [HTKN]
GO
ALTER ROLE [MSmerge_PAL_role] ADD MEMBER [MSmerge_E1F62E7E7DB44631BC4C4561B3800760]
GO
ALTER ROLE [MSmerge_PAL_role] ADD MEMBER [MSmerge_F227579622944553A50EA803F997A67D]
GO
/****** Object:  Schema [MSmerge_PAL_role]    Script Date: 5/29/2022 9:47:50 PM ******/
CREATE SCHEMA [MSmerge_PAL_role]
GO
/****** Object:  View [dbo].[Get_subcribers]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[Get_subcribers]
AS
SELECT  TENCN=PUBS.name, TENSERVER= subscriber_server
   FROM dbo.sysmergepublications PUBS,  dbo.sysmergesubscriptions SUBS
   WHERE PUBS.pubid= SUBS.PUBID  AND PUBS.publisher <> SUBS.subscriber_server
GO
/****** Object:  Table [dbo].[ChiNhanh]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ChiNhanh](
	[MACN] [nchar](10) NOT NULL,
	[TENCN] [nvarchar](100) NOT NULL,
	[DIACHI] [nvarchar](100) NOT NULL,
	[SoDT] [nvarchar](15) NOT NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_ChiNhanh] PRIMARY KEY CLUSTERED 
(
	[MACN] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UK_ChiNhanh] UNIQUE NONCLUSTERED 
(
	[TENCN] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GD_CHUYENTIEN]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GD_CHUYENTIEN](
	[MAGD] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[SOTK_CHUYEN] [nchar](9) NOT NULL,
	[NGAYGD] [datetime] NOT NULL,
	[SOTIEN] [money] NOT NULL,
	[SOTK_NHAN] [nchar](9) NOT NULL,
	[MANV] [nchar](10) NOT NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_GD_CHUYENTIEN] PRIMARY KEY CLUSTERED 
(
	[MAGD] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GD_GOIRUT]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GD_GOIRUT](
	[MAGD] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[SOTK] [nchar](9) NOT NULL,
	[LOAIGD] [nchar](2) NOT NULL,
	[NGAYGD] [datetime] NOT NULL,
	[SOTIEN] [money] NOT NULL,
	[MANV] [nchar](10) NOT NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_GD_GOIRUT] PRIMARY KEY CLUSTERED 
(
	[MAGD] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[KhachHang]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[KhachHang](
	[CMND] [nchar](10) NOT NULL,
	[HO] [nvarchar](50) NOT NULL,
	[TEN] [nvarchar](10) NOT NULL,
	[DIACHI] [nvarchar](100) NULL,
	[PHAI] [nvarchar](3) NULL,
	[NGAYCAP] [date] NOT NULL,
	[SODT] [nvarchar](15) NOT NULL,
	[MACN] [nchar](10) NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_KhachHang] PRIMARY KEY CLUSTERED 
(
	[CMND] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[NhanVien]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NhanVien](
	[MANV] [nchar](10) NOT NULL,
	[HO] [nvarchar](40) NOT NULL,
	[TEN] [nvarchar](10) NOT NULL,
	[DIACHI] [nvarchar](100) NULL,
	[PHAI] [nvarchar](3) NULL,
	[SODT] [nvarchar](15) NOT NULL,
	[MACN] [nchar](10) NULL,
	[TrangThaiXoa] [int] NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_NhanVien] PRIMARY KEY CLUSTERED 
(
	[MANV] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TaiKhoan]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TaiKhoan](
	[SOTK] [nchar](9) NOT NULL,
	[CMND] [nchar](10) NOT NULL,
	[SODU] [money] NULL,
	[MACN] [nchar](10) NULL,
	[NGAYMOTK] [datetime] NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_TaiKhoan] PRIMARY KEY CLUSTERED 
(
	[SOTK] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ChiNhanh] ADD  CONSTRAINT [MSmerge_df_rowguid_AF2BB187DE52482AAFC290DC20D2BD18]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN] ADD  CONSTRAINT [DF_GD_CHUYENTIEN_NGAYGD]  DEFAULT (getdate()) FOR [NGAYGD]
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN] ADD  CONSTRAINT [MSmerge_df_rowguid_210EA0587AB14E3DBAA6BBA008EF87FB]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[GD_GOIRUT] ADD  CONSTRAINT [DF_GD_GOIRUT_NGAYGD]  DEFAULT (getdate()) FOR [NGAYGD]
GO
ALTER TABLE [dbo].[GD_GOIRUT] ADD  CONSTRAINT [DF_GD_GOIRUT_SOTIEN]  DEFAULT ((100000)) FOR [SOTIEN]
GO
ALTER TABLE [dbo].[GD_GOIRUT] ADD  CONSTRAINT [MSmerge_df_rowguid_7899444BF0F441BCA25966AC3504C809]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[KhachHang] ADD  CONSTRAINT [MSmerge_df_rowguid_1E5B7A6F2DF04E5E9C06A96AF12327F2]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[NhanVien] ADD  CONSTRAINT [MSmerge_df_rowguid_CC54F4CE3AF5486CA62A065A158D741B]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[TaiKhoan] ADD  DEFAULT ((0)) FOR [SODU]
GO
ALTER TABLE [dbo].[TaiKhoan] ADD  CONSTRAINT [DF_TaiKhoan_NGAYMOTK]  DEFAULT (getdate()) FOR [NGAYMOTK]
GO
ALTER TABLE [dbo].[TaiKhoan] ADD  CONSTRAINT [MSmerge_df_rowguid_CC1D2B5FDA914223B874ECAC9DEBD900]  DEFAULT (newsequentialid()) FOR [rowguid]
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN]  WITH CHECK ADD  CONSTRAINT [FK_GD_CHUYENTIEN_NhanVien] FOREIGN KEY([MANV])
REFERENCES [dbo].[NhanVien] ([MANV])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN] CHECK CONSTRAINT [FK_GD_CHUYENTIEN_NhanVien]
GO
ALTER TABLE [dbo].[GD_GOIRUT]  WITH CHECK ADD  CONSTRAINT [FK_GD_GOIRUT_NhanVien] FOREIGN KEY([MANV])
REFERENCES [dbo].[NhanVien] ([MANV])
GO
ALTER TABLE [dbo].[GD_GOIRUT] CHECK CONSTRAINT [FK_GD_GOIRUT_NhanVien]
GO
ALTER TABLE [dbo].[KhachHang]  WITH CHECK ADD  CONSTRAINT [FK_KhachHang_ChiNhanh] FOREIGN KEY([MACN])
REFERENCES [dbo].[ChiNhanh] ([MACN])
GO
ALTER TABLE [dbo].[KhachHang] CHECK CONSTRAINT [FK_KhachHang_ChiNhanh]
GO
ALTER TABLE [dbo].[NhanVien]  WITH CHECK ADD  CONSTRAINT [FK_NhanVien_ChiNhanh] FOREIGN KEY([MACN])
REFERENCES [dbo].[ChiNhanh] ([MACN])
GO
ALTER TABLE [dbo].[NhanVien] CHECK CONSTRAINT [FK_NhanVien_ChiNhanh]
GO
ALTER TABLE [dbo].[TaiKhoan]  WITH CHECK ADD  CONSTRAINT [FK_TaiKhoan_ChiNhanh] FOREIGN KEY([MACN])
REFERENCES [dbo].[ChiNhanh] ([MACN])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[TaiKhoan] CHECK CONSTRAINT [FK_TaiKhoan_ChiNhanh]
GO
ALTER TABLE [dbo].[TaiKhoan]  WITH CHECK ADD  CONSTRAINT [FK_TaiKhoan_KhachHang] FOREIGN KEY([CMND])
REFERENCES [dbo].[KhachHang] ([CMND])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[TaiKhoan] CHECK CONSTRAINT [FK_TaiKhoan_KhachHang]
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN]  WITH CHECK ADD  CONSTRAINT [CK_GD_CHUYENTIEN] CHECK  (([SOTIEN]>(0)))
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN] CHECK CONSTRAINT [CK_GD_CHUYENTIEN]
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN]  WITH NOCHECK ADD  CONSTRAINT [repl_identity_range_D571EE49_692A_48B8_87DF_2CF8F2FB1DF3] CHECK NOT FOR REPLICATION (([MAGD]>=(1) AND [MAGD]<=(1001) OR [MAGD]>(1001) AND [MAGD]<=(2001)))
GO
ALTER TABLE [dbo].[GD_CHUYENTIEN] CHECK CONSTRAINT [repl_identity_range_D571EE49_692A_48B8_87DF_2CF8F2FB1DF3]
GO
ALTER TABLE [dbo].[GD_GOIRUT]  WITH CHECK ADD  CONSTRAINT [CK_LOAIGD] CHECK  (([LOAIGD]='RT' OR [LOAIGD]='GT'))
GO
ALTER TABLE [dbo].[GD_GOIRUT] CHECK CONSTRAINT [CK_LOAIGD]
GO
ALTER TABLE [dbo].[GD_GOIRUT]  WITH CHECK ADD  CONSTRAINT [CK_SOTIEN] CHECK  (([SOTIEN]>=(100000)))
GO
ALTER TABLE [dbo].[GD_GOIRUT] CHECK CONSTRAINT [CK_SOTIEN]
GO
ALTER TABLE [dbo].[GD_GOIRUT]  WITH NOCHECK ADD  CONSTRAINT [repl_identity_range_CD8E86A4_4C63_40CB_9231_0547232852EA] CHECK NOT FOR REPLICATION (([MAGD]>=(1) AND [MAGD]<=(1001) OR [MAGD]>(1001) AND [MAGD]<=(2001)))
GO
ALTER TABLE [dbo].[GD_GOIRUT] CHECK CONSTRAINT [repl_identity_range_CD8E86A4_4C63_40CB_9231_0547232852EA]
GO
ALTER TABLE [dbo].[TaiKhoan]  WITH CHECK ADD  CONSTRAINT [CK_SODU] CHECK  (([SODU]>=(0)))
GO
ALTER TABLE [dbo].[TaiKhoan] CHECK CONSTRAINT [CK_SODU]
GO
/****** Object:  StoredProcedure [dbo].[SP_CHECK_SODU]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[SP_CHECK_SODU]
(
	@sotk nchar(9)
)
as
select tk.SOTK, tk.SODU
from TaiKhoan tk
where tk.SOTK= @sotk
GO
/****** Object:  StoredProcedure [dbo].[sp_CHECK_Username]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create proc [dbo].[sp_CHECK_Username]
( @username varchar(50) )
as

select * from master.dbo.syslogins 
where name = @username
GO
/****** Object:  StoredProcedure [dbo].[SP_CHUYENTIEN]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[SP_CHUYENTIEN]
 @TKCHUYEN NVARCHAR (10) , @TKNHAN NVARCHAR (10), @SOTIEN BIGINT
 AS
SET XACT_ABORT ON
BEGIN TRANSACTION

BEGIN TRY
   UPDATE TAIKHOAN 
     SET SODU = SODU+ @SOTIEN 
     WHERE SOTK= @TKNHAN

   UPDATE TAIKHOAN 
     SET SODU = SODU -  @SOTIEN 
     WHERE SOTK= @TKCHUYEN 
 
  COMMIT
END TRY

BEGIN CATCH
   ROLLBACK
   DECLARE @ErrorMessage VARCHAR(2000)
   SELECT @ErrorMessage = 'Lỗi: ' + ERROR_MESSAGE()
   RAISERROR(@ErrorMessage, 16, 1)
END CATCH
GO
/****** Object:  StoredProcedure [dbo].[sp_DangKy]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[sp_DangKy]
  @USERNAME VARCHAR(50),
  @PASS VARCHAR(50),
  @MANV VARCHAR(50),
  @ROLE VARCHAR(50)
AS
  DECLARE @RET INT

  -- Tạo login từ biến LGNAME VÀ PASS
  EXEC @RET = SP_ADDLOGIN @USERNAME, @PASS,'NGANHANG' 
  IF (@RET = 1)  -- LOGIN NAME BI TRUNG
     RETURN 1

  -- Thêm User vào CSDL hiện hành từ biến USERNAME
  EXEC @RET= SP_GRANTDBACCESS @USERNAME, @MANV 
  IF (@RET = 1)  -- USER  NAME BI TRUNG
  BEGIN
       EXEC SP_DROPLOGIN @USERNAME
       RETURN 2
  END

  -- Thêm Role NganHang hoặc ChiNhanh cho User từ biến USERNAME VÀ ROLE
  EXEC sp_addrolemember @ROLE, @MANV

RETURN 0  -- THANH CONG
GO
/****** Object:  StoredProcedure [dbo].[sp_FindCustomer]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[sp_FindCustomer]
  @cmnd nchar(10)
AS


 IF  exists(select CMND from  dbo.KhachHang  where CMND =@cmnd)
 BEGIN
   	SELECT  CMND,HOTEN=HO+' '+TEN, DIACHI,PHAI, NGAYCAP,SODT, MACN
   	 FROM  dbo.KhachHang   WHERE CMND =@cmnd
   	                  
 END

 ELSE
  IF  exists(select CMND from  LINK0.NGANHANG.dbo.KhachHang  where CMND =@cmnd)
  BEGIN

   	SELECT  CMND,HOTEN=HO+' '+TEN, DIACHI,PHAI, NGAYCAP,SODT, MACN
   	 FROM  LINK0.NGANHANG.dbo.KhachHang nv  WHERE CMND =@cmnd
   	                  
  END
  ELSE  -- không có kh
     raiserror ( 'Khach hang ban tim khong co', 16, 1)
GO
/****** Object:  StoredProcedure [dbo].[sp_FindEMP]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[sp_FindEMP]
  @manv nchar(10)
AS


 IF  exists(select MANV from  dbo.NhanVien  where MANV =@manv)
 BEGIN
   	SELECT  MANV,HOTEN=nv.HO+' '+nv.TEN, nv.DIACHI,nv.PHAI, nv.SODT, MACN
   	 FROM  dbo.Nhanvien nv  WHERE MANV =@manv
   	                  
 END

 ELSE
  IF  exists(select MANV from  LINK0.NGANHANG.dbo.NHANVIEN  where MANV =@manv)
  BEGIN

   	SELECT  MANV,HOTEN=nv.HO+' '+nv.TEN, nv.DIACHI,nv.PHAI, nv.SODT, MACN
   	 FROM  LINK0.NGANHANG.dbo.Nhanvien nv  WHERE MANV =@manv
   	                  
  END
  ELSE  -- không có nv
     raiserror ( 'Ma nhan vien ban tim khong co', 16, 1)
GO
/****** Object:  StoredProcedure [dbo].[SP_Get_THONGTIN_Login]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[SP_Get_THONGTIN_Login]
@TENLOGIN NVARCHAR (50)
AS
DECLARE @TENUSER NVARCHAR(50)
SELECT @TENUSER=NAME FROM sys.sysusers WHERE sid = SUSER_SID(@TENLOGIN)
 
 SELECT USERNAME = @TENUSER, 
  HOTEN = (SELECT HO+ ' '+ TEN FROM NHANVIEN  WHERE MANV = @TENUSER ),
   TENNHOM= NAME
   FROM sys.sysusers 
   WHERE UID = (SELECT GROUPUID 
                 FROM SYS.SYSMEMBERS 
                   WHERE MEMBERUID= (SELECT UID FROM sys.sysusers 
                                      WHERE NAME=@TENUSER))
GO
/****** Object:  StoredProcedure [dbo].[SP_GUI]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[SP_GUI]
 @sotk NVARCHAR (10), @sotien BIGINT
 AS
SET XACT_ABORT ON
BEGIN TRANSACTION

BEGIN TRY

   UPDATE TAIKHOAN 
     SET SODU = SODU +  @sotien
     WHERE SOTK= @sotk
 
  COMMIT
END TRY

BEGIN CATCH
   ROLLBACK
   DECLARE @ErrorMessage VARCHAR(2000)
   SELECT @ErrorMessage = 'Lỗi: ' + ERROR_MESSAGE()
   RAISERROR(@ErrorMessage, 16, 1)
END CATCH
GO
/****** Object:  StoredProcedure [dbo].[sp_LK_KHACHHANG]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[sp_LK_KHACHHANG]
(
	@cn nchar(10)
)
as

select CMND, HOTEN= HO+' '+TEN, DIACHI, PHAI, NGAYCAP,SODT,MACN
from KhachHang
where MACN= @cn
order by TEN 
GO
/****** Object:  StoredProcedure [dbo].[sp_LK_TAIKHOAN]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create proc [dbo].[sp_LK_TAIKHOAN]
(
	@tungay date,
	@denngay date
)
as

select SOTK,CMND, MACN, NGAYMOTK
from TaiKhoan
where NGAYMOTK between @tungay and @denngay
GO
/****** Object:  StoredProcedure [dbo].[SP_RUT]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[SP_RUT]
 @sotk NVARCHAR (10), @sotien BIGINT
 AS
SET XACT_ABORT ON
BEGIN TRANSACTION

BEGIN TRY

   UPDATE TAIKHOAN 
     SET SODU = SODU -  @sotien
     WHERE SOTK= @sotk
 
  COMMIT
END TRY

BEGIN CATCH
   ROLLBACK
   DECLARE @ErrorMessage VARCHAR(2000)
   SELECT @ErrorMessage = 'Lỗi: ' + ERROR_MESSAGE()
   RAISERROR(@ErrorMessage, 16, 1)
END CATCH
GO
/****** Object:  StoredProcedure [dbo].[sp_SAOKE1]    Script Date: 5/29/2022 9:47:50 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[sp_SAOKE1]
(
	@sotk nchar(9),
	@tungay date,
	@denngay date
)
as 

	declare @saoke table
	(
		STT INT IDENTITY(1,1),
		SODUDAU decimal(18,2) default 0,
		SOTK nchar(9),
		LOAIGD nchar(2),
		NGAYGD datetime,
		SOTIEN money,
		SODUCUOI decimal(18,2) default 0
	)

	insert into @saoke(SOTK, NGAYGD, SOTIEN, LOAIGD) 
	select SOTK, NGAYGD, SOTIEN, LOAIGD
	from GD_GOIRUT where SOTK = @sotk AND NGAYGD between @tungay and @denngay
	union all
	select SOTK_CHUYEN, NGAYGD ,SOTIEN, LOAIGD= 'CT' 
	from GD_CHUYENTIEN where SOTK_CHUYEN = @sotk AND NGAYGD between @tungay and @denngay
	union all
	select SOTK_NHAN, NGAYGD ,SOTIEN, LOAIGD= 'NT' 
	from GD_CHUYENTIEN where SOTK_NHAN = @sotk AND NGAYGD between @tungay and @denngay
	order by NGAYGD													


	declare @soduht table
	(
		STT INT IDENTITY(1,1),
		SODUHT decimal(18,2) default 0
	)
	insert into @soduht(SODUHT)
	select soduhientai=case when sk.LOAIGD='GT' or sk.LOAIGD='NT' then tk.SODU - sk.SOTIEN else tk.SODU + sk.SOTIEN end
				from @saoke sk, TaiKhoan tk
				where sk.SOTK = tk.SOTK AND tk.SOTK = @sotk AND NGAYGD >= @tungay

	declare @sodudau decimal(18,2)
	declare @soducuoi decimal(18,2)
	declare @sodu decimal(18,2)

	declare @minID INT
	declare @maxID INT

	select @minID = (select MIN(STT) from @saoke)
	select @maxID = (select MAX(STT) from @saoke)

	set @sodudau = (select SODUHT from @soduht where STT= (select MAX(STT) from @soduht))
	update @saoke set SODUDAU = @sodudau where STT= '1'

	while @minID<=@maxID
	begin 

		select @soducuoi = case when sk.LOAIGD ='RT' or sk.LOAIGD='CT' then SODUDAU-SK.SOTIEN else SODUDAU + SK.SOTIEN end
		from @saoke sk where STT<=@minID

		update @saoke set SODUDAU = @soducuoi  where STT=@minID+1
		update @saoke set SODUCUOI = @soducuoi  where STT=@minID
           SET @minID=@minID+1					
	end
	select * from @saoke
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'‘GT’ : gởi tiền vào TK , ‘RT’ : rút tiền khỏi TK' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'GD_GOIRUT', @level2type=N'COLUMN',@level2name=N'LOAIGD'
GO
