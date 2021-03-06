if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_DetailPegawai]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_DetailPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_HitungIndex]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_HitungIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_KeluargaPegawai]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_KeluargaPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RExtraPelatihan]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RExtraPelatihan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RGaji]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RGaji]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_ROrganisasi]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_ROrganisasi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RPddkF]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RPddkF]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RPddkNF]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RPddkNF]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RPekerjaan]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RPekerjaan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RPjlnDns]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RPjlnDns]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_RPrestasi]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_RPrestasi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AU_HRD_TempatBertugas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AU_HRD_TempatBertugas]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Proc_D_GenerateIDPegawai]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Proc_D_GenerateIDPegawai]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_DetailPegawai
 @IdPegawai char(10),
 @Agama varchar(20),
 @StatusPerkawinan varchar(20),
 @GolonganDarah varchar(2),
 @Hobby varchar(100),
 @TinggiBadan varchar(20),
 @BeratBadan varchar(20),
 @JenisRambut varchar(20),
 @BentukMuka varchar(50),
 @WarnaKulit varchar(50),
 @CiriCiriKhas varchar(300),
 @CacatTubuh varchar(200)
As

SELECT IdPegawai FROM DetailPegawai WHERE IdPegawai = @IdPegawai
IF @@Rowcount = 0
	INSERT INTO DetailPegawai VALUES ( @IdPegawai, @Agama, @StatusPerkawinan, @GolonganDarah, @Hobby, @TinggiBadan, @BeratBadan, @JenisRambut, @BentukMuka, @WarnaKulit, @CiriCiriKhas, @CacatTubuh )
ELSE
	UPDATE DetailPegawai SET Agama = @Agama, StatusPerkawinan = @StatusPerkawinan, GolonganDarah = @GolonganDarah, Hobby = @Hobby, TinggiBadan = @TinggiBadan, BeratBadan = @BeratBadan, JenisRambut = @JenisRambut, BentukMuka = @BentukMuka, WarnaKulit = @WarnaKulit, CiriCiriKhas = @CiriCiriKhas, CacatTubuh = @CacatTubuh WHERE IdPegawai = @IdPegawai
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_HitungIndex
 @TglHitung datetime,
 @IdPegawai char(10),
 @KdDetailKomponenIndex varchar(5),
 @NilaiIndex int,
 @IdUser char(10),
 @KdKomponenIndex varchar(3)
As

INSERT INTO TotalScoreIndex Values (@TglHitung, @IdPegawai, @KdDetailKomponenIndex, @NilaiIndex, @IdUser)

RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_KeluargaPegawai
 @IdPegawai char(10),
 @KdHubungan char(2),
 @NoUrut char(2),
 @NamaLengkap varchar(30),
 @JenisKelamin char(1),
 @TglLahir datetime,
 @KdPekerjaan char(2),
 @KdPendidikan char(2),
 @Keterangan varchar(100),
 @OutputNoUrut char(2) output
As
Declare @tempNoUrut as char(2)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM KeluargaPegawai WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '01'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 2 )
	END

INSERT KeluargaPegawai Values (@IdPegawai,@KdHubungan,@tempNoUrut, @NamaLengkap,@JenisKelamin,@TglLahir,@KdPekerjaan,@KdPendidikan,@Keterangan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE KeluargaPegawai SET KdHubungan = @KdHubungan, NamaLengkap=@NamaLengkap,JenisKelamin=@JenisKElamin,TglLahir=@TglLahir,KdPekerjaan=@KdPekerjaan,KdPendidikan=@KdPendidikan,Keterangan=@Keterangan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RExtraPelatihan
 @IdPegawai char(10),
 @NoUrut char(3),
 @NamaPelatihan varchar(100),
 @KedudukanPeranan varchar(50),
 @TglPenyelenggaraan datetime,
 @LamaPelatihan varchar(20),
 @InstansiPenyelenggara varchar(100),
 @AlamatPenyelenggaraan char(10),
 @OutputNoUrut char(3) output
As
Declare @tempNoUrut as char(3)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatExtraPelatihan WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '001'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 3 )
	END

INSERT RiwayatExtraPelatihan Values (@IdPegawai, @tempNoUrut, @NamaPelatihan, @KedudukanPeranan, @TglPenyelenggaraan, @LamaPelatihan, @InstansiPenyelenggara, @AlamatPenyelenggaraan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatExtraPelatihan SET NamaPelatihan = @NamaPelatihan, KedudukanPeranan = @KedudukanPeranan, TglPenyelenggaraan = @TglPenyelenggaraan, LamaPelatihan = @LamaPelatihan,InstansiPenyelenggara = @InstansiPenyelenggara, AlamatPenyelenggaraan = @AlamatPenyelenggaraan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RGaji
 @IdPegawai char(10),
 @KdKomponenGaji char(2),
 @TglBerlaku datetime,
 @Jumlah money
As

DECLARE @KdKomponenGajiTemp char(2)

SELECT @KdKomponenGajiTemp=KdKomponenGaji from RiwayatGaji WHERE RiwayatGaji.IdPegawai=@IdPegawai AND KdKomponenGaji=@KdKomponenGaji
IF @@rowcount=0
	INSERT INTO RiwayatGaji Values (@IdPegawai, @KdKomponenGaji, @TglBerlaku, @Jumlah)
ELSE
	UPDATE RiwayatGaji SET TglBerlaku=@TglBerlaku, Jumlah=@Jumlah WHERE IdPegawai = @IdPegawai AND KdKomponenGaji=@KdKomponenGaji

RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_ROrganisasi
 @IdPegawai char(10),
 @NoUrut char(3),
 @NamaOrganisasi varchar(100),
 @Jabatan varchar(50),
 @TahunAwal char(4),
 @TahunAkhir char(4),
 @AlamatOrganisasi varchar(200),
 @NamaPemimpinOrganisasi varchar(50),
 @OutputNoUrut char(3) output
As
Declare @tempNoUrut as char(3)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatOrganisasi WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '001'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 3 )
	END

INSERT RiwayatOrganisasi Values (@IdPegawai, @tempNoUrut, @NamaOrganisasi, @Jabatan, @TahunAwal, @TahunAkhir, @AlamatOrganisasi, @NamaPemimpinOrganisasi)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatOrganisasi SET NamaOrganisasi = @NamaOrganisasi, Jabatan = @Jabatan, TahunAwal = @TahunAwal, TahunAkhir = @TahunAkhir, AlamatOrganisasi = @AlamatOrganisasi, NamaPemimpinOrganisasi = @NamaPemimpinOrganisasi WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RPddkF
 @IdPegawai char(10),
 @KdPendidikan char(2),
 @NoUrut char(2),
 @NamaPendidikan varchar(100),
 @Jurusan varchar(100),
 @TahunMasuk char(4),
 @TahunLulus char(4),
 @IndeksPrestasiKumulatif float,
 @GradeKelulusan varchar(50),
 @NoIjazah varchar(30),
 @TglIjazah datetime,
 @AlamatPendidikan varchar(200),
 @NamaPemimpinPendidikan char(50),
 @OutputNoUrut char(2) output
As
Declare @tempNoUrut as char(2)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatPendidikanFormal WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '01'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 2 )
	END

INSERT RiwayatPendidikanFormal Values (@IdPegawai, @KdPendidikan, @tempNoUrut, @NamaPendidikan, @Jurusan, @TahunMasuk, @TahunLulus, @IndeksPrestasiKumulatif, @GradeKelulusan,@NoIjazah, @TglIjazah,  @AlamatPendidikan, @NamaPemimpinPendidikan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatPendidikanFormal SET KdPendidikan=@KdPendidikan, NamaPendidikan = @NamaPendidikan, Jurusan = @Jurusan, TahunMasuk = @TahunMasuk, TahunLulus = @TahunLulus, IndeksPrestasiKumulatif = @IndeksPrestasiKumulatif, GradeKelulusan = @GradeKelulusan, NoIjazah=@NoIjazah, TglIjazah=@TglIjazah, AlamatPendidikan = @AlamatPendidikan,  NamaPemimpinPendidikan = @NamaPemimpinPendidikan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RPddkNF
 @IdPegawai char(10),
 @NoUrut char(3),
 @NamaPendidikan varchar(100),
 @LamaPendidikan varchar(20),
 @TglMulai datetime,
 @TglLulus datetime,
 @NoSertifikat varchar(50),
 @TglSertifikat datetime,
 @AlamatPendidikan varchar(200),
 @Keterangan char(100),
 @OutputNoUrut char(3) output
As
Declare @tempNoUrut as char(3)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatPendidikanNonFormal WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '001'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 3 )
	END

INSERT RiwayatPendidikanNonFormal Values (@IdPegawai, @tempNoUrut, @NamaPendidikan, @LamaPendidikan, @TglMulai, @TglLulus, @NoSertifikat, @TglSertifikat, @AlamatPendidikan, @Keterangan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatPendidikanNonFormal SET NamaPendidikan = @NamaPendidikan, LamaPendidikan = @LamaPendidikan, TglMulai = @TglMulai, TglLulus = @TglLulus, NoSertifikat = @NoSertifikat, TglSertifikat = @TglSertifikat, AlamatPendidikan = @AlamatPendidikan,  Keterangan = @Keterangan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RPekerjaan
 @IdPegawai char(10),
 @NoUrut char(2),
 @NamaPerusahaan varchar(100),
 @JabatanPosisi varchar(50),
 @UraianPekerjaan varchar(200),
 @TglMulai datetime,
 @TglAkhir datetime,
 @GajiPokok int,
 @NoSuratKeputusan varchar(50),
 @AlamatPerusahaan varchar(200),
 @OutputNoUrut char(2) output
As
Declare @tempNoUrut as char(2)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatPekerjaan WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '01'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 2 )
	END

INSERT RiwayatPekerjaan Values (@IdPegawai, @tempNoUrut, @NamaPerusahaan, @JabatanPosisi, @UraianPekerjaan, @TglMulai, @TglAkhir, @GajiPokok, @NoSuratKeputusan, @AlamatPerusahaan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatPekerjaan SET NamaPerusahaan = @NamaPerusahaan, JabatanPosisi = @JabatanPosisi, UraianPekerjaan= @UraianPekerjaan, TglMulai = @TglMulai, TglAkhir = @TglAkhir, GajiPokok = @GajiPokok, NoSuratKeputusan = @NoSuratKeputusan, AlamatPerusahaan = @AlamatPerusahaan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RPjlnDns
 @IdPegawai char(10),
 @NoUrut char(3),
 @KotaTujuan varchar(50),
 @NegaraTujuan varchar(50),
 @TujuanKunjungan varchar(200),
 @TglKunjungan datetime,
 @LamaKunjungan varchar(20),
 @PenyandangDana varchar(50),
 @OutputNoUrut char(3) output
As
Declare @tempNoUrut as char(3)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatPerjalananDinas WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '001'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 3 )
	END

INSERT RiwayatPerjalananDinas Values (@IdPegawai, @tempNoUrut, @KotaTujuan, @NegaraTujuan, @TujuanKunjungan, @TglKunjungan, @LamaKunjungan, @PenyandangDana)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatPerjalananDinas SET KotaTujuan = @KotaTujuan, NegaraTujuan = @NegaraTujuan, TujuanKunjungan = @TujuanKunjungan, TglKunjungan = @TglKunjungan, LamaKunjungan = @LamaKunjungan, PenyandangDana = @PenyandangDana WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_RPrestasi
 @IdPegawai char(10),
 @NoUrut char(2),
 @NamaPenghargaan varchar(100),
 @TahunDiperoleh char(4),
 @NamaInstansiPemberi varchar(100),
 @Keterangan varchar(150),
 @OutputNoUrut char(2) output
As
Declare @tempNoUrut as char(2)
IF @NoUrut IS NOT NULL
	GOTO Updating

NewData:
SELECT @tempNoUrut = MAX( CAST( NoUrut AS Integer ) ) FROM RiwayatPrestasi WHERE IdPegawai = @IdPegawai
IF @tempNoUrut IS NULL
	SET @tempNoUrut = '01'
ELSE
	BEGIN
	SET @tempNoUrut = @tempNoUrut + 1
	SET @tempNoUrut = dbo.formatnomor( @tempNoUrut, 2 )
	END

INSERT RiwayatPrestasi Values (@IdPegawai, @tempNoUrut, @NamaPenghargaan, @TahunDiperoleh, @NamaInstansiPemberi, @Keterangan)
SET @OutputNoUrut = @tempNoUrut
GOTO Ending

Updating:
UPDATE RiwayatPrestasi SET NamaPenghargaan = @NamaPenghargaan, TahunDiperoleh = @TahunDiperoleh, NamaInstansiPemberi = @NamaInstansiPemberi, Keterangan = @Keterangan WHERE IdPegawai = @IdPegawai AND NoUrut = @NoUrut
SET @OutputNoUrut = @NoUrut

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  Procedure AU_HRD_TempatBertugas
 @IdPegawai char(10),
 @KdInstalasi char(2),
 @KdRuangan char(3),
 @KdJabatan varchar(3),
 @TglMulai datetime,
 @TglAkhir datetime,
 @NoSuratKeputusan varchar(50)
As
INSERT TempatBertugas Values (@IdPegawai,@KdInstalasi, @KdRuangan, @KdJabatan, @TglMulai, @TglAkhir, @NoSuratKeputusan)
GOTO Ending

Updating:
UPDATE TempatBertugas SET KdInstalasi = @KdInstalasi, KdRuangan=@KdRuangan,KdJabatan=@KdJabatan,TglMulai=@TglMulai,TglAkhir=@TglAkhir,NoSuratKeputusan=@NoSuratKeputusan WHERE IdPegawai = @IdPegawai

Ending:
RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE Proc_D_GenerateIDPegawai
	@IdPegawai char(10),
	@KdJenisPegawai char(3),
	@NamaLengkap varchar(50),
	@JenisKelamin char(1),
	@TempatLahir varchar(50),
	@TglLahir datetime,
	@KdPangkat varchar(2),
	@KdGolongan varchar(2),
	@KdJabatan varchar(3),
	@KdPendidikanTerakhir char(2),
	@NIP varchar(10),
	@StatusAktif char(1),
	@OutputIDPegawai char(10) OUTPUT
AS
	DECLARE @nomor varchar(10)
	DECLARE @i as integer

	SELECT @nomor = MAX (CAST(RIGHT(IdPegawai,6) as integer)) FROM DataPegawai WHERE IdPegawai <> '8888888888'
	IF (@nomor IS NULL)
		BEGIN
			SET @nomor = @JenisKelamin + @KdJenisPegawai + '000001'
		END
	ELSE
		BEGIN
			SET @i = CAST(RIGHT(@nomor,6) as integer) + 1
			SET @nomor = @JenisKelamin + @KdJenisPegawai + dbo.formatNomor(@i,6)
		END

INSERT INTO DataPegawai Values(@nomor, @KdJenisPegawai, @NamaLengkap, @JenisKelamin, @TempatLahir, @TglLahir, @KdPangkat, @KdGolongan, @KdJabatan, @KdPendidikanTerakhir, @NIP, @StatusAktif)

SET @OutputIDPegawai = @nomor

RETURN @@ERROR
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

