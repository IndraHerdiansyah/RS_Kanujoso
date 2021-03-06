if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_DataKomponenIndex]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_DataKomponenIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_DetailKomponenIndex]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_DetailKomponenIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_DetailPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_DetailPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_HitungIndexPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_HitungIndexPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_JP_DataPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_JP_DataPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_KeluargaPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_KeluargaPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_KomponenIndexKaryawan]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_KomponenIndexKaryawan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_KonversiJabatanKeDetailKomponenIndex]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_KonversiJabatanKeDetailKomponenIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_KonversiPendidikanKeDetailKomponenIndex]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_KonversiPendidikanKeDetailKomponenIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_RiwayatGaji]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_RiwayatGaji]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_S_DataPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_S_DataPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_S_HitungIndex]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_S_HitungIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_Tempatbertugas]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_Tempatbertugas]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_Y_JenisPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_Y_JenisPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_DataPegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_DataPegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_S_Pegawai]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_S_Pegawai]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_DataKomponenIndex
AS
SELECT     dbo.KomponenIndex.KdJenisKomponenIndex, dbo.JenisKomponenIndex.JenisKomponenIndex, dbo.KomponenIndex.KdKomponenIndex, 
                      dbo.KomponenIndex.KomponenIndex
FROM         dbo.KomponenIndex INNER JOIN
                      dbo.JenisKomponenIndex ON dbo.KomponenIndex.KdJenisKomponenIndex = dbo.JenisKomponenIndex.KdJenisKomponenIndex

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_DetailKomponenIndex
AS
SELECT     dbo.DetailKomponenIndex.KdKomponenIndex, dbo.KomponenIndex.KomponenIndex, dbo.DetailKomponenIndex.KdDetailKomponenIndex, 
                      dbo.DetailKomponenIndex.DetailKomponenIndex, dbo.DetailKomponenIndex.NilaiIndexStandar, dbo.DetailKomponenIndex.RateIndex
FROM         dbo.DetailKomponenIndex INNER JOIN
                      dbo.KomponenIndex ON dbo.DetailKomponenIndex.KdKomponenIndex = dbo.KomponenIndex.KdKomponenIndex

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_DetailPegawai
AS
SELECT     dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap, dbo.DetailPegawai.Agama, dbo.DetailPegawai.StatusPerkawinan, 
                      dbo.DetailPegawai.GolonganDarah, dbo.DetailPegawai.Hobby, dbo.DetailPegawai.TinggiBadan, dbo.DetailPegawai.BeratBadan, 
                      dbo.DetailPegawai.JenisRambut, dbo.DetailPegawai.BentukMuka, dbo.DetailPegawai.WarnaKulit, dbo.DetailPegawai.CiriCiriKhas, 
                      dbo.DetailPegawai.CacatTubuh, dbo.DataPegawai.JenisKelamin, dbo.JenisPegawai.JenisPegawai, dbo.Jabatan.NamaJabatan
FROM         dbo.DetailPegawai RIGHT OUTER JOIN
                      dbo.DataPegawai ON dbo.DetailPegawai.IdPegawai = dbo.DataPegawai.IdPegawai LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.DataPegawai.KdJabatan = dbo.Jabatan.KdJabatan LEFT OUTER JOIN
                      dbo.JenisPegawai ON dbo.DataPegawai.KdJenisPegawai = dbo.JenisPegawai.KdJenisPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_HitungIndexPegawai
AS
SELECT     dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap, dbo.JenisPegawai.JenisPegawai, dbo.Jabatan.NamaJabatan, dbo.DataPegawai.NIP, 
                      dbo.JenisKomponenIndex.JenisKomponenIndex, dbo.KomponenIndex.KomponenIndex, dbo.DetailKomponenIndex.DetailKomponenIndex, 
                      dbo.DetailKomponenIndex.NilaiIndexStandar, dbo.DetailKomponenIndex.RateIndex, dbo.TotalScoreIndex.TglHitung
FROM         dbo.Jabatan RIGHT OUTER JOIN
                      dbo.DataPegawai ON dbo.Jabatan.KdJabatan = dbo.DataPegawai.KdJabatan LEFT OUTER JOIN
                      dbo.JenisPegawai ON dbo.DataPegawai.KdJenisPegawai = dbo.JenisPegawai.KdJenisPegawai LEFT OUTER JOIN
                      dbo.KomponenIndex INNER JOIN
                      dbo.JenisKomponenIndex ON dbo.KomponenIndex.KdJenisKomponenIndex = dbo.JenisKomponenIndex.KdJenisKomponenIndex INNER JOIN
                      dbo.DetailKomponenIndex ON dbo.KomponenIndex.KdKomponenIndex = dbo.DetailKomponenIndex.KdKomponenIndex INNER JOIN
                      dbo.TotalScoreIndex ON dbo.DetailKomponenIndex.KdDetailKomponenIndex = dbo.TotalScoreIndex.KdDetailKomponenIndex ON 
                      dbo.DataPegawai.IdPegawai = dbo.TotalScoreIndex.IdPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_JP_DataPegawai
AS
SELECT     dbo.DataPegawai.IdPegawai, dbo.JenisPegawai.JenisPegawai, dbo.DataPegawai.NamaLengkap, dbo.DataPegawai.JenisKelamin, 
                      dbo.DataPegawai.TempatLahir, dbo.DataPegawai.TglLahir, dbo.Pangkat.NamaPangkat, dbo.GolonganPegawai.NamaGolongan, 
                      dbo.Jabatan.NamaJabatan, dbo.Pendidikan.Pendidikan, dbo.DataPegawai.NIP
FROM         dbo.DataPegawai LEFT OUTER JOIN
                      dbo.Pendidikan ON dbo.DataPegawai.KdPendidikanTerakhir = dbo.Pendidikan.KdPendidikan LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.DataPegawai.KdJabatan = dbo.Jabatan.KdJabatan LEFT OUTER JOIN
                      dbo.GolonganPegawai ON dbo.DataPegawai.KdGolongan = dbo.GolonganPegawai.KdGolongan LEFT OUTER JOIN
                      dbo.Pangkat ON dbo.DataPegawai.KdPangkat = dbo.Pangkat.KdPangkat LEFT OUTER JOIN
                      dbo.JenisPegawai ON dbo.DataPegawai.KdJenisPegawai = dbo.JenisPegawai.KdJenisPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_KeluargaPegawai
AS
SELECT     TOP 100 PERCENT dbo.KeluargaPegawai.IdPegawai, dbo.KeluargaPegawai.NoUrut, dbo.HubunganKeluarga.NamaHubungan, 
                      dbo.KeluargaPegawai.NamaLengkap, dbo.KeluargaPegawai.JenisKelamin, dbo.KeluargaPegawai.TglLahir, dbo.Pekerjaan.Pekerjaan, 
                      dbo.Pendidikan.Pendidikan, dbo.KeluargaPegawai.Keterangan
FROM         dbo.KeluargaPegawai LEFT OUTER JOIN
                      dbo.Pendidikan ON dbo.KeluargaPegawai.KdPendidikan = dbo.Pendidikan.KdPendidikan LEFT OUTER JOIN
                      dbo.Pekerjaan ON dbo.KeluargaPegawai.KdPekerjaan = dbo.Pekerjaan.KdPekerjaan LEFT OUTER JOIN
                      dbo.HubunganKeluarga ON dbo.KeluargaPegawai.KdHubungan = dbo.HubunganKeluarga.Hubungan

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_KomponenIndexKaryawan
AS
SELECT     dbo.JenisKomponenIndex.JenisKomponenIndex, dbo.KomponenIndex.KomponenIndex, dbo.DetailKomponenIndex.DetailKomponenIndex, 
                      dbo.DetailKomponenIndex.NilaiIndexStandar, dbo.DetailKomponenIndex.RateIndex, dbo.ConvertDetailKomponenIndexToPendidikan.KdPendidikan, 
                      dbo.ConvertDetailKomponenIndexToJabatan.KdJabatan
FROM         dbo.JenisKomponenIndex INNER JOIN
                      dbo.KomponenIndex ON dbo.JenisKomponenIndex.KdJenisKomponenIndex = dbo.KomponenIndex.KdJenisKomponenIndex INNER JOIN
                      dbo.DetailKomponenIndex ON dbo.KomponenIndex.KdKomponenIndex = dbo.DetailKomponenIndex.KdKomponenIndex LEFT OUTER JOIN
                      dbo.ConvertDetailKomponenIndexToPendidikan ON 
                      dbo.DetailKomponenIndex.KdDetailKomponenIndex = dbo.ConvertDetailKomponenIndexToPendidikan.KdDetailKomponenIndex LEFT OUTER JOIN
                      dbo.ConvertDetailKomponenIndexToJabatan ON 
                      dbo.DetailKomponenIndex.KdDetailKomponenIndex = dbo.ConvertDetailKomponenIndexToJabatan.KdDetailKomponenIndex

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_KonversiJabatanKeDetailKomponenIndex
AS
SELECT     dbo.Jabatan.KdJabatan, dbo.Jabatan.NamaJabatan, dbo.DetailKomponenIndex.KdDetailKomponenIndex, 
                      dbo.DetailKomponenIndex.DetailKomponenIndex, dbo.DetailKomponenIndex.NilaiIndexStandar, dbo.DetailKomponenIndex.RateIndex
FROM         dbo.ConvertDetailKomponenIndexToJabatan INNER JOIN
                      dbo.DetailKomponenIndex ON 
                      dbo.ConvertDetailKomponenIndexToJabatan.KdDetailKomponenIndex = dbo.DetailKomponenIndex.KdDetailKomponenIndex INNER JOIN
                      dbo.Jabatan ON dbo.ConvertDetailKomponenIndexToJabatan.KdJabatan = dbo.Jabatan.KdJabatan

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_KonversiPendidikanKeDetailKomponenIndex
AS
SELECT     dbo.Pendidikan.KdPendidikan, dbo.Pendidikan.Pendidikan, dbo.DetailKomponenIndex.KdDetailKomponenIndex, 
                      dbo.DetailKomponenIndex.DetailKomponenIndex, dbo.DetailKomponenIndex.NilaiIndexStandar, dbo.DetailKomponenIndex.RateIndex
FROM         dbo.ConvertDetailKomponenIndexToPendidikan INNER JOIN
                      dbo.Pendidikan ON dbo.ConvertDetailKomponenIndexToPendidikan.KdPendidikan = dbo.Pendidikan.KdPendidikan INNER JOIN
                      dbo.DetailKomponenIndex ON 
                      dbo.ConvertDetailKomponenIndexToPendidikan.KdDetailKomponenIndex = dbo.DetailKomponenIndex.KdDetailKomponenIndex

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_RiwayatGaji
AS
SELECT     dbo.RiwayatGaji.IdPegawai, dbo.KomponenGaji.KomponenGaji, dbo.RiwayatGaji.TglBerlaku, dbo.RiwayatGaji.Jumlah
FROM         dbo.RiwayatGaji LEFT OUTER JOIN
                      dbo.KomponenGaji ON dbo.RiwayatGaji.KdKomponenGaji = dbo.KomponenGaji.KdKomponenGaji

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_S_DataPegawai
AS
SELECT     TOP 100 PERCENT dbo.DataPegawai.IdPegawai, dbo.JenisPegawai.JenisPegawai, dbo.DataPegawai.NamaLengkap, dbo.DataPegawai.JenisKelamin, 
                      dbo.DataPegawai.TempatLahir, dbo.DataPegawai.TglLahir, dbo.Pangkat.NamaPangkat, dbo.GolonganPegawai.NamaGolongan, 
                      dbo.Jabatan.NamaJabatan, dbo.Pendidikan.Pendidikan, dbo.DataPegawai.NIP, dbo.DataPegawai.StatusAktif
FROM         dbo.DataPegawai LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.DataPegawai.KdJabatan = dbo.Jabatan.KdJabatan LEFT OUTER JOIN
                      dbo.Pendidikan ON dbo.DataPegawai.KdPendidikanTerakhir = dbo.Pendidikan.KdPendidikan LEFT OUTER JOIN
                      dbo.GolonganPegawai ON dbo.DataPegawai.KdGolongan = dbo.GolonganPegawai.KdGolongan LEFT OUTER JOIN
                      dbo.Pangkat ON dbo.DataPegawai.KdPangkat = dbo.Pangkat.KdPangkat LEFT OUTER JOIN
                      dbo.JenisPegawai ON dbo.DataPegawai.KdJenisPegawai = dbo.JenisPegawai.KdJenisPegawai
ORDER BY dbo.DataPegawai.IdPegawai, dbo.JenisPegawai.JenisPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_S_HitungIndex
AS
SELECT DISTINCT 
                      dbo.DetailKomponenIndex.KdKomponenIndex, dbo.TotalScoreIndex.TglHitung, dbo.TotalScoreIndex.IdPegawai, 
                      dbo.TotalScoreIndex.KdDetailKomponenIndex, dbo.TotalScoreIndex.NilaiIndex, dbo.TotalScoreIndex.IdUser
FROM         dbo.TotalScoreIndex INNER JOIN
                      dbo.DetailKomponenIndex ON dbo.TotalScoreIndex.KdDetailKomponenIndex = dbo.DetailKomponenIndex.KdDetailKomponenIndex

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_Tempatbertugas
AS
SELECT     dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap, dbo.Instalasi.KdInstalasi, dbo.Instalasi.NamaInstalasi, dbo.Ruangan.KdRuangan, 
                      dbo.Ruangan.NamaRuangan, dbo.Jabatan.KdJabatan, dbo.Jabatan.NamaJabatan, dbo.TempatBertugas.TglMulai, dbo.TempatBertugas.TglAkhir, 
                      dbo.TempatBertugas.NoSuratKeputusan
FROM         dbo.Ruangan RIGHT OUTER JOIN
                      dbo.TempatBertugas ON dbo.Ruangan.KdRuangan = dbo.TempatBertugas.KdRuangan LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.TempatBertugas.KdJabatan = dbo.Jabatan.KdJabatan LEFT OUTER JOIN
                      dbo.Instalasi ON dbo.TempatBertugas.KdInstalasi = dbo.Instalasi.KdInstalasi LEFT OUTER JOIN
                      dbo.DataPegawai ON dbo.TempatBertugas.IdPegawai = dbo.DataPegawai.IdPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_Y_JenisPegawai
AS
SELECT     dbo.JenisPegawai.KdKelompokPegawai, dbo.KelompokPegawai.KelompokPegawai, dbo.JenisPegawai.KdJenisPegawai, 
                      dbo.JenisPegawai.JenisPegawai
FROM         dbo.JenisPegawai INNER JOIN
                      dbo.KelompokPegawai ON dbo.JenisPegawai.KdKelompokPegawai = dbo.KelompokPegawai.KdKelompokPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.v_DataPegawai
AS
SELECT     TOP 100 PERCENT dbo.DataPegawai.IdPegawai, dbo.DataPegawai.KdJenisPegawai, dbo.JenisPegawai.JenisPegawai, 
                      dbo.JenisPegawai.KdKelompokPegawai, dbo.KelompokPegawai.KelompokPegawai, dbo.DataPegawai.NamaLengkap, dbo.DataPegawai.JenisKelamin, 
                      dbo.Pangkat.NamaPangkat, dbo.GolonganPegawai.NamaGolongan, dbo.Jabatan.NamaJabatan
FROM         dbo.KelompokPegawai INNER JOIN
                      dbo.JenisPegawai ON dbo.KelompokPegawai.KdKelompokPegawai = dbo.JenisPegawai.KdKelompokPegawai RIGHT OUTER JOIN
                      dbo.DataPegawai ON dbo.JenisPegawai.KdJenisPegawai = dbo.DataPegawai.KdJenisPegawai LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.DataPegawai.KdJabatan = dbo.Jabatan.KdJabatan LEFT OUTER JOIN
                      dbo.Pangkat ON dbo.DataPegawai.KdPangkat = dbo.Pangkat.KdPangkat LEFT OUTER JOIN
                      dbo.GolonganPegawai ON dbo.DataPegawai.KdGolongan = dbo.GolonganPegawai.KdGolongan
ORDER BY dbo.DataPegawai.NamaLengkap

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.v_S_Pegawai
AS
SELECT     TOP 100 PERCENT dbo.DataPegawai.IdPegawai, dbo.DataPegawai.NamaLengkap, dbo.DataPegawai.JenisKelamin AS Sex, 
                      dbo.JenisPegawai.JenisPegawai, dbo.DataPegawai.KdJabatan, dbo.Jabatan.NamaJabatan, dbo.Pendidikan.KdPendidikan, 
                      dbo.Pendidikan.Pendidikan
FROM         dbo.DataPegawai LEFT OUTER JOIN
                      dbo.JenisPegawai ON dbo.DataPegawai.KdJenisPegawai = dbo.JenisPegawai.KdJenisPegawai LEFT OUTER JOIN
                      dbo.Pendidikan ON dbo.DataPegawai.KdPendidikanTerakhir = dbo.Pendidikan.KdPendidikan LEFT OUTER JOIN
                      dbo.Jabatan ON dbo.DataPegawai.KdJabatan = dbo.Jabatan.KdJabatan
ORDER BY dbo.DataPegawai.IdPegawai

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

