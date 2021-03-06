
CREATE PROCEDURE [dbo].AUD_SubRuangKerja
 @KdSubRuangKerja char(3),
 @SubRuangKerja varchar(50),
 @KdRuangKerja char(3),
 @Status char(1) -- A=Simpan & Ubah ; D=Hapus
AS
 declare @KdSubRuangKerjaTemp varchar(3)
 
 select @KdSubRuangKerjaTemp=KdSubRuangKerja from SubRuangKerja where KdSubRuangKerja=@KdSubRuangKerja
 if @@rowcount=0
 BEGIN
	SELECT @KdSubRuangKerjaTemp = MAX (KdSubRuangKerja) FROM SubRuangKerja
	IF (@KdSubRuangKerjaTemp IS NULL)
		SET @KdSubRuangKerjaTemp = '001'
	Else
		SET @KdSubRuangKerjaTemp = dbo.formatNomor(@KdSubRuangKerjaTemp+1, 3)
	insert into SubRuangKerja values(@KdSubRuangKerjaTemp,@SubRuangKerja, @KdRuangKerja)
 END
 else
 begin
	if upper(@Status)='A'
		begin
			update SubRuangKerja set SubRuangKerja = @SubRuangKerja where KdSubRuangKerja = @KdSubRuangKerja
		end
	else
		begin
			delete from SubRuangKerja where KdSubRuangKerja = @KdSubRuangKerja
		end
 end

 return @@error







