
CREATE PROCEDURE [dbo].AUD_RuangKerja
 @KdRuangKerja char(3),
 @RuangKerja varchar(50),
 @Status char(1) -- A=Simpan & Ubah ; D=Hapus
AS
 declare @KdRuangKerjaTemp varchar(3)
 
 select @KdRuangKerjaTemp=KdRuangKerja from RuangKerja where KdRuangKerja=@KdRuangKerja
 if @@rowcount=0
 BEGIN
	SELECT @KdRuangKerjaTemp = MAX (KdRuangKerja) FROM RuangKerja
	IF (@KdRuangKerjaTemp IS NULL)
		SET @KdRuangKerjaTemp = '001'
	Else
		SET @KdRuangKerjaTemp = dbo.formatNomor(@KdRuangKerjaTemp+1, 3)
	insert into RuangKerja values(@KdRuangKerjaTemp,@RuangKerja)
 END
 else
 begin
	if upper(@Status)='A'
		begin
			update RuangKerja set RuangKerja = @RuangKerja where KdRuangKerja = @KdRuangKerja
		end
	else
		begin
			delete from RuangKerja where KdRuangKerja = @KdRuangKerja
		end
 end

 return @@error







