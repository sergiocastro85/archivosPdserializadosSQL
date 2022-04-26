
USE BD
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- ========================================================================================================================
-- Author:		Sergio Andres Castro Cano
-- Create date: 25-02-2022
-- permite generar los pdf
---- ======================================================================================================================

--EXECUTE spCrearPdfOCsunegocio '192902,192905,192946,192948,192950,192951,192952,192954,192955,1929570','prueba'


CREATE PROCEDURE spCrearPdfOCsunegocio
	-- Add the parameters for the stored procedure here
	@OrdeCompra VARCHAR(MAX),
	@strUsuarioEjecuta VARCHAR(50)
	--@valorTasa DECIMAL (16,2),
	--@sUsuarioActualizacion VARCHAR(MAX),
	--@Proceso INT -- 0 Consultar, 1 Modificar
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	BEGIN TRY

	  -- Inicia el proceso de cambio a las bases de datos
	BEGIN TRANSACTION
		
		
	--DECLARE @outPutPath varchar(50) = '\\URL\publicomde\Compartida\Pruebas'
	DECLARE @outPutPath varchar(200) = '\\URL\publicomde\ArchivosPdfOC'
	, @i bigint
	, @init int
	, @data varbinary(max)
	, @fPath varchar(max)
	, @folderPath  varchar(max)


	DECLARE @Doctable TABLE (id INT IDENTITY(1,1), [FileName]  varchar(100), [Doc_Content] varBinary(max),dtmfecha DATETIME )


	INSERT INTO @Doctable( [FileName],[Doc_Content],dtmfecha)
	SELECT p.strNombreProveedor+'-'+[strNombreArchivo] AS strNombreArchivo, [binArchivo], [dtmfechaGeneracion]
	FROM dbSunegocio.[dbo].[TblSunArchivosOrdenesDeCompra] ac INNER JOIN dbSunegocio.dbo.TblSunOrdenesDeCompra oc
	ON oc.IngIdOrdenDeCompra = ac.IngIdOrdenDeCompra INNER JOIN dbSurenting.dbo.tblProveedores p ON p.IdProveedor = oc.IdProveedor
	WHERE oc.IngConsecutivo IN
		(
			SELECT CAST(value AS VARCHAR(MAX)) AS value
			FROM STRING_SPLIT(@OrdeCompra, ',')
		);

	SELECT @i = COUNT(1) FROM @Doctable




	WHILE @i >= 1

	BEGIN

			   SELECT
				@data = [Doc_Content],
				@fPath = @outPutPath +  '\' +[FileName] +'.pdf',
				@folderPath = @outPutPath
			   FROM @Doctable WHERE id = @i

		  --Create folder first
		  --SELECT * FROM @Doctable

		  EXEC sp_OACreate 'ADODB.Stream', @init OUTPUT; -- An instance created
		  EXEC sp_OASetProperty @init, 'Type', 1;
		  EXEC sp_OAMethod @init, 'Open'; -- Calling a method
		  EXEC sp_OAMethod @init, 'Write', NULL, @data; -- Calling a method
		  EXEC sp_OAMethod @init, 'SaveToFile', NULL, @fPath, 2; -- Calling a method
		  EXEC sp_OAMethod @init, 'Close'; -- Calling a method
		  EXEC sp_OADestroy @init; -- Closed the resources

		  print 'Document Generated at - '+  @fPath

		--Reset the variables for next use
		SELECT @data = NULL
		, @init = NULL
		, @fPath = NULL
		, @folderPath = NULL
		SET @i -= 1
	END

	IF ( SELECT COUNT(*) FROM @Doctable)>0
	BEGIN

	

		INSERT INTO dbSurenting.dbo.tblAuditoria
		(
			strPlaca,
			strPerfil,
			strMaquina,
			dtmFechaModificacion,
			strProceso,
            strValorAnt,
            strValorNuevo,
            strNombreMaestro,
            strVariable
		)
		SELECT TOP 1 'OC-' + CONVERT(VARCHAR(100), oc.IngConsecutivo),
			   @strUsuarioEjecuta,
			   'Reporting',
			   GETDATE(),
			   'GenerarPdfOC',
			   '0',
			   '0',
			   'RptGenerarPDFOC',
			   'NULL'
		FROM dbSunegocio.[dbo].[TblSunArchivosOrdenesDeCompra] ac
			INNER JOIN dbSunegocio.dbo.TblSunOrdenesDeCompra oc
				ON oc.IngIdOrdenDeCompra = ac.IngIdOrdenDeCompra
			INNER JOIN dbSurenting.dbo.tblProveedores p
				ON p.IdProveedor = oc.IdProveedor
		WHERE oc.IngConsecutivo IN
			  (
				  SELECT CAST(value AS VARCHAR(MAX)) AS value
				  FROM STRING_SPLIT(@OrdeCompra, ',')
			  );


		SELECT ac.[IngIdOrdenDeCompra],
			   p.strNombreProveedor + '-' + [strNombreArchivo] AS strNombreArchivo,
			   [dtmfechaGeneracion]
		FROM dbSunegocio.[dbo].[TblSunArchivosOrdenesDeCompra] ac
			INNER JOIN dbSunegocio.dbo.TblSunOrdenesDeCompra oc
				ON oc.IngIdOrdenDeCompra = ac.IngIdOrdenDeCompra
			INNER JOIN dbSurenting.dbo.tblProveedores p
				ON p.IdProveedor = oc.IdProveedor
		WHERE oc.IngConsecutivo IN
			  (
				  SELECT CAST(value AS VARCHAR(MAX)) AS value
				  FROM STRING_SPLIT(@OrdeCompra, ',')
			  );


	END 
	


		--  Finaliza el cambio a datos
	COMMIT TRANSACTION;
	END TRY
	BEGIN CATCH  
		--PRINT 1
		  --Se devuelve la transaccion en caso de ocurrir un error.
		  ROLLBACK TRANSACTION;
		  DECLARE @ErrorMessage NVARCHAR(max) = ERROR_MESSAGE() +' ROLLBACK REALIZADO '
          DECLARE @ErrorSeverity INT = ERROR_SEVERITY()
          DECLARE @ErrorState INT = ERROR_STATE()
          DECLARE @ErrorLine INT = ERROR_LINE()

          RAISERROR (@ErrorMessage,
                       @ErrorSeverity,
                       @ErrorState 
                    );
		
	END CATCH 	
END
GO
