CREATE FUNCTION [dbo].[FN_COMPRAS_VALIDACIONES](@DOCENTRY varchar(100))
RETURNS VARCHAR(200)
AS
/*
 DANIEL MORENO LOCALIZACION ECUATORIANA 
 VALIDACIONES PARA COMPRAS.
 GUATEMALA 11/05/2017
 ONESOLUTIONS S.A.
*/
BEGIN
DECLARE @ERROR VARCHAR(200) = '0';
SET @ERROR = '0';
DECLARE @CARDCODE VARCHAR(30)
DECLARE @DOCUMENTO VARCHAR(30)
DECLARE @AUTORI INT 

------FACTURA DE REEMBOLSO
    	  DECLARE @TIPO_IDENTI AS nvarchar(400) 
		  DECLARE @ID_PROVEEDOR AS nvarchar(400)
		  DECLARE @PAIS AS nvarchar(400)
		  DECLARE @TIPOPROVEEDOR AS nvarchar(400) 
		  DECLARE @CODDOC AS nvarchar(400)
		  DECLARE @ESTABLE AS nvarchar(400)
		  DECLARE @PTOEMISION AS nvarchar(400)
		  DECLARE @secuencialDocReembolso AS nvarchar(400)
		  DECLARE @numeroautorizacionDocReemb AS nvarchar(400) 
		  DECLARE @FECHAEMISION AS nvarchar(400) 
		  DECLARE @MOTIVO AS NVARCHAR(100)
		  DECLARE @RUC AS NVARCHAR(25)
		  


		  
/*AQUI VERIFICA SI DESEA GENERAR RETENCION*/     
	 IF (SELECT ISNULL(A.U_A_APLICARR,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY)= '@' 
	BEGIN				
				SET @ERROR = '---L(23) Debe de seleccionar si desea aplicar retencion al documento';
				RETURN @ERROR;
	END
	 IF (SELECT ISNULL(A.U_TI_COMPRO ,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY) = '@' 
			BEGIN
				SET @ERROR = '---L(25) Debe de seleccionar un tipo de comprobrante'
				RETURN @ERROR;
			END
	/*SI VA A GENERAR RETENCION*/
	ELSE IF (SELECT A.U_A_APLICARR  FROM OPCH A WHERE A.DocEntry = @DOCENTRY) = '01'
	BEGIN
	  IF (SELECT ISNULL(A.U_SUS_TRIBU,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY) = '@' 
			BEGIN
				SET @ERROR = '---L(24) Debe de seleccionar un sustento Tributario'
				RETURN @ERROR;
			END
		 IF (SELECT ISNULL(A.U_TI_COMPRO ,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY) = '@' 
			BEGIN
				SET @ERROR = '---L(25) Debe de seleccionar un tipo de comprobrante'
				RETURN @ERROR;
			END
	END

	/*VALIDACIONES PARA NUMATCARD*/
	  IF (SELECT ISNULL(A.NumAtCard ,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY  ) = '@'
		 BEGIN
		  SET @ERROR = '---L(30) Debe de ingresar un numero de referencia';
		  RETURN @ERROR;
		 END
	  ELSE
	  BEGIN
		IF (SELECT len (A.NumAtCard)  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) != 15
		BEGIN
			  SET @ERROR = '---L(30) El numero de referencia no puede ser menor y mayor a 15';
		      RETURN @ERROR;
		END
		IF (SELECT SUBSTRING (A.NumAtCard,1,3 )  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000'
		BEGIN
			  SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
			IF (SELECT SUBSTRING (A.NumAtCard,4, 3)  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000'
		BEGIN
			SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
		IF (SELECT SUBSTRING (A.NumAtCard,7,9 )  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000000000'
		BEGIN
			SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
	  END

  SET @CARDCODE = (SELECT A.CardCode  FROM OPCH A WHERE A.DocEntry = @DOCENTRY)
  SET @DOCUMENTO = (SELECT A.NumAtCard   FROM OPCH A WHERE A.DocEntry = @DOCENTRY)
  
  IF  (SELECT COUNT(*) FROM OPCH A WHERE A.CardCode = @CARDCODE AND A.NumAtCard = @DOCUMENTO ) > 1
  BEGIN
	 SET @ERROR = '---L(26) El numero de Referencia ya existe para el proveedor seleccionado ' ;
	 RETURN @ERROR;
  END

  /* verificar el numero de autorizacion*/
 IF (SELECT ISNULL(A.U_NO_AUTORI,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY)= '@' 
	BEGIN
 		SET @ERROR = '---L(27) Debe de ingresar un numero de autorizacion';
		RETURN @ERROR;
	END 
  ELSE
      BEGIN
		SET @AUTORI = dbo.SP_FNNumAutorizacion((SELECT ISNULL(A.U_NO_AUTORI,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY) )
			IF @AUTORI = 1 
				BEGIN
					SET @ERROR = '--L(28) Debe de ingresar un numero de autorizacion valido';
					RETURN @ERROR;
				END
			IF (SELECT ISNUMERIC(A.U_NO_AUTORI)  FROM OPCH  A WHERE A.DocEntry = @DOCENTRY ) = 0
				BEGIN
					SET @ERROR = '---L(29)Debe de ingresar un numero de autorizacion valido';
					RETURN @ERROR;
				END
	  END

/*FACTURA DE REEMBOLSO*/
	IF  (SELECT A.U_TI_COMPRO FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '41'
		BEGIN
		  	DECLARE ProdInfo CURSOR FOR SELECT B.U_T_IDENTIFICACION ,B.U_ID_PROVEEDOR,
		    B.U_PA_PROVEEDOR,B.U_T_PROVEEDOR,B.U_T_COMPROBANTE,
			B.U_SE_ESTABLE,B.U_PTO_EMISION,B.U_N_FACTURA,B.U_AUTO_REEMBOLSO ,
            B.U_FE_EMISION  
			 FROM PCH1 B WHERE B.DocEntry =  @DOCENTRY
			OPEN ProdInfo
			FETCH NEXT FROM ProdInfo INTO @TIPO_IDENTI,@ID_PROVEEDOR,@PAIS,@TIPOPROVEEDOR,
			@CODDOC,@ESTABLE,@PTOEMISION,@secuencialDocReembolso,@numeroautorizacionDocReemb,@FECHAEMISION
			WHILE @@fetch_status = 0
			BEGIN
				IF ISNULL(@TIPO_IDENTI,'@') = '@'
				BEGIN 
				   SET @ERROR = '---L(31)DEBE DE INGRESAR UN TIPO DE IDENTIFICAION PARA EL PROVEEDOR DE REEMBOLSO';
					RETURN @ERROR;
				END
				IF ISNULL(@ID_PROVEEDOR,'@') = '@'
				BEGIN 
				    SET @ERROR = '---L(31) DEBE DE INGRESAR UN ID PROVEEDOR PARA REEMBOLSO';
					RETURN @ERROR;
				END
				IF ISNULL(@PAIS,'@') = '@'
				BEGIN 
				    SET @ERROR = '---L(31) DEBE DE INGRESAR UN PAIS PARA EL PROVEEDOR DE REEMBOLSO';
					RETURN @ERROR;
				END
				IF ISNULL(@TIPOPROVEEDOR ,'@') = '@'
				BEGIN 
				    SET @ERROR = '---L(31) DEBE DE INGRESAR UN TIPO DE PROVEEDOR DE REEMBOLSO';
					RETURN @ERROR;
				END
				IF ISNULL(@CODDOC ,'@') = '@'
				BEGIN 
				    SET @ERROR = '---L(31) DEBE DE INGRESAR UN COD. DE DOCUMENTO DE REEMBOLSO';
					RETURN @ERROR;
				END
				IF ISNULL(@ESTABLE ,'@') = '@'
				BEGIN 
				   SET @ERROR = '---L(31) DEBE DE INGRESAR UNA SERIE DE ESTABLECIMIENTO PARA REEMBOLSO';
					RETURN @ERROR;
				END
------------------------------------------------------------------------------------------------
				IF ISNULL(@ESTABLE ,'@') != '@' AND @ESTABLE = '000'
				BEGIN 
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UNA SERIE DE ESTABLECIMIENTO VÁLIDO PARA EL REEMBOLSO';
				 RETURN @ERROR;
				END
------------------------------------------------------------------------------------------------
				IF ISNULL(@PTOEMISION  ,'@') = '@'
				BEGIN 
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UN PUNTO DE EMISION PARA EL REEMBOLSO';
				 RETURN @ERROR;
				END
---------------------------------------------------------------------------------------------------------------
				IF ISNULL(@PTOEMISION ,'@') != '@' AND @PTOEMISION = '000'
				BEGIN 
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UN PUNTO DE ESTABLECIMIENTO VÁLIDO PARA EL REEMBOLSO';
				 RETURN @ERROR;
				END
----------------------------------------------------------------------------------------------------------------
				IF ISNULL(@secuencialDocReembolso   ,'@') = '@'
				BEGIN 
				  SET @ERROR = '---L(31) DEBE DE INGRESAR UN NUMERO DE FACTURA';
				 RETURN @ERROR;
				END

				IF ISNULL(@numeroautorizacionDocReemb   ,'@') = '@'
				BEGIN 
				  SET @ERROR = '---L(31) DEBE DE INGRESAR UN NUMERO DE AUTORIZACION VALIDO PARA REEMBOLSO';
				 RETURN @ERROR;
				END
-----------------------------------------------------------------------------------------------------------
				IF ISNUMERIC (@numeroautorizacionDocReemb) = 0 
				begin
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UN NUMERO DE AUTORIZACION VALIDO PARA REEMBOLSO';
				 RETURN @ERROR;
				end
-----------------------------------------------------------------------------------------------------------
               IF LEN(@numeroautorizacionDocReemb) !=10 AND LEN(@numeroautorizacionDocReemb) !=37 AND LEN(@numeroautorizacionDocReemb) !=49 
			   BEGIN
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UN NUMERO DE AUTORIZACION VALIDO PARA REEMBOLSO';
				 RETURN @ERROR;
			   END 			  	
	
----------------------------------------------------------------------------------------------------------
				IF ISNULL(@FECHAEMISION ,'@') = '@'
				BEGIN 
				 SET @ERROR = '---L(31) DEBE DE INGRESAR UNA FECHA VALIDA PARA REEMBOLSO';
				 RETURN @ERROR;
				END
				FETCH NEXT FROM ProdInfo INTO @TIPO_IDENTI,@ID_PROVEEDOR,@PAIS,@TIPOPROVEEDOR,@CODDOC,@ESTABLE,@PTOEMISION,@secuencialDocReembolso,@numeroautorizacionDocReemb,@FECHAEMISION
		    END
		CLOSE ProdInfo
		DEALLOCATE ProdInfo
		END
		
	 RETURN @ERROR
	END

GO

CREATE FUNCTION [dbo].[FN_NOTAS_CREDITO_COMPRAS_VALIDACIONES](@DOCENTRY varchar(100))
RETURNS VARCHAR(200)
AS
/*
 DANIEL MORENO LOCALIZACION ECUATORIANA 
 VALIDACIONES PARA NOTAS DE CREDITO COMPRA.
 GUATEMALA 11/05/2017
 ONESOLUTIONS S.A.
*/
BEGIN
DECLARE @ERROR VARCHAR(200) = '0';
DECLARE @CARDCODE VARCHAR(30);
DECLARE @DOCUMENTO VARCHAR(30);
   
   IF (SELECT ISNULL(A.NumAtCard   ,'@')  FROM ORPC A WHERE A.DocEntry = @DOCENTRY  ) = '@'
				BEGIN
					SET @ERROR = '--L(80) Debe de ingresar un número de referencia ';
					RETURN @ERROR;
				END
   
   ELSE 
    BEGIN
	
	   IF LEN( (SELECT ISNULL(A.NumAtCard,'@')  FROM ORPC A WHERE A.DocEntry = @DOCENTRY  )) != 15
		BEGIN
					SET @ERROR = '--L(81) el numero de referencia debe ser de 15 numeros ';
					RETURN @ERROR;
		END
		 if isnumeric((SELECT A.NumAtCard   FROM ORPC A WHERE A.DocEntry = @DOCENTRY )) = 0
		 BEGIN
			SET  @ERROR = '---L(82) Solo digitos permitidos para el numero de referencia';
			RETURN @ERROR ;
		 END 
		  SET @CARDCODE = (SELECT A.CardCode  FROM ORPC A WHERE A.DocEntry = @DOCENTRY)
		  SET @DOCUMENTO = (SELECT A.NumAtCard   FROM ORPC A WHERE  A.DocEntry = @DOCENTRY )
		  IF (SELECT COUNT(A.NumAtCard) FROM ORPC A WHERE  A.CardCode = @CARDCODE AND A.NumAtCard = @DOCUMENTO  ) > 1 
			      BEGIN
					SET @ERROR = '--L(85) El numero de Referencia ya existe para el proveedor seleccionado';
					return @ERROR;
	               END
		  IF (SELECT SUBSTRING (A.NumAtCard,1,3 )  FROM ORPC A WHERE A.DocEntry = @DOCENTRY  ) = '000'
		BEGIN
			SET @ERROR = '--L(85) numero de referencia es incorrecto';
			return @ERROR;
		END

		IF (SELECT SUBSTRING (A.NumAtCard,4, 3)  FROM ORPC A WHERE A.DocEntry = @DOCENTRY ) = '000'
		BEGIN
			SET @ERROR = '--L(85) numero de referencia es incorrecto';
			return @ERROR;
		END
		IF (SELECT SUBSTRING (A.NumAtCard,7,9 )  FROM ORPC A WHERE A.DocEntry = @DOCENTRY  ) = '000000000'
		BEGIN
			SET @ERROR = '--L(85) numero de referencia es incorrecto';
			return @ERROR;
		END
	END 
		IF (SELECT ISNULL(A.U_D_MODIFICADO,'@')  FROM ORPC A WHERE A.DocEntry = @DOCENTRY ) = '@'
				BEGIN
				    SET @ERROR = 'L(83) Para la nota de crédito debe de selecionar una Doc. Modificado';
					RETURN @ERROR;
			    END
	 IF (SELECT A.DocTotal  FROM ORPC A WHERE A.DocEntry = @DOCENTRY  ) > (SELECT B.DocTotal  FROM OPCH B WHERE B.DocEntry = (SELECT A.U_D_MODIFICADO   FROM ORPC A WHERE A.DocEntry = @DOCENTRY))
		BEGIN
			SET @ERROR = '---L(84) Nota de crédito sobrepasa la cantidad de la factura '
			RETURN @ERROR
	    END	
		
		
RETURN @ERROR;
END

GO

CREATE  FUNCTION [dbo].[FN_COMPRAS_NOTAS_DEBITO_VALIDACIONES](@DOCENTRY varchar(100))
RETURNS VARCHAR(200)
AS
BEGIN
 /*
 DANIEL MORENO LOCALIZACION ECUATORIANA 
 VALIDACIONES PARA NOTAS DE DEBITO COMPRAS .
 GUATEMALA 11/05/2017
 ONESOLUTIONS S.A.
*/

DECLARE @ERROR VARCHAR(200) = '0';
SET @ERROR = '0';
DECLARE @CARDCODE VARCHAR(30)
DECLARE @DOCUMENTO VARCHAR(30)
DECLARE @AUTORI INT 

       IF (SELECT ISNULL(A.U_D_MODIFICADO,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY )= '@' 
	BEGIN				
				SET @ERROR = '---L(23) Debe de seleccionar un documento modificado ';
				RETURN @ERROR;
	END

	/*VALIDACIONES PARA NUMATCARD*/
	  IF (SELECT ISNULL(A.NumAtCard ,'@')  FROM OPCH A WHERE A.DocEntry = @DOCENTRY  ) = '@'
		 BEGIN
		  SET @ERROR = '---L(30) Debe de ingresar un numero de referencia';
		  RETURN @ERROR;
		 END
	  ELSE
	  BEGIN
		IF (SELECT len (A.NumAtCard)  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) != 15
		BEGIN
			  SET @ERROR = '---L(30) El numero de referencia no puede ser menor y mayor a 15';
		      RETURN @ERROR;
		END
		IF (SELECT SUBSTRING (A.NumAtCard,1,3 )  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000'
		BEGIN
			  SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
			IF (SELECT SUBSTRING (A.NumAtCard,4, 3)  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000'
		BEGIN
			SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
		IF (SELECT SUBSTRING (A.NumAtCard,7,9 )  FROM OPCH A WHERE A.DocEntry = @DOCENTRY ) = '000000000'
		BEGIN
			SET @ERROR = '---L(30) Numero de referencia es incorecto';
		      RETURN @ERROR;
		END
	  END

  SET @CARDCODE = (SELECT A.CardCode  FROM OPCH A WHERE A.DocEntry = @DOCENTRY)
  SET @DOCUMENTO = (SELECT A.NumAtCard   FROM OPCH A WHERE A.DocEntry = @DOCENTRY)
  
  IF  (SELECT COUNT(*) FROM OPCH A WHERE A.CardCode = @CARDCODE AND A.NumAtCard = @DOCUMENTO ) > 1
  BEGIN
	 SET @ERROR = '---L(26) El numero de Referencia ya existe para el proveedor seleccionado ' ;
	 RETURN @ERROR;
  END


RETURN @ERROR;
END 
GO


CREATE FUNCTION [dbo].[FN_VENTAS_VALIDACIONES](@DOCENTRY varchar(100))
RETURNS VARCHAR(200)
AS
/*
 DANIEL MORENO LOCALIZACION ECUATORIANA 
 VALIDACIONES PARA VENTAS.
 GUATEMALA 11/05/2017
 ONESOLUTIONS S.A.
*/
BEGIN
DECLARE @ERROR VARCHAR(200) = '0';

  DECLARE @TIPO_IDENTI AS nvarchar(400) 
		  DECLARE @ID_PROVEEDOR AS nvarchar(400)
		  DECLARE @PAIS AS nvarchar(400)
		  DECLARE @TIPOPROVEEDOR AS nvarchar(400) 
		  DECLARE @CODDOC AS nvarchar(400)
		  DECLARE @ESTABLE AS nvarchar(400)
		  DECLARE @PTOEMISION AS nvarchar(400)
		  DECLARE @secuencialDocReembolso AS nvarchar(400)
		  DECLARE @numeroautorizacionDocReemb AS nvarchar(400) 
		  DECLARE @FECHAEMISION AS nvarchar(400) 
		  DECLARE @MOTIVO AS NVARCHAR(100)
		  DECLARE @RUC AS NVARCHAR(25)

SET @ERROR = '0';
	
	IF (SELECT ISNULL(A.U_FORMA_PAGO   ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
		BEGIN
			SET @ERROR = '--L(50 ) Debe de selecciona una forma de pago';
			RETURN @ERROR;
		END

	IF (SELECT A.DocSubType   FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = 'DN'
		BEGIN
			IF (SELECT ISNULL(A.U_D_MODIFICADO ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
				BEGIN
					SET @ERROR = '--L(51) Debe de seleccionar un Doc. modificado';
					RETURN @ERROR;	
				END				
		END
/*CUANDO LA FACTURA ES DE EXPORTACION*/

			IF (SELECT A.DocSubType    FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = 'IX'
			BEGIN
			   IF (SELECT ISNULL(A.U_T_EXPORT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
						SET @ERROR = '--L(63) Debe de seleccionar un tipo de exportación';
						RETURN @ERROR;
				    END
				ELSE IF (SELECT ISNULL(A.U_T_EXPORT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '03'
					BEGIN 
						 IF (SELECT ISNULL(A.U_T_INGRE_EXT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de seleccionar un tipo de ingreso del exterior ';
								RETURN @ERROR;
							 END

						IF (SELECT ISNULL(A.U_T_INGRE_EXT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(65) Debe de seleccionar Impuesto a la Renta o Similar Ext ';
								RETURN @ERROR;
						END
						ELSE IF (SELECT ISNULL(A.U_T_INGRE_EXT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = 'SI'
							BEGIN 
								IF (SELECT ISNULL(A.U_T_EXPORT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
									BEGIN
										SET @ERROR = '--L(66) Debe ingresar valor par IR o similar en el Ext';
										RETURN @ERROR;
								    END
					END
					
				
				END
				ELSE IF (SELECT ISNULL(A.U_T_EXPORT,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '01'
					BEGIN
						IF (SELECT ISNULL(A.U_V_FOB ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de ingresar valor FOB ';
								RETURN @ERROR;
							 END
						 IF (SELECT ISNULL(A.U_D_ADUANERO  ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de ingresar distrito aduanero';
								RETURN @ERROR;
							 END
					      IF (SELECT ISNULL(A.U_ANO,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de ingresar distrito año';
								RETURN @ERROR;
							 END
						 IF (SELECT ISNULL(A.U_REGIMEN ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de ingresar Regimen';
								RETURN @ERROR;
							 END
					     IF (SELECT ISNULL(A.U_CORRELATIVO  ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
							BEGIN
								SET @ERROR = '--L(64) Debe de ingresar un correlativo';
								RETURN @ERROR;
							 END
					END

			   IF (SELECT ISNULL(A.U_INCO_TERM,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
						SET @ERROR = '--L(52) Debe de ingresar un Incor. Term';
						RETURN @ERROR;
				    END
				IF (SELECT ISNULL(A.U_LUGAR_INCOTERM,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					    SET @ERROR = '--L(53) Debe de ingresar un lugar incoTerm';
						RETURN @ERROR;
					END
				IF (SELECT ISNULL(A.U_PAIS_ORIGEN ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(54) Debe de ingresar lugar de origen';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_PUERTO_EMBARGUE  ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(55) Debe de ingresar un puerto de embarque';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_PUERTO_DESTINO,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(56) Debe de ingresar un puerto de destino';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_PAIS_DESTINO ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(57) Debe de un pais de destino';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_PAIS_ADQUISION  ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(58) Debe de un pais de adquisicion';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_TERM_TOT_SIN_IMPUESTO   ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(58) Debe de un inco. term. total sin impuesto';
					RETURN @ERROR;
				END

				IF (SELECT ISNULL(A.U_FLETE_INTERNA    ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(59) Debe un flete internacional';
					RETURN @ERROR;
				END

				IF (SELECT ISNULL(A.U_SEGURO_INTERNA ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(60) Debe ingresar seguro internacional';
					RETURN @ERROR;
				END
				IF (SELECT ISNULL(A.U_GASTOS_ADUANEROS ,'@')  FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '@'
					BEGIN
					SET @ERROR = '--L(61) Debe ingresar gastos aduaneros';
					RETURN @ERROR;
				END
				
			END

/*AL INGRESAR UN FACTURA DE REEMBOLSO DE VENTA */
IF  (SELECT A.U_TI_COMPRO    FROM OINV A WHERE A.DocEntry = @DOCENTRY ) = '41'
		BEGIN
				DECLARE ProdInfo CURSOR FOR SELECT B.U_T_IDENTIFICACION ,B.U_ID_PROVEEDOR,
				B.U_PA_PROVEEDOR,B.U_T_PROVEEDOR,B.U_T_COMPROBANTE,
				B.U_SE_ESTABLE,B.U_PTO_EMISION,B.U_N_FACTURA,B.U_AUTO_REEMBOLSO ,
				B.U_FE_EMISION,B.U_MOTIVO,B.U_N_RUC
			    FROM INV1 B WHERE B.DocEntry =  @DOCENTRY
			    OPEN ProdInfo
			    FETCH NEXT FROM ProdInfo INTO @TIPO_IDENTI,@ID_PROVEEDOR,@PAIS,@TIPOPROVEEDOR,@CODDOC,@ESTABLE,@PTOEMISION,@secuencialDocReembolso,@numeroautorizacionDocReemb,@FECHAEMISION,@MOTIVO,@RUC
			    WHILE @@fetch_status = 0
					BEGIN
					FETCH NEXT FROM ProdInfo INTO @TIPO_IDENTI,@ID_PROVEEDOR,@PAIS,@TIPOPROVEEDOR,@CODDOC,@ESTABLE,@PTOEMISION,@secuencialDocReembolso,@numeroautorizacionDocReemb,@FECHAEMISION,@MOTIVO,@RUC
						IF ISNULL(@secuencialDocReembolso   ,'@') = '@'
							BEGIN 
							 SET @ERROR = '--L(62) Debe de ingresar un numero de factura del reembolso';
							 RETURN @ERROR;
							END
						IF ISNULL(@FECHAEMISION ,'@') = '@'
							BEGIN 
							 SET @ERROR = '--L(63) DEBE DE INGRESAR UNA FECHA DE EMISION DEL REEMBOLSO';
							 RETURN @ERROR;
							END
						IF ISNULL(@numeroautorizacionDocReemb,'@') = '@'
							BEGIN 
						     SET @ERROR = '--L(64) DEBE DE INGRESAR UN NUMERO DE AUTORIZACION DE REEMBOLSO';
							 RETURN @ERROR;
							END
							IF ISNULL(@CODDOC ,'@') = '@'
								BEGIN 
								SET @ERROR = '--L(64) DEBE DE INGRESAR UN TIPO DE COMPROBANTE DEL DOCUMENTO DE REEMBOLSO';
								RETURN @ERROR;
								END

							IF ISNULL(@MOTIVO ,'@') = '@'
								BEGIN 
								SET @ERROR = '--L(64) DEBE DE INGRESAR UN MOTIVO PARA EL DOCUMENTO DE REEMBOLSO';
								RETURN @ERROR;
								END
                            IF ISNULL(@RUC ,'@') = '@'
								BEGIN 
								SET @ERROR = '--L(64) DEBE DE INGRESAR UN RUC PARA EL DOCUMENTO DE REEMBOLSO';
								RETURN @ERROR;
								END
					END
					CLOSE ProdInfo
			        DEALLOCATE ProdInfo

		END
RETURN @ERROR;
END 


GO

CREATE FUNCTION [dbo].[FN_SOCIOS_NEGOCIO_VALIDACIONES](@CARDCODE varchar(100))
RETURNS VARCHAR(200)
AS
/*
 DANIEL MORENO LOCALIZACION ECUATORIANA 
 VALIDACIONES PARA SOCIOS DE NEGOCIO.
 GUATEMALA 04/05/2017
 ONESOLUTIONS S.A.
*/
BEGIN
DECLARE @ERROR VARCHAR(200) = '0';
SET @ERROR = '0'
------------------------------------------
	IF ((SELECT ISNULL(OCRD.U_IDENTIFICACION,'@') from OCRD where OCRD.CardCode=@CARDCODE ) = '@') 
	BEGIN 
		SET @ERROR = '---L(1) Debe de seleccionar un tipo de identificacion';
		return @ERROR;
	END
	 IF  (select OCRD.U_IDENTIFICACION from OCRD where OCRD.CardCode=@CARDCODE) = '04'
		BEGIN
			IF ((SELECT OCRD.CardType  from OCRD where OCRD.CardCode=@CARDCODE ) != 'C') 
			BEGIN
				SET @ERROR = '---L(10) Debe ser cliente para consumidor final';
				RETURN @ERROR;
			END
			IF (select ISNULL(OCRD.U_DOCUMENTO ,'@')  from OCRD where OCRD.CardCode=@CARDCODE) != '9999999999999'
				BEGIN 
					SET @ERROR = '---L(4) Para consumidor final debe de ingresar 9999999999999';
					RETURN @ERROR;
				END  
		END
	IF ((select isnull(OCRD.U_PT_RELACIO ,'@') from OCRD where OCRD.CardCode=@CARDCODE AND OCRD.U_IDENTIFICACION !='04') = '@')
		BEGIN
			SET @ERROR = '---L(2) Debe de seleccionar parte relacionada';
			return @ERROR;
		END

	 IF  (select OCRD.U_IDENTIFICACION from OCRD where OCRD.CardCode=@CARDCODE) = '03'
		BEGIN
			IF (select ISNULL(OCRD.U_T_C_P,'@')  from OCRD where OCRD.CardCode=@CARDCODE) = '@'
				BEGIN 
					SET @ERROR = '---L(3) Debe de seleccionar un tipo cliente/proveedor';
					RETURN @ERROR;
				END  
		END

	
	 IF  (select ISNULL(OCRD.U_TIPO_CONTRI,'@')  from OCRD where OCRD.CardCode=@CARDCODE AND OCRD.U_IDENTIFICACION != '04' ) = '@'
		BEGIN
			SET @ERROR = '---L(5) Debe de seleccionar pago de residente o no residente';
			RETURN @ERROR;
		END

	

	IF (select ISNULL(OCRD.U_TIPO_CONTRI,'@')  from OCRD where OCRD.CardCode=@CARDCODE) = '02'
		BEGIN
			IF (select ISNULL(OCRD.U_T_R_FISCAL  ,'@')  from OCRD where OCRD.CardCode=@CARDCODE) = '@'
				BEGIN 
					SET @ERROR = '---L(6) Debe de seleccionar un regimen fiscal del exterior';
					RETURN @ERROR;
				END  
			else IF (select ISNULL(OCRD.U_D_TRIBUTACION   ,'@')  from OCRD where OCRD.CardCode=@CARDCODE) = '@'
				BEGIN 
					SET @ERROR = '---L(7) Debe de seleccionar si aplica convenio de doble tributacion';
					RETURN @ERROR;
				END  
		END

	IF (select OCRD.U_T_R_FISCAL  from OCRD where OCRD.CardCode=@CARDCODE ) = '01'
		BEGIN
			IF (select ISNULL( OCRD.U_P_R_FISCAL,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(8) Debe de seleccionar un pais de pago regimen general';
				RETURN @ERROR;
			END
			IF (select ISNULL( OCRD.U_PAIS_PAGO,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(9) Debe de seleccionar un pais de pago';
				RETURN @ERROR;
			END
		END
	   
	   	IF (select OCRD.U_T_R_FISCAL  from OCRD where OCRD.CardCode=@CARDCODE ) = '02'
		BEGIN
			IF (select ISNULL( OCRD.U_P_R_PFISCAL ,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(10) Debe de seleccionar pais de pago paraiso fiscal';
				RETURN @ERROR;
			END
			IF (select ISNULL( OCRD.U_PAIS_PAGO,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(11) Debe de seleccionar un pais de pago';
				RETURN @ERROR;
			END
		END
		IF (select OCRD.U_T_R_FISCAL  from OCRD where OCRD.CardCode=@CARDCODE ) = '03'
		BEGIN
			IF (select ISNULL( OCRD.U_D_R_FISCAL ,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(12) Debe de ingresar una denominacion del fegimen fiscal peferente';
				RETURN @ERROR;
			END
			IF (select ISNULL( OCRD.U_PAIS_PAGO,'@')   from OCRD where OCRD.CardCode=@CARDCODE) = '@'
			BEGIN
				SET @ERROR = '---L(9) Debe de seleccionar un pais de pago';
				RETURN @ERROR;
			END
		END
  
RETURN @ERROR;
END
GO