CREATE PROCEDURE SAP_CP_CorteOficinasContratos

@FechaInicial AS DATETIME,
@FechaFinal AS DATETIME,
@OficinaVentas AS NVARCHAR(100),
@Usuario AS NVARCHAR(100)

AS	

DECLARE @date1 AS DATE
DECLARE @date2 AS DATE
DECLARE @valorcero AS FLOAT

SET @date1 = ( SELECT   MIN(TA.docdate)
               FROM     owtr TA
               WHERE    TA.docdate BETWEEN @FechaInicial AND @FechaFinal
             )
SET @date2 = ( SELECT   MAX(TB.docdate)
               FROM     owtr TB
               WHERE    TB.docdate BETWEEN  @FechaInicial AND @FechaFinal
             )

SET @valorcero = ( SELECT   0
                 )

SELECT  CAST(T0.DocNum AS NVARCHAR) AS "DocNum",
        @date1 AS "De",
        @date2 AS "Hasta",
        T0.DOCDATE,
        t1.u_numtraspaso AS "Num. Trasp Promotor - Oficinas",
        T1.u_CodPromotor AS "Promotor",
        T1.U_NombrePromotor AS "Asociado",
        T1.[Dscription] AS "Plan",
        T1.[U_Serie] AS "Folio",
        T1.[U_InvInicial],
        T1.[U_Comision],
        T1.[U_Importe],
        T1.[U_FormaPago],
        '' AS "Referencia",
        T5.[Name] AS "Origen",
        '' AS "Valor Origen",
        T6.WHSNAME AS "Oficina Contratos",
        T7.WHSNAME AS "Oficina de Ventas",
        T8.[middleName] + ' ' + T8.[lastName] AS "Usuario Oficina",
        CASE WHEN T1.U_FormaPago = 'EFECTIVO'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END EFECTIVO,
        CASE WHEN T1.U_FormaPago = 'DEPOSITO'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END DEPOSITO,
        CASE WHEN T1.U_FormaPago = 'PAGARE TARJETA BANCARIA'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END PAGARE,
        CASE WHEN T1.U_FormaPago = 'TRASNFERENCIA'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END TRANSFERENCIA,
        CASE WHEN T1.U_FormaPago = 'CHEQUE'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END CHEQUE,
        CASE WHEN T1.U_FormaPago = 'BONO PABS'
             THEN CAST(ISNULL(T1.U_Importe, 0) AS FLOAT)
             ELSE CAST(( SELECT @valorcero
                       ) AS FLOAT)
        END BONO
FROM    OWTR T0
        INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry
        INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode
        INNER JOIN OHEM T3 ON T0.SlpCode = T3.salesPrson
        INNER JOIN OUSR T4 ON T0.UserSign = T4.USERID
        LEFT JOIN [@ORIGSOLICITUD] T5 ON T1.U_OrigenSolicitud = t5.Code
        INNER JOIN OWHS T6 ON T1.WHSCODE = T6.WHSCODE
        INNER JOIN OWHS T7 ON T0.FILLER = T7.WHSCODE
        INNER JOIN OHEM T8 ON T0.USERSIGN = T8.USERID
WHERE   T0.U_tipoMov = 'OFICINAS - CONTRATOS'
        AND T0.DOCDATE BETWEEN  @FechaInicial AND @FechaFinal
        AND T0.Filler = @OficinaVentas
        AND T0.UserSign = @Usuario