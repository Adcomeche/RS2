/*
  Diagnóstico de una subida relativa del 0,5 % de una tarifa Cepsa.
  Sintaxis: SQL Server 2008 o posterior (usa DATE y OUTER APPLY VALUES).
  El cliente también enumera SQL Server 2000: para ese servidor, utilizar
  la SELECT sencilla incluida en la respuesta, que devuelve h.* y d.*.
  Consultas de solo lectura. Sustituir los valores de ejemplo antes de ejecutarlas.

  La primera consulta muestra las cabeceras candidatas y TODOS los 15 tramos
  de cada fila kilométrica, incluso los vacíos. No reproduce íntegramente la
  prioridad de FindByDate ni las excepciones de LM, CE y Canarias.
*/
DECLARE @Fecha date = '2026-11-01';
DECLARE @Deposito varchar(3) = '030';
DECLARE @Emisor varchar(3) = '004';
DECLARE @TipoTarifa char(1) = 'A';       -- A: facturación activa; P: pasiva
DECLARE @TipoProducto varchar(3) = '010';
DECLARE @Destino varchar(10) = NULL;    -- NULL: mostrar todos los destinos
DECLARE @Cisterna varchar(3) = NULL;    -- NULL: mostrar todas las cisternas
DECLARE @Transportista varchar(3) = NULL; -- Solo suele discriminar la pasiva

SELECT
    h.TFBIDADE AS deposito,
    h.TFBIDTFB AS codigo_tarifa,
    h.TFBDESCR AS descripcion,
    h.TFBDTVIN AS vigente_desde,
    h.TFBFLTIT AS tipo_tarifa,
    h.TFBIDASC AS sociedad,
    h.TFBIDAEE AS emisor,
    h.TFBIDTTP AS tipologia_producto,
    h.TFBIDADZ AS destino_especifico,
    h.TFBIDATS AS tipo_cisterna,
    h.TFBIDAVE AS transportista,
    h.TFBFLMOT AS flag_vehiculo,
    h.TFBIDACN AS contrato_informativo,
    h.TFBFLUDM AS modalidad_kilometrica,
    h.TFBANNUL AS cabecera_anulada,
    d.TBRPROGD AS fila_km,
    d.TBRINFKM AS km_desde,
    d.TBRSUPKM AS km_hasta,
    d.TBRTIPOP AS tipo_recorrido,
    d.TBRANNUL AS detalle_anulado,
    t.numero_tramo,
    t.litros_desde,
    t.litros_hasta,
    t.tolerancia,
    t.moneda,
    t.precio_base AS precio_base_actual,
    CAST(CAST(t.precio_base AS decimal(28,10)) * 1.005 AS decimal(28,10)) AS precio_base_mas_0_5_pct,
    t.precio_distancia AS precio_distancia_actual,
    CAST(CAST(t.precio_distancia AS decimal(28,10)) * 1.005 AS decimal(28,10)) AS precio_distancia_mas_0_5_pct,
    CASE WHEN @Destino IS NOT NULL AND h.TFBIDADZ = @Destino THEN 'ESPECIFICO'
         WHEN h.TFBIDADZ IS NULL OR LEN(h.TFBIDADZ) = 0 THEN 'GENERAL'
         ELSE 'OTRO' END AS coincidencia_destino,
    CASE WHEN @Cisterna IS NOT NULL AND h.TFBIDATS = @Cisterna THEN 'ESPECIFICA'
         WHEN h.TFBIDATS IS NULL OR LEN(h.TFBIDATS) = 0 THEN 'GENERAL'
         ELSE 'OTRA' END AS coincidencia_cisterna,
    CASE WHEN @Transportista IS NOT NULL AND h.TFBIDAVE = @Transportista THEN 'COINCIDE'
         WHEN @Transportista IS NULL THEN 'SIN_FILTRO'
         ELSE 'OTRO' END AS coincidencia_transportista
FROM TARITFB0 AS h
LEFT JOIN TARITBR0 AS d
  ON d.TBRIDADE = h.TFBIDADE
 AND d.TBRIDTFB = h.TFBIDTFB
 AND d.TBRANNUL = 0
OUTER APPLY (VALUES
    ( 1, d.TBRLIN01, d.TBRLSU01, d.TBRTOL01, d.TBRIMP01, d.TBRIDV01, d.TBRIMB01),
    ( 2, d.TBRLIN02, d.TBRLSU02, d.TBRTOL02, d.TBRIMP02, d.TBRIDV02, d.TBRIMB02),
    ( 3, d.TBRLIN03, d.TBRLSU03, d.TBRTOL03, d.TBRIMP03, d.TBRIDV03, d.TBRIMB03),
    ( 4, d.TBRLIN04, d.TBRLSU04, d.TBRTOL04, d.TBRIMP04, d.TBRIDV04, d.TBRIMB04),
    ( 5, d.TBRLIN05, d.TBRLSU05, d.TBRTOL05, d.TBRIMP05, d.TBRIDV05, d.TBRIMB05),
    ( 6, d.TBRLIN06, d.TBRLSU06, d.TBRTOL06, d.TBRIMP06, d.TBRIDV06, d.TBRIMB06),
    ( 7, d.TBRLIN07, d.TBRLSU07, d.TBRTOL07, d.TBRIMP07, d.TBRIDV07, d.TBRIMB07),
    ( 8, d.TBRLIN08, d.TBRLSU08, d.TBRTOL08, d.TBRIMP08, d.TBRIDV08, d.TBRIMB08),
    ( 9, d.TBRLIN09, d.TBRLSU09, d.TBRTOL09, d.TBRIMP09, d.TBRIDV09, d.TBRIMB09),
    (10, d.TBRLIN10, d.TBRLSU10, d.TBRTOL10, d.TBRIMP10, d.TBRIDV10, d.TBRIMB10),
    (11, d.TBRLIN11, d.TBRLSU11, d.TBRTOL11, d.TBRIMP11, d.TBRIDV11, d.TBRIMB11),
    (12, d.TBRLIN12, d.TBRLSU12, d.TBRTOL12, d.TBRIMP12, d.TBRIDV12, d.TBRIMB12),
    (13, d.TBRLIN13, d.TBRLSU13, d.TBRTOL13, d.TBRIMP13, d.TBRIDV13, d.TBRIMB13),
    (14, d.TBRLIN14, d.TBRLSU14, d.TBRTOL14, d.TBRIMP14, d.TBRIDV14, d.TBRIMB14),
    (15, d.TBRLIN15, d.TBRLSU15, d.TBRTOL15, d.TBRIMP15, d.TBRIDV15, d.TBRIMB15)
) AS t(numero_tramo, litros_desde, litros_hasta, tolerancia, precio_base, moneda, precio_distancia)
WHERE h.TFBIDASC = '038'
  AND h.TFBIDADE = @Deposito
  AND h.TFBIDAEE = @Emisor
  AND h.TFBFLTIT = @TipoTarifa
  AND h.TFBIDTTP = @TipoProducto
  AND h.TFBDTVIN <= @Fecha
  AND h.TFBANNUL = 0
  AND (@Destino IS NULL OR h.TFBIDADZ = @Destino OR h.TFBIDADZ IS NULL OR LEN(h.TFBIDADZ) = 0)
  AND (@Cisterna IS NULL OR h.TFBIDATS = @Cisterna OR h.TFBIDATS IS NULL OR LEN(h.TFBIDATS) = 0)
ORDER BY h.TFBDTVIN DESC, h.TFBIDTFB, d.TBRPROGD, t.numero_tramo;

/*
  Segunda consulta: importes YA calculados. Ajustar el periodo y el depósito.
  Los importes PPD/PAD no se recalculan al editar TARITFB0/TARITBR0.
*/
DECLARE @Desde date = '2026-11-01';
DECLARE @HastaExclusivo date = '2026-12-01';
DECLARE @DepositoImpacto varchar(3) = '030';

SELECT
    ppt.PPTIDPPT AS prefactura,
    ppt.PPTFLTIP AS tipo_prefactura,
    ppt.PPTFLSTA AS estado_prefactura,
    p.PPDIDPPD AS linea,
    p.PPDDTVIA AS fecha_viaje,
    p.PPDIDVVT AS viaje,
    p.PPDIDVVC AS entrega,
    p.PPDIDVVP AS descarga,
    p.PPDIDACL AS cliente,
    p.PPDIDAEE AS emisor,
    p.PPDIDADE AS deposito,
    p.PPDIDADZ AS destino,
    p.PPDIDAPR AS producto,
    p.PPDIDTFI AS codigo_tarifa_interna,
    p.PPDQKMRI AS km_internos,
    p.PPDQTALT AS litros_descargados,
    p.PPDQTAKG AS kg_descargados,
    p.PPDVALTI AS tarifa_interna,
    p.PPDIMPVI AS importe_interno,
    p.PPDVALTA AS tarifa_cepsa,
    p.PPDIMPVA AS importe_cepsa,
    p.PPDFLANO AS mascara_anomalias,
    a.PADIDACN AS contrato_cargo,
    a.PADIDADD AS concepto_cargo,
    a.PADIMTAR AS tarifa_cargo,
    a.PADIMIMP AS importe_cargo,
    a.PADIDDAM AS documento_generado
FROM PREFPPT0 AS ppt
JOIN PREFPPD0 AS p ON p.PPDIDPPT = ppt.PPTIDPPT
LEFT JOIN PREFPAD0 AS a
  ON a.PADIDPPT = p.PPDIDPPT
 AND a.PADIDPPD = p.PPDIDPPD
WHERE p.PPDIDASC = '038'
  AND p.PPDIDADE = @DepositoImpacto
  AND p.PPDDTVIA >= @Desde
  AND p.PPDDTVIA < @HastaExclusivo
ORDER BY ppt.PPTIDPPT, p.PPDIDPPD, a.PADIDADD;
