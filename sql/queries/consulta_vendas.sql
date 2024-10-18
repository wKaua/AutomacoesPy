DECLARE @DiaAtual INT = DAY(GETDATE());

DECLARE @MesPassado DATE;
DECLARE @MesAtual DATE;

IF @DiaAtual = 1
BEGIN    
    SET @MesPassado = CONVERT(DATE, CONCAT(MONTH(DATEADD(MONTH, -2, GETDATE())), '/', 01, '/', YEAR(GETDATE())));
    SET @MesAtual = EOMONTH(DATEADD(MONTH, -1, GETDATE()))
END
ELSE
BEGIN
    SET @MesPassado = CONVERT(DATE, CONCAT(MONTH(DATEADD(MONTH, -1, GETDATE())), '/', 01, '/', YEAR(GETDATE())));
    SET @MesAtual = EOMONTH(GETDATE())
END

SELECT 
    CASE 
        WHEN v.cd_empresa IS NULL 
        THEN v.cd_laboratorio 
        ELSE v.cd_empresa 
    END AS CD_ENTIDADE,
    e.DS_FANTASIA, 
    CAST(CONCAT(MONTH(v.dt_venda), '/', 01, '/', YEAR(GETDATE())) AS DATE) AS DATA,
    COUNT(v.ds_etiqueta) AS VOLUME, 
    CASE 
        WHEN DATEDIFF(DAY, v.dt_venda, GETDATE()) <= 90
            AND e.dt_cadastro >= '2022-12-31' 
            AND v.ds_finalidade <> 'TERCEIRO'
        THEN 'VENDA NOVA'
        ELSE 'VENDA MANUTENÇÃO'
    END AS STATUS_VENDA,
    f.ds_fantasia AS EXECUTIVO
FROM powerbi.consultavendas v 
LEFT JOIN morales.dbo.tbl_entidades e 
    ON (CASE 
            WHEN v.cd_empresa IS NULL 
            THEN v.cd_laboratorio 
            ELSE v.cd_empresa
        END) = e.cd_entidade
LEFT JOIN powerbi.executivos f 
    ON e.cd_vendedor = f.cd_vendedor 
WHERE v.dt_venda BETWEEN @MesPassado AND @MesAtual
GROUP BY 
    CASE 
        WHEN v.cd_empresa IS NULL 
        THEN v.cd_laboratorio 
        ELSE v.cd_empresa 
    END, 
    e.DS_FANTASIA,
    CAST(CONCAT(MONTH(v.dt_venda), '/', 01, '/', YEAR(GETDATE())) AS DATE),
    f.ds_fantasia,
    CASE 
        WHEN DATEDIFF(DAY, v.dt_venda, GETDATE()) <= 90 
            AND e.dt_cadastro >= '2022-12-31'
            AND v.ds_finalidade <> 'TERCEIRO'
        THEN 'VENDA NOVA'
        ELSE 'VENDA MANUTENÇÃO'
    END
ORDER BY 
    COUNT(v.ds_etiqueta) DESC