CREATE PROCEDURE BaseConsolidada.PROJ_FEM()
    BEGIN 

        -- Definição de Cidade a partir no nome da Delagacia
        UPDATE BaseConsolidada.BaseFem
        SET CIDADE = 'RIBEIRAO PRETO'
        WHERE int64_field_0 = 1003;

        UPDATE BaseConsolidada.BaseFem
        SET CIDADE = 'RIBEIRAO PRETO'
        WHERE int64_field_0 = 1004;

        UPDATE BaseConsolidada.BaseFem
        SET CIDADE = 'ITAPETININGA'
        WHERE int64_field_0 = 322 or int64_field_0 = 323 or int64_field_0 = 324;

        UPDATE BaseConsolidada.BaseFem
        SET CIDADE = 'S.PAULO'
        WHERE int64_field_0 = 1363 or int64_field_0 = 1364;

        --Update estado civil considerando campo ignorado para os vazios
        UPDATE BaseConsolidada.BaseFem
        SET ESTADOCIVIL = 'Ignorado'
        WHERE ESTADOCIVIL IS NULL;

        --Escolaridade
        UPDATE BaseConsolidada.BaseFem
        SET GRAUINSTRUCAO = 'Não Informado'
        WHERE GRAUINSTRUCAO IS NULL;

        --Cor 
        UPDATE BaseConsolidada.BaseFem
        SET CORCUTIS = 'Não informada'
        WHERE CORCUTIS IS  NULL;

        --UF
        UPDATE BaseConsolidada.BaseFem
        SET UF = 'SP'
        WHERE UF IS  NULL;

    END;

--Criação e Insert Tabela resumida baseada na tabela principal
CREATE TABLE BaseConsolidada.FEMINICIDIO_FINAL(
    NUM_BO INT64, 
    DATAOCORRENCIA DATE,
    PERIDOOCORRENCIA STRING,
    BO_AUTORIA STRING,
    CIDADE STRING,
    UF STRING,
    DESCRICAOLOCAL STRING,
    RUBRICA STRING,
    STATUS STRING,
    NOMEPESSOA STRING,
    ESTADOCIVIL STRING,
    GRAUINSTRUCAO STRING,
    CORCUTIS STRING,
    NATUREZAVINCULADA STRING
);
INSERT BaseConsolidada.FEMINICIDIO_FINAL(

        NUM_BO,
        DATAOCORRENCIA,
        PERIDOOCORRENCIA,
        BO_AUTORIA,
        CIDADE,
        UF,
        DESCRICAOLOCAL,
        RUBRICA,
        STATUS,
        NOMEPESSOA,
        ESTADOCIVIL,
        GRAUINSTRUCAO,
        CORCUTIS,
        NATUREZAVINCULADA
)
SELECT  NUM_BO,
        DATAOCORRENCIA,
        PERIDOOCORRENCIA,
        BO_AUTORIA,
        CIDADE,
        UF,
        DESCRICAOLOCAL,
        RUBRICA,
        STATUS,
        NOMEPESSOA,
        ESTADOCIVIL,
        GRAUINSTRUCAO,
        CORCUTIS,
        NATUREZAVINCULADA
FROM BaseConsolidada.BaseFem;
    
    
call BaseConsolidada.PROJ_FEM();
