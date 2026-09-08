-- Migrazione per installazioni che hanno già eseguito CreazioneCampi.sql.
-- La regola funzionale è: codice Excel univoco per ditta e conto.

IF EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'UX_DESTDIV_dd_conto_dd_hhCodDestExc'
      AND object_id = OBJECT_ID('dbo.DESTDIV')
)
BEGIN
    DROP INDEX UX_DESTDIV_dd_conto_dd_hhCodDestExc ON dbo.DESTDIV;
END;
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'UX_DESTDIV_codditt_dd_conto_dd_hhCodDestExc'
      AND object_id = OBJECT_ID('dbo.DESTDIV')
)
BEGIN
    CREATE UNIQUE INDEX UX_DESTDIV_codditt_dd_conto_dd_hhCodDestExc
    ON dbo.DESTDIV (
        codditt,
        dd_conto,
        dd_hhCodDestExc
    )
    WHERE dd_hhCodDestExc IS NOT NULL;
END;
GO
