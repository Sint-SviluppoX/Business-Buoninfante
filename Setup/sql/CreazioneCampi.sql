-- 1. Creazione dei campi
ALTER TABLE dbo.DESTDIV
ADD
    dd_hhCodDestExc INT NULL,
    dd_hhGiornoConsegna VARCHAR(3) NULL;
GO


-- 2. Il codice, se valorizzato, deve essere positivo
ALTER TABLE dbo.DESTDIV
ADD CONSTRAINT CK_DESTDIV_dd_hhCodDestExc
CHECK (
    dd_hhCodDestExc IS NULL
    OR dd_hhCodDestExc > 0
);
GO


-- 3. Il codice deve essere univoco SOLO all'interno della stessa ditta e dello stesso dd_conto
CREATE UNIQUE INDEX UX_DESTDIV_codditt_dd_conto_dd_hhCodDestExc
ON dbo.DESTDIV (
    codditt,
    dd_conto,
    dd_hhCodDestExc
)
WHERE dd_hhCodDestExc IS NOT NULL;
GO


-- 4. Giorni della settimana ammessi
ALTER TABLE dbo.DESTDIV
ADD CONSTRAINT CK_DESTDIV_dd_hhGiornoConsegna
CHECK (
    dd_hhGiornoConsegna IS NULL
    OR dd_hhGiornoConsegna IN (
        'Lun',
        'Mar',
        'Mer',
        'Gio',
        'Ven',
        'Sab',
        'Dom'
    )
);
GO
