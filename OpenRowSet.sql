selec 

EXEC sp_configure 'show advanced options', 1
RECONFIGURE
GO
EXEC sp_configure 'ad hoc distributed queries', 1
RECONFIGURE
GO

create table exemplo (
coluna1 varchar(100),
coluna2 varchar(100),
coluna3 varchar(100))

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0' , N'AllowInProcess' , 1
GO

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0' , N'DynamicParameters' , 1
GO

INSERT INTO dbo.exemplo 

SELECT * FROM OPENROWSET('Microsoft.Jet.OLEDB.12.0','Excel 12.0;Database=C:\dosoft\exemplo.xls;HDR=YES', 'SELECT * FROM [Planilha1$]')

SELECT * FROM OPENROWSET('Microsoft.Jet.OLEDB.4.0',
  'Excel 8.0;Database=C:\dosoft\exemplo.xls', [Planilha1$])

icacls C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp /grant vs:(R,W)

SELECT * FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0',  
'Data Source=C:\dosoft\exemplo.xls;Extended Properties=EXCEL 5.0'),'SELECT * FROM [Planilha1$]')