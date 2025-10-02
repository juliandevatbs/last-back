USE SRLSQL;
 
GO
 
-- Bloqueamos el log asignando el modelo de recuperacion a Simple.
 
ALTER DATABASE SRLSQL
 
SET RECOVERY SIMPLE;
 
GO
 
-- Reducimos el archivo a 10 MB.
 
DBCC SHRINKFILE (SRLSQL_log, 100);
 
GO
 
-- reasignamos el modelo de recuperación.
 
ALTER DATABASE SRLSQL
 
SET RECOVERY FULL;
 
GO
 