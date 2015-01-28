# Convert Microsoft Access database to PostgreSQL dump files (VBA recipe)
by Daniel Clark <daniel@qcode.co.uk> [28/01/2015].

Use DAOs (Database Access Objects) to convert Microsoft Access database into Postgresql DDL statements to recreate database.

[MSAccessTables2PGDump](ms_access2pg_dump.bas#L71)
==================================================
Creates PostgreSQL dump file to recreate basic table datamodel.
Handles:
 * Conversion of MS Access data types to PostgreSQL data types.
 * Conversion of MS Access default values to PostgreSQL default values.
 * Not null constraints.
 * Column validation Rules (column check constraints).
 * Table validation Rules (table check constraints).
 * Primary key contraints.
 
note: check constraints may contain MS Acess specific functions and syntax which must be manually converted to PostgreSQL equivalents.

note: Boolean fields will default to false (if a default value is not specfied) and will enforce not null constraints.

note: Autonumber fields will enforce a not null constraint and will adopted the integer PostgreSQL data type.
 
[MSAccessIndexes2PGDump](ms_access2pg_dump.bas#L152)
====================================================
Creates PostgreSQL dump file to recreate indexes.

[MSAccessForeignKeys2PGDump](ms_access2pg_dump.bas#L217)
========================================================
Creates PostgreSQL dump file to recreate foreign key constraints.

note: If you MS Access Database performs case insenstive string matching you may have to sanitise your data before you are able to recreate foreign keys in PostgreSQL.

[MSAccessAutoNumbers2PGDump](ms_access2pg_dump.bas#L265)
========================================================
Creates PostgreSQL dump file to create sequences for MS Access autonumber fields and set default values.

note: Sequences will be initialised to the max value of the autonumber field + 1.

[MSAccessRecords2PGDump](ms_access2pg_dump.bas#L302)
====================================================
Creates PostgreSQL dump file to load record data


----------------------------------
Based on [Reverse engineer MS Access/Jet databases (Python recipe)](http://code.activestate.com/recipes/52267-reverse-engineer-ms-accessjet-databases/) by M.Keranen <mksql@yahoo.com> [07/12/2000].

----------------------------------
*[Qcode Software Limited] [qcode]*

[qcode]: http://www.qcode.co.uk "Qcode Software"
