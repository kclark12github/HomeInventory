VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCentura"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clsCentura - clsCentura.cls
'   Centura SQLBase SQLAPI Functions (Updated for version 7.0.1)...
'   Copyright © 2002, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   11/08/01    None        Ken Clark       Updated for v2.0;
'   10/26/01    None        Ken Clark       Corrected missed condition allowing NewDatabaseName to be appended to the end of the
'                                           database list in the SQL.INI file (by adding AddDBToPath() function to further
'                                           simplify the logic);
'                                           Removed FiRReMsgBox call if "SQL.INI Path" registry entry cannot be found;
'   10/25/01    None        Ken Clark       Corrected logic in AddDBToINI() to account for "*";
'   07/06/01    None        Ken Clark       Added AddDBtoINI(), DeleteDBfromINI(), and RemoveDBFromServerPath();
'                                           Added SQLBaseBackupLogs(), SQLBaseBackupSnapShot(), SQLBaseRestoreLogs() and
'                                           SQLBaseRestoreSnapShot();
'   02/04/01    None        Ken Clark       Moved GrantUser() into modDBEngine.bas;
'   01/25/01    None        Ken Clark       Added SQLBaseCreateDatabase() and SQLBaseDeleteDatabase() routines;
'   10/03/00    None        Ken Clark       Added error reporting and transaction capabilities to GrantUser;
'   08/11/99    None        JAD             IsInSQLBase to always return False until SQLBase bug is fixed
'                                           Added GrantUser function
'   07/13/99    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Private mConstants As SIASDB.clsSIASDBConstants

'Registry Keys (SQL.INI)...
Const KeyWS32 As String = "WIN32CLIENT.WS32"
Const KeyServerPath As String = "SERVERPATH"
Const SQLINIKeyName As String = "Software\Centura\SQLBase\Clients\7.0"
Const SQLINIKeyValue As String = "SQL.INI Path"

Const iOverwrite As Integer = 1
Const iLocal As Integer = 1

'information filter flags...
Const SQLRPNM = &H100    'process number
Const SQLRCLN = &H200    'client name
Const SQLRUSN = &H400    'user name
Const SQLRDBN = &H800    'database name
Const SQLXGSI = &H8000   'extended GSI information flag
'server information flags...
Const SQLGPWD = &H1      'send password
Const SQLGCUR = &H2      'cursor information
Const SQLGDBS = &H4      'database information
Const SQLGCFG = &H8      'configuration information
Const SQLGSTT = &H10     'statistics
Const SQLGPRC = &H20     'process information
'Error Text Constants...
Const SQLXMSG = &H1      'Message Text
Const SQLXREA = &H2      'Reason
Const SQLXREM = &H4      'Remedy

' Global definitions SQLAPI
' =========================

' Data Type Map
' -------------
' SQLTAPI ->      integer
' SQLTCUR ->      integer
' SQLTDAP ->      string
' SQLTDAL ->      integer
' SQLTSVH ->      integer
' SQLTFLG ->      integer
' SQLTPTY ->      integer

'---- Return Codes ----------
Private Const SQL_SUCCESS As Integer = 0
Private Const END_OF_FETCH As Integer = 1

'---- Program Data Types ----
Private Const SQLPBUF As Integer = 1  ' buffer
Private Const SQLPSTR As Integer = 2  ' string (zero terminated)
Private Const SQLPUCH As Integer = 3  ' unsigned char
Private Const SQLPSCH As Integer = 4  ' char
Private Const SQLPUIN As Integer = 5  ' unsigned integer
Private Const SQLPSIN As Integer = 6  ' integer
Private Const SQLPULO As Integer = 7  ' unsigned long
Private Const SQLPSLO As Integer = 8  ' long
Private Const SQLPFLT As Integer = 9  ' float
Private Const SQLPDOU As Integer = 10 ' double
Private Const SQLPNUM As Integer = 11 ' SQLBASE internal numeric format
Private Const SQLPDAT As Integer = 12 ' SQLBASE internal datetime format
Private Const SQLPUPD As Integer = 13 ' unsigned packed decimal
Private Const SQLPSPD As Integer = 14 ' packed decimal
Private Const SQLPDTE As Integer = 15 ' date only format
Private Const SQLPTIM As Integer = 16 ' time only format
Private Const SQLPUSH As Integer = 17 ' unsigned short
Private Const SQLPSSH As Integer = 18 ' short
Private Const SQLPNST As Integer = 19 ' numeric string
Private Const SQLPNBU As Integer = 20 ' numeric buffer
Private Const SQLPEBC As Integer = 21 ' EBCDIC buffer format
'------ Database Datatypes ------------
Private Const SQLDCHR As Integer = 1  ' character
Private Const SQLDNUM As Integer = 2  ' numeric
Private Const SQLDDAT As Integer = 3  ' date-time
Private Const SQLDLON As Integer = 4  ' long
Private Const SQLDDTE As Integer = 5  ' date only
Private Const SQLDTIM As Integer = 6  ' time only
'------ External Data Types ------------
Private Const SQLEINT As Integer = 1  ' integer
Private Const SQLESMA As Integer = 2  ' small integer
Private Const SQLEFLO As Integer = 3  ' float
Private Const SQLECHR As Integer = 4  ' character
Private Const SQLEVAR As Integer = 5  ' varchar
Private Const SQLELON As Integer = 6  ' long varchar
Private Const SQLEDEC As Integer = 7  ' decimal
Private Const SQLEDAT As Integer = 8  ' date only
Private Const SQLETIM As Integer = 9  ' time only
Private Const SQLETMS As Integer = 10 ' timestamp
Private Const SQLEMON As Integer = 11 ' money
Private Const SQLEDOU As Integer = 12 ' double
Private Const SQLEGPH As Integer = 13 ' graphic
Private Const SQLEVGP As Integer = 14 ' vargraphic
Private Const SQLELGP As Integer = 15 ' long vargraphic
Private Const SQLEBIN As Integer = 16 ' binary
Private Const SQLEVBI As Integer = 17 ' varbinary
Private Const SQLELBI As Integer = 18 ' long varbinary
Private Const SQLEBOO As Integer = 19 ' boolean
'------ Null indicators ----------------
Private Const Data_Contains_Nulls As Integer = -1
Private Const Data_Contains_No_Nulls As Integer = 0
Private Const SQLNULL As Integer = 0
'------ Bulk Execute Flags -------------
Private Const Bulk_Execute_Off As Integer = 0
Private Const Bulk_Execute_On As Integer = 1
'------ Fetch Status Codes -------------
Private Const FETRTRU As Integer = 1  ' data was truncated
Private Const FETRSIN As Integer = 2  ' signed number fetched into unsigned field
Private Const FETRDNN As Integer = 3  ' data was not numeric
Private Const FETRNOF As Integer = 4  ' numeric overflow
Private Const FETRDTN As Integer = 5  ' data type not supportted
Private Const FETRDND As Integer = 6  ' data is not a date
Private Const FETRNUL As Integer = 7  ' data is null

' SET and GET PARAMETER TYPES
'  Global parameters

Private Const SQLPDDB As Integer = 1            ' default database name
Private Const SQLPDUS As Integer = 2            ' default user name
Private Const SQLPDPW As Integer = 3            ' default password
Private Const SQLPGBC As Integer = 4            ' global cursor value
Private Const SQLPLRD As Integer = 5            ' local result set directory
Private Const SQLPDBM As Integer = 6            ' db mode - see below
Private Const SQLPDBD As Integer = 7            ' dbdir
Private Const SQLPCPG As Integer = 8            ' code page information
Private Const SQLPNIE As Integer = 9            ' null indicator error
Private Const SQLPCPT As Integer = 10           ' connect pass thru to backend
Private Const SQLPTPD As Integer = 11           ' temp dir
Private Const SQLPDTR As Integer = 12           ' distributed transaction mode
Private Const SQLPPSW As Integer = 15           ' server password
Private Const SQLPOOJ As Integer = 16           ' oracle outer join
Private Const SQLPNPF As Integer = 17           ' net prefix
Private Const SQLPNLG As Integer = 18           ' net log
Private Const SQLPNCT As Integer = 19           ' net check type
Private Const SQLPNCK As Integer = 20           ' net check
Private Const SQLPLCK As Integer = 22           ' locks
Private Const SQLPINT As Integer = 25           ' interrupt
Private Const SQLPERF As Integer = 27           '  error file
Private Const SQLPDIO As Integer = 28           '  direct I/O
Private Const SQLPSWR As Integer = 29           '  default write
Private Const SQLPCTY As Integer = 31           '  country
Private Const SQLPCSD As Integer = 32           '  commit server daemon
Private Const SQLPCSR As Integer = 33           '  commit server
Private Const SQLPCCK As Integer = 36           '  client check
Private Const SQLPCTS As Integer = 37           '  characterset
Private Const SQLPCGR As Integer = 38           '  cache group
Private Const SQLPAIO As Integer = 39           '  asyncio
Private Const SQLPANL As Integer = 40           '  apply net log
Private Const SQLPGRS As Integer = 41           '  get reentracy state
Private Const SQLPSTF As Integer = 42           '  set SQLTrace flags
Private Const SQLPCLG As Integer = 43           '  set commit-order logging

'  Server specific parameters

Private Const SQLPHEP As Integer = 1001         ' HEAP size for TSR executables
Private Const SQLPCAC As Integer = 1002         ' CACHE size in Kbytes
Private Const SQLPBRN As Integer = 1003         ' brand of database
Private Const SQLPVER As Integer = 1004         ' release version (ex. "4.0.J")
Private Const SQLPPRF As Integer = 1005         ' server profiling
Private Const SQLPPDB As Integer = 1006         ' partitioned database
Private Const SQLPGCM As Integer = 1007         ' group commit count
Private Const SQLPGCD As Integer = 1008         ' group commit delay ticks
Private Const SQLPDLK As Integer = 1009         ' number of deadlocks
Private Const SQLPCTL As Integer = 1010         ' command time limit
Private Const SQLPAPT As Integer = 1011         ' process timer activated
Private Const SQLPOSR As Integer = 1012         ' OS sample rate
Private Const SQLPAWS As Integer = 1013         ' OS Averaging window size
Private Const SQLPWKL As Integer = 1014         ' Work Limit
Private Const SQLPWKA As Integer = 1015         ' Work Space allocation
Private Const SQLPUSR As Integer = 1016         ' Number of users
Private Const SQLPTMO As Integer = 1017         ' time out
Private Const SQLPTSS As Integer = 1018         ' thread stack size
Private Const SQLPTHM As Integer = 1019         ' thread mode
Private Const SQLPSTC As Integer = 1020         ' sortcache size in kilobytes
Private Const SQLPSIL As Integer = 1021         ' silent mode
Private Const SQLPSPF As Integer = 1022         ' server prefix
Private Const SQLPSVN As Integer = 1024         ' server name
Private Const SQLPROM As Integer = 1025         ' read-only mode (0 or 1)
Private Const SQLPSTA As Integer = 1026         ' enable stats gathering
Private Const SQLPCSV As Integer = 1027         ' commit server
Private Const SQLPTTP As Integer = 1028         ' trace for 2PC

'  Database specific parameters

Private Const SQLPDBN As Integer = 2001         ' database name
Private Const SQLPDDR As Integer = 2002         ' database directory
Private Const SQLPLDR As Integer = 2003         ' log directory
Private Const SQLPLFS As Integer = 2004         ' log file size in Kbytes
Private Const SQLPCTI As Integer = 2005         ' checkpoint time interval in mins
Private Const SQLPLBM As Integer = 2006         ' log backup mode? (0 or 1)
Private Const SQLPPLF As Integer = 2007         ' Pre-allocate log files? (0 or 1)
Private Const SQLPTSL As Integer = 2008         ' transaction span limit
Private Const SQLPROT As Integer = 2009         ' read-only transactions (0, 1, 2)
Private Const SQLPHFS As Integer = 2010         ' history file size in Kbytes
Private Const SQLPREC As Integer = 2011         ' recovery
Private Const SQLPEXE As Integer = 2012         ' name of executable
Private Const SQLPNLB As Integer = 2013         ' next log to backup
Private Const SQLPROD As Integer = 2014         ' read-only database (0 or 1)
Private Const SQLPEXS As Integer = 2015         ' database file extension size
Private Const SQLPPAR As Integer = 2016         ' partitioned database (0 or 1)
Private Const SQLPNDB As Integer = 2017         ' NEWDB
Private Const SQLPLGF As Integer = 2018         ' log file offset
Private Const SQLPDTL As Integer = 2019         ' command timelimit
Private Const SQLPSMN As Integer = 2020         ' show main db
Private Const SQLPCVC As Integer = 2021         ' catalog version counter
Private Const SQLPDBS As Integer = 2022         ' database block size
Private Const SQLPUED As Integer = 2023         ' update external dictionary

'  Cursor specific parameters
'------ Get/Set Parameters -----------------

Private Const SQLPISO As Integer = 3001         ' isolation level (SQLILRR etc..)
Private Const SQLPWTO As Integer = 3002         ' lock wait timeout in seconds
Private Const SQLPPCX As Integer = 3003         ' preserve context (0 or 1)
Private Const SQLPFRS As Integer = 3004         ' front end result sets
Private Const SQLPLDV As Integer = 3005         ' load version (ex. "3.6.22")
Private Const SQLPAUT As Integer = 3006         ' autocommit
Private Const SQLPRTO As Integer = 3007         ' rollback trans on lock timeout
Private Const SQLPSCR As Integer = 3008         ' scroll mode (0 or 1)
Private Const SQLPRES As Integer = 3009         ' restriction mode (0 or 1)
Private Const SQLPFT  As Integer = 3010         ' fetch through
Private Const SQLPNPB As Integer = 3011         ' no pre-build in RL mode
Private Const SQLPPWD As Integer = 3012         ' current password
Private Const SQLPDB2 As Integer = 3013         ' DB2 compatibility mode
Private Const SQLPREF As Integer = 3014         ' referential integrity checking
Private Const SQLPBLK As Integer = 3015         ' bulk-execute mode
Private Const SQLPOBL As Integer = 3016         ' optimized bulk-execute mode
Private Const SQLPLFF As Integer = 3017         ' LONG data allowed in FERS
Private Const SQLPDIS As Integer = 3018         ' When to return Describe info
Private Const SQLPCMP As Integer = 3019         ' Compress messages sent to server
Private Const SQLPCHS As Integer = 3020         ' chained cmd has SELECT (0 or 1)
Private Const SQLPOPL As Integer = 3021         ' optimizer level
Private Const SQLPRID As Integer = 3022         ' ROWID
Private Const SQLPEMT As Integer = 3023         ' Error Message Tokens
Private Const SQLPCLN As Integer = 3024         ' client name
Private Const SQLPLSS As Integer = 3025         ' last compiled SQL statement
Private Const SQLPEXP As Integer = 3026         ' explain query plan
Private Const SQLPCXP As Integer = 3027         ' cost of execution plan
Private Const SQLPOCL As Integer = 3028         ' optimizercostlevel
Private Const SQLPTST As Integer = 3029          ' distributed transaction status
Private Const SQLP2PP As Integer = 3030         ' 2-phase protocol (SQL2STD, etc.)
'  defined for Load/Unload parsed parameters - cursor specific
Private Const SQLPCLI As Integer = 3031         ' ON CLIENT option
Private Const SQLPFNM As Integer = 3032         ' load/unload file name
Private Const SQLPOVR As Integer = 3033         ' file OVERWRITE flag
Private Const SQLPTFN As Integer = 3034         ' A Temporary file name
Private Const SQLPTRC As Integer = 3035         ' Trace stored procedures
Private Const SQLPTRF As Integer = 3036         ' Tracefile for stored procedures
Private Const SQLPCTF As Integer = 3037         ' control file flag

Private Const SQLPMID As Integer = 3038         ' mail id
Private Const SQLPAID As Integer = 3039         ' adapter id
Private Const SQLPNID As Integer = 3040         ' network id
Private Const SQLPUID As Integer = 3041         ' user application id
Private Const SQLPCIS As Integer = 3042         ' client identification strings
Private Const SQLPIMB As Integer = 3043         ' input message buffer size
Private Const SQLPOMB As Integer = 3044         ' output message buffer size
Private Const SQLPWFC As Integer = 3045         ' which fetchable command
Private Const SQLPRFE As Integer = 3046         ' Return on First Error-bulk insert
Private Const SQLPCUN As Integer = 3047         ' Current cursor user name
Private Const SQLPOFF As Integer = 3048         ' Optimize First Fetch

'   Application Specific ie. applicable to all cursors that
'    belong to the same application  (3700 -  3799)

Private Const SQLPCCB As Integer = 3700         ' Connect Closure Behaviour
Private Const SQLPTTV As Integer = 3701         ' Thread Timeout Value

'  Static attributes

Private Const SQLPFAT As Integer = 4000         ' first attribute
Private Const SQLPBRS As Integer = 4001         ' back end result sets
Private Const SQLPMUL As Integer = 4002         ' multi-user
Private Const SQLPDMO As Integer = 4003         ' demonstration version
Private Const SQLPLOC As Integer = 4004         ' local version of database
Private Const SQLPFPT As Integer = 4005         ' 1st participant
Private Const SQLPLAT As Integer = 4006         ' last attribute
Private Const SQLPCAP As Integer = 4007         ' API capability level
Private Const SQLPSCL As Integer = 4008         ' server capability level
Private Const SQLPRUN As Integer = 4009         ' runtime version

'  Server specific parameters

Private Const SQLPPLV As Integer = 5001         ' print level
Private Const SQLPALG As Integer = 5002         ' activity log
Private Const SQLPTMS As Integer = 5003         ' time stamp
Private Const SQLPPTH As Integer = 5004         ' path name seperator
Private Const SQLPTMZ As Integer = 5005         ' time zone
Private Const SQLPTCO As Integer = 5006         ' time colon only

'  SQL Server & Router/Gateway specific parameters

Private Const SQLPESTR As Integer = 5101        ' get server error # and string
Private Const SQLPMSTR As Integer = 5102        ' get server msg# and string
Private Const SQLPMAPC As Integer = 5103        ' MapGTICursors
Private Const SQLPUPRE As Integer = 5104        ' get user prefix
Private Const SQLPORID As Integer = 5105        ' Oracle RowID
Private Const SQLPERRM As Integer = 5106        ' error mapping
Private Const SQLPRTS  As Integer = 5107        ' SQL Server - Return Status
Private Const SQLPSAUT As Integer = 5108        ' SQL Server - Autocommit
Private Const SQLPROW  As Integer = 5109        ' SQL Server - Command Type
Private Const SQLPEHC  As Integer = 5110        ' SQL Server - Enhanced Cursors
Private Const SQLPGFS  As Integer = 5111        ' SQL Server - Get Fetch Status
Private Const SQLPLBUF As Integer = 5112        ' Longbuffer setting
Private Const SQLPDPH  As Integer = 5113        ' SQL Server - DBProcess handle
Private Const SQLPCKE  As Integer = 5114        ' SQL Server - CHECK EXISTS
Private Const SQLPWTX  As Integer = 5115        ' SQL Server - DBWRITETEXT
Private Const SQLPYLD  As Integer = 5116        ' SQL Server - YieldOnServerCall
Private Const SQLPOBN  As Integer = 5117        ' ODBC Router - backend brand
Private Const SQLPOBV  As Integer = 5118        ' ODBC Router - backend version
Private Const SQLPODN  As Integer = 5119        ' ODBC Router - driver name
Private Const SQLPODV  As Integer = 5120        ' ODBC Router - driver version
Private Const SQLPOCV  As Integer = 5121        ' ODBC Router - ODBC version
Private Const SQLPRSYS As Integer = 5122        ' DRDA - EXEC SQL CONNECT TO remote system name
Private Const SQLPLAB  As Integer = 5123        ' DB2 - return label information if exists
Private Const SQLPCID  As Integer = 5124        ' DB2 - Set Current SQLID default
Private Const SQLPNUMST As Integer = 5125        ' AS/400 Number of Statements
Private Const SQLPBNDRW As Integer = 5126       ' Oracle- bind SQLPBUF as RAW
Private Const SQLPNLS  As Integer = 5127        ' Informix - NLS database

Private Const SQLPFRW As Integer = 5200         ' fetchrow

'  Sybase System 10 parameters (reserved 5201 - 5225)

Private Const SQLPNESTR As Integer = 5201       ' SYB - get next error from client
Private Const SQLPNMSTR As Integer = 5202       ' SYB - get next error from server
Private Const SQLPCESTR As Integer = 5203       ' SYB - get client message count
Private Const SQLPCMSTR As Integer = 5204       ' SYB - get server message count
Private Const SQLPBRW As Integer = 5226         ' buffrow

'  ODBC specific parameters - Refer to ODBC spec for definition

Private Const SQLP_ACTIVE_CONNECTIONS As Integer = 5500
Private Const SQLP_ACTIVE_STATEMENTS As Integer = 5501
Private Const SQLP_DATA_SOURCE_NAME As Integer = 5502
Private Const SQLP_DRIVER_HDBC As Integer = 5503
Private Const SQLP_DRIVER_HENV As Integer = 5504
Private Const SQLP_DRIVER_HSTMT As Integer = 5505
Private Const SQLP_DRIVER_NAME As Integer = 5506
Private Const SQLP_DRIVER_VER As Integer = 5507
Private Const SQLP_FETCH_DIRECTION As Integer = 5508
Private Const SQLP_ODBC_API_CONFORMANCE As Integer = 5509
Private Const SQLP_ODBC_VER As Integer = 5510
Private Const SQLP_ROW_UPDATES As Integer = 5511
Private Const SQLP_ODBC_SAG_CLI_CONFORMANCE As Integer = 5512
Private Const SQLP_SERVER_NAME As Integer = 5513
Private Const SQLP_SEARCH_PATTERN_ESCAPE As Integer = 5514
Private Const SQLP_ODBC_SQL_CONFORMANCE As Integer = 5515
Private Const SQLP_DATABASE_NAME As Integer = 5516
Private Const SQLP_DBMS_NAME As Integer = 5517
Private Const SQLP_DBMS_VER As Integer = 5518
Private Const SQLP_ACCESSIBLE_TABLES As Integer = 5519
Private Const SQLP_ACCESSIBLE_PROCEDURES As Integer = 5520
Private Const SQLP_PROCEDURES As Integer = 5521
Private Const SQLP_CONCAT_NULL_BEHAVIOUR As Integer = 5522
Private Const SQLP_CURSOR_COMMIT_BEHAVIOUR As Integer = 5523
Private Const SQLP_CURSOR_ROLLBACK_BEHAVIOUR As Integer = 5524
Private Const SQLP_DATA_SOURCE_READ_ONLY As Integer = 5525
Private Const SQLP_DEFAULT_TXN_ISOLATION As Integer = 5526
Private Const SQLP_EXPRESSIONS_IN_ORDERBY As Integer = 5527
Private Const SQLP_IDENTIFIER_CASE As Integer = 5528
Private Const SQLP_IDENTIFIER_QUOTE_CHAR As Integer = 5529
Private Const SQLP_MAX_COLUMN_NAME_LEN As Integer = 5530
Private Const SQLP_MAX_CURSOR_NAME_LEN As Integer = 5531
Private Const SQLP_MAX_OWNER_NAME_LEN As Integer = 5532
Private Const SQLP_MAX_PROCEDURE_NAME_LEN As Integer = 5533
Private Const SQLP_MAX_QUALIFIER_NAME_LEN As Integer = 5534
Private Const SQLP_MAX_TABLE_NAME_LEN As Integer = 5535
Private Const SQLP_MULT_RESULT_SETS As Integer = 5536
Private Const SQLP_MULTIPLE_ACTIVE_TXN As Integer = 5537
Private Const SQLP_OUTER_JOINS As Integer = 5538
Private Const SQLP_OWNER_TERM As Integer = 5539
Private Const SQLP_PROCEDURE_TERM As Integer = 5540
Private Const SQLP_QUALIFIER_NAME_SEPARATOR As Integer = 5541
Private Const SQLP_QUALIFIER_TERM As Integer = 5542
Private Const SQLP_SCROLL_CONCURRENCY As Integer = 5543
Private Const SQLP_SCROLL_OPTIONS As Integer = 5544
Private Const SQLP_TABLE_TERM As Integer = 5545
Private Const SQLP_TXN_CAPABLE As Integer = 5546
Private Const SQLP_USER_NAME As Integer = 5547
Private Const SQLP_CONVERT_FUNCTIONS As Integer = 5548
Private Const SQLP_NUMERIC_FUNCTIONS As Integer = 5549
Private Const SQLP_STRING_FUNCTIONS As Integer = 5550
Private Const SQLP_SYSTEM_FUNCTIONS As Integer = 5551
Private Const SQLP_TIMEDATE_FUNCTIONS As Integer = 5552
Private Const SQLP_CONVERT_BIGINT As Integer = 5553
Private Const SQLP_CONVERT_BINARY As Integer = 5554
Private Const SQLP_CONVERT_BIT As Integer = 5555
Private Const SQLP_CONVERT_CHAR As Integer = 5556
Private Const SQLP_CONVERT_DATE As Integer = 5557
Private Const SQLP_CONVERT_DECIMAL As Integer = 5558
Private Const SQLP_CONVERT_DOUBLE As Integer = 5559
Private Const SQLP_CONVERT_FLOAT As Integer = 5560
Private Const SQLP_CONVERT_INTEGER As Integer = 5561
Private Const SQLP_CONVERT_LONGVARCHAR As Integer = 5562
Private Const SQLP_CONVERT_NUMERIC As Integer = 5563
Private Const SQLP_CONVERT_REAL As Integer = 5564
Private Const SQLP_CONVERT_SMALLINT As Integer = 5565
Private Const SQLP_CONVERT_TIME As Integer = 5566
Private Const SQLP_CONVERT_TIMESTAMP As Integer = 5567
Private Const SQLP_CONVERT_TINYINT As Integer = 5568
Private Const SQLP_CONVERT_VARBINARY As Integer = 5569
Private Const SQLP_CONVERT_VARCHAR As Integer = 5570
Private Const SQLP_CONVERT_LONGVARBINARY As Integer = 5571
Private Const SQLP_TXN_ISOLATION_OPTION As Integer = 5572
Private Const SQLP_ODBC_SQL_OPT_IEF As Integer = 5573

'** ODBC SDK 1.0 Additions **
Private Const SQLP_CORRELATION_NAME As Integer = 5574
Private Const SQLP_NON_NULLABLE_COLUMNS As Integer = 5575

'** ODBC SDK 2.0 Additions **
Private Const SQLP_DRIVER_HLIB As Integer = 5576
Private Const SQLP_DRIVER_ODBC_VER As Integer = 5577
Private Const SQLP_FILE_USAGE As Integer = 5578
Private Const SQLP_GETDATA_EXTENSIONS As Integer = 5579
Private Const SQLP_NULL_COLLATION As Integer = 5580
Private Const SQLP_ALTER_TABLE As Integer = 5581
Private Const SQLP_COLUMN_ALIAS As Integer = 5582
Private Const SQLP_GROUP_BY As Integer = 5583
Private Const SQLP_KEYWORDS As Integer = 5584
Private Const SQLP_ORDER_BY_COLUMNS_IN_SELECT As Integer = 5585
Private Const SQLP_OWNER_USAGE As Integer = 5586
Private Const SQLP_POSITIONED_STATEMENTS As Integer = 5587
Private Const SQLP_QUALIFIER_USAGE As Integer = 5588
Private Const SQLP_QUOTED_IDENTIFIER_CASE As Integer = 5589
Private Const SQLP_SPECIAL_CHARACTERS As Integer = 5590
Private Const SQLP_SUBQUERIES As Integer = 5591
Private Const SQLP_UNION As Integer = 5592
Private Const SQLP_MAX_COLUMNS_IN_GROUP_BY As Integer = 5593
Private Const SQLP_MAX_COLUMNS_IN_INDEX As Integer = 5594
Private Const SQLP_MAX_COLUMNS_IN_ORDER_BY As Integer = 5595
Private Const SQLP_MAX_COLUMNS_IN_SELECT As Integer = 5596
Private Const SQLP_MAX_COLUMNS_IN_TABLE As Integer = 5597
Private Const SQLP_MAX_INDEX_SIZE As Integer = 5598
Private Const SQLP_MAX_ROW_SIZE_INCLUDES_LONG As Integer = 5599
Private Const SQLP_MAX_ROW_SIZE As Integer = 5600
Private Const SQLP_MAX_STATEMENT_LEN As Integer = 5601
Private Const SQLP_MAX_TABLES_IN_SELECT As Integer = 5602
Private Const SQLP_MAX_USER_NAME_LEN As Integer = 5603
Private Const SQLP_TIMEDATE_ADD_INTERVALS As Integer = 5604
Private Const SQLP_TIMEDATE_DIFF_INTERVALS As Integer = 5605
Private Const SQLP_CONVERT_C_BINARY As Integer = 5606
Private Const SQLP_CONVERT_C_BIT As Integer = 5607
Private Const SQLP_CONVERT_C_CHAR As Integer = 5608
Private Const SQLP_CONVERT_C_DATE As Integer = 5609
Private Const SQLP_CONVERT_C_DOUBLE As Integer = 5610
Private Const SQLP_CONVERT_C_FLOAT As Integer = 5611
Private Const SQLP_CONVERT_C_SLONG As Integer = 5612
Private Const SQLP_CONVERT_C_SSHORT As Integer = 5613
Private Const SQLP_CONVERT_C_STINYINT As Integer = 5614
Private Const SQLP_CONVERT_C_TIME As Integer = 5615
Private Const SQLP_CONVERT_C_TIMESTAMP As Integer = 5616
Private Const SQLP_CONVERT_C_ULONG As Integer = 5617
Private Const SQLP_CONVERT_C_USHORT As Integer = 5618
Private Const SQLP_CONVERT_C_UTINYINT As Integer = 5619
Private Const SQLP_CONVERT_SQL_BIGINT As Integer = 5620
Private Const SQLP_CONVERT_SQL_BINARY As Integer = 5621
Private Const SQLP_CONVERT_SQL_BIT As Integer = 5622
Private Const SQLP_CONVERT_SQL_CHAR As Integer = 5623
Private Const SQLP_CONVERT_SQL_DATE As Integer = 5624
Private Const SQLP_CONVERT_SQL_DECIMAL As Integer = 5625
Private Const SQLP_CONVERT_SQL_DOUBLE As Integer = 5626
Private Const SQLP_CONVERT_SQL_FLOAT As Integer = 5627
Private Const SQLP_CONVERT_SQL_INTEGER As Integer = 5628
Private Const SQLP_CONVERT_SQL_LONGVARBINARY As Integer = 5629
Private Const SQLP_CONVERT_SQL_LONGVARCHAR As Integer = 5630
Private Const SQLP_CONVERT_SQL_NUMERIC As Integer = 5631
Private Const SQLP_CONVERT_SQL_REAL As Integer = 5632
Private Const SQLP_CONVERT_SQL_SMALLINT As Integer = 5633
Private Const SQLP_CONVERT_SQL_TIME As Integer = 5634
Private Const SQLP_CONVERT_SQL_TIMESTAMP As Integer = 5635
Private Const SQLP_CONVERT_SQL_TINYINT As Integer = 5636
Private Const SQLP_CONVERT_SQL_VARBINARY As Integer = 5637
Private Const SQLP_CONVERT_SQL_VARCHAR As Integer = 5638

Private Const SQLP_GET_TYPE_INFO As Integer = 5699

' Global definitions SQLAPI

' SQL/API Data Type Map     VB Data Type    Type declaration character
' SQLTAPI ->                integer         %
' SQLTCUR ->                integer         %
' SQLTDAP ->                string          $
' SQLTDAL ->                integer         %
' SQLTSVH ->                integer         %
' SQLTFLG ->                integer         %
' SQLTPTY ->                integer         %
' SQLTMOD ->                long            &
' SQLTCON ->                long            &
'------ Structures ---------------------
' gdidef is used by sqlgdi - Get Describe Information
Private Type gdidef
  gdichb As String    ' Column heading
  gdichl As Integer   ' Column heading length
  gdilbb As String    ' Label
  gdilbl As Integer   ' Label length
  gdicol As Integer   ' Column number
  gdiddt As Integer   ' data type
  gdiddl As Integer   ' data length
  gdiedt As Integer   ' external data type
  gdiedl As Integer   ' external data length
  gdipre As Integer   ' precision
  gdisca As Integer   ' scale
  gdinul As Integer   ' null indicator
  gdifil As String    ' reserved
End Type

'Kernel32...
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'SQLBase...
'------ Function Declarations ----------
Private Declare Function sqlcnc Lib "sqlwntm" (ByRef hcur As Integer, ByVal dbString As String, ByVal stringlen As Integer) As Integer
Private Declare Function sqlrel Lib "sqlwntm" (ByRef hcur As Integer) As Integer
Private Declare Function sqlcsv Lib "sqlwntm" (ByRef hServer As Integer, ByVal serverid As String, ByVal Password As String) As Integer
Private Declare Function sqlbdb Lib "sqlwntm" (ByRef hServer As Integer, ByVal DatabaseName As String, ByVal DatabaseNameLen As Integer, ByVal BackupDir As String, ByVal BackupDirLen As Integer, ByVal fLocal As Integer, ByVal fOverwrite As Integer) As Integer
Private Declare Function sqlblf Lib "sqlwntm" (ByRef hServer As Integer, ByVal DatabaseName As String, ByVal DatabaseNameLen As Integer, ByVal BackupDir As String, ByVal BackupDirLen As Integer, ByVal fLocal As Integer, ByVal fOverwrite As Integer) As Integer
Private Declare Function sqlrlf Lib "sqlwntm" (ByRef hServer As Integer, ByVal DatabaseName As String, ByVal DatabaseNameLen As Integer, ByVal BackupDir As String, ByVal BackupDirLen As Integer, ByVal fLocal As Integer, ByVal fOverwrite As Integer) As Integer
Private Declare Function sqlbss Lib "sqlwntm" (ByRef hServer As Integer, ByVal DatabaseName As String, ByVal DatabaseNameLen As Integer, ByVal BackupDir As String, ByVal BackupDirLen As Integer, ByVal fLocal As Integer, ByVal fOverwrite As Integer) As Integer
Private Declare Function sqlrss Lib "sqlwntm" (ByRef hServer As Integer, ByVal DatabaseName As String, ByVal DatabaseNameLen As Integer, ByVal BackupDir As String, ByVal BackupDirLen As Integer, ByVal fLocal As Integer, ByVal fOverwrite As Integer) As Integer
Private Declare Function sqldsv Lib "sqlwntm" (ByVal hServer As Integer) As Integer
Private Declare Function sqlgsi Lib "sqlwntm" (ByVal hServer As Integer, ByVal InfoFlags As Integer, buffer As Byte, ByVal buflen As Integer, ByRef rBufLen As Integer) As Integer
Private Declare Function sqlscl Lib "sqlwntm" (ByVal hServer As Integer, ByVal ClientName As String, ByVal ClientNameLen As Integer) As Integer
Private Declare Function sqlcre Lib "sqlwntm" (ByVal hServer As Integer, ByVal dbname As String, ByVal namelen As Integer) As Integer
Private Declare Function sqldel Lib "sqlwntm" (ByVal hServer As Integer, ByVal dbname As String, ByVal namelen As Integer) As Integer
Private Declare Function sqlded Lib "sqlwntm" (ByVal hServer As Integer, ByVal dbname As String, ByVal namelen As Integer) As Integer
Private Declare Function sqlind Lib "sqlwntm" (ByVal hServer As Integer, ByVal dbname As String, ByVal namelen As Integer) As Integer
Private Declare Function sqlerr Lib "sqlwntm" (ByVal rcd As Integer, ByVal errmsg As String) As Integer
Private Declare Function sqletx Lib "sqlwntm" (ByVal err As Integer, ByVal mtype As Integer, ByVal buf As String, ByVal buflen As Integer, ByRef retlen As Integer) As Integer
Private Declare Function sqlset Lib "sqlwntm" (ByVal hcur As Integer, ByVal nSParam As Integer, aBuffer As Any, ByVal nLen As Integer) As Integer

'message header...
Private Type hdrdefx
    hdrlen As Integer 'message length (including hdr)
    hdrrsv As Integer 'reserved
End Type

'message section header...
Private Type mshdefx
    mshflg As Integer     'section data type
    mshten As Integer     'total # of entries available
    mshnen As Integer     '# of entries contained in msg
    mshlen As Integer     '# of data bytes in msg section (including this header)
End Type

' gsidefi is used by sqlgsi - Get Server Information
Private Type gsidefx
    prcsel As Long      'number of selects
    prcins As Long      'number of inserts
    prcupd As Long      'number of updates
    prcdel As Long      'number of deletes
    prctps As Long      'number of transactions
    prcdlk As Long      'number of deadlocks
    prcelp As Long      'number of exclusive locks
    prcslp As Long      'number of shared locks
    prculp As Long      'number of update locks
    prcast As Long      'accumulative system time
    prcptp As Long      'time for prepare
    prcpte As Long      'time for execute
    prcptf As Long      'time for fetch
    prcmtt As Long      'maximum transaction time
    prcpss As String * 26  'status string
    prccln As String * 13   'client name
    prcsta As Byte      'status flag
    prcpts As String * 20   'process time stamp
    prciso As Byte      'isolation level flags
    prctmo As Long      'number of timeouts
    prcrsv As String * 27   'reserved
End Type

Private Type prcdefx
    prccol As Integer   'current output message length
    prcibl As Integer   'input message buffer length
    prcinl As Integer   'input length
    prcobl As Integer   'output message buffer length
    prcoul As Integer   'output length
    prcpnm As Byte      'process number
    prcact As Byte      'active flag
End Type

'filter structure...
Private Type fgidefx
  fgipnm As Byte            'process number
  fgicln As String * 13     'client name
  fgiunb As String * 19     'user name buffer
  fgidbn As String * 9      'database name buffer
End Type

Private iRC As Integer      ' return code
Private hServer As Integer  ' Server handle

Const BUFSIZ As Integer = 6000
Private Sub Class_Initialize()
    Set mConstants = New SIASDB.clsSIASDBConstants
End Sub
Private Sub Class_Terminate()
    Set mConstants = Nothing
End Sub
Public Function BuildSQLBaseError(err As Integer) As String
    Dim rawSQLerror As String * 1024
    Dim rLength As Integer
    Dim SQLError As String
        
    Call sqletx(err, SQLXMSG + SQLXREA + SQLXREM, rawSQLerror, Len(rawSQLerror), rLength)
    SQLError = Mid(rawSQLerror, 1, rLength)
    BuildSQLBaseError = SQLError
End Function
Public Function AddDBtoINI(DBEngine As SIASDB.dbeDBEngineEnum, NewServer As String, NewDatabaseName As String) As Boolean
    Dim SQLiniDir As String
    Dim SQLiniPath As String
    Dim NewIniPath As String
    Dim ErrorString As String
    Dim strServerPath As String
    Dim strServerName As String
    Dim iFileUnit As Integer
    Dim oFileUnit As Integer
    Dim SQLstring As String
    Dim fFoundWS32 As Boolean
    Dim fFoundServer As Boolean
    Dim strTime As String
    Dim dtNow As Date
    
    AddDBtoINI = False
    
    dtNow = Now
    strTime = Right(Format(Year(dtNow), "0000"), 2) & Format(Month(dtNow), "00") & Format(Day(dtNow), "00") & Format(Hour(dtNow), "00") & Format(Minute(dtNow), "00")
    
    If DBEngine <> dbeSQLBase Then
        err.Raise dberrDBENotSQLBase, mConstants.App.EXEName & " (AddDBtoINI)", mConstants.dberrDBENotSQLBaseStr
        GoTo ExitSub
    End If
    
    SQLiniDir = GetRegistryKeyValue(HKEY_CURRENT_USER, SQLINIKeyName, SQLINIKeyValue)
    'If SQLiniDir = vbNullString Then SQLiniDir = GetEnvironmentVariable("SQLBASE")
    If SQLiniDir = vbNullString Then
        err.Raise dberrDBESQLININTF, mConstants.App.EXEName & " (AddDBtoINI)", mConstants.dberrDBESQLININTFStr
        GoTo ExitSub
    Else
        SQLiniPath = SQLiniDir & "\SQL.INI"
    End If
    
    strServerPath = GetINIKey(SQLiniPath, KeyWS32, KeyServerPath, vbNullString)
    If strServerPath = vbNullString Then
        err.Raise dberrDBEServerPathNTF, mConstants.App.EXEName & " (AddDBtoINI)", Replace(Replace(Replace(mConstants.dberrDBEServerPathNTFStr, "%1", KeyServerPath), "%2", KeyWS32), "%3", SQLiniPath)
        GoTo ExitSub
    Else
        strServerName = Trim(ParseStr(strServerPath, 1, ",", """"))
        If UCase(NewServer) = UCase(strServerName) Then
            'Note: AddDBToPath does not return success/failure, but changed/unchanged...
            '      We don't want to unnecessarily update SQL.INI...
            If AddDBToPath(strServerPath, NewDatabaseName, False) Then Call SaveINIKey(SQLiniPath, KeyWS32, KeyServerPath, strServerPath)
        Else
            'We have to see if there are multiple serverpath entries in the INI file...
            'GetINIKey will have only returned the first...
            iFileUnit = FreeFile
            Open SQLiniPath For Input Access Read As #iFileUnit
            
            NewIniPath = SQLiniDir & "\SQL" & strTime & ".INI"
            oFileUnit = FreeFile
            Open NewIniPath For Output As #oFileUnit
            
            fFoundWS32 = False
            fFoundServer = False
            Do While Not EOF(iFileUnit)
                Line Input #iFileUnit, SQLstring
                If Trim(UCase(SQLstring)) = "[" & KeyWS32 & "]" Then fFoundWS32 = True
                If fFoundWS32 And KeyServerPath = Trim(UCase(ParseStr(SQLstring, 1, "=", """"))) Then
                    strServerName = Trim(UCase(ParseStr(SQLstring, 2, "=", """")))
                    strServerName = Trim(UCase(ParseStr(strServerName, 1, ",", """")))
                    If UCase(NewServer) = UCase(strServerName) Then
                        'Even though AddDBToPath is a function, call it here as a subroutine
                        'because the following logiv doesn't really care if it changed anything or not...
                        Call AddDBToPath(SQLstring, NewDatabaseName, True)
                        fFoundServer = True
                        fFoundWS32 = False  'This will speed up the rest of the copy operation...
                    End If
                End If
                Print #oFileUnit, SQLstring
            Loop
            Close #iFileUnit
            Close #oFileUnit
            If fFoundServer Then
                Kill SQLiniPath
                Name NewIniPath As SQLiniPath
            Else
                Kill NewIniPath
                err.Raise dberrDBENewServerPathNTF, mConstants.App.EXEName & " (AddDBtoINI)", Replace(Replace(Replace(mConstants.dberrDBENewServerPathNTFStr, "%1", KeyServerPath & "=" & NewServer), "%2", KeyWS32), "%3", SQLiniPath)
                GoTo ExitSub
            End If
        End If
    End If
    AddDBtoINI = True

ExitSub:
    Exit Function
End Function
Private Function AddDBToPath(SQLstring As String, NewDatabaseName As String, fReplaceWildCard As Boolean) As Boolean
    Dim i As Integer
    Dim fFoundDB As Boolean
    Dim strDBList As String
    Dim Token As String
    
    AddDBToPath = False
    strDBList = ParseStr(SQLstring, 2, "/", """")
    If strDBList = "*" Then
        If fReplaceWildCard Then
            SQLstring = Replace(SQLstring, "/*", "/" & NewDatabaseName)
            AddDBToPath = True
        End If
    Else
        fFoundDB = False
        For i = 1 To TokenCount(strDBList, ",")
            Token = Trim(UCase(ParseStr(strDBList, i, ",", """")))
            If Token = UCase(NewDatabaseName) Then
                fFoundDB = True
                Exit For
            End If
        Next
        If Not fFoundDB Then
            SQLstring = SQLstring & "," & NewDatabaseName
            AddDBToPath = True
        End If
    End If
End Function
Public Function DeleteDBfromINI(DBEngine As SIASDB.dbeDBEngineEnum, ServerName As String, DatabaseName As String) As Boolean
    Dim i As Integer
    Dim SQLiniDir As String
    Dim SQLiniPath As String
    Dim NewIniPath As String
    Dim ErrorString As String
    Dim strServerPath As String
    Dim strServerName As String
    Dim iFileUnit As Integer
    Dim oFileUnit As Integer
    Dim SQLstring As String
    Dim fFoundWS32 As Boolean
    Dim fFoundServer As Boolean
    Dim strTime As String
    Dim dtNow As Date
    
    DeleteDBfromINI = False
    
    dtNow = Now
    strTime = Right(Format(Year(dtNow), "0000"), 2) & Format(Month(dtNow), "00") & Format(Day(dtNow), "00") & Format(Hour(dtNow), "00") & Format(Minute(dtNow), "00")
    
    If DBEngine <> dbeSQLBase Then
        err.Raise dberrDBENotSQLBase, mConstants.App.EXEName & " (DeleteDBfromINI)", mConstants.dberrDBENotSQLBaseStr
        GoTo ExitSub
    End If
    
    SQLiniDir = GetRegistryKeyValue(HKEY_CURRENT_USER, SQLINIKeyName, SQLINIKeyValue)
    'If SQLiniDir = vbNullString Then SQLiniDir = GetEnvironmentVariable("SQLBASE")
    If SQLiniDir = vbNullString Then
        err.Raise dberrDBESQLININTF, mConstants.App.EXEName & " (DeleteDBfromINI)", mConstants.dberrDBESQLININTFStr
        GoTo ExitSub
    Else
        SQLiniPath = SQLiniDir & "\SQL.INI"
    End If
    
    strServerPath = GetINIKey(SQLiniPath, KeyWS32, KeyServerPath, vbNullString)
    If strServerPath = vbNullString Then
        err.Raise dberrDBEServerPathNTF, mConstants.App.EXEName & " (DeleteDBfromINI)", Replace(Replace(Replace(mConstants.dberrDBEServerPathNTFStr, "%1", KeyServerPath), "%2", KeyWS32), "%3", SQLiniPath)
        GoTo ExitSub
    Else
        strServerName = Trim(ParseStr(strServerPath, 1, ",", """"))
        If UCase(ServerName) = UCase(strServerName) Then
            Call RemoveDBFromServerPath(strServerPath, DatabaseName)
            Call SaveINIKey(SQLiniPath, KeyWS32, KeyServerPath, strServerPath)
        Else
            'We have to see if there are multiple serverpath entries in the INI file...
            'GetINIKey will have only returned the first...
            iFileUnit = FreeFile
            Open SQLiniPath For Input Access Read As #iFileUnit
            
            NewIniPath = SQLiniDir & "\SQL" & strTime & ".INI"
            oFileUnit = FreeFile
            Open NewIniPath For Output As #oFileUnit
            
            fFoundWS32 = False
            fFoundServer = False
            Do While Not EOF(iFileUnit)
                Line Input #iFileUnit, SQLstring
                If Trim(UCase(SQLstring)) = "[" & KeyWS32 & "]" Then fFoundWS32 = True
                If fFoundWS32 And KeyServerPath = Trim(UCase(ParseStr(SQLstring, 1, "=", """"))) Then
                    strServerName = Trim(UCase(ParseStr(SQLstring, 2, "=", """")))
                    strServerName = Trim(UCase(ParseStr(strServerName, 1, ",", """")))
                    If UCase(ServerName) = UCase(strServerName) Then
                        SQLstring = Trim(UCase(ParseStr(SQLstring, 2, "=", """")))
                        Call RemoveDBFromServerPath(SQLstring, DatabaseName)
                        SQLstring = KeyServerPath & "=" & SQLstring
                        fFoundServer = True
                        fFoundWS32 = False  'This will speed up the rest of the copy operation...
                    End If
                End If
                Print #oFileUnit, SQLstring
            Loop
            Close #iFileUnit
            Close #oFileUnit
            If fFoundServer Then
                Kill SQLiniPath
                Name NewIniPath As SQLiniPath
            Else
                Kill NewIniPath
                err.Raise dberrDBENewServerPathNTF, mConstants.App.EXEName & " (DeleteDBfromINI)", Replace(Replace(Replace(mConstants.dberrDBENewServerPathNTFStr, "%1", KeyServerPath & "=" & ServerName), "%2", KeyWS32), "%3", SQLiniPath)
                GoTo ExitSub
            End If
        End If
    End If
    DeleteDBfromINI = True

ExitSub:
    Exit Function
End Function
Private Sub RemoveDBFromServerPath(strServerPath As String, strDB As String)
    Dim i As Integer
    Dim strDBList As String
    Dim Token As String
    Dim fFoundDB As Boolean
    Dim strTemp As String
    
    strTemp = ParseStr(strServerPath, 1, "/", """") & "/"
    strDBList = ParseStr(strServerPath, 2, "/", """")
    For i = 1 To TokenCount(strDBList, ",")
        Token = Trim(ParseStr(strDBList, i, ",", """"))
        If Token <> strDB Then
            If i > 1 Then strTemp = strTemp & ","
            strTemp = strTemp & Token
        End If
    Next
    strServerPath = strTemp
End Sub
Private Function StringZ(stStringVal As String) As String
    Dim i As Integer
    i = InStr(stStringVal, 0)
    If i > 0 Then
        StringZ = Mid(stStringVal, 1, i - 1)
    Else
        StringZ = stStringVal
    End If
End Function
Public Function IsInSQLBase(ByVal Server As String, ByVal DatabaseName As String, ByVal UserName As String, ByVal ClientName As String) As Boolean
    Dim buffer(BUFSIZ) As Byte
    Dim fFound As Boolean
    Dim i As Integer
    Dim iLen As Integer
    Dim InfoFlags As Integer
    Dim iPos As Integer
    Dim iPrc As Integer
    Dim rBufLen As Integer
    Dim fgidef As fgidefx
    Dim hdrdef As hdrdefx
    Dim mshdef As mshdefx
    Dim prcdef As prcdefx
    Dim gsidef As gsidefx
    
    '?To always return False until SQLBase bug is fixed????
    IsInSQLBase = False
    Exit Function
    
    fFound = False
    
    InfoFlags = SQLXGSI + SQLGPRC
    
    fgidef.fgipnm = 0
    If DatabaseName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRDBN
        fgidef.fgidbn = UCase(DatabaseName) & String(Len(fgidef.fgidbn) - Len(DatabaseName), vbNullChar)
    Else
        fgidef.fgidbn = String(Len(fgidef.fgidbn), vbNullChar)
    End If
    If UserName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRUSN
        fgidef.fgiunb = UCase(UserName) & String(Len(fgidef.fgiunb) - Len(UserName), vbNullChar)
    Else
        fgidef.fgiunb = String(Len(fgidef.fgiunb), vbNullChar)
    End If
    If ClientName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRCLN
        fgidef.fgicln = ClientName & String(Len(fgidef.fgicln) - Len(ClientName), vbNullChar)
    Else
        fgidef.fgicln = String(Len(fgidef.fgicln), vbNullChar)
    End If
    CopyMemory buffer(0), fgidef, Len(fgidef)
    
    iRC = sqlcsv(hServer, Server, "SYSADM")
    iRC = sqlgsi(hServer, InfoFlags, buffer(0), Len(fgidef), rBufLen)
    If rBufLen < Len(hdrdef) Then
        MsgBox "Message too small to contain header", vbExclamation
        GoTo ExitSub
    End If
    
    CopyMemory hdrdef, buffer(0), Len(hdrdef)
    iPos = Len(hdrdef)
    iLen = hdrdef.hdrlen - Len(hdrdef)
    While iLen >= Len(mshdef)
        CopyMemory mshdef, buffer(iPos), Len(mshdef)
        iPos = iPos + Len(mshdef)
        iLen = iLen - mshdef.mshlen
        
        Select Case mshdef.mshflg
            Case SQLGPRC           'process information
                For iPrc = 1 To mshdef.mshnen
                    CopyMemory prcdef, buffer(iPos), Len(prcdef)
                    iPos = iPos + Len(prcdef)
                    CopyMemory gsidef, buffer(iPos), Len(gsidef)
                    iPos = iPos + Len(gsidef)
                                        
                    fFound = True
                Next iPrc
            Case Else
        End Select
    Wend

ExitSub:
    iRC = sqldsv(hServer)
    IsInSQLBase = fFound
End Function
Public Sub SQLBasePrintGetProcessInfo(ByVal Server As String, ByVal DatabaseName As String, ByVal UserName As String, ByVal ClientName As String)
    Dim buffer(BUFSIZ) As Byte
    Dim p As Long
    Dim i As Integer
    Dim iLen As Integer
    Dim InfoFlags As Integer
    Dim iPos As Integer
    Dim iPrc As Integer
    Dim rBufLen As Integer
    Dim fgidef As fgidefx
    Dim hdrdef As hdrdefx
    Dim mshdef As mshdefx
    Dim prcdef As prcdefx
    Dim gsidef As gsidefx
    
    InfoFlags = SQLXGSI + SQLGPRC
    
    fgidef.fgipnm = 0
    If DatabaseName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRDBN
        fgidef.fgidbn = UCase(DatabaseName) & String(Len(fgidef.fgidbn) - Len(DatabaseName), vbNullChar)
    Else
        fgidef.fgidbn = String(Len(fgidef.fgidbn), vbNullChar)
    End If
    If UserName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRUSN
        fgidef.fgiunb = UCase(UserName) & String(Len(fgidef.fgiunb) - Len(UserName), vbNullChar)
    Else
        fgidef.fgiunb = String(Len(fgidef.fgiunb), vbNullChar)
    End If
    If ClientName <> vbNullString Then
        InfoFlags = InfoFlags + SQLRCLN
        fgidef.fgicln = ClientName & String(Len(fgidef.fgicln) - Len(ClientName), vbNullChar)
    Else
        fgidef.fgicln = String(Len(fgidef.fgicln), vbNullChar)
    End If
    CopyMemory buffer(0), fgidef, Len(fgidef)
    
    iRC = sqlcsv(hServer, Server, "SYSADM")
    iRC = sqlgsi(hServer, InfoFlags, buffer(0), Len(fgidef), rBufLen)
    If rBufLen < Len(hdrdef) Then
        MsgBox "Message too small to contain header", vbExclamation
        GoTo ExitSub
    End If
    
    CopyMemory hdrdef, buffer(0), Len(hdrdef)
    iPos = Len(hdrdef)
    iLen = hdrdef.hdrlen - Len(hdrdef)
    While iLen >= Len(mshdef)
        CopyMemory mshdef, buffer(iPos), Len(mshdef)
        iPos = iPos + Len(mshdef)
        iLen = iLen - mshdef.mshlen
        
        Select Case mshdef.mshflg
            Case SQLGCUR           'cursor information
            Case SQLGDBS           'database information
            Case SQLGCFG           'configuration information
            Case SQLGSTT           'statistics
            Case SQLGPRC           'process information
                Debug.Print ""
                Debug.Print "Extracted Process Information:"
                Debug.Print "ClientName  " & vbTab & "Status "
                For iPrc = 1 To mshdef.mshnen
                    CopyMemory prcdef, buffer(iPos), Len(prcdef)
                    iPos = iPos + Len(prcdef)
                    CopyMemory gsidef, buffer(iPos), Len(gsidef)
                    iPos = iPos + Len(gsidef)
                    
                    Debug.Print gsidef.prccln & String(13 - Len(gsidef.prccln), " ") & vbTab & gsidef.prcpss & "(" & gsidef.prcsta & ")"
                    'Debug.Print "CurOutLen" & vbTab & "InBufLen" & vbTab & "InLen" & vbTab & "OutBufLen" & vbTab & "OutLen" & vbTab & "PNum" & vbTab & "Status"
                    'With prcdef
                    '    Debug.Print .prccol & vbTab & .prcibl & vbTab & .prcinl & vbTab & .prcobl & vbTab & .prcoul & vbTab & .prcpnm & vbTab & .prcact
                    'End With
                Next iPrc
            Case Else
        End Select
    Wend

ExitSub:
    iRC = sqldsv(hServer)
End Sub
Public Function SQLBaseSetClientName(ByVal Server As String, ByVal ClientName As String) As Integer
    SQLBaseSetClientName = sqlcsv(hServer, Server, "SYSADM")
    If SQLBaseSetClientName = 0 Then
        SQLBaseSetClientName = sqlscl(hServer, ClientName, Len(ClientName))
        iRC = sqldsv(hServer)
    End If
End Function
Public Function SQLBaseBackupDatabase(ByVal Server As String, ByVal DatabaseName As String, ByVal BackupDir As String, ByRef ErrorMessage) As Integer
    Dim iStatus As Integer
    Dim iCursor As Integer
    Dim lbmset As Integer   'Log Backup Mode
    
    ErrorMessage = vbNullString
    iStatus = sqlcnc(iCursor, DatabaseName, Len(DatabaseName))
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iStatus = sqlset(iCursor, SQLPLBM, lbmset, 0)
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iStatus = sqlcsv(hServer, Server, "SYSADM")
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iStatus = sqlbdb(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iStatus = sqlrel(iCursor)
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iStatus = sqlblf(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
    If iStatus <> 0 Then
        ErrorMessage = BuildSQLBaseError(iStatus)
        GoTo ExitSub
    End If
    iRC = sqldsv(hServer)
    
ExitSub:
    SQLBaseBackupDatabase = iStatus
End Function
Public Function SQLBaseBackupLogs(ByVal Server As String, ByVal DatabaseName As String, ByVal BackupDir As String, ByRef ErrorMessage) As Integer
    Dim iStatus As Integer
    iStatus = sqlcsv(hServer, Server, "SYSADM")
    If iStatus = 0 Then
        iStatus = sqlblf(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
        If iStatus = 0 Then ErrorMessage = vbNullString Else ErrorMessage = BuildSQLBaseError(iStatus)
        '09279 GCI EOR Cannot find specified protocol entry
        iRC = sqldsv(hServer)
    End If
    SQLBaseBackupLogs = iStatus
End Function
Public Function SQLBaseBackupSnapShot(ByVal Server As String, ByVal DatabaseName As String, ByVal BackupDir As String, ByRef ErrorMessage) As Integer
    Dim iStatus As Integer
    iStatus = sqlcsv(hServer, Server, "SYSADM")
    If iStatus = 0 Then
        iStatus = sqlbss(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
        If iStatus = 0 Then ErrorMessage = vbNullString Else ErrorMessage = BuildSQLBaseError(iStatus)
        '09279 GCI EOR Cannot find specified protocol entry
        iRC = sqldsv(hServer)
    End If
    SQLBaseBackupSnapShot = iStatus
End Function
Public Function SQLBaseCreateDatabase(ByVal Server As String, ByVal DatabaseName As String) As Integer
    SQLBaseCreateDatabase = sqlcsv(hServer, Server, "SYSADM")
    If SQLBaseCreateDatabase = 0 Then
        SQLBaseCreateDatabase = sqlcre(hServer, DatabaseName, Len(DatabaseName))
        '09279 GCI EOR Cannot find specified protocol entry
        iRC = sqldsv(hServer)
    End If
End Function
Public Function SQLBaseDeleteDatabase(ByVal Server As String, ByVal DatabaseName As String) As Integer
    SQLBaseDeleteDatabase = sqlcsv(hServer, Server, "SYSADM")
    If SQLBaseDeleteDatabase = 0 Then
        SQLBaseDeleteDatabase = sqldel(hServer, DatabaseName, Len(DatabaseName))
        iRC = sqldsv(hServer)
    End If
End Function
Public Function SQLBaseInstallDatabase(ByVal Server As String, ByVal DatabaseName As String) As Integer
    SQLBaseInstallDatabase = sqlcsv(hServer, Server, "SYSADM")
    If SQLBaseInstallDatabase = 0 Then
        SQLBaseInstallDatabase = sqlind(hServer, DatabaseName, Len(DatabaseName))
        iRC = sqldsv(hServer)
    End If
End Function
Public Function SQLBaseDeinstallDatabase(ByVal Server As String, ByVal DatabaseName As String) As Integer
    SQLBaseDeinstallDatabase = sqlcsv(hServer, Server, "SYSADM")
    If SQLBaseDeinstallDatabase = 0 Then
        SQLBaseDeinstallDatabase = sqlded(hServer, DatabaseName, Len(DatabaseName))
        iRC = sqldsv(hServer)
    End If
End Function
Public Function SQLBaseRestoreLogs(ByVal Server As String, ByVal DatabaseName As String, ByVal BackupDir As String, ByRef ErrorMessage) As Integer
    Dim iStatus As Integer
    iStatus = sqlcsv(hServer, Server, "SYSADM")
    If iStatus = 0 Then
        iStatus = sqlrlf(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
        If iStatus = 0 Then ErrorMessage = vbNullString Else ErrorMessage = BuildSQLBaseError(iStatus)
        '09279 GCI EOR Cannot find specified protocol entry
        iRC = sqldsv(hServer)
    End If
    SQLBaseRestoreLogs = iStatus
End Function
Public Function SQLBaseRestoreSnapShot(ByVal Server As String, ByVal DatabaseName As String, ByVal BackupDir As String, ByRef ErrorMessage) As Integer
    Dim iStatus As Integer
    iStatus = sqlcsv(hServer, Server, "SYSADM")
    If iStatus = 0 Then
        iStatus = sqlrss(hServer, DatabaseName, Len(DatabaseName), BackupDir, Len(BackupDir), iLocal, iOverwrite)
        If iStatus = 0 Then ErrorMessage = vbNullString Else ErrorMessage = BuildSQLBaseError(iStatus)
        '09279 GCI EOR Cannot find specified protocol entry
        iRC = sqldsv(hServer)
    End If
    SQLBaseRestoreSnapShot = iStatus
End Function

