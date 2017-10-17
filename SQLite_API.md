
SQLite API

Public Declare PtrSafe Function sqlite_open Lib "sqlite.dll" (ByRef file As String, ByVal db_mode As Integer) As Integer
Public Declare PtrSafe Function sqlite_open_in_memory Lib "sqlite.dll" (ByRef file As String) As Integer
Public Declare PtrSafe Function sqlite_open_cache Lib "sqlite.dll" (ByRef file As String) As Integer
Public Declare PtrSafe Function sqlite_select Lib "sqlite.dll" (ByRef sql As String, ByRef Result() As String, ByVal db_mode As Integer) As Long
Public Declare PtrSafe Function sqlite_sql Lib "sqlite.dll" (ByRef sql As String, ByVal db_mode As Integer) As Integer
Public Declare PtrSafe Function sqlite_close Lib "sqlite.dll" (ByRef db_mode As Integer, ByRef update As Integer) As Integer
Public Declare PtrSafe Function sqlite_close_in_memory Lib "sqlite.dll" (ByRef update As Integer) As Integer
Public Declare PtrSafe Function sqlite_close_cache Lib "sqlite.dll" (ByVal update As Integer) As Integer


ADO API

Public Declare PtrSafe Function ado_connect Lib "dbconnect.dll" (ByRef ConnectionString as String) As Integer
Public Declare PtrSafe Function ado_select Lib "dbconnect.dll" (ByRef sql As String, ByRef result() as String) As Long
Public Declare PtrSafe Function ado_close Lib "dbconnect.dll" (ByRef ConnectionString As String) As Integer

SQLite API

Public Declare PtrSafe Function sqlite_open Lib "sqlite.dll" (ByRef file As String, ByRef db_mode As Integer) As Integer
- SQLite DB Open, Disk Access
- file : 경로를 포함한 DB file 이름
- db_mode : 1 또는 2. 두 개의 세션을 동시에 사용하기 위한 구분자임

Public Declare PtrSafe Function sqlite_open_in_memory Lib "sqlite.dll" (ByRef file As String) As Integer
- SQLite DB Open, Memory Access
- file : 경로를 포함한 DB file 이름

Public Declare PtrSafe Function sqlite_open_cache Lib "sqlite.dll" (ByRef file As String) As Integer
- SQLite Memory Cache
- file : blank("") 입력

Public Declare PtrSafe Function sqlite_select Lib "sqlite.dll" (ByRef sql As String, ByRef Result() as String, ByVal db_mode As Integer) As Long
- Query String 전송 후 Result 배열(1차원)으로 결과 반환
- 반환 배열 Index 는 1 부터 값이 들어옴(즉, index 0 은 비어있음)
- 함수 반환값은 결과 레코드의 갯수
- Select 문 외에 다른 Query String 을 넣어도 무방함
- Select 문으로 2개 이상의 컬럼을 지정한 경우(col1, col2 또는 *)에는 "|" separator 을 삽입하여 string concatenation 처리해서 반환
- db_mode : 1 - File DB(1), 2 - File DB(2), 3 - Memory DB, 4 - Memory Cache (반드시 "ByVal" 로 정의해 줘야 하는 것에 주의!)

Public Declare PtrSafe Function sqlite_sql Lib "sqlite.dll" (ByRef sql As String, ByVal db_mode As Integer) As Integer
- Query 명령어 실행
- 수행작업은 sqlite_select 와 동일하나 결과 반환이 없음
- Create 나 Insert 같은 명령어를 사용할 수 있음
- db_mode : 1 - File DB(1), 2 - File DB(2), 3 - Memory DB, 4 - Memory Cache (반드시 "ByVal" 로 정의해 줘야 하는 것에 주의!)

Public Declare PtrSafe Function sqlite_close Lib "sqlite.dll" (ByRef update As Integer) As Integer
- File DB Close
- db_mode : 1 또는 2. 두 개의 세션을 동시에 사용하기 위한 구분자임
- update : 의미 없음(무조건 저장)

Public Declare PtrSafe Function sqlite_close_in_memory Lib "sqlite.dll" (ByRef update As Integer) As Integer
- Memory DB Close
- sqlite_open_in_memory 로 DB Open 한 경우 update = 0 이면 변경내용을 DB 로 저장하지 않고 그냥 Close
- sqlite_open_in_memory 로 DB Open 한 경우 update = 1 이면 변경내용을 DB 로 저장하고 Close(데이터가 많은 경우 상당시간 소요될 수 있음)

Public Declare PtrSafe Function sqlite_close_cache Lib "sqlite.dll" (ByVal db_mode As Integer, ByVal update As Integer) As Integer
- Memory Cache Close
- update : 의미 없음


ADO API

Public Declare PtrSafe Function ado_connect Lib "dbconnect.dll" (ByRef ConnectionString as String) As Integer
- DB 객체 생성
- ConnectString : ODBC Connection String 전달

Public Declare PtrSafe Function ado_select Lib "dbconnect.dll" (ByRef sql As String, ByRef result() as String) As Long
- Query String 전송 후 Result 배열(1차원)으로 결과 반환
- 반환 배열 Index 는 1 부터 값이 들어옴(즉, index 0 은 비어있음)
- 함수 반환값은 결과 레코드의 갯수
- Select 문 외에 다른 Query String 을 넣어도 무방함
- Select 문으로 2개 이상의 컬럼을 지정한 경우(col1, col2 또는 *)에는 "|" separator 을 삽입하여 string concatenation 처리해서 반환

Public Declare PtrSafe Function ado_close Lib "dbconnect.dll" (ByRef ConnectionString As String) As Integer
- DB Close


Error Code
#define SQLITE_ERROR        1   /* Generic error */
#define SQLITE_INTERNAL     2   /* Internal logic error in SQLite */
#define SQLITE_PERM         3   /* Access permission denied */
#define SQLITE_ABORT        4   /* Callback routine requested an abort */
#define SQLITE_BUSY         5   /* The database file is locked */
#define SQLITE_LOCKED       6   /* A table in the database is locked */
#define SQLITE_NOMEM        7   /* A malloc() failed */
#define SQLITE_READONLY     8   /* Attempt to write a readonly database */
#define SQLITE_INTERRUPT    9   /* Operation terminated by sqlite3_interrupt()*/
#define SQLITE_IOERR       10   /* Some kind of disk I/O error occurred */
#define SQLITE_CORRUPT     11   /* The database disk image is malformed */
#define SQLITE_NOTFOUND    12   /* Unknown opcode in sqlite3_file_control() */
#define SQLITE_FULL        13   /* Insertion failed because database is full */
#define SQLITE_CANTOPEN    14   /* Unable to open the database file */
#define SQLITE_PROTOCOL    15   /* Database lock protocol error */
#define SQLITE_EMPTY       16   /* Not used */
#define SQLITE_SCHEMA      17   /* The database schema changed */
#define SQLITE_TOOBIG      18   /* String or BLOB exceeds size limit */
#define SQLITE_CONSTRAINT  19   /* Abort due to constraint violation */
#define SQLITE_MISMATCH    20   /* Data type mismatch */
#define SQLITE_MISUSE      21   /* Library used incorrectly */
#define SQLITE_NOLFS       22   /* Uses OS features not supported on host */
#define SQLITE_AUTH        23   /* Authorization denied */
#define SQLITE_FORMAT      24   /* Not used */
#define SQLITE_RANGE       25   /* 2nd parameter to sqlite3_bind out of range */
#define SQLITE_NOTADB      26   /* File opened that is not a database file */
#define SQLITE_NOTICE      27   /* Notifications from sqlite3_log() */
#define SQLITE_WARNING     28   /* Warnings from sqlite3_log() */
#define SQLITE_ROW         100  /* sqlite3_step() has another row ready */
#define SQLITE_DONE        101  /* sqlite3_step() has finished executing */