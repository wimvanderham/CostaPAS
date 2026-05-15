&ANALYZE-SUSPEND _VERSION-NUMBER UIB_v9r12
&ANALYZE-RESUME
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS Procedure 
/*------------------------------------------------------------------------
    File        : dynquerytt.p
    Purpose     : Returns a dynamic temp-table as the result of a 
                  query string

    Syntax      : run dynquerytt.p
                     (input ipcQueryString as char, /* Query string, must start with each */
                      input ipcFieldList   as char, /* List of field names, if multiple tables are involved, the tablename must be specified */
                      input ipcFieldNames  as char, /* Corresponding list of name in the temp-table */
                      input ipcTTName      as char, /* Name of the dynamic temp-table */
                      output ophDynTT as table-handle). /* Output handle of dynamic temp-table */

    Description :

    Author(s)   : Wim van der Ham
    Created     : 10 September 2002
    Notes       :
    
    Modified    : 25 settembre 2002
    By          : Wim van der Ham
    Reason      : Support of array fields during buffer-copy
                  Note: This means that when you specify a field in the fieldlist
                  which is an array the field in the temp-table will be created 
                  like an array field too and ALL values will be copied to the TT
                  
    Modified    : 29 Settembre 2002
    By          : Wim van der Ham
    Reason      : Further support of array fields, this time you can specify a
                  single array field in the format fieldname[1] and get back the
                  first element from that array field. The fieldname in the TT
                  will be created like the database field but with NO extent.
    
    Modified    : 21 Febbraio 2003
    By          : Wim van der Ham
    Reason      : Separated creation of dynamic buffers and fields to reflect
                  the order of field list in the dynamic temp-table
  
    Modified    : 6 Ottobre 2004
    By          : Wim van der Ham
    Reason      : Se query string inizia con "FOR", togliere "FOR"
    
    Modified    : 27 Gennaio 2011
    By          : Wim van der Ham
    Reason      : Se input TTName è "ttCount" allora, conta solo
    
    Modified    : 05 Marzo 2012
    By          : Wim van der Ham
    Reason      : Aggiunto supporto per special field COUNT(SubQuery)
    
    Modified    : 14 Maggio 2012
    By          : Wim van der Ham e Roberta Mosso
    Reason      : Possibilità di passare proprietà dentro il parametro ipcTTName
                  Formato deve essere: Proprietà=Valore~nProprietà2=Valore2 etc.
                  
    Modified    : 16 Luglio 2012
    By          : Wim van der Ham e Roberta Mosso
    Reason      : Aggiunta di parametro Distinct in ipcTTName.
                  Elenco di fields per raggruppare e contare per il gruppo creato
                  
    Modified    : 03 Settembre 2012
    By          : Wim van der Ham e Roberta Mosso
    Reason      : Aggiunta di parametro CheckOnly in ipcTTName
                  Serve per verificare la query string (senza restituire dati)
                  Corretto errore sulla gestione di errori in QUERY-PREPARE                  
  ----------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/


/* ***************************  Definitions  ************************** */
DEFINE INPUT  PARAMETER ipcQueryString AS CHARACTER  NO-UNDO.
DEFINE INPUT  PARAMETER ipcFieldList   AS CHARACTER  NO-UNDO.
DEFINE INPUT  PARAMETER ipcFieldNames  AS CHARACTER  NO-UNDO.
DEFINE INPUT  PARAMETER ipcTTName      AS CHARACTER  NO-UNDO.
DEFINE OUTPUT PARAMETER TABLE-HANDLE ophTT.

/* Error codes */
&global-define errNotSupported        -20
&global-define errNoTable             -10
&global-define errTableNotReady       -11
&global-define errWrongDefinition     -12
&global-define errUnmatchedQuotes     -13
&global-define errUnqualifiedField    -14
&global-define errUnknownError        -999
&global-define errExternalProc        -50
&global-define errProcedureNotFound   293
&global-define errWrongQueryDef       3322
&global-define errMaxBuffersReached   7318
&global-define errUnknownTable        545
&GLOBAL-DEFINE errUnknownSpecialField -60
&GLOBAL-DEFINE errTemp-Table-prepare  -61
&GLOBAL-DEFINE errAssignFieldsError   -62
&GLOBAL-DEFINE errFieldDataMismatch   -63
&GLOBAL-DEFINE errNoArrayField        366
&GLOBAL-DEFINE errSubscriptIncorrect  367
&GLOBAL-DEFINE errSubScriptError      368

DEFINE VARIABLE hTTError       AS HANDLE     NO-UNDO.
define variable hBufError      as handle     no-undo.
define variable hComboKey      as handle     no-undo.
define variable hQueryDef      as handle     no-undo.
define variable hKeyFields     as handle     no-undo.
define variable hDescFields    as handle     no-undo.
define variable hDelimiter     as handle     no-undo.
define variable hTTQuery       as handle     no-undo.

DEFINE VARIABLE hTTBuffer   AS HANDLE     NO-UNDO.
DEFINE VARIABLE hTTField    AS HANDLE     NO-UNDO.

define variable lvc_error      as character  no-undo.
define variable lvi_errorcnt   as integer    no-undo.
DEFINE VARIABLE lviRowCounter  AS INTEGER    NO-UNDO.

/* Temp-variable to prebuild the querystring */
define variable cQueryString   as character  no-undo.
DEFINE VARIABLE iSeqNr AS INTEGER    NO-UNDO.
DEFINE VARIABLE lvlDistinct AS LOGICAL     NO-UNDO. /* check se viene chiesto il distinct in ipcTTName */
DEFINE VARIABLE lvcDistinctFieldList AS CHARACTER   NO-UNDO. /* elenco campi distinct */
DEFINE VARIABLE lQueryPrepareOk      AS LOGICAL     NO-UNDO.

/* This is used to keep track of all the tables used
   in the query */
define temp-table ttTables no-undo
   field tOrder  as INTEGER   FORMAT "z9"    LABEL "Table Nr"
   field tDb     as CHARACTER FORMAT "X(15)" LABEL "Database"
   field tTable  as CHARACTER FORMAT "X(20)" LABEL "Table"
   field tHandle as handle
      index pk as primary unique
         tOrder
      index tbl
         tDb
         tTable.

/* This is used to keep track of all the fields used
   in the query */
define temp-table ttFields no-undo
   field tOrder    as integer
   field tTable    as INTEGER   FORMAT "z9"    LABEL "Table Nr"
   FIELD tDBField  AS CHARACTER FORMAT "X(20)" LABEL "DB Field"
   field tField    as CHARACTER FORMAT "X(20)" LABEL "TT Field"
   FIELD iIndex    AS INTEGER   FORMAT "zz9"   LABEL "Index"
   field tFormat   as CHARACTER FORMAT "X(10)" LABEL "Format"
   FIELD cSubQuery AS CHARACTER FORMAT "X(40)" LABEL "Sub Query"
   field tHandle   as handle
      index pk as primary unique
         tOrder.

/* This temp-table stores all errors */
DEFINE TEMP-TABLE ttError NO-UNDO RCODE-INFORMATION
   FIELD iSeqNr AS INT FORMAT "zz9"
   FIELD iError AS INT FORMAT "zzzzz9"
   FIELD cError AS CHAR FORMAT "X(60)"
INDEX idxSeq AS PRIMARY UNIQUE iSeqNr.

DEFINE TEMP-TABLE ttCount NO-UNDO
   FIELD NrRecords AS INTEGER
.

DEFINE TEMP-TABLE ttProperty 
   FIELD PropertyName AS CHARACTER
   FIELD PropertyValue AS CHARACTER
INDEX idx1 AS PRIMARY UNIQUE PropertyName.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Procedure
&Scoped-define DB-AWARE no



/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME


/* ************************  Function Prototypes ********************** */

&IF DEFINED(EXCLUDE-addError) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD addError Procedure 
FUNCTION addError RETURNS LOGICAL
  ( /* parameter-definitions */ 
    INPUT ipiErrorNr AS INT,
    INPUT ipcError   AS CHAR
  )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-addSpecialField) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD addSpecialField Procedure 
FUNCTION addSpecialField RETURNS LOGICAL
  ( /* parameter-definitions */ 
    INPUT iphTT      AS HANDLE,
    INPUT ipcSpecial AS CHAR,
    INPUT ipcName    AS CHAR,
    INPUT ipcFormat  AS CHAR
     )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getBaseNames) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD getBaseNames Procedure 
FUNCTION getBaseNames RETURNS CHARACTER
  ( /* parameter-definitions */ 
    INPUT ipcFields AS CHAR 
  )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getEndChar) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD getEndChar Procedure 
FUNCTION getEndChar RETURNS INTEGER
  ( /* parameter-definitions */
    INPUT ipcString AS CHARACTER,
    INPUT ipiStart  AS INTEGER
  )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getProperty) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD getProperty Procedure 
FUNCTION getProperty RETURNS CHARACTER
  (INPUT ipcName AS CHAR /* parameter-definitions */ )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-openQuote) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD openQuote Procedure 
FUNCTION openQuote RETURNS LOGICAL
  ( ipc_string as character )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Procedure
   Allow: 
   Frames: 0
   Add Fields to: Neither
   Other Settings: CODE-ONLY COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
/* DESIGN Window definition (used by the UIB) 
  CREATE WINDOW Procedure ASSIGN
         HEIGHT             = 14.33
         WIDTH              = 60.2.
/* END WINDOW DEFINITION */
                                                                        */
&ANALYZE-RESUME

 


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK Procedure 


/* ***************************  Main Block  *************************** */
IF ipcQueryString BEGINS "FOR" THEN
   ipcQueryString = TRIM(SUBSTRING(ipcQueryString, 4)).

IF NUM-ENTRIES(ipcTTName, "~n") GT 1 THEN DO:
   IF NUM-ENTRIES(ENTRY(1, ipcTTName, "~n"), "=") LE 1 THEN
      ipcTTName = SUBSTITUTE("TTName=&1", ipcTTName).
END.
ELSE DO:
   ipcTTName = SUBSTITUTE("TTName=&1", ipcTTName).
END.
RUN setProperties
   (INPUT ipcTTName).

OUTPUT TO VALUE(SUBSTITUTE("dynquerytt_&1.log", getProperty("TTName"))).
PUT UNFORMATTED TODAY " " STRING(TIME, "HH:MM:SS") SKIP.
PUT UNFORMATTED "Query String: " ipcQueryString SKIP.
PUT UNFORMATTED "Field list  : " ipcFieldList   SKIP.
PUT UNFORMATTED "Field names : " ipcFieldNames  SKIP.



lvcDistinctFieldList = getProperty("distinct").
lvlDistinct = FALSE.
IF lvcDistinctFieldList NE "" THEN
   lvlDistinct = TRUE.


FOR EACH ttProperty:
   PUT UNFORMATTED SUBSTITUTE("&1: &2", ttProperty.PropertyName, ttProperty.PropertyValue) SKIP.
END.

/* Parse query string to extract tables and fields */
IF NUM-ENTRIES(ipcFieldNames) NE NUM-ENTRIES(ipcFieldList)
THEN DO:
   /* Number of Field NAMES doesn't match number of fields in list */
   ipcFieldNames = getBaseNames(ipcFieldList).
   PUT UNFORMATTED "Corrected Field names : " ipcFieldNames SKIP.
END.

RUN parseQueryString 
   (INPUT ipcQueryString,
    INPUT ipcFieldList,
    INPUT ipcFieldNames) NO-ERROR.

IF ERROR-STATUS:ERROR = TRUE THEN DO:
   PUT UNFORMATTED 
      RETURN-VALUE SKIP.
   ophTT = hTTError.
END.
ELSE DO:
   PUT UNFORMATTED
      "Parsing ok" SKIP.

   FOR EACH ttTables:
      DISPLAY
         ttTables EXCEPT tHandle
      WITH STREAM-IO TITLE " Tables Parsed ".
   END.

   FOR EACH ttFields:
      DISPLAY
         ttFields EXCEPT tHandle
      WITH STREAM-IO TITLE " Fields in List " WIDTH 132.
   END.

   IF getProperty("TTName") EQ "ttCount" THEN DO:
      RUN countData
         (INPUT ipcQueryString) NO-ERROR.
      
      IF ERROR-STATUS:ERROR = TRUE THEN DO:
         PUT UNFORMATTED
            RETURN-VALUE SKIP.
         ophTT = hTTError.
      END.
      ELSE DO:
         PUT UNFORMATTED
            "Counted data" SKIP
            "Number of records: " lviRowCounter SKIP.

         CREATE ttCount.
         ASSIGN
            ttCount.NrRecords = lviRowCounter
         .
         ophTT = TEMP-TABLE ttCount:HANDLE.

      END.
   END.
   ELSE DO:
      RUN getData
         (INPUT ipcQueryString) NO-ERROR.
      IF ERROR-STATUS:ERROR = TRUE THEN DO:
         PUT UNFORMATTED
            RETURN-VALUE SKIP.
         ophTT = hTTError.
      END.
      ELSE DO:
         PUT UNFORMATTED
            "Got data" SKIP
            "Number of records: " lviRowCounter SKIP.

         /*
         RUN outputData.
         */

         DELETE OBJECT ophTT. /* Will be postphoned until procedure ends */

      END.
   END.

END.

PUT UNFORMATTED TODAY " " STRING(TIME, "HH:MM:SS") SKIP.
OUTPUT CLOSE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&IF DEFINED(EXCLUDE-addRecords) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE addRecords Procedure 
PROCEDURE addRecords :
/*------------------------------------------------------------------------------
  Purpose:     Adds free form records
  Parameters:  input ipcRecordData a list of CHR(1) separated records in which 
               the fields values are separated with CHR(2)
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ipcRecordData AS CHARACTER  NO-UNDO.

DEFINE VARIABLE iRecordNr   AS INTEGER    NO-UNDO.
DEFINE VARIABLE cRecordData AS CHARACTER  NO-UNDO.
DEFINE VARIABLE iFieldNr    AS INTEGER    NO-UNDO.

   DO iRecordNr = 1 TO NUM-ENTRIES(ipcRecordData, CHR(1)):
      /* For all initial records */
      cRecordData = ENTRY(iRecordNr, ipcRecordData, CHR(1)).
      IF NUM-ENTRIES(cRecordData, CHR(2)) EQ NUM-ENTRIES(ipcFieldNames) THEN DO:
         /* Number of fields match with data */
         hTTBuffer:BUFFER-CREATE().
         
         lviRowCounter = lviRowCounter + 1.

         DO iFieldNr = 1 TO NUM-ENTRIES(ipcFieldNames):
            /* Loop through temp-table fields */
            hTTField = hTTBuffer:BUFFER-FIELD(iFieldNr).
            ASSIGN
               hTTField:STRING-VALUE = ENTRY(iFieldNr, cRecordData, CHR(2))
            NO-ERROR.
            IF ERROR-STATUS:ERROR = TRUE THEN DO:
               lvc_error = "".
               repeat lvi_errorcnt = 1 to error-status:num-messages:
                  addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
                  
                  lvc_error = lvc_error 
                              + (if lvc_error = "" then "" else ",")
                              + string(error-status:get-number(lvi_errorcnt)) + ";"
                              + error-status:get-message(lvi_errorcnt).
               end.
               addError({&errAssignFieldsError}, lvc_error).
               return error lvc_error.
            END.
         END. /* Loop through temp-table fields */

      END. /* Number of fields match with data */
      ELSE DO:
         /* Mismatch in number of fields */
         lvc_error = SUBSTITUTE("Field data mismatch. Expected &1 fields, received &2 fields",
                                NUM-ENTRIES(ipcFieldNames),
                                NUM-ENTRIES(cRecordData, CHR(2))).
         addError({&errFieldDataMismatch}, lvc_error).
         RETURN ERROR lvc_error.
      END. /* Mismatch in number of fields */
   END. /* For all initial records */
      

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-countData) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE countData Procedure 
PROCEDURE countData :
/*------------------------------------------------------------------------------
  Purpose:     Counts the data and stores it in a dynamic temp-table
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ipc_querydef AS CHARACTER  NO-UNDO.

define variable lvh_q1             as handle     no-undo.
define variable lvh_fwork          as handle     no-undo.
define variable lvh_work           as handle     no-undo.
define variable lvl_firstiteration as logical    no-undo.
define variable lvl_firstkey       as logical    no-undo.
define variable lvl_firstdesc      as logical    no-undo.
define variable lvc_keyvalue       as character  no-undo.
define variable lvc_descvalue      as character  no-undo.
define variable lvc_keydelimiter   as character  no-undo.
define variable lvc_descdelimiter  as character  no-undo.
define variable lvl_chartype       as logical    no-undo.

define variable lvc_result         as character  no-undo.
define variable lvc_procedure      as character  no-undo.
DEFINE VARIABLE iIndex             AS INTEGER    NO-UNDO.

DEFINE VARIABLE cRowIDS AS CHARACTER  NO-UNDO.

   create query lvh_q1.
   
   for each ttTables:
      /* All tables in query */
      if ttTables.tDb <> "?" then
         create buffer lvh_work for table (ttTables.tDb + "." + ttTables.tTable) no-error.
      else
         create buffer lvh_work for table ttTables.tTable no-error.
   
      /* Everything ok? */
      if error-status:error = false and valid-handle(lvh_work) then do:
         ttTables.tHandle = lvh_work.
         lvh_q1:add-buffer(lvh_work).
         if error-status:error then do:
               lvc_error = "".
               repeat lvi_errorcnt = 1 to error-status:num-messages:
                  addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
                  lvc_error = lvc_error 
                              + (if lvc_error = "" then "" else ",")
                              + string(error-status:get-number(lvi_errorcnt)) + ";"
                              + error-status:get-message(lvi_errorcnt).
               end.
               if lvc_error <> "" then
                  return error lvc_error.
               else DO:
                  addError({&errUnknownError}, "Couldn't assign buffer " + ttTables.tTable + " to query.").
                  return error "{&errUnknownError};" 
                     + "Couldn't assign buffer " + ttTables.tTable + " to query.".
               END.
         end.
   
      end.
      else if error-status:error then do:
            lvc_error = "".
            repeat lvi_errorcnt = 1 to error-status:num-messages:
               lvc_error = lvc_error 
                           + (if lvc_error = "" then "" else ",")
                           + string(error-status:get-number(lvi_errorcnt)) + ";"
                           + error-status:get-message(lvi_errorcnt).
               addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
            end.
            return error lvc_error.
      end.
      ELSE DO:
         addError({&errUnknownError}, "Unkown error occured while creating buffer for " + ttTables.tTable).
         return error "{&errUnknownError};Unkown error occured while creating buffer for " 
            + ttTables.tTable.
      END.
   end. /* All tables in query */
   
/*    IF ipcInitRecords NE "" THEN DO:          */
/*       /* Create initial records */           */
/*       RUN addRecords (INPUT ipcInitRecords). */
/*    END. /* Create initial records */         */

   /* Time to prepare and a PRESELECT query for counting */
   lQueryPrepareOk = lvh_q1:query-prepare("preselect " + ipc_querydef) no-error.
   if lQueryPrepareOk = FALSE THEN do: /*error-status:error then do: 03/09/12 error-status:error in questo caso darebbe false */
      lvc_error = "".
      repeat lvi_errorcnt = 1 to error-status:num-messages:
         addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
         lvc_error = lvc_error 
                     + (if lvc_error = "" then "" else ",")
                     + string(error-status:get-number(lvi_errorcnt)) + ";"
                     + error-status:get-message(lvi_errorcnt).
      end.
      return error lvc_error.
   end.
   
   IF NOT getProperty("CheckOnly") EQ "true" THEN DO:
      /* 03-09-12 se non devo fare solo il check vado avanti... altrimenti lascio che il programma vada avanti per svuotamenti e simili */

      lvh_q1:query-open no-error.
      if error-status:error then do:
         lvc_error = "".
   
         repeat lvi_errorcnt = 1 to error-status:num-messages:
            addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
            lvc_error = lvc_error 
                        + (if lvc_error = "" then "" else ",")
                        + string(error-status:get-number(lvi_errorcnt)) + ";"
                        + error-status:get-message(lvi_errorcnt).
         end.
         return error lvc_error.
      end.
      
      if lvh_q1:is-open then do:
         lviRowCounter = lvh_q1:NUM-RESULTS.
   
         lvh_q1:query-close().
      end.
      
   /*    IF ipcEndRecords NE "" THEN DO:          */
   /*       /* Create end records */              */
   /*       RUN addRecords (INPUT ipcEndRecords). */
   /*    END. /* Create end records */            */
   END.

   /* Do cleanup */
   delete object lvh_q1.
   for each ttFields:
      if valid-handle(ttFields.tHandle) then
         delete object ttFields.tHandle no-error.
   end.
   for each ttTables:
      if valid-handle(ttTables.tHandle) then
         delete object ttTables.tHandle no-error.
   end.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getCount) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE getCount Procedure 
PROCEDURE getCount :
/*------------------------------------------------------------------------------
  Purpose:     Returns count of a subquery
  Parameters:  INPUT  PARAMETER ipcSubQuery   AS CHARACTER   NO-UNDO.
               OUTPUT PARAMETER opiCount      AS INTEGER     NO-UNDO.

  Notes:       Can use values of fields available in all tables of the MAIN
               query. Access a value by specifiying @table.field or 
               @db.table.field
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ipcSubQuery   AS CHARACTER   NO-UNDO.
DEFINE OUTPUT PARAMETER opiCount      AS INTEGER     NO-UNDO.

DEFINE BUFFER ttTables FOR ttTables.

DEFINE VARIABLE cSubQuery   AS CHARACTER   NO-UNDO.
DEFINE VARIABLE iAT         AS INTEGER     NO-UNDO.
DEFINE VARIABLE iEnd        AS INTEGER     NO-UNDO.
DEFINE VARIABLE iLunghezza  AS INTEGER     NO-UNDO.
DEFINE VARIABLE cVariabile  AS CHARACTER   NO-UNDO.
DEFINE VARIABLE cNomeDato   AS CHARACTER   NO-UNDO.
DEFINE VARIABLE hBuffer     AS HANDLE      NO-UNDO.
DEFINE VARIABLE hField      AS HANDLE      NO-UNDO.
DEFINE VARIABLE cValoreDato AS CHARACTER   NO-UNDO.
DEFINE VARIABLE hTT         AS HANDLE      NO-UNDO.
DEFINE VARIABLE hBufferTT   AS HANDLE      NO-UNDO.

   ASSIGN
      cSubQuery = ipcSubQuery
   .
   DO WHILE INDEX(cSubQuery, "@") NE 0:
      iAT = INDEX(cSubQuery, "@").
      iEnd = getEndChar(cSubQuery, iAT).

      iLunghezza = iEnd - iAT.
      cVariabile = SUBSTRING(cSubQuery, iAT, iLunghezza).
      cNomeDato  = TRIM(cVariabile, "@").

      ASSIGN
         hField = ?
      .

      IF NUM-ENTRIES(cNomeDato, ".") EQ 3 THEN DO:
         FIND FIRST ttTables
         WHERE ttTables.tDb    EQ ENTRY(1, cNomeDato, ".")
         AND   ttTables.tTable EQ ENTRY(2, cNomeDato, ".") NO-ERROR.
         IF AVAILABLE ttTables THEN DO:
            hBuffer = ttTables.tHandle.
            ASSIGN
               hField = hBuffer:BUFFER-FIELD(ENTRY(3, cNomeDato, "."))
            NO-ERROR.
         END.
      END.
      IF NUM-ENTRIES(cNomeDato, ".") EQ 2 THEN DO:
         FIND FIRST ttTables
         WHERE ttTables.tTable EQ ENTRY(1, cNomeDato, ".") NO-ERROR.
         IF AVAILABLE ttTables THEN DO:
            hBuffer = ttTables.tHandle.
            ASSIGN
               hField = hBuffer:BUFFER-FIELD(ENTRY(2, cNomeDato, "."))
            NO-ERROR.
         END.
      END.

      IF VALID-HANDLE(hField) THEN DO:
         CASE hField:DATA-TYPE:
            WHEN "INTEGER" OR
            WHEN "DECIMAL" OR
            WHEN "LOGICAL" THEN
               cValoreDato = SUBSTITUTE("&1", hField:BUFFER-VALUE).
            WHEN "CHARACTER" THEN
               cValoreDato = QUOTER(hField:BUFFER-VALUE).
            WHEN "DATE" THEN
               cValoreDato = SUBSTITUTE("DATE(&1)",
                                        INTEGER(hField:BUFFER-VALUE)).
         END CASE.

      END.
      ELSE DO:
         cValoreDato = "".
      END.
      
      SUBSTRING(cSubQuery, iAT, iLunghezza) = cValoreDato.

   END.

   /* Count */
   RUN sy/p/dynquerytt.p
     (INPUT  cSubQuery,
      INPUT "#ROWIDS",
      INPUT "ROWIDS",
      INPUT "ttCount",
      OUTPUT TABLE-HANDLE hTT).

   hBufferTT = hTT:DEFAULT-BUFFER-HANDLE.
   hBufferTT:FIND-FIRST().
   IF hBufferTT:AVAILABLE THEN DO:
      opiCount = hBufferTT::NrRecords.
   END.

   DELETE OBJECT hTT.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getData) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE getData Procedure 
PROCEDURE getData :
/*------------------------------------------------------------------------------
  Purpose:     Gets the data and stores it in a dynamic temp-table
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ipc_querydef AS CHARACTER  NO-UNDO.

define variable lvh_q1             as handle     no-undo.
define variable lvh_fwork          as handle     no-undo.
define variable lvh_work           as handle     no-undo.
define variable lvl_firstiteration as logical    no-undo.
define variable lvl_firstkey       as logical    no-undo.
define variable lvl_firstdesc      as logical    no-undo.
define variable lvc_keyvalue       as character  no-undo.
define variable lvc_descvalue      as character  no-undo.
define variable lvc_keydelimiter   as character  no-undo.
define variable lvc_descdelimiter  as character  no-undo.
define variable lvl_chartype       as logical    no-undo.
DEFINE VARIABLE cWhere             AS CHARACTER   NO-UNDO.

define variable lvc_result         as character  no-undo.
define variable lvc_procedure      as character  no-undo.
DEFINE VARIABLE iIndex             AS INTEGER    NO-UNDO.
DEFINE VARIABLE lvcField1          AS CHARACTER   NO-UNDO.
DEFINE VARIABLE lvcField           AS CHARACTER   NO-UNDO.
DEFINE VARIABLE lvhFieldTT         AS HANDLE      NO-UNDO.
DEFINE VARIABLE lvhFieldDB         AS HANDLE      NO-UNDO.
DEFINE VARIABLE lvcWhere           AS CHARACTER   NO-UNDO.
DEFINE VARIABLE ix                 AS INTEGER     NO-UNDO.

DEFINE VARIABLE cRowIDS     AS CHARACTER  NO-UNDO.

DEFINE VARIABLE iCount      AS INTEGER     NO-UNDO.

   create query lvh_q1.
   
   CREATE TEMP-TABLE ophTT.
   
   for each ttTables:
      /* All tables in query */
      if ttTables.tDb <> "?" then
         create buffer lvh_work for table (ttTables.tDb + "." + ttTables.tTable) no-error.
      else
         create buffer lvh_work for table ttTables.tTable no-error.
   
      /* Everything ok? */
      if error-status:error = false and valid-handle(lvh_work) then do:
         ttTables.tHandle = lvh_work.
         lvh_q1:add-buffer(lvh_work).
         if error-status:error then do:
               lvc_error = "".
               repeat lvi_errorcnt = 1 to error-status:num-messages:
                  addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
                  lvc_error = lvc_error 
                              + (if lvc_error = "" then "" else ",")
                              + string(error-status:get-number(lvi_errorcnt)) + ";"
                              + error-status:get-message(lvi_errorcnt).
               end.
               if lvc_error <> "" then
                  return error lvc_error.
               else DO:
                  addError({&errUnknownError}, "Couldn't assign buffer " + ttTables.tTable + " to query.").
                  return error "{&errUnknownError};" 
                     + "Couldn't assign buffer " + ttTables.tTable + " to query.".
               END.
         end.
   
      end.
      else if error-status:error then do:
            lvc_error = "".
            repeat lvi_errorcnt = 1 to error-status:num-messages:
               lvc_error = lvc_error 
                           + (if lvc_error = "" then "" else ",")
                           + string(error-status:get-number(lvi_errorcnt)) + ";"
                           + error-status:get-message(lvi_errorcnt).
               addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
            end.
            return error lvc_error.
      end.
      ELSE DO:
         addError({&errUnknownError}, "Unkown error occured while creating buffer for " + ttTables.tTable).
         return error "{&errUnknownError};Unkown error occured while creating buffer for " 
            + ttTables.tTable.
      END.
   end. /* All tables in query */
   
   /* Go and create all the field handles */
   for each ttFields:
      /* For each ttFields */
      IF ttFields.tTable EQ 0 THEN DO:
         /* Special field */
         IF addSpecialField(ophTT, ttFields.tDBField, ttFields.tField, ttFields.tFormat) = FALSE THEN DO:
            addError({&errUnknownSpecialField}, "Unknown special field " + ttFields.tDBField + " " + ttFields.tField).
            RETURN ERROR "{&errUnknownSpecialField};Unknown special field " + ttFields.tDBField + " " + 
               ttFields.tField.
         END.
      end. /* Special field */
      ELSE DO:
         /* Normal field from table in query */
         FIND ttTables WHERE ttTables.tOrder = ttFields.tTable.
         lvh_work  = ttTables.tHandle.
         lvh_fwork = lvh_work:buffer-field(ttFields.tDBField) no-error.
         if error-status:error = false and valid-handle(lvh_fwork) THEN DO:
            /* Correct field */
            ttFields.tHandle = lvh_fwork.
            /* Add it to the dynamic temp-table */
            IF ttFields.tHandle:EXTENT GT 0 THEN DO:
               /* Database field is an arry */
               IF ttFields.iIndex GT ttFields.tHandle:EXTENT THEN DO:
                  addError({&errSubscriptIncorrect}, SUBSTITUTE("Subscript incorrect for field &1, DB field has array of &2 elements and you've asked for element &3",
                                                                ttFields.tDBField,
                                                                ttFields.tHandle:EXTENT,
                                                                ttFields.iIndex)).
                  return error "{&errSubscriptIncorrect};" +
                     SUBSTITUTE("Subscript incorrect for field &1, DB field has array of &2 elements and you've asked for element &3",
                                ttFields.tDBField,
                                ttFields.tHandle:EXTENT,
                                ttFields.iIndex).
               END.
               IF ttFields.iIndex GT 0 THEN DO:
                  /* Asked for a specific element instead of whole array, set extent to 0 */
                  ophTT:ADD-NEW-FIELD(ttFields.tField, 
                                      ttFields.tHandle:DATA-TYPE, 
                                      0, 
                                      ttFields.tHandle:FORMAT, 
                                      ttFields.tHandle:INITIAL, 
                                      ttFields.tHandle:LABEL + " " + STRING(ttFields.iIndex), 
                                      ttFields.tHandle:COLUMN-LABEL + " " + STRING(ttFields.iIndex)).
               END. /* Asked for a specific element instead of whole array, set extent to 0 */
               ELSE DO:
                  /* Asked for whole array */
                  ophTT:ADD-LIKE-FIELD(ttFields.tField, ttFields.tHandle).
               END. /* Asked for whole array */
            END. /* Database field is an arry */
            ELSE DO:
               /* Database field is NO array field */
               IF ttFields.iIndex GT 0 THEN DO:
                  addError({&errNoArrayField}, SUBSTITUTE("Subscript incorrect for field &1, DB field is no array field you've asked for element &2",
                                                                ttFields.tDBField,
                                                                ttFields.iIndex)).
                  return error "{&errNoArrayField};" +
                     SUBSTITUTE("Subscript incorrect for field &1, DB field is no array field but you've asked for element &2",
                                ttFields.tDBField,
                                ttFields.iIndex).
               END.
               ELSE DO:
                  ophTT:ADD-LIKE-FIELD(ttFields.tField, ttFields.tHandle).
               END.
            END. /* Database field is NO array field */
         END. /* Correct field */
         else if error-status:error then do:
            lvc_error = "".
            repeat lvi_errorcnt = 1 to error-status:num-messages:

               addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
               lvc_error = lvc_error 
                           + (if lvc_error = "" then "" else ",")
                           + string(error-status:get-number(lvi_errorcnt)) + ";"
                           + error-status:get-message(lvi_errorcnt).
            end.
            return error lvc_error.
         end.
         else DO:
            addError({&errUnknownError}, "Couldn't create a handle for the field " + ttFields.tDBField).
            return error "{&errUnknownError};" 
               + "Couldn't create a handle for the field " + ttFields.tDBField.
         END.
      END. /* Normal field from table in query */
   END. /* For each ttFields */

   /* Parameter DISTINCT ? */
   IF lvlDistinct THEN DO:
      /* Add counter field */
      ophTT:ADD-NEW-FIELD("iCount", "INTEGER", ?, "zz,zz9", ?, "Count").
      /* Create Index on DistintFieldList */
      lvcField1 = ENTRY(1,lvcDistinctFieldList).
      IF INDEX(lvcField1,".") > 0 THEN DO:
         lvcField1 = ENTRY(NUM-ENTRIES(lvcField1,"."),lvcField1,".").
      END.
      ophTT:ADD-NEW-INDEX(SUBSTITUTE("ind&1", lvcField1), YES, YES).
      DO ix = 1 TO NUM-ENTRIES(lvcDistinctFieldList):
         lvcField = entry(ix,lvcDistinctFieldList).
         IF (INDEX(lvcField,".") > 0) THEN
            lvcField = ENTRY(NUM-ENTRIES(lvcField,"."),lvcField,".").
         ophTT:ADD-INDEX-FIELD(SUBSTITUTE("ind&1", lvcField1), lvcField).
      END. 
   END. /* Add counter field */

   /* Finish definition of temp-table */
   IF ophTT:TEMP-TABLE-PREPARE(getProperty("TTName")) = FALSE THEN DO:
      addError({&errTemp-Table-prepare}, "Couldn't prepare temp-table " + getProperty("TTName")).
      RETURN ERROR "{&errTemp-Table-prepare};Couldn't prepare temp-table " + getProperty("TTName").
   END.
   hTTBuffer = ophTT:DEFAULT-BUFFER-HANDLE.
   
/*    IF ipcInitRecords NE "" THEN DO:          */
/*       /* Create initial records */           */
/*       RUN addRecords (INPUT ipcInitRecords). */
/*    END. /* Create initial records */         */

   /* Time to prepare and open the query */
   lQueryPrepareOk = lvh_q1:query-prepare("for " + ipc_querydef) no-error.
   if lQueryPrepareOk = FALSE THEN do: /*error-status:error then do: 03-09-12 */
      lvc_error = "".
      addError(0, SUBSTITUTE("Query: FOR &1", ipc_querydef)).
      repeat lvi_errorcnt = 1 to error-status:NUM-MESSAGES:
         addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
         lvc_error = lvc_error 
                     + (if lvc_error = "" then "" else ",")
                     + string(error-status:get-number(lvi_errorcnt)) + ";"
                     + error-status:get-message(lvi_errorcnt).
      end.
      return error lvc_error.
   end.

   IF NOT getProperty("CheckOnly") EQ "true" THEN DO:
      /* 03-09-12 se non devo fare solo il check vado avanti... altrimenti lascio che il programma vada avanti per svuotamenti e simili */

      lvh_q1:query-open no-error.
      if error-status:error then do:
         
         lvc_error = "".
         repeat lvi_errorcnt = 1 to error-status:num-messages:
            
            addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
            lvc_error = lvc_error 
                        + (if lvc_error = "" then "" else ",")
                        + string(error-status:get-number(lvi_errorcnt)) + ";"
                        + error-status:get-message(lvi_errorcnt).
         end.
         return error lvc_error.
      end.
      
      if lvh_q1:is-open then do:
         lviRowCounter = 0.
      
         lvh_q1:get-first.
      
         lvl_firstiteration  = true.
         /* And now it's time to get the data */
         repeat while not lvh_q1:query-off-end:
            IF getProperty("MaxRecords") NE "" THEN DO:
               IF lviRowCounter GT INTEGER(getProperty("MaxRecords")) THEN
                  LEAVE.
            END.
            lviRowCounter = lviRowCounter + 1.
      
            IF lvlDistinct EQ TRUE THEN DO:
               /* Find record in hTTBuffer with lvcDistinctFieldList */
               /* If N/A --> BUFFER-CREATE() */
               lvcWhere = "".
               DO ix = 1 TO NUM-ENTRIES(lvcDistinctFieldList):
                  lvcField = entry(ix,lvcDistinctFieldList).
                  IF (INDEX(lvcField,".") > 0) THEN
                     lvcField = ENTRY(NUM-ENTRIES(lvcField,"."),lvcField,".").
                  lvhFieldDB = lvh_work:buffer-field(lvcField). /** **/
                  IF lvcWhere EQ "" THEN
                     if lvhFieldDB:data-type = "CHARACTER" then
                        lvcWhere = SUBSTITUTE("&1 = &2",lvhFieldDB:NAME,quoter(lvhFieldDb:BUFFER-VALUE)).
                     else
                        lvcWhere = SUBSTITUTE("&1 = '&2'",lvhFieldDB:NAME,lvhFieldDb:BUFFER-VALUE).
                  ELSE
                     if lvhFieldDB:data-type = "CHARACTER" then
                        lvcWhere = SUBSTITUTE("&3 AND &1 = &2",lvhFieldDB:NAME,quoter(lvhFieldDb:BUFFER-VALUE,lvcWhere)).
                     else
                        lvcWhere = SUBSTITUTE("&3 AND &1 = '&2'",lvhFieldDB:NAME,lvhFieldDb:BUFFER-VALUE,lvcWhere). 
               END.
               hTTBuffer:FIND-FIRST (substitute("&1 &2", "WHERE",lvcWhere),NO-LOCK) NO-ERROR.
               IF hTTBuffer:AVAILABLE = FALSE THEN DO TRANSACTION:
                  /* Create record nella TT e valorizzare campi */
                  hTTBuffer:BUFFER-CREATE().
                  
   
                  DO ix = 1 TO NUM-ENTRIES(lvcDistinctFieldList):
                     lvcField = entry(ix,lvcDistinctFieldList).
                     IF (INDEX(lvcField,".") > 0) THEN
                        lvcField = ENTRY(NUM-ENTRIES(lvcField,"."),lvcField,".").
                     lvhFieldTT = hTTBuffer:BUFFER-FIELD(lvcField).
                     lvhFieldDB = lvh_work:BUFFER-FIELD(lvcField).
                     lvhFieldTT:BUFFER-VALUE = lvhFieldDB:BUFFER-VALUE.
                  END.
                  hTTBuffer:BUFFER-FIELD("iCount"):BUFFER-VALUE = 0.
               END.                                               
               /* Increment counter */
               hTTBuffer:BUFFER-FIELD("iCount"):BUFFER-VALUE = hTTBuffer:BUFFER-FIELD("iCount"):BUFFER-VALUE + 1.
            END. 
            ELSE DO:
               /* Normal getData --> buffer copy every record from DB Query to TT */
               hTTBuffer:BUFFER-CREATE().
            END. /* Normal getData --> buffer copy every record from DB Query to TT */
      
            /* Get's the values for the fields */
            for each ttFields:
               IF ttFields.tTable = 0 THEN DO:
                  /* Special field */
                  hTTField = hTTBuffer:BUFFER-FIELD(ttFields.tField).
                  CASE ttFields.tDBField:
                     WHEN "LINE-COUNTER" THEN DO:
                        hTTField:BUFFER-VALUE = lviRowCounter.
                     END.
                     WHEN "ROWIDS" THEN DO:
                        cROWIDS = "".
                        FOR EACH ttTables:
                           IF ttTables.tHandle:AVAILABLE THEN DO:
                              /* Got a record for this table */
                              IF cROWIDS NE "" THEN
                                 cROWIDS = cROWIDS + ",".
                              cROWIDS = cROWIDS +
                                 STRING(ttTables.tHandle:ROWID).
                           END. /* Got a record for this table */
                           ELSE DO:
                              cROWIDS = cROWIDS + ",".
                           END.
                        END.
                        hTTField:BUFFER-VALUE = cROWIDS.
                     END.
                     WHEN "COUNT" THEN DO:
                        RUN getCount
                           (INPUT  ttFields.cSubQuery,
                            OUTPUT iCount).
                        hTTField:BUFFER-VALUE = iCount.
                     END.
                  END CASE.
               END. /* Special field */
               ELSE DO:
                  assign 
                     lvh_fwork    = ttFields.tHandle
                  .
                  /* Copy DB data to TT field */
                  hTTField = hTTBuffer:BUFFER-FIELD(ttFields.tField).
                  IF hTTField:EXTENT EQ 0 THEN DO:
                     /* Normal field */
                     IF lvh_fwork:EXTENT GT 0 THEN DO:
                        /* Database field has extent, pick specific element */
                        ASSIGN
                           hTTField:BUFFER-VALUE = lvh_fwork:BUFFER-VALUE(ttFields.iIndex)
                        NO-ERROR.
                     END. /* Database field has extent, pick specific element */
                     ELSE DO:
                        ASSIGN
                           hTTField:BUFFER-VALUE = lvh_fwork:BUFFER-VALUE 
                        NO-ERROR.
                     END.
                     IF ERROR-STATUS:ERROR = TRUE THEN DO:
                        lvc_error = "".
                        repeat lvi_errorcnt = 1 to error-status:num-messages:
                                    
                           addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
                           lvc_error = lvc_error 
                                       + (if lvc_error = "" then "" else ",")
                                       + string(error-status:get-number(lvi_errorcnt)) + ";"
                                       + error-status:get-message(lvi_errorcnt).
                        end.
                        return error lvc_error.
                     END.
                  END. /* Normal field */
                  ELSE DO:
                     /* Array field */
                     DO iIndex = 1 TO hTTField:EXTENT:
                        ASSIGN
                           hTTField:BUFFER-VALUE(iIndex) = lvh_fwork:BUFFER-VALUE(iIndex)
                        NO-ERROR.
                        IF ERROR-STATUS:ERROR = TRUE THEN DO:
                           lvc_error = "".
                           repeat lvi_errorcnt = 1 to error-status:num-messages:
   
                              addError(ERROR-STATUS:GET-NUMBER(lvi_errorcnt), ERROR-STATUS:GET-MESSAGE(lvi_errorcnt)).
                              lvc_error = lvc_error 
                                          + (if lvc_error = "" then "" else ",")
                                          + string(error-status:get-number(lvi_errorcnt)) + ";"
                                          + error-status:get-message(lvi_errorcnt).
                           end.
                           return error lvc_error.
                        END.
                     END.
                  END. /* Array field */
               END.
            end.
      
            lvh_q1:get-next.
         end.
         lvh_q1:query-close.
      end.
   END. /* 03-09-12 */   
/*    IF ipcEndRecords NE "" THEN DO:          */
/*       /* Create end records */              */
/*       RUN addRecords (INPUT ipcEndRecords). */
/*    END. /* Create end records */            */

   /* Do cleanup */
   delete object lvh_q1.
   for each ttFields:
      if valid-handle(ttFields.tHandle) then
         delete object ttFields.tHandle no-error.
   end.
   for each ttTables:
      if valid-handle(ttTables.tHandle) then
         delete object ttTables.tHandle no-error.
   end.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-outputData) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE outputData Procedure 
PROCEDURE outputData :
/*------------------------------------------------------------------------------
  Purpose:     Outputs the temp-table data for debugging purposes
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE hQuery       AS HANDLE     NO-UNDO.
DEFINE VARIABLE hBuffer      AS HANDLE     NO-UNDO.
DEFINE VARIABLE iFieldNr     AS INTEGER    NO-UNDO.
DEFINE VARIABLE hField       AS HANDLE     NO-UNDO.

   CREATE QUERY  hQuery.
   hBuffer = ophTT:DEFAULT-BUFFER-HANDLE.
   hQuery:ADD-BUFFER(hBuffer).

   hQuery:QUERY-PREPARE(SUBSTITUTE("FOR EACH &1", hBuffer:NAME)).
   hQuery:QUERY-OPEN.

   IF hBuffer:AVAILABLE THEN DO:
      /* We got a record, output the labels */
      DO iFieldNr = 1 TO hBuffer:NUM-FIELDS:
         hField = hBuffer:BUFFER-FIELD(iFieldNr).
         PUT UNFORMATTED
            hField:LABEL.
         IF iFieldNr LT hBuffer:NUM-FIELDS THEN DO:
            PUT UNFORMATTED
               "~t".
         END.
      END.
      PUT UNFORMATTED
         SKIP.
   END.
   
   DO WHILE hQuery:QUERY-OFF-END = FALSE:
      /* Now loop through the fields of the current record */
      DO iFieldNr = 1 TO hBuffer:NUM-FIELDS:
         hField = hBuffer:BUFFER-FIELD(iFieldNr).
         PUT UNFORMATTED
            hField:STRING-VALUE.
         IF iFieldNr LT hBuffer:NUM-FIELDS THEN DO:
            PUT UNFORMATTED
               "~t".
         END.
      END.
      PUT UNFORMATTED
         SKIP.

      hQuery:GET-NEXT.
   END. /* Now loop through the fields of the current record */

   hQuery:QUERY-CLOSE().

   DELETE OBJECT hQuery.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-parseQueryString) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE parseQueryString Procedure 
PROCEDURE parseQueryString :
/*------------------------------------------------------------------------------
  Purpose:     Takes the querydefinition and makes something useful out of it
  Parameters:  input char: Querydefinition
               input char: Fields requested
  Notes:       
------------------------------------------------------------------------------*/
define input  parameter ipc_querydef  as character  no-undo.
DEFINE INPUT  PARAMETER ipc_fields    AS CHARACTER  NO-UNDO.
DEFINE INPUT  PARAMETER ipc_names     AS CHARACTER  NO-UNDO.

define variable         ia               as integer    no-undo.
define variable         ib               as integer    no-undo.
define variable         lvc_work         as character  no-undo.
define variable         lvc_listfields   as character  no-undo.
define variable         lvc_format       as character  no-undo.
DEFINE VARIABLE         lvc_name         AS CHARACTER  NO-UNDO.
DEFINE VARIABLE         cSubScript       AS CHARACTER  NO-UNDO.
DEFINE VARIABLE         iIndex           AS INTEGER    NO-UNDO.
DEFINE VARIABLE         lvc_subquery     AS CHARACTER   NO-UNDO.
DEFINE VARIABLE         iDistinctField   AS INTEGER     NO-UNDO.
DEFINE VARIABLE         cDistinctField   AS CHARACTER   NO-UNDO.

empty temp-table ttTables no-error.
empty temp-table ttFields no-error.

/* First we should extract all the tables to be used (max 18)
   every part of the query definition is separated by ','
   and the tablename should be preceeded by a 'each', 'first' or 'last' except
   for the first table that must be preceeded by a 'each'
*/


   /* Extracing the tables from the query */
   repeat ia = 1 to num-entries(ipc_querydef):
   
      /* Make sure the user don't go over the 18-buffers limit in dynamic-queries */
      if ia > 18 THEN DO:
         addError({&errMaxBuffersReached}, "You can have a maximum of 18 tables in one query").
         return error "{&errMaxBuffersReached};" 
            + "You can have a maximum of 18 tables in one query".
      END.
   
      lvc_work = trim(entry(ia,ipc_querydef)).
   
      if ia = 1 and entry(1,lvc_work," ") <> "each" THEN DO:
         addError({&errWrongQueryDef}, "The first record specification must be a EACH").
         return error "{&errWrongQueryDef};" 
                  + "The first record specification must be a EACH".
      END.
   
      /* Make sure we got the entire table-block */
      do while openQuote(lvc_work):
         if ia >= num-entries(ipc_querydef) then do:
            addError({&errUnmatchedQuotes}, "Found a unmatched quote").
            return error "{&errUnmatchedQuotes};Found a unmatched quote".
         end.
         ia = ia + 1.
         lvc_work = lvc_work + "," + trim(entry(ia,ipc_querydef)).
      end.
   
      if can-do("first,last,each",entry(1,lvc_work," ")) then do:
         assign lvc_work = trim(substring(lvc_work,index(lvc_work," ")))
                lvc_work = trim(entry(1,lvc_work," ")).
   
         create ttTables.
         assign ttTables.tOrder = ia
                ttTables.tDB    = (if num-entries(lvc_work,".") > 1 then 
                                       entry(1,lvc_work,".") else "?")
                ttTables.tTable = (if num-entries(lvc_work,".") > 1 then 
                                       entry(2,lvc_work,".") else lvc_work).
      end.
      else DO:
         addError({&errWrongQueryDef}, "Subsequent record specifications should have either EACH, FIRST or LAST").
         return error "{&errWrongQueryDef};" 
             + "Subsequent record specifications should have either EACH, FIRST or LAST".
      END.
   end.

   /* Now we have all the tables used, now we should get all the fields */
   assign lvc_listfields = ipc_fields.

   IF lvlDistinct EQ TRUE THEN DO:
      DO iDistinctField = 1 TO NUM-ENTRIES(lvcDistinctFieldList):
         cDistinctField = "".
         cDistinctField = ENTRY(iDistinctField, lvcDistinctFieldList).
         IF LOOKUP(cDistinctField, lvc_listfields) EQ 0 THEN DO:
            IF lvc_listfields NE "" THEN DO:
               lvc_listfields = lvc_listfields + ",".
               ipc_names = ipc_names + ",".
            END.
            ASSIGN
               lvc_listfields = SUBSTITUTE("&1&2", lvc_listfields, cDistinctField)
               ipc_names      = SUBSTITUTE("&1&2", ipc_names, ENTRY(NUM-ENTRIES(cDistinctField, "."), cDistinctField, "."))
            .
         END.
      END.
      ipc_names      = SUBSTITUTE("&1,&2", ipc_names, "iCount").
   END.

   repeat ia = 1 to num-entries(lvc_listfields):
      assign lvc_work     = trim(entry(ia,lvc_listfields))
             lvc_name     = TRIM(ENTRY(ia, ipc_names))
             lvc_format   = ""
             lvc_subquery = ""
         .
      IF lvc_work BEGINS "#" THEN DO:
         IF NUM-ENTRIES(lvc_work, "(") GT 1 THEN DO:
            lvc_subquery = TRIM(ENTRY(2, lvc_work, "("), ")").
            lvc_work = ENTRY(1, lvc_work, "(").
         END.
      END.

      if num-entries(lvc_work," ") > 1 then do: 
         /* This field has a format string defined, extract it */
         if entry(2,lvc_work," ") = "format" and num-entries(lvc_work," ") > 2 then do:
               lvc_format = trim(substring(lvc_work,
                                           (if index(lvc_work,"'") > 0 and index(lvc_work,'"') > 0 then
                                              (min(index(lvc_work,"'"),index(lvc_work,'"')))
                                            else if index(lvc_work,"'") > 0 then
                                                   (index(lvc_work,"'"))
                                            else if index(lvc_work,'"') > 0 then
                                                   (index(lvc_work,'"'))
                                            else 1))).
               if lvc_format begins '"' then
                  lvc_format = trim(lvc_format,'"').
               else if lvc_format begins "'" then
                  lvc_format = trim(lvc_format,"'").
         end. /* if entry(2,lvc_work," ") = "format" ... */

         lvc_work = entry(1,lvc_work," ").

      end. /* if num-entries(lvc_work," ") > 1 */


      IF lvc_work BEGINS "#" THEN DO:
         /* Special field */
         create ttFields.
         assign ttFields.tOrder    = ia
                ttFields.tTable    = 0
                ttFields.tDBField  = SUBSTRING(lvc_work, 2)
                ttFields.tField    = lvc_name
                ttFields.tFormat   = lvc_format
                ttFields.cSubQuery = lvc_subquery.
      END. /* Special field */
      ELSE DO:
         /* Normal field */
         IF INDEX(lvc_work, "[") GT 0 THEN DO:
            /* Array field */
            cSubScript = SUBSTRING(lvc_work, INDEX(lvc_work, "[") + 1,
                                   INDEX(lvc_work, "]") - (INDEX(lvc_work, "[") + 1)).
            lvc_work = SUBSTRING(lvc_work, 1, INDEX(lvc_work, "[") - 1).
            cSubScript = TRIM(cSubScript, "[]").
            ASSIGN
               iIndex = INTEGER(cSubScript)
            NO-ERROR.
            IF ERROR-STATUS:ERROR = TRUE THEN DO:
               addError({&errSubScriptError}, "Array index is not an integer.").
               return error "{&errSubScriptError};Array index is not an integer.".
            END.
         END. /* Array field */
         ELSE DO:
            /* No array field */
            iIndex = 0.
         END. /* No array field */
         
         case num-entries(lvc_work,"."):
            when 1 then do:   /* This is a field without table/database declaration */
               find last ttTables no-error.
               if avail ttTables then do:
                  if ttTables.tOrder > 1 THEN DO:
                     addError({&errUnqualifiedField}, "Unqualified field. You have to specify the tablename when using joins.").
                     return error "{&errUnqualifiedField};Unqualified field."
                          + " You have to specify the tablename when using joins.".
                  END.
                  else do:
                     create ttFields.
                     assign ttFields.tOrder = ia                            /* Order it shows up in the declaration */
                            ttFields.tTable = ttTables.tOrder
                            ttFields.tDBField = lvc_work /* Corresponding DB field */
                            ttFields.tField = lvc_name                     /* The name of the field in the TT */
                            ttFields.tFormat = lvc_format                   /* The format string */
                            ttFields.iIndex  = iIndex
                     .
                  end.
               end. /* if avail ttTables */
            end.
            when 2 then do: /* This is a field with a table declaration */
               find ttTables where ttTables.tTable = entry(1,lvc_work,".") no-error.
               if avail ttTables then do:
                  create ttFields.
                  assign ttFields.tOrder = ia
                         ttFields.tTable = ttTables.tOrder
                         ttFields.tDBField = entry(2,lvc_work,".")
                         ttFields.tField   = lvc_name
                         ttFields.tFormat = lvc_format
                         ttFields.iIndex  = iIndex
                  .
               end. /* if avail ttTables */
               ELSE DO:
                  addError({&errUnknownTable}, "Unknown table " + entry(1,lvc_work,".") + " for field " + entry(2,lvc_work,".")).
                  return error "{&errUnknownTable};"
                     + "Unknown table " + entry(1,lvc_work,".") 
                               + " for field " + entry(2,lvc_work,".") .
               END.
            end.
            when 3 then do:  /* This is a field with a table and database declaration */
               find first ttTables where ttTables.tDb    = entry(1,lvc_work,".")
                                     and ttTables.tTable = entry(2,lvc_work,".") no-error.
               if avail ttTables then do:
                  create ttFields.
                  assign ttFields.tOrder = ia
                         ttFields.tTable = ttTables.tOrder
                         ttFields.tDBField = entry(3,lvc_work,".")
                         ttFields.tField   = lvc_name
                         ttFields.tFormat  = lvc_format
                         ttfields.iIndex   = iIndex
                  .
               end. /* if avail ttTables */
               ELSE DO:
                  addError({&errUnknownTable}, "Unknown table " + entry(1, lvc_work, ".") + "." + entry(2,lvc_work,".") + " for field " + entry(3,lvc_work,".")).
                  return error "{&errUnknownTable};"
                     + "Unknown table " + entry(1, lvc_work, ".") + "." +
                     entry(2,lvc_work,".") + " for field " + entry(3,lvc_work,".").
               END.
            end.
            OTHERWISE DO:
               addError({&errUnknownTable}, "Too many entries in field specification " + lvc_work).
               return error "{&errUnknownTable};"
                  + "Too many entries in field specification " + lvc_work.
            END.
         end case.
      END.
   end. /* repeat ia = 1 to num-entries(lvc_listfields): */

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-setProperties) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE setProperties Procedure 
PROCEDURE setProperties :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE INPUT  PARAMETER ipcPropertyList AS CHARACTER   NO-UNDO.
DEFINE VARIABLE iProperty AS INTEGER     NO-UNDO.
DEFINE VARIABLE cPropNameValue AS CHARACTER   NO-UNDO.
DEFINE VARIABLE cPropName AS CHARACTER   NO-UNDO.
DEFINE VARIABLE cPropValue AS CHARACTER   NO-UNDO.

DO iProperty = 1 TO NUM-ENTRIES(ipcPropertyList,"~n"):
   cPropNameValue = ENTRY(iProperty,ipcPropertyList,"~n").
   cPropName = ENTRY(1,cPropNameValue,"=").
   cPropValue = SUBSTRING(cPropNameValue,INDEX(cPropNameValue, "=") + 1).
   FIND ttProperty WHERE ttProperty.PropertyName = cPropName NO-ERROR.
   IF NOT AVAILABLE ttProperty THEN DO:
      CREATE ttProperty.
      ttProperty.PropertyName = cPropName.
   END.
   ASSIGN ttProperty.PropertyValue = cPropValue.
END.


END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

/* ************************  Function Implementations ***************** */

&IF DEFINED(EXCLUDE-addError) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION addError Procedure 
FUNCTION addError RETURNS LOGICAL
  ( /* parameter-definitions */ 
    INPUT ipiErrorNr AS INT,
    INPUT ipcError   AS CHAR
  ) :
/*------------------------------------------------------------------------------
  Purpose:  Adds an error to the error temp-table
    Notes:  
------------------------------------------------------------------------------*/
DEFINE VARIABLE hField AS HANDLE     NO-UNDO.

   IF VALID-HANDLE(hBufError) = FALSE THEN DO:
      /* Initialize dynamic error temp-table */
      CREATE TEMP-TABLE hTTError.
      hTTError:ADD-NEW-FIELD("SeqNr", "INTEGER", 0, "z9").
      hTTError:ADD-NEW-FIELD("Number", "INTEGER", 0, "-zzzz").
      hTTError:ADD-NEW-FIELD("Error", "CHARACTER",0,"X(255)").
      
      hTTError:TEMP-TABLE-PREPARE("Error").
      hBufError = hTTError:DEFAULT-BUFFER-HANDLE.
   END.

   hBufError:BUFFER-CREATE().

   ASSIGN
      iSeqNr = iSeqNr + 1
   .

   hField = hBufError:BUFFER-FIELD("SeqNr").
   hField:BUFFER-VALUE = iSeqNr.

   hField = hBufError:BUFFER-FIELD("Number").
   hField:BUFFER-VALUE = ipiErrorNr.

   hField = hBufError:BUFFER-FIELD("Error").
   hfield:BUFFER-VALUE = ipcError.

   RETURN TRUE.

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-addSpecialField) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION addSpecialField Procedure 
FUNCTION addSpecialField RETURNS LOGICAL
  ( /* parameter-definitions */ 
    INPUT iphTT      AS HANDLE,
    INPUT ipcSpecial AS CHAR,
    INPUT ipcName    AS CHAR,
    INPUT ipcFormat  AS CHAR
     ) :
/*------------------------------------------------------------------------------
  Purpose:  Adds a special field to the temp-table
    Notes:  
------------------------------------------------------------------------------*/
DEFINE VARIABLE lOk       AS LOGICAL     NO-UNDO.
DEFINE VARIABLE cSubQuery AS CHARACTER   NO-UNDO.

   IF NUM-ENTRIES(ipcSpecial, "(") GT 1 THEN DO:
      ASSIGN
         cSubQuery  = ENTRY(2, ipcSpecial, "(")
         ipcSpecial = ENTRY(1, ipcSpecial, "(")
         cSubQuery  = TRIM(cSubQuery, ")")
      .
   END.

   CASE ipcSpecial:
      WHEN "ROWIDS" THEN DO:
         /* Create a field to keep the rowids of the database records */
         ASSIGN
            lOk = iphTT:ADD-NEW-FIELD(ipcName, "CHARACTER", 0, ?, ?, ipcSpecial)
         .
      END. /* Create a field to keep the rowids of the database records */
      WHEN "LINE-COUNTER" THEN DO:
         /* Add an integer field to count records */
         ASSIGN
            lOk = iphTT:ADD-NEW-FIELD(ipcName, "INTEGER", 0, ipcFormat, ?, ipcSpecial)
         .
      END. /* Add an integer field to count records */
      WHEN "COUNT" THEN DO:
         /* Count on sub query */
         ASSIGN
            lOk = iphTT:ADD-NEW-FIELD(ipcName,  "INTEGER", 0, ipcFormat, ?, SUBSTITUTE("&1 &2", ipcSpecial, ipcName), SUBSTITUTE("&1!&2", ipcSpecial, ipcName))
         .
/*          ADD-NEW-FIELD( field-name-exp ,                           */
/*                         datatype-exp [ ,                           */
/*                         extent-exp [ ,                             */
/*                         format-exp     [ ,                         */
/*                         initial-exp [ ,                            */
/*                         label-exp [ , column-label-exp ] ] ] ] ] ) */


      END. /* Count on sub query */
      WHEN "TOGGLE" THEN DO:
         /* Add a logical field to allow selection */
         ASSIGN
            lOk = iphTT:ADD-NEW-FIELD(ipcName, "LOGICAL", 0, ipcFormat, FALSE, ipcName)
         .
      END. /* Add a logical field to allow selection */
      OTHERWISE DO:
         lOk = FALSE.
      END.
   END CASE.

   RETURN lOk.   /* Function return value. */

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getBaseNames) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION getBaseNames Procedure 
FUNCTION getBaseNames RETURNS CHARACTER
  ( /* parameter-definitions */ 
    INPUT ipcFields AS CHAR 
  ) :
/*------------------------------------------------------------------------------
  Purpose:  Extracts the base field name from a list of database fields
    Notes:  
------------------------------------------------------------------------------*/
DEFINE VARIABLE iFieldNr   AS INTEGER    NO-UNDO.
DEFINE VARIABLE cFullField AS CHARACTER  NO-UNDO.
DEFINE VARIABLE cFieldName AS CHARACTER  NO-UNDO.
DEFINE VARIABLE cFieldList AS CHARACTER  NO-UNDO.

   DO iFieldNr = 1 TO NUM-ENTRIES(ipcFields):
      cFullField = ENTRY(iFieldNr, ipcFields).
      IF NUM-ENTRIES(cFullField, ".") GT 1 THEN DO:
         cFieldName = ENTRY(NUM-ENTRIES(cFullField, "."), cFullField, ".").
      END.
      ELSE DO:
         cFieldName = cFullField.
      END.
      /* Remove index brackets */
      cFieldName = REPLACE(REPLACE(cFieldName, "[", ""), "]", "").
      IF iFieldNr GT 1 THEN DO:
         cFieldList = cFieldList + ",".
      END.
      cFieldList = cFieldList + cFieldName.
   END.

   RETURN cFieldList.

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getEndChar) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION getEndChar Procedure 
FUNCTION getEndChar RETURNS INTEGER
  ( /* parameter-definitions */
    INPUT ipcString AS CHARACTER,
    INPUT ipiStart  AS INTEGER
  ) :
/*------------------------------------------------------------------------------
  Purpose:  Returns the end character of a variabile starting at a position
    Notes:  End characters are: [ ], )
------------------------------------------------------------------------------*/
DEFINE VARIABLE iChar     AS INTEGER     NO-UNDO.
DEFINE VARIABLE cChar     AS CHARACTER   NO-UNDO.
DEFINE VARIABLE cEndChars AS CHARACTER   NO-UNDO INITIAL " )*-+/".

   DO iChar = ipiStart + 1 TO LENGTH(ipcString):
      cChar = SUBSTRING(ipcString, iChar, 1).

      IF INDEX(cEndChars, cChar) NE 0 THEN
         RETURN iChar.
   END.

   RETURN LENGTH(ipcString) + 1.

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-getProperty) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION getProperty Procedure 
FUNCTION getProperty RETURNS CHARACTER
  (INPUT ipcName AS CHAR /* parameter-definitions */ ) :
/*------------------------------------------------------------------------------
  Purpose:  
    Notes:  
------------------------------------------------------------------------------*/

   FIND ttProperty WHERE ttProperty.PRopertyName = ipcName NO-LOCK NO-ERROR.
   IF AVAILABLE ttProperty THEN
     RETURN ttProperty.PropertyValue.
   ELSE
     RETURN "".   /* Function return value. */

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-openQuote) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION openQuote Procedure 
FUNCTION openQuote RETURNS LOGICAL
  ( ipc_string as character ) :
/*------------------------------------------------------------------------------
  Purpose:  Checks a input string for open quotes and other things
    Notes:  This is made for the purpose that functions can use comma-separated
            items that could interfere with the extraction of tables in 
            dynComboTNG
------------------------------------------------------------------------------*/
   define variable ia         as integer    no-undo.
   define variable lvi_count  as integer    no-undo.
   define variable lvc_quoter as character  no-undo.
   define variable lvl_open   as logical    no-undo.

   /* First we find the first occurance of any of the allowed quotecharacters */
   do ia = 1 to length(ipc_string):
      if substring(ipc_string,ia,1) = "(" 
            and not lvl_open then
         lvi_count = lvi_count + 1.
      else if substring(ipc_string,ia,1) = ")"
            and not lvl_open then
         lvi_count = lvi_count - 1.
      else if substring(ipc_string,ia,1) = '"'
            and not (lvl_open and lvc_quoter <> '"') then
         assign lvl_open = not lvl_open
                lvc_quoter = '"'.
      else if substring(ipc_string,ia,1) = "'"
            and not (lvl_open and lvc_quoter <> "'") then
         assign lvl_open = not lvl_open
                lvc_quoter = "'".
   end.

   if lvl_open or lvi_count > 0 then
      return true.
   else
      return false.

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

