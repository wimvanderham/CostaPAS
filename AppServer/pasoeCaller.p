
/*------------------------------------------------------------------------
    File        : pasoeCaller.p
    Purpose     : Wrapper procedure to call PASOE

    Syntax      :

    Description : Wrapper procedure to call PASOE

    Author(s)   : Wim van der Ham
    Created     : Fri May 15 09:47:01 CEST 2026
    Notes       :
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

BLOCK-LEVEL ON ERROR UNDO, THROW.

USING Progress.Json.ObjectModel.ObjectModelParser.
USING Progress.Json.*.
USING Progress.Json.ObjectModel.*.

DEFINE INPUT  PARAMETER ipcJSONIn  AS LONGCHAR NO-UNDO.
DEFINE OUTPUT PARAMETER opcJSONOut AS LONGCHAR NO-UNDO.

{pasoeCaller.i}

DEFINE VARIABLE oJSONIn     AS JSONObject        NO-UNDO.
DEFINE VARIABLE oJSONOut    AS JSONObject        NO-UNDO.
DEFINE VARIABLE oParser     AS ObjectModelParser NO-UNDO.
DEFINE VARIABLE oJsonArray  AS JsonArray         NO-UNDO.
DEFINE VARIABLE oJsonObject AS JsonObject        NO-UNDO.
DEFINE VARIABLE oData       AS JsonArray         NO-UNDO.

DEFINE VARIABLE lcInput   AS LONGCHAR  NO-UNDO.
DEFINE VARIABLE lOk       AS LOGICAL   NO-UNDO.
DEFINE VARIABLE cMessage  AS CHARACTER NO-UNDO.
DEFINE VARIABLE cCall     AS CHARACTER NO-UNDO.


DEFINE VARIABLE hTT       AS HANDLE NO-UNDO.
DEFINE VARIABLE dtStart   AS DATETIME  NO-UNDO.
DEFINE VARIABLE dtEnd     AS DATETIME  NO-UNDO.
DEFINE VARIABLE iDuration AS INTEGER   NO-UNDO.
DEFINE VARIABLE iSeconds  AS INTEGER   NO-UNDO.
DEFINE VARIABLE cTime     AS CHARACTER NO-UNDO.

/* ********************  Preprocessor Definitions  ******************** */


/* ***************************  Main Block  *************************** */

FIX-CODEPAGE (lcInput) = "UTF-8".

lcInput = ipcJSONIn.

/* Parse the JSON file into a JSON object */
ASSIGN 
   lOk = TRUE 
   cMessage = SUBSTITUTE ("JSON Object parsed successfully.")
.

oJSONOut = NEW JsonObject().

oParser = NEW ObjectModelParser().
oJSONIn = CAST(oParser:Parse(lcInput), "JsonObject").

cCall = oJSONIn:GetCharacter("call").

dtStart = NOW.
CASE cCall:
   WHEN "roundtrip" OR 
   WHEN "" THEN DO:
      ASSIGN 
         lOk = TRUE 
         cMessage = "Roundtrip"
      .
      oJSONOut:ADD ("OK",      lOk).
      oJSONOut:ADD ("Message", cMessage).
      oJSONOut:ADD ("JSONIn",  oJSONIn).
      
   END.      
   WHEN "query" THEN DO:
      
      RUN callQuery
         (OUTPUT hTT,
          OUTPUT lOk,
          OUTPUT cMessage).
          
      oJSONOut:ADD ("OK",      lOk).
      oJSONOut:ADD ("Message", cMessage).
      
      IF lOk EQ TRUE THEN DO:
         oData = NEW JsonArray().
   
         hTT:WRITE-JSON ("JsonArray", oData).
         
         oJSONOut:ADD ("Data",  oData).
         IF hTT:PRIVATE-DATA NE ? THEN
            oJSONOut:ADD ("PrivateData", hTT:PRIVATE-DATA).
      END.
      ELSE DO:
         oData = NEW JsonArray().
   
         hTT:WRITE-JSON ("JsonArray", oData).
         
         oJSONOut:ADD ("Errors",  oData).
         IF hTT:PRIVATE-DATA NE ? THEN
            oJSONOut:ADD ("PrivateData", hTT:PRIVATE-DATA).
         
      END.      
   END.
   OTHERWISE DO:
      ASSIGN 
         lOk = FALSE 
         cMessage = SUBSTITUTE ("Call '&1' not implemented.", cCall)
      .
   END.         
END.

/* Add duration information */
dtEnd = NOW.
iDuration = INTERVAL (dtEnd, dtStart, "milliseconds").
iSeconds  = INTERVAL (dtEnd, dtStart, "seconds").
cTime     = STRING (iSeconds, "HH:MM:SS").

oJSONOut:ADD ("Start", dtStart).
oJSONOut:ADD ("End", dtEnd).
oJSONOut:ADD ("Duration", iDuration).
oJSONOut:ADD ("Time", cTime).

/* Prepare output */
oJSONOut:WRITE (INPUT-OUTPUT opcJSONOut, TRUE, "UTF-8").


CATCH oError AS Progress.Lang.Error :
   DEFINE VARIABLE iMessage AS INTEGER   NO-UNDO.
   
   ASSIGN 
      lOk = FALSE 
      cMessage = "** Errors:~n"
   .
   DO iMessage = 1 TO oError:NumMessages:
      cMessage = SUBSTITUTE ("&1&2&3. &4",
         cMessage,
         (IF cMessage NE "" THEN "~n" ELSE ""),
         iMessage,
         oError:GetMessage(iMessage)).
   END.
                             
   oJSONOut:ADD ("OK",      lOk).
   oJSONOut:ADD ("Message", cMessage).
   oJSONOut:ADD ("JSONIN",  oJSONIn).

   oJSONOut:WRITE (INPUT-OUTPUT opcJSONOut).
   
END CATCH.




/* **********************  Internal Procedures  *********************** */

PROCEDURE callQuery:
   /*------------------------------------------------------------------------------
    Purpose:
    Notes:
   ------------------------------------------------------------------------------*/
   DEFINE OUTPUT PARAMETER ophTT      AS HANDLE    NO-UNDO.
   DEFINE OUTPUT PARAMETER oplOk      AS LOGICAL   NO-UNDO.
   DEFINE OUTPUT PARAMETER opcMessage AS CHARACTER NO-UNDO.

   DEFINE VARIABLE oParameters AS JsonArray.

   DEFINE VARIABLE cQueryString AS CHARACTER NO-UNDO.
   DEFINE VARIABLE cFieldList   AS CHARACTER NO-UNDO.
   DEFINE VARIABLE cFieldNames  AS CHARACTER NO-UNDO.
   DEFINE VARIABLE cTTName      AS CHARACTER NO-UNDO.
   DEFINE VARIABLE hTT          AS HANDLE    NO-UNDO.
   
   oParameters = NEW JsonArray().
   oParameters = oJSONIn:GetJsonArray("Parameters").
   
   TEMP-TABLE ttParameter:READ-JSON ("JsonArray", oParameters).

   FIND ttParameter
   WHERE ttParameter.cName EQ "QueryString".
   cQueryString = ttParameter.cValue.
   
   FIND  ttParameter
   WHERE ttParameter.cName EQ "FieldList".
   cFieldList = ttParameter.cValue.
   
   FIND  ttParameter
   WHERE ttParameter.cName EQ "FieldNames".
   cFieldNames = ttParameter.cValue.
   
   FIND  ttParameter
   WHERE ttParameter.cName EQ "TTName".
   cTTName = ttParameter.cValue.
      
   
   RUN sy/p/dynquerytt.p
      (INPUT  cQueryString,
       INPUT  cFieldList,
       INPUT  cFieldNames,
       INPUT  cTTName,
       OUTPUT TABLE-HANDLE ophTT).

   IF ophTT:NAME NE cTTName THEN DO:
      ASSIGN 
         oplOk = FALSE 
         opcMessage = "Errors during call to dynquerytt.p"
      .
   END.
   ELSE DO:
      ASSIGN
         oplOk      = TRUE 
         opcMessage = SUBSTITUTE ("Query executed succesfully.")
      .
   END.
   
END PROCEDURE.
