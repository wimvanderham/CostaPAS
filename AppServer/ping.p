
/*------------------------------------------------------------------------
    File        : ping.p
    Purpose     : Return information from the AppServer

    Syntax      :

    Description : Simple ping.p procedure to run on the AppServer

    Author(s)   : Wim van der Ham (WITS)
    Created     : Fri Jan 17 12:27:55 CET 2025
    Notes       :
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

BLOCK-LEVEL ON ERROR UNDO, THROW.

USING Progress.Json.*.
USING Progress.Json.ObjectModel.*.

DEFINE OUTPUT PARAMETER opcOutput AS LONGCHAR NO-UNDO.

DEFINE VARIABLE oOutput  AS JsonObject.
DEFINE VARIABLE oPath    AS JsonArray.
DEFINE VARIABLE oDB      AS JsonArray.
DEFINE VARIABLE iPath    AS INTEGER    NO-UNDO.
DEFINE VARIABLE iDB      AS INTEGER    NO-UNDO.
DEFINE VARIABLE lOk      AS LOGICAL    NO-UNDO.
DEFINE VARIABLE cMessage AS CHARACTER  NO-UNDO.

/* ********************  Preprocessor Definitions  ******************** */


/* ***************************  Main Block  *************************** */

oOutput = NEW JsonObject().

ASSIGN 
   lOk      = TRUE
   cMessage = SUBSTITUTE ("Hello World!")
   .
    
oOutput:ADD ("OK", lOk).
oOutput:ADD ("Message", cMessage).

oPath = NEW JsonArray().
DO iPath = 1 TO NUM-ENTRIES (PROPATH):
   oPath:Add(ENTRY (iPath, PROPATH)).
END.

oOutput:ADD ("PROPATH", oPath).

oDB = NEW JsonArray().

DO iDB = 1 TO NUM-DBS:
   oDB:ADD (LDBNAME(iDB)).
END.
oOutput:ADD("DB", oDB).

FILE-INFO:FILE-NAME = ".".
oOutput:ADD ("WKRDIR", FILE-INFO:FULL-PATHNAME).

oOutput:WRITE (INPUT-OUTPUT opcOutput, TRUE).

CATCH oError AS Progress.Lang.Error :
   ASSIGN 
      lOk      = FALSE 
      cMessage = oError:GetMessage(1)
      .

   oOutput = NEW JsonObject().
   oOutput:ADD ("OK",lok).
   oOutput:ADD ("Message",cMessage).
   oOutput:WRITE (INPUT-OUTPUT opcOutput, TRUE).
   
END CATCH.


