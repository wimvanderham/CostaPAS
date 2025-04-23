
/*------------------------------------------------------------------------
    File        : testpasoe.p
    Purpose     : Testare connessione e ping di un PASOE

    Syntax      :

    Description : Test connessione e ping di un PASOE

    Author(s)   : Wim van der Ham (WITS)
    Created     : Fri Jan 17 11:56:09 CET 2025
    Notes       :
       
    Modified by : Wim van der Ham (WITS)
    Modified on : 23 Apr 2025
    Reason      : Added option to use "classic AppServer" in Direct Connect       
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

BLOCK-LEVEL ON ERROR UNDO, THROW.

/* ********************  Preprocessor Definitions  ******************** */
DEFINE VARIABLE cHost             AS CHARACTER NO-UNDO.
DEFINE VARIABLE iPort             AS INTEGER   NO-UNDO.
DEFINE VARIABLE cProtocol         AS CHARACTER NO-UNDO INIT "apsv".
DEFINE VARIABLE cConnectionString AS CHARACTER NO-UNDO.
DEFINE VARIABLE hAppServer        AS HANDLE    NO-UNDO.
DEFINE VARIABLE lOk               AS LOGICAL   NO-UNDO.
DEFINE VARIABLE lcOutput          AS LONGCHAR  NO-UNDO.

/* ***************************  Main Block  *************************** */

DEFINE FRAME fr-Parameters
   cHost       LABEL "Host"      COLON 20 HELP "Host name or IP Address" FORMAT "X(120)" 
      VIEW-AS FILL-IN SIZE-CHARS 40 BY 1
   iPort       LABEL "HTTP Port" COLON 20 HELP "HTTP Port number"        FORMAT "zzzz9" 
   cProtocol   LABEL "Protocol"  COLON 20 HELP "Protocol"                FORMAT "X(4)" 
      VIEW-AS RADIO-SET HORIZONTAL RADIO-BUTTONS "AppServer", "apsv", "Classic AppServer", "cas", "REST", "rest", "SOAP", "soap", "WEB", "web"
   cConnectionString LABEL "Connection String" COLON 20 HELP "Full connection string" FORMAT "X(120)" 
      VIEW-AS FILL-IN SIZE-CHARS 40 BY 1
   SKIP (1)
   lcOutput    LABEL "Output"    COLON 20 HELP "Output from ping.p" FORMAT "X(200)"
      VIEW-AS EDITOR LARGE SIZE-CHARS 77 BY 18
WITH TITLE " Connection Parameters " CENTERED ROW 3 SIDE-LABELS 1 DOWN WIDTH 100.
   
DISPLAY 
   cHost
   iPort
   cProtocol
   cConnectionString
WITH FRAME fr-Parameters.

CREATE SERVER hAppServer.

REPEAT WITH FRAME fr-Parameters:
   UPDATE 
      cHost
      iPort
      cProtocol
   .
   
   CASE cProtocol:
      WHEN "cas" THEN
         cConnectionString = SUBSTITUTE ("-URL AppServerDC://&1:&2",
                                         cHost,
                                         iPort).
      WHEN "apsv" THEN 
         cConnectionString = SUBSTITUTE ("-URL http://&1:&2/&3",
                                         cHost,
                                         iPort,
                                         cProtocol) .
      OTHERWISE DO:
         MESSAGE SUBSTITUTE ("Protocollo '&1' non implementato.", cProtocol)
         VIEW-AS ALERT-BOX WARNING.
         UNDO, RETRY.
      END.
   END CASE.
                                                  
   DISPLAY
      cConnectionString
   .
   
   STATUS DEFAULT "Connecting to AppServer...".
   
   ASSIGN 
      lOk = hAppServer:CONNECT (cConnectionString)
   NO-ERROR.
   STATUS DEFAULT "".
   
   IF lOk EQ FALSE THEN DO:
      MESSAGE ERROR-STATUS:GET-MESSAGE (1)
      VIEW-AS ALERT-BOX WARNING.
      UNDO, RETRY.
   END.
   
   MESSAGE "Connected?" hAppServer:CONNECTED ()
   VIEW-AS ALERT-BOX.
                                      
   IF hAppServer:CONNECTED() EQ TRUE THEN DO:
      RUN ping.p ON hAppServer
         (OUTPUT lcOutput).
         
      DISPLAY 
      lcOutput
      WITH FRAME fr-Parameters.
      
      lcOutput:READ-ONLY = TRUE.
      
      UPDATE 
      lcOutput
      WITH FRAME fr-Parameters.
   END.
END.
   
CATCH oError AS Progress.Lang.Error :
   DEFINE VARIABLE iMessage AS INTEGER   NO-UNDO.
   DEFINE VARIABLE cMessage AS CHARACTER NO-UNDO.
   
   IF oError:NumMessages GE 1 THEN DO:
      DO iMessage = 1 TO oError:NumMessages:
         cMessage = SUBSTITUTE ("&1&2&3. &4",
                                cMessage,
                                (IF cMessage NE "" THEN "~n" ELSE ""),
                                iMessage,
                                oError:GetMessage(iMessage)).
      END.
                                
      MESSAGE 
      cMessage
      VIEW-AS ALERT-BOX ERROR.
   END.   
END CATCH.


FINALLY:
   IF VALID-HANDLE(hAppServer) EQ TRUE THEN 
   DO:
      IF hAppServer:CONNECTED () EQ TRUE THEN 
         hAppServer:DISCONNECT().
      DELETE OBJECT hAppServer.
   END.
END FINALLY.

