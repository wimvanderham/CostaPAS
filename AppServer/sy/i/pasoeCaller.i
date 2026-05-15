
/*------------------------------------------------------------------------
    File        : pasoeCaller.i
    Purpose     : Include file with temp-table definitions for pasoeCaller

    Syntax      :

    Description : Include file with temp-table definitions for pasoeCaller

    Author(s)   : Wim van der Ham
    Created     : Fri May 15 10:47:24 CEST 2026
    Notes       :
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */
DEFINE TEMP-TABLE ttParameter SERIALIZE-NAME "Parameter"
   FIELD cName  AS CHARACTER SERIALIZE-NAME "Name"
   FIELD cValue AS CHARACTER SERIALIZE-NAME "Value"
INDEX indName IS UNIQUE cName.


/* ********************  Preprocessor Definitions  ******************** */


/* ***************************  Main Block  *************************** */
