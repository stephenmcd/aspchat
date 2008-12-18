<%


' CONSTANTS


' ********************************************************************************************************************************************
Const cBUFFER_LOAD = 3 ' number of seconds between post retrievals
Const cBUFFER_EVENT = 10 ' maximum number of posts stored in memory
Const cBUFFER_EXIT = 10 ' number of seconds on top of cBUFFER_LOAD to allow for each user's client-siode recursive getter call to take before timing user out
Const cCHAR_DELIMIT = "|" ' data delimter for passing arrays between client and server
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cDEFAULT_COLOUR = "#000"
Const cCSS_CLASS_POST = "post"
Const cCSS_CLASS_PALETTE = "palette"
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cMSG_ENTER = " enters" ' entry message
Const cMSG_EXIT = " leaves" ' exit message
Const cMSG_POST_PUBLIC = " says" ' public post message prefix
Const cMSG_POST_PRIVATE = " says privately" ' private post message prefix
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cOP_NAME = "o" ' query-string operation name
Const cOP_VAL_GET = "g" ' query-string operation value for post getting
Const cOP_VAL_POST = "p" ' query-string operation value for posting
Const cOP_VAL_COLOUR = "c" ' query-string operation value for colour
Const cOP_VAL_FRAME = "f" ' query-string operation value for initial iframe
Const cOP_VAL_USER = "u" ' query-string operation value for initial user list retrieval
Const cOP_VAL_VALID = "v" ' query-string operation value for validating availability of user name
Const cOP_VAL_EXIT = "x" ' query-string operation value for exiting
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cEVENT_TYPE_POST = 0 ' event type for post
Const cEVENT_TYPE_ENTER = 1 ' event type for user entry
Const cEVENT_TYPE_EXIT = 2 ' event type for user exit
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cEVENT_FIELD_INDEX = 0 ' array index for event index
Const cEVENT_FIELD_TYPE = 1 ' array index for event type
Const cEVENT_FIELD_TIME = 2 ' array index for event time
Const cEVENT_FIELD_USER = 3 ' array index for event user name
Const cEVENT_FIELD_DATA = 4 ' array index for event data (post)
Const cEVENT_FIELD_COLOUR = 5 ' array index for event data (post)
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
Const cUSER_FIELD_INDEX = 0 ' array index for user current time
Const cUSER_FIELD_TIME = 1 ' array index for user current time
Const cUSER_FIELD_NAME = 2 ' array index for user name
' ********************************************************************************************************************************************


%>