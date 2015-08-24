<%


' FUNCTIONS


' ********************************************************************************************************************************************
' adds a value to an array
Sub addToArray(varArray, varValue)

  ' either initialize array or resize
  If not IsArray(varArray) Then
    ReDim varArray(0)
  Else
    ReDim preserve varArray(UBound(varArray) + 1)
  End If

  ' add value
  varArray(UBound(varArray)) = varValue

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' adds an event to the app event array
Sub addEvent(intType, strTime, strUserName, varEventData)

  Dim intEvent
  Dim arrEvent

  With Application

    ' initialize new event list with this event
    Call addToArray(arrEvent, Array(.Contents("event_index") + 1, intType, strTime, strUserName, varEventData))

    ' add all existing events to new event list, bar the earliest
    For intEvent = LBound(.Contents("event_list")) to UBound(.Contents("event_list"))
      If intEvent < cBUFFER_EVENT - 1 Then
        Call addToArray(arrEvent, .Contents("event_list")(intEvent))
      End If
    Next

    ' set app event array to new event list and increment app event index
    Call .lock()
    .Contents("event_list") = arrEvent
    .Contents("event_index") = .Contents("event_index") + 1
    Call .unlock()

  End With

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' maintains the app level user list
Sub setUsers(strUserName)

  Dim arrOldUser
  ReDim arrNewUser(-1)

  ' initialize new user list with current user and time
  If Len(strUserName) > 0 Then
    Call addToArray(arrNewUser, Array(Session.Contents("user_index"), Now(), strUserName))
  End If

  ' go thru current user list, building new user list excluding timed out users
  For each arrOldUser in Application.Contents("user_list")
    If arrOldUser(cUSER_FIELD_NAME) <> strUserName Then
      If DateDiff("s", arrOldUser(cUSER_FIELD_TIME), Now()) > (cBUFFER_LOAD + cBUFFER_EXIT) Then
        ' add exit event
        Call addEvent(cEVENT_TYPE_EXIT, arrOldUser(cUSER_FIELD_TIME), arrOldUser(cUSER_FIELD_NAME), "")
      Else
        ' add user to new user list
        Call addToArray(arrNewUser, arrOldUser)
      End If
    End If
  Next

  ' set app user list to new user list
  With Application
    Call .lock()
    .Contents("user_list") = arrNewUser
    Call .unlock()
  End With

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' return array of events with higher index than user's event index
Function getEvents(intUserEventIndex)

  Dim intEvent, intUser
  Dim strPrivateTo
  Dim varEventData, arrUser
  Dim arrOneEvent
  ReDim arrAllEvent(-1)

  For intEvent = UBound(Application.Contents("event_list")) to LBound(Application.Contents("event_list")) step -1
    arrOneEvent = Application.Contents("event_list")(intEvent)
    varEventData = ""

    ' build event data string
    If CLng(arrOneEvent(cEVENT_FIELD_INDEX)) > CLng(intUserEventIndex) Then

      If arrOneEvent(cEVENT_FIELD_TYPE) = cEVENT_TYPE_POST Then

        If InStr(arrOneEvent(cEVENT_FIELD_DATA), cMSG_POST_PRIVATE) = 1 Then
          ' get list of user IDs that post is to
          strPrivateTo = Mid(arrOneEvent(cEVENT_FIELD_DATA), Len(cMSG_POST_PRIVATE) + 1, InStr(Len(cMSG_POST_PRIVATE), arrOneEvent(cEVENT_FIELD_DATA), ": ") - Len(cMSG_POST_PRIVATE) - 1)
          If arrOneEvent(cEVENT_FIELD_USER) = Session.Contents("user_name") Then

            ' private post from user
            For each arrUser in Application.Contents("user_list")
              If InStr(cCHAR_DELIMIT & strPrivateTo & cCHAR_DELIMIT, cCHAR_DELIMIT & arrUser(cUSER_FIELD_INDEX) & cCHAR_DELIMIT) > 0 Then
                Call addToArray(varEventData, arrUser(cUSER_FIELD_NAME))
              End If
            Next
            If IsArray(varEventData) Then
              For intUser = LBound(varEventData) to UBound(varEventData)
                Select Case intUser
                  Case LBound(varEventData)
                    ' first name
                    varEventData(intUser) = " to " & varEventData(intUser)
                  Case UBound(varEventData)
                    ' last name
                    varEventData(intUser) = " and " & varEventData(intUser)
                  Case Else
                    varEventData(intUser) = ", " & varEventData(intUser)
                End Select
              Next
              varEventData = Replace(arrOneEvent(cEVENT_FIELD_DATA), strPrivateTo, Join(varEventData, ""), 1, 1, 1)
            End If

          ElseIf InStr(cCHAR_DELIMIT & strPrivateTo & cCHAR_DELIMIT, cCHAR_DELIMIT & Session.Contents("user_index") & cCHAR_DELIMIT) > 0 Then

            ' private post to user
            varEventData = Replace(arrOneEvent(cEVENT_FIELD_DATA), strPrivateTo, "", 1, 1, 1)

          End If
        Else
          ' public post
          varEventData = arrOneEvent(cEVENT_FIELD_DATA)
        End If

        ' add time and user name post is from
        If Len(varEventData) > 0 Then
          varEventData = "[" & LCase(FormatDateTime(arrOneEvent(cEVENT_FIELD_TIME), vbLongTime)) & "] " & Server.htmlEncode(arrOneEvent(cEVENT_FIELD_USER)) & Server.htmlEncode(varEventData)
        End If

      Else

        ' entry/exit event
        varEventData = arrOneEvent(cEVENT_FIELD_USER)
        If arrOneEvent(cEVENT_FIELD_TYPE) = cEVENT_TYPE_ENTER Then
          varEventData = arrOneEvent(cEVENT_FIELD_DATA) & cCHAR_DELIMIT & varEventData
        End If

      End If

    End If

    If Len(varEventData) > 0 Then
      ' update user's event index
      Session.Contents("event_index") = arrOneEvent(cEVENT_FIELD_INDEX)
      ' return event
      Call addToArray(arrAllEvent, arrOneEvent(cEVENT_FIELD_TYPE) & cCHAR_DELIMIT & varEventData)
    End If

  Next

  getEvents = arrAllEvent

End Function
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' adds events for data posted by user
Sub addPost(strPost, strColour, strUserIDTo, blnExit)

  Dim blnEnter
  Dim arrUser

  blnEnter = CBool(Len(Session.Contents("user_name")) = 0)

  ' user entering - initialize
  If blnEnter Then

    ' ensure username doesn't exist
    If not blnExit Then
      For each arrUser in Application.Contents("user_list")
        blnExit = CBool(blnExit or arrUser(cUSER_FIELD_NAME) = strPost)
      Next
    End If

    If not blnExit Then

      ' initialize session vars
      With Session
        .Contents("user_name") = strPost
        .Contents("user_colour") = strColour
        .Contents("start_time") = Now()
        .Contents("event_index") = Application.Contents("event_index")
      End With
      With Application
        Call .lock()
        .Contents("user_index") = .Contents("user_index") + 1
        Session.Contents("user_index") = .Contents("user_index")
        Call .unlock()
      End With

      ' add entry event
      Call addEvent(cEVENT_TYPE_ENTER, Now(), strPost, Session.Contents("user_index"))

      ' set post to entry message
      strPost = cMSG_ENTER

    Else

      ' username already exists or exit attempted w/o username, cancel post
      strPost = ""

    End If

  End If

  If Len(strPost) > 0 Then

    ' handle hooks or standard post
    Select Case strPost
      Case "!browser"
        ' hook for user agent
        strPost = " has the browser: " & Request.ServerVariables("HTTP_USER_AGENT")
      Case "!address"
        ' hook for ip address
        strPost = " has the ip address: " & Request.ServerVariables("REMOTE_ADDR")
      Case "!time"
        ' hook for time logged in
        strPost = " has been logged in for " & Int(DateDiff("s", Session.Contents("start_time"), Now) / 60) & " minute" & String(CInt(CBool(Int(DateDiff("s", Session.Contents("start_time"), Now) / 60) <> 1)) * -1, "s") & " and " & DateDiff("s", Session.Contents("start_time"), Now) mod 60 & " second" & String(CInt(CBool(DateDiff("s", Session.Contents("start_time"), Now) mod 60 <> 1)) * -1, "s")
      Case Else
        If blnExit Then
          Call addEvent(cEVENT_TYPE_EXIT, Now, Session.Contents("user_name"), "")
        ElseIf Len(strUserIDTo) > 0 Then
          ' private message
          strPost = cMSG_POST_PRIVATE & strUserIDTo & ": " & strPost
        ElseIf not blnEnter Then
          ' standard post
          strPost = cMSG_POST_PUBLIC & ": " & strPost
        End If
    End Select

    ' add post and current posts to post array
    Call addEvent(cEVENT_TYPE_POST, Now(), Session.Contents("user_name"), strPost)

  End If

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' revert user post index back by post buffer on reload
Sub frameLoad()

  With Session
    If Len(.Contents("user_name")) > 0 Then
      .Contents("event_index") = .Contents("event_index") - cBUFFER_EVENT
    End If
  End With

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' returns array of user names
Function getUsers()

  Dim arrOneUser
  ReDim arrAllUser(-1)

  Call setUsers(Session.Contents("user_name"))

  For each arrOneUser in Application.Contents("user_list")
    Call addToArray(arrAllUser, arrOneUser(cUSER_FIELD_INDEX) & cCHAR_DELIMIT & arrOneUser(cUSER_FIELD_NAME))
  Next

  getUsers = arrAllUser

End Function
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' returns true if user exists
Function isUser(strUserName)

  Dim blnIsUser
  Dim arrUser

  blnIsUser = false

  Call setUsers("")

  For each arrUser in Application.Contents("user_list")
    blnIsUser = CBool(blnIsUser or arrUser(cUSER_FIELD_NAME) = strUserName)
  Next

  isUser = blnIsUser

End Function
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' main routine always called
Sub main(strOperation)

  ' initialize app level vars
  With Application
    If not IsArray(.Contents("event_list")) Then
      Call .lock()
      .Contents("user_index") = -1
      .Contents("user_list") = Array()
      .Contents("event_index") = -1
      .Contents("event_list") = Array()
      Call .unlock()
    End If
  End With

  ' handle requested operation
  With Response
    Select Case strOperation

      Case cOP_VAL_GET ' get events
        ' maintain the app user list
        Call setUsers(Session.Contents("user_name"))
        ' write out events for user
        Call .write(Join(getEvents(Session.Contents("event_index")), vbLf))
        Call .end()

      Case cOP_VAL_POST ' set post
        ' add events for posted data
        Call addPost(Request.QueryString(cOP_VAL_POST), Request.QueryString(cOP_VAL_COLOUR), Request.QueryString(cOP_VAL_USER), false)
        Call .end()

      Case cOP_VAL_FRAME ' frame
        Call frameLoad()
        Call .end()

      Case cOP_VAL_USER ' get users
        ' write out all user names when called from client-side initialize JS
        Call .write(Join(getUsers(), vbLf))
        Call .end()

      Case cOP_VAL_VALID ' validate availability of user name
        Call .write(LCase(not isUser(Request.QueryString(cOP_VAL_VALID))))
        Call .end()

      Case cOP_VAL_EXIT ' exit
        ' add exit message, clear session contents and redirect to self w/o query-string
        Call addPost(cMSG_EXIT, Session.Contents("user_colour"), "", true)
        Call Session.Contents.removeAll()
        Call .redirect(Request.ServerVariables("PATH_INFO"))

    End Select
  End With

End Sub
' ********************************************************************************************************************************************


' ********************************************************************************************************************************************
' returns HTML string for colour palette
Function getPalette()

  Dim strPalette, strR, strG, strB
  Dim intColour
  Dim arrHex, arrColour
  
  arrHex = Array("0", "3", "6", "9", "c", "f")
  
  ' build array of web-safe colours
  For each strR in arrHex
    For each strG in arrHex
      For each strB in arrHex
        Call addToArray(arrColour, "#" & String(1, strR) & String(1, strG) & String(1, strB))
      Next
    Next
  Next
  
  ' build palette HTML
  strPalette = "<table cellpadding=""0"" cellspacing=""0"" class=""" & cCSS_CLASS_PALETTE & """><tr>"
  For intColour = LBound(arrColour) to UBound(arrColour)
    If intColour > 0 and intColour mod 36 = 0 Then
      strPalette = strPalette & "</tr><tr>"
    End If
    strPalette = strPalette & "<td class=""" & cCSS_CLASS_PALETTE & """ style=""background-color:'" & arrColour(intColour) & "';"">&nbsp;</td>"
  Next
  strPalette = strPalette & "</tr></table>"
  
  getPalette = strPalette  

End Function
' ********************************************************************************************************************************************


%>