<script type="text/javascript">
  <!--

  var oHTTP;
  var blnLoaded = false;

  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  initialization
  function init(strUserName) {

    var arrUserName;

    // ensure http request available
    try {
      // IE
      oHTTP = new ActiveXObject('Microsoft.XMLHTTP');
      blnLoaded = true;
    } catch (oError) {
      try {
        // Moz, Safari and whatever else supports this
        oHTTP = new XMLHttpRequest();
        blnLoaded = true;
      } catch (oError) {
        blnLoaded = false;
      }
    }

    // handle errors
    if (!blnLoaded) {
      alert('Your browser is not supported');
    } else {
      // validate user name
      blnLoaded = eval(http('<% = cOP_NAME %>=<% = cOP_VAL_VALID %>&<% = cOP_VAL_VALID %>=' + urlEncode(strUserName)));
      if (!blnLoaded) {
        alert('The name you entered is already in use, please choose another.');
      }
    }

    if (blnLoaded) {

      // display logged in features
      document.getElementById('hidden').style.display = 'inline';

      // set user list
      arrUserName = http('<% = cOP_NAME %>=<% = cOP_VAL_USER %>').split('\n');
      for (var i = 0; i < arrUserName.length; i++) {
        userAdd(arrUserName[i]);
      }

      // start post retrieval routine
      scroller();

    }

    return blnLoaded;

  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  recursive routine for post retrieval and auto scroll
  function scroller() {

    if (getter()) {
      if (document.forms[0].scroll.checked) {
        frames[0].scrollBy(0, 10000);
      }
      if (document.forms[0].activity.checked) {
        alert('New Activity!');
        document.forms[0].post.focus();
      }
    }
    window.setTimeout('scroller()', <% = cBUFFER_LOAD * 1000 %>);

  }

  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  sends and returns data
  function http(strQuery) {

    var strResponse = '';

    if (blnLoaded) {
      try {
        with (oHTTP) {
          open('GET', document.location.protocol + '//' + document.location.host + document.location.pathname + '?r=' + Math.random() + '&' + strQuery, false);
          send('');
          if (responseText) {
            strResponse = responseText;
          }
        }
      } catch (oError) {
        alert('Connection Error - Click OK to retry');
      }
    }

    return strResponse;

  }

  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  get posts
  function getter() {

    var blnReturn = false;
    var arrEvent = http('<% = cOP_NAME %>=<% = cOP_VAL_GET %>').split('\n');

    // loop thru events
    if (arrEvent.length > 0) {
      for (var i = 0; i < arrEvent.length; i++) {
        // determine event type
        with (arrEvent[i]) {
          switch (substr(0, indexOf('<% = cCHAR_DELIMIT %>'))) {

            case '<% = cEVENT_TYPE_POST %>':
              // add to post iframe
              frames[0].document.write('<p class="post;">' + substr(indexOf('<% = cCHAR_DELIMIT %>') + 1) + '</p>');
              blnReturn = true;
              break;

            case '<% = cEVENT_TYPE_ENTER %>':
              // add user to list
              userAdd(substr(indexOf('<% = cCHAR_DELIMIT %>') + 1));
              break;

            case '<% = cEVENT_TYPE_EXIT %>':
              // remove user from list
              userRemove(substr(indexOf('<% = cCHAR_DELIMIT %>') + 1));
              break;

          }
        }
      }
    }

    return blnReturn;

  }

  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  send post
  function poster(oForm) {

    var varToUserID = new Array();

    // get selected user IDs
    for (var i = 0; i < oForm.users.length; i++) {
      if (oForm.users[i].selected) {
        varToUserID.length++;
        varToUserID[varToUserID.length - 1] = oForm.users[i].value;
      }
    }
    varToUserID = varToUserID.join('<% = cCHAR_DELIMIT %>');
    if (varToUserID.length > 0) {
      varToUserID = '&<% = cOP_VAL_USER %>=' + varToUserID;
    }

    with (oForm.post) {
      if (value.length > 0 && value != defaultValue && (blnLoaded || init(value))) {
        http('<% = cOP_NAME %>=<% = cOP_VAL_POST %>&<% = cOP_VAL_POST %>=' + urlEncode(value) + varToUserID);
        value = '';
      }
      focus();
    }

    return false;

  }
  
  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  encodes data for use in query string
  function urlEncode(strValue) {
    return escape(strValue).replace(new RegExp('\\+', 'g'), '%2B');
  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  removes default value from post text box when focused on
  function postClick(oPost) {

    with (oPost) {
      if (value == defaultValue) {
        value = '';
      }
    }

  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  add user to the user list
  function userAdd(strUserDetail) {

    var blnAdd = true;
    var intUserIndex = 0;
    var intUserID = strUserDetail.substr(0, strUserDetail.indexOf('<% = cCHAR_DELIMIT %>'));
    var strUserName = strUserDetail.substr(strUserDetail.indexOf('<% = cCHAR_DELIMIT %>') + 1);

    // first entry upon application start will return an empty user in initial user population
    if (!isNaN(intUserID) && strUserName.length > 0) {

      with (document.forms[0]) {

        // ensure username doesn't exist - double handling could occur upon entry or reload
        for (var i = 0; i < users.length; i++) {
          blnAdd = (blnAdd && users[i].text != strUserName);
        }

        if (blnAdd) {

          users.length++;
          if (users.length > 1 && strUserName.toLowerCase() > users[users.length - 2].text.toLowerCase()) {
            // user is bottom in list - set new user index to bottom
            intUserIndex = users.length - 1;
          } else {
            // from bottom to top of list - move users down list until new user position found, then set new user index to current position
            for (var i = users.length - 2; i >= 0; i--) {
              if (users[i].text.toLowerCase() > strUserName.toLowerCase()) {
                users[i + 1].value = users[i].value;
                users[i + 1].text = users[i].text;
                users[i + 1].selected = users[i].selected;
              } else if (intUserIndex == 0) {
                intUserIndex = i + 1;
              }
            }
          }
          // add new user to list
          users[intUserIndex].value = intUserID;
          users[intUserIndex].text = strUserName;
          users[intUserIndex].selected = false;

        }

      }
    }

  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  remove user from the user list
  function userRemove(strUserName) {

    var blnFound = false;

    with (document.forms[0]) {

      // from top to bottom of list, once user to remove is found, move users up list then finally when at end, remove last user in list
      for (var i = 0; i < users.length; i++) {
        blnFound = (blnFound || users[i].text == strUserName);
        if (blnFound) {
          if (i < users.length - 1) {
            users[i].value = users[i + 1].value;
            users[i].text = users[i + 1].text;
            users[i].selected = users[i + 1].selected;
          } else {
            users.length--;
          }
        }
      }

    }

  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  //  redirect when leave button clicked
  function userExit() {

    with (document) {
      location = location.href.split('?')[0] + '?<% = cOP_NAME %>=<% = cOP_VAL_EXIT %>';
    }

  }

  //-->
</script>
