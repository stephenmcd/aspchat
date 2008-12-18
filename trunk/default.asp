<% Option Explicit : On Error Resume Next %>
<!-- #include file="constants_vbs.asp" -->
<!-- #include file="functions_vbs.asp" -->
<% Call main(Request.QueryString(cOP_NAME)) %>
<html>
	<head>
		<title>Chat</title>
		<!-- #include file="functions_js.asp" -->
		<style type="text/css">
			<!--

			body {overflow:hidden;}
			div {white-space:nowrap;}
			iframe {width:80%;height:94%;float:left;}
			input {vertical-align:top;}
			label {cursor:pointer;vertical-align:top;}
			#users{width:15%;height:95%;}
			#post {width:45%;vertical-align:top;}
			#hidden {display:none;}
			table.palette {border:0px;display:inline;margin-right:5px;}
			td.palette {font-size:1px;width:2px;height:4px;cursor:pointer;}

			//-->
		</style>
	</head>
	<body <% If Len(Session.Contents("user_name")) > 0 Then %> onload="init('');"<% End If %>>
		<form onsubmit="return poster(this);">

			<div style="white-space:nowrap;" />
				<iframe src="<% = Request.ServerVariables("PATH_INFO") %>?<% = cOP_NAME %>=<% = cOP_VAL_FRAME %>"></iframe>
				<select name="users" id="users" multiple="multiple" ondblclick="this.selectedIndex = -1;" onclick="this.form.post.focus();"></select>
			</div>

			<div style="white-space:nowrap;" />

				<% = getPalette() %>
				<input type="hidden" name="colour" value="<% If Len(Session.Contents("user_colour")) > 0 Then %><% = Session.Contents("user_colour") %><% Else %><% = cDEFAULT_COLOUR %><% End If %>">
				<input type="text" name="post" id="post" value="<% If Len(Session.Contents("user_name")) = 0 Then %>Enter your name...<% End If %>" onclick="postClick(this);" />
				<input type="submit" value="Send" />

				<div id="hidden">
					<input type="button" value="Leave" onclick="userExit();" />
					<input type="checkbox" name="scroll" id="scroll" checked="checked" onclick="this.form.post.focus();" />
					<label for="scroll">Scroll</label>
					<input type="checkbox" name="activity" id="activity" onclick="return false;this.form.post.focus();" />
					<label for="activity">Activity Alert</label>
				</div>

			</div>

		</form>
	</body>
</html>