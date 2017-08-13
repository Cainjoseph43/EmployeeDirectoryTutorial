<div class="ms-Grid-row">
<div class="ms-sm12">
<input class="ms-SearchBoxSmallExample" type="text" id="search_box" placeholder="Search Here..." style="margin:50px 25px 25px 25px" />
</div></div>
<div id="employeeDirectory" class="ms-Grid-row">
</div>

<script type="text/javascript">
function searchText(filter)
    {
        var searchText;

        if (filter !== undefined && filter != null)
        {
            searchText = filter.toLowerCase(); // Alpha filter
        }
        else
        {
            searchText = $(".search input[type=text]").val().toLowerCase();
        }

        $("span.listSearch").each(function ()
        {
            var obj = $(this);

            if ($.trim(searchText) == "")
            {
                obj.show();
                return true;
            }

            var entityText = $(this).text().toLowerCase();

            if (searchText.length == 1)
            {
                if (entityText.charAt(0) == searchText)
                {
                    obj.show();
                }
                else
                {
                    obj.hide();
                }
            }
            else if (entityText.search(searchText) > -1)
            {
                obj.show();
            }
            else {
                obj.hide();
            }
        });
    }


(function () {
	"use strict";
  // Some samples will use the tenant name here like "tenant.onmicrosoft.com"
  // I prefer to user the subscription Id 
  var subscriptionId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX";
  // Copy the client ID of your AAD app here once you have registered one, configured the required permissions, and
  // allowed implicit flow https://msdn.microsoft.com/en-us/office/office365/howto/get-started-with-office-365-unified-api
  var clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
	
  window.config = {
    //tenant: 'https://udigtech.onmicrosoft.com',
    subscriptionId: subscriptionId,                 
    clientId: clientId,     
    postLogoutRedirectUri: window.location.origin,
    endpoints: {
      graphApiUri: 'https://graph.microsoft.com'
    },
    cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
  };
  var authContext = new AuthenticationContext(config);	
  // Check For & Handle Redirect From AAD After Login
  var isCallback = authContext.isCallback(window.location.hash);
  authContext.handleWindowCallback();	
  if (isCallback && !authContext.getLoginError()) {
    window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
  }	
  // If not logged in force login
  var user = authContext.getCachedUser();
  if (user) {
    // Logged in already
  } 
  else {
    // NOTE: you may want to render the page for anonymous users and render
    // a login button which runs the login function upon click.
    authContext.login();
  }	
  // Acquire token for Files resource.
  authContext.acquireToken(config.endpoints.graphApiUri, function (error, token) {
    // Handle ADAL Errors.
    if (error || !token) {
      console.log('ADAL error occurred: ' + error);
      return;
    }
    // Execute GET request to Files API.
    //var currentUserApiBaseUri = graphApiUri + "/beta/" + config.subscriptionId + "/users/" + user.userName;
    //var filesUri = currentUserApiBaseUri + "/files";
    var filesUri = config.endpoints.graphApiUri + "/v1.0/users/";
    $.ajax({
      type: "GET",
      url: filesUri,
      headers: {
        'Authorization':'Bearer ' + token,
      },
      success: function(data){
        var object = data.value;
          for (var i = 0; i < object.length; i++) {

            //Change '-' to '.' for both mobile numbers
            if(object[i].mobilePhone == null){
                var m2 = object[i].mobilePhone;
            }
            else{
            var m = object[i].mobilePhone;
            var m1 = m.toString();
            var m2 = m.replace(/-/g,'.');
            }

            //Change '-' to '.' for business numbers
            if(object[i].businessPhones == ''){
                var o2 = object[i].businessPhones;
            }
            else{
            var o = object[i].businessPhones;
            var o1 = o.toString();
            var o2 = o1.replace(/-/g,'.');
            }

            var name = object[i].displayName;
            var lowerName = name.toLowerCase();
            if(object[i].mail == null || object[i].jobTitle == null){
                //Do nothing
            }
                else {
            var employee = `
              <span class="listSearch"><div class="ms-Grid-col ms-sm12 ms-md6 ms-lg6 ms-xl6 ms-xxl4 empCard">
              <div class="cardContent">
                <p class="empName" style="line-height:1.2;" id="`+ lowerName +`"><span style="display:none">` + lowerName + `</span> `+ object[i].displayName +`<br/>
                <span class="jobTitle">`+ object[i].jobTitle +`</span></p>
                <p class="empLocation">`+ object[i].officeLocation +`</p>
                <p style="line-height:100%">
                    O: <span class="empPhone"><a href="tel:`+ o2 +`">`+o2+`</a></span><br/>
                    M: <span class="empPhone"><a href="tel:`+ m2 +`">`+ m2 +`</a></span><br/>
                    E: <span class="empPhone"><a href="mailto:`+ object[i].mail +`">`+ object[i].mail +`</a></span>
              </p></div></div></span>
            `;
            $('#employeeDirectory').append(employee); 
          }      
          }
      },
    }).fail(function () {
      console.log('Fetching files from AD Failed.');
    });
  });
  //Search for the typed keyword each time a key is pressed
       $("#search_box").keyup(function() {
            var valThis = $(this).val();
            
            var lowerValThis = valThis.toLowerCase();
            console.log(lowerValThis);
            if (lowerValThis.length > 1) {
                searchText(lowerValThis);
                //$.when($('span.listSearch').find("div:not(:Contains(" + lowerValThis + "))").parent().slideUp()).done();
                //$('span:has(' + valThis + ')').hide();
            } else {
                $.when($('span.listSearch').find("div:not(:Contains(" + lowerValThis + ")").parent().show()).done();
            } //change function
        }); //Search Box

})();

</script>
