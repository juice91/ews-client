var ews = require('ews-javascript-api');
window.ews=ews;


/*
TEST request
var request = require("request");

var options = { method: 'GET',
  url: 'https://postman-echo.com/headers',
  headers:
    { 'postman-token': '3a1dd073-ad46-bfc1-c87b-1e0a29df0317',
      'cache-control': 'no-cache',
      'my-sample-header': 'Lorem ipsum dolor sit amet' } };

request(options, function (error, response, body) {
  if (error) throw new Error(error);

  console.log(body);
});
*/


/*

Auto Discover
var ews = require('ews-javascript-api');
//create ExchangeService object
var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2013);
exch.Credentials = new ews.ExchangeCredentials("userName", "password");
//set ews endpoint url to use
exch.Url = new ews.Uri("https://outlook.office365.com/Ews/Exchange.asmx"); // you can also use exch.AutodiscoverUrl

var attendee =[ new ews.AttendeeInfo("email1@domain.com"), new ews.AttendeeInfo("email2@domain.com")];
//create timewindow object o request avaiability suggestions for next 48 hours, DateTime and TimeSpan object is created to mimic portion of .net datetime/timespan object using momentjs
var timeWindow = new ews.TimeWindow(ews.DateTime.Now, ews.DateTime.Now.AddDays(2));
exch.GetUserAvailability(attendee, timeWindow, ews.AvailabilityData.FreeBusyAndSuggestions)
  .then(function (availabilityResponse) {
    //do what you want with user availability
  }, function (errors) {
    //log errors or do something with errors
  });
        */



/*
User setting
//import ews module
var ews = require('ews-javascript-api');
//create AutodiscoverService object
var autod = new ews.AutodiscoverService(new ews.Uri("https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"), ews.ExchangeVersion.Exchange2010);
//you can omit url and it will autodiscover the url, version helps throw error on client side for unsupported operations.example - //var autod = new ews.AutodiscoverService(ews.ExchangeVersion.Exchange2010);
//set credential for service
autod.Credentials = new ews.ExchangeCredentials("userName", "password");
//create array to include list of desired settings
var settings = [
  ews.UserSettingName.InternalEwsUrl,
  ews.UserSettingName.ExternalEwsUrl,
  ews.UserSettingName.UserDisplayName,
  ews.UserSettingName.UserDN,
  ews.UserSettingName.EwsPartnerUrl,
  ews.UserSettingName.DocumentSharingLocations,
  ews.UserSettingName.MailboxDN,
  ews.UserSettingName.ActiveDirectoryServer,
  ews.UserSettingName.CasVersion,
  ews.UserSettingName.ExternalWebClientUrls,
  ews.UserSettingName.ExternalImap4Connections,
  ews.UserSettingName.AlternateMailboxes
];
//get the setting
autod.GetUserSettings(["email1@domain.com", "email2@domain.com"], settings)
  .then(function (response) {
    //do what you want with user settings
    var tabcount = 0;
    var tabs = function () { return ews.StringHelper.Repeat("\t", tabcount); };
    console.log(autod.Url.ToString());
    //uncoment next line to see full response from autodiscover, you will need to add var util = require('util');
    //console.log(util.inspect(response, { showHidden: false, depth: null, colors: true }));
    for (var _i = 0, _a = response.Responses; _i < _a.length; _i++) {
      var resp = _a[_i];
      console.log(ews.StringHelper.Format("{0}settings for email: {1}", tabs(), resp.SmtpAddress));
      tabcount++;
      for (var setting in resp.Settings) {
        console.log(ews.StringHelper.Format("{0}{1} = {2}", tabs(), ews.UserSettingName[setting], resp.Settings[setting]));
      }
      tabcount--;
    }
  }, function (e) {
    //log errors or do something with errors
  });*/
