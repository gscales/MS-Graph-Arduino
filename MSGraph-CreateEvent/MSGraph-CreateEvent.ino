#include <SPI.h>
#include <WiFiNINA.h>
#include <ArduinoHttpClient.h>
#include <ArduinoJson.h>
#include "arduino_secrets.h"


const size_t MAX_CONTENT_SIZE = 16000;
char ssid[] = SECRET_SSID;
char pass[] = SECRET_PASS;
char resource[] = "https://graph.microsoft.com";
int status = WL_IDLE_STATUS;
int keyIndex = 0;
const char loginHostname[] = "login.microsoftonline.com";

String startTime = "2019-08-15T13:00:00Z";
String endTime = "2019-08-15T14:00:00Z";
String timeZone = "AUS Eastern Standard Time";
String subject = "Test Event Happened";
String body = "Body";
String reminder = "false";

WiFiSSLClient wifiClient;

HttpClient httpClient(wifiClient, loginHostname, 443);

void setup() {
  //Initialize serial and wait for port to open:
  Serial.begin(9600);
  while (!Serial) {
    ; // wait for serial port to connect. Needed for native USB port only
  }

  // check for the WiFi module:
  if (WiFi.status() == WL_NO_MODULE) {
    Serial.println("Communication with WiFi module failed!");
    // don't continue
    while (true);
  }

  String fv = WiFi.firmwareVersion();
  if (fv < "1.0.0") {
    Serial.println("Please upgrade the firmware");
  }

  // attempt to connect to Wifi network:
  while (status != WL_CONNECTED) {
    Serial.print("Attempting to connect to WEP network, SSID: ");
    Serial.println(ssid);
    status = WiFi.begin(ssid, keyIndex, pass);

    // wait 10 seconds for connection:
    delay(10000);
  }

  // once you are connected :
  Serial.println("You're connected to the network");
  String AccessToken = OAuthRequest(resource, true);
  Serial.println("Create Event");

  String PostData = BuildCreateEvent(startTime,endTime,timeZone,subject,body,"false");
  GraphPostRequest(AccessToken,"/v1.0/me/events", PostData);
  
}

String OAuthRequest(char* resource, bool retry) {
  Serial.println("Make OAuth Login Request");

  String postData = BuildOAuthRequest(resource);
  httpClient.beginRequest();
  httpClient.setTimeout(20000);
  httpClient.post("/Common/oauth2/token");
  httpClient.sendHeader("Content-Type", "application/x-www-form-urlencoded");
  httpClient.sendHeader("Content-Length", postData.length());
  Serial.println("Headers Sent");
  httpClient.beginBody();
  Serial.println("Body Sent");
  httpClient.print(postData);
  Serial.println("End Request");
  httpClient.endRequest();


  // read the status code and body of the response
  int statusCode = httpClient.responseStatusCode();
  postData = "";
  Serial.print("Status code: ");
  Serial.println(statusCode);
  if (statusCode == 200) {
    return readAuthReponse();
  } else {
    if (retry) {
      OAuthRequest(resource, false);
    }
  }
}

String GetCurrentTime(String URL){
  HttpClient httpClient(wifiClient, "worldtimeapi.org", 80);
  httpClient.beginRequest();
  httpClient.get(URL);
  httpClient.sendHeader("Content-Type", "application/json");
  httpClient.endRequest();
  // read the status code and body of the response
  int statusCode = httpClient.responseStatusCode();
  String response = httpClient.responseBody();

  Serial.print("Status code: ");
  Serial.println(statusCode);
  Serial.print("Response: ");
  Serial.println(response);
}

String BuildOAuthRequest(char* resource) {
  String UserName = SECRET_OFFICE365_USERNAME;
  String Password = SECRET_OFFICE365_PASSWORD;
  String ClientId = SECRET_OFFICE365_CLIENTID;

  char passwordCA[Password.length() + 1];
  Password.toCharArray(passwordCA, Password.length() + 1);
  return ("resource=" + URLEncode(resource) + "&client_id=" + ClientId + "&grant_type=password&username="
          + UserName + "&password=" + URLEncode(passwordCA) + "&scope=openid");

}

String readAuthReponse() {
  DynamicJsonDocument jsonDoc(MAX_CONTENT_SIZE);
  auto error = deserializeJson(jsonDoc, httpClient.responseBody());
  if (error) {
    Serial.print(F("deserializeJson() failed with code "));
    Serial.println(error.c_str());
    return error.c_str();
  } else {
    String token = jsonDoc["access_token"];
    return token;
  }

}

String BuildCreateEvent(String StartTime, String EndTime,String TimeZone,String Subject,String Body,String Reminder ) {
  String content = "{\r\n";
  content += "\"subject\": \"" + Subject + "\",\r\n";
  content += "     \"body\": {\r\n";
  content += "        \"ContentType\": \"Text\",\r\n";
  content += "         \"Content\": \"" + Body + "\"\r\n";
  content += "    },\r\n";
  content += "    \"isReminderOn\": " + Reminder + ",\r\n";  
  content += "    \"start\": {\r\n";
  content += "    \"dateTime\": \"" + StartTime + "\",\r\n";
  content += "    \"timeZone\": \"" + TimeZone + "\"\r\n";
  content += "    },\r\n";
  content += "    \"end\": {\r\n";
  content += "    \"dateTime\": \"" + EndTime + "\",\r\n";
  content += "    \"timeZone\": \"" + TimeZone + "\"\r\n";
  content += "    }\r\n";
  content += "}";
  return content;
}

String GraphPostRequest(String AccessToken,String GraphEndPoint, String postData) {
  HttpClient httpClient(wifiClient, "graph.microsoft.com", 443);
  httpClient.beginRequest();
  httpClient.post(GraphEndPoint);
  httpClient.sendHeader("Authorization", ("Bearer " + AccessToken));
  httpClient.sendHeader("Content-Type", "application/json");
  httpClient.sendHeader("Content-Length", postData.length());
  Serial.println("Headers Sent");
  httpClient.beginBody();
  Serial.println("Body Sent");
  httpClient.print(postData);
  Serial.println("End Request");
  httpClient.endRequest();
  // read the status code and body of the response
  int statusCode = httpClient.responseStatusCode();
  String response = httpClient.responseBody();

  Serial.print("Status code: ");
  Serial.println(statusCode);
  Serial.print("Response: ");
  Serial.println(response);

}

void loop() {

}



String URLEncode(const char* msg)
{
  const char *hex = "0123456789abcdef";
  String encodedMsg = "";

  while (*msg != '\0') {
    if ( ('a' <= *msg && *msg <= 'z')
         || ('A' <= *msg && *msg <= 'Z')
         || ('0' <= *msg && *msg <= '9') ) {
      encodedMsg += *msg;
    } else {
      encodedMsg += '%';
      encodedMsg += hex[*msg >> 4];
      encodedMsg += hex[*msg & 15];
    }
    msg++;
  }
  return encodedMsg;
}
