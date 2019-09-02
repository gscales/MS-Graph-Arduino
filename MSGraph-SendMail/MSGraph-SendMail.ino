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
String toAddress = "glenscales@yahoo.com";
String subject = "Test Message";
String body = "  Hello\r\n rgds\r\n Glen";

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
  Serial.println("Send Message");

  String PostData = BuildSendEmailRequest(toAddress, subject, body, "true");
  GraphPostRequest(AccessToken,"/v1.0/me/sendMail", PostData);
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

String BuildSendEmailRequest(String To, String Subject, String Body, String SaveToSentItems) {
  String content = "{";
  content += "        \"Message\": {";
  content += "           \"Subject\": \"" + Subject + "\",";
  content += "            \"Body\": {";
  content += "                \"ContentType\": \"Text\",";
  content += "                \"Content\": \"" + Body + "\"";
  content += "                       },";
  content += "            \"ToRecipients\": [";
  content += "                {";
  content += "                    \"EmailAddress\": {";
  content += "                        \"Address\": \"" + To + "\"";
  content += "                    }";
  content += "                }";
  content += "            ]";
  content += "        },";
  content += "        \"SaveToSentItems\": \"" + SaveToSentItems + "\"";
  content += "    }";
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
