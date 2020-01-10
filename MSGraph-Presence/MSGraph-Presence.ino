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
    // wait 5 seconds for connection:
  delay(5000);
  //String AccessToken = OAuthRequest(resource, true);
  //Serial.println("Get users Profile");
  //GetRequest(AccessToken);
  String TokenEndPoint = GetTenantDeviceoAuthURL();
  int tknPoint = TokenEndPoint.indexOf("token");
  String DeviceCodeEndPoint = TokenEndPoint.substring(0,tknPoint) + "v2.0/devicecode";
  DeviceCodeEndPoint.replace("https://login.microsoftonline.com","");
  Serial.println(DeviceCodeEndPoint);
  String DeviceCode = OAuthRequest(DeviceCodeEndPoint,true);

  
  int MaxLoop = 0;
  String Token = "";
   while((MaxLoop < 40) && (Token == "")){
    MaxLoop = MaxLoop++;
    Token = TokenPollRequest(DeviceCode,TokenEndPoint);
  }
  GetPresenceRequest(Token);
  Serial.println("Done");
  
}

String GetTenantDeviceoAuthURL(){
  String Domain = SECRET_OFFICE365_DOMAIN;
  String URL = "/" + Domain + "/.well-known/openid-configuration";
  httpClient.setTimeout(5000);
  httpClient.beginRequest();
  httpClient.get(URL);
  httpClient.endRequest();
  int statusCode = httpClient.responseStatusCode();
  String response = httpClient.responseBody();

  Serial.print("Status code: ");
  Serial.println(statusCode);
  Serial.print("Response: ");
  Serial.println(response);
  DynamicJsonDocument jsonDoc(MAX_CONTENT_SIZE);
  auto error = deserializeJson(jsonDoc, response);
  if (error) {
    Serial.print(F("deserializeJson() failed with code "));
    Serial.println(error.c_str());
    return error.c_str();
  } else {
    String TokenEndPoint = jsonDoc["token_endpoint"];
    Serial.println(TokenEndPoint);
    return TokenEndPoint;
  }
}

String TokenPollRequest(String DeviceCode,String Url){
  Serial.print("TokenPollRequest");
  String ClientID = SECRET_OFFICE365_CLIENTID;
  String postData = "resource=" + URLEncode("https://graph.microsoft.com") + "&grant_type=urn:ietf:params:oauth:grant-type:device_code&client_id=" + ClientID + "&code=" + DeviceCode;
  Serial.print(postData);
  httpClient.beginRequest();
  httpClient.setTimeout(5000);
  httpClient.post(Url);
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
      String TokenResponse = httpClient.responseBody();
      Serial.println(TokenResponse);
      int TokenStart = TokenResponse.indexOf("access_token");
      int TokenEnd = TokenResponse.indexOf("refresh_token");
      String AccessToken = TokenResponse.substring((TokenStart + 15),(TokenEnd-3));  
      return AccessToken;
  } else {
    Serial.println(httpClient.responseBody().substring(0,200));
    delay(7000);
    return "";
     
  }

}


String OAuthRequest(String Url,bool retry) {
  Serial.println("Make OAuth Login Request");
  String ClientID = SECRET_OFFICE365_CLIENTID;
  String postData = "client_id=" + ClientID + "&scope=user.read%20openid%20profile";
  httpClient.beginRequest();
  httpClient.setTimeout(5000);
  httpClient.post(Url);
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
  String BodyResponse = httpClient.responseBody();
  Serial.println(BodyResponse);
  DynamicJsonDocument jsonDoc(MAX_CONTENT_SIZE);
  auto error = deserializeJson(jsonDoc,BodyResponse);
  if (error) {
    Serial.print(F("deserializeJson() failed with code "));
    Serial.println(error.c_str());
    return error.c_str();
  } else {
    String DeviceCode = jsonDoc["device_code"];
    String UserCode = jsonDoc["user_code"];
    Serial.println("Your User Code is " + UserCode);    
    return DeviceCode;
  }
  
}



String GetPresenceRequest(String AccessToken) {
  HttpClient httpClient(wifiClient, "graph.microsoft.com", 443);
  httpClient.beginRequest();
  httpClient.get("/beta/me/presence");
  httpClient.sendHeader("Authorization", ("Bearer " + AccessToken));
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
