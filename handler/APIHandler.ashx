<%@ WebHandler Language = "C#" Class="Handler" %>

using System;
using System.IO;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Serilog;
using System.Web.Script.Serialization;
using System.Web;
using MarvalSoftware.UI.WebUI.ServiceDesk.RFP.Plugins;
using System.Linq;
using System.Xml.Linq;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public class Handler : PluginHandler
{
    public class RequestData
    {
        public string action { get; set; }

        public string apptoken { get; set; }
        public string siteId { get; set; }
        public string folderId { get; set; }
        public string folderPath { get; set; }
        public string microsoftToken { get; set; }
        public string combinedText { get; set; }
        public string attachmentName { get; set; }

    }

    public class fullResponse
    {
        public int responseCode { get; set; }//res code
        public string responseDes { get; set; } //res desc
        public string responseBody { get; set; } //res body
    }

    public class SharePointUploadResponse
    {
        public HttpStatusCode StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string Content { get; set; }
        public bool IsSuccess { get; set; }
        public string Error { get; set; }
        public string ContentType { get; set; }
        public WebHeaderCollection Headers { get; set; }
    }


    private string APIKey { get; set; }


    private string Password { get; set; }
    private string Username { get; set; }
    private string Host { get; set; }

    private string DBName { get; set; }
    private string MarvalHost { get; set; }
    private string AssignmentGroups { get; set; }

    private string ClientID { get { return this.GlobalSettings["@@ClientID"]; } }
    private string MarvalAPIKey { get { return this.GlobalSettings["@@MarvalAPIKey"]; } }
    private string CustomAttribute { get { return this.GlobalSettings["@@CustomAttribute"]; } }
    private string AutomateSubfolderCreation { get { return this.GlobalSettings["@@AutomateSubfolderCreation"]; } }
    private string SuggestDirectoryName { get { return this.GlobalSettings["@@SuggestDirectoryName"]; } }
    private string createFolderOption { get { return this.GlobalSettings["@@createFolderOption"]; } }

    private string MSMBaseUrl
    {
        get
        {
            return "https://" + HttpContext.Current.Request.Url.Host + MarvalSoftware.UI.WebUI.ServiceDesk.WebHelper.ApplicationPath;
        }
    }



    private int MsmRequestNo { get; set; }

    private int lastLocation { get; set; }

    public override bool IsReusable { get { return false; } }

    private HttpWebRequest BuildRequest(string uri = null, string body = null, string method = "GET")
    {
        //https://stackoverflow.com/a/2904963
        ServicePointManager.Expect100Continue = true;
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
        var request = WebRequest.Create(new UriBuilder(uri).Uri) as HttpWebRequest;
        request.Method = method.ToUpperInvariant();
        request.ContentType = "application/json";

        if (body != null)
        {
            using (var writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(body);
            }
        }

        return request;
    }
    //     private HttpWebRequest BuildRequest(string uri = null, string body = null, string method = "GET", string token = null)
    // {
    //     ServicePointManager.Expect100Continue = true;
    //     ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

    //     var request = WebRequest.Create(new UriBuilder(uri).Uri) as HttpWebRequest;
    //     request.Method = method.ToUpperInvariant();
    //     request.ContentType = "application/json";

    //     if (!string.IsNullOrWhiteSpace(token))
    //     {
    //         request.Headers["Authorization"] = "Bearer " + token;
    //     }

    //     if (body != null)
    //     {
    //         using (var writer = new StreamWriter(request.GetRequestStream()))
    //         {
    //             writer.Write(body);
    //         }
    //     }

    //     return request;
    // }
    private string GetRequest(string url, string token)
    {
        try
        {
            // Create a web request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "application/json";
            request.Headers["Authorization"] = "Bearer " + token;

            // // Write data to request body
            // using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
            // {
            //     writer.Write(data);
            // }

            // Get response
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    return reader.ReadToEnd();
                }
            }
        }
        catch (WebException webEx)
        {
            // If we have a response, we can read the error message from the response body
            if (webEx.Response != null)
            {
                using (var errorResponse = (HttpWebResponse)webEx.Response)
                {
                    using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        string errorText = reader.ReadToEnd();
                        // Return or log the error text
                        return errorText;
                    }
                }
            }

            // If we have no response, return the exception message
            return webEx.Message;
        }
        catch (Exception ex)
        {
            // Handle other exceptions
            return ex.ToString();
        }
    }


    private string PostRequest(string url, string data)
    {
        try
        {
            // Create a web request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";

            // Write data to request body
            using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(data);
            }

            // Get response
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    return reader.ReadToEnd();
                }
            }
        }
        catch (WebException webEx)
        {
            // If we have a response, we can read the error message from the response body
            if (webEx.Response != null)
            {
                using (var errorResponse = (HttpWebResponse)webEx.Response)
                {
                    using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        string errorText = reader.ReadToEnd();
                        // Return or log the error text
                        return errorText;
                    }
                }
            }

            // If we have no response, return the exception message
            return webEx.Message;
        }
        catch (Exception ex)
        {
            // Handle other exceptions
            return ex.ToString();
        }
    }

    private void AddMsmNote(int requestNumber, string note)
    {
        Log.Information("Adding note with ID " + requestNumber);
        IDictionary<string, object> body = new Dictionary<string, object>();
        body.Add("id", requestNumber);
        body.Add("content", note);
        body.Add("type", "public");
        string jsonNote = JsonHelper.ToJson(body);
        Log.Information("Have json note as " + jsonNote);//change msm base url
        var httpWebRequest = BuildRequest(this.MSMBaseUrl + string.Format("/api/serviceDesk/operational/requests/{0}/notes/", requestNumber), JsonHelper.ToJson(body), "POST");
        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
        ProcessRequest2(httpWebRequest);//build req then process it
    }

    private string ProcessRequest(HttpWebRequest request)
    {
        fullResponse myRes = new fullResponse();
        try
        {
            //var resStatus = ((HttpWebResponse)request.WebResponse).StatusCode;
            //request.Headers.Add("Authorization", "Bearer " + this.UserAPIKey);
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            var res = "";
            var resStatus = ((HttpStatusCode)response.StatusCode);
            using (StreamReader reader = new StreamReader(response.GetResponseStream()))
            {
                res = reader.ReadToEnd();
            }
            HttpContext.Current.Response.StatusCode = (int)resStatus;
            HttpContext.Current.Response.ContentType = "application/json";
            HttpContext.Current.Response.Write(res);
            HttpContext.Current.Response.End();
            return null;

        }
        catch (WebException webEx)
        {
            var result = "";
            var errStatus = ((HttpWebResponse)webEx.Response).StatusCode;
            var errResp = webEx.Response;

            //myRes.responseCode = Int32.Parse(errStatus.ToString());
            Log.Information("err is" + errStatus.ToString());
            myRes.responseDes = ((HttpWebResponse)errResp).StatusDescription;
            var res = "";
            using (StreamReader reader = new StreamReader(errResp.GetResponseStream()))
            {
                res = reader.ReadToEnd();
            }
            //myRes.responseBody = res;
            HttpContext.Current.Response.StatusCode = (int)errStatus;
            HttpContext.Current.Response.ContentType = "application/json";
            HttpContext.Current.Response.Write(res);
            HttpContext.Current.Response.End();

            return null;

        }
    }

    public override void HandleRequest(HttpContext context)
    {
        var param = context.Request.HttpMethod;
        var browserObject = context.Request.Browser;
        HttpWebRequest httpWebRequest;
        HttpWebRequest request;
        string microsoftToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkZ2dnRSbDNQU25Jck54REVHTHdVSExIVmsybm1BWXZUb0ZMVk9RTFNWaEUiLCJhbGciOiJSUzI1NiIsIng1dCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyIsImtpZCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC85YmNhM2NkZS1iM2Y3LTRjOWQtYWU4NC1lYmNkZWUzODI5OTAvIiwiaWF0IjoxNzUyMTA2MzkyLCJuYmYiOjE3NTIxMDYzOTIsImV4cCI6MTc1MjExMDQ2OSwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVdRQW0vOFpBQUFBclRXbmhER2llSVZCK1B5SzRzeGp0TlBkSzBGSS95QkI5Y2dFL2xITEJxdzlkWmdldjdBQ3I2RGU2SXJKU0YzQmovU3NNamFqb3ZmNEpoa1E5a1huNmVINWJ2VnhqeDNmODZpRTUyYVdSRGVldVM0OVAyWFVEeFhLZUlGanErQ0IiLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1hcnZhbCBBdXN0cmFsaWEgT2ZmaWNlIDM2NSBBZG1pbiIsImFwcGlkIjoiYjZhODQ5NzYtZDNiYy00ZjdkLWFlNDItYTUwNDcyYTY0MTE2IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJOZ3V5ZW4iLCJnaXZlbl9uYW1lIjoiRHlsYW4iLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiI0LjE5Ni43Ni4yNTMiLCJuYW1lIjoiRHlsYW4gTmd1eWVuIiwib2lkIjoiZWRmYzJhOTYtMzFkNS00NGFkLTg3NGQtYjUzZGJkZTZjMmMyIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDA0QTU1NjI2RDIiLCJyaCI6IjEuQVI4QTNqekttX2V6blV5dWhPdk43amdwa0FNQUFBQUFBQUFBd0FBQUFBQUFBQUJNQWZjZkFBLiIsInNjcCI6IkRpcmVjdG9yeS5SZWFkLkFsbCBEaXJlY3RvcnkuUmVhZFdyaXRlLkFsbCBHcm91cE1lbWJlci5SZWFkLkFsbCBHcm91cE1lbWJlci5SZWFkV3JpdGUuQWxsIE1haWwuU2VuZCBNYWlsLlNlbmQuU2hhcmVkIFNpdGVzLlJlYWQuQWxsIFNpdGVzLlJlYWRXcml0ZS5BbGwgU2l0ZXMuU2VsZWN0ZWQgVGVhbU1lbWJlci5SZWFkLkFsbCBUZWFtTWVtYmVyLlJlYWRXcml0ZS5BbGwgVGVhbU1lbWJlci5SZWFkV3JpdGVOb25Pd25lclJvbGUuQWxsIFVzZXIuUmVhZCBVc2VyLlJlYWRXcml0ZS5BbGwgVXNlckF1dGhlbnRpY2F0aW9uTWV0aG9kLlJlYWRXcml0ZS5BbGwgcHJvZmlsZSBvcGVuaWQgZW1haWwiLCJzaWQiOiIwMDRmYWU4OS00NmEwLTVlOWUtYjRmYy1iNDQxZjYxZGRkYjYiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiIzTnRKMzZ4Y01HMWZ2cDFjeFRhblJjb3ZWU1YyU2ZJRjlHbE5ra3VHY1ZNIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkVVIiwidGlkIjoiOWJjYTNjZGUtYjNmNy00YzlkLWFlODQtZWJjZGVlMzgyOTkwIiwidW5pcXVlX25hbWUiOiJkeWxhbi5uZ3V5ZW5AbWFydmFsLmNvbS5hdSIsInVwbiI6ImR5bGFuLm5ndXllbkBtYXJ2YWwuY29tLmF1IiwidXRpIjoiUXFHbmFudFBuVXlEVTFQdkFHSWpBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19mdGQiOiJOQkx4WnNHWTFkckZWX3NIeVJWSFNCZ3RXb0U4WDg4TjFjNVBvX0J5ekE0QmMzZGxaR1Z1WXkxa2MyMXoiLCJ4bXNfaWRyZWwiOiIxIDE0IiwieG1zX3N0Ijp7InN1YiI6ImZDMDlXY1VoY2R1MUZyWlAwTVFDaFo1WFA0ZWZBZFFWV1pncURDd0cyV1kifSwieG1zX3RjZHQiOjE1MDY1ODY4MjJ9.JuiUZ4PGO4HfwSH-vuYHpefPDAZGHoZvsaN1JZQnnH5BPRInWCKdyuV7ltD0LEILo3Iq0a53cmLEvSB54vfgH3a6KavEeqGAbJExTWenle9OwS1K9vIZfGCgV6KSd2-yvjSEOxe9kjbsRc6zI275sg5kW6nItVB5IJRZTHZALDl4_QCX_XuIaI0Z_I2E4-0ji-yaZNlMPpOxGlEbuzdnkmgGM6bn4l4VX6woQmlreEsAQnZ2T55ND3iF7OqjY70eeVGD2unjJsYukVrC5yJ8ws9LqXaKTVvjF0mC_tirNRssXBbUxgrYh7cKbToj5Fq_-k_roNQS2Hco2s-NDJ7XQQ";
        //MsmRequestNo = !string.IsNullOrWhiteSpace(context.Request.Params["requestNumber"]) ? int.Parse(context.Request.Params["requestNumber"]) : 0;
        //lastLocation = !string.IsNullOrWhiteSpace(context.Request.Params["lastLocation"]) ? int.Parse(context.Request.Params["lastLocation"]) : 0;

        //this.MarvalHost = context.Request.Params["host"] ?? string.Empty;
        var getParamVal = context.Request.Params["endpoint"] ?? string.Empty;
        Log.Information("endpoint is, ", getParamVal);
        Log.Information("msm base yurl is " + MSMBaseUrl);
        switch (param)
        {


            case "GET":
                MsmRequestNo = !string.IsNullOrWhiteSpace(context.Request.Params["requestNumber"]) ? int.Parse(context.Request.Params["requestNumber"]) : 0;
                lastLocation = !string.IsNullOrWhiteSpace(context.Request.Params["lastLocation"]) ? int.Parse(context.Request.Params["lastLocation"]) : 0;

                this.MarvalHost = context.Request.Params["host"] ?? string.Empty;
                //var getParamVal = context.Request.Params["endpoint"] ?? string.Empty;
                // Trace.Write("paramval is" + getParamVal);
                // Log.information
                Log.Information("paramval is" + getParamVal);
                if (getParamVal == "none")
                {



                    context.Response.Write("Hi");
                }
                // }else if (getParamVal == "getClientID"){
                //     var response = PostRequest("https://graph.microsoft.com/v1.0/sites/root");
                //      context.Response.Write(response);
                // }
                else if (getParamVal == "ChatbotHostOverride")
                {
                    context.Response.Write("Hi");
                }
                else if (getParamVal == "ClientID")
                {
                    context.Response.Write(ClientID);
                    // }else if (getParamVal == "getSites"){
                    //     var appToken = context.Request.Params["apptoken"];
                    //     Log.Information("apptoken is: " + appToken);
                    //     string ex = GetRequest("https://graph.microsoft.com/v1.0/sites?search=*", appToken);
                    //     context.Response.Write(ex);
                    // }}
                }else if (getParamVal == "SuggestDirectoryName")
                {
                    context.Response.Write(SuggestDirectoryName);
                }else if (getParamVal == "createFolderOption")
                {
                    context.Response.Write(createFolderOption);
                }
                else if (getParamVal == "CustomAttribute")
                {
                    context.Response.Write(CustomAttribute);

                } else if (getParamVal == "AutomateSubfolderCreation")
                {
                    context.Response.Write(AutomateSubfolderCreation);
                }
                else if (getParamVal == "generatePassword")
                {
                    context.Response.Write("Hi");
                }
                else if (getParamVal == "TenantID")
                {
                    context.Response.Write("Hi");
                }
                else if (getParamVal == "getprivatekey")
                {
                    var response = PostRequest("https", "");
                    context.Response.Clear();
                    context.Response.ContentType = "application/octet-stream"; // or "text/plain" if it's text
                    context.Response.AddHeader("Content-Disposition", "attachment; filename=privatekey.txt");
                    context.Response.Write(response);
                    context.Response.Flush();
                    context.Response.End();

                }
                else if (getParamVal == "getpublickey")
                {
                    var response = PostRequest("https:", "");
                    context.Response.Clear();
                    context.Response.ContentType = "application/octet-stream"; // or "text/plain" if it's text
                    context.Response.AddHeader("Content-Disposition", "attachment; filename=publickey.txt");
                    context.Response.Write(response);
                    context.Response.Flush();
                    context.Response.End();
                }
                else if (getParamVal == "getchatbotsnippet")
                {
                    var response = PostRequest("https:/i/server/downloadFile", "{");
                    context.Response.Clear();
                    context.Response.ContentType = "application/octet-stream"; // or "text/plain" if it's text
                    context.Response.AddHeader("Content-Disposition", "attachment; filename=chatbotsnippet.txt");
                    context.Response.Write(response);
                    context.Response.Flush();
                    context.Response.End();
                }
                else if (getParamVal == "databaseValue")
                {
                    string jsontwo = this.GetCustomersJSON(context.Request.Params["CIId"]);

                    context.Response.Write(jsontwo);
                }
                else if (getParamVal == "AADObjectGUIDLocation")
                {
                    context.Response.Write("Hi");
                }
                else if (getParamVal == "SecretKey")
                {
                    context.Response.Write("Hi");
                } else if (getParamVal == "getAllNotes")
                {
                    try
                    {
                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];

                        string attachmentUrl = MSMBaseUrl + "/api/serviceDesk/operational/requests/" + reqId + "/notes";
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        Log.Information("we are at line 367");

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            context.Response.ContentType = "application/octet-stream"; // optionally set MIME type
                            context.Response.OutputStream.Write(attachmentData, 0, attachmentData.Length);
                        }
                        else
                        {
                            Log.Warning("No attachment data received from MSM API");
                            context.Response.Write("No attachment data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                } else if (getParamVal == "getCustomAttributeValues")
                {
                    try
                    {
                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];

                        string attachmentUrl = MSMBaseUrl + "/api/serviceDesk/operational/requests/" + reqId;
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        ;

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            context.Response.ContentType = "application/octet-stream"; // optionally set MIME type
                            context.Response.OutputStream.Write(attachmentData, 0, attachmentData.Length);
                        }
                        else
                        {
                            Log.Warning("No attachment data received from MSM API");
                            context.Response.Write("No attachment data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                }
                else if (getParamVal == "getAllAttachments")
                {
                    try
                    {
                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];

                        string attachmentUrl = MSMBaseUrl + "/api/serviceDesk/operational/requests/" + reqId + "/attachments";
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        Log.Information("we are at line 367");

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            context.Response.ContentType = "application/octet-stream"; // optionally set MIME type
                            context.Response.OutputStream.Write(attachmentData, 0, attachmentData.Length);
                        }
                        else
                        {
                            Log.Warning("No attachment data received from MSM API");
                            context.Response.Write("No attachment data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                } else if (getParamVal == "getAllEmails")
                {
                    try
                    {//not even using this
                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];

                        string attachmentUrl = "https://localhost/MSM/RFP/Forms/RequestEmailAuditViewer.aspx?id=" + reqId;
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        Log.Information("we are at line 367");

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            context.Response.ContentType = "application/octet-stream"; // optionally set MIME type
                            context.Response.OutputStream.Write(attachmentData, 0, attachmentData.Length);
                        }
                        else
                        {
                            Log.Warning("No emails data received from MSM API");
                            context.Response.Write("No emails data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during email : " + e.Message, e);
                        context.Response.Write("Error getting emails: " + e.Message);
                    }
                }

                else
                {
                    context.Response.Write("Something is not working");
                }
                break;
            case "POST":
                Log.Information("we are at 434, with getparamval" + getParamVal);

                if (getParamVal == "getAttachment") //upload attachment
                {
                    Log.Information("we are in if");

                    try
                    {
                        Log.Information("we are in try in attachment");

                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];
                        var attachmentName = context.Request.QueryString["attachmentName"];
                        //var microsoftAccessToken = context.Request.QueryString["microsoftToken"];
                        using (var reader = new StreamReader(context.Request.InputStream))
                        {
                            requestBody = reader.ReadToEnd();//read contents from body in frontend
                        }
                        dynamic parsedBody = JsonConvert.DeserializeObject(requestBody);

                        // Step 1: Get the attachment from the MSM API
                        string attachmentUrl = MSMBaseUrl+"/api/serviceDesk/operational/requests/" + reqId + "/attachments/" + identifer + "/content?mode=Attachment";
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        Log.Information("we are at line 367");

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            // Step 2: Upload to SharePoint
                            //string getAllDocumentsUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/items/root/children";
                            string sharePointUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/root:/test/" + attachmentName + ":/content";
                            // string newUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,"+parsedBody.siteId+"+attachmentName+":/content";
                            string url3 = "https://graph.microsoft.com/v1.0/sites/" + parsedBody.siteId + "/drive/root:/" + parsedBody.folderId + "/" + attachmentName + ":/content";
                            // Create request for SharePoint upload
                            Log.Information("url3 of uploading to " + url3);
                            HttpWebRequest sharePointRequest = (HttpWebRequest)WebRequest.Create(url3);
                            sharePointRequest.Method = "PUT";
                            sharePointRequest.Headers["Authorization"] = "Bearer " + parsedBody.microsoftToken; //get token from frontend

                            sharePointRequest.ContentType = "application/octet-stream";
                            sharePointRequest.ContentLength = attachmentData.Length;
                            Log.Information("data len " + attachmentData.Length);
                            // Write the attachment data to the request stream
                            using (Stream requestStream = sharePointRequest.GetRequestStream())
                            {
                                requestStream.Write(attachmentData, 0, attachmentData.Length);
                            }
                            // Execute the SharePoint upload request
                            //var sharePointResponse = this.ProcessRequest2(sharePointRequest);
                            //Log.Information("Attachment uploaded successfully to SharePoint", sharePointResponse);
                            Log.Information("we are at line 486");
                            context.Response.Write(this.ProcessRequest2(sharePointRequest));
                        }
                        else
                        {
                            Log.Warning("No attachment data received from MSM API");
                            context.Response.Write("No attachment data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                }else if (getParamVal == "uploadCombinedNotes")
                {
                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        string json = reader.ReadToEnd();
                        var data = JsonConvert.DeserializeObject<RequestData>(json); // this should match the frontend JSON shape
                        Log.Information("data combined text is " + data.combinedText);

                        // Write combined text to a temporary .txt file
                        string tempFilePath = Path.GetTempFileName();
                        string targetFileName = Path.ChangeExtension(tempFilePath, ".txt");

                        File.WriteAllText(targetFileName, data.combinedText); // write note content to file

                        // Read file into byte[] for uploading
                        byte[] attachmentData = File.ReadAllBytes(targetFileName);

                        // Upload to SharePoint via Graph API
                        string url3 = "https://graph.microsoft.com/v1.0/sites/" + data.siteId
                            + "/drive/root:/" + data.folderId
                            + "/" + data.attachmentName + ":/content";

                        Log.Information("Uploading to URL: " + url3);

                        HttpWebRequest sharePointRequest = (HttpWebRequest)WebRequest.Create(url3);
                        sharePointRequest.Method = "PUT";
                        sharePointRequest.Headers["Authorization"] = "Bearer " + data.microsoftToken;
                        sharePointRequest.ContentType = "application/octet-stream";
                        sharePointRequest.ContentLength = attachmentData.Length;

                        Log.Information("Uploading " + attachmentData.Length + " bytes");

                        using (Stream requestStream = sharePointRequest.GetRequestStream())
                        {
                            requestStream.Write(attachmentData, 0, attachmentData.Length);
                        }

                        // Read SharePoint response
                        var uploadResponse = this.ProcessRequest2(sharePointRequest); // assuming this returns a JSON string or message

                        Log.Information("Upload finished");

                        context.Response.ContentType = "application/json";
                        context.Response.Write(JsonConvert.SerializeObject(new
                        {
                            status = "success",
                            file = data.attachmentName,
                            response = uploadResponse
                        }));
                        return;
                    }
                }

                if (getParamVal == "uploadNote") //upload attachment
                {
                    Log.Information("we are in if");

                    try
                    {
                        Log.Information("we are in try in upload notes");

                        string requestBody;
                        var identifer = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];
                        var attachmentName = context.Request.QueryString["attachmentName"];
                        //var microsoftAccessToken = context.Request.QueryString["microsoftToken"];
                        using (var reader = new StreamReader(context.Request.InputStream))
                        {
                            requestBody = reader.ReadToEnd();//read contents from body in frontend
                        }
                        dynamic parsedBody = JsonConvert.DeserializeObject(requestBody);

                        // Step 1: Get the attachment from the MSM API
                        string attachmentUrl = MSMBaseUrl + "/api/serviceDesk/operational/requests/" + reqId + "/notes/" + identifer;
                        Log.Information("url here is" + attachmentUrl);
                        httpWebRequest = BuildRequest(attachmentUrl);
                        httpWebRequest.Headers["Authorization"] = "Bearer " + MarvalAPIKey;
                        httpWebRequest.Method = "GET";
                        // Get the attachment data as byte array
                        //byte[] attachmentData = this.ProcessRequestAsBytes(httpWebRequest);
                        string noteJson = this.ProcessRequest2(httpWebRequest); // raw JSON string
                        dynamic noteData = JsonConvert.DeserializeObject(noteJson); //deserialise so we cna access the data in the body
                        Log.Information("noteJson = " + noteJson);
                        Log.Information("contentSummary = " + noteData.entity.data.contentSummary);


                        // Build a plain-text file content
                        string textContent =
                        "Created On: " + noteData.entity.data.createdOn + "\r\n" +
                        "Author: " + noteData.entity.data.author.name + "\r\n" +
                        "Summary: " + noteData.entity.data.contentSummary;



                        // Convert the text to byte array for upload
                        byte[] attachmentData = Encoding.UTF8.GetBytes(textContent); //need to remember that needs to be in bytes format when we are doing stream write

                        // Optional: append .txt extension to make it clearer
                        if (!attachmentName.EndsWith(".txt"))
                        {
                            attachmentName += ".txt";
                        }

                        Log.Information("we are at line 367");

                        if (attachmentData != null && attachmentData.Length > 0)
                        {
                            // Step 2: Upload to SharePoint
                            //string getAllDocumentsUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/items/root/children";
                            string sharePointUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/root:/test/" + attachmentName + ":/content";
                            // string newUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,"+parsedBody.siteId+"+attachmentName+":/content";
                            string url3 = "https://graph.microsoft.com/v1.0/sites/" + parsedBody.siteId + "/drive/root:/" + parsedBody.folderId + "/" + attachmentName + ":/content";
                            Log.Information("url3 of uploading to " + url3); //are getting here
                                                                             // Create request for SharePoint upload
                            HttpWebRequest sharePointRequest = (HttpWebRequest)WebRequest.Create(url3);
                            sharePointRequest.Method = "PUT";
                            Log.Information("line 611");
                            sharePointRequest.Headers["Authorization"] = "Bearer " + parsedBody.microsoftToken; //get token from frontend

                            sharePointRequest.ContentType = "application/octet-stream";
                            sharePointRequest.ContentLength = attachmentData.Length;
                            Log.Information("data len " + attachmentData.Length);
                            // Write the attachment data to the request stream
                            using (Stream requestStream = sharePointRequest.GetRequestStream())
                            {
                                requestStream.Write(attachmentData, 0, attachmentData.Length);
                            }
                            // Execute the SharePoint upload request
                            //var sharePointResponse = this.ProcessRequest2(sharePointRequest);
                            //Log.Information("Attachment uploaded successfully to SharePoint", sharePointResponse);
                            Log.Information("we are at line 486");
                            context.Response.Write(this.ProcessRequest2(sharePointRequest));
                            //ddMsmNote(Int32.Parse(reqId), "testing note from backend handler");
                            Log.Information("we are creating a note");
                        }
                        else
                        {
                            Log.Warning("No attachment data received from MSM API");
                            context.Response.Write("No attachment data found");
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                }
                else if (getParamVal == "createUploadNote")
                {
                    var reqNum = context.Request.QueryString["reqId"];
                    AddMsmNote(Int32.Parse(reqNum), "Successfully uploaded all attachments!");
                }


                else if (getParamVal == "uploadEmail") //upload attachment
                {
                    Log.Information("we are in if");
                    string json;
                    string url = "";
                    string attachmentName = context.Request.QueryString["attachmentName"];


                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        json = reader.ReadToEnd();
                    }

                    dynamic parsedBody;
                    Log.Information("Raw JSON payload received from frontend: " + json);
                    try
                    {

                        parsedBody = JsonConvert.DeserializeObject(json);//read body from request
                        Log.Information("parsed body in emails ios " + parsedBody.folderPath);

                    }

                    catch (JsonException)
                    {
                        context.Response.StatusCode = 400; // Bad Request
                        context.Response.Write("Invalid JSON");
                        context.Response.End();
                        return;
                    }
                    try
                    {

                        // Step 2: Upload to SharePoint
                        //string getAllDocumentsUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/items/root/children";
                        string sharePointUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,04f24f61-1573-410f-b54d-3ab2c7784161,6ee23755-585f-477d-bf49-4a114bca65df/drive/root:/test/" + attachmentName + ":/content";
                        // string newUrl = "https://graph.microsoft.com/v1.0/sites/marvaluk.sharepoint.com,"+parsedBody.siteId+"+attachmentName+":/content";
                        string url3 = "https://graph.microsoft.com/v1.0/sites/" + parsedBody.siteId + "/drive/root:/" + parsedBody.folderId + "/" + attachmentName + ":/content";
                        Log.Information("url3 of uploading to " + url3);
                        //Log.Information("parsed ")
                        // Create request for SharePoint upload
                        HttpWebRequest sharePointRequest = (HttpWebRequest)WebRequest.Create(url3);
                        sharePointRequest.Method = "PUT";
                        Log.Information("line 682");
                        sharePointRequest.Headers["Authorization"] = "Bearer " + parsedBody.microsoftToken; //get token from frontend

                        string subject = context.Request.QueryString["attachmentName"];
                        attachmentName = Path.GetFileNameWithoutExtension(subject) + ".eml";

                        // Construct .eml content

                        string emlContent =
"From: " + parsedBody.address + "\r\n" +
"To: unknown@recipient.com\r\n" +
"Subject: " + attachmentName + "\r\n" +
"Date: " + parsedBody.messageDate + "\r\n" +
"Content-Type: text/plain; charset=UTF-8\r\n" +
"\r\n" +
"This is an auto-uploaded email record for subject: " + attachmentName + "\r\n";


                        // Convert to byte array
                        byte[] attachmentData = Encoding.UTF8.GetBytes(emlContent);


                        sharePointRequest.ContentType = "application/octet-stream";
                        sharePointRequest.ContentLength = attachmentData.Length;
                        Log.Information("data len " + attachmentData.Length);
                        // Write the attachment data to the request stream
                        using (Stream requestStream = sharePointRequest.GetRequestStream())
                        {
                            requestStream.Write(attachmentData, 0, attachmentData.Length);
                        }
                        // Execute the SharePoint upload request
                        //var sharePointResponse = this.ProcessRequest2(sharePointRequest);
                        //Log.Information("Attachment uploaded successfully to SharePoint", sharePointResponse);
                        Log.Information("we are at line 486");
                        context.Response.Write(this.ProcessRequest2(sharePointRequest));


                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during attachment upload: " + e.Message, e);
                        context.Response.Write("Error uploading file: " + e.Message);
                    }
                }
                else if (getParamVal == "getAllFolders")
                {
                    string json;
                    string url = "";


                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        json = reader.ReadToEnd();
                    }

                    RequestData data;
                    try
                    {

                        data = JsonConvert.DeserializeObject<RequestData>(json);//read body from request
                    }
                    catch (JsonException)
                    {
                        context.Response.StatusCode = 400; // Bad Request
                        context.Response.Write("Invalid JSON");
                        context.Response.End();
                        return;
                    }
                    //var action = data.action;
                    var apptoken = data.apptoken;
                    Log.Information("apptoken is", apptoken);
                    Log.Information("data is" + data);
                    string siteId = data.siteId;
                    string folderPath = data.folderPath;
                    Log.Information("Site id from my frontend is " + siteId);

                    if (!String.IsNullOrEmpty(folderPath))//if exists
                    {
                        url = "https://graph.microsoft.com/v1.0/sites/"+siteId+"/drive/root:/" + folderPath + ":/children";
                    }
                    else
                    {
                        url = "https://graph.microsoft.com/v1.0/sites/"+siteId+"/drive/items/root/children"; //meaning get root folders
                    }

                    Log.Information("apptoken is: " + apptoken);
                    Log.Information("url of get all folders is " + url);
                    string ex = GetRequest((url), apptoken);
                    context.Response.Write(ex);
                }
                else if (getParamVal == "getSites")
                {
                    string json;

                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        json = reader.ReadToEnd();
                    }

                    RequestData data;
                    try
                    {

                        data = JsonConvert.DeserializeObject<RequestData>(json);
                    }
                    catch (JsonException)
                    {
                        context.Response.StatusCode = 400; // Bad Request
                        context.Response.Write("Invalid JSON");
                        context.Response.End();
                        return;
                    }
                    //var action = data.action;
                    var apptoken = data.apptoken;
                    Log.Information("apptoken is", apptoken);
                    Log.Information("data is" + data);

                    Log.Information("apptoken is: " + apptoken);
                    string ex = GetRequest("https://graph.microsoft.com/v1.0/sites?search=*", apptoken);
                    context.Response.Write(ex);
                }
                else if (getParamVal == "createFolder")
                {
                    Log.Information("Entering 'createFolder' block.");

                    try
                    {
                        Log.Information("Processing folder creation...");

                        // Extract query parameters
                        var identifier = context.Request.QueryString["identifier"];
                        var reqId = context.Request.QueryString["reqId"];
                        var attachmentName = context.Request.QueryString["attachmentName"];
                        Log.Information("Processing folder creation... with the attachment name " + attachmentName);

                        // Read JSON body from request
                        string requestBody;
                        using (var reader = new StreamReader(context.Request.InputStream))
                        {
                            requestBody = reader.ReadToEnd();
                        }

                        dynamic parsedBody = JsonConvert.DeserializeObject(requestBody);
                        string siteId = parsedBody.siteId;
                        string folderId = parsedBody.folderId;//should be the folder path
                        string microsoftToken2 = parsedBody.microsoftToken;

                        // Create SharePoint folder creation URL
                        string url = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/drive/items/root:/" + folderId + ":/children";

                        // Setup request to create folder
                        HttpWebRequest request2 = (HttpWebRequest)WebRequest.Create(url);
                        request2.Method = "POST";
                        request2.Headers["Authorization"] = "Bearer " + microsoftToken2;
                        request2.ContentType = "application/json";

                        // Build request body to create folder
                        var body = new Dictionary<string, object>
{
    { "name", attachmentName },
    { "folder", new { } },
    { "@microsoft.graph.conflictBehavior", "rename" }
};



                        string jsonBody = JsonConvert.SerializeObject(body);

                        using (var streamWriter = new StreamWriter(request2.GetRequestStream()))
                        {
                            streamWriter.Write(jsonBody);
                        }

                        // Send request to Graph API
                        var responseContent = this.ProcessRequest2(request2);
                        Log.Information("Folder created successfully.");
                        context.Response.Write(responseContent);
                    }
                    catch (Exception e)
                    {
                        Log.Error("Exception during folder creation: " + e.Message, e);
                        context.Response.Write("Error creating folder: " + e.Message);
                    }
                }



                break;
        }
    }
    private string ProcessRequest2(HttpWebRequest request)
    {
        try
        {
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (Stream responseStream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(responseStream))
                    {
                        string responseText = reader.ReadToEnd();
                        // Log the response status
                        //Log.Information("HTTP Status:"+ response.StatusCode +response.StatusDescription);
                        // For SharePoint uploads, successful responses are usually 200 or 201
                        if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created)
                        {
                            //Log.Information("SharePoint upload successful");
                            return responseText;
                        }
                        else
                        {
                            // Log.Warning($"Unexpected status code: {response.StatusCode}");
                            return responseText;
                        }
                    }
                }
            }
        }
        catch (WebException webEx)
        {
            // Handle web-specific errors
            if (webEx.Response != null)
            {
                using (HttpWebResponse errorResponse = (HttpWebResponse)webEx.Response)
                {
                    using (Stream errorStream = errorResponse.GetResponseStream())
                    {
                        using (StreamReader errorReader = new StreamReader(errorStream))
                        {
                            string errorText = errorReader.ReadToEnd();
                            Log.Error("Web Exception - Status: {errorResponse.StatusCode}, Error:" + errorText);
                            throw new Exception("SharePoint API Error: {errorResponse.StatusCode} - " + errorText);
                        }
                    }
                }
            }
            else
            {
                Log.Error("Web Exception without response: "+webEx.Message);
                throw new Exception("Network Error: {webEx.Message}");
                return "";
            }
        }
        catch (Exception ex)
        {
            return "";
            Log.Error("General Exception in ProcessRequest: " +ex.Message);
            throw;
        }
        return "";
    }

    // Alternative version that returns more detailed response info
    private SharePointUploadResponse ProcessRequestDetailed(HttpWebRequest request)
    {
        try
        {
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (Stream responseStream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(responseStream))
                    {
                        string responseText = reader.ReadToEnd();
                        return new SharePointUploadResponse
                        {
                            StatusCode = response.StatusCode,
                            StatusDescription = response.StatusDescription,
                            Content = responseText,
                            IsSuccess = response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created,
                            ContentType = response.ContentType,
                            Headers = response.Headers
                        };
                    }
                }
            }
        }
        catch (WebException webEx)
        {
            if (webEx.Response != null)
            {
                using (HttpWebResponse errorResponse = (HttpWebResponse)webEx.Response)
                {
                    using (Stream errorStream = errorResponse.GetResponseStream())
                    {
                        using (StreamReader errorReader = new StreamReader(errorStream))
                        {
                            string errorText = errorReader.ReadToEnd();
                            return new SharePointUploadResponse
                            {
                                StatusCode = errorResponse.StatusCode,
                                StatusDescription = errorResponse.StatusDescription,
                                Content = errorText,
                                IsSuccess = false,
                                Error = "SharePoint API Error: {errorResponse.StatusCode} - " + errorText
                            };
                        }
                    }
                }
            }
            else
            {
                return new SharePointUploadResponse
                {
                    IsSuccess = false,
                    Error = "Network Error: "+webEx.Message
                };
            }
        }
        catch (Exception ex)
        {
            return new SharePointUploadResponse
            {
                IsSuccess = false,
                Error = "General Exception: "+ex.Message
            };
        }
    }
    // Helper method to process request and return byte array
    private byte[] ProcessRequestAsBytes(HttpWebRequest request)
    {
        try
        {
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream responseStream = response.GetResponseStream())
            using (MemoryStream memoryStream = new MemoryStream())
            {
                responseStream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
        catch (Exception ex)
        {
            Log.Error("Error processing request as bytes: " + ex.Message, ex);
            return null;
        }
    }
    private string GetDBString()
    {
        string connectionString = "";

        string msmdLocation = GetAppPath("MSM");
        string path = msmdLocation;
        string newPath = Path.GetFullPath(Path.Combine(path, @"..\"));
        string openFilePath = newPath + "connectionStrings.config";

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(openFilePath);

        XmlNodeList nodeList = xmlDoc.SelectNodes("/appSettings/add[@key='DatabaseConnectionString']");

        if (nodeList.Count > 0)
        {
            // Get the value attribute of the node
            connectionString = nodeList[0].Attributes["value"].Value;
        }

        else
        {
        }
        return connectionString;
    }
    private string GetAppPath(string productName)
    {
        const string foldersPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders";
        var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);

        var subKey = baseKey.OpenSubKey(foldersPath);
        if (subKey == null)
        {
            baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            subKey = baseKey.OpenSubKey(foldersPath);
        }
        return subKey != null ? subKey.GetValueNames().FirstOrDefault(kv => kv.Contains(productName)) : "ERROR";
    }
    internal class JsonHelper
    {
        public static string ToJson(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }

        public static dynamic FromJson(string json)
        {
            return JObject.Parse(json);
        }
    }

    private string GetCustomersJSON(string CIId)
    {
        string connString = GetDBString();
        using (SqlConnection conn = new SqlConnection())
        {
            conn.ConnectionString = connString;
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.CommandText = "select guid from directoryRelationship where CIId = " + CIId;
                cmd.Connection = conn;
                conn.Open();
                string returnVal = "";
                using (SqlDataReader sdr = cmd.ExecuteReader())
                {
                    sdr.Read();
                    returnVal = Convert.ToString(sdr["guid"]);
                }
                conn.Close();

                return returnVal;
            }
        }
    }

}
