using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace OutlookService
{
    /// <summary>
    /// Provides methods to call the Outlook API.
    /// </summary>
    public class OutlookService
    {
        private const string apiEndpoint = "https://outlook.office.com/api/";

        private string apiVersion;
        private string userEmail;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookService"/> class, configured
        /// to the specified API version.
        /// </summary>
        /// <param name="version">The version of the Outlook API to use. If not specified, the latest released version is used.</param>
        public OutlookService(Apiversion version = Apiversion.V2)
        {
            apiVersion = version.ToString("g").ToLower();
        }

        /// <summary>
        /// Makes a generic API call to the Outlook API
        /// </summary>
        /// <param name="accessToken">The access token to use for the API call.</param>
        /// <param name="apiUrl">The API URL segment beginning with the user specification. Do not include a query string, instead
        /// add query parameters in the <paramref name="queryParams"/> parameter. For example, <code>/Me/Messages</code>.</param>
        /// <param name="method">The HTTP method to use for the call. Default is <code>GET</code>.</param>
        /// <param name="preferHeaders">A list of key/value paris to add to the <code>Prefer</code> header.</param>
        /// <param name="queryParams">A list of query parameters to add to the API request.</param>
        /// <param name="payload">A JSON payload to add to the request.</param>
        public async Task<HttpResponseMessage> MakeApiCall(string accessToken, string apiUrl, string method = "GET", List<KeyValuePair<string, string>> preferHeaders = null,
            List<KeyValuePair<string, string>> queryParams = null, string payload = null)
        {
            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage(new HttpMethod(method), FullyQualifyApiUrl(apiUrl, queryParams));

                // Headers
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("OutlookService", "1.0"));
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                request.Headers.Add("return-client-request-id", "true");

                // If user is set, add it as the anchor
                if (!string.IsNullOrEmpty(userEmail))
                {
                    request.Headers.Add("X-AnchorMailbox", userEmail);
                }

                // If additional headers were supplied, set them
                if (preferHeaders != null)
                {
                    foreach(KeyValuePair<string, string> header in preferHeaders)
                    {
                        request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
                    }
                }

                // Content
                if ((string.Compare(method, "POST", true) == 0 || string.Compare(method, "PATCH", true) == 0) &&
                    !string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload);
                    request.Content.Headers.ContentType.MediaType = "application/json";
                }

                return await httpClient.SendAsync(request);
            }
        }

        private string GetBaseApiUrl()
        {
            return string.Format("{0}{1}", apiEndpoint, apiVersion);
        }

        private string FullyQualifyApiUrl(string partialUrl, List<KeyValuePair<string, string>> queryParams)
        {
            if (partialUrl.StartsWith("/"))
            {
                return string.Format("{0}{1}{2}", GetBaseApiUrl(), partialUrl, GetQueryString(queryParams));
            }

            return string.Format("{0}/{1}{2}", GetBaseApiUrl(), partialUrl, GetQueryString(queryParams));
        }

        private string GetQueryString(List<KeyValuePair<string, string>> queryParams)
        {
            if (queryParams == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder("?");
            var separator = "";

            foreach (var parm in queryParams)
            {
                builder.AppendFormat("{0}{1}={2}", separator, WebUtility.UrlEncode(parm.Key), WebUtility.UrlEncode(parm.Value));
                separator = "&";
            }

            return builder.ToString();
        }
    }
}
