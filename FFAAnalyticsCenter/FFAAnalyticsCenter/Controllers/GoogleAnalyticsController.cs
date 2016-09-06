using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web;
using System.Web.Mvc;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Analytics.v3;
using Google.Apis.Analytics.v3.Data;
using Google.Apis.Services;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;

namespace FFAAnalyticsCenter.Controllers
{
    public class GoogleAnalyticsController : Controller
    {
       public AnalyticsService Service { get; set; }

       public GoogleAnalyticsController(string keyPath, string accountEmailAddress)
        {
            var certificate = new X509Certificate2(keyPath, "notasecret", X509KeyStorageFlags.Exportable);

            var credentials = new ServiceAccountCredential(
               new ServiceAccountCredential.Initializer(accountEmailAddress)
               {
                   Scopes = new[] { AnalyticsService.Scope.AnalyticsReadonly }
               }.FromCertificate(certificate));

            Service = new AnalyticsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = "FFAGoogleAnalytics"
            });
        }
        public AnalyticDataPoint GetAnalyticsData(string profileId, string dimensions, string metrics, string sort, string filter, DateTime startDate, DateTime endDate, int maxResults)
        {
            AnalyticDataPoint data = new AnalyticDataPoint();
            if (!profileId.Contains("ga:"))
                profileId = string.Format("ga:{0}", profileId);

            var request = BuildAnalyticRequest(profileId, dimensions, metrics, sort, filter, startDate, endDate, 1, maxResults);
            GaData response = request.Execute();

            data.ColumnHeaders = response.ColumnHeaders;

            if (response.Rows == null || response.Rows.Count == 0)
            {
                var dummyData = new List<string>();
                for (var y = 0; y < data.ColumnHeaders.Count; y++)
                {
                    dummyData.Add("0");
                }
                dummyData[0] = "No Data Returned";
                data.Rows.Add(dummyData);
            }
            else
            {
                data.Rows.AddRange(response.Rows);
            }

            return data;
        }
        private DataResource.GaResource.GetRequest BuildAnalyticRequest(string profileId, string dimensions, string metrics,
                                                                            string sort, string filter, DateTime startDate,
                                                                            DateTime endDate, int startIndex, int maxResults)
        {
            DataResource.GaResource.GetRequest request = Service.Data.Ga.Get(profileId, startDate.ToString("yyyy-MM-dd"),
                                                                                endDate.ToString("yyyy-MM-dd"), string.Join(",", metrics));
            request.Dimensions = dimensions;
            request.StartIndex = startIndex;
            request.Filters =  filter;
            request.Sort = sort;
            request.MaxResults = maxResults;
            return request;
        }
        public class AnalyticDataPoint
        {
            public AnalyticDataPoint()
            {
                Rows = new List<IList<string>>();
            }

            public IList<GaData.ColumnHeadersData> ColumnHeaders { get; set; }
            public List<IList<string>> Rows { get; set; }
        }
	}
}