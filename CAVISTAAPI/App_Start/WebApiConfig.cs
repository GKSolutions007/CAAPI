using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Controllers;
using System.Web.Http.Cors;
using System.Web.Http.Filters;

namespace CAVISTAAPI
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();
            string allowedOrigins = "https://localhost:44349";
            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
            var cors = new EnableCorsAttribute(origins: allowedOrigins, headers: "*", methods: "*") //   Allow only frontend URL
            {
                SupportsCredentials = true //  Allows cookies to be sent
            };
            config.EnableCors(cors);
            // Add a global message handler to append embedding headers for PDF responses
            config.MessageHandlers.Add(new PdfEmbeddingHeaderHandler(allowedOrigins));

            config.Formatters.JsonFormatter.SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/json"));
            config.Filters.Add(new SessionEnabledAttribute());
        }
    }
    public class SessionEnabledAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(HttpActionContext actionContext)
        {
            var session = HttpContext.Current.Session;
            base.OnActionExecuting(actionContext);
        }
    }
    public class PdfEmbeddingHeaderHandler : System.Net.Http.DelegatingHandler
    {
        private readonly string _allowedOrigin;
        public PdfEmbeddingHeaderHandler(string allowedOrigin)
        {
            _allowedOrigin = allowedOrigin;
        }
        protected override System.Threading.Tasks.Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request, System.Threading.CancellationToken cancellationToken)
        {
            return base.SendAsync(request, cancellationToken).ContinueWith(task =>
            {
                var response = task.Result;
                try
                {
                    if (response != null && response.Content != null)
                    {
                        var mediaType = response.Content.Headers.ContentType?.MediaType;
                        if (string.Equals(mediaType, "application/pdf", StringComparison.OrdinalIgnoreCase))
                        {
                            // Do NOT set Access-Control-Allow-Origin here to avoid duplicates with Web API CORS
                            if (!response.Headers.Contains("Vary"))
                                response.Headers.Add("Vary", "Origin");
                            // Prefer CSP over X-Frame-Options for modern control
                            if (!response.Headers.Contains("Content-Security-Policy"))
                                response.Headers.Add("Content-Security-Policy", $"frame-ancestors {_allowedOrigin}");
                            // Ensure inline display if filename is present downstream
                            if (response.Content.Headers.ContentDisposition == null)
                                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline");
                        }
                    }
                }
                catch { }
                return response;
            }, cancellationToken);
        }
    }
}
