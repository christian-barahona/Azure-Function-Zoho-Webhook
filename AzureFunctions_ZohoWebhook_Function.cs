using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ZohoWebhook
{
    public static class Function
    {
        [FunctionName( "Function" )]
        public static async Task<IActionResult> RunAsync(
            [HttpTrigger( AuthorizationLevel.Function, "get", "post", Route = null )] HttpRequest req,
            ILogger log )
        {
            log.LogInformation( "C# HTTP trigger function processed a request. \r\n" );

            string quoteNumber = req.Query[ "quote_number" ];
            string quoteType = req.Query[ "quote_type" ];

            string requestBody = await new StreamReader( req.Body ).ReadToEndAsync( );
            dynamic data = JsonConvert.DeserializeObject( requestBody );
            quoteNumber = quoteNumber ?? data?.quote_number;
            quoteType = quoteType ?? data?.quote_type;

            var instantiate = new HtmlToPdf( log );
            await instantiate.GeneratePdfAsync( "", "test" );

            return quoteNumber != null && quoteType != null
                ? (ActionResult)new OkObjectResult( $"Quote # {quoteNumber}, Quote Type: {quoteType}" )
                : new BadRequestObjectResult( "Quote number and quote type required" );
        }
    }

    public class HtmlToPdf
    {
        public HtmlToPdf( ILogger log )
        {
            m_Log = log;
            m_HttpClient = new HttpClient( );
        }

        public async Task GeneratePdfAsync( string quoteNumber, string quoteType )
        {
            var templatesPath = $@"d:\home\templates";
            var pdfsPath = $@"d:\home\pdfs";
            await ConvertAsync( );

            async Task<bool> InsertDataIntoHtmlFileAsync( )
            {
                var htmlFile = $@"{templatesPath}\{quoteType}.html";

                if( !File.Exists( htmlFile ) )
                {
                    m_Log.LogInformation( "Template does not exist!" );

                    return false;
                }
                else
                {
                    var ( quote, account ) = await GetDataAsync( quoteNumber );
                    var variablePattern = @"\${(.*?)\.(.*?)}";
                    var html = File.ReadAllText( htmlFile );
                    var templateVariables = Regex.Matches( html, variablePattern );

                    var i = 0;
                    foreach (Match foo in templateVariables)
                    {
                        i++;
                        m_Log.LogInformation( $"Var{i}: {foo.Groups[ 0 ].Value}" );
                    }

                    i = 0;
                    foreach( Match templateVariable in templateVariables )
                    {
                        i++;
                        var placeholder = templateVariable.Groups[ 0 ].Value;  // ${Quotes.Quote No.}
                        var module = templateVariable.Groups[ 1 ].Value;  // Quotes
                        var field = templateVariable.Groups[ 2 ].Value;  // Quote No.
                        var data = "";
                        m_Log.LogInformation( $"Field{i}: {field}" );

                        if( module == "Quotes" || module == "Products" || module == "Opportunities")
                        {
                            switch( field )
                            {
                                case "Deal_Name":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Deal_Name" ][ "name" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Contact_Name":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Contact_Name" ][ "name" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Account_Name":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Account_Name" ][ "name" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Product_Name":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Product_Details" ][ 0 ][ "product" ][ "name" ].ToString( ); 
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Product_Code":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Product_Details" ][ 0 ][ "product" ][ "Product_Code" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Product_Description":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Product_Details" ][ 0 ][ "product_description" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                case "Quantity":
                                    m_Log.LogInformation( $"Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ "Product_Details" ][ 0 ][ "quantity" ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                                default:
                                    m_Log.LogInformation( $"Else Module: {placeholder}" );
                                    data = quote[ "data" ][ 0 ][ field ].ToString( );
                                    m_Log.LogInformation( $"Data: {data}" );
                                    break;
                            }

                            html = html.Replace( placeholder, data );
                        }

                        m_Log.LogInformation( "\r\n" );
                    }

                    File.WriteAllText( htmlFile, html );

                    return true;
                }
            }

            async Task ConvertAsync( )
            {
                var result = await InsertDataIntoHtmlFileAsync( );

                if( result )
                {
                   var renderer = new IronPdf.HtmlToPdf( );

                   renderer.PrintOptions.MarginTop = 0;
                   renderer.PrintOptions.MarginLeft = 0;
                   renderer.PrintOptions.MarginRight = 0;
                   renderer.PrintOptions.MarginBottom = 0;

                   var pdf = renderer.RenderHTMLFileAsPdf( $@"{templatesPath}\{quoteType}.html" );

                   pdf.SaveAs( $@"{pdfsPath}\{quoteNumber}.pdf" );
                }
            }
        }
        
        public async Task<dynamic> GetAccessTokenAsync( )
        {
            var url = $"https://accounts.zoho.com/oauth/v2/token?refresh_token={m_RefreshToken}&client_id={m_ClientId}&client_secret={m_ClientSecret}&grant_type=refresh_token";
            var response = await m_HttpClient.PostAsync( url, default );
            var result = await response.Content.ReadAsAsync<dynamic>( );

            return result.access_token;
        }

        public async Task<JObject> GetJsonAsync( string module, string id )
        {
            var accessToken = await GetAccessTokenAsync( );
            var url = $"https://www.zohoapis.com/crm/v2/{module}/{id}";
            m_HttpClient.DefaultRequestHeaders.Add( "Authorization", "Zoho-oauthtoken " + accessToken );
            var response = await m_HttpClient.GetAsync( url );
            var result = await response.Content.ReadAsStringAsync( );
            var json = JObject.Parse( result );

            return json;
        }

        public async Task<(JObject quote, JObject account)> GetDataAsync( string quoteNumber )
        {
            var quote = await GetJsonAsync( "quote", quoteNumber );
            var accountNumber = quote[ "data" ][ 0 ][ "Account_Name" ][ "id" ].ToString( );
            var account = await GetJsonAsync( "account", accountNumber );

            return ( quote, account );
        }

        private readonly ILogger m_Log;
        private readonly HttpClient m_HttpClient;
        private readonly string m_RefreshToken = "";
        private readonly string m_ClientId = "";
        private readonly string m_ClientSecret = "";
    }

}
