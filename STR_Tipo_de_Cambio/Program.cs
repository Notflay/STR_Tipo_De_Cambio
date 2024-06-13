using STR_Tipo_de_Cambio.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;
using ScrapySharp.Extensions;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text;
using System.Net;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace STR_Tipo_de_Cambio
{
    class Program
    {
        static void Main(string[] args)
        {
            // Hacer logica dewl fro 
            // Obtener data del archivo de conexion.xml
            string filePath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "conexion.xml");

            try
            {
                // Cargar el archivo XML
                XDocument xmlDoc = XDocument.Load(filePath);

                // Obtener la lista de elementos SBO
                List<SBO> sboList = new List<SBO>();

                foreach (var sboElement in xmlDoc.Descendants("SBO"))
                {
                    SBO sbo = new SBO();
                    foreach (var addElement in sboElement.Elements("add"))
                    {
                        string key = (string)addElement.Attribute("key");
                        string value = (string)addElement.Attribute("value");

                        // Asignar valores al objeto SBO
                        switch (key)
                        {
                            case "SAP_SERVIDOR":
                                sbo.SAP_SERVIDOR = value;
                                break;
                            case "SAP_BASE":
                                sbo.SAP_BASE = value;
                                break;
                            case "SAP_TIPO_BASE":
                                sbo.SAP_TIPO_BASE = value;
                                break;
                            case "SAP_DBUSUARIO":
                                sbo.SAP_DBUSUARIO = value;
                                break;
                            case "SAP_DBPASSWORD":
                                sbo.SAP_DBPASSWORD = value;
                                break;
                            case "SAP_USUARIO":
                                sbo.SAP_USUARIO = value;
                                break;
                            case "SAP_PASSWORD":
                                sbo.SAP_PASSWORD = value;
                                break;
                            default:
                                // Opción predeterminada, puedes manejarla según tus necesidades
                                break;
                        }
                    }
                    // Agregar el objeto SBO a la lista
                    sboList.Add(sbo);
                }

                // Mostrar la lista de objetos SBO
                foreach (var sbo in sboList)
                {
                    Console.WriteLine($"ServerSAP: {sbo.SAP_SERVIDOR}, SBOCompany: {sbo.SAP_BASE}, SBOUser: {sbo.SAP_DBUSUARIO}, SBOPassword: {sbo.SAP_DBPASSWORD}");

                    SAPConnector.Conectar(sbo);

                    bool procesoFinaliza = false;

                    //
                    IntegrarTipoDeCambioSelenium(sbo).Wait();
                    //integrarTipoCambioADI(sbo).Wait();

                    SAPConnector.Desconectar();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo XML: {ex.Message}");
            }

        }
        private static async Task IntegrarTipoDeCambioSelenium(SBO sbo)
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); // Para ejecución en segundo plano
            IWebDriver driver = new ChromeDriver(options);

            string date = System.DateTime.Now.ToString("yyyy-MM-dd");

            // Que el sistema detecté el tipo de Cambio según el formato
            CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
            culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
            culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
            System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;

            try
            {
                List<string> TipoCambio = new List<string>();

                string url = "https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx";

                int intentos = 0;

                string html = "";
                while (intentos < 3) // Intenta 3 veces, puedes ajustar este número según tu necesidad
                {
                    try
                    {
                        driver.Navigate().GoToUrl(url);

                        html = driver.PageSource;

                        // Verificar si el HTML contiene el mensaje "Request unsuccessful"
                        if (html.Contains("Request unsuccessful"))
                        {
                            Console.WriteLine("El mensaje 'Request unsuccessful' fue detectado. Refrescando la página...");

                            driver.Navigate().Refresh();

                            intentos++;
                            System.Threading.Thread.Sleep(1000);
                        }
                        else
                        {
                            break;
                        }
                    }
                    catch (WebDriverException ex)
                    {
                        Console.WriteLine($"Error al cargar la página: {ex.Message}");

                        // Incrementar el contador de intentos
                        intentos++;

                        // Esperar un tiempo antes de intentar de nuevo (por ejemplo, 5 segundos)
                        System.Threading.Thread.Sleep(1000);
                    }
                }

             //   string html = driver.PageSource;

                // Si el contador de intentos es igual al número máximo de intentos, mostrar un mensaje de error
                if (intentos == 3)
                {
                    Console.WriteLine("Se ha alcanzado el número máximo de intentos. La página no se pudo cargar.");
                }


                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                HtmlNode Body = doc.DocumentNode.CssSelect("body").First();
                string sbody = Body.InnerHtml;
                try
                {
                    Log.WriteToFile(doc.DocumentNode.CssSelect(".APLI_fila2").ToString());
                }
                catch (Exception)
                {

                    throw;
                }
    
                foreach (var Node in doc.DocumentNode.CssSelect(".APLI_fila2"))
                {
                    TipoCambio.Add(Node.InnerHtml);
                }
                if (string.IsNullOrEmpty(TipoCambio[1]))
                    throw new Exception("Aun no se ha actualizado el tipo de Cambio SBS");
                double tipoCambio = Convert.ToDouble(TipoCambio[1]);

                // Cerrar el navegador
                driver.Quit();

                SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                Console.WriteLine("Actualizando Tipo de Cambio " + tipoCambio  + " " + date + ".....");
                Log.WriteToFile("Actualizando Tipo de Cambio " + date + ".....");
                bo.SetCurrencyRate("USD", System.DateTime.Now, tipoCambio, true);
                Log.WriteToFile("Tipo de Cambio del dia " + date + " : " + tipoCambio.ToString("F2"));

            }
            catch (Exception ex)
            {
                // Cerrar el navegador
                driver.Quit();
                Log.WriteToFile($"Metodo IntegrarTipoDeCambioSelenium - Error al Actualizar en SAP {sbo.SAP_BASE} -:" + ex.Message);
                IntegrarTipoCambioSBS(sbo).Wait();
            }

        }
        private static async Task IntegrarTipodeCambioSUNAT(SBO sbo)
        {
            try
            {
                // Que el sistema detecté el tipo de Cambio según el formato
                CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
                culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
                culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
                System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;

                string date = System.DateTime.Now.ToString("dd/MM/yyyy");

                List<string> TipoCambio = new List<string>();

                HttpClient client = new HttpClient();
                HttpRequestMessage request = new HttpRequestMessage();
                request.RequestUri = new Uri("https://www.sunat.gob.pe/a/txt/tipoCambio.txt");
                request.Method = HttpMethod.Get;
                var response = await client.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    string content = response.Content.ReadAsStringAsync().Result;
                    List<string> valores = content.Split('|').ToList();
                    string fecha = DateTime.Parse(valores[0]).ToShortDateString();
                    if (fecha == DateTime.Now.ToShortDateString())
                    {
                        double cambioVenta = Convert.ToDouble(valores[2]);

                        SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        Log.WriteToFile($"Actualizando Tipo de Cambio {sbo.SAP_BASE} - " + date + ".....");
                        bo.SetCurrencyRate("USD", System.DateTime.Now, cambioVenta, true);
                        Log.WriteToFile($"Tipo de Cambio del dia {sbo.SAP_BASE} - " + date + " : " + cambioVenta);
                    }
                    else
                    {
                        throw new Exception($"No se encuentra con el cambio del día de hoy");
                    };
                }
                else {
                    throw new Exception($"Error al llamar al endpoint https://www.sunat.gob.pe/a/txt/tipoCambio.txt");
                }

            }
            catch (Exception ex)
            {
                Log.WriteToFile($"Metodo CONSULTA API - Error al Actualizar en SAP {sbo.SAP_BASE} -:" + ex.Message);
                IntegrarTipoCambioSBS(sbo).Wait();
            }

        }
        private static async Task IntegrarTipoDeCambioAPI(SBO sbo)
        {
            try
            {
                CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
                culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
                culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
                System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;

                string date = System.DateTime.Now.ToString("dd/MM/yyyy");

                // Valida de Otras formas si no hay del día de hoy
                // ccabceca-3f19-4e37-8479-070a87a33843-41ac9b69-2e3e-4991-aed4-34ba12980d71

                HttpClient client1 = new HttpClient();

                HttpRequestMessage htp = new HttpRequestMessage();
                htp.Method = HttpMethod.Post;

                TipoCambioRequest tipoCambioRequest = new TipoCambioRequest();
                tipoCambioRequest.token = "ccabceca-3f19-4e37-8479-070a87a33843-41ac9b69-2e3e-4991-aed4-34ba12980d71";

                tipoCambioRequest.tipo_cambio = new Tipo_cambioReDet()
                {
                    moneda = "PEN",
                    fecha_inicio = DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"),
                    fecha_fin = DateTime.Now.AddDays(+1).ToString("dd/MM/yyyy"),
                };

                HttpContent ctc = new StringContent(JsonConvert.SerializeObject(tipoCambioRequest), System.Text.Encoding.UTF8, "application/json");

                var response2 = await client1.PostAsync("https://ruc.com.pe/api/v1/consultas", ctc);

                if (response2.IsSuccessStatusCode)
                {
                    var respons = response2.Content.ReadAsStringAsync();
                    TipoCambioResponse tipoCambios = JsonConvert.DeserializeObject<TipoCambioResponse>(respons.Result);
                    int i = tipoCambios.exchange_rates.FindIndex(x => x.fecha == DateTime.Now.ToString("dd/MM/yyyy"));
                    if (i != -1)
                    {
                        SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        Log.WriteToFile($"Actualizando Tipo de Cambio {sbo.SAP_BASE} - " + date + ".....");
                        bo.SetCurrencyRate("USD", System.DateTime.Now, tipoCambios.exchange_rates[i].venta, true);
                        Log.WriteToFile($"Tipo de Cambio del dia {sbo.SAP_BASE} - " + date + " : " + tipoCambios.exchange_rates[i].venta);
                    }
                    else
                    {
                        throw new Exception("No se encuentra con el tipo de cambio del día de hoy");
                    }
                }
                else
                {
                    throw new Exception($"Error al llamar al endpoint https://ruc.com.pe/api/v1/consultas");
                }
            }
            catch (Exception ex)
            {
                Log.WriteToFile($"Metodo CONSULTA API - Error al Actualizar en SAP {sbo.SAP_BASE} -:" + ex.Message);
                IntegrarTipodeCambioSUNAT(sbo).Wait();
            }
        }
        private static async Task IntegrarTipoCambioSBS(SBO sbo)
        {
            try
            {
                // Que el sistema detecté el tipo de Cambio según el formato
                CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
                culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
                culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
                System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;

                string date = System.DateTime.Now.ToString("dd/MM/yyyy");

                List<string> TipoCambio = new List<string>();
                var handler = new HttpClientHandler();
                var client = new HttpClient(handler);
                var rq = new HttpRequestMessage();

                var cookies = new CookieContainer();
                handler.CookieContainer = cookies;

                rq.RequestUri = new Uri("https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx");
                rq.Method = HttpMethod.Get;
                // rq.Headers.Add("Content-Type", "application/json");
                //rq.Headers.Add("Upgrade-Insecure-Requests", "1");
                //rq.Headers.Add("Sec-Fetch-Dest", "document");
                //rq.Headers.Add("Sec-Fetch-Mode", "navigate");
                //rq.Headers.Add("Sec-Fetch-Site", "none");
                //rq.Headers.Add("Sec-Fetch-User", "?1");
                //rq.Headers.Add("Priority", "u=1");
                //rq.Headers.Add("TE", "trailers");
                //rq.Headers.Add("Host", "www.sbs.gob.pe");
                //rq.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8");
                //rq.Headers.Add("Accept-Language", "es-ES,es;q=0.5");
                //rq.Headers.Add("Accept-Encoding", "gzip, deflate, br, zstd");
                //rq.Headers.Add("Connection", "keep-alive");
                //rq.Headers.Add("DNT", "1");
                //rq.Headers.Add("Sec-GPC", "1");
                //rq.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0");
                rq.Headers.Add("Upgrade-Insecure-Requests", "1");
                rq.Headers.Add("Sec-Ch-Ua", "\"Google Chrome\";v=\"125\", \"Chromium\";v=\"125\", \"Not.A/Brand\";v=\"24\"");
                rq.Headers.Add("Sec-Ch-Ua-Mobile", "?0");
                rq.Headers.Add("Sec-Ch-Ua-Platform", "?0");
                rq.Headers.Add("Sec-Fetch-Dest", "document");
                rq.Headers.Add("Sec-Fetch-Mode", "navigate");
                rq.Headers.Add("Sec-Fetch-Site", "none");
                rq.Headers.Add("Sec-Fetch-User", "?1");
                rq.Headers.Add("Priority", "u=0, i");
                rq.Headers.Add("Referer", "https://www.google.com/");
                rq.Headers.Add("TE", "trailers");
                rq.Headers.Add("Host", "www.sbs.gob.pe");
                rq.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7");
                rq.Headers.Add("Accept-Language", "es-ES,es;q=0.5");
                rq.Headers.Add("Accept-Encoding", "gzip, deflate, br, zstd");
                rq.Headers.Add("Connection", "keep-alive");
                rq.Headers.Add("DNT", "1");
                rq.Headers.Add("Sec-GPC", "1");
                rq.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0");
                //var responseCookies = new List<System.Net.Cookie>();

                var handler1 = new HttpClientHandler();
                var client1 = new HttpClient(handler1);
                var cookies1 = new CookieContainer();

                //foreach (var cookie in responseCookies)
                //{
                //    cookies1.Add(cookie);
                //}

                cookies1.Add(ParseCookie("dtCookie=v_4_srv_1_sn_11E9A519CE524C01BC3652B7184423BB_perc_100000_ol_0_mul_1_app-3Aa7babc1dd8d57c64_0; Path=/; Domain=sbs.gob.pe;", rq.RequestUri));

                handler.CookieContainer = cookies1;

                var response = await client.GetAsync("https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx");
               // responseCookies.Add(ParseCookie(value, rq.RequestUri));

                /*
                var responseCookies = new List<System.Net.Cookie>();
                foreach (var header in response.Headers)
                {
                    if (header.Key.ToLower() == "set-cookie")
                    {
                        var t = header;
                        foreach (var value in header.Value)
                        {
                            responseCookies.Add(ParseCookie(value, rq.RequestUri));
                        }
                    }
                }

                var handler1 = new HttpClientHandler();
                var client1 = new HttpClient(handler1);
                var cookies1 = new CookieContainer();

                foreach (var cookie in responseCookies)
                {
                    cookies1.Add(cookie);
                }

                handler1.CookieContainer = cookies1;

                var segundaSolicitud = await client1.GetAsync("https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx");
                */
                var contenido = await response.Content.ReadAsStringAsync();
                
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(contenido);

                HtmlNode Body = doc.DocumentNode.CssSelect("body").First();
                string sbody = Body.InnerHtml;

                foreach (var Node in doc.DocumentNode.CssSelect(".APLI_fila2"))
                {
                    TipoCambio.Add(Node.InnerHtml);
                }
                if (string.IsNullOrEmpty(TipoCambio[1]))
                    throw new Exception("Aun no se ha actualizado el tipo de Cambio SBS");
                double tipoCambio = Convert.ToDouble(TipoCambio[1]);
                //double tipoCambio = Convert.ToDouble(TipoCambio[1]);

                SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                Log.WriteToFile("Actualizando Tipo de Cambio " + date + ".....");
                bo.SetCurrencyRate("USD", System.DateTime.Now, tipoCambio, true);
                Log.WriteToFile("Tipo de Cambio del dia " + date + " : " + tipoCambio);

            }
            catch (Exception ex)
            {
                Log.WriteToFile($"Metodo SBS - Error al Actualizar en SAP {sbo.SAP_BASE} :" + ex.Message);
                IntegrarTipoDeCambioAPI(sbo).Wait();
            }
        }
        static System.Net.Cookie ParseCookie(string cookieString, Uri defaultUri)
        {
            // Aquí puedes implementar tu lógica de análisis de cookies
            var parts = cookieString.Split(';');
            var cookiePart = parts[0];
            var keyValue = cookiePart.Split('=');

            var cookie = new System.Net.Cookie(keyValue[0].Trim(), keyValue[1].Trim());

            // Establecer el dominio predeterminado si no se especifica en la cookie
            cookie.Domain = defaultUri.Host;

            // Analizar los otros atributos de la cookie si están presentes
            foreach (var part in parts.Skip(1))
            {
                var attribute = part.Trim().Split('=');
                if (attribute.Length == 2)
                {
                    var attributeName = attribute[0].Trim();
                    var attributeValue = attribute[1].Trim();
                    switch (attributeName.ToLower())
                    {
                        case "domain":
                            cookie.Domain = attributeValue;
                            break;
                        case "path":
                            cookie.Path = attributeValue;
                            break;
                        // Otros atributos de la cookie, como Secure, HttpOnly, etc., pueden ser manejados aquí
                        // case "secure":
                        //     cookie.Secure = true;
                        //     break;
                        // case "httponly":
                        //     cookie.HttpOnly = true;
                        //     break;
                        default:
                            // Otros atributos no manejados
                            break;
                    }
                }
            }

            return cookie;
        }
        private static async Task integrarTipoCambioADI(SBO sbo)
        {
            try
            {
                // Que el sistema detecté el tipo de Cambio según el formato
                CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
                culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
                culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
                System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;

                string date = System.DateTime.Now.ToString("dd/MM/yyyy");

                List<string> TipoCambio = new List<string>();

                HtmlWeb oWeb = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = oWeb.Load("https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx#");


                HtmlNode Body = doc.DocumentNode.CssSelect("body").First();
                string sbody = Body.InnerHtml;

                foreach (var Node in doc.DocumentNode.CssSelect(".APLI_fila2"))
                {
                    TipoCambio.Add(Node.InnerHtml);
                }
                if (string.IsNullOrEmpty(TipoCambio[1]))
                    throw new Exception("Aun no se ha actualizado el tipo de Cambio SBS");
                double tipoCambio = Convert.ToDouble(TipoCambio[1]);
                //double tipoCambio = Convert.ToDouble(TipoCambio[1]);

                SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                Log.WriteToFile("Actualizando Tipo de Cambio " + date + ".....");
                bo.SetCurrencyRate("USD", System.DateTime.Now, tipoCambio, true);
                Log.WriteToFile("Tipo de Cambio del dia " + date + " : " + tipoCambio);

            }
            catch (Exception ex)
            {
                Log.WriteToFile($"Metodo SBS - Error al Actualizar en SAP {sbo.SAP_BASE} :" + ex.Message);
            }
        }

        public class TipoCambioRequest
        {
            public string token { get; set; }
            public Tipo_cambioReDet tipo_cambio { get; set; }
        }

        public class Tipo_cambioReDet
        {
            public string moneda { get; set; }
            public string fecha_inicio { get; set; }
            public string fecha_fin { get; set; }
        }

        public class TipoCambioResponse
        {
            public bool success { get; set; }
            public List<ResponseCambioDet> exchange_rates { get; set; }
        }

        public class ResponseCambioDet
        {
            public string fecha { get; set; }
            public string moneda { get; set; }
            public double compra { get; set; }
            public double venta { get; set; }
        }
    }
}