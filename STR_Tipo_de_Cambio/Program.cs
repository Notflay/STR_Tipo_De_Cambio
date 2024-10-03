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
using System.Configuration;

namespace STR_Tipo_de_Cambio
{
    public class Program
    {
        static async Task Main(string[] args)
        {
            // Obtener la ruta del archivo de conexión
            string filePath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "conexion.xml");

            try
            {
                CultureInfo culturaPersonalizada = new CultureInfo("es-PE");
                culturaPersonalizada.NumberFormat.NumberDecimalSeparator = ".";
                culturaPersonalizada.NumberFormat.NumberGroupSeparator = ",";
                System.Threading.Thread.CurrentThread.CurrentCulture = culturaPersonalizada;
                // Cargar el archivo XML y obtener la lista de conexiones SBO
                List<SBO> sboList = ObtenerListaDeSBOs(filePath);

                // Obtener el tipo de cambio una sola vez
                double tipoCambio = await ObtenerTipoCambio();

                if (tipoCambio == 0)
                {
                    Console.WriteLine("No se pudo obtener el tipo de cambio.");
                    return;
                }

                // Conectar a cada base de datos SAP y actualizar el tipo de cambio
                foreach (var sbo in sboList)
                {
                    try
                    {
                        Console.WriteLine($"Conectando a ServerSAP: {sbo.SAP_SERVIDOR}, SBOCompany: {sbo.SAP_BASE}...");

                        SAPConnector.Conectar(sbo);

                        // Actualizar el tipo de cambio en SAP
                        ActualizarTipoDeCambioEnSAP(sbo, tipoCambio);

                        SAPConnector.Desconectar();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al procesar SBO {sbo.SAP_BASE}: {ex.Message}");
                        Log.WriteToFile($"Error al procesar SBO {sbo.SAP_BASE}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error general: {ex.Message}");
                Log.WriteToFile($"Error general: {ex.Message}");
            }
        }

        private static List<SBO> ObtenerListaDeSBOs(string filePath)
        {
            List<SBO> sboList = new List<SBO>();

            try
            {
                XDocument xmlDoc = XDocument.Load(filePath);

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
                        }
                    }
                    sboList.Add(sbo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo XML: {ex.Message}");
                throw;
            }

            return sboList;
        }

        private static async Task<double> ObtenerTipoCambio()
        {
            try
            {
                return await ObtenerTipoCambioRamo();
            }
            catch (Exception)
            {
                try
                {
                    // Intentar obtener el tipo de cambio desde SBS
                    return await ObtenerTipoCambioDesdeSBS();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al obtener tipo de cambio desde SBS: {ex.Message}");
                    Log.WriteToFile($"Error al obtener tipo de cambio desde SBS: {ex.Message}");

                    // Si falla, intentar obtenerlo desde la SUNAT
                    try
                    {
                        return await ObtenerTipoCambioDesdeSUNAT();
                    }
                    catch (Exception exSunat)
                    {
                        Console.WriteLine($"Error al obtener tipo de cambio desde SUNAT: {exSunat.Message}");
                        Log.WriteToFile($"Error al obtener tipo de cambio desde SUNAT: {exSunat.Message}");

                        throw;
                        // return 0; // Devuelve 0 si no se pudo obtener el tipo de cambio
                    }
                }
            } 
        }
        private static async Task<double> ObtenerTipoCambioRamo()
        {
            try
            {
                string _link = ConfigurationManager.AppSettings["tpocambio_endpoint"];
                int _ejecucion = Convert.ToInt32(ConfigurationManager.AppSettings["dia_later"]);
                string _uri = $"{_link}tipoCambio?later={_ejecucion}";
                HttpClient client = new HttpClient();
                HttpRequestMessage request = new HttpRequestMessage();
                request.RequestUri = new Uri(_uri);
                request.Method = HttpMethod.Get;

                Log.WriteToFile($"ObtenerTipoCambioRamo - GET - {_uri}");

                var response = await client.SendAsync(request);
            
                if (response.IsSuccessStatusCode)
                {
                    return Convert.ToDouble(response.Content.ReadAsStringAsync().Result);
                }
                else
                {
                    throw new Exception("Error al llamar al endpoint de RAMO.");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private static async Task<double> ObtenerTipoCambioDesdeSBS()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); // Para ejecución en segundo plano
            using (IWebDriver driver = new ChromeDriver(options))
            {
                string url = "https://www.sbs.gob.pe/app/pp/sistip_portal/paginas/publicacion/tipocambiopromedio.aspx";
                int intentos = 0;
                string html = "";

                while (intentos < 3)
                {
                    try
                    {
                        driver.Navigate().GoToUrl(url);
                        html = driver.PageSource;

                        if (!html.Contains("Request unsuccessful"))
                            break;

                        Console.WriteLine("El mensaje 'Request unsuccessful' fue detectado. Refrescando la página...");
                        await Task.Delay(1000); // Esperar antes de reintentar
                    }
                    catch (WebDriverException ex)
                    {
                        Console.WriteLine($"Error al cargar la página: {ex.Message}");
                        await Task.Delay(1000); // Esperar antes de reintentar
                    }

                    intentos++;
                }

                if (intentos == 3)
                {
                    throw new Exception("Se ha alcanzado el número máximo de intentos. La página no se pudo cargar.");
                }

                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                var nodes = doc.DocumentNode.SelectNodes("//td[contains(@class, 'APLI_fila2')]");
                if (nodes == null || nodes.Count < 2 || string.IsNullOrEmpty(nodes[1].InnerText))
                {
                    throw new Exception("Aún no se ha actualizado el tipo de Cambio SBS o el formato de la página ha cambiado.");
                }

                if (double.TryParse(nodes[1].InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out double tipoCambio))
                {
                    return tipoCambio;
                }
                else
                {
                    throw new Exception("No se pudo convertir el tipo de cambio a un valor numérico.");
                }
            }
        }

        private static async Task<double> ObtenerTipoCambioDesdeSUNAT()
        {
            try
            {
                HttpClient client = new HttpClient();
                HttpRequestMessage request = new HttpRequestMessage();
                request.RequestUri = new Uri("https://www.sunat.gob.pe/a/txt/tipoCambio.txt");
                request.Method = HttpMethod.Get;
                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    List<string> valores = content.Split('|').ToList();
                    string fecha = DateTime.Parse(valores[0]).ToShortDateString();
                    if (fecha == DateTime.Now.ToShortDateString())
                    {
                        return Convert.ToDouble(valores[2]);
                    }
                    else
                    {
                        throw new Exception("El tipo de cambio no está actualizado para hoy.");
                    }
                }
                else
                {
                    throw new Exception("Error al llamar al endpoint de SUNAT.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al obtener tipo de cambio desde SUNAT: {ex.Message}");
            }
        }

        private static void ActualizarTipoDeCambioEnSAP(SBO sbo, double tipoCambio)
        {
            try
            {
                string date = System.DateTime.Now.ToString("dd/MM/yyyy");

                SAPbobsCOM.SBObob bo = SAPConnector.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                Log.WriteToFile($"Actualizando Tipo de Cambio {sbo.SAP_BASE} - {date}.....");
                bo.SetCurrencyRate("USD", ConfigurationManager.AppSettings["dia_later"] == "1" ? System.DateTime.Now.AddDays(1) : System.DateTime.Now, tipoCambio, true);
                Log.WriteToFile($"Tipo de Cambio del día {sbo.SAP_BASE} - {date}: {tipoCambio}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al actualizar tipo de cambio en SAP {sbo.SAP_BASE}: {ex.Message}");
                Log.WriteToFile($"Error al actualizar tipo de cambio en SAP {sbo.SAP_BASE}: {ex.Message}");
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