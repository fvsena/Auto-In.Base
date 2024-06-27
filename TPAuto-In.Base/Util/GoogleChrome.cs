using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Net;
using System.Threading;

namespace TPAuto_In.Base.Util
{
    public class GoogleChrome
    {
        public IWebDriver Driver = null;
        public string ChromeDriverProcessID = string.Empty;
        public static int ChromeID;
        public string DownloadPath;

        /// <summary>
        /// Cria uma instncia da classe de controle do Driver
        /// </summary>
        /// <param name="showBrowser">Exibe ou no a janela do navegador</param>
        /// <param name="binaryLocation">Configura um caminho customizado do binrio do Chrome</param>
        /// <param name="downloadPath">Altera o caminho padro de downloads</param>
        /// <param name="implicitTimeout">Configura um timeout predefinido para captura de elementos</param>
        public GoogleChrome(bool showBrowser = true, string binaryLocation = null, string downloadPath = null,
                            int? implicitTimeout = null, bool foraDaTela = false, string userProfile = null, string driverPath = null)
        {
            var chromeDriverService = ChromeDriverService.CreateDefaultService(driverPath);
            chromeDriverService.HideCommandPromptWindow = true;


            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
            //chromeOptions.AddUserProfilePreference("print.always_print_silent", true);

            //Janela oculta
            if (!showBrowser)
            {
                chromeOptions.AddArguments("headless");
            }

            if (foraDaTela)
            {
                //options.AddArgument("--window-position=-32000,-32000");
                chromeOptions.AddArgument("--window-position=-32000,-32000");
            }

            if (!string.IsNullOrEmpty(userProfile))
            {
                chromeOptions.AddArgument("--user-data-dir=" + userProfile);
                //chromeOptions.AddArgument("--profile-directory=Default");
            }

            //Configura local do exe do Chrome manualmente
            if (!string.IsNullOrEmpty(binaryLocation))
            {
                chromeOptions.BinaryLocation = binaryLocation;
            }

            //Configura local padro de download
            if (!string.IsNullOrEmpty(downloadPath))
            {
                chromeOptions.AddUserProfilePreference("download.default_directory", downloadPath);
                DownloadPath = downloadPath;

            }
            chromeOptions.AddArguments("--disable-notifications");
            chromeOptions.AddArguments("--start-maximized");
            chromeOptions.AddUserProfilePreference("credentials_enable_service", false);
            chromeOptions.AddUserProfilePreference("profile.password_manager_enabled", false);
            //chromeOptions.AddAdditionalCapability("useAutomationExtension", false);

            //Inicia o Driver sem usar os Options
            //driver = new ChromeDriver(chromeDriverService);

            //Inicia o Driver usando os Options do Chrome
            Driver = new ChromeDriver(chromeDriverService, chromeOptions);

            ChromeDriverProcessID = chromeDriverService.ProcessId.ToString();

            if (implicitTimeout.HasValue)
            {
                Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(implicitTimeout.Value);
                Driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromMilliseconds(implicitTimeout.Value));
            }

        }

        public GoogleChrome(string debugPort, int tentativas = 3, int delay = 3, bool encerraTodosDrivers = false, bool hideConsole = true)
        {
            if (encerraTodosDrivers)
            {
                EncerraTodosChromeDriver();
            }

            Thread.Sleep(3000);

            for (int i = 0; i < tentativas; i++)
            {
                try
                {
                    ChromeOptions options = new ChromeOptions();
                    options.DebuggerAddress = debugPort;

                    var chromeDriverService = ChromeDriverService.CreateDefaultService();
                    chromeDriverService.HideCommandPromptWindow = hideConsole;
                    Driver = new ChromeDriver(chromeDriverService, options);
                    ChromeDriverProcessID = chromeDriverService.ProcessId.ToString();
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(delay * 1000);
                }
            }
        }

        public string GetDriverSource()
        {
            try
            {
                if (Driver == null)
                {
                    return "";
                }

                return Driver.PageSource;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void EncerraTodosChromeDriver()
        {
            System.Diagnostics.Process[] chromeDriverProcesses = System.Diagnostics.Process.GetProcessesByName("chromedriver");
            foreach (var chromeDriverProcess in chromeDriverProcesses)
            {
                chromeDriverProcess.Kill();
            }
        }

        /// <summary>
        /// Encerra o processo do Chromedriver e todas as janelas do Chrome que este iniciou
        /// </summary>
        /// <param name="chromeDriverProcessID">ID do processo gerado pelo Chromedriver</param>
        public void EncerraDriver()
        {
            try
            {
                System.Diagnostics.Process[] chromeDriverProcesses = System.Diagnostics.Process.GetProcessesByName("chromedriver");
                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    if (chromeDriverProcess.Id.ToString().Equals(ChromeDriverProcessID))
                    {
                        chromeDriverProcess.Kill();
                        break;
                    }
                }

                EncerrarNavegador(ChromeDriverProcessID);
            }
            catch
            {

            }
        }

        /// <summary>
        /// Faz a navegação para um link determinado
        /// </summary>
        /// <param name="url"></param>
        public void Navegar(string url)
        {
            Driver.Navigate().GoToUrl(url);
        }

        /// <summary>
        /// Captura a tela atual e salva em um arquivo png
        /// </summary>
        /// <param name="caminho">Caminho do arquivo que será gravado</param>
        public void PrintScreen(string caminho)
        {
            Screenshot ss = ((ITakesScreenshot)Driver).GetScreenshot();
            string screenshot = ss.AsBase64EncodedString;
            byte[] screenshotAsByteArray = ss.AsByteArray;
            ss.SaveAsFile(caminho, ScreenshotImageFormat.Png);
        }

        /// <summary>
        /// Executa uma funo javascript
        /// </summary>
        /// <param name="script">Nome da funo</param>
        public void ExecutarScript(string script)
        {
            ((IJavaScriptExecutor)Driver).ExecuteScript(script);
            Thread.Sleep(50);
        }

        /// <summary>
        /// Executa uma funo javascript
        /// </summary>
        /// <param name="script">Nome da funo</param>
        /// <param name="elemento">Elemento que ser usado de parametro da funo javascript</param>
        public void ExecutarScript(string script, IWebElement elemento)
        {
            ((IJavaScriptExecutor)Driver).ExecuteScript(script, elemento);
            Thread.Sleep(50);
        }

        /// <summary>
        /// Localiza um elemento utilizando um critrio de busca
        /// </summary>
        /// <param name="by">Critrio de busca do elemento</param>
        /// <param name="tentativas">Quantidade de tentativas de localizar o elemento</param>
        /// <param name="delay">Tempo de espera entre cada tentativa de localizar o elemento</param>
        /// <param name="lancaExcecao">Permite ou no lanar exceo caso o elemento no seja localizado aps todas as tentativas</param>
        /// <returns>Elemento IWebElement existente na pgina</returns>
        public IWebElement LocalizaElemento(By by, int tentativas, int delay, bool lancaExcecao)
        {
            IWebElement elemento = null;
            for (int i = 0; i < tentativas; i++)
            {
                try
                {
                    elemento = Driver.FindElement(by);
                }
                catch (Exception ex)
                {
                    elemento = null;
                    Thread.Sleep(delay * 1000);
                }

                if (elemento != null)
                {
                    break;
                }
            }

            if (elemento == null && lancaExcecao)
            {
                throw new Exception($"Elemento no localizado");
            }

            return elemento;
        }

        /// <summary>
        /// Localiza um elemento utilizando um critrio de busca
        /// </summary>
        /// <param name="by">Critrio de busca do elemento</param>
        /// <param name="tentativas">Quantidade de tentativas de localizar o elemento</param>
        /// <param name="delay">Tempo de espera entre cada tentativa de localizar o elemento</param>
        /// <param name="lancaExcecao">Permite ou no lanar exceo caso o elemento no seja localizado aps todas as tentativas</param>
        /// <returns>Elemento IWebElement existente na pgina</returns>
        public IList<IWebElement> LocalizaElementos(By by, int tentativas, int delay, bool lancaExcecao)
        {
            //Console.WriteLine("Localizando " + by.Criteria);
            IList<IWebElement> elementos = null;
            for (int i = 0; i < tentativas; i++)
            {
                try
                {
                    elementos = Driver.FindElements(by);
                }
                catch (Exception ex)
                {
                    elementos = null;
                    Thread.Sleep(delay * 1000);
                }

                if (elementos != null && elementos.Count > 0)
                {
                    break;
                }
                else
                {
                    Thread.Sleep(delay * 1000);
                }
            }

            if ((elementos == null || elementos.Count == 0) && lancaExcecao)
            {
                throw new Exception($"Elemento no localizado");
            }

            return elementos;
        }

        public IWebElement LocalizaElementoPropriedade(IList<IWebElement> elementos, string propriedade, string valor, bool lancaExcecao)
        {
            foreach (var elemento in elementos)
            {
                var valorPropriedade = elemento.GetAttribute(propriedade);
                if (!string.IsNullOrEmpty(valorPropriedade) && valorPropriedade == valor)
                {
                    return elemento;
                }
            }
            if (lancaExcecao)
            {
                throw new Exception($"Não foi possível localizar um elemento com a propriedade [{propriedade}] de valor [{valor}]");
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Encerra todas as janelas do Chrome de acordo com o ID do processo pai
        /// </summary>
        /// <param name="chromeDriverProcessID">ID do processo pai</param>
        public void EncerrarNavegador(string chromeDriverProcessID)
        {

            System.Diagnostics.Process[] chromeProcess = System.Diagnostics.Process.GetProcessesByName("chrome");
            foreach (var process in chromeProcess)
            {
                int parentPid = 0;
                try
                {
                    using (ManagementObject mo = new ManagementObject($"win32_process.handle='{process.Id}'"))
                    {
                        mo.Get();
                        parentPid = Convert.ToInt32(mo["ParentProcessId"]);

                        if (parentPid.ToString() == chromeDriverProcessID)
                        {
                            process.Kill();
                        }
                    }
                }
                catch
                {

                }
            }
        }

        /// <summary>
        /// Encerra todas as janelas do Chrome de acordo com o ID do processo pai
        /// </summary>
        /// <param name="chromeDriverProcessID">ID do processo pai</param>
        public static void EncerrarNavegadores()
        {

            System.Diagnostics.Process[] chromeProcess = System.Diagnostics.Process.GetProcessesByName("chrome");
            foreach (var process in chromeProcess)
            {
                try
                {
                    process.Kill();
                }
                catch
                {

                }
            }
        }

        /// <summary>
        /// Encerra todas as janelas do Chrome de acordo com o ID do processo pai
        /// </summary>
        /// <param name="chromeDriverProcessID">ID do processo pai</param>
        public void EncerrarNavegador(int processId)
        {

            var chromeProcess = System.Diagnostics.Process.GetProcessById(processId);
            try
            {
                chromeProcess.Kill();
            }
            catch
            {

            }
        }

        /// <summary>
        /// Aguarda o download de um determinado arquivo
        /// </summary>
        /// <param name="nomeArquivo">Nome do arquivo</param>
        /// <param name="caminhoDownload">Diretório onde será gravado</param>
        public void AguardarDownload(string nomeArquivo, int maximoTentativas = 3000)
        {
            try
            {
                for (var i = 0; i < maximoTentativas; i++)
                {
                    if (File.Exists(DownloadPath + nomeArquivo)) { break; }
                    Thread.Sleep(1000);
                }

                if (!File.Exists(DownloadPath + nomeArquivo))
                {
                    throw new Exception("Não foi possível localizar o arquivo baixado");
                }

                Thread.Sleep(2000);
                var length = new FileInfo(DownloadPath + nomeArquivo).Length;
                for (var i = 0; i < 30; i++)
                {
                    Thread.Sleep(500);
                    var newLength = new FileInfo(DownloadPath + nomeArquivo).Length;
                    if (newLength == length && length != 0) { break; }
                    length = newLength;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public CookieContainer CapturarCookies()
        {
            CookieContainer cookieContainer = new CookieContainer();

            CookieCollection cc = new CookieCollection();
            foreach (OpenQA.Selenium.Cookie cook in Driver.Manage().Cookies.AllCookies)
            {
                System.Net.Cookie cookie = new System.Net.Cookie();
                cookie.Name = cook.Name;
                cookie.Value = cook.Value;
                cookie.Domain = cook.Domain;
                cookieContainer.Add(cookie);
            }

            return cookieContainer;
        }

        public IntPtr ObterHandle(string chromeDriverProcessID)
        {
            IntPtr intptr = new IntPtr();
            System.Diagnostics.Process[] chromeProcess = System.Diagnostics.Process.GetProcessesByName("chrome");
            foreach (var process in chromeProcess)
            {
                int parentPid = 0;
                try
                {
                    using (ManagementObject mo = new ManagementObject($"win32_process.handle='{process.Id}'"))
                    {
                        mo.Get();
                        parentPid = Convert.ToInt32(mo["ParentProcessId"]);

                        if (parentPid.ToString() == chromeDriverProcessID)
                        {
                            intptr = process.MainWindowHandle;
                            break;
                        }
                    }
                }
                catch
                {

                }
            }
            return intptr;
        }

        //TO DO: IMPLEMENTAR MÉTODO DE NAVEGAÇÃO ENTRE OS HANDLES
        private bool SearchString(string[] search, string url)

        {

            var match = Array.FindAll(search, n => url.Contains(n));

            return (search.Length == match.Length);

        }

    }
}
