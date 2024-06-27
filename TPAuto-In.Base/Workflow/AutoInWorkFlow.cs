using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestStack.White;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems;
using TPAuto_In.Base.Util;
using TPAuto_In.Base.Configuration;
using TP_Automation.Logger.Result;
using TP_Automation.Logger.Model;
using TestStack.White.Factory;

namespace TPAuto_In.Base.Workflow
{
    public class AutoInWorkFlow
    {
        private GoogleChrome Chrome = null;

        public void AbreChrome(string url = null, bool newWindow = true)
        {
            ////ABRE CHROME VIA PROCESS
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.FileName = "chrome";
            if (string.IsNullOrEmpty(url))
            {
                process.StartInfo.Arguments = $"about:blank {(newWindow ? "--new-window" : "")} --remote-debugging-port=9222";
            }
            else
            {
                process.StartInfo.Arguments = url + $" {(newWindow ? "--new-window" : "")} --remote-debugging-port=9222";
            }

            process.Start();
        }

        private bool ProcessoMarcacaoPonto(ref AutomationResult automation, ref LogData log, out string mensagem)
        {
            mensagem = "";
            if (automation.Automation.AutomationProcesses.FirstOrDefault(x => x.Id.Equals(Global.IdProcessoMarcacaoPonto)).Active)
            {
                //LOG LOOP CONTROLE LEGAL
                var logCL = new LogProcessData();
                logCL.StartDate = DateTime.Now;
                logCL.IdProcessAutomation = Global.IdProcessoMarcacaoPonto;

                bool tpLoginEnabled = false;

                switch (Global.TipoLogin)
                {
                    case "0":
                        tpLoginEnabled = false;
                        break;
                    case "1":
                        tpLoginEnabled = true;
                        break;
                    case "2":
                        tpLoginEnabled = Global.TPLoginEnabled;
                        break;
                }


                //Wait Controle Legal to loadup
                var tempoCL = Convert.ToInt32(automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Tempo Máximo Espera Controle Legal")).Value);
                bool isCLReady = Global.SkipCL;

                if (!tpLoginEnabled)
                {
                    while (isCLReady == false)
                    {
                        isCLReady = IsCLReady(tempoCL, out bool tempoExcedido);
                        if (tempoExcedido)
                        {
                            logCL.EndDate = DateTime.Now;
                            logCL.Success = false;
                            logCL.Message = "Tempo máximo de espera atingido no processo do CL";
                            log.LogsProcessData.Add(logCL);

                            log.Success = false;
                            log.Message = "Tempo máximo de espera atingido no processo do CL";
                            log.EndDate = DateTime.Now;
                            mensagem = "Tempo máximo de espera atingido no processo do CL";
                            return false;
                        }
                    }
                }
                else
                {
                    string overrideUser = Global.OverrideUser;
                    string login = string.IsNullOrEmpty(overrideUser) ? Environment.UserName : overrideUser;
                    var caminhoArquivo = Global.CaminhoIdentificadorTPLogin.Replace("{{USERNAME}}", login);
                    while (isCLReady == false)
                    {
                        isCLReady = IsTPLoginReady(caminhoArquivo, tempoCL, out bool tempoExcedido);
                        if (tempoExcedido)
                        {
                            logCL.EndDate = DateTime.Now;
                            logCL.Success = false;
                            logCL.Message = "Tempo máximo de espera atingido no processo do TP LOGIN";
                            log.LogsProcessData.Add(logCL);

                            log.Success = false;
                            log.Message = "Tempo máximo de espera atingido no processo do CL";
                            log.EndDate = DateTime.Now;
                            mensagem = "Tempo máximo de espera atingido no processo do CL";
                            return false;
                        }
                    }
                }



                logCL.EndDate = DateTime.Now;
                logCL.Success = true;
                logCL.Message = "Processo OK";
                log.LogsProcessData.Add(logCL);
            }

            return true;
        }

        private bool IsCLReady(int tempoMaximo, out bool tempoExcedido)
        {

            tempoExcedido = false;
            try
            {
                Console.WriteLine("Localizando a janela do controle legal");
                List<System.Diagnostics.Process> processes = System.Diagnostics.Process.GetProcesses().Where(p => p.MainWindowTitle == "Controle Legal").ToList();
                DateTime dataInicio = DateTime.Now;
                //AGUARDA O PROCESSO DO CONTROLE LEGAL SER CARREGADO
                while (processes.Count < 1)
                {
                    if (Convert.ToInt32(DateTime.Now.Subtract(dataInicio).TotalSeconds) > tempoMaximo)
                    {
                        tempoExcedido = true;
                        return false;
                    }

                    processes = System.Diagnostics.Process.GetProcesses().Where(p => p.MainWindowTitle == "Controle Legal").ToList();
                    Thread.Sleep(5000);

                }

                //LOCALIZA O BOTÃO IDENTIFICADOR DE LOGIN
                foreach (var process in processes)
                {
                    var id = process.Id;
                    Application application = Application.Attach(id);
                    Window window = application.GetWindow(process.MainWindowTitle, InitializeOption.NoCache);

                    SearchCriteria searchCriteria = SearchCriteria.ByText("Saída");

                    Button button = null;
                    while (button == null)
                    {
                        Console.WriteLine("Localizando o botão sair");
                        Thread.Sleep(1000);
                        button = window.Get<Button>(searchCriteria);
                        while (!button.Visible)
                        {
                            if (Convert.ToInt32(DateTime.Now.Subtract(dataInicio).TotalSeconds) > tempoMaximo)
                            {
                                tempoExcedido = true;
                                return false;
                            }

                            Thread.Sleep(10000);
                        }
                    }

                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool IsTPLoginReady(string nomeArquivo, int tempoMaximo, out bool tempoExcedido)
        {
            tempoExcedido = false;
            try
            {
                Console.WriteLine("Localizando o arquivo que identifica o login no TP Logi//n");
                DateTime dataInicio = DateTime.Now;

                bool fileExists = File.Exists(nomeArquivo);
                while (!fileExists && !tempoExcedido)
                {
                    Thread.Sleep(2000);

                    fileExists = File.Exists(nomeArquivo);
                    tempoExcedido = Convert.ToInt32(DateTime.Now.Subtract(dataInicio).TotalSeconds) > tempoMaximo;
                }

                if (!fileExists)
                {
                    return fileExists;
                }
                else
                {
                    new FileInfo(nomeArquivo).Delete();
                    return fileExists;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exceção: " + ex.Message);
                return false;
            }
        }

        public bool IniciarLoginSistemas(ref LogData logData, ref AutomationResult automation, out string mensagem)
        {
            mensagem = "";
            bool sucesso = false;

            try
            {
                //AGUARDA MARCACAO DE PONTO
                if (!ProcessoMarcacaoPonto(ref automation, ref logData, out mensagem))
                {
                    return false;
                }


                //TO DO: IMPLEMENTAR CÓDIGO DO PROCESSO DO AUTO-IN

                sucesso = true;
            }
            catch (Exception ex)
            {
                sucesso = false;
                mensagem = ex.Message;
            }

            return sucesso;
        }
    }
}
