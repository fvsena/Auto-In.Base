using Castle.Core.Logging;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestStack.White.Configuration;
using TP_Automation.Logger.Interface;
using TP_Automation.Logger.Model;
using TP_Automation.Logger.Result;
using TPAuto_In.Base.Configuration;
using TPAuto_In.Base.View;
using TPAuto_In.Base.Workflow;

namespace TPAuto_In.Base
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static bool Ended = false;
        static string mensagem = "";
        static bool sucesso = false;

        static void Main(string[] args)
        {
            var log = new Log();
            LogData logData = new LogData();
            logData.LogsProcessData = new List<LogProcessData>();
            string mensagem = "";

            try
            {
                var automation = new Manager().GetAutomation(Global.AutomationId, Global.IsProd);
                if (!automation.Automation.Active)
                {
                    if (!string.IsNullOrEmpty(Global.Homologadores))
                    {
                        List<string> listaHomologadores = Global.Homologadores.Split(';').ToList();
                        if (!listaHomologadores.Any(x => x.ToUpper().Equals(Environment.UserName.ToUpper())))
                        {
                            //EXISTE LISTA DE HOMOLOGADORES MAS USUÁRIO NÃO ESTÁ CADASTRADO
                            return;
                        }
                    }
                    else
                    {
                        //NÃO HÁ HOMOLOGADORES CADASTRADOS
                        return;
                    }
                }

                //CARREGA OS PARAMETROS
                Global.ProcessarParametros(automation);

                //OBTÉM O USUÁRIO EM EXECUÇÃO
                var user = ObterUsuarioExecucao(Global.IsProd, Global.OverrideUser);

                //VALIDA SE O USER ESTÁ EM CÉLULA QUE EXECUTA O AUTOIN
                if (!CelulaExecutaAutoIn(ref user, ref automation))
                {
                    Ended = true;
                    return;
                }

                //EXIBE O ICONE DO AUTOIN
                new Thread(() =>
                {
                    ExibeForm(Global.AutomationId, Global.IsProd, Global.IsDebug, Global.LiberaEncerramentoManual, Global.TempoEncerramento, ref logData, ref log);

                }).Start();

                ////////////////////////////////////////////////////////////////
                //INICIO DO LOG
                logData.StartDate = DateTime.Now;
                logData.Hostname = Environment.MachineName;
                logData.Username = Environment.UserName;
                logData.AutomationID = automation.Automation.Id;
                logData.ClientID = Global.ClientId;

                var handle = GetConsoleWindow();
                CoreAppXmlConfiguration.Instance.LoggerFactory = new ConsoleFactory(LoggerLevel.Error);

                //Console Handler
                if (Global.ExibeConsole)
                {
                    ShowWindow(handle, SW_SHOW);
                }
                else
                {
                    ShowWindow(handle, SW_HIDE);
                }

                var workflow = new AutoInWorkFlow();
                workflow.AbreChrome(newWindow: true);

                sucesso = workflow.IniciarLoginSistemas(ref logData, ref automation, out mensagem);

                if (!sucesso)
                {
                    logData.Success = false;
                    logData.Message = mensagem;
                }
                else
                {
                    logData.Success = true;
                    logData.Message = "Processo OK";
                }

                //Play warning sound at the end
                Thread.Sleep(1000);
                System.Media.SystemSounds.Hand.Play();

                //form.Close();
                Ended = true;
                logData.EndDate = DateTime.Now;
                var result = log.SaveLog(logData, Global.IsProd);
            }
            catch (Exception ex)
            {
                Ended = true;
                logData.Success = false;
                logData.Message = ex.Message;
                logData.EndDate = DateTime.Now;
                var result = log.SaveLog(logData, Global.IsProd, Global.IsDebug);
                Console.WriteLine("Resultado log:" + result.MsgCatch);

                System.Media.SystemSounds.Beep.Play();
            }
        }

        private static UserResult ObterUsuarioExecucao(bool isProd, string overrideUser)
        {
            //OBTÉM DADOS DO USUÁRIO EM EXECUÇÃO
            UserResult user;
            if (!string.IsNullOrEmpty(overrideUser))
            {
                user = new Manager().GetUser("TPB\\" + overrideUser, isProd);
            }
            else
            {
                user = new Manager().GetUser("TPB\\" + Environment.UserName, isProd);
            }

            return user;
        }

        private static bool CelulaExecutaAutoIn(ref UserResult user, ref AutomationResult automation)
        {
            if (user == null || (user != null && user.User == null))
            {
                Console.WriteLine("Usuário não localizado no MOP");
                Thread.Sleep(5000);
                return false;
            }

            bool celulaExecutaAutoIn = false;

            var listaCelulas = Global.Celulas.Split(';').ToList();

            foreach (var celulaUsuario in user.User.EmployeeIdentifications)
            {
                if (listaCelulas.Any(x => x.Equals(celulaUsuario.CellCode)))
                {
                    celulaExecutaAutoIn = true;
                    break;
                }
            }
            return celulaExecutaAutoIn;
        }

        private static void ExibeForm(int automationId, bool isProd, bool isDebug, bool liberaEncerramentoManual, int tempoEncerramento, ref LogData logData, ref Log log)
        {
            try
            {
                FrmAutoIn form = new FrmAutoIn(logData.StartDate);
                Rectangle workingArea = Screen.GetWorkingArea(form);
                form.Location = new Point(workingArea.Right - form.Bounds.Width, workingArea.Bottom - form.Bounds.Height);
                form.Show();
                bool verificarLiberacaoEncerramento = true;
                Application.DoEvents();

                DateTime dtVerificacao = DateTime.Now;
                while (!Ended)
                {

                    if (liberaEncerramentoManual && verificarLiberacaoEncerramento && DateTime.Now.Subtract(dtVerificacao).TotalSeconds > tempoEncerramento)
                    {
                        verificarLiberacaoEncerramento = false;
                        form.LiberaEncerramento();
                    }

                    if (DateTime.Now.Subtract(dtVerificacao).TotalSeconds > 120)
                    {
                        dtVerificacao = DateTime.Now;

                        var automationAtualizado = new Manager().GetAutomation(automationId, isProd);
                        if (automationAtualizado.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Force Shutdown")).Equals("1"))
                        {
                            logData.Success = false;
                            logData.Message = "Automação encerrada remotamente via Force Shutdown";
                            logData.EndDate = DateTime.Now;
                            var result = log.SaveLog(logData, isProd, isDebug);
                            Console.WriteLine("Resultado log:" + result.MsgCatch);

                            Application.Exit();
                            Environment.Exit(0);
                        }
                    }
                    Thread.Sleep(1000);
                    Application.DoEvents();
                }

                if (!sucesso)
                {
                    form.TrayMessage(mensagem + " | Realize o login manualmente.");
                }
                else
                {
                    form.TrayMessage("Processo de login finalizado. Bom trabalho!");
                }
            }
            catch (Exception ex)
            {
                logData.Success = false;
                logData.Message = "Falha no processo do Frm Auto IN: " + ex.Message;
                logData.EndDate = DateTime.Now;
                var result = log.SaveLog(logData, isProd, isDebug);
                Console.WriteLine("Resultado log:" + result.MsgCatch);
            }

            Application.Exit();
            Environment.Exit(0);
        }
    }
}
