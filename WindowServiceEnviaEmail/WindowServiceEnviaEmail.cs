using System;
using System.Configuration;
using System.ServiceProcess;
using System.Timers;
using ServicoEmail;
using Modelos;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace WindowServiceEnviaEmail
{
    public partial class WindowServiceEnviaEmail : ServiceBase
    {
        Timer timer1 = new Timer();
        int getCallType = 0;
        List<ItensListaModel> parametro = null;

        ClientContext contextoSharepoint;
        public WindowServiceEnviaEmail()
        {
            InitializeComponent();
            ErroModel registroErro = new ErroModel();

            string site = "https://classsolutions.sharepoint.com/sites/leandro.silva";
            string usuario = "leandro.silva@class-solutions.com.br";
            string senha = "Class685947";

            contextoSharepoint = DataSharepoint.AutenticaAcessoSPOnline(site, usuario, senha, out registroErro);
            parametro = DataSharepoint.obtemLista(contextoSharepoint, out registroErro, "Parametro");

            int strTime = Convert.ToInt32(ConfigurationManager.AppSettings["callDuration"]);
            getCallType = Convert.ToInt32(ConfigurationManager.AppSettings["CallType"]);
            if (getCallType == 1)
            {
                double inter = (double)GetNextInterval(parametro);
                timer1.Interval = inter;
                timer1.Elapsed += new ElapsedEventHandler(ServiceTimer_Tick);
            }
            else
            {
                timer1 = new Timer();
                timer1.Interval = strTime * 1000;
                timer1.Elapsed += new ElapsedEventHandler(ServiceTimer_Tick);
            }
        }

        private void ServiceTimer_Tick(object sender, ElapsedEventArgs e)
        {
            ErroModel registroErro = new ErroModel();
            var itens = DataSharepoint.obtemLista(contextoSharepoint, out registroErro, "Teste");
            string retornoLista = "<h1>Itens Recuperados</h1>";
            foreach (var item in itens)
            {
                retornoLista += $"<p>Item : {item.Title} - ID da pessoa cadastrada: {item.ID_Pessoa} - Nome Da pessoa: {item.Nome_Pessoa}</p>";
            }
            retornoLista += registroErro.Mensagem;
            string exeption = "";
            bool retornoEnvioEmail = Email.EnviaEmail("Leandrots82@gmail.com",
                "Teste de envio de e-mail",
                $"<h1>E-mail diário automatico.</h1><p>E-mail disparado por serviço windows services.</p><p>Este e-mail foi disparado a partir da dll de serviço de disparo de e-mail (deu certo)</p><br/><br/>{retornoLista}",
                exeption);

            if (!retornoEnvioEmail)
            {
                retornoEnvioEmail = ServicoEmail.EnviaEmail("Leandrots82@gmail.com",
                "Email diário",
                $"<h1>E-mail diário automatico.</h1><p>E-mail disparado por serviço windows services.</p><p>Este e-mail foi disparado a partir do método de disparo de e-mail implementado dentro do serviço, a dll não funcionou.</p><p>{exeption}</p><br/><br/>{retornoLista}",
                out exeption);
            }
            if (getCallType == 1)
            {
                timer1.Stop();
                System.Threading.Thread.Sleep(1000000);
                SetTimer();
            }
        }

        private double GetNextInterval(List<ItensListaModel> itens)
        {
            var timeString = ConfigurationManager.AppSettings["StartTime"];
            var t0 = itens[0].Title;
            DateTime t = DateTime.Parse(timeString);
            TimeSpan ts = new TimeSpan();
            int x;
            ts = t - System.DateTime.Now;
            if (ts.TotalMilliseconds < 0)
            {
                ts = t.AddDays(1) - System.DateTime.Now;//Here you can increase the timer interval based on your requirments.   
            }
            return ts.TotalMilliseconds;
        }
        private void SetTimer()
        {
            try
            {
                double inter = (double)GetNextInterval(parametro);
                timer1.Interval = inter;
                timer1.Start();
            }
            catch (Exception ex)
            {
            }
        }

        private void IniciaServico()
        {

        }

        protected override void OnStart(string[] args)
        {
            timer1.AutoReset = true;
            timer1.Enabled = true;
            ServiceLog.WriteErrorLog("Daily Reporting service started");
        }

        protected override void OnStop()
        {
            timer1.AutoReset = false;
            timer1.Enabled = false;
            ServiceLog.WriteErrorLog("Daily Reporting service stopped");
        }
    }
}
