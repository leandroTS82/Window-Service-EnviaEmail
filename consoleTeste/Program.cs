using System;

namespace consoleTeste
{
    public class Program
    {
        static void Main(string[] args)
        {
            string exeption = "";
            bool retornoEnvioEmail = Mail.EnviaEmail("Gracekellluz@Gmail.com; Leandrots82@gmail.com",
                "Teste de envio de e-mail",
                "Estou enviando este email como teste no console",
                out exeption);
            Console.WriteLine($"E-mail enviado {retornoEnvioEmail}\n{exeption}");
            Console.ReadKey();
        }
    }
}
