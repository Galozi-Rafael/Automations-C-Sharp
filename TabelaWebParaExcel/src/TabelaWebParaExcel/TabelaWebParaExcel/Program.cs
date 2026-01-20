using DocumentFormat.OpenXml.Wordprocessing;
using ExtrairExcel.Config;
using ExtrairExcel.Services;
using Microsoft.Extensions.Configuration;
using System;
using System.Threading.Tasks;

namespace ExtrairExcel
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Iniciando o robô...");

            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .Build();

            var settings = configuration.Get<AppSettings>();

            if (settings == null)
            {
                Console.WriteLine("Erro: não foi possível carregar as configurações.");
                return;
            }

            var webservice = new WebService();
            var page = await webservice.OpenPageAsync(settings.Login.Url);
            
            Console.WriteLine("Página aberta com sucesso!");

            await webservice.LoginAsync(page, settings.Login.Username, settings.Login.Password);

            Console.WriteLine("Login efetuado com sucesso!");

            Console.WriteLine("Começando o scrape.");

            var table = await webservice.ScrapeTransactionsTableAsync(page);

            foreach (var row in table)
            {
                foreach (var cell in row)
                {
                    Console.Write(cell + ",");
                }

                Console.WriteLine();
            }

            Console.WriteLine();
            Console.WriteLine("Scrape concluído com sucesso!");
            Console.Write("Pressione qualquer tecla para sair...");
            Console.ReadKey();
        }
    }
    
}

