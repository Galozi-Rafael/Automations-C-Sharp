using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using Microsoft.Playwright;

namespace ExtrairExcel.Services
{
    public class WebService
    {
        // Abre uma página no navegador e navega até a URL especificada.
        public async Task<IPage> OpenPageAsync(string url)
        {
            var playwright = await Playwright.CreateAsync();

            var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = true,
            });

            var page = await browser.NewPageAsync();

            await page.GotoAsync(url);

            return page;
        }

        // Realiza o login na página usando o nome de usuário e senha fornecidos.
        public async Task LoginAsync(IPage page, string username, string password)
        {
            await page.WaitForSelectorAsync("#username");
            await page.FillAsync("#username", username);
            await page.FillAsync("#password", password);
            await page.ClickAsync("#log-in");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }


        // Realiza o scrape da tabela de transações na página e retorna os dados como uma lista de listas de strings.
        public async Task<List<List<string>>> ScrapeTransactionsTableAsync(IPage page)
        {
            var tableSelector = "div.element-wrapper";  //"table.table.table-padded";

            await page.WaitForSelectorAsync(tableSelector);

            var RowElements = await page.QuerySelectorAllAsync($"{tableSelector} tbody tr");

            var result = new List<List<string>>();

            var headers = await page.QuerySelectorAllAsync($"{tableSelector} thead tr th");

            if (headers.Count > 0)
            {
                var headerData = new List<string>();

                foreach (var header in headers)
                {
                    var text = (await header.InnerTextAsync()).Trim();
                    headerData.Add(NormalizeSpace(text));
                }
                result.Add(headerData);
            }
            

            foreach (var row in RowElements)
            {
                var cellElements = await row.QuerySelectorAllAsync("td");
                
                if (cellElements.Count == 0)
                {
                    continue;
                }
                    
                var rowData = new List<string>();

                foreach (var cell in cellElements)
                {
                    var text = (await cell.InnerTextAsync()).Trim();
                    rowData.Add(NormalizeSpace(text));
                }
                
                result.Add(rowData);
            }

            return result;
        }

        // Normaliza os espaços em branco em uma string, removendo espaços extras e quebras de linha.
        private string NormalizeSpace(string text)
        {
            return string.Join(" ",
                text.Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries));
        }
    }
}
