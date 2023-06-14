// See https://aka.ms/new-console-template for more information
using Capstone.Models;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OfficeOpenXml;
using System.Diagnostics;

namespace Capstone
{
    public static class Program
    {
        public static void Main(string[] args)
        {

            var chromeOptions = new ChromeOptions();
            string downloadDirectory = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("bin")) + "data";
            chromeOptions.AddUserProfilePreference("download.default_directory", @$"{downloadDirectory}");
            List<Company> companies = new();
            try
            {
                var jsonText = File.ReadAllText(downloadDirectory + "/companies.txt");
                companies = JsonConvert.DeserializeObject<List<Company>>(jsonText);
            }
            catch (Exception)
            {
                companies = GetBistCompanies(chromeOptions);
            }
            DownloadFinancialReports(chromeOptions, downloadDirectory, companies);
            PrepareExcelFiles(downloadDirectory);
            CollectFinancialData(companies);
            WriteGraphData(companies, downloadDirectory + "/JsonData");
            #region Data Okuma 
            //var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}//var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}//var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}//var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}//var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}//var d = new DirectoryInfo($@"{downloadDirectory}/Upload");
            //var files = d.GetFiles("*.xlsx");
            //foreach (var file in files)
            //{
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (var package = new ExcelPackage(new FileInfo(file.FullName.Replace('\\', '/'))))
            //    {
            //        var firstSheet = package.Workbook.Worksheets["A.V.O.D. KURUTULMUŞ GIDA VE TARIM ÜRÜNLERİ SANAYİ TİCARET A.Ş."];
            //        Console.WriteLine("Sheet 1 Data");
            //        Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A3"].Text}");
            //    }
            //}
            #endregion

        }


        public static List<Company> GetBistCompanies(ChromeOptions options)
        {
            List<Company> companies = new List<Company>();
            using (var driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl("https://www.kap.org.tr/tr/bist-sirketler");
                int numberOfCompanies = int.Parse(driver.FindElement(By.XPath("/html/body/div[7]/div/div/div/div[3]/p/span")).Text);
                int tableNo = 2;
                int rowNo = 2;
                for (int i = 0; i < numberOfCompanies; i++)
                {
                    try
                    {
                        driver.FindElement(By.XPath("/html/body/div[7]/div/div/div/div[3]/div/div[2]/div[" + tableNo + "]/div[" + rowNo + "]/div[1]"));
                    }
                    catch (Exception)
                    {
                        tableNo += 2;
                        rowNo = 2;
                        i--;
                        continue;
                    }
                    Company company = new();
                    for (int columnNo = 1; columnNo < 5; columnNo++)
                    {
                        string xPath = "/html/body/div[7]/div/div/div/div[3]/div/div[2]/div[" + tableNo + "]/div[" + rowNo + "]/div[" + columnNo + "]";
                        string value = driver.FindElement(By.XPath(xPath)).Text;
                        switch (columnNo)
                        {
                            case 1: company.Code = value; break;
                            case 2: company.Name = value; break;
                            case 3: company.City = value; break;
                            case 4: company.IndependentAuditingFirm = value; break;
                        }
                    }
                    companies.Add(company);
                    rowNo++;
                }
            }
            WriteCompaniesToATxt(companies);
            return companies;
        }
        public static void WriteCompaniesToATxt(List<Company> companies)
        {
            string txtFilePath = Environment.CurrentDirectory + "/../../../data/companies.txt";

            TextWriter writer = null;
            try
            {
                var contentsToWriteToFile = JsonConvert.SerializeObject(companies);
                writer = new StreamWriter(txtFilePath, false);
                writer.Write(contentsToWriteToFile);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }
        }
        public static void DownloadFinancialReports(ChromeOptions options, string downloadDirectory, List<Company> companies)
        {
            for (int i = 0; i < companies.Count / 50; i++)
            {
                using (var driver = new ChromeDriver(options))
                {
                    driver.Manage().Window.Maximize();
                    for (int j = 0; j < 50 && i * 50 + j < companies.Count; j++)
                    {
                        try
                        {
                            driver.Navigate().GoToUrl("https://www.kap.org.tr/");
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("/html/body/div[5]/form/input[1]")).SendKeys(companies[i * 50 + j].Code);
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("/html/body/div[5]/div/div[3]/a[1]")).Click();
                            Thread.Sleep(1000);
                            string a = "";
                            int columnNo = 0;
                            while (a != "Bildirim Sorgu")
                            {
                                columnNo++;
                                a = driver.FindElement(By.XPath("/html/body/div[7]/div/div/div[1]/div[2]/a[" + columnNo + "]/div")).Text;
                            }
                            driver.FindElement(By.XPath("/html/body/div[7]/div/div/div[1]/div[2]/a[" + columnNo + "]/div")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("//html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[1]/isteven-multi-select/span/button")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[1]/isteven-multi-select/span/div/div[2]/div[2]/div/label/span")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[3]/a")).Click();
                            Thread.Sleep(1000);
                            columnNo = 0;
                            a = "";
                            try
                            {
                                a = driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/disclosure-list/div/div/div/div[3]/div/span")).Text;
                                if (a == "Gösterilecek Bildirim Bulunamadı...")
                                    continue;
                            }
                            catch (Exception) { }
                            while (a != "Finansal Rapor")
                            {
                                columnNo++;
                                a = driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/disclosure-list/div/div/div/div[1]/disclosure-list-item[" + columnNo + "]/div/div/div/div/div[3]/div/div[3]/span")).Text;
                            }
                            driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/disclosure-list/div/div/div/div[1]/disclosure-list-item[" + columnNo + "]/div/div/div/div/div[3]/div/div[1]/span")).Click();
                            Thread.Sleep(2000);
                            DateTime date = DateTime.Parse(driver.FindElement(By.XPath("/html/body/div[11]/div/div/div[2]/div/div[5]/div/div[1]/div[2]")).Text);
                            Actions actions = new Actions(driver);
                            WebElement excelButton = (WebElement)driver.FindElement(By.XPath("/html/body/div[11]/div/div/div[3]/a[3]"));
                            actions.ContextClick(excelButton);
                            Thread.Sleep(2000);
                            driver.FindElement(By.XPath("/html/body/div[11]/div/div/div[3]/a[3]")).Click();
                            Thread.Sleep(5000);
                            File.Move(@$"{downloadDirectory}/Bildirimler.xls", @$"{downloadDirectory}/{companies[i * 50 + j].Name}({date.ToShortDateString()}).xls");
                            companies[i * 50 + j].xlsFile = @$"{downloadDirectory}/Upload/{companies[i * 50 + j].Name}({date.ToShortDateString()}).xlsx";
                            Thread.Sleep(1000);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
            }
        }
        public static void PrepareExcelFiles(string downloadDirectory)
        {
            var ps1File = @$"{downloadDirectory}\xlsConvert.ps1";
            var startInfo = new ProcessStartInfo()
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy ByPass -File \"{ps1File}\" \"{downloadDirectory}\"",
                UseShellExecute = false
            };
            var result = Process.Start(startInfo);
            result.WaitForExit();
        }
        public static List<Company> CollectFinancialData(List<Company> companies)
        {
            foreach (var company in companies)
            {
                if (!string.IsNullOrEmpty(company.xlsFile))
                    ReadExcel(company);
            }
            return companies;
        }
        public static void ReadExcel(Company company)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(company.xlsFile.Replace('\\', '/'))))
            {
                string worksheetName = company.Name;
                if (worksheetName.Length > 31)
                    worksheetName = worksheetName.Substring(0, 31);
                var worksheet = package.Workbook.Worksheets[company.Name];
                company.Data = new();
                company.Data.OperatingExpenses = new();
                company.Data.Ratios = new();
                company.Data.Tax = new();
                for (int i = 1; i < 1000; i++)
                {
                    while (worksheet.Cells["B" + i].Text == "Hasılat" && company.Data.Revenue == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.Revenue = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Satışların Maliyeti" && company.Data.CostOfRevenue == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.CostOfRevenue = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Esas Faaliyetin Karı" && company.Data.OperatingProfit == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.OperatingProfit = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "DÖNEM KARI (ZARARI)" && company.Data.NetProfit == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.NetProfit = result;
                        break;
                    } 
                    while (worksheet.Cells["B" + i].Text == "BRÜT KAR (ZARAR)" && company.Data.GrossProfit == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.GrossProfit = result;
                        break;
                    } 
                    while (worksheet.Cells["B" + i].Text == "Genel Yönetim Giderleri" && company.Data.OperatingExpenses.ManagementExpenses == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.OperatingExpenses.ManagementExpenses = result;
                        break;
                    } 
                    while (worksheet.Cells["B" + i].Text == "Pazarlama Giderleri" && company.Data.OperatingExpenses.MarketingExpenses == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.OperatingExpenses.MarketingExpenses = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Araştırma ve Geliştirme Giderleri" && company.Data.OperatingExpenses.RD_Expenses == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.OperatingExpenses.RD_Expenses = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Esas Faaliyetlerden Diğer Giderler" && company.Data.OperatingExpenses.OtherExpenses == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.OperatingExpenses.OtherExpenses = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Dönem Vergi (Gideri) Geliri" && company.Data.Tax.Tax1 == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.Tax.Tax1 = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Ertelenmiş Vergi (Gideri) Geliri" && company.Data.Tax.Tax2 == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.Tax.Tax2 = result;
                        break;
                    }
                    while (worksheet.Cells["B" + i].Text == "Amortisman ve İtfa Gideri İle İlgili Düzeltmeler" && company.Data.Amortization == 0 && i < 1000)
                    {
                        string revenue = worksheet.Cells["H" + i].Text;
                        revenue = revenue.Replace(".", "");
                        _ = int.TryParse(revenue, out int result);
                        company.Data.Amortization = result;
                        break;
                    }
                }
                if (company.Data.Revenue != 0)
                    CalculateRatios(company);
            }
        }
        public static void CalculateRatios(Company company)
        {
            var asd = JsonConvert.SerializeObject(company.Data, Formatting.Indented);
            Console.WriteLine(asd);
            return;
        }
        public static void WriteGraphData(List<Company> companies, string JsonPath)
        {
            foreach (Company company in companies)
            {
                if (company.Data != null && company.Data.Revenue != 0)
                {
                    try
                    {
                        File.Create($"{JsonPath}/{company.Name.Replace(".","")}.json");
                        var data = JsonConvert.SerializeObject(company, Formatting.Indented);
                        using (StreamWriter outputFile = new StreamWriter($"{JsonPath}/{company.Name}.json"))
                        {
                            outputFile.WriteLine(data);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
        }

    }
}


