// See https://aka.ms/new-console-template for more information
using Capstone.Models;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;

namespace Capstone
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var chromeOptions = new ChromeOptions();
            string downloadDirectory = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("bin")) + "data";
            chromeOptions.AddUserProfilePreference("download.default_directory", @$"{downloadDirectory}");
            using (var driver = new ChromeDriver(chromeOptions))
            {
                driver.Manage().Window.Maximize();
                List<Company> companies = new();
                try
                {
                    var jsonText = File.ReadAllText(downloadDirectory + "/companies.txt");
                    companies = JsonConvert.DeserializeObject<List<Company>>(jsonText);
                }
                catch (Exception)
                {
                    companies = GetBistCompanies(driver);
                }
                DownloadFinancialReports(driver, downloadDirectory, companies);
            }
        }
        public static List<Company> GetBistCompanies(ChromeDriver driver)
        {
            driver.Navigate().GoToUrl("https://www.kap.org.tr/tr/bist-sirketler");
            int numberOfCompanies = int.Parse(driver.FindElement(By.XPath("/html/body/div[7]/div/div/div/div[3]/p/span")).Text);
            List<Company> companies = new List<Company>();
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
        public static void DownloadFinancialReports(ChromeDriver driver, string downloadDirectory, List<Company> companies)
        {
            for (int i = 0; i < companies.Count; i++)
            {
                driver.Navigate().GoToUrl("https://www.kap.org.tr/");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[5]/form/input[1]")).SendKeys(companies[i].Code);
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
                File.Move(@$"{downloadDirectory}/Bildirimler.xls", @$"{downloadDirectory}/{companies[i].Name}({date.ToShortDateString()}).xls");
                Thread.Sleep(1000);
            }
        }
    }
}


