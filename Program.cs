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
            using (var driver = new ChromeDriver())
            {
                driver.Manage().Window.Maximize();
                List<Company> companies = new();
                try
                {
                    var jsonText = File.ReadAllText(Environment.CurrentDirectory + "/companies.txt");
                    companies = JsonConvert.DeserializeObject<List<Company>>(jsonText);
                }
                catch (Exception)
                {
                    companies = GetBistCompanies(driver);
                }
                driver.Navigate().GoToUrl("https://www.kap.org.tr/");
                driver.FindElement(By.XPath("/html/body/div[5]/form/input[1]")).SendKeys(companies[0].Code);
                driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/div/div[3]/a[1]")).Click();
                driver.FindElement(By.XPath("/html/body/div[7]/div/div/div[1]/div[2]/a[6]/div")).Click();
                driver.FindElement(By.XPath("//html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[1]/isteven-multi-select/span/button")).Click();
                driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[1]/isteven-multi-select/span/div/div[2]/div[2]/div/label/span")).Click();
                driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/div/div[3]/a")).Click();
                driver.FindElement(By.XPath("/html/body/div[10]/div/div/div[2]/div/div[2]/div/disclosure-list/div/div/div/div[1]/disclosure-list-item[3]/div/div/div/div/div[3]/div/div[1]/span")).Click();
                Actions actions = new Actions(driver);
                WebElement excelButton = (WebElement)driver.FindElement(By.XPath("/html/body/div[11]/div/div/div[3]/a[3]"));
                actions.ContextClick(excelButton);
                driver.FindElement(By.XPath("/html/body/div[11]/div/div/div[3]/a[3]")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/form/div")).Click();
            }
        }
        public static List<Company> GetBistCompanies(ChromeDriver driver)
        {
            driver.Navigate().GoToUrl("https://www.kap.org.tr/tr/bist-sirketler");
            int numberOfCompanies = int.Parse(driver.FindElement(By.XPath("/html/body/div[7]/div/div/div/div[3]/p/span")).Text);
            List<Company> companies = new List<Company>();
            int tableNo = 2;
            int rowNo = 2;
            for(int i = 0; i < numberOfCompanies; i++)
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
                Company company = new ();
                for(int columnNo = 1; columnNo < 5; columnNo++)
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
            string txtFilePath = Environment.CurrentDirectory + "/companies.txt";

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
    }
}


