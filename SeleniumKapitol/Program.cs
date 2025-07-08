using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string excelInputPath = Path.Combine(baseDirectory, "Kontakty.xlsx");
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string excelOutputPath = Path.Combine(desktopPath, "pecovatel_vysledky.xlsx");

        // ✅ Automatické zjištění počtu řádků
        int maxRadek;
        using (var package = new ExcelPackage(new FileInfo(excelInputPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            maxRadek = worksheet.Dimension.Rows; // Získání skutečného počtu řádků
        }

        Console.WriteLine($"****************************************************************");
        Console.WriteLine($"************ Vítá vás vyhledávácí software HEIMDALL ************");
        Console.WriteLine($"****************************************************************");
        Console.WriteLine($"      ************* Publisher: Morava Hemp *************");
        Console.WriteLine($"               ********************************");
        Console.WriteLine($"                        **************\n");

        // ✅ Výběr mezi rodným číslem a telefonním číslem
        Console.WriteLine("Co chcete vyhledávat?");
        Console.WriteLine("1 - Rodná čísla");
        Console.WriteLine("2 - Telefonní čísla");
        Console.Write("Zadejte volbu (1/2): ");

        int volba;
        while (!int.TryParse(Console.ReadLine(), out volba) || (volba != 1 && volba != 2))
        {
            Console.Write("Neplatný vstup. Zadejte 1 pro rodná čísla nebo 2 pro telefonní čísla: ");
        }

        int sloupec = (volba == 1) ? 3 : 1; // ✅ Sloupec v Excelu (3 = Rodná čísla, 1 = Telefonní čísla)

        Console.Write($"Od jakého řádku chcete začít kontrolovat? (1 - {maxRadek}): ");
        int startRow;
        while (!int.TryParse(Console.ReadLine(), out startRow) || startRow < 1 || startRow > maxRadek)
        {
            Console.Write($"Neplatný vstup. Zadejte číslo řádku mezi 1 a {maxRadek}: ");
        }

        Console.Write("Kolik čísel chcete načíst? ");
        int pocetRadku;
        while (!int.TryParse(Console.ReadLine(), out pocetRadku) || pocetRadku <= 0)
        {
            Console.Write("Neplatný vstup. Zadejte prosím číslo větší než 0: ");
        }

        // ✅ Zajištění, že nepřekročíme maxRadek
        if (startRow + pocetRadku - 1 > maxRadek)
        {
            pocetRadku = maxRadek - startRow + 1;
            Console.WriteLine($"Počet řádků byl upraven na {pocetRadku}, aby nepřekročil řádek {maxRadek}.");
        }

        List<string> seznamCisel = new List<string>();

        using (var package = new ExcelPackage(new FileInfo(excelInputPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            for (int row = startRow; row < startRow + pocetRadku; row++)
            {
                string cislo = worksheet.Cells[row, sloupec].Text.Trim();
                if (!string.IsNullOrEmpty(cislo))
                {
                    seznamCisel.Add(cislo);
                }
            }
        }

        Console.WriteLine($"Načteno {seznamCisel.Count} čísel od řádku {startRow}.\n");

        var excelPackage = new ExcelPackage();
        var worksheetOut = excelPackage.Workbook.Worksheets.Add("Výsledky");

        worksheetOut.Cells[1, 1].Value = (volba == 1) ? "Rodné číslo" : "Telefonní číslo";
        worksheetOut.Cells[1, 2].Value = "Výsledek";

        int rowIndex = 2;

        ChromeDriverService service = ChromeDriverService.CreateDefaultService();
        service.HideCommandPromptWindow = true;

        ChromeOptions options = new ChromeOptions();
        options.AddArgument("start-maximized");
        options.AddArgument("window-position=-32000,-32000");

        using (IWebDriver driver = new ChromeDriver(service, options))
        {
            try
            {
                driver.Navigate().GoToUrl("https://portal.kapitol.cz/");
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                var usernameInput = wait.Until(d => d.FindElement(By.Id("username")));
                usernameInput.SendKeys("XXX");

                var passwordInput = wait.Until(d => d.FindElement(By.Id("password")));
                passwordInput.SendKeys("XXX");

                var loginButton = wait.Until(d => d.FindElement(By.Id("kc-login")));
                loginButton.Click();

                Console.WriteLine("Přihlášení dokončeno.");
                driver.Navigate().GoToUrl("https://kappka.kapitol.cz/secured/consultant?division=KF");
                Console.WriteLine("Otevřena stránka Kappka.\n");

                foreach (var (cislo, index) in seznamCisel.Select((value, i) => (value, i + 1)))
                {
                    Console.WriteLine($"Zpracovávám {index}/{seznamCisel.Count}: {cislo}");

                    // ✅ Výběr správného vstupního pole
                    string inputXpath = (volba == 1)
                        ? "//input[@placeholder='Rodné číslo']"
                        : "//input[@placeholder='Telefon']";

                    var inputField = wait.Until(d => d.FindElement(By.XPath(inputXpath)));
                    inputField.Clear();
                    inputField.SendKeys(cislo);

                    var searchButton = wait.Until(d => d.FindElement(By.XPath("//button[contains(text(), 'Vyhledat klienta')]")));

                    if (searchButton.GetAttribute("disabled") == null)
                    {
                        try
                        {
                            searchButton.Click();
                        }
                        catch (Exception)
                        {
                            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                            js.ExecuteScript("arguments[0].click();", searchButton);
                        }
                        Console.WriteLine("Kliknuto na tlačítko 'Vyhledat klienta'.");
                    }
                    else
                    {
                        Console.WriteLine("Tlačítko 'Vyhledat klienta' je zakázané, přeskočeno.\n");
                    }

                    Thread.Sleep(150);

                    var pecovatelElements = driver.FindElements(By.XPath("//td[contains(@class, 'ng-star-inserted')]/div[contains(@class, 'd-flex')]/div"));



                    bool pozitivni = false;
                    string nalezenyPecovatel = "Nebyl nalezen žádný odpovídající klient.";

                    if (pecovatelElements.Count > 0)
                    {
                        foreach (var element in pecovatelElements)
                        {
                            string text = element.Text.Trim();

                            if (!string.IsNullOrEmpty(text))
                            {
                                nalezenyPecovatel = text;
                            }

                            if (text.StartsWith("KAPITOL") && !text.Contains("GOLD") && !text.Contains("AUTO"))
                            {
                                pozitivni = true;
                                break;
                            }
                        }
                    }

                    Console.WriteLine($"Pečovatel pro {cislo}: {nalezenyPecovatel}\n");

                    if (pozitivni)
                    {
                        worksheetOut.Cells[rowIndex, 1].Value = cislo;
                        worksheetOut.Cells[rowIndex, 2].Value = "pozitivní";
                        rowIndex++;
                    }
                }

                excelPackage.SaveAs(new FileInfo(excelOutputPath));
                Console.WriteLine($"Výsledky uloženy do souboru: {excelOutputPath}\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Došlo k chybě: {ex.Message}");
            }
            finally
            {
                driver.Quit();
                Console.WriteLine("Program dokončen.");
                Console.ReadLine();
            }
        }
    }
}
