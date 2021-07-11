using System;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace RPA
{
    class Program
    {
        static void Main(string[] args)
        {
            String user = "";
            String pass = "";

            //Desktop como carpeta de descarga
            var options = new ChromeOptions();
            String path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Console.WriteLine(path);
            options.AddUserProfilePreference("download.default_directory", path);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("download.directory_upgrade", true);

            try
            {
                using (ChromeDriver driver = new ChromeDriver(options))
                {
                    int sleepTime = 5000;
                    driver.Manage().Window.Maximize();

                    ////Ingresar a CANVAS
                    driver.Url = "https://canvas.unapec.edu.do/";

                    //Seleccionar modulo de estudiantes
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.CssSelector("body > header > div > div.row.header-info > " +
                        "div.botones-opcion.col-12 > div > div:nth-child(1) > a")).Click();

                    //Ingresar correo
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.CssSelector("#i0116")).SendKeys(user);
                    driver.FindElement(By.CssSelector("#idSIButton9")).Click();

                    //Ingresar contraseña
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.CssSelector("#i0118")).SendKeys(pass);
                    driver.FindElement(By.CssSelector("#idSIButton9")).Click();

                    //No recordar contraseña
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.CssSelector("#idBtn_Back")).Submit();

                    //Seleccionar la primera materia
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.Id("global_nav_courses_link")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("#nav-tray-portal > span > span > div > " +
                        "div > div > div > div > ul:nth-child(3) > li:nth-child(1) > a")).Click();

                    //Ir a archivos
                    Thread.Sleep(sleepTime);
                    driver.FindElement(By.ClassName("files")).Click();

                    //Buscar los elementos descargables
                    Thread.Sleep(sleepTime);
                    var elementosDescargables = driver.FindElements(By.CssSelector("#content > div > div.ef-main > div > div > div > div > div.ef-name-col > a"));
                    
                    // Buscar los elementos que sean válidos para descargar: pdf, word o powerpoint
                    var elementosValido = elementosDescargables.Where(r => r.Text.ToLower().Contains(".ppt") ||
                                        r.Text.ToLower().Contains(".pdf") || r.Text.ToLower().Contains(".doc"));

                    Thread.Sleep(sleepTime);
                    IWebElement downloadButton = null;

                    // Buscamos un elemento que se pueda descargar.
                    // Existen elementos que aparecen pero no se pueden descargar porque no
                    // esta lista la fecha

                    foreach(var elemento in elementosValido)
                    {
                        try
                        {
                            elemento.Click();
                            downloadButton = driver.FindElement(By.CssSelector("body > span > span > span > div > div.ef-file-preview-header > div > a > span"));
                            break;
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    // Descargar
                    downloadButton.Click();
                    Thread.Sleep(sleepTime);

                    //driver.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Ha ocurrido un error!");
            }
        }
    }
}
