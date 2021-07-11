﻿using System;
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
            String path = "";

            var options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", path);
           
            ChromeDriver driver = new ChromeDriver(options);
            driver.Manage().Window.Maximize();

            //Ingresar a CANVAS
            driver.Url = "https://canvas.unapec.edu.do/";

            //Seleccionar modulo de estudiantes
            Thread.Sleep(5000);
            IWebElement button = driver.FindElement(By.CssSelector("body > header > div > div.row.header-info > div.botones-opcion.col-12 > div > div:nth-child(1) > a"));
            button.Click();

            //Ingresar correo
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("#i0116")).SendKeys(user);
            driver.FindElement(By.CssSelector("#idSIButton9")).Click();

            //Ingresar contraseña
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("#i0118")).SendKeys(pass);
            driver.FindElement(By.CssSelector("#idSIButton9")).Click();

            //No recordar contraseña
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("#idBtn_Back")).Submit();

            //Seleccionar la primera materia
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("#DashboardCard_Container > div > div > div:nth-child(1)")).Click();

            //Ir a archivos
            Thread.Sleep(5000);
            driver.FindElement(By.ClassName("files")).Click();

            //Visualizar el primer archivo
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("#content > div > div.ef-main > div > div > div > div:nth-child(3) > div.ef-name-col > a")).Click();

            //Descargar
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("body > span > span > span > div > div.ef-file-preview-header > div > a > span")).Click();

            //driver.Close();
        }
    }
}
