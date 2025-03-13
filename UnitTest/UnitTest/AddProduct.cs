using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class AddProduct
    {
        [TestMethod]
        public void TestAddProduct()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            // Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("admin");
            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("@dmIn12");
            IWebElement login = wait.Until(d => d.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root")));
            login.Click();
            Thread.Sleep(5000);

            // Chuyển đến trang sản phẩm
            IWebElement product = driver.FindElement(By.XPath("//div[@role='button' and .//div[text()='Sản phẩm']]"));
            product.Click();
            Thread.Sleep(1000);

            string filePath = @"C:\DATN\datainputandouput.xlsx";
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(7);

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        IWebElement addProduct = driver.FindElement(By.CssSelector("a[href='/dashboard/app/products/create']"));
                        addProduct.Click();
                        Thread.Sleep(2000);

                        string productName = row.GetCell(0)?.ToString();
                        string skuCode = row.GetCell(1)?.ToString();
                        string productType = row.GetCell(2)?.ToString();
                        string productQuantity = row.GetCell(3)?.ToString();
                        string productWarranty = row.GetCell(4)?.ToString();
                        string productIntroduce = row.GetCell(5)?.ToString();
                        string productPrice = row.GetCell(6)?.ToString();
                        string productPriceTT = row.GetCell(7)?.ToString();
                        string productImages = row.GetCell(8)?.ToString();

                        IWebElement nameField = driver.FindElement(By.Name("name"));
                        nameField.SendKeys(Keys.Control + "a");
                        nameField.SendKeys(Keys.Delete);
                        nameField.SendKeys(productName);

                        IWebElement codeField = driver.FindElement(By.Name("sku"));
                        codeField.SendKeys(Keys.Control + "a");
                        codeField.SendKeys(Keys.Delete);
                        codeField.SendKeys(skuCode);

                        IWebElement typeField = driver.FindElement(By.Name("variantName"));
                        typeField.SendKeys(Keys.Control + "a");
                        typeField.SendKeys(Keys.Delete);
                        typeField.SendKeys(productType);

                        IWebElement quantityField = driver.FindElement(By.Name("quantity"));
                        quantityField.SendKeys(Keys.Control + "a");
                        quantityField.SendKeys(Keys.Delete);
                        quantityField.SendKeys(productQuantity);

                        IWebElement warrantyField = driver.FindElement(By.Name("warrantyPeriod"));
                        warrantyField.SendKeys(Keys.Control + "a");
                        warrantyField.SendKeys(Keys.Delete);
                        warrantyField.SendKeys(productWarranty);

                        IWebElement introduceField = driver.FindElement(By.CssSelector("div.ql-editor"));
                        introduceField.SendKeys(Keys.Control + "a");
                        introduceField.SendKeys(Keys.Delete);
                        introduceField.SendKeys(productIntroduce);

                        // brand
                        IWebElement inputElement = driver.FindElements(By.CssSelector("input[aria-autocomplete='list']"))[1];
                        inputElement.Click();
                        inputElement.SendKeys(Keys.Delete);
                        IWebElement option = driver.FindElement(By.XPath($"//li[contains(text(), 'Samsung')]"));
                        option.Click();

                        // category
                        IWebElement categoryField = driver.FindElements(By.CssSelector("input[aria-autocomplete='list']"))[2];
                        categoryField.Click();
                        categoryField.SendKeys(Keys.Delete);
                        IWebElement options = driver.FindElement(By.XPath($"//li[contains(text(), 'Laptop')]"));
                        options.Click();

                        // price
                        IWebElement priceField = driver.FindElement(By.Name("price"));
                        priceField.SendKeys(Keys.Control + "a");
                        priceField.SendKeys(Keys.Delete);
                        priceField.SendKeys(productPrice);

                        // priceTT
                        IWebElement priceAllField = driver.FindElement(By.Name("marketPrice"));
                        priceAllField.SendKeys(Keys.Control + "a");
                        priceAllField.SendKeys(Keys.Delete);
                        priceAllField.SendKeys(productPriceTT);

                        // enter image
                        IWebElement imageField = driver.FindElements(By.CssSelector(".MuiSwitch-input.PrivateSwitchBase-input.css-mraihx"))[0];
                        imageField.Click();
                        IWebElement urlInput = driver.FindElements(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"))[2];
                        urlInput.SendKeys(productImages);

                        if (productName.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else if (productName.Length > 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (skuCode.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else if (productType.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép bỏ trống");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productType.Length > 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productQuantity.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép bỏ trống");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productQuantity.Length >= 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }

                        //else if (productQuantity <= 0)
                        else if (productWarranty.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép bỏ trống");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productWarranty.Length >= 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productIntroduce.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else if (productIntroduce.Length >= 50)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 50 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productPrice.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else if (productPrice.Length >= 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productPriceTT.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else if (productPriceTT.Length >= 25)
                        {
                            row.CreateCell(9).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            // click add button
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productImages.Length == 0)
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                        }
                        else
                        {
                            row.CreateCell(9).SetCellValue("Pass");
                            IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                            addButton.Click();
                            Thread.Sleep(5000);
                        }

                    }
                }

                using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeStream);
                }
            }

            driver.Quit();
        }
    }
}