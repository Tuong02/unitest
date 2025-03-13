using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTest
{
    [TestClass]
    public class AddCart
    {
        [TestMethod]
        public void TestAddCart()
        {
            string filePath = @"C:\DATN\datainputandouput.xlsx";
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(12); // Lấy sheet thứ 2

                // Nhập tìm kiếm lần đầu tiên
                IWebElement searchInput = wait.Until(d => d.FindElement(By.CssSelector(".css-tihntr-MuiInputBase-input-MuiOutlinedInput-input")));
                searchInput.SendKeys(Keys.Enter);
                //Thread.Sleep(1000);

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy giá trị tìm kiếm từ file Excel
                        string search = row.GetCell(0)?.ToString();

                        // Nhập từ khóa tìm kiếm
                        IWebElement input = wait.Until(d => d.FindElement(By.CssSelector(".css-mqdhaa-MuiInputBase-input-MuiOutlinedInput-input")));
                        input.SendKeys(Keys.Control + "a");
                        input.SendKeys(search);
                        //Thread.Sleep(1000);

                        // Nhấn nút tìm kiếm
                        IWebElement submit = wait.Until(d => d.FindElement(By.CssSelector(".css-150zjap-MuiButtonBase-root-MuiButton-root")));
                        submit.Click();
                        //Thread.Sleep(3000);

                        string xpath = $"//a[@aria-label='{search}']";
                        IWebElement btnChose = wait.Until(d => d.FindElement(By.XPath(xpath)));
                        btnChose.Click();
                        //Thread.Sleep(3000);

                        // Thêm vào giỏ hàng
                        IWebElement addCart = wait.Until(d => d.FindElement(By.CssSelector(".css-dwtauf-MuiButtonBase-root-MuiButton-root")));
                        addCart.Click();

                        row.CreateCell(1).SetCellValue("Pass");

                        // Nhập tìm kiếm lần đầu tiên
                        IWebElement btnSearch = wait.Until(d => d.FindElement(By.CssSelector(".css-tihntr-MuiInputBase-input-MuiOutlinedInput-input")));
                        btnSearch.SendKeys(Keys.Enter);

                    }
                }

                // Lưu lại file Excel sau khi ghi kết quả
                using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeStream);
                }
            }

            driver.Quit();
        }
    }
}
