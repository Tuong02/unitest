using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;

namespace UnitTest
{
    [TestClass]
    public class Search
    {
        [TestMethod]
        public void TestSearch()
        {
            string filePath = @"E:\datainputandouput.xlsx";
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
                ISheet sheet = workbook.GetSheetAt(1); // Lấy sheet 2

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        string search = row.GetCell(0)?.ToString();
                        wait.Until(d => d.FindElement(By.CssSelector("div.header-input input[type='text']")));
                        IWebElement searchInput = driver.FindElement(By.CssSelector("div.header-input input[type='text']"));
                        IWebElement searchIcon = driver.FindElement(By.CssSelector("div.header__input--icon"));
                        searchInput.Clear();
                        searchInput.SendKeys(search);
                        searchIcon.Click();

                        try
                        {
                            // Chờ và kiểm tra xem có kết quả tìm kiếm hay không
                            wait.Until(d => d.FindElements(By.CssSelector("div.search")).Count > 0);

                            // Kiểm tra xem có kết quả tìm kiếm hoặc không có kết quả tìm kiếm
                            IWebElement searchResult = null;
                            IWebElement searchNoResult = null;
                            try
                            {
                                searchResult = driver.FindElement(By.CssSelector("div.search a.search-item"));
                            }
                            catch (NoSuchElementException)
                            {
                                // Không có kết quả tìm kiếm
                                searchNoResult = driver.FindElement(By.CssSelector("div.search.no-product"));
                            }

                            if (searchResult != null)
                            {
                                // Lấy thông tin sản phẩm nếu có kết quả tìm kiếm
                                string productName = searchResult.FindElement(By.CssSelector("div.content b")).Text;
                                string productPrice = searchResult.FindElement(By.CssSelector("div.content .price")).Text;
                                string productImageSrc = searchResult.FindElement(By.CssSelector("img")).GetAttribute("src");

                                Console.WriteLine($"Product Name: {productName}");
                                Console.WriteLine($"Product Price: {productPrice}");
                                Console.WriteLine($"Product Image URL: {productImageSrc}");
                                row.CreateCell(1).SetCellValue("Hiển thị kết quả tìm kiếm thành công");
                            }
                            else if (searchNoResult != null)
                            {
                                // Ghi thất bại nếu không có kết quả tìm kiếm
                                Console.WriteLine("Không có kết quả tìm kiếm.");
                                row.CreateCell(1).SetCellValue("No data");
                            }
                        }
                        catch (WebDriverTimeoutException)
                        {
                            Console.WriteLine("Không có kết quả tìm kiếm hoặc trang không hiển thị đúng.");
                            row.CreateCell(1).SetCellValue("No data");
                        }
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
