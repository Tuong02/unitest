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
    public class Search
    {
        [TestMethod]
        public void TestSearch()
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
                ISheet sheet = workbook.GetSheetAt(1); // Lấy sheet thứ 2

                // Nhập tìm kiếm lần đầu tiên
                IWebElement searchInput = driver.FindElement(By.CssSelector(".css-tihntr-MuiInputBase-input-MuiOutlinedInput-input"));
                searchInput.SendKeys(Keys.Enter);
                Thread.Sleep(1000);

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy giá trị tìm kiếm từ file Excel
                        string search = row.GetCell(0)?.ToString();

                        // Nhập từ khóa tìm kiếm
                        IWebElement input = driver.FindElement(By.CssSelector(".css-mqdhaa-MuiInputBase-input-MuiOutlinedInput-input"));
                        input.SendKeys(Keys.Control + "a");
                        input.SendKeys(search);
                        Thread.Sleep(1000);

                        // Nhấn nút tìm kiếm
                        IWebElement submit = driver.FindElement(By.CssSelector(".css-150zjap-MuiButtonBase-root-MuiButton-root"));
                        submit.Click();
                        Thread.Sleep(3000);

                        try
                        {
                            // Lấy kết quả tìm kiếm
                            Thread.Sleep(4000);
                            string resultSearch = driver.FindElement(By.CssSelector(".MuiTypography-root.MuiTypography-body2.css-1ws7hsa-MuiTypography-root")).Text;
                            if (resultSearch.Contains("Tìm thấy 0/3434 sản phẩm"))
                            {
                                row.CreateCell(1).SetCellValue("No data");
                            }
                            else
                            {
                                row.CreateCell(1).SetCellValue("Pass");
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
