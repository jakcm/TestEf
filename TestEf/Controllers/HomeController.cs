using Aspose.Cells;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TestEf.Models;

namespace TestEf.Controllers
{
    public class HomeController : Controller
    {
        private ICollection GetDataBulk(int startIndex, int bulkCount, int maxCount = 300000)
        {
            ICollection result = null;
            using (TestEfEntities context = new TestEfEntities())
            {
                result = context.Product.OrderBy(i => i.ID).Where(i => i.ID <= maxCount)
                    .Skip(startIndex).Take(bulkCount).ToList();
            }

            return result;
        }

        private int ExportDataBulk(Func<int, int, ICollection> getDataFunc)
        {
            int totalCount = 0;
            int bulkCount = 20000;
            string templateFile = System.Web.HttpContext.Current.Server.MapPath(@"~\Excel\Template\ExcelTemplate.xlsx");

            string exportFileName = DateTime.Now.Ticks.ToString();
            exportFileName = System.Web.HttpContext.Current.Server.MapPath(@"~\Excel\Export\" + exportFileName + ".xlsx");

            object[,] templateValueArray = null;
            using (FileStream fileStream = new FileStream(templateFile, FileMode.Open))
            {
                WorkbookDesigner designer = new WorkbookDesigner(new Workbook(fileStream));
                Cells cells = designer.Workbook.Worksheets[0].Cells;

                templateValueArray = cells.ExportArray(1, 0, 1, cells.Columns.Count);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            while (totalCount <= 1000000)
            {
                ICollection bulkData = getDataFunc(totalCount, bulkCount);
                if (bulkData == null || bulkData.Count <= 0)
                {
                    break;
                }

                string loadFileName = (totalCount <= 0 ? templateFile : exportFileName);
                using (FileStream fileStream = new FileStream(loadFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (Workbook workbook = new Workbook(fileStream))
                    {
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);
                        using (Cells cells = workbook.Worksheets[0].Cells)
                        {
                            if (totalCount > 0)
                            {
                                cells.ImportTwoDimensionArray(templateValueArray, totalCount + 1, 0);
                            }

                            designer.SetDataSource("Product", bulkData);
                            designer.Process();
                            designer.ClearDataSource();

                            XlsSaveOptions xlsSaveOptions = new XlsSaveOptions() { ClearData = true };
                            workbook.Save(exportFileName, xlsSaveOptions);
                            cells.Dispose();
                        }
                        workbook.Dispose();
                    }
                    fileStream.Close();
                    fileStream.Dispose();
                }

                totalCount += bulkData.Count;
                if (bulkData.Count < bulkCount)
                {
                    break;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return totalCount;
        }

        private void ExportToExcel()
        {
        }

        public ActionResult Get10W()
        {
            Stopwatch watch = Stopwatch.StartNew();
            int productCount = ExportDataBulk((startIndex, bulkCount) =>
            {
                return this.GetDataBulk(startIndex, bulkCount, 100000);
            });

            watch.Stop();
            ViewData["ProductCount"] = productCount + "; " + (watch.ElapsedMilliseconds / 1000) + "S";
            return View();
        }

        public ActionResult Get30W()
        {
            Stopwatch watch = Stopwatch.StartNew();
            int productCount = ExportDataBulk((startIndex, bulkCount) =>
            {
                return this.GetDataBulk(startIndex, bulkCount, 300000);
            });

            watch.Stop();
            ViewData["ProductCount"] = productCount + "; " + (watch.ElapsedMilliseconds / 1000) + "S";
            return View();
        }

        public ActionResult Get50W()
        {
            Stopwatch watch = Stopwatch.StartNew();
            int productCount = ExportDataBulk((startIndex, bulkCount) =>
            {
                return this.GetDataBulk(startIndex, bulkCount, 500000);
            });

            watch.Stop();
            ViewData["ProductCount"] = productCount + "; " + (watch.ElapsedMilliseconds / 1000) + "S";
            return View();
        }

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
    }
}