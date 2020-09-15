using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using NUnit.Framework.Constraints;

namespace EasyExcel
{
    public class Elements
    {
        public Application excel;
        public Worksheet sheet;
        public Workbook book;


        /// <summary>
        /// Запускаем экземпляр excel
        /// </summary>
        public void Start()
        {
            excel = new Application();

            // Скрываем excel от пользователя
            excel.Visible = false;
            excel.DisplayAlerts = false;
        }

        /// <summary>
        /// Открываем файл для чтения/записи
        /// </summary>
        /// <param name="file">Имя файла Excel</param>
        public void Open(string file)
        {
            excel.Workbooks.Open(file);
        }

        /// <summary>
        /// Создание книги
        /// </summary>
        public void createWorkbook()
        {
            book = excel.Workbooks.Add(Type.Missing);
        }

        /// <summary>
        /// Функция создает 'X' листов в новом документе
        /// </summary>
        /// <param name="number">Количество листов</param>
        public void createSheet(int number = 1)
        {
            if (number > 0)
            {
                book.Sheets.Add(book.Sheets[1], Type.Missing, number, Type.Missing);
            }
            else { throw new Exception("number of sheets can't be below 0"); }
        }

        /// <summary>
        /// Функция устанавливает рабочий лист, с которого в последствии можно будет чтитать данные
        /// </summary>
        /// <param name="index">Номер листа (начинается с '1')</param>
        public void setWorksheet(int index)
        {
            if (index <= excel.Worksheets.Count && index > 0)
            {
                sheet = (Worksheet)excel.Worksheets.Item[index];
            }
            else
            {
                throw new Exception("Out of range");
            }
        }

        public bool Save(string file)
        {
            // (по default файл сохраняет в папку Documents)
            excel.Application.ActiveWorkbook.SaveAs(file);
            return true;
        }

        /// <summary>
        /// Выход из excel и очистка памяти
        /// </summary>
        public void Stop()
        {
            // Выходим из приложения
            excel.Application.Quit();
            excel.Quit();

            // Очищаем память
            excel = null;
            sheet = null;
            book = null;

            GC.Collect();
        }
    }



    [TestFixture]
    class ElementsTest
    {
        Elements elements;
        string path = Environment.CurrentDirectory + "/";

        [SetUp]
        public void Start()
        {
            elements = new Elements();
            elements.Start();
        }

        [Test]
        public void createDoc()
        {
            Assert.DoesNotThrow(() =>
            {
                elements.createWorkbook();
                elements.createSheet(1);
                elements.Save(path+"test.xlsx");
            });
        }

        [TestCase(1, true)]
        [TestCase(10, true)]
        [TestCase(0, false)]
        [TestCase(-5, false)]
        public void sheetsNumber(int number, bool result)
        {
            elements.createWorkbook();
            if (result)
            {
                Assert.DoesNotThrow(() =>
                {
                    elements.createSheet(number);
                });
            }
            else
            {
                Assert.Throws<Exception>(() =>
                {
                    elements.createSheet(number);
                });
            }
            elements.Save(path+"test.xlsx");
        }

        [TestCase(1, 1, true)]
        [TestCase(1, 3, false)]
        [TestCase(2, 1, true)]
        [TestCase(4, 0, false)]
        [TestCase(4, -1, false)]
        [TestCase(10, 7, true)]
        public void setSheet(int number, int count, bool result)
        {
            elements.createWorkbook();
            elements.createSheet(number);

            if (result)
            {
                Assert.DoesNotThrow(() => {
                    elements.setWorksheet(count);
                });
            }
            else
            {
                Assert.Throws<Exception>(() => {
                    elements.setWorksheet(count);
                });
            }
        }

        [TestCase("123")]
        [TestCase("db_1.txt")]
        [TestCase("MY_BiG_DATABASE_9919128.txt")]
        public void savedFileName(string name)
        {
            elements.createWorkbook();

            Assert.DoesNotThrow(() => { 
                elements.Save(path+name);          
            });
        }

        [TearDown]
        public void Stop()
        {
            elements.Stop();
        }
    }
}
