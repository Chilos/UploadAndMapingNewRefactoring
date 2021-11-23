using System.Data;
using System.Reflection;
using excelObj = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System;

namespace UploadAndMapingNew
{
    public sealed class UploadNameOfFieldsFromExcel : ICloneable
    {
        private DataTable _paramsFromExcel = new DataTable(); //Лист, который будет содержать названия колонок из файла Excel
        private DataTable _paramsFromExcelDocs = new DataTable();
        private DataTable _paramsFromExcelVehisle = new DataTable();
        public DataTable paramsFromExcel { get { return _paramsFromExcel; } set { _paramsFromExcel = value; } }
        public DataTable paramsFromExcelDocs { get { return _paramsFromExcelDocs; } set { _paramsFromExcelDocs = value; } }
        public DataTable paramsFromExcelVehisle { get { return _paramsFromExcelVehisle; } set { _paramsFromExcelVehisle = value; } }

        public object Clone()
        {
            return new UploadNameOfFieldsFromExcel
            {
                paramsFromExcelDocs = paramsFromExcelDocs.Copy(),
                paramsFromExcelVehisle = paramsFromExcelVehisle.Copy()
            };                        
        }
              

        public DataTable GetFieldNamesFromExcel(OpenFileDialog ofd)
        {
            //После выбора файла создается новы объект Application
            excelObj.Application app = new excelObj.Application();
            //Который может содержать одну или более книг, ссылки на которые содержит свойство _workbook
            excelObj.Workbook workbook;
            //книги могут содержать одну или несколько страниц, ссылки на которые содержит войство Worksheet
            excelObj.Worksheet newSheet;
            //Страницы могут содержать ячейки или группы ячеек, ссылки на которые содержит свойство Range
            excelObj.Range sheetRange;

            //Для открытия документа используется метод Excele.Workbooks.Open
            //Основным параметром является путь к файлу, остальные можно оставить пустыми
            workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value);

            _paramsFromExcelDocs.Columns.Add(new DataColumn("Параметры на вход"));//Создадим новую колонку, в которую потом будем класть полученные назания колонок в виде строк
            _paramsFromExcelVehisle.Columns.Add(new DataColumn("Параметры на вход"));//Создадим новую колонку, в которую потом будем класть полученные назания колонок в виде строк


            //Сначала пишем сюда колонки с листа Заказы
            newSheet = (excelObj.Worksheet)workbook.Worksheets["Заказы"]; //это доступ по названию листа
            sheetRange = newSheet.UsedRange;

            for (int i = 1; i <= sheetRange.Columns.Count; i++) // в Excel строки начинаются не с 0, а с 1, поэтому i = 1
            {
                DataRow dr = _paramsFromExcelDocs.NewRow();
                dr[0] = (sheetRange.Cells[1, i] as excelObj.Range).Value2;
                _paramsFromExcelDocs.Rows.Add(dr);
            }

            //Теперь дописываем колонки с листа Машины
            newSheet = (excelObj.Worksheet)workbook.Worksheets["Авто"]; //это доступ по названию листа
            sheetRange = newSheet.UsedRange;

            for (int i = 1; i <= sheetRange.Columns.Count; i++)
            {
                DataRow dr = _paramsFromExcelVehisle.NewRow();
                dr[0] = (sheetRange.Cells[1, i] as excelObj.Range).Value2;
                _paramsFromExcelVehisle.Rows.Add(dr);
            }

            app.Quit();

            return paramsFromExcel;
        }
    }
}
