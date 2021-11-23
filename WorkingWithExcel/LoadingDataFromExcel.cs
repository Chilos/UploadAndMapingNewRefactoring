using System.Data;
using System.Reflection;
using excelObj = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System;

namespace UploadAndMapingNew
{        
    public sealed class LoadingDataFromExcel
    {
        public DataTable dataFromExcelDocs = new DataTable();
        public DataTable dataFromExcelVehicle = new DataTable();

        private excelObj.Workbook _workbook;
        private excelObj.Worksheet _nwSheet;
        private excelObj.Range _shtRange;

        //TODO Может если сделать через Microsoft.Jet.... будет по быстрее работать загрузка из файла?
        public void GetDataFromExcel(OpenFileDialog ofd)
        {
            excelObj.Application app = new excelObj.Application();

            _workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value);
            GetIncomingExcelDataTableDocs();
            GetIncomingExcelDataTableVehicle();
            AppQuit(app);


        }
        
        protected void GetIncomingExcelDataTableDocs()
        {
            try
            {
                _nwSheet = (excelObj.Worksheet)_workbook.Worksheets["Заказы"]; //это доступ по названию листа
            }
            catch (Exception e)
            {
                MessageBox.Show(e + "Проверь название листа должно быть 'Заказы'");
            }
            _shtRange = _nwSheet.UsedRange;

            for (int Cnum = 1; Cnum <= _shtRange.Columns.Count; Cnum++) // в цикле получим все названия колонок, и положим их в качеситве назвния столбцов
            {                                                          // в Excel строки начинаются не с 0, а с 1, поэтому i = 1
                dataFromExcelDocs.Columns.Add
                    (new DataColumn((_shtRange.Cells[1, Cnum] as excelObj.Range).Value2.ToString()));

            }

            //Теперь дописываем остальные строки в цикле
            for (int Rnum = 2; Rnum <= _shtRange.Rows.Count; Rnum++)
            {


                DataRow dr = dataFromExcelDocs.NewRow();
                for (int Cnum = 1; Cnum <= _shtRange.Columns.Count; Cnum++)
                {

                    if ((_shtRange.Cells[Rnum, Cnum] as excelObj.Range).Value2 != null)
                    {
                        dr[Cnum - 1] =
                        (_shtRange.Cells[Rnum, Cnum] as excelObj.Range).Value2.ToString();
                    }
                }
                dataFromExcelDocs.Rows.Add(dr);
                dataFromExcelDocs.AcceptChanges();
            }

        }

        protected void GetIncomingExcelDataTableVehicle()
        {
            //Теперь дописываем с листа Машины           
            try
            {
                _nwSheet = (excelObj.Worksheet)_workbook.Worksheets["Авто"]; //это доступ по названию листа
            }
            catch (Exception e)
            {
                MessageBox.Show(e + "Проверь название листа, должно быть 'Авто'");
            }
            _shtRange = _nwSheet.UsedRange;

            for (int Cnum = 1; Cnum <= _shtRange.Columns.Count; Cnum++) // в цикле получим все названия колонок, и положим их в качеситве назвния столбцов
            {                                                          // в Excel строки начинаются не с 0, а с 1, поэтому i = 1
                dataFromExcelVehicle.Columns.Add
                    (new DataColumn((_shtRange.Cells[1, Cnum] as excelObj.Range).Value2));//.ToString())); 
            }
            for (int Rnum = 2; Rnum <= _shtRange.Rows.Count; Rnum++)
            {
                DataRow dr = dataFromExcelVehicle.NewRow();
                for (int Cnum = 1; Cnum <= _shtRange.Columns.Count; Cnum++)
                {
                    if ((_shtRange.Cells[Rnum, Cnum] as excelObj.Range).Value2 != null)
                    {
                        dr[Cnum - 1] =
                        (_shtRange.Cells[Rnum, Cnum] as excelObj.Range).Value2.ToString();
                    }
                }
                dataFromExcelVehicle.Rows.Add(dr);
                dataFromExcelVehicle.AcceptChanges();
            }
        }
                
        public void AppQuit(excelObj.Application app)
        {
            app.Quit();//закрыть Excel файл, что бы не висел в процессах
        }
    }
}