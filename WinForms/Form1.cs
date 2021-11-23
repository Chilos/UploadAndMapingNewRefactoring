using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace UploadAndMapingNew
{
    public partial class Form1 : Form
    {
        private OpenFileDialog openExcelFile;
        private List<GetStoredProceduresParameters> listOfStoredProcedureParams;
        public Form1()
        {
            InitializeComponent();
            textBox5.Text = "";
            textBox4.Text = "Mobile_SupG_Avtoexpress";
            textBox2.Text = "";
            textBox3.Text = "";
            listOfStoredProcedureParams = new List<GetStoredProceduresParameters>();

        }

        private void button1_Click(object sender, EventArgs e)
        {                        
            MessageBox.Show(@"ПРОВЕРЬ ЗАГРУЖАЕМЫЙ ФАЙЛ ДЛЯ ДОСТАВКИ:
                                   Названия листов:
                    - Лист с накладными и Клиентами - Заказы
                    - Лист с машинами - Авто
                                   Корректность данных:
                    1) Названия колонок должны идти первой строкой
                    2) НАД названием колонок НЕ должно присутсвовать пустых или заполненых строк
                    3) Названия колонок НЕ должны повторяться
                    4) Название колонки должно быть одной строчкой, НЕ должно быть группировок
                    5) Названия колонок Vehisle_Class (класс машины), Delivery_window_Start (начало окна доставки), Delivery_window_End(конец кна достаки),Volium (объем накладной), Veit (вес накладной), service_time (время на обслуживание на ТТ), Lat, Lon  должны быть именно такими 
                    5) Поля Delivery_window_Start и Delivery_window_End должны иметь СТРОГО текстовый формат");

            openExcelFile = new OpenFileDialog();
            openExcelFile.DefaultExt = "*.xls; *.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты, доступные в поле "Файлы типа" диалогового окна.
            openExcelFile.Filter = "Excel 2016(*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls";
            if (openExcelFile.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openExcelFile.FileName; 
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            EstablishConnectionToSql connect = new EstablishConnectionToSql(textBox5.Text, textBox4.Text, textBox3.Text, textBox2.Text);
            connect.GetConnection();
            
            GetStoredProceduresParameters setClientExParametrs = new GetStoredProceduresParameters();
            GetStoredProceduresParameters setObjectsAttributeExParametrs = new GetStoredProceduresParameters();
            GetStoredProceduresParameters setDocumentExParametrs = new GetStoredProceduresParameters();
            GetStoredProceduresParameters setDocAttributeParametrs = new GetStoredProceduresParameters();
            GetStoredProceduresParameters setAgentExParametrs = new GetStoredProceduresParameters();
            GetStoredProceduresParameters setObjectsAttributes = new GetStoredProceduresParameters();
            listOfStoredProcedureParams.Add(setClientExParametrs);
            listOfStoredProcedureParams.Add(setObjectsAttributeExParametrs);
            listOfStoredProcedureParams.Add(setDocumentExParametrs);
            listOfStoredProcedureParams.Add(setDocAttributeParametrs);
            listOfStoredProcedureParams.Add(setAgentExParametrs);
            listOfStoredProcedureParams.Add(setObjectsAttributes);
            setClientExParametrs.GetParams("DMT_Set_ClientEx", connect.connection);            
            setObjectsAttributeExParametrs.GetParams("DMT_Set_FacesAttributeEx", connect.connection);            
            setDocumentExParametrs.GetParams("DMT_Set_DocumentEx", connect.connection);            
            setDocAttributeParametrs.GetParams("DMT_Set_DocAttributeEx", connect.connection);            
            setAgentExParametrs.GetParams("DMT_set_AgentEx", connect.connection);            
            setObjectsAttributes.GetParams("DMT_Set_ObjectsAttribute", connect.connection);

            LoadingDataFromExcel dataFromExcel = new LoadingDataFromExcel();
            dataFromExcel.GetDataFromExcel(openExcelFile);            
            UploadNameOfFieldsFromExcel nameOfExcelFields = new UploadNameOfFieldsFromExcel();
            nameOfExcelFields.GetFieldNamesFromExcel(openExcelFile);

            MappingFields Mapping = new MappingFields(listOfStoredProcedureParams, nameOfExcelFields, dataFromExcel);

            connect.CloseConntion();
            Mapping.Show();            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
                           
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
                         
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
                         
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
                         
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
