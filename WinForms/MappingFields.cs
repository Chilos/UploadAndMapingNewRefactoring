using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace UploadAndMapingNew
{
    public partial class MappingFields : Form
    {
        private GetStoredProceduresParameters setClientExParametrs { get; set; }
        private GetStoredProceduresParameters setObjectsAttributeExParametrs { get; set; }
        private GetStoredProceduresParameters setDocumentExParametrs { get; set; }
        private GetStoredProceduresParameters setDocAttributeParametrs { get; set; }
        private GetStoredProceduresParameters setAgentExParametrs { get; set; }
        private GetStoredProceduresParameters setObjectsAttributes { get; set; }

        private UploadNameOfFieldsFromExcel nameOfExcelFieldsClients { get; set; }
        private UploadNameOfFieldsFromExcel nameOfExcelFieldsClientsAttributes { get; set; }
        private UploadNameOfFieldsFromExcel nameOfExcelFieldsDocuments { get; set; }
        private UploadNameOfFieldsFromExcel nameOfExcelFieldsDocumetsAttributes { get; set; }
        private UploadNameOfFieldsFromExcel nameOfExcelFieldsVihicle { get; set; }
        private UploadNameOfFieldsFromExcel nameOfExcelFieldsVihicleAttributes { get; set; }

        private DataTable tableWithMapedFieldsForCients { get; set; }
        private DataTable tableWithMapedFieldsForCientsAttributes { get; set; }
        private DataTable tableWithMapedFieldsForDocument { get; set; }
        private DataTable tableWithMapedFieldsForDocumentAttributes { get; set; }
        private DataTable tableWithMapedFieldsForVihicle { get; set; }
        private DataTable tableWithMapedFieldsForVihicleAttributes { get; set; }

        LoadingDataFromExcel dataFromExcel;

        public MappingFields(List<GetStoredProceduresParameters> listOfStoredProcedureParams, UploadNameOfFieldsFromExcel nameOfExcelFields, LoadingDataFromExcel DataFromExcel)
        {
            InitializeComponent();
            setClientExParametrs = listOfStoredProcedureParams[0];
            setObjectsAttributeExParametrs = listOfStoredProcedureParams[1];
            setDocumentExParametrs = listOfStoredProcedureParams[2];
            setDocAttributeParametrs = listOfStoredProcedureParams[3];
            setAgentExParametrs = listOfStoredProcedureParams[4];
            setObjectsAttributes = listOfStoredProcedureParams[5];
            nameOfExcelFieldsClients = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            nameOfExcelFieldsClientsAttributes = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            nameOfExcelFieldsDocuments = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            nameOfExcelFieldsDocumetsAttributes = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            nameOfExcelFieldsVihicle = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            nameOfExcelFieldsVihicleAttributes = (UploadNameOfFieldsFromExcel)nameOfExcelFields.Clone();
            tableWithMapedFieldsForCients = new DataTable();
            tableWithMapedFieldsForCientsAttributes = new DataTable();
            tableWithMapedFieldsForDocument = new DataTable();
            tableWithMapedFieldsForDocumentAttributes = new DataTable();
            tableWithMapedFieldsForVihicle = new DataTable();
            tableWithMapedFieldsForVihicleAttributes = new DataTable();
            this.dataFromExcel = DataFromExcel;
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForCients);
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForCientsAttributes);
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForDocument);
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForDocumentAttributes);
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForVihicle);
            SetFIeldForTableWithMapedFields(tableWithMapedFieldsForVihicleAttributes);
        }

        public MappingFields()
        {
            InitializeComponent();
        }

        //Поскольку мы используем DataGridView в WinForms то надо каждому DataGridView проставить источник данных DataTable откуда брать  данные
        private void MappingFields_Load_1(object sender, EventArgs e)
        {
            ExcelParamsForClientDataGridView.AutoGenerateColumns = true;
            ExcelParamsForClientDataGridView.DataSource = nameOfExcelFieldsClients.paramsFromExcelDocs;
            ExcelParamsForClientAttributesdataGridView.AutoGenerateColumns = true;
            ExcelParamsForClientAttributesdataGridView.DataSource = nameOfExcelFieldsClientsAttributes.paramsFromExcelDocs;
            ExcelParamsForDocumentesdataGridView.AutoGenerateColumns = true;
            ExcelParamsForDocumentesdataGridView.DataSource = nameOfExcelFieldsDocuments.paramsFromExcelDocs;
            ExcelParamsForDocumentAttributesDataGridView.AutoGenerateColumns = true;
            ExcelParamsForDocumentAttributesDataGridView.DataSource = nameOfExcelFieldsDocumetsAttributes.paramsFromExcelDocs;
            ExcelParamsForVihicleDataGridView.AutoGenerateColumns = true;
            ExcelParamsForVihicleDataGridView.DataSource = nameOfExcelFieldsVihicle.paramsFromExcelVehisle;
            ExcelParamsForVehisleAttributesdataGridView.AutoGenerateColumns = true;
            ExcelParamsForVehisleAttributesdataGridView.DataSource = nameOfExcelFieldsVihicleAttributes.paramsFromExcelVehisle;

            ClientStoredProceduresParamsDataGridView.AutoGenerateColumns = true;
            ClientStoredProceduresParamsDataGridView.DataSource = setClientExParametrs.storedProcedureParams;
            ClientAttrIbutesStoredProceduresParamsDataGridView.AutoGenerateColumns = true;
            ClientAttrIbutesStoredProceduresParamsDataGridView.DataSource = setObjectsAttributeExParametrs.storedProcedureParams;
            DocumentesStoredProceduresParamsDataGridView.AutoGenerateColumns = true;
            DocumentesStoredProceduresParamsDataGridView.DataSource = setDocumentExParametrs.storedProcedureParams;
            DocumentAttributesStoredProcedureParamsDataGridView.AutoGenerateColumns = true;
            DocumentAttributesStoredProcedureParamsDataGridView.DataSource = setDocAttributeParametrs.storedProcedureParams;
            VihicleStoredProcedureParamsDataGridView.AutoGenerateColumns = true;
            VihicleStoredProcedureParamsDataGridView.DataSource = setAgentExParametrs.storedProcedureParams;
            VehisleAttributesStoredProcedureParamsdataGridView.AutoGenerateColumns = true;
            VehisleAttributesStoredProcedureParamsdataGridView.DataSource = setObjectsAttributes.storedProcedureParams;

            MapedFieldsForClientsDataGridView.AutoGenerateColumns = true;
            MapedFieldsForClientsDataGridView.DataSource = tableWithMapedFieldsForCients;
            MapedFieldsForClientsAttributesDataGridView.AutoGenerateColumns = true;
            MapedFieldsForClientsAttributesDataGridView.DataSource = tableWithMapedFieldsForCientsAttributes;
            MapedFieldsForDocumentDataGridView.DataSource = tableWithMapedFieldsForDocument;
            MapedFielsDocumetAttrbutesDataGridView.AutoGenerateColumns = true;
            MapedFielsDocumetAttrbutesDataGridView.DataSource = tableWithMapedFieldsForDocumentAttributes;
            MapedFieldsVihicleDataGridView.AutoGenerateColumns = true;
            MapedFieldsVihicleDataGridView.DataSource = tableWithMapedFieldsForVihicle;
            MapedFieldsForVehisleAttributesdataGridView.AutoGenerateColumns = true;
            MapedFieldsForVehisleAttributesdataGridView.DataSource = tableWithMapedFieldsForVihicleAttributes;
        }



        AddJoin addJoinClient = new AddJoin();
        private void AddJoinClient_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForCients.Rows.Add();

            tableWithMapedFieldsForCients = addJoinClient.AddJoinParams(
                ExcelParamsForClientDataGridView
                , ClientStoredProceduresParamsDataGridView
                , MapedFieldsForClientsDataGridView);
        }

        private void AddDefault_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForCients.Rows.Add();
            tableWithMapedFieldsForCients = addJoinClient.AddjoinDefaulValue(
                ClientStoredProceduresParamsDataGridView
                , MapedFieldsForClientsDataGridView);
        }

        AddJoin addJoinClientAttributes = new AddJoin();
        private void AddJoinCientAttributes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForCientsAttributes.Rows.Add();
            tableWithMapedFieldsForCientsAttributes = addJoinClientAttributes.AddJoinParams(
                ExcelParamsForClientAttributesdataGridView
                , ClientAttrIbutesStoredProceduresParamsDataGridView
                , MapedFieldsForClientsAttributesDataGridView);
        }
        private void AddDefaultForClientAttr_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForCientsAttributes.Rows.Add();
            tableWithMapedFieldsForCientsAttributes = addJoinClientAttributes.AddjoinDefaulValue(
                ClientAttrIbutesStoredProceduresParamsDataGridView
                , MapedFieldsForClientsAttributesDataGridView);
        }

        AddJoin addJoinDocumentes = new AddJoin();
        private void AddJoinDocumentes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForDocument.Rows.Add();
            tableWithMapedFieldsForDocument = addJoinDocumentes.AddJoinParams(
                ExcelParamsForDocumentesdataGridView
                , DocumentesStoredProceduresParamsDataGridView
                , MapedFieldsForDocumentDataGridView);
        }
        private void AddDefaultForDocumentes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForDocument.Rows.Add();
            tableWithMapedFieldsForDocument = addJoinDocumentes.AddjoinDefaulValue(
                DocumentesStoredProceduresParamsDataGridView
                , MapedFieldsForDocumentDataGridView);
        }

        AddJoin addJoinDocumentAttributes = new AddJoin();
        private void AddJoinDocumentAttributes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForDocumentAttributes.Rows.Add();
            tableWithMapedFieldsForDocumentAttributes = addJoinDocumentAttributes.AddJoinParams(
                ExcelParamsForDocumentAttributesDataGridView
                , DocumentAttributesStoredProcedureParamsDataGridView
                , MapedFielsDocumetAttrbutesDataGridView);
        }
        private void AddDefaultForDocumentesAttributes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForDocumentAttributes.Rows.Add();
            tableWithMapedFieldsForDocumentAttributes = addJoinDocumentAttributes.AddjoinDefaulValue(
                
                DocumentAttributesStoredProcedureParamsDataGridView
                , MapedFielsDocumetAttrbutesDataGridView);
        }

        AddJoin addJoinVihicle = new AddJoin();
        private void AddJoinVihicle_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForVihicle.Rows.Add();
            tableWithMapedFieldsForVihicle = addJoinVihicle.AddJoinParams(
                ExcelParamsForVihicleDataGridView
                , VihicleStoredProcedureParamsDataGridView
                , MapedFieldsVihicleDataGridView);
        }

        private void AddDefaultForDocumentesVihicle_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForVihicle.Rows.Add();
            tableWithMapedFieldsForVihicle = addJoinVihicle.AddjoinDefaulValue(
                VihicleStoredProcedureParamsDataGridView
                , MapedFieldsVihicleDataGridView);
        }

        AddJoin AddJoinForVehisleAttributes = new AddJoin();
        private void AddJoinVehisleAttributes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForVihicleAttributes.Rows.Add();
            tableWithMapedFieldsForVihicleAttributes = AddJoinForVehisleAttributes.AddJoinParams(
                ExcelParamsForVehisleAttributesdataGridView
                , VehisleAttributesStoredProcedureParamsdataGridView
                , MapedFieldsForVehisleAttributesdataGridView);
        }

        private void AddDefaultForVehisleAttributes_Click(object sender, EventArgs e)
        {
            tableWithMapedFieldsForVihicleAttributes.Rows.Add();
            tableWithMapedFieldsForVihicleAttributes = AddJoinForVehisleAttributes.AddjoinDefaulValue(
                VehisleAttributesStoredProcedureParamsdataGridView
                , MapedFieldsForVehisleAttributesdataGridView);
        }

        private void StartUpload_Click_1(object sender, EventArgs e)
        {
            EstablishConnectionToSql conectToload = new EstablishConnectionToSql("cdcs3", "Mobile_SupG_Avtoexpress", "alexis", "admin");
            conectToload.GetConnection();

            CompilingStoredProcedureCall compilingSetClientEx = new CompilingStoredProcedureCall(tableWithMapedFieldsForCients, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetFacesAttributeEx360 = new CompilingStoredProcedureCall(tableWithMapedFieldsForCientsAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetFacesAttributeEx361 = new CompilingStoredProcedureCall(tableWithMapedFieldsForCientsAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocumente = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocument, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocAttribute1486 = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocumentAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocAttribute1487 = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocumentAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocAttribute831 = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocumentAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocAttribute383 = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocumentAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetDocAttribute384 = new CompilingStoredProcedureCall(tableWithMapedFieldsForDocumentAttributes, conectToload, dataFromExcel.dataFromExcelDocs);
            CompilingStoredProcedureCall compilingSetAgent = new CompilingStoredProcedureCall(tableWithMapedFieldsForVihicle, conectToload, dataFromExcel.dataFromExcelVehicle);
                        
            compilingSetClientEx.compilingStoredProcedureCall("DMT_Set_ClientEx", ConstantsAttributesId.ATTRID_NOT_NEED_FOR_STORED_PROCEDURE);
            compilingSetClientEx.UpdateUflagForClients();
            compilingSetClientEx.SetPaymentTypeIfNotExist();
            compilingSetClientEx.SetPriceListIfNotExits();
            //compilingSetClientEx.SetStoreIfNotExists();
            compilingSetClientEx.compilingAttrValueSetStoredProcedureCall(ConstantsAttributesId.ATTRID_VEHISLE_CLASS);
            compilingSetFacesAttributeEx360.compilingStoredProcedureCall("DMT_Set_FacesAttributeEx", ConstantsAttributesId.ATTRID_LAT);
            compilingSetFacesAttributeEx361.compilingStoredProcedureCall("DMT_Set_FacesAttributeEx", ConstantsAttributesId.ATTRID_LON);
            compilingSetDocumente.SetStoreIfNotExists();
            compilingSetDocumente.compilingStoredProcedureCall("DMT_Set_DocumentEx", ConstantsAttributesId.ATTRID_NOT_NEED_FOR_STORED_PROCEDURE);
            compilingSetDocAttribute1486.compilingStoredProcedureCall("DMT_Set_DocAttributeEx", ConstantsAttributesId.ATTRID_ORDER_VEIT);
            compilingSetDocAttribute1487.compilingStoredProcedureCall("DMT_Set_DocAttributeEx", ConstantsAttributesId.ATTRID_ORDER_VOLIUM);
            compilingSetDocAttribute831.compilingStoredProcedureCall("DMT_Set_DocAttributeEx", ConstantsAttributesId.ATTRID_VEHISLE_CLASS);
            compilingSetDocAttribute383.compilingStoredProcedureCall("DMT_Set_DocAttributeEx", ConstantsAttributesId.ATTRID_DELIVERY_TIME_START);
            compilingSetDocAttribute384.compilingStoredProcedureCall("DMT_Set_DocAttributeEx", ConstantsAttributesId.ATTRID_DELIVERY_TIME_END);
            compilingSetAgent.compilingStoredProcedureCall("DMT_set_AgentEx", ConstantsAttributesId.ATTRID_NOT_NEED_FOR_STORED_PROCEDURE);

            ExecuteStoredProcedures ExecSetClient = new ExecuteStoredProcedures(compilingSetClientEx.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetFacesAttributeEx360 = new ExecuteStoredProcedures(compilingSetFacesAttributeEx360.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetFacesAttributeEx361 = new ExecuteStoredProcedures(compilingSetFacesAttributeEx361.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocumente = new ExecuteStoredProcedures(compilingSetDocumente.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocAttributes1486 = new ExecuteStoredProcedures(compilingSetDocAttribute1486.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocAttributes1487 = new ExecuteStoredProcedures(compilingSetDocAttribute1487.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocAttributes831 = new ExecuteStoredProcedures(compilingSetDocAttribute831.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocAttributes383 = new ExecuteStoredProcedures(compilingSetDocAttribute383.listOfProcedureCall);
            ExecuteStoredProcedures ExecSetDocAttributes384 = new ExecuteStoredProcedures(compilingSetDocAttribute384.listOfProcedureCall);

            int numberOfClientsLoaded = ExecSetClient.StoredProcedureExecuter();
            int numberOfAttr360 = ExecSetFacesAttributeEx360.StoredProcedureExecuter();
            int numberOfAttr361 = ExecSetFacesAttributeEx361.StoredProcedureExecuter();
            int numberOfDocuments = ExecSetDocumente.StoredProcedureExecuter();
            //  TODO Проверить соответсвия кол-ва документов, клиентов и атрибутов, ну имашин
            ExecSetDocAttributes1486.StoredProcedureExecuter();
            ExecSetDocAttributes1487.StoredProcedureExecuter();
            ExecSetDocAttributes831.StoredProcedureExecuter();
            ExecSetDocAttributes383.StoredProcedureExecuter();
            ExecSetDocAttributes384.StoredProcedureExecuter();

            conectToload.CloseConntion();

            MessageBox.Show("Клиентов загружено - " 
                + numberOfClientsLoaded + "Атрибутов 360 загружено " 
                + numberOfAttr360 + "АТрибутов 361 загружено " 
                + numberOfAttr361 + "Документов загружено " 
                + numberOfDocuments);

        }
        private void SetFIeldForTableWithMapedFields(DataTable TableName)
        {
            TableName.Columns.Add(new DataColumn("Параметры Excel"));
            TableName.Columns.Add(new DataColumn("Параметры хранимых процедур"));
            TableName.Columns.Add(new DataColumn("DataType"));
            TableName.Columns.Add(new DataColumn("AttrId"));
        }

        private void Client_Click(object sender, EventArgs e)
        {

        }
        private void ClientAttributs_Click(object sender, EventArgs e)
        {

        }
        private void Documents_Click(object sender, EventArgs e)
        {

        }
        private void DocAttributes_Click(object sender, EventArgs e)
        {

        }
        private void Vehicle_Click(object sender, EventArgs e)
        {

        }

        private void ClientStoredProceduresParams_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ExcelParams_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void MapedFieldsDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ExcelParamsForClientAttributesdataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ClientAttrIbutesStoredProceduresParamsDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void MapedFieldsForClientsAttributesDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ExcelParamsForDocumentesdataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void CleareJoin_Click(object sender, EventArgs e)
        {

        }

        private void ExcelParamsForAgentDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Vehisle_Attributes_Click(object sender, EventArgs e)
        {

        }
    }
}
