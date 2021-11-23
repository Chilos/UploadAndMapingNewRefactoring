using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace UploadAndMapingNew
{
   sealed class CompilingStoredProcedureCall
    {
        private readonly DataTable _mapedFields = new DataTable();
        private readonly SqlConnection _connection;
        private readonly DataTable _tableWithDataFromExcel;
        private SqlCommand _sqlCommand;
        public List<SqlCommand> listOfProcedureCall { get; set; }

        public CompilingStoredProcedureCall(DataTable mapedFields, EstablishConnectionToSql esteblishedConection, DataTable tableWithDataFromExcel)
        {
            _mapedFields = mapedFields;
            _connection = esteblishedConection.connection;
            _tableWithDataFromExcel = tableWithDataFromExcel;
            listOfProcedureCall = new List<SqlCommand>();
        }

        /// <summary>
        /// attrId подаем либо номер атрибута для процедур, которые грузят атрибуты, либо 1 для проц, которые с атрибутами не работают. В таблице _mapedFields поле @AttrText будем мапить 
        /// с несколькими полями
        /// В _mapedFields поле с индексом [3] будем заполнять attrId того атрибута, с которым смапили AttrText
        /// На кажый атрибут будем создавать отдельный инстанс класа компиляции процедуры, но подавать туда один и тот же _mapedFields, а вот нужный AttrText будем отбирать по attrId
        /// </summary>        
        /// <param name="attrId"></param>
        public void compilingStoredProcedureCall(string storedProcedureName, int attrId)
        {
            string cdate = DateTime.Today.ToString();
            cdate = (cdate.Substring(0, cdate.Length - 8));
            string pathToLogWithConVerErr = @"d:\\YandexFiles\ErrOfConvering" + cdate + ".txt";
            WriteToLog WriteLogWithConverErr = new WriteToLog(pathToLogWithConVerErr);

            for (int rowindex = 0; rowindex < _tableWithDataFromExcel.Rows.Count; rowindex++)
            {
                _sqlCommand = new SqlCommand(storedProcedureName, _connection);
                _sqlCommand.CommandType = CommandType.StoredProcedure;

                CompilingDefaultPartOfProcedureCall(attrId);

                bool isExAttrValueIdEmpty = false;
                for (int cellInex = 0; cellInex < _tableWithDataFromExcel.Columns.Count; cellInex++)
                {
                    List<string[]> mapedParamMassList = FindeProcParamNameAndType(cellInex, attrId);
                    switch ((mapedParamMassList.Count != 0))
                    {
                        case true:
                            AddParamIntoSqlCommand(mapedParamMassList, rowindex, cellInex, WriteLogWithConverErr, storedProcedureName);
                            isExAttrValueIdEmpty = CheckIsExAttrValueIdIsEmpty(storedProcedureName, attrId, rowindex, cellInex);
                            break;
                        case false:
                            continue;
                    }
                }
                switch (isExAttrValueIdEmpty == true)
                {
                    case true:
                        continue;
                    case false:
                        listOfProcedureCall.Add(_sqlCommand);
                        break;
                }
            }

        }

        private void AddParamIntoSqlCommand(List<string[]> paramMassList, int rowindex, int cellInex, WriteToLog WriteLogWithConverErr, string storedProcedureName)
        {

            for (int indexParamMassList = 0; indexParamMassList < paramMassList.Count(); indexParamMassList++)
            {
                string[] paramMass = paramMassList[indexParamMassList];
                string paramType = paramMass[1];
                string procParamName = paramMass[0];

                switch (paramType)
                {
                    case "nvarchar": //секция 
                        try
                        {
                            switch (CheckIsFieldEmpty(rowindex, cellInex))
                            {
                                case true:
                                    _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = DBNull.Value;
                                    break;
                                default:
                                    if (procParamName.ToString().ToLower() == "@OtherFields".ToLower() && storedProcedureName.ToLower() == "DMT_Set_DocumentEx".ToLower())
                                    {
                                        string paramValueStr = "#######" + _tableWithDataFromExcel.Rows[rowindex][cellInex].ToString() + "######";

                                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = paramValueStr;
                                    }
                                    else if (procParamName.ToString().ToLower() == "@OtherFields".ToLower() && storedProcedureName.ToLower() == "DMT_set_AgentEx".ToLower())
                                    {
                                        string paramValueStr = "#5210005#" + "<831>:" + _tableWithDataFromExcel.Rows[rowindex][cellInex].ToString() + "##";
                                    }
                                    else
                                    {
                                        string paramValueStr = _tableWithDataFromExcel.Rows[rowindex][cellInex].ToString();
                                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = paramValueStr;
                                    }
                                    break;
                            }
                        }

                        catch (Exception)
                        {
                            string errMessage = @"Не порлучилост сконвертировать в строку в ячейке"
                            + cellInex + "в строке"
                            + rowindex
                            + "При сборке процедуры"
                            + storedProcedureName;
                            WriteLogWithConverErr.writeToLog(errMessage);
                        }
                        break;
                    case "int"://секция INT
                        try
                        {
                            switch (CheckIsFieldEmpty(rowindex, cellInex))
                            {     
                                case true:
                                    _sqlCommand.Parameters.Add(procParamName, SqlDbType.Int).Value = 0;
                                    break;
                                default:
                                    switch (_tableWithDataFromExcel.Rows[rowindex][cellInex].ToString().ToLower())
                                    {
                                        case "доставка":
                                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 2;
                                            break;
                                        case "возврат":
                                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 9;
                                            break;
                                        default:
                                            int paramValueInt = System.Convert.ToInt32(_tableWithDataFromExcel.Rows[rowindex][cellInex]);
                                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.Int).Value = paramValueInt;
                                            break;
                                    }
                                    break;
                            }
                            break;

                        }
                        catch (Exception)
                        {
                            string errMessage = @"Не порлучилост сконвертировать в INT (целое число) ячейке"
                            + cellInex
                            + "в строке"
                            + rowindex
                            + "При сборке процедуры"
                            + storedProcedureName;
                            WriteLogWithConverErr.writeToLog(errMessage);
                        }

                        break;
                    case "datetime":
                        try
                        {
                            if (CheckIsFieldEmpty(rowindex, cellInex) == true)
                            {
                                _sqlCommand.Parameters.Add(procParamName, SqlDbType.DateTime).Value = DBNull.Value;
                            }
                            DateTime paramValueDateTime = System.Convert.ToDateTime(_tableWithDataFromExcel.Rows[rowindex][cellInex]);
                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.DateTime).Value = paramValueDateTime;
                        }
                        catch (Exception)
                        {
                            string errMessage = @"Не порлучилост сконвертировать в Дату со временем в ячейке"
                            + cellInex
                            + "в строке"
                            + rowindex
                            + "При сборке процедуры"
                            + storedProcedureName;
                            WriteLogWithConverErr.writeToLog(errMessage);
                        }
                        break;
                    case "decimal":
                        try
                        {
                            if (CheckIsFieldEmpty(rowindex, cellInex) == true)
                            {
                                _sqlCommand.Parameters.Add(procParamName, SqlDbType.Decimal).Value = 0;
                            }
                            decimal paramValueDecimal = System.Convert.ToDecimal(_tableWithDataFromExcel.Rows[rowindex][cellInex]);
                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.Decimal).Value = paramValueDecimal;
                        }
                        catch (Exception)
                        {
                            string errMessage = @"Не порлучилост сконвертировать в число с лавающей точкой в ячейке"
                            + cellInex
                            + "в строке"
                            + rowindex
                            + "При сборке процедуры"
                            + storedProcedureName;
                            WriteLogWithConverErr.writeToLog(errMessage);
                        }
                        break;
                    case "мoney":
                        try
                        {
                            if (CheckIsFieldEmpty(rowindex, cellInex) == true)
                            {
                                _sqlCommand.Parameters.Add(procParamName, SqlDbType.Money).Value = 0;
                            }
                            decimal paramValueMoney = System.Convert.ToDecimal(_tableWithDataFromExcel.Rows[rowindex][cellInex]);
                            _sqlCommand.Parameters.Add(procParamName, SqlDbType.Decimal).Value = paramValueMoney;
                        }
                        catch (Exception)
                        {
                            string errMessage = @"Не порлучилост сконвертировать в число с лавающей точкой в ячейке"
                            + cellInex
                            + "в строке"
                            + rowindex
                            + "При сборке процедуры"
                            + storedProcedureName;
                            WriteLogWithConverErr.writeToLog(errMessage);
                        }
                        break;
                }

            }
        }

        private bool CheckIsFieldEmpty(int rowindex, int cellInex)
        {
            if (_tableWithDataFromExcel.Rows[rowindex][cellInex].ToString() == "" || _tableWithDataFromExcel.Rows[rowindex][cellInex].ToString() == null)
            { return true; }
            else
            { return false; }
        }
        private void CompilingDefaultPartOfProcedureCall(int attrId)
        {

            for (int rowIndex = 0; rowIndex < _mapedFields.Rows.Count; rowIndex++)
            {
                string procParamName = _mapedFields.Rows[rowIndex][1].ToString();

                if (_mapedFields.Rows[rowIndex][0].ToString() == "default")
                {
                    if (_mapedFields.Rows[rowIndex][1].ToString() == "@AttrID")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.Int).Value = attrId;
                        continue;
                    }
                    if ((_mapedFields.Rows[rowIndex][1].ToString() == "@fcomment" && _mapedFields.Rows[rowIndex][2].ToString() == "nvarchar") ||
                        (_mapedFields.Rows[rowIndex][1].ToString() == "@Comment" && _mapedFields.Rows[rowIndex][2].ToString() == "nvarchar"))
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = "Коментарий к точке";
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][1].ToString() == "@ExAttr" && _mapedFields.Rows[rowIndex][2].ToString() == "nvarchar")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = DBNull.Value;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][1].ToString() == "@ExAttrValue" && _mapedFields.Rows[rowIndex][2].ToString() == "nvarchar")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = DBNull.Value;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][1].ToString() == "@AttrValueID" && _mapedFields.Rows[rowIndex][2].ToString() == "nvarchar")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = DBNull.Value;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "nvarchar")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = DBNull.Value;
                        continue;
                    }
                    //Секция INT                    
                    if ((_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@activeFlag") ||
                        (_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@ActiveFlag"))
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 1;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@ftype")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 7;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@DictId")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 2;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@Sort")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 1;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "int" && _mapedFields.Rows[rowIndex][1].ToString() == "@CleanOtherVals")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.NVarChar).Value = 1;
                        continue;
                    }
                    if (_mapedFields.Rows[rowIndex][2].ToString() == "int")
                    {
                        _sqlCommand.Parameters.Add(procParamName, SqlDbType.Int).Value = 0;
                        continue;
                    }

                }
            }
        }
        private List<string[]> FindeProcParamNameAndType(int cellInex, int attrId)
        {
            List<string[]> listParamsMass = new List<string[]>();

            for (int mapTableIndex = 0; mapTableIndex < _mapedFields.Rows.Count; mapTableIndex++)

            {
                string excelFieldNameFromMapedFieldsTable = _mapedFields.Rows[mapTableIndex][0].ToString().ToLower();
                string excelFieldNameFromExcelDataTable = _tableWithDataFromExcel.Columns[cellInex].ToString().ToLower();
                string attrIdFromMapedFieldsTable = _mapedFields.Rows[mapTableIndex][3].ToString();
                string defaultAttrId = "0";//Значение по умолчанию, подается на вход в процедуру, когда собираем ХП НЕ для выгрузки атрибутов, а например для клиентов или документов

                if ((excelFieldNameFromMapedFieldsTable == excelFieldNameFromExcelDataTable && attrIdFromMapedFieldsTable == attrId.ToString())
                    || (excelFieldNameFromMapedFieldsTable == excelFieldNameFromExcelDataTable && attrIdFromMapedFieldsTable == defaultAttrId))
                {
                    string[] paramMass = new string[3];
                    paramMass[0] = _mapedFields.Rows[mapTableIndex][1].ToString();
                    paramMass[1] = _mapedFields.Rows[mapTableIndex][2].ToString();
                    paramMass[2] = _mapedFields.Rows[mapTableIndex][3].ToString();
                    listParamsMass.Add(paramMass);
                }
            }
            return listParamsMass;
        }

        //ХП при загрузке клинтов ТХЮ почему то не ставит им uflag = 1, приходиться апдэйтить
        public void UpdateUflagForClients()
        {
            string comandText = "  Update DS_Faces Set Uflag = 1 where fType = 7 and fActiveFlag = 1";
            SqlCommand CommandUpdateUflag = new SqlCommand(comandText, _connection);
            CommandUpdateUflag.CommandType = CommandType.Text;
            listOfProcedureCall.Add(CommandUpdateUflag);
        }

        public void SetPaymentTypeIfNotExist()
        {
            bool rowExists = CheckIsPaymentTypeExists();
            if (rowExists == false)
            {
                SqlCommand SetPaymentType = new SqlCommand("DMT_Set_PaymentTypeEx", _connection);
                SetPaymentType.CommandType = CommandType.StoredProcedure;
                SetPaymentType.Parameters.Add("@ExId", SqlDbType.NVarChar).Value = "PtDefault";
                SetPaymentType.Parameters.Add("@MarkUp", SqlDbType.Int).Value = 0;
                SetPaymentType.Parameters.Add("@Name", SqlDbType.NVarChar).Value = "PtDefault";
                SetPaymentType.Parameters.Add("@Comment", SqlDbType.NVarChar).Value = DBNull.Value;
                SetPaymentType.Parameters.Add("@ActiveFlag", SqlDbType.Int).Value = 1;
                SetPaymentType.Parameters.Add("@OtherFields", SqlDbType.NVarChar).Value = DBNull.Value;
                listOfProcedureCall.Add(SetPaymentType);
            }
        }

        public void SetPriceListIfNotExits()
        {
            bool rowExists = CheckIsPriceListExists();
            if (rowExists == false)
            {
                SqlCommand SetPriceList = new SqlCommand("DMT_Set_PriceLists", _connection);
                SetPriceList.CommandType = CommandType.StoredProcedure;
                SetPriceList.Parameters.Add("@exid", SqlDbType.NVarChar).Value = "PlDefault";
                SetPriceList.Parameters.Add("@name", SqlDbType.NVarChar).Value = "PlDefault";
                SetPriceList.Parameters.Add("@activeflag", SqlDbType.Int).Value = 1;
                listOfProcedureCall.Add(SetPriceList);
            }

        }

        public void SetStoreIfNotExists()
        {
            int cellIndesOfCoumnWIthStoreExidInExcel;
            for (int rowNumberMapedFieldsTable = 0; rowNumberMapedFieldsTable < _mapedFields.Rows.Count; rowNumberMapedFieldsTable++)
            {
                string storedProcedureParamName = _mapedFields.Rows[rowNumberMapedFieldsTable][1].ToString();
                if (storedProcedureParamName.ToLower() == "@otherfields".ToLower())
                {
                    string excelNameOfStoreExid = _mapedFields.Rows[rowNumberMapedFieldsTable][0].ToString();
                    cellIndesOfCoumnWIthStoreExidInExcel = _tableWithDataFromExcel.Columns.IndexOf(excelNameOfStoreExid);
                }
                else
                    continue;

                for (int rowNumberWithStoreExid = 0; rowNumberWithStoreExid < _tableWithDataFromExcel.Rows.Count; rowNumberWithStoreExid++)
                {

                    bool rowExists = CheckIsStoreActiveAndExists(cellIndesOfCoumnWIthStoreExidInExcel, rowNumberWithStoreExid);
                    if (rowExists == false)
                    {
                        string exid = _tableWithDataFromExcel.Rows[rowNumberWithStoreExid][cellIndesOfCoumnWIthStoreExidInExcel].ToString();
                        SqlCommand SetStore = new SqlCommand("DMT_set_Store", _connection);
                        SetStore.CommandType = CommandType.StoredProcedure;
                        SetStore.Parameters.Add("@ExId", SqlDbType.NVarChar).Value = exid;
                        SetStore.Parameters.Add("@ActiveFlag", SqlDbType.Int).Value = 1;
                        SetStore.Parameters.Add("@Name", SqlDbType.NVarChar).Value = "Склад" + exid;
                        SetStore.Parameters.Add("@ShortName", SqlDbType.NVarChar).Value = "Склад" + exid;
                        SetStore.Parameters.Add("@ServerExId", SqlDbType.NVarChar).Value = DBNull.Value;
                        SetStore.Parameters.Add("@AgentExId", SqlDbType.NVarChar).Value = DBNull.Value;
                        SetStore.Parameters.Add("@StoreTypeExId", SqlDbType.NVarChar).Value = DBNull.Value;
                        SetStore.Parameters.Add("@OwnerDistID", SqlDbType.Int).Value = DBNull.Value;

                        listOfProcedureCall.Add(SetStore);
                    }
                }

            }
        }


        private bool CheckIsStoreActiveAndExists(int cellIndesOfCoumnWIthStoreExidInExcel, int rowNumberWithStoreExid)
        {
            string comandText;
            string storeExId = _tableWithDataFromExcel.Rows[rowNumberWithStoreExid][cellIndesOfCoumnWIthStoreExidInExcel].ToString();
            comandText = "select * from DS_Faces where Ftype = 6 and factiveflag = 1 and OwnerDistId = dbo.Get_DistId() and exid ='" + storeExId + "'";

            SqlCommand CommandCheckIsStoreExists = new SqlCommand(comandText, _connection);
            CommandCheckIsStoreExists.CommandType = CommandType.Text;
            SqlDataReader recordset = CommandCheckIsStoreExists.ExecuteReader();
            bool rowExists = recordset.HasRows;
            recordset.Close();
            return rowExists;
        }

        private bool CheckIsPaymentTypeExists()
        {
            string comandText = "select * from DS_PaymentTypes where OwnerDistId = dbo.Get_DistId() and ActiveFlag = 1";
            SqlCommand CommandCheckIsExists = new SqlCommand(comandText, _connection);
            CommandCheckIsExists.CommandType = CommandType.Text;
            SqlDataReader recordset = CommandCheckIsExists.ExecuteReader();
            bool rowExists = recordset.HasRows;
            recordset.Close();
            return rowExists;
        }
        private bool CheckIsPriceListExists()
        {
            string comandText = "select * from DS_Pricelists where OwnerDistId = dbo.Get_DistId() and NotActive = 0";
            SqlCommand CommandCheckIsExists = new SqlCommand(comandText, _connection);
            CommandCheckIsExists.CommandType = CommandType.Text;
            SqlDataReader recordset = CommandCheckIsExists.ExecuteReader();
            bool rowExists = recordset.HasRows;
            recordset.Close();
            return rowExists;
        }

        public void compilingAttrValueSetStoredProcedureCall(int attrId)
        {
            int vehisleClassCellNumber = 0;
            for (int cellIndex = 0; cellIndex < _tableWithDataFromExcel.Columns.Count; cellIndex++)
            {
                if (_tableWithDataFromExcel.Columns[cellIndex].ColumnName.ToLower() == "Vehisle_Class_Docs".ToLower())
                {
                    vehisleClassCellNumber = cellIndex;
                    break;
                }
            }

            for (int rowNumber = 0; rowNumber < _tableWithDataFromExcel.Rows.Count; rowNumber++)
            {
                string nameOfClass = _tableWithDataFromExcel.Rows[rowNumber][vehisleClassCellNumber].ToString();
                if (nameOfClass != null && nameOfClass != "")
                {
                    SqlCommand SetAttrValue = new SqlCommand("DMT_Set_AttributeValueEx", _connection);
                    SetAttrValue.CommandType = CommandType.StoredProcedure;
                    SetAttrValue.Parameters.Add("@AttrID", SqlDbType.Int).Value = attrId;
                    SetAttrValue.Parameters.Add("@ExAttrId", SqlDbType.NVarChar).Value = DBNull.Value;
                    SetAttrValue.Parameters.Add("@AttrValueID", SqlDbType.Int).Value = DBNull.Value;
                    SetAttrValue.Parameters.Add("@AttrValueName", SqlDbType.NVarChar).Value = nameOfClass;
                    SetAttrValue.Parameters.Add("@ExAttrValueId", SqlDbType.NVarChar).Value = nameOfClass;
                    SetAttrValue.Parameters.Add("@AttrValueSystemFlag", SqlDbType.Int).Value = 0;
                    SetAttrValue.Parameters.Add("@ActiveFlag", SqlDbType.Int).Value = 1;
                    SetAttrValue.Parameters.Add("@OtherFields", SqlDbType.NVarChar).Value = DBNull.Value;
                    listOfProcedureCall.Add(SetAttrValue);
                }
            }
        }
        //Этот метод нужен, что бы проверять на пустое значение для атрибута 813, это перечислимый атрибут и в DMT_Set_DocAttributeEx нельзя передавать параметр @ExAttrValueId = null, поэтому в случае пустоты, ВООБЩЕ не надо добавлять этот вызов в лист вызова ХП
        //Ну точнее можно @ExAttrValueId = null, но при AttrValueId !=0, а у меня этот параметр всегда равен 0.
        private bool CheckIsExAttrValueIdIsEmpty(string storedProcedureName, int attrId, int rowindex, int cellIndex)
        {
            if ((storedProcedureName.ToLower() == "DMT_Set_DocAttributeEx".ToLower() && _tableWithDataFromExcel.Rows[rowindex][cellIndex].ToString() == null && attrId == 831) || 
                (storedProcedureName.ToLower() == "DMT_Set_DocAttributeEx".ToLower() && _tableWithDataFromExcel.Rows[rowindex][cellIndex].ToString() == "" && attrId == 831))
                return true;
            else
                return false;
        }
    }
}
