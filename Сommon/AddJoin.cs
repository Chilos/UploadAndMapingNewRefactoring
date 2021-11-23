using System.Data;
using System.Windows.Forms;

namespace UploadAndMapingNew
{
    sealed class AddJoin
    {
        int counterForRowsInDataGridView;

        public AddJoin()
        {
            counterForRowsInDataGridView = 0;
        }

        //TODO МАРКЕР Здесь хардкод на то какие атрибуты берем в зависимости от названия полей экселя
        public DataTable AddJoinParams(DataGridView excelFields, DataGridView storedProceduresParams, DataGridView mapedFields)
        {
            DataTable dataTableMapedFields = (DataTable)mapedFields.DataSource;

            switch (storedProceduresParams.CurrentRow.Cells[0].Value.ToString().ToLower())
            {
                case "@otherfields":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    int rowIndexExcelFields = excelFields.CurrentRow.Index;
                    excelFields.Rows.RemoveAt(rowIndexExcelFields);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;

                case "@attrtext":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    switch (excelFields.CurrentCell.Value.ToString().ToLower())
                    {
                        case "lat":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_LAT;
                            break;
                        case "lon":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_LON;
                            break;
                        case "service_time":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_SERVICE_TIME_ON_OUTLET; //Время обслуживания на точке AttrId = 364
                            break;
                        case "veit":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_ORDER_VEIT; //Вес накладной  AttrId = 1486
                            break;
                        case "volium":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_ORDER_VOLIUM; //Объем накладной накладной  AttrId = 1487
                            break;
                        case "vehisle_class":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_VEHISLE_CLASS; //Класс машины  AttrId = 831
                            break;
                        case "delivery_window_start":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DELIVERY_TIME_START; // AttrId = 383
                            break;
                        case "delivery_window_end":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DELIVERY_TIME_END; // AttrId = 384
                            break;
                        case "vehisle_volum":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_BODY_VOLUME; // Объем кузова
                            break;
                        case "load_capacity":
                            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_LOAD_CAPACITY_KG; // Грузоподъемность в кг
                            break;
                    }
                    rowIndexExcelFields = excelFields.CurrentRow.Index;
                    excelFields.Rows.RemoveAt(rowIndexExcelFields);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;

                case "@exid":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    int rowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
                    storedProceduresParams.Rows.RemoveAt(rowIndexStoredProcedureParams);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;
                case "@docidd":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    rowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
                    storedProceduresParams.Rows.RemoveAt(rowIndexStoredProcedureParams);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;
                case "@docdate":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    rowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
                    storedProceduresParams.Rows.RemoveAt(rowIndexStoredProcedureParams);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;
                case "@name":
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    rowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
                    storedProceduresParams.Rows.RemoveAt(rowIndexStoredProcedureParams);
                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;
                default:
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = excelFields.CurrentCell.Value;//Возьмем текущуюю ячейку из грида с заголовками Эксель и
                                                                                                               //присвоим ее нулевой яейке в строке с номером counterRows
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;//Возьмем текущую ячейку из грида
                                                                                                                                  //с параметрами процдуры  и прихнем ее в
                                                                                                                                  //1ую ячейку строки с номером counterRows
                                                                                                                                  //Удалим значения из строк, которые смапили.
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
                    dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;
                    rowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
                    storedProceduresParams.Rows.RemoveAt(rowIndexStoredProcedureParams);
                    rowIndexExcelFields = excelFields.CurrentRow.Index;
                    excelFields.Rows.RemoveAt(rowIndexExcelFields);

                    counterForRowsInDataGridView++;
                    return dataTableMapedFields;

            }
        }

        public DataTable AddjoinDefaulValue(DataGridView storedProceduresParams, DataGridView mapedFields)
        {
            DataTable dataTableMapedFields = (DataTable)mapedFields.DataSource;
            dataTableMapedFields.Rows[counterForRowsInDataGridView][0] = "default";
            dataTableMapedFields.Rows[counterForRowsInDataGridView][1] = storedProceduresParams.CurrentRow.Cells[0].Value;
            dataTableMapedFields.Rows[counterForRowsInDataGridView][2] = storedProceduresParams.CurrentRow.Cells[1].Value;
            dataTableMapedFields.Rows[counterForRowsInDataGridView][3] = ConstantsAttributesId.ATTRID_DEFAULT;

            int RowIndexStoredProcedureParams = storedProceduresParams.CurrentRow.Index;
            storedProceduresParams.Rows.RemoveAt(RowIndexStoredProcedureParams);

            counterForRowsInDataGridView++;

            return dataTableMapedFields;
        }
               
    }
}
