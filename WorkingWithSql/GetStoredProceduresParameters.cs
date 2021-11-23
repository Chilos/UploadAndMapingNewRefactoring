using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;


namespace UploadAndMapingNew
{
    //Получим параметрых хранимых процедур
    public sealed class GetStoredProceduresParameters
    {        
        public DataTable storedProcedureParams { get; set; } 

        public void GetParams(string storedProcedureName, SqlConnection conn)
        {
            string cmdText = @"select  
   'Parameter_name' = name,  
   'Type'   = type_name(user_type_id),  
   'Length'   = max_length,  
   'Prec'   = case when type_name(system_type_id) = 'uniqueidentifier' 
              then precision  
              else OdbcPrec(system_type_id, max_length, precision) end,  
   'Scale'   = OdbcScale(system_type_id, scale),  
   'Param_order'  = parameter_id,  
   'Collation'   = convert(sysname, 
                   case when system_type_id in (35, 99, 167, 175, 231, 239)  
                   then ServerProperty('collation') end)  
  from sys.parameters where object_id = object_id('" + storedProcedureName + "') ";

            using (SqlCommand sqlCommand = new SqlCommand(cmdText, conn))
            {
                sqlCommand.CommandType = CommandType.Text;
                SqlDataAdapter DataAdapter = new SqlDataAdapter(sqlCommand);
                DataSet dataSetwithStoredProceduresParams = new DataSet();
                DataAdapter.Fill(dataSetwithStoredProceduresParams);

                storedProcedureParams = dataSetwithStoredProceduresParams.Tables[0];
                dataSetwithStoredProceduresParams.Dispose();
            }
        }

    }
}
