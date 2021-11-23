using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace UploadAndMapingNew
{
    sealed class ExecuteStoredProcedures
    {
        private readonly List<SqlCommand> _listOfProcerdureCall;

        public ExecuteStoredProcedures(List<SqlCommand> ListOfProcerdureCall)
        {
            _listOfProcerdureCall = ListOfProcerdureCall;
        }
        //TODO У меня вылезает сообщение, о том что таймаут превышен "Помощник по отладке управляемого кода "ContextSwitchDeadlock" : "CLR не удалось перейти из COM-контекста 0x150b1d8 в COM-контекст 0x150b120 за 60 секунд. ", можно как-то сделать что бы оно не лезло, а продолжалос выполнение автоматичеки? 
        //TODO добавить что бы каждая проца вызывалась в своем потоке, пока пул потоков не кончиться. А как кончится не создавать новые, а ждать пока освободятся меющиеся.        

        public int StoredProcedureExecuter()
        {
            int numberOfExecutedStoredProcedures = 0;
            int errCount = 0;
            string cDate = DateTime.Today.ToString();
            cDate = (cDate.Substring(0,cDate.Length - 8));
            string pathToLogWithConVerErr = @"d:\\YandexFiles\ErrOfExecuting"+cDate+".txt";
            WriteToLog LogWithExecErr = new WriteToLog(pathToLogWithConVerErr);
            for (int indexOfProcedureCallList = 0; indexOfProcedureCallList < _listOfProcerdureCall.Count; indexOfProcedureCallList++)
            {
                try
                {
                    int numberOfExecutedStoredProcerures = _listOfProcerdureCall[indexOfProcedureCallList].ExecuteNonQuery();                    
                    numberOfExecutedStoredProcedures++;
                }
                catch (Exception e)
                {
                    errCount++;
                    LogWithExecErr.writeToLog(e.ToString());
                    string paramName;
                    string paramValue;
                    string unitedParaAndValue = null;
                    for (int ParamIndex = 0; ParamIndex < _listOfProcerdureCall[indexOfProcedureCallList].Parameters.Count; ParamIndex++)
                    {
                        paramName = _listOfProcerdureCall[indexOfProcedureCallList].Parameters[ParamIndex].ParameterName.ToString();
                        paramValue = _listOfProcerdureCall[indexOfProcedureCallList].Parameters[ParamIndex].Value.ToString();
                        unitedParaAndValue +=(paramName + "=" + paramValue);
                    }
                    LogWithExecErr.writeToLog(_listOfProcerdureCall[indexOfProcedureCallList].CommandText + unitedParaAndValue);
                    LogWithExecErr.writeToLog("--------------------------------------------------");
                }
            }
            if (errCount > 0) { MessageBox.Show("Возникло " + errCount + " ошибок, смотри лог"); }
            return numberOfExecutedStoredProcedures;
        }

    }
}
