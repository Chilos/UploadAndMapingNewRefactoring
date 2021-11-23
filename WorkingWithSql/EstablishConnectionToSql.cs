using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace UploadAndMapingNew
{
    sealed class EstablishConnectionToSql 
    {
        private readonly string _dataSourse; //сервак
        private readonly string _initalCatalog; //БД
        private readonly string _password; //Пароль
        private readonly string _login; //Логин
        private  SqlConnection _connection;
        public SqlConnection connection { get { return _connection; } set { _connection = value; } }

        internal EstablishConnectionToSql(string dataSourse, string initalCatalog, string password, string login)
        {
            //TODO Если убрать отсюда this, то как правильно назвать входящие параметры конструктора? 
            _dataSourse = dataSourse;
            _initalCatalog = initalCatalog;
            _password = password;
            _login = login;
        }

        public void GetConnection()
        {
            string connectionString = @"Data Source = " + _dataSourse + "; Database = " + _initalCatalog + "; User Id = " + _login + "; Password = " + _password;
            connection = new SqlConnection(connectionString);
            
            try
            {
                connection.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка:" + e.Message); //TODO переделать на события, что бы форма по событию кидала MessageBox. Потому что по феншую слой работы с источниками данных не должен ничего знать о UI
            }

        }

        public void CloseConntion()
        {
            connection.Close();
            connection.Dispose();
        }

    }
}
