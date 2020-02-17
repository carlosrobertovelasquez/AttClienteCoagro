using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
namespace CoagroEnvioFacturas
{
    public abstract class Conexion
    {
        private readonly string connectionString;

        public Conexion()
        {
            connectionString = ConfigurationManager.ConnectionStrings["sConexion"].ToString();
        }
        protected SqlConnection GetConnection()
        {
            return new SqlConnection(connectionString);
        }


    }
}
