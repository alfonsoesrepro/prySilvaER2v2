using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace prySilvaER2v2.BaseDatos
{
    public class CConexion
    {
        // Propiedades públicas para ser accedidas por otras clases del proyecto
        public OleDbConnection CNN;
        public DataSet DS;

        // Propiedad privada para la gestión de errores
        private string ERROR = "";

        // Constructor: inicializa las propiedades
        public CConexion()
        {
            CNN = null;
            DS = null;
            ERROR = "";
        }

        // Método para establecer la conexión
        public bool Conectar(string cadenaConexion)
        {
            bool resultado = false;
            CNN = new OleDbConnection();
            CNN.ConnectionString = cadenaConexion;
            try
            {
                CNN.Open();
                // Se instancia el DataSet que servirá de contenedor de tablas en memoria
                DS = new DataSet();
                resultado = true;
            }
            catch (Exception ex)
            {
                ERROR = ex.Message;
            }
            return resultado;
        }

        // Método para cerrar la conexión de forma segura
        public bool Desconectar()
        {
            bool resultado = false;
            try
            {
                if (CNN != null && CNN.State == ConnectionState.Open)
                {
                    CNN.Close();
                    resultado = true;
                }
            }
            catch (Exception ex)
            {
                ERROR = ex.Message;
            }
            return resultado;
        }

        // Método para recuperar el último error ocurrido
        public string ObtenerError()
        {
            return ERROR;
        }

        // Ejecuta un comando parametrizado (OleDbCommand). La conexión debe estar abierta.
        public bool EjecutarComando(OleDbCommand comando)
        {
            try
            {
                if (CNN == null || CNN.State != ConnectionState.Open)
                {
                    ERROR = "Conexión no establecida.";
                    return false;
                }
                comando.Connection = CNN;
                comando.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                ERROR = ex.Message;
                return false;
            }
        }

        // Ejecuta una cadena SQL sin parámetros
        public bool EjecutarComando(string sql)
        {
            try
            {
                if (CNN == null || CNN.State != ConnectionState.Open)
                {
                    ERROR = "Conexión no establecida.";
                    return false;
                }
                using (var cmd = new OleDbCommand(sql, CNN))
                {
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                ERROR = ex.Message;
                return false;
            }
        }

        // Ejecuta una consulta y devuelve un DataTable con el resultado
        public DataTable EjecutarConsulta(string sql)
        {
            try
            {
                if (CNN == null || CNN.State != ConnectionState.Open)
                {
                    ERROR = "Conexión no establecida.";
                    return null;
                }
                using (var da = new OleDbDataAdapter(sql, CNN))
                {
                    var dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                ERROR = ex.Message;
                return null;
            }
        }
    }
}