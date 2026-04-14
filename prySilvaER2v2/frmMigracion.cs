using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using prySilvaER2v2.BaseDatos;

namespace prySilvaER2v2
{
    public partial class frmMigracion : Form
    {
        public frmMigracion()
        {
            InitializeComponent();
        }

        private void Log(string mensaje)
        {
            txtInfo.AppendText(mensaje + Environment.NewLine);
            Application.DoEvents();
        }

        private void cmdIniciar_Click(object sender, EventArgs e)
        {
            string articuloFile = Path.Combine(Application.StartupPath, "Articulos.txt");
            string categoriaFile = Path.Combine(Application.StartupPath, "Categorias.txt");
            string dbFile = Path.Combine(Application.StartupPath, "Distribuidora.mdb");

            try
            {
                txtInfo.Clear();
                Log("Inicio de migración...");

                // 1) Crear archivos de ejemplo si no existen
                CreateSampleFiles(categoriaFile, articuloFile);

                // 2) Crear base de datos Access (si no existe)
                if (!File.Exists(dbFile))
                {
                    Log("Creando base de datos 'Distribuidora.mdb'...");
                    bool creado = CreateAccessDatabase(dbFile);
                    Log(creado ? "Base de datos creada correctamente." : "Error creando base de datos.");
                }
                else
                {
                    Log("La base de datos ya existe. Se utilizará la existente.");
                }

                // 3) Conectar a la base de datos
                var conexion = new CConexion();
                string cadena = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbFile};";
                if (!conexion.Conectar(cadena))
                {
                    Log("No se pudo conectar a la base de datos: " + conexion.ObtenerError());
                    return;
                }
                Log("Conexión establecida.");

                // 4) Crear tablas (si no existen)
                // Eliminamos tablas si existen por simplicidad antes de crear
                conexion.EjecutarComando("DROP TABLE Articulos");
                conexion.EjecutarComando("DROP TABLE Categorias");

                string sqlCategorias = "CREATE TABLE Categorias (Id INTEGER PRIMARY KEY, Descripcion TEXT(100))";
                string sqlArticulos = "CREATE TABLE Articulos (Codigo INTEGER PRIMARY KEY, Nombre TEXT(100), Precio DOUBLE, Stock INTEGER, CategoriaId INTEGER)";

                if (conexion.EjecutarComando(sqlCategorias))
                    Log("Tabla 'Categorias' creada.");
                else
                    Log("Error creando 'Categorias': " + conexion.ObtenerError());

                if (conexion.EjecutarComando(sqlArticulos))
                    Log("Tabla 'Articulos' creada.");
                else
                    Log("Error creando 'Articulos': " + conexion.ObtenerError());

                // 5) Leer y migrar Categorias
                Log("Migrando categorías desde archivo: " + categoriaFile);
                int contCat = 0;
                foreach (var linea in File.ReadAllLines(categoriaFile))
                {
                    if (string.IsNullOrWhiteSpace(linea)) continue;
                    var partes = linea.Split(';');
                    if (partes.Length < 2)
                    {
                        Log("Línea inválida en Categorias.txt: " + linea);
                        continue;
                    }
                    int id;
                    if (!int.TryParse(partes[0].Trim(), out id))
                    {
                        Log("Id inválido en línea: " + linea);
                        continue;
                    }
                    string descripcion = partes[1].Trim();

                    var cmd = new OleDbCommand("INSERT INTO Categorias (Id, Descripcion) VALUES (?, ?)");
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.Parameters.AddWithValue("@Descripcion", descripcion);

                    if (conexion.EjecutarComando(cmd))
                    {
                        contCat++;
                    }
                    else
                    {
                        Log("Error insertando categoría: " + conexion.ObtenerError());
                    }
                }
                Log($"Categorías migradas: {contCat}");

                // 6) Leer y migrar Articulos
                Log("Migrando artículos desde archivo: " + articuloFile);
                int contArt = 0;
                foreach (var linea in File.ReadAllLines(articuloFile))
                {
                    if (string.IsNullOrWhiteSpace(linea)) continue;
                    var partes = linea.Split(';');
                    if (partes.Length < 5)
                    {
                        Log("Línea inválida en Articulos.txt: " + linea);
                        continue;
                    }
                    int codigo;
                    if (!int.TryParse(partes[0].Trim(), out codigo))
                    {
                        Log("Código inválido en línea: " + linea);
                        continue;
                    }
                    string nombre = partes[1].Trim();
                    double precio;
                    if (!double.TryParse(partes[2].Trim(), out precio))
                    {
                        Log("Precio inválido en línea: " + linea);
                        continue;
                    }
                    int stock;
                    if (!int.TryParse(partes[3].Trim(), out stock))
                    {
                        Log("Stock inválido en línea: " + linea);
                        continue;
                    }
                    int categoriaId;
                    if (!int.TryParse(partes[4].Trim(), out categoriaId))
                    {
                        Log("CategoriaId inválido en línea: " + linea);
                        continue;
                    }

                    var cmd = new OleDbCommand("INSERT INTO Articulos (Codigo, Nombre, Precio, Stock, CategoriaId) VALUES (?, ?, ?, ?, ?)");
                    cmd.Parameters.AddWithValue("@Codigo", codigo);
                    cmd.Parameters.AddWithValue("@Nombre", nombre);
                    cmd.Parameters.AddWithValue("@Precio", precio);
                    cmd.Parameters.AddWithValue("@Stock", stock);
                    cmd.Parameters.AddWithValue("@CategoriaId", categoriaId);

                    if (conexion.EjecutarComando(cmd))
                    {
                        contArt++;
                    }
                    else
                    {
                        Log("Error insertando artículo: " + conexion.ObtenerError());
                    }
                }
                Log($"Artículos migrados: {contArt}");

                conexion.Desconectar();
                Log("Migración finalizada.");
            }
            catch (Exception ex)
            {
                Log("Error en la migración: " + ex.Message);
            }
        }

        private void CreateSampleFiles(string categoriaFile, string articuloFile)
        {
            if (!File.Exists(categoriaFile))
            {
                var categorias = new string[] {
                    "1;Procesadores",
                    "2;Placas Madre",
                    "3;Memoria RAM",
                    "4;Discos SSD"
                };
                File.WriteAllLines(categoriaFile, categorias, Encoding.UTF8);
                Log("Archivo 'Categorias.txt' creado de ejemplo.");
            }
            else
            {
                Log("Archivo 'Categorias.txt' ya existe.");
            }

            if (!File.Exists(articuloFile))
            {
                var articulos = new string[] {
                    "101;Intel i7;4500.5;10;1",
                    "102;ASUS Prime;1200;5;2",
                    "103;Kingston 8GB;800;20;3",
                    "104;Samsung 970;3500;7;4"
                };
                File.WriteAllLines(articuloFile, articulos, Encoding.UTF8);
                Log("Archivo 'Articulos.txt' creado de ejemplo.");
            }
            else
            {
                Log("Archivo 'Articulos.txt' ya existe.");
            }
        }

        private bool CreateAccessDatabase(string path)
        {
            try
            {
                // Use ADOX Catalog via COM to create a new .mdb
                Type catalogType = Type.GetTypeFromProgID("ADOX.Catalog");
                if (catalogType == null)
                {
                    Log("No se encontró ADOX en el sistema. Asegúrese de tener instalado 'Microsoft ADO Ext. 2.8 for DDL and Security'.");
