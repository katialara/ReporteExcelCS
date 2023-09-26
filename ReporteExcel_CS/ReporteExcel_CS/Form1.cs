using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ReporteExcel_CS
{
    public partial class Form1 : Form
    {
        static string conexionstring = "server= LAPTOP-HAO3D8QC\\SQLEXPRESS; database= Tienda_Ropa; integrated security= true";
        SqlConnection conexion = new SqlConnection(conexionstring);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            conexion.Open();
            string query = "Select * from empleado";
            SqlCommand comando = new SqlCommand(query, conexion);
            SqlDataAdapter data = new SqlDataAdapter(comando);
            DataTable tabla = new DataTable(); 
            data.Fill(tabla);
            dataGridView1.DataSource = tabla;
        }

        private void button2_Click(object sender, EventArgs e)
        {           
            Excel.Application excelApp = new Excel.Application();

            // Abrir el archivo existente
            Excel.Workbook workbook = excelApp.Workbooks.Open("C:\\Users\\pauli\\OneDrive\\Documentos\\Excel2CS.xlsx");
            
            // Obtener la hoja de trabajo en la que deseas exportar los datos
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Puedes cambiar el número de la hoja según tus necesidades

            worksheet.Cells[1, 1] = "IdEmpleado";
            worksheet.Cells[1, 2] = "Nombres";
            worksheet.Cells[1, 3] = "Apellido Paterno";
            worksheet.Cells[1, 4] = "Apellido Materno";
            worksheet.Cells[1, 5] = "Estado_civil";
            worksheet.Cells[1, 6] = "Fecha";
            worksheet.Cells[1, 7] = "Telefono";
            worksheet.Cells[1, 8] = "Correo";
            worksheet.Cells[1, 9] = "RFC";
            worksheet.Cells[1, 10] = "NSS";
            worksheet.Cells[1, 11] = "Direccion";
            worksheet.Cells[1, 12] = "Ciudad";
            worksheet.Cells[1, 12] = "Estado";
            worksheet.Cells[1, 12] = "Estatus";
            // Recorrer las filas y columnas del DataGridView

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        worksheet.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    else
                    {
                        worksheet.Cells[i + 1, j + 1] = ""; 
                    }
                }
            }

            // Guardar los cambios en el archivo Excel
            workbook.Save();

            // Cerrar y liberar recursos
            workbook.Close(false);
            excelApp.Quit();
            
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            Marshal.ReleaseComObject(worksheet);


        }
    }
}
