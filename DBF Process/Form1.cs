using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using dBASE.NET;
using System.Globalization;
using System.IO;


namespace DBF_Process
{
    public partial class Paso1 : Form
    {
        public static int WeeksInYear(DateTime date)
        {
            GregorianCalendar cal = new GregorianCalendar(GregorianCalendarTypes.Localized);
            return cal.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);
        }
        public class Field
        {
            public string FieldName { get; set; }
            public string FieldType { get; set; }
            public int FieldLength { get; set; }
        } 
        public Paso1()
        {
            InitializeComponent();
        }
        
        public Dbf dbf = new Dbf();
        public DataSet dbfDS = new DataSet();
        
        
        private void button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Encoding iso = Encoding.GetEncoding("ISO-8859-1");
            Encoding utf8 = Encoding.Unicode;
            FileInfo fi = new FileInfo(this.textBox1.Text); 
            DataTable dt = ParseDBF.ReadDBF(this.textBox1.Text);
            dt.TableName = "TablaDatos"; //Nombra el data table
            dbfDS.Tables.Add(dt); //Agrega el data table al dataset
            string campo = textBox2.Text;
            string formato = comboBox1.SelectedItem.ToString();
            dbfDS.Tables["TablaDatos"].Columns.Add("MASTERDATE*8", typeof(DateTime));
            dbfDS.Tables["TablaDatos"].Columns.Add("YEARNWEEK*6", typeof(String));
            string day = string.Empty;
            string month = string.Empty;
            string year = string.Empty;
            HashSet<string> semanas = new HashSet<string>();
            if (formato == "Fecha")
            {
                campo = campo + "*8";
                foreach (DataRow row in dbfDS.Tables["TablaDatos"].Rows)
                {
                    DateTime tempdate;
                    try
                    {
                        tempdate = DateTime.Parse(row[campo].ToString());
                    }
                    catch
                    {
                        MessageBox.Show("!Hay registros sin fecha");
                        tempdate = new DateTime(2020, 1, 1);
                    }
                    row["MASTERDATE*8"] = tempdate;
                    row["YEARNWEEK*6"] = tempdate.Year.ToString() + WeeksInYear(tempdate).ToString("00");
                    semanas.Add(tempdate.Year.ToString() + WeeksInYear(tempdate).ToString("00"));
                }
            }
            else if(formato == "mm/dd/aaaa")
            {
                campo = campo + "*10";
                foreach (DataRow row in dbfDS.Tables["TablaDatos"].Rows)
                {
                    day = row[campo].ToString().Substring(3,2);
                    month = row[campo].ToString().Substring(0, 2);
                    year = row[campo].ToString().Substring(6, 4);
                    DateTime tempdate;
                    try
                    {
                        tempdate = new DateTime(Int32.Parse(year), Int32.Parse(month), Int32.Parse(day));
                    }
                    catch 
                    {
                        MessageBox.Show("!Hay fwchas con formato erroneo");
                        tempdate = new DateTime(2020, 1, 1);
                    }
                    row["MASTERDATE*8"] = tempdate;
                    row["YEARNWEEK*6"] = tempdate.Year.ToString() + WeeksInYear(tempdate).ToString("00");
                    semanas.Add(tempdate.Year.ToString() + WeeksInYear(tempdate).ToString("00"));
                }
            }
           
            List<Field> DSFields = new List<Field>();
            foreach(DataColumn Column in dt.Columns)
            {
                Field field = new Field();
                field.FieldName = Column.ColumnName.Substring(0,Column.ColumnName.IndexOf("*", 0));
                field.FieldType = Column.DataType.ToString();
                int location = Column.ColumnName.Length - Column.ColumnName.IndexOf("*", 0) - 1;
                field.FieldLength = Int32.Parse(Column.ColumnName.Substring(Column.ColumnName.IndexOf("*", 0) + 1, location));
                DSFields.Add(field);
            }

            foreach (Field field in DSFields) 
            {
                if (field.FieldType == "System.String")
                {
                    dbf.Fields.Add(new DbfField(field.FieldName, DbfFieldType.Character, Convert.ToByte(field.FieldLength)));
                }
                else if (field.FieldType == "System.Int32")
                {
                    dbf.Fields.Add(new DbfField(field.FieldName, DbfFieldType.Numeric, Convert.ToByte(field.FieldLength)));
                }
                else if (field.FieldType == "System.DateTime")
                {
                    dbf.Fields.Add(new DbfField(field.FieldName, DbfFieldType.Date, 8));
                }
                else
                {
                    dbf.Fields.Add(new DbfField(field.FieldName, DbfFieldType.Character, Convert.ToByte(field.FieldLength)));
                }
            }
            foreach(string semana in semanas)
            {
                var query = from r in dt.AsEnumerable()
                            where r.Field<string>("YEARNWEEK*6") == semana//Convert.ToInt32(Folio)
                            select r;
                DataTable dtDestino = query.CopyToDataTable<DataRow>();
  
                DbfRecord record;
                foreach (DataRow row in dtDestino.Rows)
                {
                    record = dbf.CreateRecord();
                    for (int i = 0; i < dtDestino.Columns.Count; i++)
                    {
                        if (row[i].GetType().ToString() == "System.String")
                        {
                            byte[] utfBytes = utf8.GetBytes(row[i].ToString());
                            byte[] isoBytes = Encoding.Convert(utf8, iso, utfBytes);
                            string msg = iso.GetString(isoBytes);
                            record.Data[i] = msg;
                        }
                        else record.Data[i] = row[i];
                    }

                }
                dbf.Write(fi.DirectoryName + "//" + textBox3.Text + "S" + semana.Substring(4,2) + ".dbf", DbfVersion.FoxBaseDBase3NoMemo);
                dbf.Records.Clear();

            }

            MessageBox.Show("Proceso concluido exitosamente");
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            this.textBox1.Text = openFileDialog1.FileName;
        }
       
    }
}
