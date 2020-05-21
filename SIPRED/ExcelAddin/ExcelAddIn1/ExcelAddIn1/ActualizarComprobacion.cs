using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class ActualizarComprobacion : Form
    {
        public int _iYear;
        public oComprobacion _oComprobacion;
        public Form _Form;
        string accion = "";
        int TpPlantilla = 0;
        public ActualizarComprobacion(int IdCruce, oComprobacion _Comprobacion, int IDPlantilla, string Accion, int _Year)
        {
            InitializeComponent();
            
            txtNro.Text = IdCruce.ToString();
            _iYear = _Year;
            _oComprobacion = _Comprobacion;
            if (_Comprobacion != null)
            {
                txtConcepto.Text = _Comprobacion.Concepto;
                string[] Formula = _Comprobacion.Formula.Split('=');
                txtcelda.Text = Formula[0];
                txtformula.Text = Formula[1];
                txtCondicion.Text = _Comprobacion.Condicion;
                if (txtCondicion.Text.Trim() != "")
                {
                    txtCondicion.ReadOnly = false;
                    chkCondicionar.Checked = true;
                }
                txtNota.Text = _Comprobacion.Nota;
                


            }
            accion = Accion;

            TpPlantilla = IDPlantilla;
        }

        private void btguardar_Click(object sender, EventArgs e)
        {
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            string _Title = "Conexión de Red";
            if (_Connection)
            {
                _Message = (txtConcepto.Text.Trim() == "") ? "- Debe indicar concepto." : "";
                _Message += (txtcelda.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar celda." : "";
                _Message += (txtformula.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar fórmula." : "";
                _Message += (chkCondicionar.Checked && txtCondicion.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar condición." : "";

                if (_Message.Length > 0)
                {
                    MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DialogResult _response = DialogResult.None;
                string Formulax = (txtcelda.Text + "=" + txtformula.Text);
                string condicion = "";
                
                if (chkCondicionar.Checked)
                    condicion = txtCondicion.Text;


                oComprobacion _Template = new oComprobacion()
                {
                    IdComprobacion = Convert.ToInt32(txtNro.Text),
                    IdTipoPlantilla = TpPlantilla,//buscar
                    Concepto = txtConcepto.Text,
                    Formula = Formulax,
                    Condicion = condicion,
                    Nota = txtNota.Text,
                };

                if (accion == "A")
                {
                    
                    KeyValuePair<bool, string[]> _result = new lComprobacionesAdmin(_Template, accion).Add();

                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Comprobación agregada con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    //if (_result.Key) this.Hide();
                }
                else
                if (accion == "M")
                {
                    KeyValuePair<bool, string[]> _result = new lComprobacionesAdmin(_Template, accion).Update();

                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Comprobación modificada con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    //if (_result.Key) this.Hide();
                }

                string _Path = Configuration.Path;
                oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{_Path}\\jsons\\Plantillas.json");
                oPlantilla _Temp = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _Template.IdTipoPlantilla && o.Anio == _iYear);

                if (_oComprobacion != null)
                {
                    //Libro Actual de Excel.
                    Excel.Worksheet xlSht;
                    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    string[] Formula = _oComprobacion.Formula.Split('=');
                    string[] _celdaBase = Formula[0].Replace("[","").Replace("]", "").Split(',');
                    string[] _celdaFin = Formula[1].Replace("[", "").Replace("]", "").Split(',');
                    Excel.Range _RangeO;
                    Excel.Range _RangeR;

                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_celdaBase[0]);
                    for (int a = 1; a < 1000; a++)
                    {
                        _RangeO = (Excel.Range)xlSht.get_Range($"A" + a.ToString());

                        if(_RangeO != null)
                        {
                            if (_RangeO.Value.ToString() == _celdaBase[1])
                            {
                                _RangeR = (Excel.Range)xlSht.get_Range($"{Generales.ColumnAdress(Int32.Parse(_celdaBase[2]))}" + a.ToString());
                                _RangeR.Formula = "";
                                _RangeR.Value = "";
                                _RangeR.NumberFormat = "@";
                                _RangeR.Value2 = "";

                                a = 1001;
                            }
                        }
                    }
                }
                
                _Form.Close();
                FileJson(_Temp, TpPlantilla.ToString());
                _Form.Show();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkCondicionar_CheckedChanged(object sender, EventArgs e)
        {
            txtCondicion.ReadOnly = !chkCondicionar.Checked;
            if (!chkCondicionar.Checked)
                txtCondicion.Text = "";
        }

        private void txtNota_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void FileJson(oPlantilla _Template, string _Tipo)
        {
            this.TopMost = false;
            //this.Enabled = false;
            //this.Hide();
            FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
            _FileJsonfrm._Form = this;
            _FileJsonfrm._Process = true;
            _FileJsonfrm._Update = true;
            _FileJsonfrm._Automatic = true;
            _FileJsonfrm._Template = _Template;
            _FileJsonfrm._Tipo = _Tipo;
            _FileJsonfrm._window = this.Text;
            _FileJsonfrm.Show();
            
            this.TopMost = true;
            this.Close();
            return;
        }
    }
}
