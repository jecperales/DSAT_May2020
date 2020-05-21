﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ExcelAddIn1
{
    public class Generales
    {
        /// <summary>Función para Proteger y Desproteger las hojas de un archivo de Excel.
        /// <para>Desprotege y Protege un archivo de Excel. Referencia: <see cref="Proteccion(bool)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="Proteccion(bool)"/>
        /// </summary>
        public static void Proteccion(bool accion)
        {
            int f;
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            using (ExcelPackage _package = new ExcelPackage(_Excel))
            {
                for (f = 1; f <= _package.Workbook.Worksheets.Count(); f++)
                {
                    Excel.Worksheet xlSht = wb.Worksheets[f];
                    if (!accion)
                    {
                        //xlSht.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);
                    }
                    else
                    {
                        //xlSht.Protect(ExcelAddIn.Access.Configuration.PwsExcel, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);
                    }
                }
            }
        }
        /// <summary>Función para Insertar el Indice.
        /// <para>Inserta el Indice en el archivo de Excel. Referencia: <see cref="InsertIndice(Excel.Worksheet, int, Excel.Range, bool, int)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="InsertIndice(Excel.Worksheet, int, Excel.Range, bool, int)"/>
        /// </summary>
        public static void InsertIndice(Excel.Worksheet xlSht, int CantReg, Excel.Range currentCell, bool ConFormula, int NroPrincipal)
        {
            Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string _NameFile = wb.Name;
            string _AnexoFile = sheet.Name;
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            Excel.Range currentFind = null;
            Excel.Range currentFindExpl = null;

            Excel.Range RangeLocked = null;
            int NroRow = currentCell.Row;
            int NroColumn = currentCell.Column;
            string IndicePrevio = "";
            long IndiceInicial = 0;

            int iTotalColumns = 0;
            int k = 1;
            int i = 1;
            long indiceNvo = 0;
            int CantExpl = 0;
            currentCell = (Excel.Range)xlSht.Cells[NroRow, 1];
            IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
            currentFindExpl = (Excel.Range)xlSht.Cells[NroRow + 1, 1];

            if (currentFindExpl.get_Value(Type.Missing) != null)
                if (currentFindExpl.get_Value(Type.Missing).ToString().ToUpper().Trim() == "EXPLICACION")
                    NroRow++;

            IndiceInicial = Convert.ToInt64(IndicePrevio) + 100;
            int rowexpl = 0;
            List<int> FilasExplicacion = new List<int>();
            int CantRango = 0; long IndiceInicialx = IndiceInicial;

            foreach (Excel.Name cname in Globals.ThisAddIn.Application.Names)
            {
                if (cname.Name == "IA_0" + Convert.ToString(IndiceInicialx))
                {
                    CantRango++;
                    IndiceInicialx = IndiceInicialx + 100;
                    rowexpl = cname.RefersToRange.Cells.Row + 1;
                    currentFindExpl = (Excel.Range)xlSht.Cells[rowexpl, 1];

                    if (currentFindExpl.get_Value(Type.Missing) != null)
                        if (currentFindExpl.get_Value(Type.Missing).ToString().ToUpper().Trim() == "EXPLICACION")
                        {
                            CantExpl++;
                            if (!FilasExplicacion.Contains(rowexpl + CantReg))//los indices que tienen explicacion la fila actual + los registros que ingresó nvos
                                FilasExplicacion.Add(rowexpl + CantReg);
                        }
                }
            }

            currentFind = currentCell.Find(IndiceInicial, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                     Type.Missing, Type.Missing);

            int NroRowx = 0;
            CantRango = CantRango + CantExpl;
            int NroPrincipalAux = DameRangoPrincipal(NroPrincipal, xlSht);

            while (i <= CantReg)
            {
                indiceNvo = Convert.ToInt64(IndicePrevio) + 100;
                Excel.Range rangej = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));
                rangej.Select();
                rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                var rangeall = xlSht.get_Range(string.Format("{0}:{0}", NroPrincipalAux - 1, Type.Missing));
                var rangeaCopy = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));
                iTotalColumns = xlSht.UsedRange.Columns.Count;
                rangeall.Copy();
                rangeaCopy.PasteSpecial(Excel.XlPasteType.xlPasteFormulas, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                rangeaCopy.Locked = false;

                k = 1;

                while (k <= iTotalColumns)
                {
                    if (!(rangeaCopy.Cells[k].HasFormula))
                        rangeaCopy.Cells[k].Value = "";

                    k = k + 1;
                }

                xlSht.Cells[NroRow + i, 1] = "0" + Convert.ToString(indiceNvo);
                sheet.Controls.Remove("IA_0" + indiceNvo);
                AddNamedRange(NroRow + i, 1, "IA_0" + Convert.ToString(indiceNvo));
                currentCell = (Excel.Range)xlSht.Cells[NroRow + i, 1];
                IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
                //currentCell = xlSht.Range[xlSht.Cells[NroRow + i, 1], xlSht.Cells[NroRow + i, 3]];
                //currentCell.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);

                ((Excel.Range)xlSht.Cells[NroRow + 1, 2]).NumberFormat = "General"; // le doy formato text al concepto
                ((Excel.Range)xlSht.Cells[NroRow + 1, 2]).WrapText = true;
                RangeLocked = (Excel.Range)xlSht.Cells[NroRow + i, 1];
                RangeLocked.Locked = true; // con esto bloqueo solo la primera columna

                i++;
            }

            if (currentFind != null)
            {
                NroRowx = NroRow + CantReg;
                currentCell = (Excel.Range)xlSht.Cells[NroRowx, 1];
                IndicePrevio = currentCell.get_Value(Type.Missing).ToString();

                int j = 1;

                while (j <= CantRango)
                {
                    if (!FilasExplicacion.Contains(NroRowx + j))
                    {
                        indiceNvo = Convert.ToInt64(IndicePrevio) + 100;
                        xlSht.Cells[NroRowx + j, 1] = "0" + Convert.ToString(indiceNvo);
                        sheet.Controls.Remove("IA_0" + indiceNvo);
                        AddNamedRange(NroRowx + j, 1, "IA_0" + Convert.ToString(indiceNvo));
                        currentCell = (Excel.Range)xlSht.Cells[NroRowx + j, 1];
                        IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
                    }

                    j++;
                }
            }

            string NombreHoja = xlSht.Name.ToUpper().Replace(" ", "");
            List<oSubtotal> ColumnasST = DameColumnasST((NombreHoja));
            Excel.Range Sum_Range = null;
            int NroFinal = NroRow + CantReg + CantRango;

            int _Rango = 0;
            string _Renglon;
            string _Columna;
            string _rCelda = "";

            foreach (oSubtotal ST in ColumnasST)
            {
                Sum_Range = xlSht.get_Range(ST.Columna + (NroPrincipalAux).ToString(), ST.Columna + (NroPrincipalAux).ToString());
                Sum_Range.Formula = "=sum(" + ST.Columna + (NroPrincipalAux + 1).ToString() + ":" + ST.Columna + (NroFinal).ToString() + ")";
                _Rango = NroFinal - NroPrincipalAux; // nroprincipalaux siempre tiene el numero de la fila padre al restar te da la cantidad de rows insertadas
                _Renglon = (Sum_Range.Row).ToString();
                _Columna = Generales.ColumnAdress(Sum_Range.Column);
                _rCelda = _Columna + "" + _Renglon;
                InsertarReferencia(_NameFile, _AnexoFile, _rCelda, _Rango, _Columna, _Renglon, CantReg);
            }

            Sum_Range = xlSht.get_Range("B" + (NroPrincipal).ToString(), "B" + (NroPrincipal).ToString());
            Sum_Range.Select();
           
        }
        public static void ActualizarReferencia(string _Archivo, string _Anexo, string _Celda, int _Cantidad, string _Column, string _Row,int CantEliminar, string Accion)
        {
            string _Path = ExcelAddIn.Access.Configuration.Path;

            if (File.Exists(_Path + "\\references\\" + _Archivo + ".json"))
            {
                string _jCadena = "";
                StreamReader _fJason = new StreamReader(_Path + "\\references\\" + _Archivo + ".json");

                while (_fJason.Peek() >= 0)
                {
                    var _json = _fJason.ReadLine();
                    oIndices _Indices = JsonConvert.DeserializeObject<oIndices>(_json);

                    if (_Indices.Anexo + _Indices.Column + _Indices.Row == _Anexo + _Column + _Row)
                    {
                        if (Accion == "E")
                        {//Eliminar
                            _Indices.Cantidad -= CantEliminar;
                        }
                        else if (Accion == "A")
                        {//Agregar
                            _Indices.Cantidad += 1;
                        }
                        _json = JsonConvert.SerializeObject(_Indices);
                    }
                    if (_Indices.Column != _Column)
                    {
                        if (_Indices.Row == _Row)
                        {
                            if (Accion == "E")
                            {//Eliminar
                                _Indices.Cantidad -= CantEliminar;
                            }
                            else if (Accion == "A")
                            {//Agregar
                                _Indices.Cantidad += 1;
                            }
                                _json = JsonConvert.SerializeObject(_Indices);
                        }
                    }
                    if (_Indices.Anexo + _Indices.Row != _Anexo + _Row)
                    {
                        if (Convert.ToInt32(_Row) < Convert.ToInt32(_Indices.Row))
                        {
                            if (Accion == "E")
                            {//Eliminar
                                _Indices.Row = (Convert.ToInt32(_Indices.Row) - CantEliminar).ToString();
                            }
                            else if (Accion == "A")
                            {//Agregar
                                _Indices.Row = (Convert.ToInt32(_Indices.Row) + 1).ToString();
                            }
                            _json = JsonConvert.SerializeObject(_Indices);
                        }
                    }

                    _jCadena += _json;
                    _jCadena += string.Format(Environment.NewLine);
                }

                _fJason.Close();
                File.WriteAllText(_Path + "\\references\\" + _Archivo + ".json", _jCadena);
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                wb.Save();
            }
        }
        public static void InsertarReferencia(string _Archivo, string _Anexo, string _Celda, int _Cantidad, string _Column, string _Row, int _Posicion)
        {
            string _Path = ExcelAddIn.Access.Configuration.Path;
            var Indices = new oIndices
            {
                Archivo = _Archivo,
                Anexo = _Anexo,
                Celda = _Celda,
                Cantidad = _Cantidad,
                Column = _Column,
                Row = _Row
            };

            string jReference = JsonConvert.SerializeObject(Indices);

            if (Directory.Exists(_Path + "\\references"))
            {
                if (!File.Exists(_Path + "\\references\\" + _Archivo + ".json"))
                {
                    File.WriteAllText(_Path + "\\references\\" + _Archivo + ".json", jReference);
                }
                else
                {
                    string _jCadena = "";
                    StreamReader _fJason = new StreamReader(_Path + "\\references\\" + _Archivo + ".json");
                    
                    while(_fJason.Peek() >= 0)
                    {
                        var _json = _fJason.ReadLine();
                        oIndices _Indices = JsonConvert.DeserializeObject<oIndices>(_json);

                        //if(Convert.ToInt32(_Row) < Convert.ToInt32(_Indices.Row))
                        //{
                        //    Match m = Regex.Match(_Indices.Celda, "(\\d+)");
                        //    string _num = "0";
                        //    if (m.Success){_num = m.Value;}

                        //    _Indices.Row = (Convert.ToInt32(_Indices.Row) + _Posicion).ToString();
                        //    _json = JsonConvert.SerializeObject(_Indices);
                        //}
                        if (_Indices.Column + _Indices.Row != _Column + _Row)
                        {
                            _jCadena += _json;
                            _jCadena += string.Format(Environment.NewLine);
                        }
                    }

                    _fJason.Close();
                    //_jCadena += string.Format(Environment.NewLine);
                    _jCadena += jReference;
                    File.WriteAllText(_Path + "\\references\\" + _Archivo + ".json", _jCadena);
                }
              
            }
            else
            {
                Directory.CreateDirectory(_Path + "\\references");
                if (!File.Exists(_Path + "\\references\\" + _Archivo + ".json"))
                {
                    File.WriteAllText(_Path + "\\references\\" + _Archivo + ".json", jReference);
                }
                else
                {
                    string _jCadena = "";
                    StreamReader _fJason = new StreamReader(_Path + "\\references\\" + _Archivo + ".json");

                    while (_fJason.Peek() >= 0)
                    {
                        _jCadena = _jCadena + _fJason.ReadLine();
                    }

                    _fJason.Close();
                    _jCadena = _jCadena + jReference;
                    File.WriteAllText(_Path + "\\references\\" + _Archivo + ".json", _jCadena);
                }
            }

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            wb.Save();
        }
        /// <summary>Función para Insertar la Explicación.
        /// <para>Inserta la Explicación en el archivo de Excel. Referencia: <see cref="InsertaExplicacion(Excel.Worksheet, Excel.Range, string)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="InsertaExplicacion(Excel.Worksheet, Excel.Range, string)"/>
        /// </summary>
        public static void InsertaExplicacion(Excel.Worksheet xlSht, Excel.Range currentCell, string Explicacion)
        {
            var rangej = xlSht.get_Range(string.Format("{0}:{0}", currentCell.Row + 1, Type.Missing));
            rangej.Select();
            
            rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            xlSht.Cells[currentCell.Row + 1, 1] = " EXPLICACION ";
            xlSht.Cells[currentCell.Row + 1, 2] = Explicacion;

            ((Excel.Range)xlSht.Cells[currentCell.Row + 1, 2]).NumberFormat = "General";
            ((Excel.Range)xlSht.Cells[currentCell.Row + 1, 2]).WrapText = true;

            currentCell.Select();
            currentCell = xlSht.Range[xlSht.Cells[currentCell.Row + 1, 1], xlSht.Cells[currentCell.Row + 1, 2]];
            currentCell.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);
            currentCell.Locked = true;

            int iTotalColumns = xlSht.UsedRange.Columns.Count;
            int k = 3;
            while (k <= iTotalColumns)
            {
                currentCell = xlSht.Range[xlSht.Cells[currentCell.Row, k], xlSht.Cells[currentCell.Row, k]];
                currentCell.Locked = true;
                k++;
            }
            //ref
            string NombreHoja = xlSht.Name.ToUpper().Replace(" ", "");
            List<oSubtotal> ColumnasST = Generales.DameColumnasST(NombreHoja);
            int _Registro = 1;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string _NameFile = wb.Name;
            int NroPrincipal = currentCell.Row + 1;
            int row = DameRangoPrincipal(NroPrincipal, xlSht);
            foreach (oSubtotal ST in ColumnasST)
            {
                if (_Registro == 1)
                {
                    _Registro += 1;
                    Generales.ActualizarReferencia(_NameFile, xlSht.Name.ToUpper(), ST.Columna + row.ToString(), 0, ST.Columna, row.ToString(), 0, "A");
                }
            }
          }
        /// <summary>Función para Agregar Rango con Nombre.
        /// <para>Agrega el Rango con Nombre en el archivo de Excel. Referencia: <see cref="AddNamedRange(int, int, string)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="AddNamedRange(int, int, string)"/>
        /// </summary>
        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;
            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            
            Excel.Range cell = worksheet.Cells[row, col];
            try
            {
                NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);
            }
            catch
            {

            }
        }
        /// <summary>Función para Agregar Rango con Nombre. pasando la hoja
        /// <para>Agrega el Rango con Nombre en el archivo de Excel. Referencia: <see cref="AddNamedRange(int, int, string, Excel.Worksheet)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="AddNamedRange(int, int, string, Excel.Worksheet)"/>
        /// </summary>
        public static void AddNamedRange(int row, int col, string myrango, Excel.Worksheet xlsht)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(xlsht);

            Excel.Range cell = worksheet.Cells[row, col];
            try
            {
                NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);
            }
            catch
            {

            }
        }
        /// <summary>Función para obtener el Rango Principal.
        /// <para>Obtiene el Rango Principal en el archivo de Excel. Referencia: <see cref="DameRangoPrincipal(int, Excel.Worksheet)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="DameRangoPrincipal(int, Excel.Worksheet)"/>
        /// </summary>
        public static int DameRangoPrincipal(int NroPrincipal, Excel.Worksheet xlSht)
        {
            int NroPrincipalAux = NroPrincipal;
            Excel.Range objRange = (Excel.Range)xlSht.Cells[NroPrincipal, 2];
            var ConceptoPrevio = objRange.get_Value(Type.Missing);
            if (ConceptoPrevio != null)
            {
                ConceptoPrevio = ConceptoPrevio.ToString();

                if  (EsConceptoValido(ConceptoPrevio))
                {
                    while (NroPrincipalAux > 0)
                    {
                        objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                        ConceptoPrevio = objRange.get_Value(Type.Missing);

                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                           
                            if (EsConceptoValido(ConceptoPrevio))
                                break;
                        }
                        NroPrincipalAux--;
                    }
                }
                else
                {
                    while (NroPrincipalAux > 0)
                    {
                        objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                        ConceptoPrevio = objRange.get_Value(Type.Missing);
                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                            if (EsConceptoValido(ConceptoPrevio))
                                break;
                        }
                        NroPrincipalAux--;
                    }
                }
            }
            else
            {
                while (NroPrincipalAux > 0)
                {
                    objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                    ConceptoPrevio = objRange.get_Value(Type.Missing);
                    if (ConceptoPrevio != null)
                    {
                        ConceptoPrevio = ConceptoPrevio.ToString();
                        if (EsConceptoValido(ConceptoPrevio))
                            break;
                        
                    }
                    NroPrincipalAux--;
                }
            }
            return NroPrincipalAux;
        }
        /// <summary>Función que valida si el Concepto es Valido.
        /// <para>Obtiene Verdadero o Falso en el archivo de Excel. Referencia: <see cref="EsConceptoValido(string)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="EsConceptoValido(string)"/>
        /// </summary>
        public static bool EsConceptoValido(string Concepto)
        {
            bool CncValido = false;
            List<oConcepto> ConceptVal = new List<oConcepto>();
            ConceptVal = DameConceptosValidos();

            foreach (oConcepto c in ConceptVal)
            {
                if (Concepto.Length >= c.Caracteres)
                {
                    if (Concepto.ToUpper().Substring(0, c.Caracteres).Contains(c.Descripcion.ToUpper()))
                    {
                        CncValido = true;
                        break;
                    }
                }
            }
            return CncValido;
        }
        /// <summary>Función que obtiene los Sub Totales.
        /// <para>Obtiene los Sub Totales en el archivo de Json. Referencia: <see cref="DameColumnasST(string)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="DameColumnasST(string)"/>
        /// </summary>
        public static List<oSubtotal> DameColumnasST(string Hoja)
        {
            List<oSubtotal> Subtotales = new List<oSubtotal>();
          
            string _Path = ExcelAddIn.Access.Configuration.Path;
            bool _Connection = new lSerializados().CheckConnection(ExcelAddIn.Access.Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\Indices.json"))
                {
                    if (_Connection)
                    {
                        KeyValuePair<bool, System.Data.DataTable> _TipoPlantilla = new lSerializados().ObtenerUpdate();

                        foreach (DataRow _Row in _TipoPlantilla.Value.Rows)
                        {
                            string _IdTipoPlantilla = _Row["IdTipoPlantilla"].ToString();
                            string _Fecha_Modificacion = _Row["Fecha_Modificacion"].ToString();
                            string _Linea = null;

                            if (File.Exists(_Path + "\\jsons\\Update" + _IdTipoPlantilla + ".txt"))
                            {
                                StreamReader sw = new StreamReader(_Path + "\\Jsons\\Update" + _IdTipoPlantilla + ".txt");
                                _Linea = sw.ReadLine();
                                sw.Close();

                                if (_Linea != null)
                                {
                                    if (_Linea != _Fecha_Modificacion)
                                    {
                                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                                        _FileJsonfrm._Process = true;
                                        _FileJsonfrm._Update = true;
                                        _FileJsonfrm._window = "";
                                        _FileJsonfrm.Show();
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                    else
                    {
                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                        _FileJsonfrm._Process = true;
                        _FileJsonfrm._window = "";
                        _FileJsonfrm.Show();
                    }
                }
            }
            else
            {
                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }
                if (!_Connection)
                {
                    MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                else
                {
                    FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                    _FileJsonfrm._Process = true;
                    _FileJsonfrm._window = "";
                    _FileJsonfrm.Show();
                }
            }

            oRootobject _Root = Assembler.LoadJson<oRootobject>($"{_Path}\\jsons\\Indices.json");
           
            Subtotales = _Root.Subtotales;

            return Subtotales.Where(x => x.Hoja == Hoja.Trim()).ToList();
        }
        /// <summary>Función para obtener los Conceptos Validos.
        /// <para>Obtiene los Conceptos Validos en el archivo de Json. Referencia: <see cref="DameConceptosValidos()"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="DameConceptosValidos()"/>
        /// </summary>
        public static List<oConcepto> DameConceptosValidos()
        {
            List<oConcepto> Conceptos = new List<oConcepto>();
            string _Path = ExcelAddIn.Access.Configuration.Path;

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (!File.Exists(_Path + "\\jsons\\Indices.json"))
                {
                    FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                    _FileJsonfrm._Process = true;
                    _FileJsonfrm._window = "";
                    _FileJsonfrm.Show();
                }
            }
            else
            {
                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                _FileJsonfrm._Process = true;
                _FileJsonfrm._window = "";
                _FileJsonfrm.Show();
            }

            oRootobject _Root = Assembler.LoadJson<oRootobject>($"{_Path}\\jsons\\Indices.json");
            Conceptos = _Root.Conceptos;
            return Conceptos;
        }
        /// <summary>Función para convertir de numero a letras.
        /// <para>Convierte de número a letras en el campo específico. Referencia: <see cref="ColumnAdress(int)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="ColumnAdress(int)"/>
        /// </summary>
        public static string ColumnAdress(int col)
        {
            if (col <= 26)
            {
                return Convert.ToChar(col + 64).ToString();
            }
            int div = col / 26;
            int mod = col % 26;
            if (mod == 0) { mod = 26; div--; }
            return ColumnAdress(div) + ColumnAdress(mod);
        }


        /// <summary>Función para obtener la contraseña de la hoja activa de Excel.
        /// <para>Obtiene la contraseña de la hoja activa de Excel para desbloquearla. Referencia: <see cref="_Macro(bool, Excel.Worksheet, string)"/> se agrega la referencia ExcelAddIn.Generales para invocarla.</para>
        /// <seealso cref="_Macro(bool, Excel.Worksheet, string)"/>
        /// </summary>
        public static void _Macro(bool accion, Excel.Worksheet xlSht, string PwsExcel)
        {
            int a, b, c, d, e, f, a1, a2, a3, a4, a5, a6;
            string _Contrasena = PwsExcel;

            try
            {
                if (!accion)
                {
                    xlSht.Unprotect(_Contrasena);

                    if (xlSht.ProtectContents == false)
                    {
                        //Byte[] _data = ASCIIEncoding.ASCII.GetBytes(_Contrasena);
                        //ConfigurationManager.AppSettings["VAL7"] = Convert.ToBase64String(_data);
                        return;
                    }
                }
                else
                {
                    xlSht.Protect(_Contrasena, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);

                    if (xlSht.ProtectContents == true)
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                for (a = 65; a <= 66; a++)
                {
                    for (b = 65; b <= 66; b++)
                    {
                        for (c = 65; c <= 66; c++)
                        {
                            for (d = 65; d <= 66; d++)
                            {
                                for (e = 65; e <= 66; e++)
                                {
                                    for (a1 = 65; a1 <= 66; a1++)
                                    {
                                        for (a2 = 65; a2 <= 66; a2++)
                                        {
                                            for (a3 = 65; a3 <= 66; a3++)
                                            {
                                                for (a4 = 65; a4 <= 66; a4++)
                                                {
                                                    for (a5 = 65; a5 <= 66; a5++)
                                                    {
                                                        for (a6 = 65; a6 <= 66; a6++)
                                                        {
                                                            for (f = 32; f <= 126; f++)
                                                            {
                                                                _Contrasena = "";
                                                                _Contrasena += Convert.ToChar(a);
                                                                _Contrasena += Convert.ToChar(b);
                                                                _Contrasena += Convert.ToChar(c);
                                                                _Contrasena += Convert.ToChar(d);
                                                                _Contrasena += Convert.ToChar(e);
                                                                _Contrasena += Convert.ToChar(a1);
                                                                _Contrasena += Convert.ToChar(a2);
                                                                _Contrasena += Convert.ToChar(a3);
                                                                _Contrasena += Convert.ToChar(a4);
                                                                _Contrasena += Convert.ToChar(a5);
                                                                _Contrasena += Convert.ToChar(a6);
                                                                _Contrasena += Convert.ToChar(f);

                                                                try
                                                                {
                                                                    if (!accion)
                                                                    {
                                                                        xlSht.Unprotect(_Contrasena);

                                                                        if (xlSht.ProtectContents == false)
                                                                        {
                                                                            Byte[] _data = ASCIIEncoding.ASCII.GetBytes(_Contrasena);
                                                                            ConfigurationManager.AppSettings["VAL7"] = Convert.ToBase64String(_data);
                                                                            //Byte[] _data = ASCIIEncoding.ASCII.GetBytes(_Contrasena);
                                                                            //Configuration configServ = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                                                                            //configServ.AppSettings.Settings["VAL7"].Value = Convert.ToBase64String(_data);
                                                                            //configServ.Save(ConfigurationSaveMode.Modified);
                                                                            return;
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        xlSht.Protect(_Contrasena, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);

                                                                        if (xlSht.ProtectContents == true)
                                                                        {
                                                                            Byte[] _data = ASCIIEncoding.ASCII.GetBytes(_Contrasena);
                                                                            ConfigurationManager.AppSettings["VAL7"] = Convert.ToBase64String(_data);
                                                                            //Byte[] _data = ASCIIEncoding.ASCII.GetBytes(_Contrasena);
                                                                            //Configuration configServ = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                                                                            //configServ.AppSettings.Settings["VAL7"].Value = Convert.ToBase64String(_data);
                                                                            //configServ.Save(ConfigurationSaveMode.Modified);
                                                                            return;
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception exc) { }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return;
            }
        }
    }
}
