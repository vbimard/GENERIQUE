using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/////windows
///forms
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Security.Principal;
//ini file
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
/////almacam
///wpm
using Wpm.Implement.Manager;
using Wpm.Schema.Kernel;
using Actcut.EstimationManager;
using Actcut.ResourceManager;
using Alma.NetWrappers;
using Wpm.Implement.ComponentEditor;
using Actcut.ActcutModelManager;
using Actcut.ActcutModel;
using System.Collections;
using Actcut.CommonModel;

namespace AF_ImportTools
//  namespace Import_GP
{
    #region declaration interface

    public interface IImport
    {
        bool CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary, Dictionary<string, string> CentreFrais_Dictionnary = null, bool Sans_Donnees_Technique = false);
        bool CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary);

    }



    #endregion
    #region Test
    //CETTE classe ne sert que pour les tests
    public class ImportToolTest
    { }
    #endregion


    #region declaration enum ou type
    public enum TypeTole : int { Chute = 2, Tole = 1 };
    // public struct Vector { double X; double Y ; double Norme ;}
    #endregion

    #region structure_declaration
    public struct Champ
    {
        public string fieldname;
        public Type fieldtype;
        public int position;
        public int maxSize;
        public object defaultValue;
    }

    #endregion
    #region Import_Param
    /// <summary>
    /// import param seciton 
    /// definir les sections = nom de dictonnaire  
    /// </summary>
    public class Import_Param {


        Dictionary<string, string> Param_Model; // stock les différents model
        Dictionary<string, object> Parameters;// stock les paramètres // normalement recuperer sur l'interface almacam
        Dictionary<string, string> Param_Directory;// stock les chemin

        public virtual void Set_Default_Param_Model(string type, string value)
        {

            //var Param_Model = new Dictionary<string, string>();
            Param_Model = new Dictionary<string, string>();
            {
                Param_Model.Add("MODEL_CA", "0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#_MATERIAL#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#CUSTOMER#string;10#_QUANTITY#integer;11#QUANTITY#double;12#ECOQTY#string;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;16#FORMATCLIP#string;17#IDMAT#string;18#IDLNBOM#string;19#NUMMAG#string;20#FILENAME#string;21#_DESCRIPTION#string;22#AF_CDE#string;23#DELAI_INT#date;24#EN_RANG#string;25#EN_PERE_PIECE#string;26#ID_PIECE_CFAO#string");
                Param_Model.Add("MODEL_DM", "0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string");
                Param_Model.Add("MODEL_PATH", "0#TECHNOLOGIE#string;1#ImportCda#string;0#ImportDM#string;2#ExportRp#string;3#ExportDT#string;4#Centredefrais#string;5#Destination_Path#string;6#Source_Path#string");

            }

        }

        public virtual void Set_Default_Parameters(string type, string value)
        {

            //var Parameters = new Dictionary<string, object>();
            Parameters = new Dictionary<string, object>();
            {
                Parameters.Add("STRING_FORMAT_DOUBLE", "{ 0:0.00###}");
                Parameters.Add("ALMACAM_EDITOR_NAME", "Wpm.Implement.Editor.exe");
                Parameters.Add("CLIPPER_MACHINE_CF", "CLIP");
                Parameters.Add("VERBOSE_LOG", true);
                Parameters.Add("IMPORT_AUTO", true);

            }

        }
        public virtual void Set_Default_Param_Directory(string type, string value)
        {
            //var Param_Directory = new Dictionary<string, string>();
            Param_Directory = new Dictionary<string, string>();
            {
                Param_Directory.Add("IMPORT_CA", @"C:\Alma\Datas\_Clipper\Import_OF\CAHIER_AFFAIRE.csv");
                Param_Directory.Add("IMPORT_DM", @"C:\Alma\Datas\_Clipper\Import_Stock\DISPO_MAT.csv");
                Param_Directory.Add("IMPORT_Rp", @"C:\Alma\Datas\_Clipper\Export_GPAO");
                Param_Directory.Add("IMPORT_Dt", @"C:\Alma\Datas\_Clipper\Export_DT");
                Param_Directory.Add("EMF_DIRECTORY", @"C:\Alma\Datas\_Clipper\Emf");
                Param_Directory.Add("SHEET_REQUIREMENT_DIRECTORY", @"C:\Alma\Datas\_Clipper\Export_Sheet_requirements");
                Param_Directory.Add("APPLICATION1", @"C:\AlmaCAM\Bin\AlmaCamUser1.exe");
                Param_Directory.Add("LOG_DIRECTORY", "");
                Param_Directory.Add("WORKSHOP_OPTION", "");

            }
        }

        public void Read_Param()
        { }
        public void Write_Param()
        { }
        public void Create_default_Param()
        { }

        public string GetParameterValueAsString(string name) {
            try {
                string result = null;
                if (this.Parameters.ContainsKey(name))
                {
                    result = Parameters[name].ToString();
                }
                else if (this.Param_Directory.ContainsKey(name)) {
                    result = Param_Directory[name].ToString();
                }
                else if (this.Param_Model.ContainsKey(name)) {
                    result = Param_Model[name].ToString();
                } else
                {
                    //ecriture log
                    result = null;
                }
                return result;

            }
            catch (Exception ie) {
                MessageBox.Show(ie.Message);
                return null;
            }


        }
        void GetParameterValueAsInteger(string name) { }
        void GetParameterValueAsDouble(string name) { }
        void GetParameterValueAsBoolean(string name) { }
    }




    #endregion
    #region DataModel
    /// <summary>
    /// la Class DataModel recupere un dictionnaire d objets fortement typés ordonnés par le data_model_string
    /// le data_model string est une chaine contenun sur une ligne unique et listant index|nomchamp|type;index2|nomchamp2|type2...
    /// la method set field structure un dictionnaire d'objet ensuite traité et intégré dans la base almacam
    /// </summary>
    public static class Data_Model
    {
        //declaration des delegates
        //definiton du dictionnaire de champ pour control de l'integrite te validation fichier
        public static Dictionary<string, Champ> _Field_Dictionnary = new Dictionary<string, Champ>();
        private static int _LineNumber = 0;
        private static int _colnumber = 0;
        //rempli le dictionnaire de type


        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <param name="linedictionnary"></param>
        /// <returns></returns>
        public static object GetLineDictionnaryObject(string key, ref Dictionary<string, object> linedictionnary)
        {
            object result;
            try
            {
                bool flag = linedictionnary.ContainsKey(key);
                object obj;
                if (flag)
                {
                    obj = linedictionnary[key];
                }
                else
                {
                    Alma_Log.Write_Log_Important(key + "not found in line dictionnary");
                    obj = null;
                }
                result = obj;
            }
            catch
            {
                result = null;
            }
            return result;
        }

        public static void setFieldDictionnary(string Data_Model_String)
        {
            string[] fieldlist; Champ newField;
            //
            _Field_Dictionnary.Clear();
            //premier niveau
            fieldlist = Data_Model_String.Split(';');
            //2eme niveau
            foreach (string field in fieldlist)
            {

                string[] fieldinfos;
                //object o;
                fieldinfos = field.Split('#');
                newField.defaultValue = "";
                //on ecrit le dictionnaire de type
                newField.position = Convert.ToInt32(fieldinfos[0]);
                newField.fieldname = fieldinfos[1];
                newField.fieldtype = Data_Model.getTypeOf(fieldinfos[2], out newField.defaultValue);
                newField.maxSize = 1000;
                _Field_Dictionnary.Add(newField.position.ToString(), newField);

            }

            ;
        }
        /// <summary>
        /// traite une ligne avant le split
        /// </summary>
        /// <param name="line"></param>
        /// <param name="caracteres_to_be_removed"></param>
        public static void TreatLine(ref string line, string caracteres_to_be_removed)
        {
            try
            {
                Data_Model.RemoveSpecialCharacters(ref line, caracteres_to_be_removed);
                Debug.Print(line);
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log("error in Trealine methode : " + ex.Message);
            }
        }
        /// <summary>
        /// enleve les caractères spéciaux
        /// </summary>
        /// <param name="line"></param>
        /// <param name="caracteres_to_be_removed"></param>
        public static void RemoveSpecialCharacters(ref string line, string caracteres_to_be_removed)
        {
            try
            {
                line = System.Text.RegularExpressions.Regex.Replace(line, caracteres_to_be_removed, "");
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log("RemoveSpecialCharacters : " + ex.Message);
            }
        }
        /// <summary>
        /// retourne un Type sur la base d'une chaine de caractère enumérant ce type  comme ci dessous
        /// "0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string..."
        ///         /// </summary>
        /// <param name="strType">string, int, bool..</param>
        /// <returns>typeof(string), typeof(int),typeof(bool)</returns>
        public static string ConvertToString(object ObjetctoConvert)
        {
            Type mtype = ObjetctoConvert.GetType();


            switch (mtype.ToString())
            {
                case "string": return ObjetctoConvert.ToString();
                case "int": return ObjetctoConvert.ToString();
                case "integer": return ObjetctoConvert.ToString();
                case "long": return ObjetctoConvert.ToString();
                case "dbl": return ObjetctoConvert.ToString();
                case "double": return ObjetctoConvert.ToString();
                case "bool": return ObjetctoConvert.ToString();
                case "boolean": return ObjetctoConvert.ToString();
                case "date": {
                        DateTime dt;
                        dt = (DateTime)ObjetctoConvert;
                        return dt.ToString("MM / dd / yyyy");
                    }

            }

            return null;
        }
        /// <summary>
        /// retourne un Type sur la base d'une chaine de caractère enumérant ce type  comme ci dessous
        /// "0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string..."
        ///         /// </summary>
        /// <param name="strType">string, int, bool..</param>
        /// <returns>typeof(string), typeof(int),typeof(bool)</returns>
        static Type getTypeOf(string strType)
        {
            switch (strType)
            {
                case "string": return typeof(string);
                case "int": return typeof(int);
                case "integer": return typeof(int);
                case "long": return typeof(Int32);
                case "dbl": return typeof(Double);
                case "double": return typeof(Double);
                case "bool": return typeof(Boolean);
                case "boolean": return typeof(Boolean);
                case "date": return typeof(DateTime);
            }

            return typeof(string);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strType"></param>
        /// <param name="defaultvalue"></param>
        /// <returns></returns>
        // ImportTools.Data_Model
        private static Type getTypeOf(string strType, out object defaultvalue)
        {
            defaultvalue = "";
            Type result = typeof(string);

            if (strType == "boolean")
            {
                defaultvalue = false;
                result = typeof(bool);
                return result;
            }
            else if (strType == "string")
            {
                defaultvalue = "";
                result = typeof(string);
                return result;
            }
            else if (strType == "double")
            {
                defaultvalue = 0.0;
                result = typeof(double);
                return result;
            }

            else if (strType == "int")
            {
                defaultvalue = 0;
                result = typeof(int);
                return result;
            }

            else if (strType == "long")
            {
                defaultvalue = 0;
                result = typeof(long);
                return result;
            }
            else if (strType == "integer")
            {
                defaultvalue = 0;
                result = typeof(int);
                return result;
            }
            else if (strType == "date")
            {
                defaultvalue = DateTime.Parse("05/06/2013");
                result = typeof(DateTime);
                return result;
            }
            else if (strType == "dbl")
            {
                defaultvalue = 0.0;
                result = typeof(double);
                return result;
            }
            else if (strType == "bool")
            {
                defaultvalue = false;
                result = typeof(bool);
                return result;
            }

            return result;

        }




















        /// <summary>
        /// lit une ligne, verifie conformité et validation erreur de lecture
        /// retourne une liste d'objet fromatés
        /// </summary>
        /// <param name="csvline">ligne csv avec separateur ; </param>
        /// <param name="lineNumber">nulero de ligne en cours d'etude</param>
        /// <returns>liste d'objet</returns>
        /// //public static List<object> ReadCsvLine(string csvline, int lineNumber, Dictionary<string, champ> FieldDictionnary)
        public static List<object> ReadCsvLine(string csvline, int lineNumber)
        {
            List<object> ls = new List<object>();
            try
            {
                string[] line;

                line = csvline.Split(';');
                if (_Field_Dictionnary.Count() != line.Count()) { throw new InvalidDataException("Nombre de colonnes incorrecte à la ligne numéro " + lineNumber); }

                int colnumber = 0;
                foreach (string vvalue in line)
                {
                    if (vvalue.Length > _Field_Dictionnary[colnumber.ToString()].maxSize) { throw new InvalidDataException("Nombre invalide de caractere sur champ numéro" + colnumber.ToString() + " intitulé : " + _Field_Dictionnary[colnumber.ToString()].fieldname + "   incorrecte à la ligne numéro " + lineNumber); }
                    if (_Field_Dictionnary[colnumber.ToString()].fieldtype == typeof(System.DateTime))
                    {
                        //cas du date time
                        ls.Add(Convert.ToDateTime(vvalue));
                    }
                    else
                    {
                        ls.Add(Convert.ChangeType(vvalue, _Field_Dictionnary[colnumber.ToString()].fieldtype));
                    }

                    colnumber++;
                }

                return ls;

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return ls;
            }
        }
        /// <summary>
        /// retourne un dictionnaire de type <string,object> d'objet nommés <"nomduchamp",objet>
        /// </summary>
        /// <param name="csvline">ligne csv</param>
        /// <returns>dictionnaire <string,object> </returns>
        public static Dictionary<string, object> ReadCsvLine_With_Dictionnary(string csvline)
        {
            Dictionary<string, object> line_Dic = new Dictionary<string, object>();
            try
            {
                string[] line;
                string separateurdecimal;
                separateurdecimal = ",";
                separateurdecimal = string.Format("{0,0}", .5).Substring(1, 1);


                line = csvline.Split(';');
                if (_Field_Dictionnary.Count() != line.Count()) {
                    Alma_Log.Write_Log_Important(System.Reflection.MethodBase.GetCurrentMethod().Name);
                    Alma_Log.Write_Log(csvline);
                    Alma_Log.Write_Log("le nombre de colonnes est incorrecte. \r\n Le fichier d'import contient plus de colonne que celles decrites dans le model, verfier le model ou bien le contenu du fichier csv");
                    //return null;
                    throw new InvalidDataException("le nombre de colonnes est incorrecte. \r\n Le fichier d'import contient plus de colonne que celles decrites dans le model, verfier le model ou bien le contenu du fichier csv");
                }

                _colnumber = 0;

                foreach (string vvalue in line)
                {

                    if (vvalue.Length != 0)
                    {   // on trim //
                        vvalue.Trim();
                        if (vvalue.Length > _Field_Dictionnary[_colnumber.ToString()].maxSize) { throw new InvalidDataException("Nombre invalide de caractere sur champ numéro" + _colnumber.ToString() + " intitulé : " + _Field_Dictionnary[_colnumber.ToString()].fieldname); }



                        if (_Field_Dictionnary[_colnumber.ToString()].fieldtype == typeof(System.DateTime))
                        {
                            //get_day month year
                            string inputdate = vvalue;
                            //DateTime dt3 = Convert.ToDateTime(inputdate+" 00:00:00");
                            if (inputdate == "") { inputdate = "24/06/1973"; }
                            IFormatProvider culture = new System.Globalization.CultureInfo("fr-FR", true);
                            DateTime dt2 = DateTime.Parse(inputdate, culture, System.Globalization.DateTimeStyles.AssumeLocal);
                            line_Dic.Add(_Field_Dictionnary[_colnumber.ToString()].fieldname, dt2);

                        }

                        else if (_Field_Dictionnary[_colnumber.ToString()].fieldtype.ToString() == "System.String")
                        {
                            //get_day month year
                            string text_value = vvalue;
                            //acune conversion
                            line_Dic.Add(_Field_Dictionnary[_colnumber.ToString()].fieldname, text_value.ToString());

                        }
                        else
                        {

                            line_Dic.Add(_Field_Dictionnary[_colnumber.ToString()].fieldname, Convert.ChangeType(vvalue.Replace(".", separateurdecimal), _Field_Dictionnary[_colnumber.ToString()].fieldtype));
                        }
                    }
                    _colnumber++;
                }

                return line_Dic;

            }
            catch (Exception e)
            {
                Alma_Log.Write_Log_Important(System.Reflection.MethodBase.GetCurrentMethod().Name);
                Alma_Log.Write_Log(csvline);
                Alma_Log.Write_Log(string.Format(" Erreur de conversion de données dans la fonction ReadCsvLine_With_Dictionnary  \r\n erreur possible mauvais type de donnée dans le csv (quantité en decimale...)..."));
                //System.Windows.Forms.MessageBox.Show(string.Format(" Erreur de conversion de données dans la fonction ReadCsvLine_With_Dictionnary {0} \r\n erreur possible mauvais type de donnée dans le csv (quantité en decimale...) :\r\n {1}",
                Alma_Log.Write_Log(_Field_Dictionnary[_colnumber.ToString()].fieldname + ":" + e.Message);
                //return null;
                return line_Dic;
            }
        }
        /// <summary>
        /// implement des valeurs par defaut en focntion du type et de la ligne d'entree
        /// </summary>
        /// <param name="csvline"></param>
        /// <param name="trim_values"></param>
        /// <returns></returns>
        public static Dictionary<string, object> ReadCsvLine_With_Dictionnary2(string csvline, bool trim_values)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            Dictionary<string, object> result;
            try
            {
                string newValue = string.Format("{0,0}", 0.5).Substring(1, 1);
                string[] array = csvline.Split(new char[] { ';' });
                bool flag = Data_Model._Field_Dictionnary.Count<KeyValuePair<string, Champ>>() != array.Count<string>();
                if (flag)
                {
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name);
                    Alma_Log.Write_Log(csvline);
                    Alma_Log.Write_Log("Le nombre de colonnes est incorrecte. \r\n Le fichier d'import contient plus de colonne que celles decrites dans le model, verfier le model ou bien le contenu du fichier csv");
                    MessageBox.Show(string.Concat(new object[]
                    {
                "Le fichier d'import n'est pas conforme a la descritpion du model, il manque des champs : le nombre de champs demandés est de ",
                Data_Model._Field_Dictionnary.Count<KeyValuePair<string, Champ>>(),
                " alors que le fichier n'en contient que ",
                array.Count<string>()
                    }));
                    throw new InvalidDataException("le nombre de colonnes est incorrecte. \r\n Le fichier d'import contient plus de colonne que celles decrites dans le model, verfier le model ou bien le contenu du fichier csv");
                }
                Data_Model._colnumber = 0;
                string[] array2 = array;
                for (int i = 0; i < array2.Length; i++)
                {
                    string text = array2[i];
                    bool flag2 = text.Length != 0;
                    if (flag2)
                    {
                        bool flag3 = text.Length > Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].maxSize;
                        if (flag3)
                        {
                            throw new InvalidDataException("Nombre invalide de caractere sur champ numéro" + Data_Model._colnumber.ToString() + " intitulé : " + Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname);
                        }
                        bool flag4 = Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldtype == typeof(DateTime);
                        if (flag4)
                        {
                            string text2 = text;
                            bool flag5 = text2 == "";
                            if (flag5)
                            {
                                text2 = "24/06/1973";
                            }
                            IFormatProvider provider = new CultureInfo("fr-FR", true);
                            DateTime dateTime = DateTime.Parse(text2, provider, DateTimeStyles.AssumeLocal);
                            dictionary.Add(Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname, dateTime);
                        }
                        else if (trim_values)
                        {
                            dictionary.Add(Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname, Convert.ChangeType(text.Replace(".", newValue).Trim(), Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldtype));
                        }
                        else
                        {
                            dictionary.Add(Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname, Convert.ChangeType(text.Replace(".", newValue), Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldtype));
                        }
                    }
                    else
                    {
                        dictionary.Add(Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname, Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].defaultValue);
                    }
                    Data_Model._colnumber++;
                }
                result = dictionary;
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name);
                Alma_Log.Write_Log(csvline);
                Alma_Log.Write_Log(string.Format(" Erreur de conversion de données dans la fonction ReadCsvLine_With_Dictionnary  \r\n erreur possible mauvais type de donnée dans le csv (quantité en decimale...)...", new object[0]));
                Alma_Log.Write_Log(Data_Model._Field_Dictionnary[Data_Model._colnumber.ToString()].fieldname + ":" + ex.Message);
                result = dictionary;
            }
            return result;
        }
        /// <summary>
        /// lit une ligne, verifie conformité et validation erreur de lecture
        /// retourne une liste d'objet fromatéss
        /// </summary>
        /// <param name="csvline">ligne csv avec separateur ; </param>
        /// <returns>liste d'objet</returns>
        public static List<object> ReadCsvLine(string csvline)
        {

            List<object> ls = new List<object>();

            try
            {
                string[] line;

                line = csvline.Split(';');
                if (_Field_Dictionnary.Count() != line.Count()) { throw new InvalidDataException("Nombre de colonnes incorrecte à la ligne numéro " + _LineNumber.ToString()); }

                int colnumber = 0;
                foreach (string vvalue in line)
                {
                    if (vvalue.Length > _Field_Dictionnary[colnumber.ToString()].maxSize || vvalue.Trim() == string.Empty)
                    {
                        ls.Add(null);
                    }
                    //throw new InvalidDataException("Nombre invalide de caractere sur champ numéro" + colnumber.ToString() + " intitulé : " + _FieldDictionnary[colnumber.ToString()].fieldname + "   incorrecte à la ligne numéro " + _LineNumber.ToString()); }
                    else
                    {

                        if (_Field_Dictionnary[colnumber.ToString()].fieldtype == typeof(DateTime))
                        {
                            //   date francaise type fr dd/mm/yyy et non l'inverse
                            ls.Add(DateTime.Parse(vvalue, new CultureInfo("fr-FR")));
                        }
                        else
                        {
                            ls.Add(Convert.ChangeType(vvalue, _Field_Dictionnary[colnumber.ToString()].fieldtype));
                        }

                    }


                    colnumber++;
                }

                return ls;

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return ls;
            }
        }
        /// <summary>
        /// retourn l'index d'un champs
        /// </summary>
        /// <param name="Index">numero du champs</param>
        /// <returns>integer</returns>
        /// 
        public static string getFieldName(int Index)
        {
            return _Field_Dictionnary[Index.ToString()].fieldname;
        }
        /// <summary>
        /// retourne le nom du champs 
        /// </summary>
        /// <param name="FieldName">nom du champs </param>
        /// <returns>index</returns>
        public static Int32 getFieldNumber(string FieldName)
        {
            int i = 0; string result = "";

            foreach (string s in _Field_Dictionnary.Keys)
            {
                if (_Field_Dictionnary[i.ToString()].fieldname == FieldName)
                { result = i.ToString(); }

                i++;
            }
            return Convert.ToInt32(result);
        }

        /// <summary>
        /// retourne un champ existe dans le dictionnaire de champ
        /// </summary>
        /// <param name="keyname">nom du champ</param>
        /// <returns>true/false</returns>
        public static bool ExistsInDictionnary(string keyname)
        {
            try
            {
                if (_Field_Dictionnary.ContainsKey(keyname) == false)
                {
                    MessageBox.Show(string.Format("keyname {0} not found in dictionnary ", keyname));
                    //Alma_Log.Write_Log(   string.Format("keyname {0} not found in dictionnary ", keyname)
                }
                return _Field_Dictionnary.ContainsKey(keyname);

            }
            catch (KeyNotFoundException) { MessageBox.Show(string.Format("keyname {0} not found", keyname)); return false; }
        }
        /// <summary>
        /// retourne si une clé exite dans le dictuionnaire pointé
        /// </summary>
        /// <param name="keyname">nom de la clé</param>
        /// <param name="dictionnary">dictionnaire par reference</param>
        /// <returns>true/false</returns>
        public static bool ExistsInDictionnary(string keyname, ref Dictionary<string, object> dictionnary)
        {
            try
            {
                if (dictionnary.ContainsKey(keyname) == false)
                {
                    //MessageBox.Show(string.Format("keyname {0} not found in dictionnary ", keyname));
                    //string.Format("keyname {0} not found in dictionnary ", keyname)
                    Alma_Log.Write_Log(string.Format("keyname {0} not found in dictionnary ", keyname));
                }
                return dictionnary.ContainsKey(keyname);

            }
            catch (KeyNotFoundException) { MessageBox.Show(string.Format("keyname {0} not found", keyname)); return false; }
        }
        /// <summary>
        /// retourne si une clé existe dans un dictionnaire pointé
        /// </summary>
        /// <param name="keyname">nom de la clé</param>
        /// <param name="dictionnary">dictionnaire par reference</param>
        /// <returns>true/false</returns>
        public static bool ExistsInDictionnary(string keyname, ref Dictionary<string, string> dictionnary)
        {
            try
            {


                if (dictionnary.ContainsKey(keyname) == false)
                {
                    //MessageBox.Show(string.Format("keyname {0} does not existe in dictionnary ", keyname));
                    Alma_Log.Write_Log(string.Format("Filed dictionnary, keyname {0} not found in dictionnary ", keyname)); return false;
                }

                return dictionnary.ContainsKey(keyname);

            }
            catch (KeyNotFoundException) { Alma_Log.Write_Log(string.Format("keyname {0} not found in dictionnary ", keyname)); return false; }//MessageBox.Show(string.Format("keyname {0} not found", keyname)); return false; }
        }
        /// <summary>
        /// retourne un objet si l'objet exist dans le dictionnaire
        /// </summary>
        /// <param name="keyname">cle a rechercher</param>
        /// <param name="dictionnary"></param>
        /// <returns>objet ou null </returns>
        public static object ReturnObject_If_ExistsInDictionnary(string keyname, ref Dictionary<string, object> dictionnary)
        {
            object item = null;
            if (ExistsInDictionnary(keyname, ref dictionnary)) { item = dictionnary[keyname]; }
            return item;
        }
        /// <summary>
        /// update un item et les champs associés
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="item"></param>
        /// <param name="line_dictionnary"></param>
        public static void update_Item(IContext contextlocal, IEntity item, Dictionary<string, object> line_dictionnary)
        {
            try
            {
                foreach (var field in line_dictionnary)
                {
                    item.SetFieldValue(field.Key, field.Value);
                }
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); }
        }
    }


    #endregion
    /// <summary>
    ///  information sur la base de données
    /// </summary>
    #region DataBase

    public static class DataBase
    {

        /// <summary>
        /// recupere la derniere base ouverte par almacam ou bien ouvre la base demandée
        /// </summary>
        /// <param name="DatabaseName">nom de la base a connecter si vide recupere celle du registre</param>
        /// <param name="_Context">recupere le nouveau context </param>
        /// <returns></returns>
        public static IContext Connect(ref string mDatabaseName)
        {
            try
            {

                //int databaseAmount = 0;
                bool databasefound = false;
                IModelsRepository mRepository = new ModelsRepository();
                IContext contextelocal = null;
                //si le databasname est vide alors on recherche dans le registre
                if (mDatabaseName == "")
                {
                    mDatabaseName = Alma_RegitryInfos.GetLastDataBase();
                    if (mDatabaseName != "") { databasefound = true; }

                }
                else
                {
                    databasefound = mRepository.DatabaseExist(mDatabaseName);
                }

                //creation du model repository
                if (databasefound)
                {




                    if (mRepository.DatabaseExist(mDatabaseName))
                    {

                        contextelocal = mRepository.GetModelContext(mDatabaseName);  //nom de la base;
                    }
                    databasefound = true;
                }
                else
                {
                    MessageBox.Show(mDatabaseName + " : not found");
                    contextelocal = null;
                    databasefound = false;
                }





                return contextelocal;
            }

            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }

        }

        public static List<IModel> GetDataBaseList()
        {
            try
            {

                List<IModel> lstmodel = new List<IModel>();
                IModelsRepository mRepository = new ModelsRepository();
                lstmodel = mRepository.ModelList.ToList<IModel>();


                return lstmodel;
            }

            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }

        }


    }


    #endregion
    /// <summary>
    /// la classe material retourne l'ensemble des information sur une matiere via divers methodes
    /// oiuyr l'instant seul les index et les nom amtiere sont rendus public
    /// </summary>
    public static class Material
    {

        public static string Material_Name { get; set; }
        public static Int32 Material_Id32 { get; set; }
        public static string Quality_Name { get; set; }
        public static Int32 Quality_Id32 { get; set; }
        public static string Quality_set_Name { get; set; }
        public static Int32 Quality_Set_Id32 { get; set; }
        /// <summary>
        /// recupere le nom de la matiere en fonction de l'epaisseur et de la nuance
        /// </summary>
        /// <param name="contextlocal">contexte</param>
        /// <param name="nuance">grade as string</param>
        /// <param name="Thickness">thickness as double</param>
        /// <returns>undef or nommatiere</returns>
        public static string getMaterial_Name(IContext contextlocal, string nuance, double Thickness) {
            //string material_name = null;
            try {
                IEntityList grades = null;
                IEntityList materials = null;
                IEntity grade = null;
                string material = null;


                grades = contextlocal.EntityManager.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, nuance);
                grades.Fill(false);

                if (grades.Count() > 0) {
                    grade = grades.FirstOrDefault();
                    materials = contextlocal.EntityManager.GetEntityList("_MATERIAL", LogicOperator.And, "_THICKNESS", ConditionOperator.Equal, Thickness, "_QUALITY", ConditionOperator.Equal, grade.Id32);//("_MATERIAL", "_QUALITY", ConditionOperator.Equal, grade.Id32);//("_MATERIAL", LogicOperator.And, "_THICKNESS", ConditionOperator.Equal, Thickness, "_QUALITY", ConditionOperator.Like, grade);
                    materials.Fill(false);
                    if (materials.Count > 0) {
                        material = materials.FirstOrDefault().GetFieldValueAsString("_NAME");
                    }
                    else { material = string.Empty;
                        Alma_Log.Write_Log(string.Format("Material {0} not Found ", nuance + " " + Thickness + "mm"));
                    }
                    //.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, nuance);

                }
                else { material = string.Empty;
                    Alma_Log.Write_Log(string.Format("Material {0} not Found ", nuance + " " + Thickness + "mm"));
                }


                return material;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return string.Empty;
            }
        }
        /// <summary>
        /// retourne un nom matiere en fonction de l id 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="materialID"></param>
        /// <returns></returns>
        public static string getMaterial_Name(IContext contextlocal, Int32 materialID) {
            IEntityList materials = null;
            IEntity material = null;
            try
            {
                materials = contextlocal.EntityManager.GetEntityList("_MATERIAL", "ID", ConditionOperator.Equal, materialID);
                //materials = contextlocal.EntityManager.GetEntityList("_MATERIAL");
                materials.Fill(false);

                if (materials.Count() > 0 && materials.FirstOrDefault().Status.ToString() == "Normal")
                { material = materials.FirstOrDefault(); }
                else { material = null; }

                return material.GetFieldValueAsString("_NAME");
            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
                return null;
            }
        }
        /// <summary>
        /// retourne une reference de matiere a partie d'une matiere epaisseur..
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nuance"></param>
        /// <param name="Thickness"></param>
        /// <returns></returns>
        public static IEntity getMaterial_Entity(IContext contextlocal, string nuance, double Thickness)
        {
            //string material_name = null;


            try
            {
                IEntityList grades = null;
                IEntityList materials = null;
                IEntity grade = null;
                IEntity material = null;
                nuance = nuance.Replace('§', '*');
                grades = contextlocal.EntityManager.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, nuance);
                grades.Fill(false);

                if (grades.Count() > 0)
                {
                    grade = grades.FirstOrDefault();
                    materials = contextlocal.EntityManager.GetEntityList("_MATERIAL", LogicOperator.And, "_THICKNESS", ConditionOperator.Equal, Thickness, "_QUALITY", ConditionOperator.Equal, grade.Id32);//("_MATERIAL", "_QUALITY", ConditionOperator.Equal, grade.Id32);//("_MATERIAL", LogicOperator.And, "_THICKNESS", ConditionOperator.Equal, Thickness, "_QUALITY", ConditionOperator.Like, grade);
                    materials.Fill(false);
                    if (materials.Count > 0)
                    {
                        material = materials.FirstOrDefault();
                    }
                    else { material = null; }
                    //.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, nuance);

                }
                else { material = null; }


                return material;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return null;
            }
        }

        public static void get_Material_Infos(IEntity Material) {

            try {

                Material_Name = Material.GetFieldValueAsString("_NAME");
                Material_Id32 = Material.Id32;
                IEntity quality; IEntity quality_set;
                quality = Material.GetFieldValueAsEntity("_QUALITY");
                Quality_Id32 = quality.Id32;
                Quality_Name = quality.GetFieldValueAsString("_NAME");
                quality_set = quality.GetFieldValueAsEntity("_QUALITY_SET");
                Quality_set_Name = quality_set.GetFieldValueAsString("_NAME");




            } catch (Exception ie) { MessageBox.Show(ie.Message); }

        }
        /// <summary>
        /// retourne le nom complet de la nuance
        /// </summary>
        /// <param name="materialID"></param>
        /// <returns></returns>
        //public static string get_Nuance_Name(Int32 materialID) { return ""; }
        /// <summary>
        /// 
        /// revoir la mrpemier valeur trouvée de renvoier la nuance d'une matiere donnée
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        public static string get_Nuance_Name(IEntity material) {
            try {

                String rst = "";
                get_Material_Infos(material);
                rst = Quality_Name;

                return rst;
            }
            catch (Exception ie) {
                MessageBox.Show(ie.Message); return "";
            }



        }
        public static bool Exists_In_Database(String material_name) { return false; }
        public static bool Exists_In_machine_Database(String material_name, IEntity machine) { return false; }
        /// <summary>
        /// recupere le nom de la matiere d'un part
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="part"></param>
        /// <returns></returns>
        public static string getPart_Material(IContext contextlocal, IEntity part)
        {
            IEntity material;
            Int32 material_id;
            string materialname;

            try {
                material_id = part.GetFieldValueAsEntity("_MATERIAL").Id32;
                material = contextlocal.EntityManager.GetEntity(material_id, "_MATERIAL");
                materialname = material.GetFieldValueAsString("_NAME");
                return materialname;
            }
            catch (Exception ie)
            {
                Alma_Log.Error("MATIERE NON TROUVEE : ref " + part.GetFieldValueAsString("_NAME"), MethodBase.GetCurrentMethod().Name + "Reference non trouvée : import impossible: " + ie.Message);
                return "";
            }



        }
        /// <summary>
        /// retourne l'id de la matiere
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="part"></param>
        /// <returns></returns>
        public static Int32 getPart_Material_Id(IContext contextlocal, IEntity part) {
            Int32 material_id;
            try {
                material_id = part.GetFieldValueAsEntity("_MATERIAL").Id32;
                return material_id;
            }
            catch (Exception ie)
            {
                Alma_Log.Error("MATIERE NON TROUVEE : ref " + part.GetFieldValueAsString("_NAME"), MethodBase.GetCurrentMethod().Name + "Reference non trouvée : import impossible: " + ie.Message);

                return 0;
            }
        }


    }

    #region Export

    public class PartInfo
    {
        //infos  des part recupérée ---> "SI IL N Y A PAS DE SEQUENCE"
        /// <summary>
        /// retourne les infos de Geometrie et la machine par defaut d'une part
        /// </summary>
        /// <param name="Reference"></param>
        /// 
        //private int Zero_Value=0;
        //public string Reference { get;set;}
        // zone privé

        double? surface = 0;
        double? surfaceBrute = 0;
        double? weight = 0;
        double? height = 0;
        double? width = 0;
        double? thickness = 0;
        double? perimeter = 0;
        double? partTime = 0;
        double? quote_part_cyle_time = 0;
        //int? defaultMachine  = 0;

        string defaultMachineName = "";
        string emfFile = "";
        string name = "";
        string costcenter = "";
        string material = "";
        string iddevis = "";

        ///zone public
        public string Name { get { return name; } set { name = value; } }
        public double Surface { get { if (surface.HasValue) { return Convert.ToDouble(surface); } { return 0; } } set { surface = value; } }
        public double SurfaceBrute { get { if (surfaceBrute.HasValue) { return Convert.ToDouble(surfaceBrute); } { return 0; } } set { surfaceBrute = value; } }
        public double Weight { get { if (weight.HasValue) { return Convert.ToDouble(weight); } { return 0; } } set { weight = value; } }
        public double Height { get { if (height.HasValue) { return Convert.ToDouble(height); } { return 0; } } set { height = value; } }
        public double Width { get { if (width.HasValue) { return Convert.ToDouble(width); } { return 0; } } set { width = value; } }
        public double Thickness { get { if (thickness.HasValue) { return Convert.ToDouble(thickness); } { return 0; } } set { thickness = value; } }
        public double Perimeter { get { if (perimeter.HasValue) { return Convert.ToDouble(perimeter); } { return 0; } } set { perimeter = value; } }


        public double Profiles_Amount { get; set; } = 0;
        public double Internal_Profiles_Amount { get; set; } = 0;
        public double External_Profiles_Amount { get; set; } = 0;

        public double AlmaCam_PartTime { get; set; } = 0;
        public double Quote_part_cyle_time { get { if (quote_part_cyle_time.HasValue) { return Convert.ToDouble(quote_part_cyle_time); } { return 0; } } set { quote_part_cyle_time = value; } }
        //public int DefaultMachineid { get { if (defaultMachine.HasValue) { return Convert.ToInt32(defaultMachine); } { return 0; } } set { defaultMachine = value; } }
        public string DefaultMachineName { get { return defaultMachineName; } set { defaultMachineName = value; } }
        public string EmfFile { get { if (!string.IsNullOrEmpty(emfFile)) { return (emfFile.ToString()); } { return ""; } } set { emfFile = value; } }
        public string Machinablepart_emfFile { get; set; }

        public double PartTime { get { if (partTime.HasValue) { return Convert.ToDouble(partTime); } { return 0; } } set { partTime = value; } }
        public string Costcenter { get { if (!string.IsNullOrEmpty(costcenter)) { return costcenter; } { return null; } } set { costcenter = value; } }
        public string Material { get { if (!string.IsNullOrEmpty(material)) { return material; } { return null; } } set { material = value; } }
        public string Iddevis { get { if (!string.IsNullOrEmpty(iddevis)) { return iddevis; } { return null; } } set { iddevis = value; } }
        //ientity
        public IEntity Quotepart { get; set; } = null;
        public IEntity MaterialEntity { get; set; } = null;
        public IEntity DefaultMachineEntity { get; set; } = null;
        public IEntity Part_To_Produce_IEntity { get; set; } = null;

        public SpecificFields Specific_Part_Fields = new SpecificFields();
        /// <summary>
        /// renvoie les information de bases d'une reference avec les ifnos de la machinable part selectionnée
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="Part"></param>
        public void GetPartinfos_FromMachinablePart(ref IContext contextlocal, IEntity MachinablePart)
        {   //recuperation des infos de part
            //IEntity reference = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, MachinablePart.GetFieldInternalValueAsEntity("_PREPARATION"));
            IEntity reference = MachinablePart.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE");
            Name = reference.GetFieldValue("_NAME").ToString();
            IEntity material_entity = null;
            material_entity = reference.GetFieldValueAsEntity("_MATERIAL");
            MaterialEntity = material_entity;
            material = material_entity.GetFieldValueAsString("_NAME");
            thickness = reference.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS");
            //recuperation de la machine par defaut
            IEntity defaultMachine = null;
            defaultMachine = MachinablePart.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");

            IEntity centrefrais;
            centrefrais = Get_CostCenter(defaultMachine);
            Costcenter = centrefrais.GetFieldValueAsString("_CODE");
            Quotepart = reference.GetFieldValueAsEntity("_QUOTE_PART");

            if (reference.GetFieldValueAsEntity("_QUOTE_PART") != null) { Quote_part_cyle_time = Quotepart.GetFieldValueAsDouble("_CORRECTED_CYCLE_TIME"); }

            defaultMachineName = defaultMachine.GetFieldValueAsString("_NAME");
            DefaultMachineEntity = defaultMachine;
            //on recherche si le status en normal --> non oboslete et validé
            if (MachinablePart.Status.ToString() == "Normal" && MachinablePart.ValidData == true)
            {
                weight = (MachinablePart.GetFieldValueAsDouble("_WEIGHT"));
                perimeter = (MachinablePart.GetFieldValueAsDouble("_PERIMET"));
                surfaceBrute = (MachinablePart.GetFieldValueAsDouble("_SURFEXT"));
                surface = (MachinablePart.GetFieldValueAsDouble("_SURFACE"));
                height = (MachinablePart.GetFieldValueAsDouble("_DIMENS1"));
                width = MachinablePart.GetFieldValueAsDouble("_DIMENS2");
                emfFile = MachinablePart.GetImageFieldValueAsLinkFile("_PREVIEW");
                AlmaCam_PartTime = MachinablePart.GetFieldValueAsDouble("_TOTALTIME");
                getCustomPartInfos(MachinablePart);
            }




        }
        /// <summary>
        /// renvoie les information de bases sans ouverture de la piece
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="Part"></param>
        public void GetPartinfos(ref IContext contextlocal, IEntity Part)
        {   //recuperation des infos de part
            name = Part.GetFieldValue("_NAME").ToString();
            IEntity material_entity = null;
            material_entity = Part.GetFieldValueAsEntity("_MATERIAL");
            MaterialEntity = material_entity;
            material = material_entity.GetFieldValueAsString("_NAME");
            thickness = Part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS");
            //recuperation de la machine par defaut
            IEntity defaultMachine = null;
            defaultMachine = Part.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
            //on set le default machinie id;

            //recuperation de la liste des preparations pour la part
            IEntityList preparations = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, Part.Id);
            preparations.Fill(false);
            //cost center
            IEntity centrefrais;
            centrefrais = Get_CostCenter(defaultMachine);
            Costcenter = centrefrais.GetFieldValueAsString("_CODE");
            Quotepart = Part.GetFieldValueAsEntity("_QUOTE_PART");
            if (Part.GetFieldValueAsEntity("_QUOTE_PART") != null) { Quote_part_cyle_time = Quotepart.GetFieldValueAsDouble("_CORRECTED_CYCLE_TIME"); }

            foreach (IEntity preparation in preparations)
            {
                if (preparation.ImplementedEntityType.Key == "_MACHINABLE_PART")
                {
                    IEntity machinablePart = preparation.ImplementedEntity;
                    IEntity machine = machinablePart.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
                    defaultMachineName = machine.GetFieldValueAsString("_NAME");
                    DefaultMachineEntity = defaultMachine;
                    if (machine.Id == defaultMachine.Id)
                    {
                        //on recherche si le status en normal --> non oboslete et validé
                        if (machinablePart.Status.ToString() == "Normal" && machinablePart.ValidData == true)
                        {
                            weight = (machinablePart.GetFieldValueAsDouble("_WEIGHT"));
                            perimeter = (machinablePart.GetFieldValueAsDouble("_PERIMET"));
                            surfaceBrute = (machinablePart.GetFieldValueAsDouble("_SURFEXT"));
                            surface = (machinablePart.GetFieldValueAsDouble("_SURFACE"));
                            height = (machinablePart.GetFieldValueAsDouble("_DIMENS1"));
                            width = machinablePart.GetFieldValueAsDouble("_DIMENS2");
                            emfFile = machinablePart.GetImageFieldValueAsLinkFile("_PREVIEW");
                            AlmaCam_PartTime = machinablePart.GetFieldValueAsDouble("_TOTALTIME");


                            getCustomPartInfos(machinablePart);
                        }

                    }
                }
            }

        }
        /// <summary>
        /// renovie les infos avec ouverture de la iece dans le drafter
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="Part"></param>
        /// <param name="withtopologie"></param>
        public void GetPartinfos(IContext contextlocal, IEntity Part, bool withtopologie)
        {
            this.name = Part.GetFieldValue("_NAME").ToString();
            IEntity fieldValueAsEntity = Part.GetFieldValueAsEntity("_MATERIAL");
            this.MaterialEntity = fieldValueAsEntity;
            this.material = fieldValueAsEntity.GetFieldValueAsString("_NAME");
            this.thickness = new double?(Part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS"));
            IEntity fieldValueAsEntity2 = Part.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
            IEntityList entityList = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, Part.Id);
            entityList.Fill(false);
            IEntity entity = this.Get_CostCenter(fieldValueAsEntity2);
            this.Costcenter = entity.GetFieldValueAsString("_CODE");
            this.Quotepart = Part.GetFieldValueAsEntity("_QUOTE_PART");
            bool flag = Part.GetFieldValueAsEntity("_QUOTE_PART") != null;
            if (flag)
            {
                this.Quote_part_cyle_time = this.Quotepart.GetFieldValueAsDouble("_CORRECTED_CYCLE_TIME");
            }
            foreach (IEntity current in entityList)
            {
                bool flag2 = current.ImplementedEntityType.Key == "_MACHINABLE_PART";
                if (flag2)
                {
                    IEntity implementedEntity = current.ImplementedEntity;
                    IEntity fieldValueAsEntity3 = implementedEntity.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
                    this.defaultMachineName = fieldValueAsEntity3.GetFieldValueAsString("_NAME");
                    this.DefaultMachineEntity = fieldValueAsEntity2;
                    bool flag3 = fieldValueAsEntity3.Id == fieldValueAsEntity2.Id;
                    if (flag3)
                    {
                        bool flag4 = implementedEntity.Status.ToString() == "Normal" && implementedEntity.ValidData;
                        if (flag4)
                        {
                            this.weight = new double?(implementedEntity.GetFieldValueAsDouble("_WEIGHT"));
                            this.perimeter = new double?(implementedEntity.GetFieldValueAsDouble("_PERIMET"));
                            this.surfaceBrute = new double?(implementedEntity.GetFieldValueAsDouble("_SURFEXT"));
                            this.surface = new double?(implementedEntity.GetFieldValueAsDouble("_SURFACE"));
                            this.height = new double?(implementedEntity.GetFieldValueAsDouble("_DIMENS1"));
                            this.width = new double?(implementedEntity.GetFieldValueAsDouble("_DIMENS2"));
                            this.Machinablepart_emfFile = implementedEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                            if (withtopologie)
                            {
                                MachinablePart_Infos.Get_Basic_Infos_MachinablePart(ref contextlocal, implementedEntity);
                                this.Profiles_Amount = MachinablePart_Infos.Profiles_amount;
                                this.Internal_Profiles_Amount = MachinablePart_Infos.Profiles_Internal_profiles_amount;
                                this.External_Profiles_Amount = MachinablePart_Infos.Profiles_External_profiles_amount;
                            }
                            this.getCustomPartInfos(implementedEntity);
                        }
                    }
                }
            }
        }


        /// <summary>
        /// renvoie les information de bases sans ouverture de la piece
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="Part"></param>
        /// <param name="centee de frais de la ligne de piece a produire">machine concernée</param>
        public void GetPartinfos(ref IContext contextlocal, IEntity Part, IEntity Centrefrais)
        {   //recuperation des infos de part
            name = Part.GetFieldValue("_NAME").ToString();
            IEntity material_entity = null;
            Boolean preparation_exists = false;
            material_entity = Part.GetFieldValueAsEntity("_MATERIAL");
            MaterialEntity = material_entity;
            material = material_entity.GetFieldValueAsString("_NAME");
            thickness = Part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS");
            //recuperation de la machine par defaut

            //recuperation de la liste des preparations pour la part
            IEntityList preparations = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, Part.Id);
            preparations.Fill(false);
            //cost center
            //IEntity centrefrais;
            IEntityList machinelist;
            machinelist = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
            machinelist.Fill(false);

            //Get_Machine_From_Cost_Center(Centrefrais);
            //nesting_to_close = current_nesting; //nestings_list.Where(x => x.GetFieldValueAsString("_NAME") == nesting_name).FirstOrDefault();
            IEntity machine = machinelist.Where(x => x.GetFieldValueAsEntity("CENTREFRAIS_MACHINE").Id == Centrefrais.Id).FirstOrDefault();

            Costcenter = Centrefrais.GetFieldValueAsString("_CODE");
            Quotepart = Part.GetFieldValueAsEntity("_QUOTE_PART");
            if (Part.GetFieldValueAsEntity("_QUOTE_PART") != null) { Quote_part_cyle_time = Quotepart.GetFieldValueAsDouble("_CORRECTED_CYCLE_TIME"); }

            foreach (IEntity preparation in preparations)
            {
                if (preparation.ImplementedEntityType.Key == "_MACHINABLE_PART")
                {
                    IEntity machinablePart = preparation.ImplementedEntity;
                    IEntity currentmachine = machinablePart.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
                    string MachineName = currentmachine.GetFieldValueAsString("_NAME");

                    if (machine.Id == currentmachine.Id)
                    {
                        //
                        preparation_exists = true;
                        //on recherche si le status en normal --> non oboslete et validé
                        if (machinablePart.Status.ToString() == "Normal" && machinablePart.ValidData == true)
                        {
                            weight = (machinablePart.GetFieldValueAsDouble("_WEIGHT"));
                            perimeter = (machinablePart.GetFieldValueAsDouble("_PERIMET"));
                            surfaceBrute = (machinablePart.GetFieldValueAsDouble("_SURFEXT"));
                            surface = (machinablePart.GetFieldValueAsDouble("_SURFACE"));
                            height = (machinablePart.GetFieldValueAsDouble("_DIMENS1"));
                            width = machinablePart.GetFieldValueAsDouble("_DIMENS2");
                            emfFile = machinablePart.GetImageFieldValueAsLinkFile("_PREVIEW");
                            AlmaCam_PartTime = machinablePart.GetFieldValueAsDouble("_TOTAL_TIME");
                            getCustomPartInfos(machinablePart);
                        }

                    }
                }
            }


            if (preparation_exists == false)
            {   //on envoié rien et message
                //aucune prepar n'existe pour cette pieces et cette machine
                Alma_Log.Write_Log("No preparaiton exists for this machine");
                MessageBox.Show("Aucune preparation de cette piece n'existe pour cette machines, certaines données seront ignorées", " _GetPartinfos ", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }



        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //methode de recuperation des propriété dans une classe derivée//a améliorer en collant ca dans la classe geometrie
        public derivedclass As<derivedclass>() where derivedclass : PartInfo
        {
            var derivedtype = typeof(derivedclass);
            var basetype = typeof(PartInfo);
            var instance = Activator.CreateInstance(derivedtype);

            //PropertyInfo[] properties = type.GetProperties();
            PropertyInfo[] properties = basetype.GetProperties();
            foreach (var property in properties)
            {
                property.SetValue(instance, property.GetValue(this, null), null);
            }

            return (derivedclass)instance;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////
        public virtual void getCustomPartInfos(IEntity machinablepart) {; }
        /// <summary>
        /// retourne  true si la preparation exist pour la machine  par defaut 
        /// </summary>
        /// <param name="contextlocal">contexte</param>
        /// <param name="Part">entité part</param>
        /// <param name="machine">entité machine</param>
        /// <returns>true/false</returns>
        public bool IsPartDefault_Preparation(IEntity Part, IEntity machine)
        {
            if (machine != null) {
                { IEntity defaultMachine = Part.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
                    return (defaultMachine.Id32 == machine.Id32) ? true : false; }
            } else { return false; }

        }
        /// <summary>
        /// retourne  true si la preparation exist pour la machine  par defaut 
        /// </summary>
        /// <param name="contextlocal">contexte</param>
        /// <param name="Part">entité part</param>
        /// <param name="machine">nom de la machine</param>
        /// <returns>true/false</returns>
        public bool IsPartDefault_Preparation(IEntity Part, string machineName)
        {
            if (machineName != null)
            {
                IEntity defaultMachine = Part.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
                return string.Compare(defaultMachine.DefaultValue.ToString(), machineName) == 0 ? true : false;
            }
            else { return false; }

        }

        /// <summary>
        /// retourne  le nom du centre de frais de la machine demandées
        /// </summary>
        /// <param name="contextlocal">contexte</param>
        /// <param name="Part">entité part</param>
        /// <param name="machine">nom de la machine</param>
        /// <returns>true/false</returns>
        public IEntity Get_CostCenter(IEntity machine)
        {
            IEntity costcenter = null;
            if (machine != null) {
                costcenter = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
            }
            return costcenter;

        }


        public void Get_AlmaCamEstimation(IEntity machinablepart, IContext contextlocal)
        {

        }


        // ImportTools.PartInfo
        public void GetPartinfos(ref IContext contextlocal, ref IEntity Part, bool withtopologie)
        {
            this.name = Part.GetFieldValue("_NAME").ToString();
            IEntity fieldValueAsEntity = Part.GetFieldValueAsEntity("_MATERIAL");
            this.MaterialEntity = fieldValueAsEntity;
            this.material = fieldValueAsEntity.GetFieldValueAsString("_NAME");
            this.thickness = new double?(Part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS"));
            IEntity fieldValueAsEntity2 = Part.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
            IEntityList entityList = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, Part.Id);
            entityList.Fill(false);
            IEntity entity = this.Get_CostCenter(fieldValueAsEntity2);
            this.Costcenter = entity.GetFieldValueAsString("_CODE");
            this.Quotepart = Part.GetFieldValueAsEntity("_QUOTE_PART");
            bool flag = Part.GetFieldValueAsEntity("_QUOTE_PART") != null;
            if (flag)
            {
                this.Quote_part_cyle_time = this.Quotepart.GetFieldValueAsDouble("_CORRECTED_CYCLE_TIME");
            }
            foreach (IEntity current in entityList)
            {
                bool flag2 = current.ImplementedEntityType.Key == "_MACHINABLE_PART";
                if (flag2)
                {
                    IEntity implementedEntity = current.ImplementedEntity;
                    IEntity fieldValueAsEntity3 = implementedEntity.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
                    this.defaultMachineName = fieldValueAsEntity3.GetFieldValueAsString("_NAME");
                    this.DefaultMachineEntity = fieldValueAsEntity2;
                    bool flag3 = fieldValueAsEntity3.Id == fieldValueAsEntity2.Id;
                    if (flag3)
                    {
                        bool flag4 = implementedEntity.Status.ToString() == "Normal" && implementedEntity.ValidData;
                        if (flag4)
                        {
                            this.weight = new double?(implementedEntity.GetFieldValueAsDouble("_WEIGHT"));
                            this.perimeter = new double?(implementedEntity.GetFieldValueAsDouble("_PERIMET"));
                            this.surfaceBrute = new double?(implementedEntity.GetFieldValueAsDouble("_SURFEXT"));
                            this.surface = new double?(implementedEntity.GetFieldValueAsDouble("_SURFACE"));
                            this.height = new double?(implementedEntity.GetFieldValueAsDouble("_DIMENS1"));
                            this.width = new double?(implementedEntity.GetFieldValueAsDouble("_DIMENS2"));
                            this.Machinablepart_emfFile = implementedEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                            if (withtopologie)
                            {
                                MachinablePart_Infos.Get_Basic_Infos_MachinablePart(ref contextlocal, implementedEntity);
                                this.Profiles_Amount = MachinablePart_Infos.Profiles_amount;
                                this.Internal_Profiles_Amount = MachinablePart_Infos.Profiles_Internal_profiles_amount;
                                this.External_Profiles_Amount = MachinablePart_Infos.Profiles_External_profiles_amount;
                            }
                            this.getCustomPartInfos(implementedEntity);
                        }
                    }
                }
            }
        }



    }


    /// <summary>
    /// recuperation des informations de géometrie d'une piece ou d'un chute.
    /// </summary>
    public class Geometric_Infos : IDisposable
    {

        public double Surface { get; set; }//{ get{return Surface;} set{Surface=0;} }
        public double SurfaceBrute { get; set; }//{ get { return SurfaceBrute; } set { SurfaceBrute = 0; } }
        public double Weight { get; set; }//{ get { return Weight; } set { Weight = 0; } }
        public double Height { get; set; }//{ get { return Longueur; } set { Longueur = 0; } }
        public double Width { get; set; }//{ get { return Largeur; } set { Largeur = 0; } }
        public double Perimeter { get; set; }//{ get { return Perimeter; } set { Perimeter = 0; } }
        public string EmfFile { get; set; }//{ get; set; }
        public double Thickness { get; set; }//{ get{ return Thickness;} set{Thickness=0;} }
        public Int32 Material_Id { get; set; }
        public string Material_Name { get; set; }
        public string NumLot { get; set; }
        public List<Topologie> TableTopologiques { get; set; }  //sous la forme '

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }




    }



    public static class MachinablePart_Infos
    {
        public static double DimX { get; set; }//dimension x mm
        public static double DimY { get; set; }//dimension y mm
        public static double Perimeter { get; set; }//perimeter mm
        public static double Surface { get; set; }//surface mm2
        public static double Weight { get; set; }//weight kg
        public static double TotalTime { get; set; }//total machining part
        public static int defaultMachineid { get; }

        public static double Profiles_Internal_profiles_amount { get; set; }
        public static double Profiles_External_profiles_amount { get; set; }
        public static double Profiles_amount { get; set; }

        /// <summary>
        /// retourne la machinable part avec la machine par defaut. (voir pour selection des machinable part)
        /// </summary>
        /// <param name="contextlocal">context</param>
        /// <param name="reference">reference pointée</param>
        /// <returns></returns>

        public static void Get_Basic_Info_MachinablePart(IContext contextlocal, IEntity reference)

        {
            IEntity machinablePart;
            IEntityList machinableParts;
            IEntityList preparationList;

            //IEntity machine;
            machinableParts = contextlocal.EntityManager.GetEntityList("_MACHINABLE_PART");
            preparationList = contextlocal.EntityManager.GetEntityList("_PREPARATION", "_REFERENCE", ConditionOperator.Equal, reference);
            machinablePart = machinableParts.TakeWhile(x => x.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE") == reference).FirstOrDefault();
            DimX = machinablePart.GetFieldValueAsDouble("_DIMENS1");
            DimY = machinablePart.GetFieldValueAsDouble("_DIMENS2");
            Perimeter = machinablePart.GetFieldValueAsDouble("_PERIMET");
            Surface = machinablePart.GetFieldValueAsDouble("_SURFACE");
            Weight = machinablePart.GetFieldValueAsDouble("_WEIGHT");
            TotalTime = machinablePart.GetFieldValueAsDouble("_TOTALTIME");

        }

        public static void Get_Basic_Infos_MachinablePart(ref IContext contextlocal, IEntity machinablePart)
        {
            MachinablePart_Infos.DimX = machinablePart.GetFieldValueAsDouble("_DIMENS1");
            MachinablePart_Infos.DimY = machinablePart.GetFieldValueAsDouble("_DIMENS2");
            MachinablePart_Infos.Perimeter = machinablePart.GetFieldValueAsDouble("_PERIMET");
            MachinablePart_Infos.Surface = machinablePart.GetFieldValueAsDouble("_SURFACE");
            MachinablePart_Infos.Weight = machinablePart.GetFieldValueAsDouble("_WEIGHT");
            MachinablePart_Infos.TotalTime = machinablePart.GetFieldValueAsDouble("_TOTALTIME");
            Topologie topologie = new Topologie();
            topologie.GetCuttingTopologie(ref machinablePart, ref contextlocal);
            MachinablePart_Infos.Profiles_Internal_profiles_amount = topologie.Topo_Internal_Profiles_Amount;
            MachinablePart_Infos.Profiles_External_profiles_amount = topologie.Topo_External_Profiles_Amount;
            MachinablePart_Infos.Profiles_amount = topologie.Topo_Profiles_Amount;
        }

        public static void Get_Basic_Info_FromReference(IContext contextlocal, IEntity reference)

        {
            IEntity currentreference;
            //IEntityList defaultMachine;
            IEntityList referencelistList;

            //IEntity machine;
            referencelistList = contextlocal.EntityManager.GetEntityList("_REFERENCE");
            currentreference = referencelistList.TakeWhile(x => x.GetFieldValueAsEntity("_REFERENCE") == reference).FirstOrDefault();
            DimX = currentreference.GetFieldValueAsDouble("_DIMENS1");
            DimY = currentreference.GetFieldValueAsDouble("_DIMENS2");
            Perimeter = currentreference.GetFieldValueAsDouble("_PERIMET");
            Surface = currentreference.GetFieldValueAsDouble("_SURFACE");
            Weight = currentreference.GetFieldValueAsDouble("_WEIGHT");
            TotalTime = Perimeter / 1000;

        }




    }



    public class Vector : IDisposable
    {
        public double X { get { return (x2 - x1); } }//{ get{return Surface;} set{Surface=0;} }
        public double Y { get { return (y2 - y1); } }//{ get { return SurfaceBrute; } set { SurfaceBrute = 0; } }
        public double Norme { get { return Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2)); } }//{ get { return Weight; } set { Weight = 0; } }
        public double x1 { get; set; }//{ get { return Longueur; } set { Longueur = 0; } }
        public double x2 { get; set; }//{ get { return Largeur; } set { Largeur = 0; } }
        public double y1 { get; set; }//{ get { return Longueur; } set { Longueur = 0; } }
        public double y2 { get; set; }//{ get { return Largeur; } set { Largeur = 0; } }
        public double angle { get; set; }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        public double CalculateAngle() {
            double anglex; double angley; double tan;
            anglex = Math.Asin(X / Norme);
            angley = Math.Cos(Y / Norme);
            tan = Math.Atan2(Y, X);


            return 0;
        }




    }



    public class Topologie : IDisposable
    {
        public string Tool_ID { get; set; }

        public long Topo_ContoursAmount { get; set; } = 0;
        public double Topo_Perimeter { get; set; } = 0;
        public double Topo_Surface { get; set; } = 0;
        public long Topo_SharpeAnglesAmount { get; set; } = 0;
        public double Topo_PartTime { get; set; } = 0;
        public long Topo_NbAmorcages { get; set; } = 0;



        public long Topo_Profiles_Amount { get; set; } = 0;


        public long Topo_Internal_Profiles_Amount
        {
            get
            {
                return this.Topo_Profiles_Amount - this.Topo_External_Profiles_Amount;
            }
        }

        public long Topo_External_Profiles_Amount { get; set; } = 0;
        public double Topo_MarkingPerimeter { get; set; } = 0;

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        // ImportTools.Topologie
        public void GetCuttingTopologie(ref IEntity machinablepart, ref IContext contextlocal)
        {
            DrafterModule drafterModule = new DrafterModule();
            long num = 0L;
            long num2 = 0L;
            drafterModule.Init(contextlocal.Model.Id32, 1);
            drafterModule.OpenMachinablePart(machinablepart.Id32);
            int num3;
            int i = drafterModule.FirstProfile(out num3);
            while (i > 0)
            {

                // Modification Sébastien DAVID pour projet BO: Ajout de la longueur de marquage.
                if (num3 == 2)
                {
                    // Le profil est de la découpe
                    num += 1L;
                    double num4;
                    double num5;
                    double num6;
                    double num7;
                    drafterModule.GetProfileDimension(i, out num4, out num5, out num6, out num7);
                    bool flag2 = drafterModule.IsClosedProfile(i);
                    if (flag2)
                    {
                        this.Topo_Surface += drafterModule.GetProfileSurface(i);
                        this.Topo_Perimeter += drafterModule.GetProfilePerimeter(i);
                        bool flag3 = drafterModule.IsExternalProfile(i);
                        if (flag3)
                        {
                            num2 += 1L;
                        }
                    }
                    else
                    {
                        this.Topo_Surface += 0.0;
                        this.Topo_Perimeter += drafterModule.GetProfilePerimeter(i);
                    }
                    bool flag4 = drafterModule.IsRightMaterialProfile(i);
                    bool flag5 = drafterModule.IsExternalProfile(i);
                }
                if (num == 1)
                {
                    // le profil est du marquage.
                    this.Topo_MarkingPerimeter += drafterModule.GetProfilePerimeter(i);
                }
                i = drafterModule.NextProfile(out num3);
                // fin Modification Sébastien DAVID pour projet BO: Ajout de la longueur de marquage prévu pour évolution mais pas livré au client si pas demandé.
                // En effet, le document CSV généré doit être en conformité avec le document actuel fait par actcut (macro Info_DPR). Il contient dans les colonnes
                // supplémentaires les commentaires.



            }
            this.Topo_Profiles_Amount = num;
            this.Topo_NbAmorcages = num;
            this.Topo_External_Profiles_Amount = num2;
        }

        public void GetTopologie(ref IEntity machinablepart, ref IContext contextlocal) {
            //On ouvre la machinable part avec le drafter
            Actcut.ActcutModelManager.DrafterModule drafter = new Actcut.ActcutModelManager.DrafterModule();
            int tooling;
            int profile;
            int element;
            int type;
            long anglevif = 0;



            //Topo2d.Part p = new Topo2d.Part();
            drafter.Init(contextlocal.Model.Id32, 1);
            drafter.OpenMachinablePart(machinablepart.Id32);

            profile = drafter.FirstProfile(out tooling);

            //coupe
            while (profile > 0)
            {
                //coupe
                if (tooling == 2)
                {
                    double xMin; double yMin; double xMax; double yMax;
                    drafter.GetProfileDimension(profile, out xMin, out yMin, out xMax, out yMax);

                    if (drafter.IsClosedProfile(profile) == true)
                    {
                        Topo_Surface = Topo_Surface + drafter.GetProfileSurface(profile);
                        Topo_Perimeter = Topo_Perimeter + drafter.GetProfilePerimeter(profile);
                    }
                    else
                    {
                        Topo_Surface = Topo_Surface + 0;
                        Topo_Perimeter = Topo_Perimeter + drafter.GetProfilePerimeter(profile);
                    }

                    //bool isClosed = drafter.IsClosedProfile(profile);

                    bool isRightMaterial = drafter.IsRightMaterialProfile(profile);
                    bool isExternal = drafter.IsExternalProfile(profile);

                    /*

                                    if (isExternal == false && tooling == 2)
                                    {
                                        drafter.SetProfileTooling(profile, 1, 1);
                                    }
                    */

                    element = drafter.FirstElement(profile, out type);

                    while (element > 0)
                    {
                        double xStart = 0; double yStart = 0; double xEnd = 0; double yEnd = 0; double xCenter = 0; double yCenter = 0; int antiClockWise = 0; double scalaire = 0; double norme = 1;
                        double angle = 0;
                        if (type == 0)
                        {
                            drafter.GetLine(element, out xStart, out yStart, out xEnd, out yEnd);
                            //recuperation des angles vifs ici (faire un simple produit vectoriel)
                            //xx* -yy* /longeurs vecteurs = arcos (X)= angle 
                            // si angle <120 alors anglesvif = anglevif +1
                            //
                            scalaire = (xEnd - xStart) + (yEnd - yStart);
                            norme = Math.Sqrt(Math.Abs(Math.Pow(xEnd - xStart, 2) + Math.Pow((yEnd - yStart), 2)));
                            angle = Math.Acos(scalaire / norme);
                            if (angle < (120 * Math.PI / 180)) { anglevif = anglevif + 1; }
                            //

                        }
                        //anglevif = anglevif + 1;
                        else
                            drafter.GetArc(element, out xStart, out yStart, out xEnd, out yEnd, out xCenter, out yCenter, out antiClockWise);
                        element = drafter.NextElement(profile, element, out type);
                    }

                    profile = drafter.NextProfile(out tooling);
                }





            }








        }

        public static double clipper_Estimate_Part(IContext contextlocal, IEntity machinablepart)

        {
            double Total_Time = 0;
            IEntity machine;
            IEntity material;
            machine = machinablepart.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");

            material = SimplifiedMethods.Machinable_Part_Get_Implement_Material(contextlocal, machinablepart);

            AF_ImportTools.Machine_Info.GetFeeds(contextlocal, machine, material);

            return Total_Time;
        }






    }



    /// <summary>
    /// informations de part : 
    /// </summary>


    public partial class Nested_PartInfo : Geometric_Infos
    {
        //infos recuperer en private
        /// <summary>
        /// retourne les infos de Geometrie pour la machine par defaut
        /// </summary>
        /// <param name="Reference"></param>
        /// 
        //private int Zero_Value=0;
        public long Nested_Quantity { get; set; }

        public string DefaultMachineName { get; set; }
        //public double PartTime { get { return PartTime; } set { PartTime = Perimeter / 2000; } }
        public double Part_Time { get; set; }
        public string Part_Name { get; set; }
        public string Part_Reference { get; set; }
        //public double Part_Balanced_Weight { get { return Part_Balanced_Weight; } set { Part_Balanced_Weight = Weight * Ratio_Consommation; } }
        public double Part_Balanced_Weight { get; set; }
        public double Part_Balanced_Surface { get; set; }
        public long Total_Nested_Quantity { get; set; }
        public double Part_Total_Nested_Weight { get; set; }  //poinds toatl des pieces
        public double Part_Total_Nested_Weight_ratio { get; set; }  //poinds toatl des pieces /poids de la tole consommée
        public Boolean Part_IsGpao = true;  //pad defaut toutes les pieces proviennent de la gpao, les autre sont nommée pieces fantomes si Isgpao=false c'est une peice fantome
        //part to produce
        public IEntity Part_To_Produce_IEntity;
        //champs specifiques
        public SpecificFields Nested_PartInfo_specificFields = new SpecificFields();
        /// </summary>
        public double Ratio_Consommation { get; set; }
        //nom champ, valeur


        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //methode de recuperation des propriété dans une classe derivée//a améliorer en collant ca dans la classe geometric
        public derivedclass As<derivedclass>() where derivedclass : Nested_PartInfo
        {
            var derivedtype = typeof(derivedclass);
            var basetype = typeof(Nested_PartInfo);
            var instance = Activator.CreateInstance(derivedtype);

            //PropertyInfo[] properties = type.GetProperties();
            PropertyInfo[] properties = basetype.GetProperties();
            foreach (var property in properties)
            {
                property.SetValue(instance, property.GetValue(this, null), null);
            }

            return (derivedclass)instance;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// <summary>
        /// recupere les infos de part
        /// </summary>
        /// <param name="Name">nom de la piece</param>
        /// <param name="stage">etape </param>
        public virtual void GetInfos(string Name, string stage)
        {
            // IEntity Part_To_Produce;

        }

    }


    /// <summary>
    /// entité chute: peut etre remplacer par une part infos
    /// </summary>
    public class Offcut_Infos : Geometric_Infos
    {
        public long Offcut_Quantity { get; set; }
        public string Offcut_Name { get; set; }
        public double Offcut_Ratio { get; set; }


        /// <summary>
        /// selon le mode 2 ou 3 on reucpere les infos de chute dans le format ou pas
        ///    //verification de l'option du stock ou  de la creation d'une tole personnalisée à la volée 
        ///bool manageStock = ActcutModelOptions.IsManageStock(contextlocal);
        /// </summary>
        /// 

        public SpecificFields Specific_Offcut_Fields = new SpecificFields();


        public virtual void GetOffcut_Infos(IContext contextlocal, IEntity Nested_Part)
        {
            //recuperation de la machine par defaut
            //Name = Nested_Part.GetFieldValue("_NAME").ToString();
        }


        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //methode de recuperation des propriété dans une classe derivée//a améliorer en collant ca dans la classe geometrie
        public derivedclass As<derivedclass>() where derivedclass : Offcut_Infos
        {
            var derivedtype = typeof(derivedclass);
            var basetype = typeof(Offcut_Infos);
            var instance = Activator.CreateInstance(derivedtype);

            //PropertyInfo[] properties = type.GetProperties();
            PropertyInfo[] properties = basetype.GetProperties();
            foreach (var property in properties)
            {
                property.SetValue(instance, property.GetValue(this, null), null);
            }

            return (derivedclass)instance;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////



    }

    //contient le liste des 
    public class Generic_GP_Infos : IDisposable
    {
        //informations generiques aux placements
        //couple placement chutes
        public List<Nest_Infos_2> nestinfoslist = new List<Nest_Infos_2>();
        //public SpecificFields Generic_GP_Infos_specificFields;

        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //methode de recuperation des propriété dans une classe derivée//a améliorer en collant ca dans la classe geometric
        public derivedclass As<derivedclass>() where derivedclass : Generic_GP_Infos
        {
            var derivedtype = typeof(derivedclass);
            var basetype = typeof(Generic_GP_Infos);
            var instance = Activator.CreateInstance(derivedtype);

            //PropertyInfo[] properties = type.GetProperties();
            PropertyInfo[] properties = basetype.GetProperties();
            foreach (var property in properties)
            {
                property.SetValue(instance, property.GetValue(this, null), null);
            }

            return (derivedclass)instance;
        }


        ///purge auto
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
        /////construction des infos de nesting
        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <param name="stage">closed </param>
        /// <returns></returns>
        public virtual void GetNestInfosBySheet(IEntity to_cut_sheet_entity)
        {


            //creation du dictionnaire pour l'etat des tole en fonction de l'etat des placements
            Dictionary<string, string> Get_associated_Sheet_Type =
                new Dictionary<string, string>();

            Get_associated_Sheet_Type.Add("_CLOSED_NESTING", "_CUT_SHEET");
            Get_associated_Sheet_Type.Add("_TO_CUT_NESTING", "_TO_CUT_SHEET");

            //recuperation de la liste des toles coupées
            //string stage = Entity.EntityType.Key;
            //IEntityList state_sheets;


            //recuperation du nestinfos2
            Nest_Infos_2 nestinfo2data = new Nest_Infos_2();
            //set nesting id
            nestinfo2data.Nesting = to_cut_sheet_entity.GetFieldValueAsEntity("_TO_CUT_NESTING");//to_cut_sheet_entity.GetFieldValueAsEntity(Get_associated_Sheet_Type[to_cut_sheet_entity.EntityType.Key]);//to_cut_sheet_entity.GetFieldValueAsEntity("_TO_CUT_NESTING");
            nestinfo2data.NestingId = nestinfo2data.Nesting.Id;
            //on verifie si la tole a ete tournée
            nestinfo2data.IS_ROTATED = nestinfo2data.Nesting.GetFieldValueAsBoolean("_IS_ROTATED");
            //on regarde les toles une a une
            nestinfo2data.Get_NestInfos(to_cut_sheet_entity);
            //en mode bysheet la multiplicité est de 1
            nestinfo2data.Tole_Nesting.Mutliplicity = 1;
            ///nestinfo2data.GetInfos()
            nestinfo2data.Get_OffcutInfos(nestinfo2data);

            nestinfo2data.GetPartsInfos(to_cut_sheet_entity);
            //calculus
            nestinfo2data.ComputeNestInfosCalculus();
            //nestinfo2data.Get_OffcutInfos(nestinfo2data.Tole_Nesting.StockEntity);

            //section specifique
            SetSpecific_Tole_Infos(nestinfo2data.Tole_Nesting);
            SetSpecific_Part_Infos(nestinfo2data.Nested_Part_Infos_List);
            SetSpecific_Offcut_Infos(nestinfo2data.Offcut_infos_List);
            //
            if (nestinfo2data.Calculus_CheckSum_OK != true)
            {//on log

            }
            this.nestinfoslist.Add(nestinfo2data);
            nestinfo2data = null;



        }

        /////construction des infos de nesting
        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <param name="stage">closed or to cut</param>
        /// <returns></returns>
        public void GetNestInfosByNesting(IContext contextlocal, IEntity nesting, string stage)
        {
            //creation du dictionnaire pour l'etat des tole en fonction de l'etat des placements
            Dictionary<string, string> Get_associated_Sheet_Type =
                new Dictionary<string, string>();

            Get_associated_Sheet_Type.Add("_CLOSED_NESTING", "_CUT_SHEET");
            Get_associated_Sheet_Type.Add("_TO_CUT_NESTING", "_TO_CUT_SHEET");

            //recuperation de la liste des toles coupées
            //string stage = Entity.EntityType.Key;Get_associated_Sheet_Type[stage]
            IEntityList state_sheets;

            state_sheets = contextlocal.EntityManager.GetEntityList("_TO_CUT_SHEET", "_TO_CUT_NESTING", ConditionOperator.Equal, nesting.Id);
            state_sheets.Fill(false);
            //creation des nest_infos2 pour chaque cloture
            // un nestinfo2 contient la liste des pieces et des chutes generées ainsi que les calculs de ratio.
            foreach (IEntity currentsheet in state_sheets)
            {
                //recuperation du nestinfos2
                Nest_Infos_2 nestinfo2data = new Nest_Infos_2();
                //set nesting id
                nestinfo2data.Nesting = nesting;
                nestinfo2data.NestingId = nesting.Id;
                nestinfo2data.NestingMultiplicity = nesting.GetFieldValueAsLong("_QUANTITY");
                //on regarde les toles une a une
                nestinfo2data.Get_NestInfos(currentsheet);
                nestinfo2data.GetPartsInfos(currentsheet);
                ///nestinfo2data.GetInfos()
                /// recuperation de l'option WORKSHOP_OPTION  : GlobalCloseSeparated : en standard envoie coupe puis cloture
                WorkShopOptionType WORKSHOP_OPTION = ActcutModelOptions.GetWorkShopOption(contextlocal);
                //nestinfo2data.Get_OffcutInfos(nestinfo2data )
                nestinfo2data.Get_OffcutInfos(nestinfo2data, WORKSHOP_OPTION);

                //calculus
                nestinfo2data.ComputeNestInfosCalculus();
                //nestinfo2data.Get_OffcutInfos(nestinfo2data.Tole_Nesting.StockEntity);

                //section specifique
                SetSpecific_Tole_Infos(nestinfo2data.Tole_Nesting);
                SetSpecific_Part_Infos(nestinfo2data.Nested_Part_Infos_List);
                //pas d'infos de stock car le stock n'est pas créée a l'envoie a la coupe

                SetSpecific_Offcut_Infos(nestinfo2data.Offcut_infos_List);

                this.nestinfoslist.Add(nestinfo2data);
                nestinfo2data = null;

            }



            // return this.nestinfoslist;
        }

        ///  ecriture du fichier de sortie
        /// </summary>
        /// <param name="nestinfos">variables de type nestinfos2 preconstuit sur le nestinfos2</param>
        /// <param name="export_gpao_file">chemin vers le fichier de sortie</param>
        /// 
        public virtual void Export_NestInfosToFile(string export_gpao_path)
        {

        }
        //export un fichier de planning extension .planning
        public virtual void Export_NestInfosToFilePlanning(string export_gpao_path)
        {

        }

        public virtual void SetSpecific_Generic_GP_Infos(string export_gpao_path)
        {


        }
        public virtual void SetSpecific_Tole_Infos(Tole Tole)
        {

        }

        public virtual void SetSpecific_Offcut_Infos(List<Tole> Offcut_infos_List)
        {

        }

        public virtual void SetSpecific_Part_Infos(List<Nested_PartInfo> Nested_Part_Infos_List)
        {

        }
    }
    /// <summary>
    /// cette class stock les elements necessaires pour les calcul de retour gp 
    /// elle contient les information commune entre chute et tole neuves
    /// </summary>
    public class Tole : IDisposable
    {   //
        //public Int64   NestingId;   //id du nesting
        public IEntity SheetEntity { get; set; } //format
        public IEntity StockEntity { get; set; } //stock utiliser pour le placement
        public string Sheet_Name { get; set; } //nom du format  
        public double Sheet_Weight { get; set; }//poids de la tole
        public double Sheet_Length { get; set; }//long de la tole
        public double Sheet_Width { get; set; }//larg de la tole
        public double Sheet_Surface { get; set; } // surfaces  de la tole 

        /// <summary>
        /// cumul des surfaces
        /// </summary>
        public double Sheet_Total_Surface { get; set; } // surfaces  de la tole 
        public double Sheet_Total_Weight { get; set; }// surfaces  de la tole 

        public string Sheet_Reference { get; set; } //nom de la tole du stock selon son etat
        public string Stock_Name { get; set; }  //nom du stock
        public string Stock_Coulee { get; set; } // numelro de coule heat number
        public Int64 Stock_qte_initiale { get; set; }    // qte iniatale
        public Int64 Stock_qte_reservee { get; set; }    // qte reservee
        public Int64 Stock_qte_Utilisee { get; set; }    // qte utilisee

        public string To_Cut_Sheet_Name { get; set; } //nom de la tole du stock
        public string To_Cut_Sheet_NoPgm { get; set; } //nom cu programme cn
        public string To_Cut_Sheet_Pgm_Name { get; set; } //nom du programme cn        
        public string To_Cut_Sheet_Extract_FullName { get; set; } //chemin complet


        public string State_Sheet_Name { get; set; } //nom de la tole du stock selon son etat

        public Int64 Sheet_Id { get; set; }             //id du foirmat de tole (format)
        public Int64 Stock_id { get; set; }       //id du stock de tole du placement (tole reele)
        public Int32 Material_Id { get; set; }    //matiere id de la tole
        public string MaterialName { get; set; }  //matiere nom de la tole
        public IEntity Grade { get; set; }
        public string GradeName { get; set; }      //grade de la tole       
        public double Thickness { get; set; }      //epaisseur de la tole
        public long Mutliplicity { get; set; }      //multiplicité du placement
        //public long  Sheet_muliplicity { get; set; }   //multiplicité de la tole; pour l'obetnuir on regarde le nombre de stock de meme parent (sheet_id)
        public string Sheet_EmfFile { get; set; }  //apercu

        public Boolean no_Offcuts = false;// true si pas de chute //
        public Boolean no_Stock = false;// true si pas de stock //
        public Boolean Sheet_Is_rotated = false;
        public long OffcutGenerated { get; set; }  //nombre de chutes enfants

        //public SpecificFields Tole_specificFields;

        //discitonnaire de champs specifiques
        public SpecificFields Specific_Tole_Fields = new SpecificFields();


        ///purge auto
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }


    }

    /// <summary>
    /// nouvelle class derivable contenant des listes de placements avec des listes de chutes et leurs propriétés
    /// nestinfos2 contient une liste de Tole avec piece et chute correspondante
    /// Tole_Nesting : le information relatif a la tole utilisee pour le placement en cours d'etude
    /// Offcut_infos_List : contient la liste des chutes associées à la tole du placement
    /// Nested_Part_Infos_List : contient la liste des Pieces du placement
    /// </summary>
    [Obsolete]
    public class Nest_Infos_2 : IDisposable
    {
        //propriete
        public Tole Tole_Nesting { get; set; }      //placement
        public IEntity Nesting { get; set; }        //placement
        public IEntity NestingSheetEntity { get; set; }     //format
        public IEntity NestingStockEntity { get; set; }     //stock utiliser pour le placement
        public Int64 NestingId;                             //id du nesting
        ///
        public IEntity Machine_Entity;
        public Int32 DefaultMachine_Id { get; set; }        //id de la machine par defaut
        public string Nesting_MachineName { get; set; }     //nom de la machine par defaut
        public string Nesting_CentreFrais_Machine { get; set; }  //clipper machine centre de frais
        public double LongueurCoupe { get; set; } // longeur de coupe *
        public Int64  NestingMultiplicity { get; set; } = 1;    //multiplicité placement
        public double Nesting_FrontWaste { get; set; } //chute au front
        public double Nesting_TotalWaste { get; set; } //chute totale
        public double Nesting_FrontWasteKg { get; set; } //chute au front en kg
        public double Nesting_TotalWasteKg { get; set; }//chute totale en kg       
        public double Nesting_TotalTime { get; set; } //in seconds       
        public double NestingSheet_loadingTimeInit { get; set; }  //temps de chanregement        
        public double NestingSheet_loadingTimeEnd { get; set; }//temps de chanregement fin
        public Boolean IS_ROTATED = false;

       // public SpecificFields Nest_Infos_2_specificFields;
        public SpecificFields Nest_Infos_2_Fields = new SpecificFields();
        

        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   

        ///
        /// <summary>
        /// offcutlist
        /// </summary>
        public List<Tole> Offcut_infos_List { get; set; }
        /// <summary>
        /// partlist
        /// </summary>
        public List<Nested_PartInfo> Nested_Part_Infos_List = null;
        /// <summary>
        /// calculus GP
        /// </summary>
        public double Calculus_Parts_Total_Surface { get; set; }//somme des surfaces pieces 
        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   
        public double Calculus_Parts_Total_Weight { get; set; }//somme des surfaces pieces 
        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   
        public double Calculus_Parts_Total_Time { get; set; } = 0;//somme des surfaces pieces 
        public double Calculus_Offcuts_Total_Surface { get; set; } = 0;//somme des surfaces chutes
        public double Calculus_Offcuts_Total_Weight { get; set; } = 0;//somme des surfaces chutes      
        public double Calculus_Offcut_Ratio { get; set; } = 0;//somme des surfaces chutes



        //calculus
        public double Calculus_Ratio_Consommation { get; set; }
        public double Calculus_CheckSum = 1;
        public Boolean Calculus_CheckSum_OK = false;
        ///purge auto
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
        #region parts
        
        /// <summary>
        /// calcul de laliste des pieces placée dans le placement
        /// atention, les pieces fantome ne sont pas prise en compte
        /// </summary>
        /// <param name="nestedpart"></param>
        public void Get_NestedPartInfos(IEntity nestedpart)
        {
            //piece par toles
           
            IEntity machinable_Part = null;
            IEntity to_produce_reference = null;
    
            Nested_PartInfo nested_Part_Infos = new Nested_PartInfo();
            //
          
        
            nested_Part_Infos.Part_To_Produce_IEntity = nestedpart.GetFieldValueAsEntity("_TO_PRODUCE_REFERENCE");
            //on set matiere et epaisseur a celle du nesting
            nested_Part_Infos.Material_Id = Tole_Nesting.Material_Id; //   Sheet_Material_Id;
            nested_Part_Infos.Material_Name = Tole_Nesting.MaterialName; //  Sheet_MaterialName;
            nested_Part_Infos.Thickness = Tole_Nesting.Thickness; //Sheet_Thickness;

            //recuperation des infos du part to produce

            nested_Part_Infos.Part_Time = nestedpart.GetFieldValueAsDouble("_TOTALTIME");
            nested_Part_Infos.Nested_Quantity = nestedpart.GetFieldValueAsLong("_QUANTITY");
            nested_Part_Infos.Nested_Quantity = nestedpart.GetFieldValueAsLong("_QUANTITY");
            //repercution des infos de machinable part
            machinable_Part = nestedpart.GetFieldValueAsEntity("_MACHINABLE_PART");
            nested_Part_Infos.Surface = machinable_Part.GetFieldValueAsDouble("_SURFACE");
            nested_Part_Infos.Part_Total_Nested_Weight = nested_Part_Infos.Surface * nested_Part_Infos.Nested_Quantity;
            nested_Part_Infos.SurfaceBrute = machinable_Part.GetFieldValueAsDouble("_SURFEXT");
            nested_Part_Infos.Weight = machinable_Part.GetFieldValueAsDouble("_WEIGHT");
            nested_Part_Infos.Part_Total_Nested_Weight = nested_Part_Infos.Weight * nested_Part_Infos.Nested_Quantity;
            //nested_Part_Infos.EmfFile = machinable_Part.GetImageFieldValueAsLinkFile("_PREVIEW");
            //nested_Part_Infos.EmfFile = SimplifiedMethods.GetPreview(@machinable_Part.GetImageFieldValueAsLinkFile("_PREVIEW"), machinable_Part);
            nested_Part_Infos.EmfFile = SimplifiedMethods.GetPreview(machinable_Part);
            nested_Part_Infos.Width = machinable_Part.GetFieldValueAsDouble("_DIMENS1");
            nested_Part_Infos.Height = machinable_Part.GetFieldValueAsDouble("_DIMENS2");

            //reference to produce
            to_produce_reference = nestedpart.GetFieldValueAsEntity("_TO_PRODUCE_REFERENCE");
            nested_Part_Infos.Part_Reference = to_produce_reference.GetFieldValueAsString("_NAME");
            nested_Part_Infos.Part_Name= to_produce_reference.GetFieldValueAsString("_NAME");
            //custom_Fields
            //Nested_Part_Info.Custom_Nested_Part_Infos

            //nested_Part_Infos.c
            //ajout des methodes specifiques
            Get_NestedPart_CustomInfos(to_produce_reference, nested_Part_Infos);

            //calcul de la surface total des pieces
            //on ne somme que les pieces qui ont un uid gpao (numero de gamme ou autre..)
            if (nested_Part_Infos.Part_IsGpao == true)
            {
                Calculus_Parts_Total_Surface += nested_Part_Infos.Surface * nested_Part_Infos.Nested_Quantity;
                Calculus_Parts_Total_Weight += nested_Part_Infos.Weight * nested_Part_Infos.Nested_Quantity ;
                Calculus_Parts_Total_Time += nested_Part_Infos.Part_Time * nested_Part_Infos.Nested_Quantity ;
                //ajout à la liste les pieces qui ne sont pas de sieces fantomes
                Nested_Part_Infos_List.Add(nested_Part_Infos);
            }
           
            //tole
           // IEntity offcut_IEntity;





        }
        
        
        #endregion

        #region offcut
        // public virtual void Get_OffcutInfos(IEntity NestingStockEntity)
        public virtual void Get_OffcutInfos(Nest_Infos_2 CurrentNesting)
        {


            //recuperation des chute de meme parent stock
            //IEntityList parentstocklist;
            IEntityList sheets, stocks;
            Offcut_infos_List = new List<Tole>();
            /*
            parentstocklist = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_PARENT_STOCK", ConditionOperator.Equal, CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
            parentstocklist.Fill(false);*/
            //recuperation du sheet du placement 
            if (CurrentNesting.Tole_Nesting.StockEntity != null) { 
            sheets = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);
            sheets.Fill(false);

            //construction de la liste des chutes
            foreach (IEntity sheet in sheets)
            {
                stocks = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet.Id);///NestingStockEntity.Id);
                stocks.Fill(false);


                foreach (IEntity offcut in stocks)
                {

                    Tole offcut_tole = new Tole();
                    offcut_tole.StockEntity = offcut;
                    //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                    offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                    offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                    offcut_tole.Thickness = Tole_Nesting.Thickness;  //

                    offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                    offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;

                    ///sheet
                    offcut_tole.SheetEntity = offcut.GetFieldValueAsEntity("_SHEET");
                    offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                    offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                    offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                    offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                    //pour la tole totalsurface = surface

                    offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                    offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                    offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                    //pour la tole totalweight= weigth

                    //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                    offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                    offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                    /////
                    if (offcut != null)
                    {
                        ////stock 
                        offcut_tole.StockEntity = offcut;
                        ///////on egalise la multiplicité avec celle de la tole mere (a verifier si fiable)
                        offcut_tole.Mutliplicity = CurrentNesting.Tole_Nesting.Mutliplicity;
                        offcut_tole.Stock_Name = offcut.GetFieldValueAsString("_NAME");
                        offcut_tole.Stock_Coulee = offcut.GetFieldValueAsString("_HEAT_NUMBER");
                        offcut_tole.Stock_qte_initiale = offcut.GetFieldValueAsInt("_QUANTITY");
                        offcut_tole.Stock_qte_reservee = offcut.GetFieldValueAsInt("_BOOKED_QUANTITY");
                        offcut_tole.Stock_qte_Utilisee = offcut.GetFieldValueAsInt("_USED_QUANTITY");

                        Tole_Nesting.no_Offcuts = false;
                        Tole_Nesting.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                        //////
                        Offcut_infos_List.Add(offcut_tole);
                    }
                    else { Tole_Nesting.no_Offcuts = true; }
                }
            }

            }



        }
        public virtual void Get_OffcutInfos(Nest_Infos_2 CurrentNesting,WorkShopOptionType Workshop_Option)
        {
            //creation de la liste des futures toles de type chute
            Offcut_infos_List = new List<Tole>();

            switch (Workshop_Option) { 
            
                //fermeture tole a tole
                case   WorkShopOptionType.GlobalCloseOneClic:

                    //recuperation des infos commune des chutes chutes de meme parent stock
                    //IEntityList parentstocklist;
                    /*
                    parentstocklist = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_PARENT_STOCK", ConditionOperator.Equal, CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
                    parentstocklist.Fill(false);*/

                    IEntityList sheets, stocks;
                    sheets = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);///NestingStockEntity.Id);
                    sheets.Fill(false);
                    //construction de la liste des chutes
                    foreach (IEntity sheet in sheets)
                    {

                        stocks = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet.Id);///NestingStockEntity.Id);
                        stocks.Fill(false);


                        foreach (IEntity offcut in stocks)
                        {

                            Tole offcut_tole = new Tole();
                            offcut_tole.StockEntity = offcut;
                            //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                            offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                            offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                            offcut_tole.Thickness = Tole_Nesting.Thickness;  //
                            offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                            offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;
                                                                            ///sheet
                            offcut_tole.SheetEntity = offcut.GetFieldValueAsEntity("_SHEET");
                            offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                            offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                            offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                            offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                            offcut_tole.Sheet_Total_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE") * CurrentNesting.NestingMultiplicity;

                            //pour la tole totalsurface = surface

                            offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                            offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                            offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                            offcut_tole.Sheet_Total_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT") * CurrentNesting.NestingMultiplicity;

                            //pour la tole totalweight= weigth

                            //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                            offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                            offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                            /////
                            if (offcut != null)
                            {
                                ////stock 
                                offcut_tole.StockEntity = offcut;
                                ///////on egalise la multiplicité avec celle de la tole mere (a verifier si fiable)
                                offcut_tole.Mutliplicity = CurrentNesting.Tole_Nesting.Mutliplicity;
                                offcut_tole.Stock_Name = offcut.GetFieldValueAsString("_NAME");
                                offcut_tole.Stock_Coulee = offcut.GetFieldValueAsString("_HEAT_NUMBER");
                                offcut_tole.Stock_qte_initiale = offcut.GetFieldValueAsInt("_QUANTITY");
                                offcut_tole.Stock_qte_reservee = offcut.GetFieldValueAsInt("_BOOKED_QUANTITY");
                                offcut_tole.Stock_qte_Utilisee = offcut.GetFieldValueAsInt("_USED_QUANTITY");

                                Tole_Nesting.no_Offcuts = false;
                                Tole_Nesting.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                                //////
                                Offcut_infos_List.Add(offcut_tole);
                            }
                            else { Tole_Nesting.no_Offcuts = true; }
                        }
                    }


                    break;

                case WorkShopOptionType.GlobalCloseSeparated:

                    IEntityList sheetList = CurrentNesting.Nesting.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);//CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
                    sheetList.Fill(false);
                    //construction de la liste des chutes
                    foreach (IEntity sheet in sheetList)
                    {
                        Tole offcut_tole = new Tole();
                        offcut_tole.StockEntity = null;
                        
                        //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                        offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                        offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                        offcut_tole.Thickness = Tole_Nesting.Thickness;  //
                        offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                        offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;
                        offcut_tole.Mutliplicity = Tole_Nesting.Mutliplicity;
                        //sheet
                        offcut_tole.SheetEntity = sheet; ///offcut.GetFieldValueAsEntity("_SHEET");
                        offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                        offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                        offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                        offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                        offcut_tole.Sheet_Total_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE")* CurrentNesting.NestingMultiplicity;
                        //pour la tole totalsurface = surface

                        offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                        offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                        offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                        offcut_tole.Sheet_Total_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT") * CurrentNesting.NestingMultiplicity;
                        //pour la tole totalweight= weigth

                        //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                        offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                        offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                        ///// plus d'infos de toel car elle n'existe pas
                        offcut_tole.no_Offcuts = true;
                        ///jamais de stock dans ce mode 
                        ///
                        offcut_tole.no_Stock = true;
                        Offcut_infos_List.Add(offcut_tole);
                    }


                    break;
                default:
                throw new Exception("l'Option atelier choisie n'est pas compatible avec la librairie generique d'export GP");
                break;
            }

        }
        #endregion

        #region nestinfs
        /// <summary>
        /// recupere les nestinfos infos de la tole mere
        /// </summary>
        /// <param name="currentsheet">t</param>
        /// <param name="stage"></param>

        public virtual void Get_NestInfos(IEntity to_cut_sheet)
        {
            //IEntity nesting; // string stage;
            //nesting = Nesting;
            //affectation des infos de tole du nesting
            this.Tole_Nesting = new Tole();
            this.NestingId = Nesting.Id;
            //this.IS_ROTATED = Nesting.GetFieldValueAsBoolean("_IS_ROTATED");
            string stage = Nesting.EntityType.Key;
            IContext contextelocal = this.Nesting.Context;
            Tole_Nesting.Mutliplicity = NestingMultiplicity;
            Tole_Nesting.To_Cut_Sheet_Name = to_cut_sheet.GetFieldValueAsString("_NAME");

            ////////////////////////////////////////////
            //recuperation du format de la tole du placement
            if (Nesting.GetFieldValueAsEntity("_SHEET") != null)
            {
                //IEntity sheet = Nesting.GetFieldValueAsEntity("_SHEET"); 
                Tole_Nesting.SheetEntity = Nesting.GetFieldValueAsEntity("_SHEET");
                Tole_Nesting.Sheet_Id = Tole_Nesting.SheetEntity.Id32;
                Tole_Nesting.Sheet_Name = Tole_Nesting.SheetEntity.GetFieldValueAsString("_NAME");
                Tole_Nesting.Sheet_Weight = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                Tole_Nesting.Sheet_Length = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                Tole_Nesting.Sheet_Width = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                Tole_Nesting.Sheet_Surface = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_SURFACE");

                //pour la tole support on a poids = total poids si multiplicité =1 ce qui est le cas dans les clotures toles à tole
                Tole_Nesting.Sheet_Total_Weight = Tole_Nesting.Sheet_Weight ;
                //pour la tole support on a surface = total surface si multiplicité =1 ce qui est le cas dans les clotures toles à tole
                Tole_Nesting.Sheet_Total_Surface = Tole_Nesting.Sheet_Surface  ;
            
                Tole_Nesting.Sheet_Reference = Tole_Nesting.SheetEntity.GetFieldValueAsString("_REFERENCE");
                Tole_Nesting.no_Offcuts = true;
                Tole_Nesting.Sheet_Is_rotated = this.IS_ROTATED;
          

            }

          
            ///information programme cn
            ///
            IEntityList programCns;
            IEntity programCn;
            programCns = Nesting.Context.EntityManager.GetEntityList("_CN_FILE", "_SEQUENCED_NESTING", ConditionOperator.Equal, NestingId);
            programCn = SimplifiedMethods.GetFirtOfList(programCns);

            if (programCn != null) {
                this.Tole_Nesting.To_Cut_Sheet_NoPgm = programCn.GetFieldValueAsString("_NOPGM");
                this.Tole_Nesting.To_Cut_Sheet_NoPgm = programCn.GetFieldValueAsString("_NOPGM");
                this.Tole_Nesting.To_Cut_Sheet_Pgm_Name = programCn.GetFieldValueAsString("_NAME");
                this.Tole_Nesting.To_Cut_Sheet_Extract_FullName = programCn.GetFieldValueAsString("_EXTRACT_FULLNAME");
            }
            else {
                this.Tole_Nesting.To_Cut_Sheet_NoPgm = "0";
                this.Tole_Nesting.To_Cut_Sheet_Pgm_Name = Tole_Nesting.Sheet_Name;
                this.Tole_Nesting.To_Cut_Sheet_Extract_FullName = Tole_Nesting.Sheet_Name;
            }
            
            // information du placement
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///NESTING LAYOUT///
            /////////////////)///////////////////////////////////////////////////////////////////////////////////////////////////////////////

           //creation de la preview et recuperation du chemin
            Tole_Nesting.Sheet_EmfFile = SimplifiedMethods.GetPreview(Nesting);
            this.Nesting_TotalTime = Nesting.GetFieldValueAsDouble("_TOTALTIME");
           

            LongueurCoupe = Nesting.GetFieldValueAsDouble("_CUT_LENGTH");
            Nesting_FrontWaste = Nesting.GetFieldValueAsDouble("_FRONT_WASTE");
            Nesting_FrontWaste = Nesting.GetFieldValueAsDouble("_TOTAL_WASTE");
            //multiplicite interdite en mode 3 : closing by sheet on force a 1         

            ///validation matiere
            IEntity material = Nesting.GetFieldValueAsEntity("_MATERIAL");
            Tole_Nesting.Material_Id = material.Id32;
            Tole_Nesting.MaterialName = material.GetFieldValueAsString("_NAME");
            Tole_Nesting.Thickness = material.GetFieldValueAsDouble("_THICKNESS");

            //recuperation des grades
            Int32 gradeid = material.GetFieldValueAsInt("_QUALITY");

            IEntityList grades = null;
            //IEntity grade = null;
            grades = Nesting.Context.EntityManager.GetEntityList("_QUALITY", "ID", ConditionOperator.Equal, gradeid);
            Tole_Nesting.Grade = SimplifiedMethods.GetFirtOfList(grades);
            Tole_Nesting.GradeName = Tole_Nesting.Grade.GetFieldValueAsString("_NAME");

            ///////////////////////////////////////////////////////////////////////////////////
            //machine -->
            IEntityList machineliste;
            IEntity machine;
                       
            Machine_Entity = Nesting.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
            machineliste = Nesting.Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "ID", ConditionOperator.Equal, Machine_Entity.Id32);
            //machineliste.Fill(false);
            machine = SimplifiedMethods.GetFirtOfList(machineliste);

            //recuperation des certains parametre de la ressource
            ICutMachineResource parameterList = AF_ImportTools.SimplifiedMethods.GetRessourceParameter(machine);
            //POUR L INSTANT ON CHARGE LES PARAMETRES DE CHARGERMENT AU DECHARGEMENT

            NestingSheet_loadingTimeInit = parameterList.GetSimpleParameterValueAsDouble("PAR_TPSCHARG");
            NestingSheet_loadingTimeEnd = parameterList.GetSimpleParameterValueAsDouble("PAR_TPSDECHARG");
            

            Nesting_MachineName = machine.GetFieldValueAsString("_NAME");
            DefaultMachine_Id = machine.Id32;

            //recuperation du centre de frais
            IEntity centrefrais;
            centrefrais = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
            Nesting_CentreFrais_Machine = centrefrais.GetFieldValueAsString("_CODE");
            
            centrefrais = null;
            machine = null;
            machineliste = null;

            ////////////////////////////////////////////////////////

            ////recuperation des infos de stock
            /*information sur le stock de clipper*/
            IEntityList stocklist;
            IEntity stock;
            stocklist = Nesting.Context.EntityManager.GetEntityList("_STOCK", "ID", ConditionOperator.Equal, to_cut_sheet.GetFieldValueAsInt("_STOCK"));
            stock = SimplifiedMethods.GetFirtOfList(stocklist);

            Tole_Nesting.StockEntity = stock;
            if (stocklist.Count > 0) { 
            ////stock 
            Tole_Nesting.Stock_Name = stock.GetFieldValueAsString("_NAME");
            Tole_Nesting.Stock_Coulee = stock.GetFieldValueAsString("_HEAT_NUMBER");
            Tole_Nesting.Stock_qte_initiale = stock.GetFieldValueAsInt("_QUANTITY");
            Tole_Nesting.Stock_qte_reservee = stock.GetFieldValueAsInt("_BOOKED_QUANTITY");
            Tole_Nesting.Stock_qte_Utilisee = stock.GetFieldValueAsInt("_USED_QUANTITY");
            
            }
            stocklist = null;
            stock = null;
        }
        /// <summary>
        /// calcul les ratios...
        /// </summary>
        #region calculus
        public virtual void ComputeNestInfosCalculus()
        {
            //
            int accuracy = 5; //nombre de chiffre apres la virgule
            //calcul de la surface total des chutes
            Calculus_Offcuts_Total_Surface = 0;
            Calculus_Offcuts_Total_Weight = 0;

            if (Offcut_infos_List!=null) {

                Calculus_Offcuts_Total_Surface = Offcut_infos_List.Sum(o => o.Sheet_Surface);
                Calculus_Offcuts_Total_Weight = Offcut_infos_List.Sum(o => o.Sheet_Weight);
            }

           
            
            Calculus_Parts_Total_Surface = Nested_Part_Infos_List.Sum(o => o.Surface * o.Nested_Quantity);

            //calculus
            if ((Tole_Nesting.Sheet_Total_Surface - Calculus_Offcuts_Total_Surface) != 0)
            {
                Calculus_Ratio_Consommation = (((Tole_Nesting.Sheet_Total_Surface - Calculus_Offcuts_Total_Surface) * Tole_Nesting.Mutliplicity) / Calculus_Parts_Total_Surface);
            }

            //eciture des poids corrigés
            foreach (Nested_PartInfo p in Nested_Part_Infos_List)
            {
                if (Calculus_Ratio_Consommation != 0)
                {
                    p.Ratio_Consommation = Calculus_Ratio_Consommation;
                    p.Part_Balanced_Weight = Math.Round(p.Weight * Calculus_Ratio_Consommation, accuracy);
                    p.Part_Balanced_Surface = Math.Round(p.Surface * Calculus_Ratio_Consommation, accuracy);
                    Calculus_CheckSum += p.Weight * Calculus_Ratio_Consommation * p.Nested_Quantity;
                    p.Part_Total_Nested_Weight = p.Part_Balanced_Weight * p.Nested_Quantity;
                    p.Part_Total_Nested_Weight_ratio = p.Part_Total_Nested_Weight / Calculus_Offcuts_Total_Weight;

                }
                else
                {
                    p.Ratio_Consommation = 1;
                    p.Part_Balanced_Weight = p.Weight;
                    p.Part_Balanced_Surface = p.Surface;
                    p.Part_Total_Nested_Weight = p.Weight * 1 * p.Nested_Quantity;
                    Calculus_CheckSum = 0;

                }
            }


            //checksum des poids
            Calculus_CheckSum = Calculus_CheckSum - (Tole_Nesting.Sheet_Weight - Calculus_Offcuts_Total_Weight);

            //if (Calculus_CheckSum - (Tole_Nesting.Sheet_Weight - Calculus_Offcuts_Total_Weight) < 1)
            if (Math.Round(Calculus_CheckSum, accuracy) == 1)
            {
                Calculus_CheckSum_OK = true;
            }



        }
        #endregion


        /// <summary>
        /// NestingName : exporte les données dans le streamwriter
        /// </summary>
        /// <param name="context"></param>
        /// <param name="to_cut_sheet"> entite a exporte voué a disparaitre</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        /// <param name="export_gpao_file"></param>
        public virtual void Export_NestInfosToFile(IContext context, IEntity to_cut_sheet, string stage, StreamWriter export_gpao_file)
        {
        }
        /// <summary>
        /// ecrit le fichier de retour de l'export du dosssier technique
        /// stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// </summary>
        /// <param name="context">context par reference</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        /// <param name="nestingname">nom du placement</param>
        /// <param name="exportfile">stream vers le fichier d'export</param>
        public virtual void Export_NestInfosToFile(ref IContext context, string stage, string nestingname, StreamWriter exportfile)
        {


        }

        //public virtual void Export_NestInfosToFile(IContext context, IEntity to_cut_sheet, StreamWriter export_gpao_file)   {
        public virtual void Export_NestInfosToFile(IContext context, IEntity nesting_entity, StreamWriter export_gpao_file)
        {


        }

        /// <summary>
        ///  ecriture du fichier de sortie
        /// </summary>
        /// <param name="nestinfos">variables de type nestinfos2 preconstuit sur le nestinfos2</param>
        /// <param name="export_gpao_file">chemin vers le fichier de sortie</param>
        public virtual void Export_NestInfosToFile(StreamWriter export_gpao_file)
        {


        }
        /// <summary>
        /// stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// 
        /// </summary>
        /// <param name="nesting_sheet">entité tole de placement</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        // public void GetPartsInfos(IEntity nesting_sheet, string stage) //IEntity Nesting)
        public void GetPartsInfos(IEntity to_cut_sheet_entity) //IEntity Nesting)
        {

            //recuperation des infos de sheet
            //IEntity to_cut_sheet;

            Nested_Part_Infos_List = new List<Nested_PartInfo>();
            IEntity current_nesting;

            //IEntity stock_Sheet;

            //
            current_nesting = Nesting;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///PARTS///
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //nested Parts-lists--> nestedpartinfos
            //on recherche les parts pointant sur le nesting//
            IEntityList nestedparts = null;

            //si stock managé on regarde les pieces placées sur la tole, sinn, sur le nesting
            //Actcut.ActcutModel.ActcutModelOptions.IsManagePartSet(contextlocal)…..stock managé, on regarde ls tocut reference sinon les proprités du nesting
            if (Actcut.CommonModel.ActcutModelOptions.IsManageStock(current_nesting.Context))
            { 
                nestedparts = current_nesting.Context.EntityManager.GetEntityList("_TO_CUT_REFERENCE", "_TO_CUT_SHEET", ConditionOperator.Equal, to_cut_sheet_entity.Id32);
                nestedparts.Fill(false);
            }
            else
            {
              nestedparts = current_nesting.Context.EntityManager.GetEntityList("_NESTED_REFERENCE", "_NESTING", ConditionOperator.Equal, current_nesting.Id32);
              nestedparts.Fill(false);
            }


            foreach (IEntity nestedpart in nestedparts)
            {///recuperation de la liste des pieces
                Get_NestedPartInfos(nestedpart);


            }


            //calculus
            //ComputeNestInfosCalculus();
        }
        #endregion
        #region virtual methodes

        //custom field infos
        public virtual void SetSpecific_Generic_NestInfos2() {}
        public virtual void Get_NestedPart_CustomInfos(IEntity nestedpart, Nested_PartInfo nestedpartinfos) { }
        public virtual void Get_Offcut_CustomInfos(IEntity offcut, Offcut_Infos offcutinfos) { }
        public virtual void Set_Offcut_CustomInfos(IEntity offcut, Offcut_Infos offcutinfos) { }
        public virtual void Get_NestInfos_CustomInfos(Tole Tole_nesting) { }

        #endregion




    }
    
    public class SpecificFields
    {
        private Dictionary<string, object> _dict=new Dictionary<string, object>() ;

       

        public void Add<T>(string key , object Value)
        {
            if (Value != null) { this._dict.Add(key, Value); }
            
            else { _dict.Add(key, "Undef"); }

        }

        public void Set(string key, object Value)
        {
            this._dict[key]= Value;
        }
        public void Remove(string key, object Value)
        {
            this._dict.Remove(key);
        }
        public bool Get<T>(string key, out T value)
        {   //
            object result;
            if (this._dict.TryGetValue(key, out result) && result is T)
            {
                value = (T)result;
                return true;
            }
            value = default(T);
            return false;
            

        }
    }

    #endregion
    #region description des Gp_Sheet_Infos : information susceptibles d'etre retournées pour les gp
    /// <summary>
    /// gp sheet contient l'equivalent d'un agi (piece, tole chute)
    /// 
    /// </summary>
    public class Gp_Sheet_Infos : IDisposable

    {

        //infos standards des sheets //
        //public Boolean Sheet_Is_Rotated = false;
        public double Sheet_Length { get; set; }//long de la tole
        public double Sheet_Width { get; set; }//larg de la tole
        public string Sheet_NumLot { get; set; }//lot de la tole
        public string Sheet_NumMatLot { get; set; }//matiere lottie de la tole
        public string Sheet_Magasin { get; set; }//matiere lottie de la tole
        public string Sheet_CastNumber { get; set; }    //matiere cast number= gismenet en francais
        public Int32 Sheet_Id { get; set; }             //id de la tole
        public Int32 Sheet_Material_Id { get; set; }    //matiere id de la tole
        public string Sheet_MaterialName { get; set; }  //matiere nom de la tole
        public string Sheet_Grade { get; set; }     //grade de la tole       
        public double Sheet_Thickness { get; set; } //epaisseur de la tole
        public long Sheet_Mutliplicity { get; set; }//multiplicité de la tole
        public string Sheet_EmfFile { get; set; }   //apercu
        //delacaration de varibale standards

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }



    }




    #endregion
    /*A prevoir */
    #region Import
    class Import_Stock : IDisposable
    {
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
        public void from_Csv()
        {



        }



    }



    #endregion



/// <summary>
/// //les classes statics
/// 
/// </summary>
/// 
#region machine
public static class Machine_Info
    {
        /// <summary>
        /// retourne la machine par defaut de la part
        /// </summary>
        /// <param name="reference">ientityt reference</param>
        /// <param name="defaultmachine">out : resultat de la machine par defaut</param>
        /// <returns>true si machine non null</returns>
        public static bool GetDefaultMachine(IEntity reference, out IEntity defaultmachine)
        {
            defaultmachine = null;
            Boolean rst = false;

            try
            {

                if (reference != null)
                {
                    {
                        defaultmachine = reference.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
                        if (defaultmachine != null)
                        {
                            rst = true;
                        }


                    }
                }
                return rst;
            }
            catch
            {
                return false;
            }

        }


        public static double GetFeed(IContext contextlocal, IEntity machine, IEntity material)
        {
            double feed_value = 0.01;


            try
            {

                IMachineManager machineManager = new MachineManager();
                IEntity cutMachineEntity = machine;//machineList.First();
                ICutMachine cutMachine = machineManager.GetCutMachine(contextlocal, cutMachineEntity.Id);
                // Que fait cette ligne?
                IEntity cutMachineMaterial = material;//materialList.First();
                IEntity cutMachineCuttingCondition = cutMachine.ConditionEntityList.First();
                IEntity cutMachineToolClass = cutMachine.ToolClassEntityList.First();
                IEntity cutMachineChamferAngle = cutMachine.AngleEntityList.First();

                ICutMachineResource cutMachineResource = cutMachine.GetCutMachineResource(true);
                object feed = 0;
                //recuperation de la vitesse;
                //recuperation de l'outils par defaut
                feed = cutMachineResource.GetParameterValue("EV_VITESSE", cutMachineMaterial, cutMachineCuttingCondition, cutMachineToolClass, cutMachineChamferAngle);

                feed = cutMachineResource.GetParameterValue("EV_VITESSE", cutMachineMaterial, cutMachineCuttingCondition, cutMachineToolClass, cutMachineChamferAngle);
                feed_value = Convert.ToDouble(feed);
                // Montrer comment on obtient la clé, expliquer le nombre de 
                //paramètre et leur ordre
                // Comment obtenir un double à partir de l'objet? (Ici je triche avec un ToString parce que je n'ai besoin que de l'affichage)

                string mach = cutMachineEntity.GetFieldValueAsString("_NAME");
                string mat = cutMachineMaterial.GetFieldValueAsString("_NAME");
                string cond = cutMachineCuttingCondition.GetFieldValueAsString("_NAME");
                string tool = cutMachineToolClass.GetFieldValueAsString("_NAME");
                string angle = cutMachineChamferAngle.GetFieldValueAsString("_NAME");


                if (machine != null)
                {



                }
                return feed_value;

            }
            catch { return feed_value; }

        }


        public static Dictionary<string, double> GetFeeds(IContext contextlocal, IEntity machine, IEntity material)
        {
            double feed_value = 0.01;
            Dictionary<string, double> dfeed = new Dictionary<string, double>();

            try
            {

                IMachineManager machineManager = new MachineManager();
                IEntity cutMachineEntity = machine;//machineList.First();
                ICutMachine cutMachine = machineManager.GetCutMachine(contextlocal, machine.Id);
                // Que fait cette ligne?
                IEntity cutMachineMaterial = material;//materialList.First();
                IEntity cutMachineCuttingCondition = cutMachine.ConditionEntityList.First();
                IEntity cutMachineToolClass = cutMachine.ToolClassEntityList.First();
                IEntity cutMachineChamferAngle = cutMachine.AngleEntityList.First();

                ICutMachineResource cutMachineResource = cutMachine.GetCutMachineResource(true);
                object feed = 0;
                // IEntity ToolClass = null;


                //recuperation de la vitesse;
                //cutting condition???
                //recuperation de l'outils par defaut
                feed = cutMachineResource.GetParameterValue("EV_VITESSE", cutMachineMaterial, cutMachineCuttingCondition, cutMachineToolClass, cutMachineChamferAngle);
                feed = cutMachineResource.GetParameterValue("EV_VITESSE", cutMachineMaterial, cutMachineCuttingCondition, cutMachineToolClass, cutMachineChamferAngle);
                feed_value = Convert.ToDouble(feed);

                dfeed.Add(cutMachineToolClass.ToString(), feed_value);

                // Montrer comment on obtient la clé, expliquer le nombre de 
                //paramètre et leur ordre
                // Comment obtenir un double à partir de l'objet? (Ici je triche avec un ToString parce que je n'ai besoin que de l'affichage)

                string mach = cutMachineEntity.GetFieldValueAsString("_NAME");
                string mat = cutMachineMaterial.GetFieldValueAsString("_NAME");
                string cond = cutMachineCuttingCondition.GetFieldValueAsString("_NAME");
                string tool = cutMachineToolClass.GetFieldValueAsString("_NAME");
                string angle = cutMachineChamferAngle.GetFieldValueAsString("_NAME");


                if (machine != null)
                {



                }
                return dfeed;

            }
            catch { return dfeed; }

        }

        // ImportTools.Machine_Info
        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nestingstate"></param>
        /// <param name="nesting_name"></param>
        /// <returns></returns>
        public static IEntity GetNestingMachineEntity(ref IContext contextlocal, string nestingstate, string nesting_name)
        {
            IEntity result;
            try
            {
                IEntityList entityList = contextlocal.EntityManager.GetEntityList(nestingstate, "_NAME", ConditionOperator.Equal, nesting_name);
                entityList.Fill(false);
                IEntity fieldValueAsEntity = SimplifiedMethods.GetFirtOfList(entityList).GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
                result = fieldValueAsEntity;
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log("Pas de type techno detecté " + ex.Message);
                result = null;
            }
            return result;
        }

        // ImportTools.Machine_Info
        // ImportTools.Machine_Info
        /// <summary>
        /// 
        /// </summary>
        /// <param name="technologyid"></param>
        /// <returns></returns>
        public static string GetNestingTechnologyName(long technologyid)
        {
            string result;
            try
            {
                string text = "undef";
                TechnoTypeInfo.GetAllTechnoType();
                foreach (KeyValuePair<long, string> current in TechnoTypeInfo.GetAllTechnoType())
                {
                    bool flag = current.Key == technologyid;
                    if (flag)
                    {
                        text = current.Value;
                        break;
                    }
                }
                result = text;
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log("Pas de type techno detecté " + ex.Message);
                result = "undef";
            }
            return result;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        public static string GetNestingTechnologyName(ref IContext Contextlocal, ref IEntity nesting)
        {
            string result;
            try
            {
               // string text = "";
                string key = nesting.EntityType.Key;
                IEntity nestingMachineEntity = Machine_Info.GetNestingMachineEntity(ref Contextlocal, key, nesting.GetFieldValueAsString("_NAME"));
                string nestingTechnologyName = Machine_Info.GetNestingTechnologyName(nestingMachineEntity.GetFieldValueAsLong("_TECHNOLOGY"));
                result = nestingTechnologyName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                result = null;
            }
            return result;
        }

        // ImportTools.Machine_Info
        public static IEntity GetDefaultCondition(ref IContext contextlocal, IEntity machine)
        {
            IEntity result;
            try
            {
                IEntityList entityList = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_CONDITION", "_CUT_MACHINE_TYPE", ConditionOperator.Equal, machine.Id);
                IEntity firtOfList = SimplifiedMethods.GetFirtOfList(entityList);
                result = firtOfList;
            }
            catch
            {
                Alma_Log.Write_Log_Important("Condition de coupe introuvable pour la machine selectionnée");
                result = null;
            }
            return result;
        }



    }



    #endregion

    #region marKing
    public static class MarKing {


            public static void Marque(IEntity reference, string fieldKeytoMark, string mqimageText) { 
                try{

                int contourNumber;
                string text = "";
                string textetowrite = "";
                double height; double coordx; double coordy; double angle;
               
                string texttorereplace = mqimageText ?? "MMMMMMMMM";
                ///recuperation du texte
                ///
                object ovalue = reference.GetFieldValue(fieldKeytoMark);
                textetowrite = AF_ImportTools.Data_Model.ConvertToString(ovalue).Substring(0, texttorereplace.Length);
                /////recup de la piece a produire associée
                DrafterModule df = new DrafterModule();
                IEntityList machinableParts = reference.Context.EntityManager.GetEntityList("PREPARATION");
                IEntity machinablePart = machinableParts.TakeWhile(x => x.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE") == reference).FirstOrDefault();
                df.Mqimaj();
                df.SaveMachinablePart(true);
                            
                contourNumber= df.FirstText(out text,out height,out coordx, out coordy, out angle);

                while (!text.Contains(texttorereplace))
                {
                    contourNumber=df.NextText(out text, out height, out coordx, out coordy, out angle);        
                   
                    df.AddMacroText( text,  height,  coordx,  coordy,  angle);
                }
                df.DeleteProfile(contourNumber);
                df.SaveMachinablePart(true);

                }
        
                catch (Exception ie){MessageBox.Show(ie.Message);
                    }
                
            }

    //IEntityList offcut_List = contexlocal.EntityManager.GetEntityList("_SHEET","_SEQUENCED_NESTING", ConditionOperator.Equal, nesting.Id);
    //long nested_parts_produced;
    //long offcut_
    //lest des offcut
    /*
    DrafterModule df = new DrafterModule();
    df.Mqimaj();
    df.FirstText(toto,200,0,0);
    df.AddMacroText(tutu, 200, 0, 0);
    */

}


    #endregion
    #region static methods
    /// <summary>
    /// classe static simplifiant les methode almacam
    /// </summary>
    public static class SimplifiedMethods
    {


        public static string EntityType;
        public static IContext contexlocal;
        public static string ExtendedEntityPath;
        // <summary>
        /// retourne le chemin de la prevue et la creer a la volée
        /// </summary>
        /// <param name="pathTopreview"></param>
        /// <param name="EntityToDraw"></param>
        /// <returns></returns>
        public static string GetPreview( IEntity EntityToDraw)
        {

            try
            {
                string path = EntityToDraw.GetImageFieldValueAsLinkFile("_PREVIEW");
                if (Directory.Exists(@Path.GetDirectoryName(@path)))
                {
                    if (File.Exists(@path) == false)
                    {
                        EntityToDraw.GetFieldValueInFile("_PREVIEW", ref @path);
                    }

                    return path;// pathTopreview;}}

                }

                else {
                    MessageBox.Show("Impossible de creer les EMF : veuillez configurer le dossier de sortie des emf dans le wpm.ini.");
                    //int exitCode;
                    return "";
                }

     

            }

            catch { return ""; }

        }




        // <summary>
        /// retourne le chemin de la prevue et la creer a la volée
        /// </summary>
        /// <param name="pathTopreview"></param>
        /// <param name="emfoutputdir">doccier de sorti des emf</param>
        /// <returns></returns>
        public static string GetPreview(IEntity EntityToDraw, string emfoutputdir)
        {

            try
            {
                string path = EntityToDraw.GetImageFieldValueAsLinkFile("_PREVIEW");
                path = emfoutputdir + "" + Path.GetFileName(path);
                if (Directory.Exists(@Path.GetDirectoryName(@path)))
                {
                    if (File.Exists(@path) == false)
                    {
                        EntityToDraw.GetFieldValueInFile("_PREVIEW", ref @path);
                    }

                    return path;// pathTopreview;}}

                }

                else
                {
                    MessageBox.Show("Impossible de creer les EMF : veuillez configurer le dossier de sortie des emf dans le wpm.ini.");
                    //int exitCode;
                    return "";
                }



            }

            catch { return ""; }

        }
        /// <summary>
        /// creer et copie un emf sous un autre nom
        /// </summary>
        /// <param name="EntityToDraw">entité</param>
        /// <param name="newemffilename">nouveau nom de l emf</param>
        /// <returns></returns>
        public static string CreateAndCopyPreview(IEntity EntityToDraw, string new_emf_filename)
        {

            try
            {
                string path = EntityToDraw.GetImageFieldValueAsLinkFile("_PREVIEW");
                //string folder = Path.GetDirectoryName(path);
                string newpath = Path.GetDirectoryName(path) +"\\"+ new_emf_filename;
                if (Directory.Exists(@Path.GetDirectoryName(@newpath)))
                {
                    if (File.Exists(@newpath) == false)
                    {
                        EntityToDraw.GetFieldValueInFile("_PREVIEW", ref @newpath);
                    }
                    else if (File.Exists(@newpath) == true)
                    {//update
                        File.Delete(newpath);
                        EntityToDraw.GetFieldValueInFile("_PREVIEW", ref @newpath);
                    }

                    return newpath;// pathTopreview;}}

                }

                else
                {
                    MessageBox.Show("Impossible de creer les EMF : veuillez configurer le dossier de sortie des emf dans le wpm.ini.");
                    //int exitCode;
                    return "";
                }



            }

            catch { return ""; }

        }
        /// <summary>
        /// recupere la fenetre de selection de l'entité selon l'intitulé du type d'entité
        /// </summary>
        /// <param name="contextlocal">icontext par reference</param>
        /// <param name="entitype">du type "REFERENCE" .... "GEOMETRY"</param>
        /// <returns></returns>

        public static void GetEntitySelector(ref IContext contextlocal, string entitype, ref List<IEntity> selectedEntityList, bool allow_Multiselection)
        {
            //IEntity[] ientitytable = null;
            selectedEntityList = new List<IEntity>();
            
            IEntitySelector xselector = null;
            xselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            xselector.Init(contextlocal, contextlocal.Kernel.GetEntityType(entitype));
            xselector.MultiSelect = allow_Multiselection;


            if (xselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //int index = 0;

                foreach (IEntity xentity in xselector.SelectedEntity)
                {
                                      selectedEntityList.Add(xentity);
                }



            }

            

           

        }
        
        
        
        
        /// <summary>
        /// retourn la liste des chutes générées par le nesting
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        /// 
        public static IEntityList GetOffcutList(ref IContext contextlocal, IEntity nesting)
        {
            try
            {
                IEntityList offcutlist = null;
                offcutlist = contextlocal.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting.Id);
                offcutlist.Fill(false);

                return offcutlist;
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }
        }



        /// <summary>
        /// retourn le stock pour un format donné
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        public static IEntityList GetStockList(ref IContext contextlocal, IEntity sheet_format)
        {
            try
            {
                IEntityList stocklist = null;

               
                stocklist = contextlocal.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet_format.Id);
                stocklist.Fill(false);

                return stocklist;
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }
        }


        /// <summary>
        /// retourn la liste des pieces générées par le nesting
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        public static IEntityList GetnestedPartList(ref IContext contextlocal, IEntity nesting)
        {
            try
            {
                IEntityList nestedpartlist = null;
                nestedpartlist = contextlocal.EntityManager.GetEntityList("_NESTED_REFERENCE", "_NESTING", ConditionOperator.Equal, nesting.Id);
                nestedpartlist.Fill(false);

                return nestedpartlist;
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }
        }
        /// <summary>
        /// recupere la premiere entité d'une liste d'entité
        /// </summary>
        /// <param name="entitylist"></param>
        /// <returns></returns>
        /// 
        public static IEntity GetFirtOfList(IEntityList entitylist)
        {  IEntity returned_entity=null;
            try {
                entitylist.Fill(false);
                if (entitylist.Count() > 0) { returned_entity = entitylist.FirstOrDefault(); }
                return returned_entity; }

            catch {
                Alma_Log.Error( ":Probleme de recuperation de la premiere entité de liste", MethodBase.GetCurrentMethod().Name);
               // Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name ":_ID non detecté sur la ligne a importée, line ignored"); result = false;

                return returned_entity; }

        }
        /// <summary>
        /// rertourn la liste des toles dispo pour une tole paticuliere
        /// </summary>
        /// <param name="contextlocal">context</param>
        /// <param name="Sheet">tole </param>
        /// <returns></returns>
        /// 
        public static int Get_RemainingQuantity(ref IContext contextlocal, IEntity Sheet)

        {
            //IContext contextlocal = Sheet.Context;
            IEntityList stockList = contextlocal.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, Sheet.Id);

            stockList.Fill(false);
            int sumInitialQuantity = 0;
            int sumUsedQuantity = 0;
            int QuantityInProduction = Sheet.GetFieldValueAsInt("_IN_PRODUCTION_QUANTITY");

            foreach (IEntity stock in stockList)
            {
                sumInitialQuantity = sumInitialQuantity + stock.GetFieldValueAsInt("_QUANTITY");
                sumUsedQuantity = sumUsedQuantity + stock.GetFieldValueAsInt("_USED_QUANTITY");
            }

            int remainingQuantity = sumInitialQuantity - sumUsedQuantity;

            if (remainingQuantity < 0)
            { remainingQuantity = 0; }


            return remainingQuantity;



        }


        /// <summary>
        /// retourne une extentend entity list
        /// </summary>
        /// <param name="key">Nom de la table</param>
        /// <param name="stringqueryvalue">requete</param>
        /// <returns></returns>
        public static IExtendedEntityList Extended_List_compute_Equal(string key, string stringqueryvalue)
        {
            try
            {

                IEntityType entityType = contexlocal.Kernel.GetEntityType(EntityType);
                IExtendedEntityType extendedEntityType = entityType.ExtendedEntityType;

                IConditionType conditionType1 = contexlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                extendedEntityType.GetExtendedField(@ExtendedEntityPath),
                ConditionOperator.Equal,
                contexlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter(key, stringqueryvalue));

                IQueryType queryType = new QueryType(contexlocal.Kernel, "Compute_query", entityType);
                queryType.SetFilter(conditionType1);

                IExtendedEntityList l = contexlocal.EntityManager.GetExtendedEntityList(queryType); // On creer une liste avec le resultat de la requete
                l.Fill(false);
                return l;


            }

            catch (Exception ie)
            {
                Alma_Log.Error(ie.Message, "compute Extended entityList");
                return null;
            }
        }

        /// <summary>
        /// retourne une extentend entity list
        /// </summary>
        /// <param name="key">Nom de la table</param>
        /// <param name="intquervalue">requete</param>
        /// <returns></returns>
        public static IExtendedEntityList Extended_List_compute_Equal(string key, int intqueryvalue)
        {
            try
            {

                IEntityType entityType = contexlocal.Kernel.GetEntityType(EntityType);
                IExtendedEntityType extendedEntityType = entityType.ExtendedEntityType;

                IConditionType conditionType1 = contexlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                extendedEntityType.GetExtendedField(@ExtendedEntityPath),
                ConditionOperator.Equal,
                contexlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter(key, intqueryvalue));

                IQueryType queryType = new QueryType(contexlocal.Kernel, "Compute_query", entityType);
                queryType.SetFilter(conditionType1);

                IExtendedEntityList l = contexlocal.EntityManager.GetExtendedEntityList(queryType); // On creer une liste avec le resultat de la requete
                l.Fill(false);

                return l;


            }

            catch (Exception ie)
            {
                Alma_Log.Error(ie.Message, "compute Extended entityList");
                return null;
            }
        }

        /// <summary>
        /// retourne une chaineVide sur les string null
        /// </summary>
        /// <param name="stringvalue"></param>
        /// <returns></returns>
        public static string ConvertNullStringToEmptystring(string stringvalue) {
          
            if (string.IsNullOrEmpty(stringvalue)) { stringvalue = ""; }
            return stringvalue;

         }

        /// <summary>
        ///  retourne une chaine Vide à la place d'une string null
        /// </summary>
        /// <param name="keyname">clé</param>
        /// <param name="dictionnary">dictionnaire</param>
        /// <returns></returns>
        public static string ConvertNullStringToEmptystring(string keyname, ref Dictionary<string, object> dictionnary)
        {

            {
                string item = null;
                if (AF_ImportTools.Data_Model.ExistsInDictionnary(keyname, ref dictionnary) && dictionnary[keyname].GetType() == typeof(string)) { item = dictionnary[keyname].ToString(); }
                else { return item=""; }
                return item;

            }
        }
        /// <summary>
        /// return Implemented reference pointed by the machinablepart
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="machinablePart">entity machinable part</param>
        /// <returns>entity </returns>
        public static IEntity Machinable_Part_Get_Implement_Reference(IContext contextlocal, IEntity machinablePart)
        {
            IEntity ie;
            ie = machinablePart.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE");
            return ie;

        }

        public static void Machinable_Part_Clean_content(ref IContext contextlocal, IEntity machinablepart)
        {
            DrafterModule drafterModule = new DrafterModule();
            drafterModule.Init(contextlocal.Model.Id32, 1);
            drafterModule.OpenMachinablePart(machinablepart.Id32);
            int num3;
            int i = drafterModule.FirstProfile(out num3);
            drafterModule.Clean();
        }

           
        /// <summary>
        /// Marge un champs ou plusieur champs de l'entité en utilisant le mqimage
        /// </summary>
        /// <param name="machinablpart">machiaablepart entity</param>
        /// <param name="fieldstoMark">list on string : if lis is emty the function will only place MMMMM</param>
        /// <returns></returns>
        public static Boolean Machinable_Part_MQimage(ref IContext contextlocal ,IEnumerable <IEntity> machinablparts, IEnumerable<string> fieldstoMark)
        {
            try {
                DrafterModule dm = new DrafterModule();
                foreach (var item in machinablparts)
                {
                    //
                    dm.Init(contextlocal.Model.Id32, Convert.ToInt32(item.GetFieldValueAsLong("_CUT_MACHINE_TYPE")));
                    dm.OpenMachinablePart( item.Id32);


                    long num = 0L;
                    //long num2 = 0L;
                    int num3;
                    int i = dm.FirstProfile(out num3);
                    ////sauvegarde des profiles
                    while (i > 0)
                    {
                        if (num == 1)
                        {
                            //
                            //var p = new Point();
                            
                            // le profil est du marquage.
                           /// this.Topo_MarkingPerimeter += ///dm.GetProfilePerimeter(i);
                           /// 
                           
                        }

                        i = dm.NextProfile(out num3);


                    }

                    dm.Mqimaj();






                    dm.SaveMachinablePart(true);
                    
                    /*
                    foreach (var fieldvalue in fieldstoMark)
                    {

                    }*/

                    }
                

                return true;

            } catch { return false; }

            
        }
        /// <summary>
        /// return Implemented material pointed by the machinablepart
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="machinablePart">entity machinable part</param>
        /// <returns>entity </returns>
        public static IEntity Machinable_Part_Get_Implement_Material(IContext contextlocal, IEntity machinablePart)
        {
            IEntity ie;
              ie = machinablePart.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE").GetFieldValueAsEntity("_MATERIAL");

            return ie;

        }

        /// <summary>
        /// return Implemented material pointed by the machinablepart
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="machinablePart">entity machinable part</param>
        /// <returns>entity </returns>
        public static IEntity Machinable_Part_Get_Implement_Defaultmachine(IContext contextlocal, IEntity machinablePart)
        {
            IEntity ie;
            ie = machinablePart.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE").GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE");
            return ie;

        }



        /// <summary>
        /// creer la preparation associée
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="reference"></param>
        /// <param name="ListMaChine"></param>
        public static void CreateMachinablePartFromReference(IContext contextlocal,IEntity reference, IEntity machine)

        {
            try
            {
                bool alreadyCreated=false;
                IMachineManager machineManager = new MachineManager();
                ActcutReferenceManager actcutReferencemanager = new ActcutReferenceManager();
                ICutMachine cutMachine = machineManager.GetCutMachine(contextlocal, machine.Id);
                ICutMachineInfo cutMachineInfo = cutMachine.CutMachineInfo;
                IDictionary<long, IEntity> dictCutCondition = new Dictionary<long, IEntity>();

                actcutReferencemanager.GetReferenceMachinablePartList(contextlocal, reference);
                // la preparation existe elle deja //
                foreach ( IEntity preparation in  actcutReferencemanager.GetReferenceMachinablePartList(contextlocal, reference))
                    {
                    // preparation//

                    if(preparation.GetFieldValueAsEntity("_MATERIAL") == reference.GetFieldValueAsEntity("_MATERIAL"))
                    {
                        alreadyCreated = true;

                    }

                    //         

                }                

                if (alreadyCreated == false)
                {
                    ///creation de la preparation
                    actcutReferencemanager.CreateMachinablePart(contextlocal, reference, machine, GetCutMachineCuttingCondition(contextlocal, cutMachineInfo, reference.GetFieldValueAsEntity("_MATERIAL"), dictCutCondition), reference.GetFieldValueAsString("_NAME"), 0, false);
                }
                ///
            }

               
            
            catch
            {

            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <param name="cutMachineInfo"></param>
        /// <param name="material"></param>
        /// <param name="dictCutCondition"></param>
        /// <returns></returns>
        public static IEntity GetCutMachineCuttingCondition(IContext context, ICutMachineInfo cutMachineInfo, IEntity material, IDictionary<long, IEntity> dictCutCondition)
        {
            IEntity cutMachineCuttingCondition = null;
            if (dictCutCondition.TryGetValue(material.Id, out cutMachineCuttingCondition) == false)
            {
                cutMachineCuttingCondition = cutMachineInfo.GetDefaultCondition(material); //Return null if the machine is not allowed to cut the material
                dictCutCondition.Add(material.Id, cutMachineCuttingCondition);
            }

            return cutMachineCuttingCondition;
        }


        /// <summary>
        /// met le boolean exported 
        /// ajoute la date d'export
        /// string fieldKey = stage + "_GPAO_Exported"; 
        /// string fieldKey2 = stage + "_GPAO_Exported_Dte";
        /// stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// </summary>
        /// <param name="nesting">inetity nesting contenant les champs </param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        public static void MarqueAsExported(ref IEntity nesting, string stage)
        {
            string fieldKey = stage.Substring(1, stage.Length-1) + "_GPAO_Exported";
            string fieldKey2 = stage.Substring(1, stage.Length-1) + "_GPAO_Exported_Dte";
            IField field = null;
            nesting.EntityType.TryGetField(fieldKey, out field);
            bool flag = field != null;
            if (flag)
            {
                nesting.SetFieldValue(fieldKey, true);
                field = null;
            }
            nesting.EntityType.TryGetField(fieldKey, out field);
            bool flag2 = field != null;
            if (flag2)
            {
                nesting.SetFieldValue(fieldKey2, DateTime.Now.ToString("yyMMdd hh:mm:ss"));
                field = null;
            }
            nesting.Save();
        }
        /// <summary>
        /// recupere le nom du champ a checker.
        /// </summary>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        /// <returns></returns>
        public static string Get_Marqued_FieldName(string stage)
        {

            try{
                string marquedFieldname;
                marquedFieldname=stage.Substring(1, stage.Length - 1) + "_GPAO_Exported";
                return marquedFieldname; }
            catch (Exception ie) { MessageBox.Show("Error "+ie.Message); return ""; }
        }

        /// <summary>
        /// cloture automatique
        /// stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        public static bool CloseNesting( IContext contextlocal, IEntity nesting) {
            try
            {
        
           
                /// recupere les tole en prod du nesting
                /// on set les qté utilisees used_quantité dans le stock et on sauve
                ///  on deplace  les  BookNestingSheetData en affectant les valeurs de qte necessaire au placement
                ///  on set les rejected list a 0
              
                bool rst = false;
                //On recupere la tole du placement
                IEntity CurrentNestingSheet = nesting.GetFieldValueAsEntity("_SHEET");
                long multiplicity = nesting.GetFieldValueAsLong("_QUANTITY");
                string nesting_name = nesting.GetFieldValueAsString("_NAME");
                
                IEntityList current_nesting_list = contextlocal.EntityManager.GetEntityList(nesting.EntityType.Key, "ID", ConditionOperator.Equal, nesting.Id);
                current_nesting_list.Fill(false);
               

                /*on travail sur les liste des nestind avec le meme id*/
                /*rebut à 0*/
                RejectedNestingListData rejectedPartNestingListData = new RejectedNestingListData(contextlocal, current_nesting_list);
                foreach (RejectedNestingData rejectedPartNestingData in rejectedPartNestingListData.RejectedPartNestingList)
                {
                    if (rejectedPartNestingData.NestingEntity.Id == nesting.Id)
                    {
                        foreach (RejectedPartData rejectedPartData in rejectedPartNestingData.RejectedPartDataList)
                        {
                           rejectedPartData.RejectedQuantity =0;
                        }
                    }
                }

                IActcutNestingManager actcutNestingManager = new Actcut.ActcutModelManager.ActcutNestingManager();

                /*tole*/
                //verification de l'option du stock ou  de la creation d'une tole personnalisée à la volée 
                bool manageStock = ActcutModelOptions.IsManageStock(contextlocal);
                 if ( manageStock == true && nesting.GetFieldValueAsLong("_SHEET") != 0)
                {

                    //cette liste a prereservée des toles
                    

                    BookNestingSheetData bookSheetToNestingData = new BookNestingSheetData(contextlocal, current_nesting_list, true);

                    foreach (BookSheetData bookSheetData in bookSheetToNestingData.BookSheetDataList)
                    {

                        ///on set les quantités reservées à 0    POUR REINITIALISER LES TOLE DEJA RESERVEE
                        ///sheet list --> liste des toles du stock
                        foreach (StockData stockData in bookSheetData.SheetList) { stockData.Quantity = 0; }
                        //quantité de tole a debiter
                        // sequenced_nesting_list.het
                        long reservedQty = 0;
                        //stockdata contient la toles courante du stock
                        //bookSheetData.SheetList
                        ///ici on parcours les toles du stock et on reserve les quantités à debiter dans le stock
                        foreach (StockData stockData in bookSheetData.SheetList)
                        {//on reserve une tole de chaque element de stock jusqu'a ce que le compte soit atteint.
                            if (stockData.StockDataItem.AvailableQuantity > 1 && reservedQty <bookSheetData.Quantity)
                            {
                                //
                                if (stockData.StockDataItem.AvailableQuantity < bookSheetData.Quantity - reservedQty)
                                {
                            
                                    stockData.Quantity = stockData.StockDataItem.AvailableQuantity;
                                    reservedQty += stockData.Quantity;
                                }
                                else
                                {
                                    stockData.Quantity = bookSheetData.Quantity - reservedQty;
                                    reservedQty += stockData.Quantity;
                                    break;
                                }
                            }
                        }
                        if (reservedQty < bookSheetData.Quantity)
                        {

                        }

                        if (bookSheetData.NestingEntity.GetFieldValueAsEntity("_SHEET") == null)
                        {

                        }
                        else
                        {
                            //nestingEntity = bookSheetData.NestingEntity;
                            //bookSheetData.SheetList.First().Quantity = nestingEntity.GetFieldValueAsLong("_QUANTITY");
                        }
                            
                    }

                        /*cloture avec toles specifiées*/
                        actcutNestingManager.CloseNesting(contextlocal, current_nesting_list, bookSheetToNestingData, rejectedPartNestingListData);

                }
                else
                {
                    /*cloture sur tole créée a la volée ne faisant pas partie du stock*/
                    actcutNestingManager.CloseNesting(contextlocal, current_nesting_list, null, rejectedPartNestingListData);

                }

                /*cloture*/
                //actcutNestingManager.CloseNesting(contextlocal, sequenced_nesting_list, bookSheetToNestingData, rejectedPartNestingListData);
                //actcutNestingManager.CloseNesting(contextlocal, sequenced_nesting_list, bookSheetToNestingData, null);
                return rst;

                
            }


            catch (Exception ie) { MessageBox.Show(ie.Message);return false; }
        }
        /// <summary>
        /// cloture automatique
        /// stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="nesting"></param>
        /// <returns></returns>
        public static bool CloseNestingClipper(IContext contextlocal, IEntity nesting)
        {
            try
            {


                /// recupere les tole en prod du nesting
                /// on set les qté utilisees used_quantité dans le stock et on sauve
                ///  on deplace  les  BookNestingSheetData en affectant les valeurs de qte necessaire au placement
                ///  on set les rejected list a 0

                bool rst = false;
                //On recupere la tole du placement
                IEntity CurrentNestingSheet = nesting.GetFieldValueAsEntity("_SHEET");
                long multiplicity = nesting.GetFieldValueAsLong("_QUANTITY");
                string nesting_name = nesting.GetFieldValueAsString("_NAME");
                
                
                IEntityList current_nesting_list = contextlocal.EntityManager.GetEntityList(nesting.EntityType.Key, "ID", ConditionOperator.Equal, nesting.Id);
                current_nesting_list.Fill(false);
                

                /*on travail sur les liste des nestind avec le meme id*/
                /*rebut à 0*/
                RejectedNestingListData rejectedPartNestingListData = new RejectedNestingListData(contextlocal, current_nesting_list);
                foreach (RejectedNestingData rejectedPartNestingData in rejectedPartNestingListData.RejectedPartNestingList)
                {
                    if (rejectedPartNestingData.NestingEntity.Id == nesting.Id)
                    {
                        foreach (RejectedPartData rejectedPartData in rejectedPartNestingData.RejectedPartDataList)
                        {
                            rejectedPartData.RejectedQuantity = 0;
                        }
                    }
                }

                IActcutNestingManager actcutNestingManager = new Actcut.ActcutModelManager.ActcutNestingManager();

                /*tole*/
                //verification de l'option du stock ou  de la creation d'une tole personnalisée à la volée 
                bool manageStock = ActcutModelOptions.IsManageStock(contextlocal);
                // bool manageStock = contextlocal.ParameterSetManager.GetParameterValue("_GLOBAL_CONFIGURATION", "_MANAGE_STOCK").GetValueAsBoolean();
                if (manageStock == true && nesting.GetFieldValueAsLong("_SHEET") != 0)
                {

                    //current_nesting_list

                    //verfication de la demande en tole du placement


                    //verification des tole dispos
                    BookNestingSheetData bookSheetToNestingData = new BookNestingSheetData(contextlocal, current_nesting_list, true);
                    IEntity nestingEntity = null;

                    //bookSheetToNestingData.BookSheetDataList : il s'agit de la liste des tole dispo pour un format donnée
                    //bookSheetData contient l'associatin placement tole du stock dans sheetlist
                    // reservation des toles bookSheetData dispos dans lA liste BookSheetDataList ->LISTE DES TOLES DE MEME FORMAT DANS LE  STOCK 
                    
                    foreach (BookSheetData bookSheetData in bookSheetToNestingData.BookSheetDataList)
                    {

                        ///on set les quantités reservées à 0    POUR REINITIALISER LES TOLE DEJA RESERVEE
                        ///sheet list --> liste des toles du stock
                        foreach (StockData stockData in bookSheetData.SheetList) { stockData.Quantity = 0; }
                        //quantité de tole a debiter
                        // sequenced_nesting_list.het
                        long reservedQty = 0;
                        //stockdata contient la toles courante du stock
                        //bookSheetData.SheetList
                        ///ici on parcours les toles du stock et on reserve les quantités à debiter dans le stock
                        foreach (StockData stockData in bookSheetData.SheetList)
                        {//on reserve une tole de chaque element de stock jusqu'a ce que le compte soit atteint.
                            if (stockData.StockDataItem.AvailableQuantity > 1 && reservedQty < bookSheetData.Quantity)
                            {
                                //
                                if (stockData.StockDataItem.AvailableQuantity < bookSheetData.Quantity - reservedQty)
                                {

                                    stockData.Quantity = stockData.StockDataItem.AvailableQuantity;
                                    reservedQty += stockData.Quantity;
                                }
                                else
                                {
                                    stockData.Quantity = bookSheetData.Quantity - reservedQty;
                                    reservedQty += stockData.Quantity;
                                    break;
                                }
                            }
                        }
                        if (reservedQty < bookSheetData.Quantity)
                        {

                        }

                        if (bookSheetData.NestingEntity.GetFieldValueAsEntity("_SHEET") == null)
                        {

                        }
                        else
                        {
                            //nestingEntity = bookSheetData.NestingEntity;
                            //bookSheetData.SheetList.First().Quantity = nestingEntity.GetFieldValueAsLong("_QUANTITY");
                        }

                    }
                    
                    /*cloture avec toles specifiées*/

                    
                    //on regle le quantités de tole //
                    /*actcutNestingManager.CloseNesting(contextlocal, current_nesting_list, bookSheetToNestingData, rejectedPartNestingListData);*/
                    // on force la cloture // //on regle les quanttie de pieces placée// 
                    actcutNestingManager.CloseNesting(contextlocal, current_nesting_list, null, rejectedPartNestingListData);
                    
                    
                  

                }
                else
                {
                    /*cloture sur tole créée a la volée ne faisant pas partie du stock*/
                    actcutNestingManager.CloseNesting(contextlocal, current_nesting_list, null, rejectedPartNestingListData);

                }

                /*cloture*/
                //actcutNestingManager.CloseNesting(contextlocal, sequenced_nesting_list, bookSheetToNestingData, rejectedPartNestingListData);
                //actcutNestingManager.CloseNesting(contextlocal, sequenced_nesting_list, bookSheetToNestingData, null);
                return rst;


            }


            catch (Exception ie) { MessageBox.Show(ie.Message); return false; }
        }


        public static ICutMachineResource GetRessourceParameter(IEntity machine)
        {
            try { 
            //IEntityList listeResources = machine.Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
            //listeResources.Fill(false);
            
            IMachineManager machinemanager = new MachineManager();
            ICutMachineResource ressource = machinemanager.GetCutMachineResource(machine.Context, machine, false);
            return ressource;
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); return null; }
        }

        /// <summary>
        /// converti en double
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static double GetDouble(string value, double defaultValue)
        {
            double result;

            //Try parsing in the current culture
            if (!double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out result) &&
                //Then try in US english
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result) &&
                //Then in neutral language
                !double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result = defaultValue;
            }

            return result;
        }


        /// <summary>
        /// converti en double en usi: a tester
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static double GetDoubleInvariantCulture(string value)
        {
            double result;
            //Try parsing in the current culture
            double.TryParse(value, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out result);
            return result;
        }


    }




    #endregion

    #region File_tools
    public static class File_Tools
        {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        ////// ImportTools.File_Tools
        public static void CreateDirectory(string path)
        {
            try
            {
                bool flag = !Directory.Exists(path);
                if (flag)
                {
                    Directory.CreateDirectory(path);
                    Alma_Log.Write_Log(path + " Créé");
                }
            }
            catch (Exception ex)
            {
                Alma_Log.Write_Log(ex.Message);
            }
        }
        public static Boolean replaceFile(string filepath, Boolean killIfexists) {

            Boolean replaced=false;

            if (File.Exists(filepath))
            {
                if (killIfexists == false)
                {
           
                    DialogResult dr = MessageBox.Show("Le fichier " + filepath + " existe déjà \r\n. Voulez-vous le remplacer?",
                          "Fichier existant", MessageBoxButtons.YesNo);
                    switch (dr)
                    {
                        case DialogResult.Yes:
                            File.Delete(filepath);
                            replaced = true; break;
                        case DialogResult.No:
                            replaced = false; break;
                    }
                }
                else {

                    File.Delete(filepath);
                    replaced = true; 


                }
            }
            return replaced;
        }
        /// <summary>
        /// renomme le fichier en nomfichier_date_hh_mm_ss
        /// </summary>
        /// <param name="filepath">chemin initial du fichier</param>
        public static void Rename_Csv(string filepath)
            {
                File.Move(filepath, filepath + string.Format(".{0:d_M_yyyy_HH_mm_ss}", DateTime.Now));
            }
            /// <summary>
            /// renomme le fichier en nomfichier_date_hh_mm_ss
            /// </summary>
            /// <param name="filepath">chemin initial du fichier</param>
            /// <param name="timeTag">tag ajouter pour le nouveau nom</param>
            public static void Rename_Csv(string filepath, string timeTag)
            {
                File.Move(filepath, filepath + "." + timeTag);
            }

        }
    /// <summary>
    /// ecriture des log wondows
    /// </summary>
    /// 
    #region WindowsTools
    public class AdminitraionTools
    {

      


            
            public bool IsUserAdministrator()
            {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show(ex.Message);
                isAdmin = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                isAdmin = false;
            }
            return isAdmin;
        


    }



    }


    /// <summary>
    /// class d'ecriture des w log
    /// </summary>
    public class WindowsLog
    {
        public string  wLogSource { get; set; } = "Alma_ImportTools_log_Source";//somme des surfaces chutes
        public string wLog { get; set; } = "ImportTools_AlmaCam_Log";

        //string wEvent= "Sample Event";

        public void LogStart()
        {
            if (!EventLog.SourceExists(wLogSource)) {
                EventLog.CreateEventSource(wLogSource, wLog); }


        }

        public  void WriteLogEvent(string EventToLog)
        {
            EventLog.WriteEntry(wLog, EventToLog);

        }

        public  void WriteLogWarningEvent(string EventToLog)
        {
           EventLog.WriteEntry(wLog, EventToLog, EventLogEntryType.Warning);

        }

        public void WriteLogErrorEvent(string EventToLog)
        {
            EventLog.WriteEntry(wLog, EventToLog, EventLogEntryType.Error);

        }

        public void WriteLogSuccess(string EventToLog)
        {
            EventLog.WriteEntry(wLog, EventToLog, EventLogEntryType.SuccessAudit);

        }

    }
        /// <summary>
        /// Create a New INI file to store or load data
        /// </summary>
        public class IniFile
        {
            public string path;

            [DllImport("kernel32")]
            private static extern long WritePrivateProfileString(string section,
                string key, string val, string filePath);
            [DllImport("kernel32")]
            private static extern int GetPrivateProfileString(string section,
                     string key, string def, StringBuilder retVal,
                int size, string filePath);

            /// <summary>
            /// INIFile Constructor.
            /// </summary>
            /// <PARAM name="INIPath"></PARAM>
            public IniFile(string INIPath)
            {
                path = INIPath;
            }
            /// <summary>
            /// Write Data to the INI File
            /// </summary>
            /// <PARAM name="Section"></PARAM>
            /// Section name
            /// <PARAM name="Key"></PARAM>
            /// Key Name
            /// <PARAM name="Value"></PARAM>
            /// Value Name
            public void IniWriteValue(string Section, string Key, string Value)
            {
                WritePrivateProfileString(Section, Key, Value, this.path);
            }

            /// <summary>
            /// Read Data Value From the Ini File
            /// </summary>
            /// <PARAM name="Section"></PARAM>
            /// <PARAM name="Key"></PARAM>
            /// <PARAM name="Path"></PARAM>
            /// <returns></returns>
            public string IniReadValue(string Section, string Key)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp,
                                                255, this.path);
                return temp.ToString();

            }
       

    }
    #endregion

    #region registry
    public static class Alma_RegitryInfos
        {
            public static string LastModelDatabaseName { get; set; } //nom de la machine par defaut
            private static string LastModelDatabaseNamekey=@"LastModelDatabaseName";
            private static string wpmkey = @"Software\Alma\Wpm";

            public static string GetLastDataBase()
            { string lastdatabase=null;
                        try
                       {
           
                            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(wpmkey);
                            if (key != null)
                            {
                                Object o = key.GetValue(LastModelDatabaseNamekey);
                                if (o != null) { LastModelDatabaseName = o.ToString(); }
                                 lastdatabase=LastModelDatabaseName;
                            }
                            if (lastdatabase == null) { MessageBox.Show("Database Not Found in the current user registry key  :" + LastModelDatabaseNamekey + "\\" + wpmkey); }
                            return lastdatabase;
                       }

                        catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                        {
                             //react appropriately
                            MessageBox.Show(ex.Message);
                            return lastdatabase;
                        }

            }

        /// <summary>
        /// retourne la valeur sd'une clé
        /// </summary>
        /// <param name="key_path"></param>
        /// <param name="key_name"></param>
        /// <returns></returns>
         public static string GetRegistryInfosKey(string key_path , string key_name)

        {
            try
            {
                string rst=null;
                
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key_path);
                if (key != null)
                {
                    Object o = key.GetValue(key_name);
                    if (o != null) { rst = o.ToString(); }
                    rst= "";
                }

                if (rst == null) { MessageBox.Show("Database Not Found in the current user registry key  :" + key_path + "\\" + key_name); }
                return rst;
            }

            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                //react appropriately
                MessageBox.Show(ex.Message);
                return "";
            }

        }

        }


        public enum Log_Type {
        verbose,
        discret
        }

        

        

    /// <summary>
    /// Ecriture des logs: il existe 2 log les log importants que l'utilisateur doit voir et les logs de debuggage
    /// pour visualiser les logs de debuggages, il suffit d'activer la case a cocher log verbeux
    /// </summary>
    public static class Alma_Log
        {     
              private static string temporyFolder = Path.GetTempPath();
              private static string logfile;
              private static Log_Type LogType;
                        
            public static bool Create_Log()
            {

            string timestamp=string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now) + "_AlmaCam.log";
            logfile = timestamp;
           
            Trace.Listeners.Add(new TextWriterTraceListener(temporyFolder + @"\" + logfile)); //Création d'un "listener" texte pour sortie dans un fichier texte
            Trace.AutoFlush = true; //On écrit directement si true, pas de temporisation.		
            Trace.WriteLine("###### Alma LogStart: " + Convert.ToString (DateTime.Now));
            Trace.WriteLine("###### Alma LogStart: " + temporyFolder + @"\" + logfile);
                return true;
            }

        public static bool Create_Log(bool verbosemode)
        {
             
            string timestamp = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now) + "_AlmaCam.log";
            logfile = timestamp;
            //mode verbeux ou non
            LogType = Log_Type.discret;
            if (verbosemode) { LogType = Log_Type.verbose; }
                      

            Trace.Listeners.Add(new TextWriterTraceListener(temporyFolder + @"\" + logfile)); //Création d'un "listener" texte pour sortie dans un fichier texte
            Trace.AutoFlush = true; //On écrit directement si true, pas de temporisation.		
            Trace.WriteLine("###### Alma LogStart: " + Convert.ToString(DateTime.Now));
            Trace.WriteLine("###### Alma LogStart: " + temporyFolder + @"\" + logfile);
            return true;
        }

            public static void Write_Log(string message)
            {
            
                Trace.WriteLineIf(LogType==Log_Type.verbose, message);
                //return true;
             }


            public static void Error(string message, string module)
            {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
        }

            public static void Error(Exception ex, string module)
            {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(ex.Message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
        }

            public static void Warning(string message, string module)
            {
            Trace.WriteLineIf(LogType == Log_Type.verbose, string.Format("{0}:{1}", message, module));
            }

            public static void Info(string message, string module)
            {
            Trace.WriteLineIf(LogType == Log_Type.verbose, string.Format("{0}:{1}", message, module));
             }
            //ecriture dans le log non verbeu
            public static void Write_Log_Important(string message)
            {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
                //return true;
            }

            public static void Close_Log()
            {
                //Debug.Flush();
                Trace.Close();
                //return true;
            }

            public static void Final_Open_Log()

            {
                Trace.Close();
                Process.Start("notepad.exe", temporyFolder + "\\" + logfile);                
                //return true;
            }

            public static void Final_Open_Log(long linelimit)
            {

            Trace.Close();
                if (linelimit<5000){
                Process.Start("notepad.exe", temporyFolder + "\\" + logfile);}
                else{System.Windows.Forms.MessageBox.Show("le fichier de log est disponible dans "+ temporyFolder +"\\"+ logfile);}
                //return true;
            }
        }

    public static class new_Alma_Log
    {
        private static string temporyFolder = Path.GetTempPath();
        private static string logfile;
        private static string Verbose_logfile;
        private static Log_Type LogType;
        /// <summary>
        /// creation des log verbeux et des log non verbeux
        /// </summary>
        /// 
        /// <returns>true false si log cree</returns>
        public static bool Create_Log()
        {

            string timestamp = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now) + "_AlmaCam.log";
            logfile = timestamp;
            Verbose_logfile = "verbose_"+timestamp;
            Trace.Listeners.Add(new TextWriterTraceListener(temporyFolder + @"\" + logfile));
            //Trace.Listeners.Add(new TextWriterTraceListener(temporyFolder + @"\" + Verbose_logfile)); //Création d'un "listener" texte pour sortie dans un fichier texte//Création d'un "listener" texte pour sortie dans un fichier texte
            Trace.AutoFlush = true; //On écrit directement si true, pas de temporisation.		
            Trace.WriteLine("###### Alma LogStart: " + Convert.ToString(DateTime.Now));
            Trace.WriteLine("###### Alma LogStart: " + temporyFolder + @"\" + logfile);
            return true;
        }

        public static bool Create_Log(bool verbosemode)
        {

            string timestamp = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now) + "_AlmaCam.log";
            logfile = timestamp;
            //mode verbeux ou non
            LogType = Log_Type.discret;
            if (verbosemode) { LogType = Log_Type.verbose; }


            Trace.Listeners.Add(new TextWriterTraceListener(temporyFolder + @"\" + logfile)); //Création d'un "listener" texte pour sortie dans un fichier texte
            Trace.AutoFlush = true; //On écrit directement si true, pas de temporisation.		
            Trace.WriteLine("###### Alma LogStart: " + Convert.ToString(DateTime.Now));
            Trace.WriteLine("###### Alma LogStart: " + temporyFolder + @"\" + logfile);
            return true;
        }

        public static void Write_Log(string message)
        {

            Trace.WriteLineIf(LogType == Log_Type.verbose, message);
            //return true;
        }


        public static void Error(string message, string module)
        {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
        }

        public static void Error(Exception ex, string module)
        {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(ex.Message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
        }

        public static void Warning(string message, string module)
        {
            Trace.WriteLineIf(LogType == Log_Type.verbose, string.Format("{0}:{1}", message, module));
        }

        public static void Info(string message, string module)
        {
            Trace.WriteLineIf(LogType == Log_Type.verbose, string.Format("{0}:{1}", message, module));
        }

        public static void Write_Log_Important(string message)
        {
            Trace.Indent();
            Trace.WriteLine("********************************");
            Trace.WriteLine(message);
            Trace.WriteLine("********************************");
            Trace.Unindent();
            //return true;
        }

        public static void Close_Log()
        {
            //Debug.Flush();
            Trace.Close();
            //return true;
        }

        public static void Final_Open_Log()
        {
            Trace.Close();
            Process.Start("notepad.exe", temporyFolder + "\\" + logfile);
            //return true;
        }

        public static void Final_Open_Log(long linelimit)
        {

            Trace.Close();
            if (linelimit < 5000)
            {
                Process.Start("notepad.exe", temporyFolder + "\\" + logfile);
            }
            else { System.Windows.Forms.MessageBox.Show("le fichier de log est disponible dans " + temporyFolder + "\\" + logfile); }
            //return true;
        }
    }

    public static class Lock_File
    {
        public static void CreateLockFile(string path)
        {
            File.CreateText(path);
        }

        public static void DeleteLockFile(string path)
        {
            File.Delete(path);
        }


    }
    
    public static class Alma_Time {

        public static double minutes(double val) {return val/60; }
        public static double seconds(double val) { return val; }
        public static double decimales_hours(double val) { return val / 3600 ; }

    }

    public static class ExtendedEntity_Tools
    {
        public static string EntityType;
        //public static IContext contexlocal;
        public static string ExtendedEntityPath;

        public static IContext contexlocal { get; set; }

        

        /// <summary>
        /// return a extended entity with the equal condition
        /// </summary>
        /// <param name="PathToField"> sous la forme "\_REFERENCE\_NAME"</param>
        /// <param name="value"> objet : id, string "P01" ou 12</param>
        /// <returns></returns>
        /// 

       
        public static IExtendedEntityList Extended_List_compute_Equal(string PathToField, object value)
        {

            try
            { 
            IEntityType e = contexlocal.Kernel.GetEntityType(PathToField.Split('\\')[0]);

            ExtendedEntityType extendedentitytype = (ExtendedEntityType)e.ExtendedEntityType;

            IConditionType conditionrtype1 = contexlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(extendedentitytype.GetExtendedField(PathToField),
                ConditionOperator.Equal,
                contexlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("C1", value));
            IQueryType querytype = new QueryType(contexlocal.Kernel, "MyQuery", e);
            querytype.SetFilter(conditionrtype1);
            IExtendedEntityList l = contexlocal.EntityManager.GetExtendedEntityList(e);
            l.Fill(false);


            return l;
            }
            catch (Exception ie)
            {
                Alma_Log.Error(ie.Message, "compute Extended entityList");
                return null;
            }

        }

/// <summary>
/// retourne une liste d'entité dont le champs de l'entité est egale a value
/// </summary>
/// <param name="PathToField">exemple \_REFERENCE\_NAME</param>
/// <param name="value">exemple "toto"</param>
/// <returns></returns>
        public static IEntityList EntityList_compute_Equal(string PathToField, object value)
        {

            try
            {
                IEntityList el=null;
                el = contexlocal.EntityManager.GetEntityList(PathToField.Split('\\')[0], PathToField.Split('\\')[1], ConditionOperator.Equal, value);
                el.Fill(false);
                return el;
            }
            catch (Exception ie)
            {
                Alma_Log.Error(ie.Message, "compute entityList");
                return null;
            }

        }



    }
    #endregion

    #region StockManager
    /// <summary>
    /// contient les methoses statics pour ecrire dans le stock alma
    /// fonctionne avec les neting to cut
    /// 
    /// 
    /// </summary>

    public static class StockManager
            {
        /// <summary>
        /// creer le stock en fonction de la multiplicité
        /// </summary>
        /// <param name="nesting">nesting to cut</param>
        /// <returns></returns>
                public static bool CreateStockFromNestingToCut(IEntity nesting)
                {
                    try {

                bool EXPLODE_MULTIPLICITY = true;

                string cle = nesting.GetFieldValueAsString("_REFERENCE");

                // création stock CFAO

                        string nestingName = nesting.GetFieldValueAsString("_NAME");
                       
                        int nestingMultiplicity = nesting.GetFieldValueAsInt("_QUANTITY");

                        int a = 0;
                        IEntityList _sheetList = nesting.Context.EntityManager.GetEntityList("_SHEET");
                        _sheetList.Fill(false);
                        var sheetList = _sheetList.Where(p => p.GetFieldValueAsString("_NAME").Contains(nestingName));
                ///verification des toles dans le stock
                // si les tole sont equales a la multiplicité on ne fait rien
                //int c = nesting.Context.EntityManager.GetEntity("_STOCK");
                if (EXPLODE_MULTIPLICITY==true) { 
                        foreach (IEntity sheet in sheetList)
                        {

                            for (int i = 0; i < nestingMultiplicity; i++)
                            {
                                IEntity newStock = nesting.Context.EntityManager.CreateEntity("_STOCK");
                            //    string previewFullPath = GetPreview(IEntity EntityToDraw, emfoutputdir);
                               
                                newStock.SetFieldValue("_SHEET", sheet.Id32);
                                newStock.SetFieldValue("_NAME", i.ToString());
                                newStock.SetFieldValue("_QUANTITY", 1);
                                newStock.SetFieldValue("AF_STOCK_CFAO", true);
                                newStock.Save();
                    
                        string previewFullPath = SimplifiedMethods.CreateAndCopyPreview(newStock.GetFieldValueAsEntity("_SHEET"), "_SHEET_"+ sheet.Id.ToString()+"_STOCK_"+newStock.Id.ToString()+".emf");
                        newStock.SetFieldValue("AF_STOCK_NAME", nestingName + "_"+sheet.Id + "_" + newStock.Id.ToString());
                        newStock.SetFieldValue("FILENAME", previewFullPath);

                        newStock.Save();

                            }
                        }
                }

                //////pas pris en compte pour le moment////////
                else
                {
                    IEntity newStock = nesting.Context.EntityManager.CreateEntity("_STOCK");
                    
                    newStock.SetFieldValue("_SHEET", sheetList.FirstOrDefault().Id32);
                    newStock.SetFieldValue("_NAME", 1);
                    newStock.SetFieldValue("_QUANTITY", nestingMultiplicity);
                    newStock.SetFieldValue("AF_STOCK_CFAO", true);
                    newStock.Save();

                    string previewFullPath = SimplifiedMethods.CreateAndCopyPreview(newStock.GetFieldValueAsEntity("_SHEET"), "_SHEET_" + sheetList.FirstOrDefault().Id.ToString() + "_STOCK_" + newStock.Id.ToString() + ".emf");
                    newStock.SetFieldValue("AF_STOCK_NAME", nesting.GetFieldValueAsString("_NAME") + "_"+sheetList.FirstOrDefault().Id + "_" + newStock.Id.ToString());
                    newStock.SetFieldValue("FILENAME", previewFullPath);


                }





                return true;
                    }

                    catch (Exception ie) {

                            MessageBox.Show(ie.Message);
                            return false;
                    }

                    finally { }

                }
        /// <summary>
        /// les toles crees par cette lib sont marquées comme etant des toles cfoa 
        /// supprime les toles cfao ( generalement pour anulation des placements)
        /// </summary>
        /// <param name="nestingToCut">nesting to cut</param>
        /// <returns></returns>
                public static bool DeleteNestingAssociatedStock(IEntity nestingToCut)
                        {
                            try
                            {

                                IEntity nesting = nestingToCut;
                                string nestingname = nesting.GetFieldValueAsString("_NAME");
                                IEntityList _stockList = nesting.Context.EntityManager.GetEntityList("_STOCK");
                                _stockList.Fill(false);
                                 var stockList = _stockList.Where(p => p.GetFieldValueAsString("AF_STOCK_NAME") != null && p.GetFieldValueAsString("AF_STOCK_NAME").StartsWith(nestingname));
                                foreach (IEntity stock in stockList)
                                {
                                stock.Delete();
                                }
                                return true;
                            }


                            catch (Exception ie)
                            {

                                MessageBox.Show(ie.Message);
                                return false;
                            }

                            finally { }


                        }
        /// <summary>
        /// supprime les toles non cfao donc crees par almacam en automatique
        /// </summary>
        /// <param name="Closed_Nesting">close nesting</param>
        /// <returns></returns>
        public static bool DeleteAlmaCamStock(IEntity Closed_Nesting)
                    {
                        try
                        {

                         IEntity nesting = Closed_Nesting;
                            
                        string nestingname = nesting.GetFieldValueAsString("_NAME");
                        //IEntityList _sheetList = nesting.Context.EntityManager.GetEntityList("_SHEET");
                        IEntityList _sheetList = nesting.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting.Id);
                        _sheetList.Fill(false);

                                if (_sheetList.Count() > 0)
                                {
                                    foreach (IEntity sheet in _sheetList)
                                    {


                                                IEntityList _stockList = nesting.Context.EntityManager.GetEntityList("_STOCK", LogicOperator.And, "_SHEET", ConditionOperator.Equal, sheet.Id, "AF_STOCK_CFAO", ConditionOperator.Equal, false);
                                                _stockList.Fill(false);


                                                if (_stockList.Count() > 0)
                                                {
                                                    foreach (IEntity stock in _stockList)
                                                    {

                                                        stock.Delete();
                                                    }
                                                }


                                    }
                                }


                
                           
                            return true;
                        }


                        catch (Exception ie)
                        {

                            MessageBox.Show(ie.Message);
                            return false;
                        }

                        finally { }


                    }


    }
    #endregion

    #region Nestinfos
    /// <summary>
    /// nouvelle class derivable contenant des listes de placements avec des listes de chutes et leurs propriétés
    /// nestinfos2 contient une liste de Tole avec piece et chute correspondante
    /// Tole_Nesting : le information relatif a la tole utilisee pour le placement en cours d'etude
    /// Offcut_infos_List : contient la liste des chutes associées à la tole du placement
    /// Nested_Part_Infos_List : contient la liste des Pieces du placement
    /// </summary>

    public class Nest_Infos : IDisposable
    {   
        //propriete
        public Tole Tole_Nesting { get; set; }      //placement
        public IEntity Nesting { get; set; }        //placement
        public IEntity NestingSheetEntity { get; set; }     //format
        public IEntity NestingStockEntity { get; set; }     //stock utiliser pour le placement
        public Int64   NestingId;                             //id du nesting
        ///
        public IEntity Machine_Entity;
        public Int32 DefaultMachine_Id { get; set; }        //id de la machine par defaut
        public string Nesting_MachineName { get; set; }     //nom de la machine par defaut
        public string Nesting_CentreFrais_Machine { get; set; }  //clipper machine centre de frais
        public double LongueurCoupe { get; set; } // longeur de coupe *
        public Int64 NestingMultiplicity { get; set; } = 1;    //multiplicité placement
        public double Nesting_FrontWaste { get; set; } //chute au front
        public double Nesting_TotalWaste { get; set; } //chute totale
        public double Nesting_FrontWasteKg { get; set; } //chute au front en kg
        public double Nesting_TotalWasteKg { get; set; }//chute totale en kg       
        public double Nesting_TotalTime { get; set; } //in seconds       
        public double NestingSheet_loadingTimeInit { get; set; }  //temps de chanregement        
        public double NestingSheet_loadingTimeEnd { get; set; }//temps de chanregement fin
        public Boolean IS_ROTATED = false;

        // public SpecificFields Nest_Infos_2_specificFields;
        public SpecificFields Nest_Infos_2_Fields = new SpecificFields();


        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   

        //// liste du stock reservé
        ///retourne la liste des toles selectionnées
        public List<IEntity> Booked_Stock_Entity_List = new List<IEntity>();
        ///
        /// <summary>
        /// offcutlist
        /// </summary>
        public List<Tole> Offcut_infos_List { get; set; }
        /// <summary>
        /// partlist
        /// </summary>
        public List<Nested_PartInfo> Nested_Part_Infos_List = null;
        /// <summary>
        /// calculus GP
        /// </summary>
        public double Calculus_Parts_Total_Surface { get; set; }//somme des surfaces pieces 
        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   
        public double Calculus_Parts_Total_Weight { get; set; }//somme des surfaces pieces 
        [DefaultValue(0.0000001)] //eviter l' erreur de la division par 0   
        public double Calculus_Parts_Total_Time { get; set; } = 0;//somme des surfaces pieces 
        public double Calculus_Offcuts_Total_Surface { get; set; } = 0;//somme des surfaces chutes
        public double Calculus_Offcuts_Total_Weight { get; set; } = 0;//somme des surfaces chutes      
        public double Calculus_Offcut_Ratio { get; set; } = 0;//somme des surfaces chutes
        //calculus
        public double Calculus_Ratio_Consommation { get; set; }
        public double Calculus_CheckSum = 1;
        public Boolean Calculus_CheckSum_OK = false;


        ///dispose
        ///
        ///purge auto
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }


        #region parts

        /// <summary>
        /// calcul de la liste des pieces placée dans le placement
        /// atention, les pieces fantome ne sont pas prise en compte
        /// </summary>
        /// <param name="nestedpart"></param>
        public void Get_NestedPartInfos(IEntity nestedpart)
        {
            //piece par toles

            IEntity machinable_Part = null;
            IEntity to_produce_reference = null;

            Nested_PartInfo nested_Part_Infos = new Nested_PartInfo();
            //


            nested_Part_Infos.Part_To_Produce_IEntity = nestedpart.GetFieldValueAsEntity("_TO_PRODUCE_REFERENCE");
            //on set matiere et epaisseur a celle du nesting
            nested_Part_Infos.Material_Id = Tole_Nesting.Material_Id; //   Sheet_Material_Id;
            nested_Part_Infos.Material_Name = Tole_Nesting.MaterialName; //  Sheet_MaterialName;
            nested_Part_Infos.Thickness = Tole_Nesting.Thickness; //Sheet_Thickness;

            //recuperation des infos du part to produce

            nested_Part_Infos.Part_Time = nestedpart.GetFieldValueAsDouble("_TOTALTIME");
            nested_Part_Infos.Nested_Quantity = nestedpart.GetFieldValueAsLong("_QUANTITY");
            nested_Part_Infos.Nested_Quantity = nestedpart.GetFieldValueAsLong("_QUANTITY");
            //repercution des infos de machinable part
            machinable_Part = nestedpart.GetFieldValueAsEntity("_MACHINABLE_PART");
            nested_Part_Infos.Surface = machinable_Part.GetFieldValueAsDouble("_SURFACE");
            nested_Part_Infos.Part_Total_Nested_Weight = nested_Part_Infos.Surface * nested_Part_Infos.Nested_Quantity;
            nested_Part_Infos.SurfaceBrute = machinable_Part.GetFieldValueAsDouble("_SURFEXT");
            nested_Part_Infos.Weight = machinable_Part.GetFieldValueAsDouble("_WEIGHT");
            nested_Part_Infos.Part_Total_Nested_Weight = nested_Part_Infos.Weight * nested_Part_Infos.Nested_Quantity;
            //nested_Part_Infos.EmfFile = machinable_Part.GetImageFieldValueAsLinkFile("_PREVIEW");
            //nested_Part_Infos.EmfFile = SimplifiedMethods.GetPreview(@machinable_Part.GetImageFieldValueAsLinkFile("_PREVIEW"), machinable_Part);
            nested_Part_Infos.EmfFile = SimplifiedMethods.GetPreview(machinable_Part);
            nested_Part_Infos.Width = machinable_Part.GetFieldValueAsDouble("_DIMENS1");
            nested_Part_Infos.Height = machinable_Part.GetFieldValueAsDouble("_DIMENS2");

            //reference to produce
            to_produce_reference = nestedpart.GetFieldValueAsEntity("_TO_PRODUCE_REFERENCE");
            nested_Part_Infos.Part_Reference = to_produce_reference.GetFieldValueAsString("_NAME");
            nested_Part_Infos.Part_Name = to_produce_reference.GetFieldValueAsString("_NAME");
            //custom_Fields
            //Nested_Part_Info.Custom_Nested_Part_Infos

            //nested_Part_Infos.c
            //ajout des methodes specifiques
            Get_NestedPart_CustomInfos(to_produce_reference, nested_Part_Infos);

            //calcul de la surface total des pieces
            //on ne somme que les pieces qui ont un uid gpao (numero de gamme ou autre..)
            if (nested_Part_Infos.Part_IsGpao == true)
            {
                Calculus_Parts_Total_Surface += nested_Part_Infos.Surface * nested_Part_Infos.Nested_Quantity;
                Calculus_Parts_Total_Weight += nested_Part_Infos.Weight * nested_Part_Infos.Nested_Quantity;
                Calculus_Parts_Total_Time += nested_Part_Infos.Part_Time * nested_Part_Infos.Nested_Quantity;
                //ajout à la liste les pieces qui ne sont pas de sieces fantomes
                Nested_Part_Infos_List.Add(nested_Part_Infos);
            }

            //tole
            // IEntity offcut_IEntity;





        }
        #endregion

        #region offcut
        // public virtual void Get_OffcutInfos(IEntity NestingStockEntity)
        public virtual void Get_OffcutInfos(Nest_Infos CurrentNesting)
        {


            //recuperation des chute de meme parent stock
            //IEntityList parentstocklist;
            IEntityList sheets, stocks;
            Offcut_infos_List = new List<Tole>();
            /*
            parentstocklist = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_PARENT_STOCK", ConditionOperator.Equal, CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
            parentstocklist.Fill(false);*/
            //recuperation du sheet du placement 
            if (CurrentNesting.Tole_Nesting.StockEntity != null)
            {
                sheets = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);
                sheets.Fill(false);

                //construction de la liste des chutes
                foreach (IEntity sheet in sheets)
                {
                    stocks = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet.Id);///NestingStockEntity.Id);
                    stocks.Fill(false);


                    foreach (IEntity offcut in stocks)
                    {

                        Tole offcut_tole = new Tole();
                        offcut_tole.StockEntity = offcut;
                        //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                        offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                        offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                        offcut_tole.Thickness = Tole_Nesting.Thickness;  //

                        offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                        offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;

                        ///sheet
                        offcut_tole.SheetEntity = offcut.GetFieldValueAsEntity("_SHEET");
                        offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                        offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                        offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                        offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                        //pour la tole totalsurface = surface

                        offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                        offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                        offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                        //pour la tole totalweight= weigth

                        //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                        offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                        offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                        /////
                        if (offcut != null)
                        {
                            ////stock 
                            offcut_tole.StockEntity = offcut;
                            ///////on egalise la multiplicité avec celle de la tole mere (a verifier si fiable)
                            offcut_tole.Mutliplicity = CurrentNesting.Tole_Nesting.Mutliplicity;
                            offcut_tole.Stock_Name = offcut.GetFieldValueAsString("_NAME");
                            offcut_tole.Stock_Coulee = offcut.GetFieldValueAsString("_HEAT_NUMBER");
                            offcut_tole.Stock_qte_initiale = offcut.GetFieldValueAsInt("_QUANTITY");
                            offcut_tole.Stock_qte_reservee = offcut.GetFieldValueAsInt("_BOOKED_QUANTITY");
                            offcut_tole.Stock_qte_Utilisee = offcut.GetFieldValueAsInt("_USED_QUANTITY");

                            Tole_Nesting.no_Offcuts = false;
                            Tole_Nesting.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                            //////
                            Offcut_infos_List.Add(offcut_tole);
                        }
                        else { Tole_Nesting.no_Offcuts = true; }
                    }
                }

            }



        }

        public virtual void Get_OffcutInfos()
        {
            //Nest_Infos CurrentNesting = this;

            //recuperation des chute de meme parent stock
            //IEntityList parentstocklist;
            IEntityList sheets, stocks;
            Offcut_infos_List = new List<Tole>();
            /*
            parentstocklist = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_PARENT_STOCK", ConditionOperator.Equal, CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
            parentstocklist.Fill(false);*/
            //recuperation du sheet du placement 
            if (Tole_Nesting.StockEntity != null)
            {
                sheets = Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, NestingId);
                sheets.Fill(false);

                //construction de la liste des chutes
                foreach (IEntity sheet in sheets)
                {
                    stocks = Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet.Id);///NestingStockEntity.Id);
                    stocks.Fill(false);


                    foreach (IEntity offcut in stocks)
                    {

                        Tole offcut_tole = new Tole();
                        offcut_tole.StockEntity = offcut;
                        //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                        offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                        offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                        offcut_tole.Thickness = Tole_Nesting.Thickness;  //

                        offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                        offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;

                        ///sheet
                        offcut_tole.SheetEntity = offcut.GetFieldValueAsEntity("_SHEET");
                        offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                        offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                        offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                        offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                        //pour la tole totalsurface = surface

                        offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                        offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                        offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                        //pour la tole totalweight= weigth

                        //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                        offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                        offcut_tole.Sheet_Is_rotated = IS_ROTATED;
                        /////
                        if (offcut != null)
                        {
                            ////stock 
                            offcut_tole.StockEntity = offcut;
                            ///////on egalise la multiplicité avec celle de la tole mere (a verifier si fiable)
                            offcut_tole.Mutliplicity = Tole_Nesting.Mutliplicity;
                            offcut_tole.Stock_Name = offcut.GetFieldValueAsString("_NAME");
                            offcut_tole.Stock_Coulee = offcut.GetFieldValueAsString("_HEAT_NUMBER");
                            offcut_tole.Stock_qte_initiale = offcut.GetFieldValueAsInt("_QUANTITY");
                            offcut_tole.Stock_qte_reservee = offcut.GetFieldValueAsInt("_BOOKED_QUANTITY");
                            offcut_tole.Stock_qte_Utilisee = offcut.GetFieldValueAsInt("_USED_QUANTITY");

                            Tole_Nesting.no_Offcuts = false;
                            Tole_Nesting.Sheet_Is_rotated = IS_ROTATED;
                            //////
                            Offcut_infos_List.Add(offcut_tole);
                        }
                        else { Tole_Nesting.no_Offcuts = true; }
                    }
                }

            }



        }


        public virtual void Get_OffcutInfos(Nest_Infos CurrentNesting, WorkShopOptionType Workshop_Option)
        {
            //creation de la liste des futures toles de type chute
            Offcut_infos_List = new List<Tole>();

            switch (Workshop_Option)
            {

                //fermeture tole a tole
                case WorkShopOptionType.GlobalCloseOneClic:

                    //recuperation des infos commune des chutes chutes de meme parent stock
                    //IEntityList parentstocklist;
                    /*
                    parentstocklist = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_PARENT_STOCK", ConditionOperator.Equal, CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
                    parentstocklist.Fill(false);*/

                    IEntityList sheets, stocks;
                    sheets = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);///NestingStockEntity.Id);
                    sheets.Fill(false);
                    //construction de la liste des chutes
                    foreach (IEntity sheet in sheets)
                    {

                        stocks = CurrentNesting.Tole_Nesting.StockEntity.Context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sheet.Id);///NestingStockEntity.Id);
                        stocks.Fill(false);


                        foreach (IEntity offcut in stocks)
                        {

                            Tole offcut_tole = new Tole();
                            offcut_tole.StockEntity = offcut;
                            //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                            offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                            offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                            offcut_tole.Thickness = Tole_Nesting.Thickness;  //
                            offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                            offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;
                                                                            ///sheet
                            offcut_tole.SheetEntity = offcut.GetFieldValueAsEntity("_SHEET");
                            offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                            offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                            offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                            offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                            offcut_tole.Sheet_Total_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE") * CurrentNesting.NestingMultiplicity;

                            //pour la tole totalsurface = surface

                            offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                            offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                            offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                            offcut_tole.Sheet_Total_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT") * CurrentNesting.NestingMultiplicity;

                            //pour la tole totalweight= weigth

                            //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                            offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                            offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                            /////
                            if (offcut != null)
                            {
                                ////stock 
                                offcut_tole.StockEntity = offcut;
                                ///////on egalise la multiplicité avec celle de la tole mere (a verifier si fiable)
                                offcut_tole.Mutliplicity = CurrentNesting.Tole_Nesting.Mutliplicity;
                                offcut_tole.Stock_Name = offcut.GetFieldValueAsString("_NAME");
                                offcut_tole.Stock_Coulee = offcut.GetFieldValueAsString("_HEAT_NUMBER");
                                offcut_tole.Stock_qte_initiale = offcut.GetFieldValueAsInt("_QUANTITY");
                                offcut_tole.Stock_qte_reservee = offcut.GetFieldValueAsInt("_BOOKED_QUANTITY");
                                offcut_tole.Stock_qte_Utilisee = offcut.GetFieldValueAsInt("_USED_QUANTITY");

                                Tole_Nesting.no_Offcuts = false;
                                Tole_Nesting.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                                //////
                                Offcut_infos_List.Add(offcut_tole);
                            }
                            else { Tole_Nesting.no_Offcuts = true; }
                        }
                    }


                    break;

                case WorkShopOptionType.GlobalCloseSeparated:

                    IEntityList sheetList = CurrentNesting.Nesting.Context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, CurrentNesting.NestingId);//CurrentNesting.Tole_Nesting.StockEntity.Id);///NestingStockEntity.Id);
                    sheetList.Fill(false);
                    //construction de la liste des chutes
                    foreach (IEntity sheet in sheetList)
                    {
                        Tole offcut_tole = new Tole();
                        offcut_tole.StockEntity = null;

                        //ON VALIDE LES POINTS GENERIQUES  MEME MATIERE QUE LA TOLE DU PLACEMENT
                        offcut_tole.Material_Id = Tole_Nesting.Material_Id; //  Sheet_Material_Id;
                        offcut_tole.MaterialName = Tole_Nesting.MaterialName; // Sheet_MaterialName;
                        offcut_tole.Thickness = Tole_Nesting.Thickness;  //
                        offcut_tole.Grade = Tole_Nesting.Grade; //  Sheet_Material_Id;
                        offcut_tole.GradeName = Tole_Nesting.GradeName; // Sheet_MaterialName;
                        offcut_tole.Mutliplicity = Tole_Nesting.Mutliplicity;
                        //sheet
                        offcut_tole.SheetEntity = sheet; ///offcut.GetFieldValueAsEntity("_SHEET");
                        offcut_tole.Sheet_Id = offcut_tole.SheetEntity.Id;
                        offcut_tole.Sheet_Name = offcut_tole.SheetEntity.GetFieldValueAsString("_NAME");
                        offcut_tole.Sheet_Reference = offcut_tole.SheetEntity.GetFieldValueAsString("_REFERENCE");
                        offcut_tole.Sheet_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE");
                        offcut_tole.Sheet_Total_Surface = offcut_tole.SheetEntity.GetFieldValueAsDouble("_SURFACE") * CurrentNesting.NestingMultiplicity;
                        //pour la tole totalsurface = surface

                        offcut_tole.Sheet_Length = offcut_tole.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                        offcut_tole.Sheet_Width = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                        offcut_tole.Sheet_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                        offcut_tole.Sheet_Total_Weight = offcut_tole.SheetEntity.GetFieldValueAsDouble("_WEIGHT") * CurrentNesting.NestingMultiplicity;
                        //pour la tole totalweight= weigth

                        //offcut_tole.Sheet_EmfFile = offcut_tole.SheetEntity.GetImageFieldValueAsLinkFile("_PREVIEW");
                        offcut_tole.Sheet_EmfFile = SimplifiedMethods.GetPreview(offcut_tole.SheetEntity);
                        offcut_tole.Sheet_Is_rotated = CurrentNesting.IS_ROTATED;
                        ///// plus d'infos de toel car elle n'existe pas
                        offcut_tole.no_Offcuts = true;
                        ///jamais de stock dans ce mode 
                        ///
                        offcut_tole.no_Stock = true;
                        Offcut_infos_List.Add(offcut_tole);
                    }


                    break;
                default:
                    throw new Exception("l'Option atelier choisie n'est pas compatible avec la librairie generique d'export GP");
                    break;
            }

        }
        #endregion
        
        #region nestinfos
        /// <summary>
        /// recupere les nestinfos infos de la tole mere
        /// </summary>
        /// <param name="currentsheet">t</param>
        /// <param name="stage"></param>

        public virtual void Get_NestInfos(IEntity nesting_to_cut)
        {
            IEntity to_cut_sheet;
            //IEntity nesting; // string stage;
            //nesting = Nesting;
            //affectation des infos de tole du nesting
            Tole_Nesting = new Tole();
            NestingId = nesting_to_cut.Id;
            Nesting = nesting_to_cut;
            //this.IS_ROTATED = Nesting.GetFieldValueAsBoolean("_IS_ROTATED");
            string stage = nesting_to_cut.EntityType.Key;
            IContext contextelocal = nesting_to_cut.Context;
            Tole_Nesting.Mutliplicity = nesting_to_cut.GetFieldValueAsInt("_QUANTITY");
            Tole_Nesting.SheetEntity = nesting_to_cut.GetFieldValueAsEntity("_SHEET");





            /*

            IEntityList formats = contextelocal.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting_to_cut.Id);
            formats.Fill(false);
            */

            Tole_Nesting.To_Cut_Sheet_Name = Tole_Nesting.SheetEntity.GetFieldValueAsString("_NAME");

            ////////////////////////////////////////////
            //recuperation du format de la tole du placement
            if (Tole_Nesting.SheetEntity != null)
            {
                //IEntity sheet = Nesting.GetFieldValueAsEntity("_SHEET"); 
                //Tole_Nesting.SheetEntity = Nesting.GetFieldValueAsEntity("_SHEET");
                Tole_Nesting.Sheet_Id = Tole_Nesting.SheetEntity.Id32;
                Tole_Nesting.Sheet_Name = Tole_Nesting.SheetEntity.GetFieldValueAsString("_NAME");
                Tole_Nesting.Sheet_Weight = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_WEIGHT");
                Tole_Nesting.Sheet_Length = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_LENGTH");
                Tole_Nesting.Sheet_Width = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_WIDTH");
                Tole_Nesting.Sheet_Surface = Tole_Nesting.SheetEntity.GetFieldValueAsDouble("_SURFACE");

                //pour la tole support on a poids = total poids si multiplicité =1 ce qui est le cas dans les clotures toles à tole
                Tole_Nesting.Sheet_Total_Weight = Tole_Nesting.Sheet_Weight;
                //pour la tole support on a surface = total surface si multiplicité =1 ce qui est le cas dans les clotures toles à tole
                Tole_Nesting.Sheet_Total_Surface = Tole_Nesting.Sheet_Surface;

                Tole_Nesting.Sheet_Reference = Tole_Nesting.SheetEntity.GetFieldValueAsString("_REFERENCE");
                Tole_Nesting.no_Offcuts = true;
                Tole_Nesting.Sheet_Is_rotated = this.IS_ROTATED;


            }


            ///information programme cn
            ///
            IEntityList programCns;
            IEntity programCn;
            programCns = contextelocal.EntityManager.GetEntityList("_CN_FILE", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting_to_cut.Id);
            programCn = SimplifiedMethods.GetFirtOfList(programCns);

            if (programCn != null)
            {
                
                this.Tole_Nesting.To_Cut_Sheet_NoPgm = programCn.GetFieldValueAsLong("_NOPGM").ToString();
                this.Tole_Nesting.To_Cut_Sheet_Pgm_Name = programCn.GetFieldValueAsString("_NAME");
                this.Tole_Nesting.To_Cut_Sheet_Extract_FullName = programCn.GetFieldValueAsString("_EXTRACT_FULLNAME");
            }
            else
            {
                this.Tole_Nesting.To_Cut_Sheet_NoPgm = "0";
                this.Tole_Nesting.To_Cut_Sheet_Pgm_Name = Tole_Nesting.Sheet_Name;
                this.Tole_Nesting.To_Cut_Sheet_Extract_FullName = Tole_Nesting.Sheet_Name;
            }

            // information du placement
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///NESTING LAYOUT///
            /////////////////)///////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //creation de la preview et recuperation du chemin
            Tole_Nesting.Sheet_EmfFile = SimplifiedMethods.GetPreview(nesting_to_cut);
            this.Nesting_TotalTime = nesting_to_cut.GetFieldValueAsDouble("_TOTALTIME");


            LongueurCoupe = nesting_to_cut.GetFieldValueAsDouble("_CUT_LENGTH");
            Nesting_FrontWaste = nesting_to_cut.GetFieldValueAsDouble("_FRONT_WASTE");
            Nesting_FrontWaste = nesting_to_cut.GetFieldValueAsDouble("_TOTAL_WASTE");
            //multiplicite interdite en mode 3 : closing by sheet on force a 1         

            ///validation matiere
            IEntity material = nesting_to_cut.GetFieldValueAsEntity("_MATERIAL");
            Tole_Nesting.Material_Id = material.Id32;
            Tole_Nesting.MaterialName = material.GetFieldValueAsString("_NAME");
            Tole_Nesting.Thickness = material.GetFieldValueAsDouble("_THICKNESS");

            //recuperation des grades
            Int32 gradeid = material.GetFieldValueAsInt("_QUALITY");

            IEntityList grades = null;
            //IEntity grade = null;
            grades = contextelocal.EntityManager.GetEntityList("_QUALITY", "ID", ConditionOperator.Equal, gradeid);
            Tole_Nesting.Grade = SimplifiedMethods.GetFirtOfList(grades);
            Tole_Nesting.GradeName = Tole_Nesting.Grade.GetFieldValueAsString("_NAME");

            ///////////////////////////////////////////////////////////////////////////////////
            //machine -->
            IEntityList machineliste;
            IEntity machine;

            Machine_Entity = nesting_to_cut.GetFieldValueAsEntity("_CUT_MACHINE_TYPE");
            machineliste = contextelocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "ID", ConditionOperator.Equal, Machine_Entity.Id32);
            //machineliste.Fill(false);
            machine = SimplifiedMethods.GetFirtOfList(machineliste);

            //recuperation des certains parametre de la ressource
            ICutMachineResource parameterList = AF_ImportTools.SimplifiedMethods.GetRessourceParameter(machine);
            //POUR L INSTANT ON CHARGE LES PARAMETRES DE CHARGERMENT AU DECHARGEMENT

            NestingSheet_loadingTimeInit = parameterList.GetSimpleParameterValueAsDouble("PAR_TPSCHARG");
            NestingSheet_loadingTimeEnd = parameterList.GetSimpleParameterValueAsDouble("PAR_TPSDECHARG");


            Nesting_MachineName = machine.GetFieldValueAsString("_NAME");
            DefaultMachine_Id = machine.Id32;

            //recuperation du centre de frais
            IEntity centrefrais;
            centrefrais = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
            Nesting_CentreFrais_Machine = centrefrais.GetFieldValueAsString("_CODE");

            centrefrais = null;
            machine = null;
            machineliste = null;

            ////////////////////////////////////////////////////////

            ////recuperation des infos de stock
            /*information sur le stock de clipper*/
            IEntityList stocklist=null;
            IEntity stock=null;

            stocklist = contextelocal.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, Tole_Nesting.Sheet_Id);
            stocklist.Fill(false);
            stock = SimplifiedMethods.GetFirtOfList(stocklist);

            Tole_Nesting.StockEntity = stock;
            if (stocklist.Count > 0)
            {
                ////stock 
                Tole_Nesting.Stock_Name = stock.GetFieldValueAsString("_NAME");
                Tole_Nesting.Stock_Coulee = stock.GetFieldValueAsString("_HEAT_NUMBER");
                Tole_Nesting.Stock_qte_initiale = stock.GetFieldValueAsInt("_QUANTITY");
                Tole_Nesting.Stock_qte_reservee = stock.GetFieldValueAsInt("_BOOKED_QUANTITY");
                Tole_Nesting.Stock_qte_Utilisee = stock.GetFieldValueAsInt("_USED_QUANTITY");

            }
            
            stocklist = null;
            stock = null;
        }
        /// <summary>
        /// calcul les ratios...
        /// </summary>
        #region calculus
        public virtual void ComputeNestInfosCalculus()
        {
            //
            int accuracy = 5; //nombre de chiffre apres la virgule
            //calcul de la surface total des chutes
            Calculus_Offcuts_Total_Surface = 0;
            Calculus_Offcuts_Total_Weight = 0;

            if (Offcut_infos_List != null)
            {

                Calculus_Offcuts_Total_Surface = Offcut_infos_List.Sum(o => o.Sheet_Surface);
                Calculus_Offcuts_Total_Weight = Offcut_infos_List.Sum(o => o.Sheet_Weight);
            }



            Calculus_Parts_Total_Surface = Nested_Part_Infos_List.Sum(o => o.Surface * o.Nested_Quantity);

            //calculus
            if ((Tole_Nesting.Sheet_Total_Surface - Calculus_Offcuts_Total_Surface) != 0)
            {
                Calculus_Ratio_Consommation = (((Tole_Nesting.Sheet_Total_Surface - Calculus_Offcuts_Total_Surface) * Tole_Nesting.Mutliplicity) / Calculus_Parts_Total_Surface);
            }

            //eciture des poids corrigés
            foreach (Nested_PartInfo p in Nested_Part_Infos_List)
            {
                if (Calculus_Ratio_Consommation != 0)
                {
                    p.Ratio_Consommation = Calculus_Ratio_Consommation;
                    p.Part_Balanced_Weight = Math.Round(p.Weight * Calculus_Ratio_Consommation, accuracy);
                    p.Part_Balanced_Surface = Math.Round(p.Surface * Calculus_Ratio_Consommation, accuracy);
                    Calculus_CheckSum += p.Weight * Calculus_Ratio_Consommation * p.Nested_Quantity;
                    p.Part_Total_Nested_Weight = p.Part_Balanced_Weight * p.Nested_Quantity;
                    p.Part_Total_Nested_Weight_ratio = p.Part_Total_Nested_Weight / Calculus_Offcuts_Total_Weight;

                }
                else
                {
                    p.Ratio_Consommation = 1;
                    p.Part_Balanced_Weight = p.Weight;
                    p.Part_Balanced_Surface = p.Surface;
                    p.Part_Total_Nested_Weight = p.Weight * 1 * p.Nested_Quantity;
                    Calculus_CheckSum = 0;

                }
            }


            //checksum des poids
            Calculus_CheckSum = Calculus_CheckSum - (Tole_Nesting.Sheet_Weight - Calculus_Offcuts_Total_Weight);

            //if (Calculus_CheckSum - (Tole_Nesting.Sheet_Weight - Calculus_Offcuts_Total_Weight) < 1)
            if (Math.Round(Calculus_CheckSum, accuracy) == 1)
            {
                Calculus_CheckSum_OK = true;
            }



        }
        #endregion

        #region Reserved_Stock_Infos
        public virtual void Get_Booked_Stock_Entities(IEntity nesting_to_cut)
        {
            

                 try
            {
                IContext contextlocal = nesting_to_cut.Context;
                List<IEntity> nesting_to_cut_list = new List<IEntity>();
                nesting_to_cut_list.Add(nesting_to_cut);


                bool manageStock = ActcutModelOptions.IsManageStock(contextlocal);
                manageStock = true;

                if (manageStock == true)
                {
                    BookNestingSheetData bookSheetToNestingData = new BookNestingSheetData(contextlocal, nesting_to_cut_list, true);
                    foreach (BookSheetData bookSheetData in bookSheetToNestingData.BookSheetDataList)
                    {
                        long Selected_Quantity = 0;

                        foreach (StockData stockData in bookSheetData.SheetList)
                        {
                            if (stockData.Quantity > 0) { 

                                Booked_Stock_Entity_List.Add(stockData.StockDataItem.StockEntity);
                                Selected_Quantity += stockData.Quantity;

                            }

                        }

                        if(Selected_Quantity< bookSheetData.Quantity)
                        {
                            MessageBox.Show("Les quantités selectionnées sont inferieurezs aux quantités necessaires, tous les ficheirs de retour ne seront pas crées.");
                        }



                    }


                }





            }
            catch { }
            finally { }


        }
        #endregion

        /// <summary>
        /// NestingName : exporte les données dans le streamwriter
        /// </summary>
        /// <param name="context"></param>
        /// <param name="to_cut_sheet"> entite a exporte vouée a disparaitre</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        /// <param name="export_gpao_file"></param>
        public virtual void Export_NestInfosToFile(IContext context, IEntity to_cut_sheet, string stage, StreamWriter export_gpao_file)
        {
        }
        /// <summary>
        /// ecrit le fichier de retour de l'export du dosssier technique
        /// stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// </summary>
        /// <param name="context">context par reference</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        /// <param name="nestingname">nom du placement</param>
        /// <param name="exportfile">stream vers le fichier d'export</param>
        public virtual void Export_NestInfosToFile(ref IContext context, string stage, string nestingname, StreamWriter exportfile)
        {


        }

        //public virtual void Export_NestInfosToFile(IContext context, IEntity to_cut_sheet, StreamWriter export_gpao_file)   {
        public virtual void Export_NestInfosToFile(IContext context, IEntity nesting_entity, StreamWriter export_gpao_file)
        {


        }

        /// <summary>
        ///  ecriture du fichier de sortie
        /// </summary>
        /// <param name="nestinfos">variables de type nestinfos2 preconstuit sur le nestinfos2</param>
        /// <param name="export_gpao_file">chemin vers le fichier de sortie</param>
        public virtual void Export_NestInfosToFile(StreamWriter export_gpao_file)
        {


        }
        /// <summary>
        /// stage = //list des placement stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;
        /// 
        /// </summary>
        /// <param name="nesting_sheet">entité tole de placement</param>
        /// <param name="stage">stage =  _SEQUENCED_NESTING, _CLOSED_NESTING , _TO_CUT_NESTING;</param>
        // public void GetPartsInfos(IEntity nesting_sheet, string stage) //IEntity Nesting)
        public void GetPartsInfos(IEntity to_cut_sheet_entity) //IEntity Nesting)
        {

            //recuperation des infos de sheet
            //IEntity to_cut_sheet;

            Nested_Part_Infos_List = new List<Nested_PartInfo>();
            IEntity current_nesting;

            //IEntity stock_Sheet;

            //
            current_nesting = Nesting;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///PARTS///
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //nested Parts-lists--> nestedpartinfos
            //on recherche les parts pointant sur le nesting//
            IEntityList nestedparts = null;

            //si stock managé on regarde les pieces placées sur la tole, sinn, sur le nesting
            //Actcut.ActcutModel.ActcutModelOptions.IsManagePartSet(contextlocal)…..stock managé, on regarde ls tocut reference sinon les proprités du nesting
            if (Actcut.CommonModel.ActcutModelOptions.IsManageStock(current_nesting.Context))
            {
                nestedparts = current_nesting.Context.EntityManager.GetEntityList("_TO_CUT_REFERENCE", "_TO_CUT_SHEET", ConditionOperator.Equal, to_cut_sheet_entity.Id32);
                nestedparts.Fill(false);
            }
            else
            {
                nestedparts = current_nesting.Context.EntityManager.GetEntityList("_NESTED_REFERENCE", "_NESTING", ConditionOperator.Equal, current_nesting.Id32);
                nestedparts.Fill(false);
            }


            foreach (IEntity nestedpart in nestedparts)
            {///recuperation de la liste des pieces
                Get_NestedPartInfos(nestedpart);


            }


            //calculus
            //ComputeNestInfosCalculus();
        }
        #endregion
        #region virtual methodes

        //custom field infos
        public virtual void SetSpecific_Generic_NestInfos() { }
        public virtual void Get_NestedPart_CustomInfos(IEntity nestedpart, Nested_PartInfo nestedpartinfos) { }
        public virtual void Get_Offcut_CustomInfos(IEntity offcut, Offcut_Infos offcutinfos) { }
        public virtual void Set_Offcut_CustomInfos(IEntity offcut, Offcut_Infos offcutinfos) { }
        public virtual void Get_NestInfos_CustomInfos(Tole Tole_nesting) { }

        #endregion
        public virtual Boolean Fill(IEntity nesting_to_cut)
        {
            return true;

        } 


    }
    #endregion
    #endregion
}