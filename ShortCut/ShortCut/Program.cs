using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Diagnostics;

namespace ShortCut
{
    class Program
    {
        static void Main(string[] args)
        {



            string currentdir = Directory.GetCurrentDirectory();
            string inifile = currentdir + "\\ShortCut.ini";
            StreamWriter sw=null;

            if (File.Exists(inifile)==false) {

               
                sw= File.CreateText(inifile);
                sw.Write(AlmaCam_Update.Resource1.ShorCut);
                sw.Close();
                sw.Dispose();

               
            }

            IniFile ini = new IniFile(inifile);
            string PathForUpdate = ini.IniReadValue("BASE", "PathForUpdate");
            string Pathalmacam = ini.IniReadValue("BASE", "Pathalmacam");
            string LastUpdate = ini.IniReadValue("BASE", "LastUpdate ");
            string PathToEditor = ini.IniReadValue("BASE", "PathToEditor");

             try
            {
                Console.WriteLine("Updating AlmaCam Specific Dll");

                string[] filePaths = Directory.GetFiles(PathForUpdate);
                foreach (var filename in filePaths)
                {
                    string @file = filename.ToString();

                    {
                        string dlltoupdate = Path.Combine(@Pathalmacam, Path.GetFileName(@file));
                        if (Tools.Updatedll(@file, @dlltoupdate))
                            File.Copy(@file, @Pathalmacam +"\\"+ Path.GetFileName(file),true);
                    }
                }

                ini.IniWriteValue("BASE", "LastUpdate ", DateTime.Now.ToString());

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = PathToEditor;
                Process processTemp = new Process();
                processTemp.StartInfo = startInfo;
                processTemp.EnableRaisingEvents = true;
                processTemp.Start();
            }
            catch (Exception e)
            {
                Console.WriteLine (e.Message);
                Console.ReadKey();
                
            }
            finally
            {
              
            }

        }



       



    }

    public static class Tools
    {

        public static bool Updatedll(string new_dll_fullPath, string dll_to_update_fullPath)
        {
            bool rst = true;
         
            long dest = 0;
            long source = 0;
       
          
            string newPath = Path.GetFullPath(new_dll_fullPath);
            string str = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            if (File.Exists(dll_to_update_fullPath) )
            {
                var dtdest = File.GetLastWriteTime(@dll_to_update_fullPath);
                    dest = (Int64) dtdest.Ticks;

                var dtsource = File.GetLastWriteTime(@new_dll_fullPath);
                    source= (Int64) dtsource.Ticks;


                if (source <= dest || Path.GetFileName(new_dll_fullPath) == str)
                { rst = false; }
                else
                {//on ecrit le logPath.GetDirectoryName(new_dll_fullPath) + "\\" + "Update_historique.txt"


                    using (StreamWriter historic = new StreamWriter(Path.GetDirectoryName(new_dll_fullPath) + "\\" + "Update_historique.txt", true))
                    {
                        string machine = "[ " + System.Environment.MachineName + " ]";

                        historic.WriteLine(".." + machine + dll_to_update_fullPath + " date source : {" + dtsource + "} date dest : {" + dtdest + "}");
                        historic.Close();
                        historic.Dispose();

                    }




                }
            


           
              
               

            }

            return rst;

            }

    }

    // <summary>
    /// Create a New INI file to store or load data
    /// </summary>
    public class IniFile
    {
        public string path;

        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
        [System.Runtime.InteropServices.DllImport("kernel32")]
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
}
