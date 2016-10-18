using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Bilfinger_CAM_NC_converter
{
    public partial class Form1 : Form
    {
        private string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private string filePath = string.Empty;
        private string filePathFull = string.Empty;
        private string filePathOnly = string.Empty;
        private string fileNameGlobal = string.Empty;
        private string projectNumber = string.Empty;

        public Form1()
        {
            this.InitializeComponent();
        }

        public string FilePath()
        {
            return this.filePath;
        }

        public string DesktopPath()
        {
            return this.desktopPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.label1.Text = this.textBox1.Text;

            this.ProjNum();

            bool directoryExists = Directory.Exists(@"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath);
            if (directoryExists)
            {
                Directory.Delete(@"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath, true);
            }

            Directory.CreateDirectory(@"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath);

            this.DoIt();
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XSR|*.xsr|CAM|*.cam|All|*.*";
            ofd.Title = "Select cam or xsr";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.label3.Text = ofd.FileName.ToString();
                this.filePath = ofd.FileName.ToString();
                this.filePathFull = ofd.FileName.ToString();

                int lastDotIndex = this.filePath.LastIndexOf('.');
                int lastSlashIndex = this.filePath.LastIndexOf(@"\");
                this.filePathOnly = this.filePath.Substring(0, lastSlashIndex + 1);
                this.fileNameGlobal = this.filePath.Substring(lastSlashIndex + 1, this.filePath.Length - lastSlashIndex - 1);
                this.filePath = this.filePath.Substring(lastSlashIndex + 1, lastDotIndex - lastSlashIndex - 1);
                this.label1.Text = this.filePathOnly + this.fileNameGlobal;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {
        }

        private int GetNthIndex(string s, char t, int n)
        {
            int count = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == t)
                {
                    count++;
                    if (count == n)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        private string RenameProfile(string profile)
        {
            string newProfile = profile;

            if (profile[0] == 'B' && profile[1] == 'L')
            {
                int firstStar = this.GetNthIndex(profile, '*', 1);
                if (firstStar == -1)
                {
                    newProfile = "BL " + profile.Substring(2, profile.Length - 2);
                }
                else
                {
                    newProfile = "BL " + profile.Substring(2, firstStar - 2);
                }
            }
            else
            {
                int firstNumber = profile.IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
                int profileLength = profile.Length;

                if (profile.Contains("HEAA"))
                {
                    newProfile = profile.Substring(0, 2) + profile.Substring(4, profileLength - 4) + profile.Substring(2, 2);
                }
                else if (profile.Contains("HEA") || profile.Contains("HEB") || profile.Contains("HEM"))
                {
                    newProfile = profile.Substring(0, 2) + profile.Substring(3, profileLength - 3) + profile.Substring(2, 1);
                }

                newProfile = newProfile.Replace("*", "X");
                newProfile = newProfile.Replace("  ", " ");

                if (newProfile.Substring(0, 2) == "RR" || newProfile.Substring(0, 2) == "RO" || newProfile.Substring(0, 3) == "RHS" || newProfile.Substring(0, 2) == "RO")
                {
                    newProfile = newProfile.Replace(".", ",");

                    string temp = newProfile.Substring(newProfile.Length - 3, 3);
                    if (temp[0] == 'X')
                    {
                        newProfile = newProfile + ",0";
                    }
                    else
                    {
                        if (temp[1] == 'X')
                        {
                            newProfile = newProfile + ",0";
                        }
                    }
                }
            }

            // ne e napravena proverkata za kutienite secheniq RR160*80 * 5->RR 160X80X5,0
            return newProfile;
        }

        private string RenameMainPos(string mainPos)
        {
            string newMainPos = string.Empty;
            if (mainPos[0] == 'H')
            {
                newMainPos = mainPos.Substring(1, mainPos.Length - 1);
            }
            else
            {
                newMainPos = mainPos;
            }

            return newMainPos;
        }

        private string RenameName(string name)
        {
            string newName = name;

            if (name.Contains("Level"))
            {
                newName = newName.Replace("Level", "Hoehe");
            }

            if (name == "ENDPLATE")
            {
                newName = "Blech";
            }

            if (name == "PLATE")
            {
                newName = "Blech";
            }

            if (name == "BEAM")
            {
                newName = "Traeger";
            }

            if (name == "COLUMN")
            {
                newName = "Stuetze";
            }

            if (name == "Plate")
            {
                newName = "Blech";
            }

            if (name == "Beam")
            {
                newName = "Traeger";
            }

            if (name == "Column")
            {
                newName = "Stuetze";
            }

            if (name == "plate")
            {
                newName = "Blech";
            }

            if (name == "beam")
            {
                newName = "Traeger";
            }

            if (name == "column")
            {
                newName = "Stuetze";
            }

            if (name == "Träger")
            {
                newName = "Traeger";
            }

            if (name == "Stütze")
            {
                newName = "Stuetze";
            }

            newName = newName.Replace("ü", "ue").Replace("ä", "ae").Replace("ö", "oe").Replace("ß", "ss");

            return newName;
        }

        private string GetID()
        {
            string ID = Environment.UserName;

            if (Environment.UserName == "Nikola")
            {
                ID = "PME/NMS";
            }

            if (Environment.UserName == "Goro")
            {
                ID = "PME/GPM";
            }

            if (Environment.UserName == "Aneliya")
            {
                ID = "PME/AGY";
            }

            if (Environment.UserName == "Boyan Abdo")
            {
                ID = "PME/BNA";
            }

            if (Environment.UserName == "Boris")
            {
                ID = "PME/BGP";
            }

            if (Environment.UserName == "Chavdar")
            {
                ID = "PME/CHS";
            }

            if (Environment.UserName == "Delyana")
            {
                ID = "PME/DSY";
            }

            if (Environment.UserName == "Doch")
            {
                ID = "PME/DGG";
            }

            if (Environment.UserName == "Eva")
            {
                ID = "PME/ENA";
            }

            if (Environment.UserName == "Galin")
            {
                ID = "PME/GKK";
            }

            if (Environment.UserName == "Georgi")
            {
                ID = "PME/GSK";
            }

            if (Environment.UserName == "Hristo")
            {
                ID = "PME/HBY";
            }

            if (Environment.UserName == "Ismeth")
            {
                ID = "PME/IMH";
            }

            if (Environment.UserName == "Ivan")
            {
                ID = "PME/ISP";
            }

            if (Environment.UserName == "Joro")
            {
                ID = "PME/GDG";
            }

            if (Environment.UserName == "Maria")
            {
                ID = "PME/MIO";
            }

            if (Environment.UserName == "Mariqna")
            {
                ID = "PME/MDK";
            }

            if (Environment.UserName == "Milko")
            {
                ID = "PME/MSM";
            }

            if (Environment.UserName == "nexo")
            {
                ID = "PME/PGG";
            }

            if (Environment.UserName == "Niki")
            {
                ID = "PME/NVN";
            }

            if (Environment.UserName == "Pepsi")
            {
                ID = "PME/PMY";
            }

            if (Environment.UserName == "Plamen")
            {
                ID = "PME/PDM";
            }

            if (Environment.UserName == "PM")
            {
                ID = "PME/GPM";
            }

            if (Environment.UserName == "Rado")
            {
                ID = "PME/RKG";
            }

            if (Environment.UserName == "Rumy")
            {
                ID = "PME/RAM";
            }

            if (Environment.UserName == "Silvia")
            {
                ID = "PME/SZP";
            }

            if (Environment.UserName == "Soff")
            {
                ID = "PME/SPA";
            }

            if (Environment.UserName == "Stefan")
            {
                ID = "PME/SVD";
            }

            if (Environment.UserName == "Svilen")
            {
                ID = "PME/SSS";
            }

            if (Environment.UserName == "Tedy")
            {
                ID = "PME/TRU";
            }

            if (Environment.UserName == "Tim")
            {
                ID = "PME/TIM";
            }

            if (Environment.UserName == "Tina")
            {
                ID = "PME/RVM";
            }

            if (Environment.UserName == "Vanko")
            {
                ID = "PME/IDD";
            }

            if (Environment.UserName == "Yavor")
            {
                ID = "PME/YMY";
            }

            return ID;
        }

        private string RenameDIN(string DIN)
        {
            if (DIN == "EN-14399-4")
            {
                DIN = "EN 14399-4 10.9 HV TZN";
            }

            if (DIN == "EN ISO 4017" || DIN == "933")
            {
                DIN = "DIN EN ISO 4017-4032-7089/2/FORM B 8.8 TZN";
            }

            if (DIN == "ISO 4017")
            {
                DIN = "ISO 4017-7090-4032";
            }

            if (DIN == "Mu 4032+Sch 7090")
            {
                DIN = "ISO 4032-7090";
            }

            if (DIN == "Mutter-EN-14399-4")
            {
                DIN = "EN 14399-4-6";
            }
            
            return DIN;
        }

        private string RenameBolt(string boltName, string DIN, string boltLength)
        {
            if (DIN == "6914" || DIN == "14399" || DIN == "EN-14399-4")
            {
                boltName = boltName.Substring(3, boltName.Length - 3);
                boltName = "6KT SCHR " + boltName + " MU2S";
            }

            if (DIN == "7990")
            {
                boltName = "6KT SCHR " + boltName + " MUS";
            }

            /// rev 5
            if (DIN == "933" || DIN == "ISO 4017" || DIN == "ISO 4017-7090-7093-7042") 
            {
                boltName = "6KT SCHR " + boltName + " MU2S";
            }

            if (DIN == "4762")
            {
                boltName = "ZYL.SCHRAUBE " + boltName + "X" + boltLength;
            }

            if (DIN == "Mutter-EN-14399-4")
            {
                boltName = "14399-4-MU" + boltName.Substring(11, 2);
            }

            return boltName;
        }

        private int ProjNum()
        {
            string path = this.filePathFull;
            string line;
            int counter = 0;
            StreamReader cam = new StreamReader(path, System.Text.Encoding.Default);
            int linesCount = File.ReadAllLines(path).Count();
            string[] camRow = new string[linesCount];
            while ((line = cam.ReadLine()) != null)
            {
                camRow[counter] = line;
                counter++;
            }

            cam.Close();
            this.projectNumber = camRow[2].Remove(0, 2).Replace(" ", string.Empty);

            return 0;
        }

        private int DoIt()
        {
            string path = this.filePathFull;
            string line;

            if (!File.Exists(path))
            {
                throw new FileNotFoundException();
            }

            // save the file into a string array
            int counter = 0;
            StreamReader cam = new StreamReader(path, System.Text.Encoding.Default);
            int linesCount = File.ReadAllLines(path).Count();
            string[] camRow = new string[linesCount];
            while ((line = cam.ReadLine()) != null)
            {
                camRow[counter] = line;
                counter++;
            }

            cam.Close();
            
            // file saved into camRow[]
            path = this.fileNameGlobal;

            // search for revision in the name of the file
            int indexRev = path.IndexOfAny(new char[] { '_', '-' });
            string rev = string.Empty;
            string fileName = string.Empty;
            if (indexRev != -1)
            {
                int indexRevEnd = path.IndexOf('.');
                rev = path.Substring(indexRev + 1, path.Length - indexRevEnd - indexRev);
                fileName = path.Substring(0, indexRev);
            }
            else
            {
                rev = "00";
                int indexNameEnd = path.IndexOf('.');
                fileName = path.Substring(0, indexNameEnd);
            }
            /// revision is saved in rev

            string ID = this.GetID();

            DateTime todayDate = DateTime.Now;
            string format = "ddMMyyyy";
            
            string drawingName = this.textBox1.Text;
            if (drawingName == string.Empty)
            {
                drawingName = "name";
            }

            string pos = string.Empty, 
                   mainPos = string.Empty, 
                   profile = string.Empty, 
                   length = string.Empty, 
                   width = string.Empty, 
                   numberMain = string.Empty, 
                   number = string.Empty, 
                   material = string.Empty, 
                   partName = string.Empty, 
                   rest = string.Empty;
            string boltName = string.Empty, 
                   boltLength = string.Empty, 
                   boltPos = string.Empty, 
                   DIN = string.Empty;
            path = path.Replace(".xsr", ".cam");

            /// MessageBox.Show(@"D:\Bilfinger\" + projectNumber + @"\" + filePath + @"\" + "rev_" + path);

            StreamWriter revCam = new StreamWriter(@"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath + @"\" + "rev_" + path);

            StringBuilder output = new StringBuilder();
            Dictionary<string, string> NRows = new Dictionary<string, string>();

            output.Append("0#").AppendLine();
            output.Append("1#01").AppendLine();
            output.Append(camRow[2]).AppendLine();

            output.Append("3######").AppendLine();
            output.Append("4#" + fileName + "#" + rev).AppendLine();
            output.Append("5#" + ID).AppendLine();
            output.Append("9###" + drawingName + "###" + ID + "#" + todayDate.ToString(format) + "####").AppendLine();

            for (int i = 7; i < linesCount; i++)
            {
                if (camRow[i].Substring(0, 1) == "H")
                {
                    profile = camRow[i].Substring(2, this.GetNthIndex(camRow[i], '#', 2) - 2);
                    profile = this.RenameProfile(profile);
                    length = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 2) + 1, this.GetNthIndex(camRow[i], '#', 3) - (this.GetNthIndex(camRow[i], '#', 2) + 1));
                    width = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 3) + 1, this.GetNthIndex(camRow[i], '#', 4) - (this.GetNthIndex(camRow[i], '#', 3) + 1));
                    mainPos = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 4) + 1, this.GetNthIndex(camRow[i], '#', 5) - (this.GetNthIndex(camRow[i], '#', 4) + 1));
                    mainPos = this.RenameMainPos(mainPos);
                    numberMain = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 5) + 1, this.GetNthIndex(camRow[i], '#', 6) - (this.GetNthIndex(camRow[i], '#', 5) + 1));
                    material = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 6) + 1, this.GetNthIndex(camRow[i], '#', 7) - (this.GetNthIndex(camRow[i], '#', 6) + 1));
                    partName = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 9) + 1, this.GetNthIndex(camRow[i], '#', 10) - (this.GetNthIndex(camRow[i], '#', 9) + 1));
                    partName = this.RenameName(partName);
                    rest = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 10) + 1, this.GetNthIndex(camRow[i], '#', 25) - (this.GetNthIndex(camRow[i], '#', 10)));

                    NRows.Add(mainPos, "N#" + mainPos + ".nc");

                    output.Append("H#" + profile + "#" + length + "#" + width + "#" + mainPos + "#" + numberMain + "#" + material + "###" + partName + "#" + rest).AppendLine();
                }

                if (camRow[i].Substring(0, 1) == "W")
                {
                    profile = camRow[i].Substring(2, this.GetNthIndex(camRow[i], '#', 2) - 2);
                    profile = this.RenameProfile(profile);
                    length = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 2) + 1, this.GetNthIndex(camRow[i], '#', 3) - (this.GetNthIndex(camRow[i], '#', 2) + 1));
                    width = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 3) + 1, this.GetNthIndex(camRow[i], '#', 4) - (this.GetNthIndex(camRow[i], '#', 3) + 1));
                    pos = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 4) + 1, this.GetNthIndex(camRow[i], '#', 5) - (this.GetNthIndex(camRow[i], '#', 4) + 1));
                    pos = this.RenameMainPos(pos);
                    number = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 5) + 1, this.GetNthIndex(camRow[i], '#', 6) - (this.GetNthIndex(camRow[i], '#', 5) + 1));
                    material = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 6) + 1, this.GetNthIndex(camRow[i], '#', 7) - (this.GetNthIndex(camRow[i], '#', 6) + 1));
                    partName = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 9) + 1, this.GetNthIndex(camRow[i], '#', 10) - (this.GetNthIndex(camRow[i], '#', 9) + 1));
                    partName = this.RenameName(partName);
                    rest = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 10) + 1, this.GetNthIndex(camRow[i], '#', 25) - this.GetNthIndex(camRow[i], '#', 10));

                    if (!NRows.ContainsKey(pos))
                    {
                        NRows.Add(pos, "N#" + pos + ".nc");
                    }

                    output.Append("W#" + profile + "#" + length + "#" + width + "#" + pos + "#" + number + "#" + material + "###" + partName + "#" + rest).AppendLine();
                }

                if (camRow[i].Substring(0, 1) == "A")
                {
                    pos = camRow[i].Substring(2, this.GetNthIndex(camRow[i], '#', 2) - 2);
                    number = length = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 2) + 1, this.GetNthIndex(camRow[i], '#', 3) - (this.GetNthIndex(camRow[i], '#', 2) + 1));
                    mainPos = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 3) + 1, this.GetNthIndex(camRow[i], '#', 4) - (this.GetNthIndex(camRow[i], '#', 3) + 1));
                    mainPos = this.RenameMainPos(mainPos);

                    output.Append("A#" + pos + "#" + number + "#" + mainPos + "#").AppendLine();
                }

                if (camRow[i].Substring(0, 1) == "S")
                {
                    boltName = camRow[i].Substring(2, this.GetNthIndex(camRow[i], '#', 2) - 2);
                    boltLength = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 2) + 1, this.GetNthIndex(camRow[i], '#', 3) - (this.GetNthIndex(camRow[i], '#', 2) + 1));
                    boltPos = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 4) + 1, this.GetNthIndex(camRow[i], '#', 5) - (this.GetNthIndex(camRow[i], '#', 4) + 1));
                    DIN = camRow[i].Substring(this.GetNthIndex(camRow[i], '#', 9) + 1, this.GetNthIndex(camRow[i], '#', 10) - (this.GetNthIndex(camRow[i], '#', 9) + 1));

                    boltName = this.RenameBolt(boltName, DIN, boltLength);
                    DIN = this.RenameDIN(DIN);

                    output.Append("S#" + boltName + "#" + boltLength + "##" + boltPos + "#BOLT_XLS_QTY####" + DIN + "################").AppendLine();
                }
            }

            List<string> nc = new List<string>();

            // output.Replace('#', '&');
            using (revCam)
            {
                revCam.Write(output);
                foreach (KeyValuePair<string, string> item in NRows.OrderBy(key => key.Value))
                {
                    // revCam.WriteLine(item.Value.ToString().Replace('#','&'));
                    nc.Add(item.Value.ToString().Replace("N#", string.Empty));
                }
            }

            // nc
            string ncName = null;
            foreach (var item in nc)
            {
                counter = 0;
                if (!File.Exists(this.filePathOnly + item))
                {
                    ncName = this.filePathOnly + "H" + item;
                }
                else
                {
                    ncName = this.filePathOnly + item;
                }

                StreamReader ncReader = new StreamReader(ncName);

                int ncLines = File.ReadAllLines(ncName).Count();
                string[] ncRow = new string[ncLines];
                while ((line = ncReader.ReadLine()) != null)
                {
                    ncRow[counter] = line;
                    counter++;
                }

                ncReader.Close();

                bool metBO = false;
                StreamWriter ncWriter = new StreamWriter(@"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath + @"\" + "rev_" + item);

                using (ncWriter)
                {
                    ncWriter.WriteLine(ncRow[0]);
                    ncWriter.WriteLine(ncRow[2]);
                    /// ncWriter.WriteLine(ncRow[3]);
                    ncWriter.WriteLine("  " + fileName);
                    string ncPos = ncRow[4].Replace("  ", string.Empty);

                    ncPos = this.RenameMainPos(ncPos);
                    ncWriter.WriteLine("  " + ncPos);
                    ncWriter.WriteLine("  " + ncPos);
                    ncWriter.WriteLine(ncRow[6]);
                    ncWriter.WriteLine(ncRow[7]);
                    string ncProfile = ncRow[8].Replace("  ", string.Empty);

                    ncProfile = this.RenameProfile(ncProfile);
                    ncWriter.WriteLine("  " + ncProfile);
                    for (int i = 9; i < ncLines; i++)
                    {
                        if (ncRow[i] == "BO")
                        {
                            metBO = true;
                        }

                        if (metBO)
                        {
                            ncWriter.WriteLine(ncRow[i].Replace("       5.00", "m      0.00"));
                        }
                        else
                        {
                            ncWriter.WriteLine(ncRow[i]);
                        }
                    }

                    metBO = false;
                }
            }

            MessageBox.Show("Success!\nDon't forget to fill the number of the bolts!!!" + "\nFiles are here: " +
                @"D:\Bilfinger\" + this.projectNumber + @"\" + this.filePath);

            return 0;
        }
    }
}
