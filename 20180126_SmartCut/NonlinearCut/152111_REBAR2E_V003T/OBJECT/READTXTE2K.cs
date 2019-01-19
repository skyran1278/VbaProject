using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace REBAR2ET
{
    class READTXTE2K
    {
        private string _filename;

        public READTXTE2K(DB dbi)
        {
            _filename = dbi.Get_filename_ETABS();
        }

        public void readinput(string file, DB dbi)
        {
            FileInfo f = new FileInfo(file);
            StreamReader sr = f.OpenText();
            UInt32 count = 0;
            while (sr.Peek() > 0)
            {
                string word_temp = sr.ReadLine();
                // word_temp.Replace(" ", "");

                // char[] delimiterChars = { ' ', ',', '.', ':', '\t' };
                char[] delimiterChars = {' ', ',', '\t' };

                string[] words = word_temp.Split(delimiterChars);

   

                dbi.Insert_ETABS_TXT(count, words);

                count++;
            }

            Dictionary<UInt32, string[]> out_txt = dbi.Get_ETABS_TXT();

            UInt32 count_f3 = 0;

            foreach (KeyValuePair<UInt32, string[]> it1 in out_txt)
            {
                string[] frame_section = it1.Value;
               // List<string> result = frame_section.ToList();

                // 讀取斷面性質
                //  FRAMESECTION  "B40X75"  MATERIAL "C280"  SHAPE "Rectangular"  D 0.75  B 0.4
                if (frame_section.Count() >= 7)
                {
                    if (frame_section[2] == "FRAMESECTION" & frame_section[6] == "MATERIAL")
                    {
                        string[] words_temp;
                        words_temp = frame_section;

                        string temp = frame_section[4];
                        dbi.Insert_FRAME_SECTIONS_1(temp, words_temp);

                    }
                    //  FRAMESECTION  "C55X55"  JMOD 0.00001  I2MOD 0.7  I3MOD 0.7  MMOD 0.96  WMOD 0.96
                    if (frame_section[2] == "FRAMESECTION" & frame_section[6] != "MATERIAL")
                    {
                        string[] words_temp;

                        words_temp = frame_section;
                        string temp = frame_section[4];
                        dbi.Insert_FRAME_SECTIONS_2(temp, words_temp);
                    }
                }

                // 讀取桿件指定性質
                //  LINEASSIGN  "B1"  "2F"  SECTION "B40X80"  ANG  0  MAXSTASPC 0.5  CARDINALPT 8    MESH "POINTSANDLINES"  

                if (frame_section.Count() >= 5)
                {
                    if (frame_section[2] == "LINEASSIGN")
                    {
                        string[] words_temp;
                        words_temp = frame_section;

                        dbi.Insert_FRAME_LINEASSIGN(count_f3, words_temp);
                        count_f3++;
                    }
                }



            }

            /************************************************************
            /////////測試輸出
            string Out_file = "test.e2k";

            using (StreamWriter sw2 = new StreamWriter(Out_file, false))
            {

                Dictionary<UInt32, string[]> out_txt = dbi.Get_ETABS_TXT();

                foreach (KeyValuePair<UInt32, string[]> it1 in out_txt)
                {
                    string[] out_line = it1.Value;

                    // List<string> result = out_line.ToList();

                    if (out_line.Length >= 2) //當欄位數大於2, 
                    {
                        string NewText = string.Join(" ",out_line, 0, out_line.Length); 
                        sw2.WriteLine(NewText); //寫入並換行
                    }


                }
                sw2.Close();

            }

           ************************************************************/



        }


    }

}
