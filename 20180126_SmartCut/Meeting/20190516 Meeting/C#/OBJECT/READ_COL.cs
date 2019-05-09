using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace REBAR2ET
{
    class READ_COL
    {
        private string _filename;



        public READ_COL(DB dbi)
        {
            _filename = dbi.Get_filename_COL();
        }

        public void readinput(string file, DB dbi)
        {
            FileInfo f = new FileInfo(file);
            StreamReader sr = f.OpenText();
            UInt32 count = 0;
            while (sr.Peek() > 0)
            {
                string word_temp = sr.ReadLine();
                word_temp.Replace(" ", "");

                char[] delimiterChars = { ' ','-', ',', '\t' };

                string[] words = word_temp.Split(delimiterChars);

                dbi.Insert_COL_REBAR_DATA(count, words);

                count++;
            }


        }



    }
}
