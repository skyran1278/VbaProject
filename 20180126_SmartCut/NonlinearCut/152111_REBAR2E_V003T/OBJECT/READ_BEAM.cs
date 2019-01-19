using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace REBAR2ET
{
    class READ_BEAM
    {
        private string _filename;

        public READ_BEAM(DB dbi)
        {
            _filename = dbi.Get_filename_BEAM();
        }

        public void readinput(string file, DB dbi)
        {
            FileInfo f = new FileInfo(file);
            StreamReader sr = f.OpenText();
            UInt32 count = 0;
   
            while (sr.Peek() > 0)
            {
                string word_temp = sr.ReadLine();

                char[] delimiterChars = { ' ', '-', ',', '\t' };

                string[] words = word_temp.Split(delimiterChars);


                dbi.Insert_BEAM_REBAR_DATA(count, words);

                count++;
            }


        }
    }
}
