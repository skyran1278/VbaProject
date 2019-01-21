using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace REBAR2ET
{
    class DB
    {
        /// <summary>
        /// 私有
        /// </summary>


        private string _file_ETABS,_file_COL, _file_BEAM;

        private string  _COL_hinge_i;
        private string  _COL_hinge_j;
        private string  _BEAM_hinge_i;
        private string  _BEAM_hinge_j;

        private string _COL_COVER; // unit:m
        private string _BEAM_COVER;


        private Dictionary<UInt32, string[]> _ETABS_TXT_DATA = new Dictionary<UInt32, string[]>();  // ETABS 文字檔內容
        private Dictionary<string, string[]> _FRAME_SECTIONS_1 = new Dictionary<string, string[]>();  // 尺寸資料
        private Dictionary<string, string[]> _FRAME_SECTIONS_2 = new Dictionary<string, string[]>();  // 折減資料

        private Dictionary<UInt32, string[]> _FRAME_LINEASSIGN = new Dictionary<UInt32, string[]>();  // 桿件指定性質

        private Dictionary<UInt32, string[]> _COL_REBAR_DATA = new Dictionary<UInt32, string[]>();  // 柱鋼筋量
        private Dictionary<UInt32, string[]> _BEAM_REBAR_DATA = new Dictionary<UInt32, string[]>();  // 梁鋼筋量

        /// <summary>
        /// 公有
        /// </summary>

        public string Get_filename_ETABS() { return _file_ETABS; }
        public string Get_filename_COL() { return _file_ETABS; }
        public string Get_filename_BEAM() { return _file_ETABS; }

        public string Get_COL_hinge_i() { return _COL_hinge_i; }
        public string Get_COL_hinge_j() { return _COL_hinge_j; }
        public string Get_BEAM_hinge_i() { return _BEAM_hinge_i; }
        public string Get_BEAM_hinge_j() { return _BEAM_hinge_j; }

        public string Get_COL_COVER() { return _COL_COVER; }
        public string Get_BEMA_COVER() { return _BEAM_COVER; }


        public Dictionary<UInt32, string[]> Get_ETABS_TXT() { return _ETABS_TXT_DATA; }
        public Dictionary<string, string[]> Get_FRAME_SECTIONS_1() { return _FRAME_SECTIONS_1; }
        public Dictionary<string, string[]> Get_FRAME_SECTIONS_2() { return _FRAME_SECTIONS_2; }
        public Dictionary<UInt32, string[]> Get_FRAME_LINEASSIGN() { return _FRAME_LINEASSIGN; }

        public Dictionary<UInt32, string[]> Get_COL_REBAR_DATA() { return _COL_REBAR_DATA; }
        public Dictionary<UInt32, string[]> Get_BEAM_REBAR_DATA() { return _BEAM_REBAR_DATA; }

        public void Insert_ETABS_TXT(UInt32 b, string[] value)
        {
            _ETABS_TXT_DATA.Add(b, value);
        }


        public void Insert_FRAME_SECTIONS_1(string b, string[] value)
        {
            _FRAME_SECTIONS_1.Add(b, value);
        }

        public void Insert_FRAME_SECTIONS_2(string b, string[] value)
        {
            _FRAME_SECTIONS_2.Add(b, value);
        }

        public void Insert_FRAME_LINEASSIGN(UInt32 b, string[] value)
        {
            _FRAME_LINEASSIGN.Add(b, value);
        }

        public void Insert_COL_REBAR_DATA(UInt32 b, string[] value)
        {
            _COL_REBAR_DATA.Add(b, value);
        }

        public void Insert_BEAM_REBAR_DATA(UInt32 b, string[] value)
        {
            _BEAM_REBAR_DATA.Add(b, value);
        }


        /// <summary>
        /// 建構子
        /// </summary>

        public DB(string name1, string name2, string name3, string name4, string name5, string name6, string name7, string name8,string name9)
        {
            _file_ETABS = name1;
            _file_COL = name2;
            _file_BEAM = name3;

            _COL_hinge_i = name4;
            _COL_hinge_j = name5;
            _BEAM_hinge_i = name6;
            _BEAM_hinge_j = name7;

            _COL_COVER = name8;
            _BEAM_COVER = name9;


        }


    }
}
