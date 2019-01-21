using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace REBAR2ET
{
    class WRITEe2K
    {



        public WRITEe2K()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbi"></param>
        public void write_out_e2k(DB dbi)
        {

            Dictionary<UInt32, string[]> frame_assign = dbi.Get_FRAME_LINEASSIGN();
            Dictionary<string, string[]> frame_prop1 = dbi.Get_FRAME_SECTIONS_1();
            Dictionary<string, string[]> frame_prop2 = dbi.Get_FRAME_SECTIONS_2();

            Dictionary<UInt32, string[]> frame_prop_out1 = new Dictionary<UInt32, string[]>();// 輸出
            Dictionary<UInt32, string[]> frame_prop_out2 = new Dictionary<UInt32, string[]>();// 輸出

            Dictionary<UInt32, string[]> frame_col_hinge = new Dictionary<UInt32, string[]>();// 輸出
            Dictionary<UInt32, string[]> frame_beam_hinge = new Dictionary<UInt32, string[]>();// 輸出


            Dictionary<UInt32, string[]> frame_assign_Out = new Dictionary<uint, string[]>();// 輸出

            /////柱鋼筋
            Dictionary<UInt32, string[]> COL_REBAR = dbi.Get_COL_REBAR_DATA();

            Dictionary<UInt32, string[]> COL_REBAR_OUT = new Dictionary<UInt32, string[]>();// 輸出

          

            ////////////////////////////////////////////////////////////////////////////////////////////////
            Dictionary<UInt32, StringBuilder> frame_prop_out1_temp = new Dictionary<uint, StringBuilder>();
            Dictionary<UInt32, StringBuilder> frame_prop_out2_temp = new Dictionary<uint, StringBuilder>();

            UInt32 ii = 0;
            foreach (KeyValuePair<string, string[]> it1 in frame_prop1)
            {
                string[] temp_out1 = it1.Value;

                StringBuilder temp_in1 = new StringBuilder();

                string NewText = string.Join(" ", temp_out1, 0, temp_out1.Length);

                temp_in1.Append(NewText);
                frame_prop_out1_temp.Add(ii, temp_in1);
                ii++;              
            }

            UInt32 jj = 0;
            foreach (KeyValuePair<string, string[]> it1 in frame_prop2)
            {
                string[] temp_out2 = it1.Value;

                StringBuilder temp_in2 = new StringBuilder();

                string NewText = string.Join(" ", temp_out2, 0, temp_out2.Length);

                temp_in2.Append(NewText);
                frame_prop_out2_temp.Add(jj, temp_in2);
                jj++;
            }


            ///////////////////////////////////////////////////////////////////////////////////////////////////


            string COL_hinge_i_loc = dbi.Get_COL_hinge_i();
            string COL_hinge_j_loc = dbi.Get_COL_hinge_j();
            string BEAM_hinge_i_loc = dbi.Get_BEAM_hinge_i();
            string BEAM_hinge_j_loc = dbi.Get_BEAM_hinge_j();

            string COL_COVER = dbi.Get_COL_COVER();
            string BEAM_COVER = dbi.Get_BEMA_COVER();

            UInt32 count_prop1 = 0;
            UInt32 count_prop2 = 0;

            UInt32 count_COL = 0;
            UInt32 count_hinge_col = 0;
            UInt32 count_hinge_beam = 0;

            UInt32 cout_assign = 0;

            foreach (KeyValuePair<UInt32, string[]> it1 in COL_REBAR)
            {

                string[] COL_REABR_temp = it1.Value;
                //        Story 柱編號 X向柱寬 Y向柱寬 主筋  X向支數 Y向支數     圍束區 非圍束區 X向繫筋    Y向繫筋
                //          PR	 C1	     50	      70	10-#8 	3	4	         #4@10	#4@15	  2      	1

                string col_numbe_out = "\"" + COL_REABR_temp[1] + "\"";
                string floor_number_out = "\"" + COL_REABR_temp[0] + "\"";


                string[] COL_REBAR_OUT_temp = new string[10];

                if (COL_REABR_temp.Length > 6)
                {

                    
                    COL_REBAR_OUT_temp[0] = "  CONCRETESECTION ";
                    COL_REBAR_OUT_temp[1] = "\"" + COL_REABR_temp[1] + "-" + COL_REABR_temp[0] + "\"";
                    COL_REBAR_OUT_temp[2] = " TYPE \"COLUMN\"  PATTERN ";
                    COL_REBAR_OUT_temp[3] = "\"R" + "-" + COL_REABR_temp[7] + "-" + COL_REABR_temp[8]+ "\"";
                    COL_REBAR_OUT_temp[4] = " TRANSREINF \"TIES\"  COVER ";

                    COL_REBAR_OUT_temp[5] = COL_COVER;

                    COL_REBAR_OUT_temp[6] = " REBAR ";
                    COL_REBAR_OUT_temp[7] = "\"" + COL_REABR_temp[5] + "\"";
                    COL_REBAR_OUT_temp[9] = " DESIGNCHECK \"CHECK\"";
                    count_COL++;
                    COL_REBAR_OUT.Add(count_COL, COL_REBAR_OUT_temp);// 輸出


                    foreach (KeyValuePair<UInt32, string[]> it2 in frame_assign)
                    {

                        string[] assign_temp = new string[30];

                        assign_temp = it2.Value;

                        string col_number = assign_temp[4];
                        string floor_number = assign_temp[6];

                        string[] temp1 = new string[30]; 
                        string[] temp2= new string[30];

                        if (col_number == col_numbe_out & floor_number == floor_number_out)
                        {

                            string col_section = assign_temp[9];

                            frame_prop1.TryGetValue(col_section, out temp1);
                            frame_prop2.TryGetValue(col_section, out temp2);


                            col_number = col_number.Replace("\"", "");
                            floor_number = floor_number.Replace("\"", "");

                            temp1[4] = "\"" + col_number + "-" + floor_number + "\"";


                            string NewText= string.Join(" ", temp1, 0, temp1.Length);

                            char[] delimiterChars = {',', '\t' };

                            temp1= NewText.Split(delimiterChars);


                            frame_prop_out1.Add(count_prop1, temp1);
                            count_prop1++;

                            if (temp2.Length > 0)
                            {
                                ///
                                temp2[4] = "\"" + col_number + "-" + floor_number + "\"";

                                string NewText2 = string.Join(" ", temp2, 0, temp2.Length);

                                temp2 = NewText2.Split(delimiterChars);

                                frame_prop_out2.Add(count_prop2, temp2);

                                count_prop2++;
                            }
                            /// 輸出COL HINGE
                            /// 


                            string[] temp3=new string[5];
                            string[] temp4 = new string[5];
                            // 底
                            temp3[0] = "  LINEASSIGN ";
                            temp3[1] = "\"" + col_number + "\""  +" ";
                            temp3[2] = "\"" + floor_number + "\"" + "  ";
                            temp3[3] = "HINGE \"Default-PMM\"  RDISTANCE ";
                            temp3[4] = COL_hinge_i_loc;
                            frame_col_hinge.Add(count_hinge_col, temp3);
                            ++count_hinge_col;

                            // 頂
                            temp4[0] = "  LINEASSIGN ";
                            temp4[1] = "\"" + col_number + "\"" + " ";
                            temp4[2] = "\"" + floor_number + "\"" + "  ";
                            temp4[3] = "HINGE \"Default-PMM\"  RDISTANCE ";
                            temp4[4] = COL_hinge_j_loc;

                            frame_col_hinge.Add(count_hinge_col, temp4);
                            ++count_hinge_col;

                            ///// 轉化換桿件名稱為 C2-2F  

                            col_number = col_number.Replace("\"", "");
                            floor_number = floor_number.Replace("\"", "");

                            assign_temp[9] = "\"" + col_number + "-" + floor_number + "\"";

                            frame_assign_Out.Add(cout_assign, assign_temp);

                            cout_assign++;


                            break;
                        }

                    }


                }

            }
            /////梁鋼筋
            Dictionary<UInt32, string[]> BEAM_REBAR = dbi.Get_BEAM_REBAR_DATA();

            Dictionary<UInt32, string[]> BEAM_REBAR_OUT = new Dictionary<UInt32, string[]>();// 輸出

            UInt32 count_beam = 0;



            double i_Top = 0, j_Top = 0, i_Bot = 0, j_Bot = 0;

            

            foreach (KeyValuePair<UInt32, string[]> it1 in BEAM_REBAR)
            {
                string[] BEAM_REABR_temp = it1.Value;


                string[] BEAM_REBAR_OUT_temp = new string[14];

                string beam_numbe_out = "\"" + BEAM_REABR_temp[1] + "\"";
                string floor_number_out = "\"" + BEAM_REABR_temp[0] + "\"";

                BEAM_REBAR_OUT_temp[0] = "  CONCRETESECTION ";
                BEAM_REBAR_OUT_temp[1] = "\"" + BEAM_REABR_temp[1] + "-" + BEAM_REABR_temp[0] + "\"";
                BEAM_REBAR_OUT_temp[2] = "TYPE \"BEAM\"  COVERTOP  ";
                BEAM_REBAR_OUT_temp[3] = BEAM_COVER;
                BEAM_REBAR_OUT_temp[4] = "  COVERBOTTOM ";
                BEAM_REBAR_OUT_temp[5] = BEAM_COVER;


                i_Top = Convert.ToDouble(BEAM_REABR_temp[2]);
                j_Top = Convert.ToDouble(BEAM_REABR_temp[3]);
                i_Bot = Convert.ToDouble(BEAM_REABR_temp[4]);
                j_Bot = Convert.ToDouble(BEAM_REABR_temp[5]);

                double area_rebar = 0;

                if (BEAM_REABR_temp[7] == "#10")
                {
                    area_rebar = 8.14/10000;
                }
                else if (BEAM_REABR_temp[7] == "#8")
                {
                    area_rebar = 5.07 / 10000;
                }
                else if (BEAM_REABR_temp[7] == "#7")
                {
                    area_rebar = 3.87 / 10000;
                }
                else if (BEAM_REABR_temp[7] == "#6")
                {
                    area_rebar = 2.87 / 10000;
                }
                else
                {
                    MessageBox.Show("主筋沒定義.");
                }


                double rebar_i_TOP = i_Top * area_rebar;
                double rebar_j_TOP = j_Top * area_rebar;
                double rebar_i_BOT = i_Bot * area_rebar;
                double rebar_j_BOT = j_Bot * area_rebar;

                BEAM_REBAR_OUT_temp[6] = "ATI ";
                BEAM_REBAR_OUT_temp[7] = Convert.ToString(rebar_i_TOP);

                BEAM_REBAR_OUT_temp[8] = " ABI ";
                BEAM_REBAR_OUT_temp[9] = Convert.ToString(rebar_i_BOT);

                BEAM_REBAR_OUT_temp[10] = "ATJ ";
                BEAM_REBAR_OUT_temp[11] = Convert.ToString(rebar_j_TOP);

                BEAM_REBAR_OUT_temp[12] = "ABJ ";
                BEAM_REBAR_OUT_temp[13] = Convert.ToString(rebar_j_BOT);

                BEAM_REBAR_OUT.Add(count_beam, BEAM_REBAR_OUT_temp);// 輸出

                

                foreach (KeyValuePair<UInt32, string[]> it3 in frame_assign)
                {

                    string[] assign_temp = new string[30];
                    assign_temp = it3.Value;

                    string beam_number = assign_temp[4];
                    string floor_number = assign_temp[6];



                    if (beam_number== beam_numbe_out & floor_number== floor_number_out)
                    {

                        string beam_section = assign_temp[9];

                        string[] temp5 = new string[20];
                        string[] temp6 = new string[20]; 



                        frame_prop1.TryGetValue(beam_section, out temp5);
                        frame_prop2.TryGetValue(beam_section, out temp6);

                        beam_number = beam_number.Replace("\"", "");
                        floor_number = floor_number.Replace("\"", "");

                        temp5[4] = "\"" + beam_number + "-"  + floor_number + "\"";

                        string NewText3 = string.Join(" ", temp5, 0, temp5.Length);

                        char[] delimiterChars = { ',', '\t' };

                        temp5 = NewText3.Split(delimiterChars);

                        frame_prop_out1.Add(count_prop1, temp5);
                        count_prop1++;

                        if (temp6.Length > 0)
                        {
                            temp6[4] = "\"" + beam_number + "-" +  floor_number + "\"";

                            string NewText4 = string.Join(" ", temp6, 0, temp6.Length);

                            temp6 = NewText4.Split(delimiterChars);

                            frame_prop_out2.Add(count_prop2, temp6);
                            count_prop2++;
                        }
                        /// 輸出beam HINGE
                        /// 


                        string[] temp3 = new string[5];
                        string[] temp4 = new string[5];
                        // I 端
                        temp3[0] = "  LINEASSIGN ";
                        temp3[1] = "\"" + beam_number + "\"" + " ";
                        temp3[2] = "\"" + floor_number + "\"" + " ";
                        temp3[3] = "HINGE \"Default-M3\"  RDISTANCE ";
                        temp3[4] = BEAM_hinge_i_loc;


                        frame_beam_hinge.Add(count_hinge_beam, temp3);
                        count_hinge_beam++;
                        // J 端
                        temp4[0] = "  LINEASSIGN ";
                        temp4[1] = "\"" + beam_number + "\"" + " ";
                        temp4[2] = "\"" + floor_number + "\"" + " ";
                        temp4[3] = "HINGE \"Default-M3\"  RDISTANCE ";
                        temp4[4] = BEAM_hinge_j_loc;

                        frame_beam_hinge.Add(count_hinge_beam, temp4);
                        count_hinge_beam++;

                        ///// 轉化換桿件名稱為 B2-2F  

                        beam_number = beam_number.Replace("\"", "");
                        floor_number = floor_number.Replace("\"", "");

                        assign_temp[9] = "\"" + beam_number + "-" + floor_number + "\"";

                        frame_assign_Out.Add(cout_assign, assign_temp);

                        cout_assign++;

                        break;
                    }


                }
                count_beam++;


            }

            /////////////////////////////////
            foreach (KeyValuePair<UInt32, string[]> it3 in frame_assign)
            {
                string[] assign_temp = new string[30];
                assign_temp  = it3.Value;

                string beam_number = assign_temp[4];
                string floor_number = assign_temp[6];
            
                string beam_number2 ;
                string floor_number2;

                bool check_out = false;

                foreach (KeyValuePair<UInt32, string[]>  it4 in frame_assign_Out)
                {
                    string[] assign_temp2 = it4.Value;

                    beam_number2 = assign_temp2[4];
                    floor_number2 = assign_temp2[6];


                    if(beam_number== beam_number2 & floor_number== floor_number2)
                    {
                        check_out = false;
                        break;
                    }
                    else
                    {
                        check_out = true;
                    }

                }
                
                if(check_out == true)
                {         
                    frame_assign_Out.Add(cout_assign, assign_temp);
                    cout_assign++;

                }
                

            }


                ////////////////

                /********************************************
                UInt32 cout_assign = 0;

                foreach (KeyValuePair<UInt32, string[]> it3 in frame_assign)
                {

                    string[] assign_temp = it3.Value;

                    string name_number = assign_temp[4];
                    string floor_number = assign_temp[6];

                    name_number = name_number.Replace("\"", "");
                    floor_number = floor_number.Replace("\"", "");

                    assign_temp[9] = "\"" + name_number + "-" + floor_number + "\"";


                    frame_assign_Out.Add(cout_assign, assign_temp);
                    cout_assign++;
                }
                ********************************************************/






            string Out_file = "PUSH.e2k";

            using (StreamWriter sw2 = new StreamWriter(Out_file, false))
            {

                Dictionary<UInt32, string[]> out_txt = dbi.Get_ETABS_TXT();

                foreach (KeyValuePair<UInt32, string[]> it1 in out_txt)
                {
                    string[] out_line = it1.Value;


                    if (out_line.Length >= 2) //當欄位數大於2, 
                    {
                        /////////////////////////
                        string NewText1 = string.Join(" ", out_line, 0, out_line.Length);


                        if (NewText1 == "$ FRAME SECTIONS")
                        {
                            sw2.WriteLine("$ FRAME SECTIONS");

                            ///////////////////////////////
                            foreach (KeyValuePair<UInt32, StringBuilder> it21 in frame_prop_out1_temp)
                            {
                                StringBuilder TXT_OUT_temp = it21.Value;

                                string TXT_OUT = TXT_OUT_temp.ToString();

                                sw2.WriteLine(TXT_OUT); //寫入並換行
                            }

                            foreach (KeyValuePair<UInt32, StringBuilder> it21 in frame_prop_out2_temp)
                            {
                                StringBuilder TXT_OUT_temp = it21.Value;

                                string TXT_OUT = TXT_OUT_temp.ToString();

                                sw2.WriteLine(TXT_OUT); //寫入並換行
                            }


                            /////////////////////////////////////
                            foreach (KeyValuePair<UInt32, string[]> it21 in frame_prop_out1)
                            {

                                string[] out_line_temp = it21.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }
                            ////
                            foreach (KeyValuePair<UInt32, string[]> it22 in frame_prop_out2)
                            {

                                string[] out_line_temp = it22.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }

                        }



                        if (NewText1 == "$ CONCRETE SECTIONS")
                        {

                            //////////////////////////////////////////////
                            sw2.WriteLine("$ CONCRETE SECTIONS");
                            ////
                            foreach (KeyValuePair<UInt32, string[]> it23 in COL_REBAR_OUT)
                            {

                                string[] out_line_temp = it23.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }
                            /////
                            foreach (KeyValuePair<UInt32, string[]> it24 in BEAM_REBAR_OUT)
                            {

                                string[] out_line_temp = it24.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }
                            /////
                        }

                        if (NewText1 == "$ LINE ASSIGNS")
                        {

                            sw2.WriteLine("$ LINE ASSIGNS");

                            foreach (KeyValuePair<UInt32, string[]> it25 in frame_col_hinge)
                            {

                                string[] out_line_temp = it25.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }
                            /////


                            foreach (KeyValuePair<UInt32, string[]> it26 in frame_beam_hinge)
                            {

                                string[] out_line_temp = it26.Value;
                                string NewText = string.Join(" ", out_line_temp, 0, out_line_temp.Length);
                                sw2.WriteLine(NewText); //寫入並換行
                            }

                        }

                        sw2.WriteLine(NewText1); //寫入並換行



                    }
                }


                
                



                sw2.Close();

            }





        }
    }
}



