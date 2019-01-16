using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpssLib.SpssDataset;
using SpssLib.DataReader;
using System.IO;

namespace SRJabodetabek
{
    class Program
    {
        static void Main(string[] args)
        {
            int SurveyID = 600606;
            //string SurveyPeriod = "2014-09-30";//survey period
            string AttendedOn = "2018-07-31";
            //string year = getYear(SurveyPeriod); 
            string country = "Indonesia";//survey country
            DB_insertion_SRHW_Jabodetabek iobj = new DB_insertion_SRHW_Jabodetabek();
            string questions = "id,weight,S9,S2,B1,B2_4,B2_6,B2_7,B2_8,B2_9,B2_11,B2_12,B2_13,B2_14,B2_15,B2_19,B2_21,B2_23,B4a,B5_4,B5_6,B5_7,B5_8,B5_9,B5_11,B5_12,B5_13,B5_14,B5_15,B5_19,B5_21,B5_23,B5b_1,B5b_2,B5b_3,B5b_4,B5b_5,B5b_6,B5b_7,B10,B3,B6_4,B6_6,B6_7,B6_8,B6_9,B6_11,B6_12,B6_13,B6_14,B6_15,B6_19,B6_21,B6_23,B7_4,B7_6,B7_7,B7_8,B7_9,B7_11,B7_12,B7_13,B7_14,B7_15,B7_19,B7_21,B7_23,B8_4,B8_6,B8_7,B8_8,B8_9,B8_11,B8_12,B8_13,B8_14,B8_15,B8_19,B8_21,B8_23,B37a11,B37a12,B37a13,B37a14,B37a15,B37a16,B37a17,B37a18,B37a19,B37a20,B37a21,B37a22,B37a23,B11a_1,B11a_2,B11a_3,B11a_4,B11a_5,B11b_1,B11b_2,B11b_3,B11b_4,B11b_5,B12,B13_1,B13_2,B13_3,B13_4,B13_5,B13_6,B13_7,B13_8,B14_1,B14_2,B14_3,B14_4,B14_5,B14_6,B14_7,B14_8,B15_1,B15_2,B15_3,B15_4,B15_5,B15_6,B15_7,B15_8,B15_9,B15_10,B16_1,B16_2,B16_3,B16_4,B16_5,B16_6,B17_1,B17_2,B17_3,B17_4,B17_5,B17_6,B19a,B21_1,B21_2,B21_3,B21_4,B21_5,B21_6,B21_7,B21_8,B21_9,B21_10,B21_11,B21_12,B21_13,B21_14,B21_15,B21_16,B22b_1,B22b_2,B22b_3,B22b_4,B22b_5,B22b_6,B22b_7,B22b_8,B22b_9,B22b_10,B22b_11,B22b_12,B22b_13,B22b_14,B22b_15,B24_1,B24_2,B24_3,B24_4,B24_5,B24_6,B24_7,B24_8,B24_9,B25a,B25b,B25d,B25e,B25f,B26_1,B26_2,B26_3,B26_4,B26_5,B26_6,B26_7,B26_8,B26_9,B26_10,B26_11,B26_12,B26_13,B26_14,B26_15,B26_16,B28_1,B28_2,B28_3,B28_4,B28_5,B31_1,B31_2,B31_3,B31_4,B31_5,B32,HH33d_1,HH33d_2,HH33d_3,HH33d_4,HH33d_5,HH33d_6,HH33d_7,HH33d_8,HH33d_9,HH33d_10,HH33d_11,HH33d_12,HH33d_13,HH33d_14,HH33d_15,HH33d_16,HH33f_1,HH33f_2,HH33f_3,HH33f_4,HH33f_5,HH33i_1,HH33i_2,HH33i_3,HH33i_4,HH33i_5,HH33j,B35_1,B35_2,B35_3,B35_4,B35_5,B35_6,B35_7,B35_8,B35_9,B35_10,B35_11,B35_12,B35_13,B35_14,B36_1,B36_2,B36_3,B36_4,B36_5,B36_6,B36_7,B40_1,B40_2,B40_3,B40_4,B41b_1,B41b_2,B41b_3,B41b_4,B42_1,B42_2,B43a_1,B43a_2,P1,P2_1,P2_2,P2_3,P2_4,P2_5,P2_6,P2_7,P2_8,P2_9,P2_10,P3,C4_1,C4_2,C4_3,C4_4,C4_5,C4_6,C4_7,C3_1,C3_2,C3_3,C3_4,C3_5,C3_6,C3_7,C2_1,C2_2,C2_3,C2_4,C2_5,C2_6,C2_7,C1_1,C1_2,C1_3,C1_4,C1_5,C1_6,C1_7";// dashboard variable value          
            //string questions = "CHAP39,CHAP40";// dashboard variable value          
            string[] spss_variable_name = questions.Split(',');
            SpssReader spssDataset;
            using (FileStream fileStream = new FileStream(@"D:\spssparsing\SRHW\SRHW_Jabodetabek.sav", FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                spssDataset = new SpssReader(fileStream); // Create the reader, this will read the file header
                foreach (string spss_v in spss_variable_name)
                {
                    foreach (var variable in spssDataset.Variables)  // Iterate through all the varaibles
                    {
                        if (variable.Name.Equals(spss_v))
                        {
                            foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                            {
                                string VARIABLE_NAME = spss_v;
                                string VARIABLE_NAME_SUB_NAME = variable.Name;
                                string VARIABLE_ID = label.Key.ToString();
                                string VARIABLE_VALUE = label.Value;
                                string VARIABLE_NAME_QUESTION = variable.Label;
                                string BASE_VARIABLE_NAME = variable.Name;
                                iobj.insert_Dashboard_variable_values(VARIABLE_NAME, VARIABLE_NAME_SUB_NAME, VARIABLE_ID, VARIABLE_VALUE, VARIABLE_NAME_QUESTION, SurveyID, country, BASE_VARIABLE_NAME, AttendedOn);
                            }
                        }

                    }
                }
                foreach (var record in spssDataset.Records)
                {
                    string userID = null;
                    string u_id = null;
                    string variable_name;
                    decimal Weight = 0;
                    string Country = "-- Not Available --";
                    string AgeGroup = "-- Not Available --";

                    string SES = "-- Not Available --";
                    string BrTom = "-- Not Available --";
                    string BrBread_Talk = "-- Not Available --";
                    string BrGarmelia = "-- Not Available --";
                    string BrHolland_Bakery = "-- Not Available --";
                    string BrLauw = "-- Not Available --";
                    string BrLe_Gitt = "-- Not Available --";
                    string BrMaxim = "-- Not Available --";
                    string BrMr_Bread = "-- Not Available --";
                    string BrMy_Roti = "-- Not Available --";
                    string BrParoti = "-- Not Available --";
                    string BrPrime_Bread = "-- Not Available --";
                    string BrSari_Roti = "-- Not Available --";
                    string BrSharon = "-- Not Available --";
                    string BrSwanish = "-- Not Available --";
                    string AdTom = "-- Not Available --";
                    string AdBread_Talk = "-- Not Available --";
                    string AdGarmelia = "-- Not Available --";
                    string AdHolland_Bakery = "-- Not Available --";
                    string AdLauw = "-- Not Available --";
                    string AdLe_Gitt = "-- Not Available --";
                    string AdMaxim = "-- Not Available --";
                    string AdMr_Bread = "-- Not Available --";
                    string AdMy_Roti = "-- Not Available --";
                    string AdParoti = "-- Not Available --";
                    string AdPrime_Bread = "-- Not Available --";
                    string AdSari_Roti = "-- Not Available --";
                    string AdSharon = "-- Not Available --";
                    string AdSwanish = "-- Not Available --";
                    string B5b_1 = "-- Not Available --";
                    string B5b_2 = "-- Not Available --";
                    string B5b_3 = "-- Not Available --";
                    string B5b_4 = "-- Not Available --";
                    string B5b_5 = "-- Not Available --";
                    string B5b_6 = "-- Not Available --";
                    string B5b_7 = "-- Not Available --";
                    string Bumo = "-- Not Available --";
                    string Favourite_Brand = "-- Not Available --";
                    string ConL3M_Bread_Talk = "-- Not Available --";
                    string ConL3M_Garmelia = "-- Not Available --";
                    string ConL3M_Holland_Bakery = "-- Not Available --";
                    string ConL3M_Lauw = "-- Not Available --";
                    string ConL3M_Le_Gitt = "-- Not Available --";
                    string ConL3M_Maxim = "-- Not Available --";
                    string ConL3M_Mr_Bread = "-- Not Available --";
                    string ConL3M_My_Roti = "-- Not Available --";
                    string ConL3M_Paroti = "-- Not Available --";
                    string ConL3M_Prime_Bread = "-- Not Available --";
                    string ConL3M_Sari_Roti = "-- Not Available --";
                    string ConL3M_Sharon = "-- Not Available --";
                    string ConL3M_Swanish = "-- Not Available --";
                    string ConL1M_Bread_Talk = "-- Not Available --";
                    string ConL1M_Garmelia = "-- Not Available --";
                    string ConL1M_Holland_Bakery = "-- Not Available --";
                    string ConL1M_Lauw = "-- Not Available --";
                    string ConL1M_Le_Gitt = "-- Not Available --";
                    string ConL1M_Maxim = "-- Not Available --";
                    string ConL1M_Mr_Bread = "-- Not Available --";
                    string ConL1M_My_Roti = "-- Not Available --";
                    string ConL1M_Paroti = "-- Not Available --";
                    string ConL1M_Prime_Bread = "-- Not Available --";
                    string ConL1M_Sari_Roti = "-- Not Available --";
                    string ConL1M_Sharon = "-- Not Available --";
                    string ConL1M_Swanish = "-- Not Available --";
                    string ConL1W_Bread_Talk = "-- Not Available --";
                    string ConL1W_Garmelia = "-- Not Available --";
                    string ConL1W_Holland_Bakery = "-- Not Available --";
                    string ConL1W_Lauw = "-- Not Available --";
                    string ConL1W_Le_Gitt = "-- Not Available --";
                    string ConL1W_Maxim = "-- Not Available --";
                    string ConL1W_Mr_Bread = "-- Not Available --";
                    string ConL1W_My_Roti = "-- Not Available --";
                    string ConL1W_Paroti = "-- Not Available --";
                    string ConL1W_Prime_Bread = "-- Not Available --";
                    string ConL1W_Sari_Roti = "-- Not Available --";
                    string ConL1W_Sharon = "-- Not Available --";
                    string ConL1W_Swanish = "-- Not Available --";
                    string B37a11 = "-- Not Available --";
                    string B37a12 = "-- Not Available --";
                    string B37a13 = "-- Not Available --";
                    string B37a14 = "-- Not Available --";
                    string B37a15 = "-- Not Available --";
                    string B37a16 = "-- Not Available --";
                    string B37a17 = "-- Not Available --";
                    string B37a18 = "-- Not Available --";
                    string B37a19 = "-- Not Available --";
                    string B37a20 = "-- Not Available --";
                    string B37a21 = "-- Not Available --";
                    string B37a22 = "-- Not Available --";
                    string B37a23 = "-- Not Available --";
                    string B11a_1 = "-- Not Available --";
                    string B11a_2 = "-- Not Available --";
                    string B11a_3 = "-- Not Available --";
                    string B11a_4 = "-- Not Available --";
                    string B11a_5 = "-- Not Available --";
                    string B11b_1 = "-- Not Available --";
                    string B11b_2 = "-- Not Available --";
                    string B11b_3 = "-- Not Available --";
                    string B11b_4 = "-- Not Available --";
                    string B11b_5 = "-- Not Available --";
                    string B12 = "-- Not Available --";
                    string B13_1 = "-- Not Available --";
                    string B13_2 = "-- Not Available --";
                    string B13_3 = "-- Not Available --";
                    string B13_4 = "-- Not Available --";
                    string B13_5 = "-- Not Available --";
                    string B13_6 = "-- Not Available --";
                    string B13_7 = "-- Not Available --";
                    string B13_8 = "-- Not Available --";
                    string B14_1 = "-- Not Available --";
                    string B14_2 = "-- Not Available --";
                    string B14_3 = "-- Not Available --";
                    string B14_4 = "-- Not Available --";
                    string B14_5 = "-- Not Available --";
                    string B14_6 = "-- Not Available --";
                    string B14_7 = "-- Not Available --";
                    string B14_8 = "-- Not Available --";
                    string B15_1 = "-- Not Available --";
                    string B15_2 = "-- Not Available --";
                    string B15_3 = "-- Not Available --";
                    string B15_4 = "-- Not Available --";
                    string B15_5 = "-- Not Available --";
                    string B15_6 = "-- Not Available --";
                    string B15_7 = "-- Not Available --";
                    string B15_8 = "-- Not Available --";
                    string B15_9 = "-- Not Available --";
                    string B15_10 = "-- Not Available --";
                    string B16_1 = "-- Not Available --";
                    string B16_2 = "-- Not Available --";
                    string B16_3 = "-- Not Available --";
                    string B16_4 = "-- Not Available --";
                    string B16_5 = "-- Not Available --";
                    string B16_6 = "-- Not Available --";
                    string B17_1 = "-- Not Available --";
                    string B17_2 = "-- Not Available --";
                    string B17_3 = "-- Not Available --";
                    string B17_4 = "-- Not Available --";
                    string B17_5 = "-- Not Available --";
                    string B17_6 = "-- Not Available --";
                    string B19a = "-- Not Available --";
                    string B21_1 = "-- Not Available --";
                    string B21_2 = "-- Not Available --";
                    string B21_3 = "-- Not Available --";
                    string B21_4 = "-- Not Available --";
                    string B21_5 = "-- Not Available --";
                    string B21_6 = "-- Not Available --";
                    string B21_7 = "-- Not Available --";
                    string B21_8 = "-- Not Available --";
                    string B21_9 = "-- Not Available --";
                    string B21_10 = "-- Not Available --";
                    string B21_11 = "-- Not Available --";
                    string B21_12 = "-- Not Available --";
                    string B21_13 = "-- Not Available --";
                    string B21_14 = "-- Not Available --";
                    string B21_15 = "-- Not Available --";
                    string B21_16 = "-- Not Available --";
                    string B22b_1 = "-- Not Available --";
                    string B22b_2 = "-- Not Available --";
                    string B22b_3 = "-- Not Available --";
                    string B22b_4 = "-- Not Available --";
                    string B22b_5 = "-- Not Available --";
                    string B22b_6 = "-- Not Available --";
                    string B22b_7 = "-- Not Available --";
                    string B22b_8 = "-- Not Available --";
                    string B22b_9 = "-- Not Available --";
                    string B22b_10 = "-- Not Available --";
                    string B22b_11 = "-- Not Available --";
                    string B22b_12 = "-- Not Available --";
                    string B22b_13 = "-- Not Available --";
                    string B22b_14 = "-- Not Available --";
                    string B22b_15 = "-- Not Available --";
                    string B24_1 = "-- Not Available --";
                    string B24_2 = "-- Not Available --";
                    string B24_3 = "-- Not Available --";
                    string B24_4 = "-- Not Available --";
                    string B24_5 = "-- Not Available --";
                    string B24_6 = "-- Not Available --";
                    string B24_7 = "-- Not Available --";
                    string B24_8 = "-- Not Available --";
                    string B24_9 = "-- Not Available --";
                    string B25a = "-- Not Available --";
                    string B25b = "-- Not Available --";
                    string B25d = "-- Not Available --";
                    string B25e = "-- Not Available --";
                    string B25f = "-- Not Available --";
                    string B26_1 = "-- Not Available --";
                    string B26_2 = "-- Not Available --";
                    string B26_3 = "-- Not Available --";
                    string B26_4 = "-- Not Available --";
                    string B26_5 = "-- Not Available --";
                    string B26_6 = "-- Not Available --";
                    string B26_7 = "-- Not Available --";
                    string B26_8 = "-- Not Available --";
                    string B26_9 = "-- Not Available --";
                    string B26_10 = "-- Not Available --";
                    string B26_11 = "-- Not Available --";
                    string B26_12 = "-- Not Available --";
                    string B26_13 = "-- Not Available --";
                    string B26_14 = "-- Not Available --";
                    string B26_15 = "-- Not Available --";
                    string B26_16 = "-- Not Available --";
                    string B28_1 = "-- Not Available --";
                    string B28_2 = "-- Not Available --";
                    string B28_3 = "-- Not Available --";
                    string B28_4 = "-- Not Available --";
                    string B28_5 = "-- Not Available --";
                    string B31_1 = "-- Not Available --";
                    string B31_2 = "-- Not Available --";
                    string B31_3 = "-- Not Available --";
                    string B31_4 = "-- Not Available --";
                    string B31_5 = "-- Not Available --";
                    string B32 = "-- Not Available --";
                    string HH33a = "-- Not Available --";
                    string HH33b = "-- Not Available --";
                    string HH33c = "-- Not Available --";
                    string HH33d_1 = "-- Not Available --";
                    string HH33d_2 = "-- Not Available --";
                    string HH33d_3 = "-- Not Available --";
                    string HH33d_4 = "-- Not Available --";
                    string HH33d_5 = "-- Not Available --";
                    string HH33d_6 = "-- Not Available --";
                    string HH33d_7 = "-- Not Available --";
                    string HH33d_8 = "-- Not Available --";
                    string HH33d_9 = "-- Not Available --";
                    string HH33d_10 = "-- Not Available --";
                    string HH33d_11 = "-- Not Available --";
                    string HH33d_12 = "-- Not Available --";
                    string HH33d_13 = "-- Not Available --";
                    string HH33d_14 = "-- Not Available --";
                    string HH33d_15 = "-- Not Available --";
                    string HH33d_16 = "-- Not Available --";
                    string HH33f_1 = "-- Not Available --";
                    string HH33f_2 = "-- Not Available --";
                    string HH33f_3 = "-- Not Available --";
                    string HH33f_4 = "-- Not Available --";
                    string HH33f_5 = "-- Not Available --";
                    string HH33i_1 = "-- Not Available --";
                    string HH33i_2 = "-- Not Available --";
                    string HH33i_3 = "-- Not Available --";
                    string HH33i_4 = "-- Not Available --";
                    string HH33i_5 = "-- Not Available --";
                    string HH33j = "-- Not Available --";
                    string B35_1 = "-- Not Available --";
                    string B35_2 = "-- Not Available --";
                    string B35_3 = "-- Not Available --";
                    string B35_4 = "-- Not Available --";
                    string B35_5 = "-- Not Available --";
                    string B35_6 = "-- Not Available --";
                    string B35_7 = "-- Not Available --";
                    string B35_8 = "-- Not Available --";
                    string B35_9 = "-- Not Available --";
                    string B35_10 = "-- Not Available --";
                    string B35_11 = "-- Not Available --";
                    string B35_12 = "-- Not Available --";
                    string B35_13 = "-- Not Available --";
                    string B35_14 = "-- Not Available --";
                    string B36_1 = "-- Not Available --";
                    string B36_2 = "-- Not Available --";
                    string B36_3 = "-- Not Available --";
                    string B36_4 = "-- Not Available --";
                    string B36_5 = "-- Not Available --";
                    string B36_6 = "-- Not Available --";
                    string B36_7 = "-- Not Available --";
                    string B40_1 = "-- Not Available --";
                    string B40_2 = "-- Not Available --";
                    string B40_3 = "-- Not Available --";
                    string B40_4 = "-- Not Available --";
                    string B41b_1 = "-- Not Available --";
                    string B41b_2 = "-- Not Available --";
                    string B41b_3 = "-- Not Available --";
                    string B41b_4 = "-- Not Available --";
                    string B42_1 = "-- Not Available --";
                    string B42_2 = "-- Not Available --";
                    string B43a_1 = "-- Not Available --";
                    string B43a_2 = "-- Not Available --";
                    string P1 = "-- Not Available --";
                    string P2_1 = "-- Not Available --";
                    string P2_2 = "-- Not Available --";
                    string P2_3 = "-- Not Available --";
                    string P2_4 = "-- Not Available --";
                    string P2_5 = "-- Not Available --";
                    string P2_6 = "-- Not Available --";
                    string P2_7 = "-- Not Available --";
                    string P2_8 = "-- Not Available --";
                    string P2_9 = "-- Not Available --";
                    string P2_10 = "-- Not Available --";
                    string P3 = "-- Not Available --";
                    string C4_1 = "-- Not Available --";
                    string C4_2 = "-- Not Available --";
                    string C4_3 = "-- Not Available --";
                    string C4_4 = "-- Not Available --";
                    string C4_5 = "-- Not Available --";
                    string C4_6 = "-- Not Available --";
                    string C4_7 = "-- Not Available --";
                    string C3_1 = "-- Not Available --";
                    string C3_2 = "-- Not Available --";
                    string C3_3 = "-- Not Available --";
                    string C3_4 = "-- Not Available --";
                    string C3_5 = "-- Not Available --";
                    string C3_6 = "-- Not Available --";
                    string C3_7 = "-- Not Available --";
                    string C2_1 = "-- Not Available --";
                    string C2_2 = "-- Not Available --";
                    string C2_3 = "-- Not Available --";
                    string C2_4 = "-- Not Available --";
                    string C2_5 = "-- Not Available --";
                    string C2_6 = "-- Not Available --";
                    string C2_7 = "-- Not Available --";
                    string C1_1 = "-- Not Available --";
                    string C1_2 = "-- Not Available --";
                    string C1_3 = "-- Not Available --";
                    string C1_4 = "-- Not Available --";
                    string C1_5 = "-- Not Available --";
                    string C1_6 = "-- Not Available --";
                    string C1_7 = "-- Not Available --";
                    foreach (var variable in spssDataset.Variables)
                    {
                        foreach (string spss_v in spss_variable_name)
                        {
                            if (variable.Name.Equals(spss_v))
                            {
                                variable_name = variable.Name;

                                switch (variable_name)
                                {
                                    case "id":
                                        {
                                            u_id = Convert.ToString(record.GetValue(variable));
                                            userID = find_UserId(SurveyID, AttendedOn, u_id);
                                            break;
                                        }
                                    case "weight":
                                        {
                                            Weight = Convert.ToDecimal(record.GetValue(variable));
                                            break;
                                        }
                             
                                    case "S9":
                                        {
                                            SES = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                  
                                    case "S2":
                                        {
                                            AgeGroup = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "B1": { BrTom = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_4": { BrBread_Talk = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_6": { BrGarmelia = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_7": { BrHolland_Bakery = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_8": { BrLauw = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_9": { BrLe_Gitt = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_11": { BrMaxim = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_12": { BrMr_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_13": { BrMy_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_14": { BrParoti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_15": { BrPrime_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_19": { BrSari_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_21": { BrSharon = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B2_23": { BrSwanish = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B4a": { AdTom = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_4": { AdBread_Talk = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_6": { AdGarmelia = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_7": { AdHolland_Bakery = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_8": { AdLauw = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_9": { AdLe_Gitt = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_11": { AdMaxim = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_12": { AdMr_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_13": { AdMy_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_14": { AdParoti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_15": { AdPrime_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_19": { AdSari_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_21": { AdSharon = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5_23": { AdSwanish = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_1": { B5b_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_2": { B5b_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_3": { B5b_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_4": { B5b_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_5": { B5b_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_6": { B5b_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B5b_7": { B5b_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B10": { Bumo = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B3": { Favourite_Brand = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_4": { ConL3M_Bread_Talk = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_6": { ConL3M_Garmelia = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_7": { ConL3M_Holland_Bakery = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_8": { ConL3M_Lauw = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_9": { ConL3M_Le_Gitt = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_11": { ConL3M_Maxim = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_12": { ConL3M_Mr_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_13": { ConL3M_My_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_14": { ConL3M_Paroti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_15": { ConL3M_Prime_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_19": { ConL3M_Sari_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_21": { ConL3M_Sharon = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B6_23": { ConL3M_Swanish = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_4": { ConL1M_Bread_Talk = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_6": { ConL1M_Garmelia = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_7": { ConL1M_Holland_Bakery = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_8": { ConL1M_Lauw = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_9": { ConL1M_Le_Gitt = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_11": { ConL1M_Maxim = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_12": { ConL1M_Mr_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_13": { ConL1M_My_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_14": { ConL1M_Paroti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_15": { ConL1M_Prime_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_19": { ConL1M_Sari_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_21": { ConL1M_Sharon = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B7_23": { ConL1M_Swanish = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_4": { ConL1W_Bread_Talk = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_6": { ConL1W_Garmelia = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_7": { ConL1W_Holland_Bakery = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_8": { ConL1W_Lauw = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_9": { ConL1W_Le_Gitt = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_11": { ConL1W_Maxim = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_12": { ConL1W_Mr_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_13": { ConL1W_My_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_14": { ConL1W_Paroti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_15": { ConL1W_Prime_Bread = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_19": { ConL1W_Sari_Roti = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_21": { ConL1W_Sharon = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B8_23": { ConL1W_Swanish = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a11": { B37a11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a12": { B37a12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a13": { B37a13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a14": { B37a14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a15": { B37a15 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a16": { B37a16 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a17": { B37a17 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a18": { B37a18 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a19": { B37a19 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a20": { B37a20 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a21": { B37a21 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a22": { B37a22 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B37a23": { B37a23 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11a_1": { B11a_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11a_2": { B11a_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11a_3": { B11a_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11a_4": { B11a_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11a_5": { B11a_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11b_1": { B11b_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11b_2": { B11b_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11b_3": { B11b_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11b_4": { B11b_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B11b_5": { B11b_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B12": { B12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_1": { B13_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_2": { B13_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_3": { B13_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_4": { B13_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_5": { B13_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_6": { B13_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_7": { B13_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B13_8": { B13_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_1": { B14_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_2": { B14_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_3": { B14_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_4": { B14_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_5": { B14_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_6": { B14_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_7": { B14_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B14_8": { B14_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_1": { B15_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_2": { B15_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_3": { B15_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_4": { B15_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_5": { B15_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_6": { B15_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_7": { B15_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_8": { B15_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_9": { B15_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B15_10": { B15_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_1": { B16_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_2": { B16_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_3": { B16_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_4": { B16_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_5": { B16_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B16_6": { B16_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_1": { B17_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_2": { B17_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_3": { B17_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_4": { B17_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_5": { B17_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B17_6": { B17_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B19a": { B19a = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_1": { B21_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_2": { B21_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_3": { B21_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_4": { B21_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_5": { B21_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_6": { B21_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_7": { B21_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_8": { B21_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_9": { B21_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_10": { B21_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_11": { B21_11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_12": { B21_12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_13": { B21_13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_14": { B21_14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_15": { B21_15 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B21_16": { B21_16 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_1": { B22b_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_2": { B22b_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_3": { B22b_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_4": { B22b_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_5": { B22b_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_6": { B22b_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_7": { B22b_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_8": { B22b_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_9": { B22b_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_10": { B22b_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_11": { B22b_11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_12": { B22b_12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_13": { B22b_13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_14": { B22b_14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B22b_15": { B22b_15 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_1": { B24_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_2": { B24_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_3": { B24_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_4": { B24_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_5": { B24_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_6": { B24_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_7": { B24_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_8": { B24_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B24_9": { B24_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B25a": { B25a = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B25b": { B25b = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B25d": { B25d = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B25e": { B25e = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B25f": { B25f = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_1": { B26_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_2": { B26_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_3": { B26_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_4": { B26_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_5": { B26_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_6": { B26_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_7": { B26_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_8": { B26_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_9": { B26_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_10": { B26_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_11": { B26_11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_12": { B26_12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_13": { B26_13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_14": { B26_14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_15": { B26_15 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B26_16": { B26_16 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B28_1": { B28_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B28_2": { B28_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B28_3": { B28_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B28_4": { B28_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B28_5": { B28_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B31_1": { B31_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B31_2": { B31_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B31_3": { B31_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B31_4": { B31_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B31_5": { B31_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B32": { B32 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33a": { HH33a = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33b": { HH33b = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33c": { HH33c = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_1": { HH33d_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_2": { HH33d_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_3": { HH33d_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_4": { HH33d_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_5": { HH33d_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_6": { HH33d_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_7": { HH33d_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_8": { HH33d_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_9": { HH33d_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_10": { HH33d_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_11": { HH33d_11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_12": { HH33d_12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_13": { HH33d_13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_14": { HH33d_14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_15": { HH33d_15 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33d_16": { HH33d_16 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33f_1": { HH33f_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33f_2": { HH33f_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33f_3": { HH33f_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33f_4": { HH33f_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33f_5": { HH33f_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33i_1": { HH33i_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33i_2": { HH33i_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33i_3": { HH33i_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33i_4": { HH33i_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33i_5": { HH33i_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "HH33j": { HH33j = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_1": { B35_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_2": { B35_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_3": { B35_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_4": { B35_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_5": { B35_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_6": { B35_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_7": { B35_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_8": { B35_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_9": { B35_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_10": { B35_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_11": { B35_11 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_12": { B35_12 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_13": { B35_13 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B35_14": { B35_14 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_1": { B36_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_2": { B36_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_3": { B36_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_4": { B36_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_5": { B36_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_6": { B36_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B36_7": { B36_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B40_1": { B40_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B40_2": { B40_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B40_3": { B40_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B40_4": { B40_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B41b_1": { B41b_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B41b_2": { B41b_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B41b_3": { B41b_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B41b_4": { B41b_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B42_1": { B42_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B42_2": { B42_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B43a_1": { B43a_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "B43a_2": { B43a_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P1": { P1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_1": { P2_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_2": { P2_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_3": { P2_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_4": { P2_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_5": { P2_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_6": { P2_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_7": { P2_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_8": { P2_8 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_9": { P2_9 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P2_10": { P2_10 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "P3": { P3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_1": { C4_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_2": { C4_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_3": { C4_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_4": { C4_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_5": { C4_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_6": { C4_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C4_7": { C4_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_1": { C3_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_2": { C3_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_3": { C3_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_4": { C3_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_5": { C3_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_6": { C3_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C3_7": { C3_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_1": { C2_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_2": { C2_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_3": { C2_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_4": { C2_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_5": { C2_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_6": { C2_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C2_7": { C2_7 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_1": { C1_1 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_2": { C1_2 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_3": { C1_3 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_4": { C1_4 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_5": { C1_5 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_6": { C1_6 = Convert.ToString(record.GetValue(variable)); break; }
                                    case "C1_7": { C1_7 = Convert.ToString(record.GetValue(variable)); break; }
                                }

                            }
                        }
                    }

                    if (userID != null && Weight != 0)
                    {
                        iobj.insert_Dashboard_values(userID, AttendedOn, Weight, SurveyID, country, SES, AgeGroup, BrTom, BrBread_Talk, BrGarmelia, BrHolland_Bakery, BrLauw, BrLe_Gitt, BrMaxim, BrMr_Bread, BrMy_Roti, BrParoti, BrPrime_Bread, BrSari_Roti, BrSharon, BrSwanish, AdTom, AdBread_Talk, AdGarmelia, AdHolland_Bakery, AdLauw, AdLe_Gitt, AdMaxim, AdMr_Bread, AdMy_Roti, AdParoti, AdPrime_Bread, AdSari_Roti, AdSharon, AdSwanish, B5b_1, B5b_2, B5b_3, B5b_4, B5b_5, B5b_6, B5b_7, Bumo, Favourite_Brand, ConL3M_Bread_Talk, ConL3M_Garmelia, ConL3M_Holland_Bakery, ConL3M_Lauw, ConL3M_Le_Gitt, ConL3M_Maxim, ConL3M_Mr_Bread, ConL3M_My_Roti, ConL3M_Paroti, ConL3M_Prime_Bread, ConL3M_Sari_Roti, ConL3M_Sharon, ConL3M_Swanish, ConL1M_Bread_Talk, ConL1M_Garmelia, ConL1M_Holland_Bakery, ConL1M_Lauw, ConL1M_Le_Gitt, ConL1M_Maxim, ConL1M_Mr_Bread, ConL1M_My_Roti, ConL1M_Paroti, ConL1M_Prime_Bread, ConL1M_Sari_Roti, ConL1M_Sharon, ConL1M_Swanish, ConL1W_Bread_Talk, ConL1W_Garmelia, ConL1W_Holland_Bakery, ConL1W_Lauw, ConL1W_Le_Gitt, ConL1W_Maxim, ConL1W_Mr_Bread, ConL1W_My_Roti, ConL1W_Paroti, ConL1W_Prime_Bread, ConL1W_Sari_Roti, ConL1W_Sharon, ConL1W_Swanish, B37a11, B37a12, B37a13, B37a14, B37a15, B37a16, B37a17, B37a18, B37a19, B37a20, B37a21, B37a22, B37a23, B11a_1, B11a_2, B11a_3, B11a_4, B11a_5, B11b_1, B11b_2, B11b_3, B11b_4, B11b_5, B12, B13_1, B13_2, B13_3, B13_4, B13_5, B13_6, B13_7, B13_8, B14_1, B14_2, B14_3, B14_4, B14_5, B14_6, B14_7, B14_8, B15_1, B15_2, B15_3, B15_4, B15_5, B15_6, B15_7, B15_8, B15_9, B15_10, B16_1, B16_2, B16_3, B16_4, B16_5, B16_6, B17_1, B17_2, B17_3, B17_4, B17_5, B17_6, B19a, B21_1, B21_2, B21_3, B21_4, B21_5, B21_6, B21_7, B21_8, B21_9, B21_10, B21_11, B21_12, B21_13, B21_14, B21_15, B21_16, B22b_1, B22b_2, B22b_3, B22b_4, B22b_5, B22b_6, B22b_7, B22b_8, B22b_9, B22b_10, B22b_11, B22b_12, B22b_13, B22b_14, B22b_15, B24_1, B24_2, B24_3, B24_4, B24_5, B24_6, B24_7, B24_8, B24_9, B25a, B25b, B25d, B25e, B25f, B26_1, B26_2, B26_3, B26_4, B26_5, B26_6, B26_7, B26_8, B26_9, B26_10, B26_11, B26_12, B26_13, B26_14, B26_15, B26_16, B28_1, B28_2, B28_3, B28_4, B28_5, B31_1, B31_2, B31_3, B31_4, B31_5, B32, HH33a, HH33b, HH33c, HH33d_1, HH33d_2, HH33d_3, HH33d_4, HH33d_5, HH33d_6, HH33d_7, HH33d_8, HH33d_9, HH33d_10, HH33d_11, HH33d_12, HH33d_13, HH33d_14, HH33d_15, HH33d_16, HH33f_1, HH33f_2, HH33f_3, HH33f_4, HH33f_5, HH33i_1, HH33i_2, HH33i_3, HH33i_4, HH33i_5, HH33j, B35_1, B35_2, B35_3, B35_4, B35_5, B35_6, B35_7, B35_8, B35_9, B35_10, B35_11, B35_12, B35_13, B35_14, B36_1, B36_2, B36_3, B36_4, B36_5, B36_6, B36_7, B40_1, B40_2, B40_3, B40_4, B41b_1, B41b_2, B41b_3, B41b_4, B42_1, B42_2, B43a_1, B43a_2, P1, P2_1, P2_2, P2_3, P2_4, P2_5, P2_6, P2_7, P2_8, P2_9, P2_10, P3, C4_1, C4_2, C4_3, C4_4, C4_5, C4_6, C4_7, C3_1, C3_2, C3_3, C3_4, C3_5, C3_6, C3_7, C2_1, C2_2, C2_3, C2_4, C2_5, C2_6, C2_7, C1_1, C1_2, C1_3, C1_4, C1_5, C1_6, C1_7);
                    }
                }
            }
        }
        private static string find_UserId(int SurveyID, string SurveyPeriod, string u_id)
        {
            string sum = "";
            string[] date = SurveyPeriod.Split('-');
            foreach (string d in date)
            {
                sum = sum + d;

            }
            return u_id + SurveyID + sum;
        }
    }
}
