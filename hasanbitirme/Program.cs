using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace hasanbitirme
{
    internal class Program
    {
        static void Main()
        {
            
            Console.Clear();

            Console.WriteLine("*************************************");           
            Console.WriteLine("warning('off');"); 
            
            DeleteFileIfExists("Dengkx1.xls");
            DeleteFileIfExists("Dengky1.xls");

            string DOSYA = "caspary_2.xlsx";

           
            string filePath = @"C:\Users\hasan\OneDrive\Masaüstü\caspary_2.txt";
            using (StreamWriter sw = new StreamWriter(filePath))
            {

                Console.WriteLine("+İTERASYONLU SERBEST DENGELEME SONUÇ RAPORU (DOĞRULTU-KENAR AÇI)+");
                Console.WriteLine("+Research Asistant Berkant KONAKOĞLU+");
                Console.WriteLine("*************************************");


                double[][] DOG = ReadExcelData(DOSYA, "doğrultu");
                double[][] KEN = ReadExcelData(DOSYA, "kenar");
                double[][] KOORD = ReadExcelData(DOSYA, "koord");


                double[] NN = KOORD.Select(row => row[0]).ToArray();
                double[] Y1 = KOORD.Select(row => row[1]).ToArray();
                double[] X1 = KOORD.Select(row => row[2]).ToArray();


                double top_y1 = Y1.Sum();
                double top_x1 = X1.Sum();
                double mean_y1 = Y1.Average();
                double mean_x1 = X1.Average();


                double[] koor_ort1X = new double[NN.Length];
                for (int i = 0; i < NN.Length; i++)
                {
                    koor_ort1X[i] = X1[i] - mean_x1;
                }


                double[] koor_ort1Y = new double[NN.Length];
                for (int i = 0; i < NN.Length; i++)
                {
                    koor_ort1Y[i] = Y1[i] - mean_y1;
                }

                // BÖLÜNMEYEN NOKTALARIN YAKLAŞIK KOORDİNATLARININ YAZDIRILMASI


                Console.WriteLine("++++++++++++++++++++++++++++++++++++++++++++++++++");
                Console.WriteLine("++ BÖLÜNMEYEN NOKTALARIN YAKLAŞIK KOORDİNATLARI ++");
                Console.WriteLine("++++++++++++++++++++++++++++++++++++++++++++++++++");
                Console.WriteLine("  NNO              Y(m)               X(m)");

                for (int i = 0; i < NN.Length; i++)
                {
                    sw.WriteLine($"{NN[i],5}   {Y1[i],18:F5}  {X1[i],18:F5}");
                }

                Console.WriteLine("--------------------------------------------------");
                Console.WriteLine();



                double[] DN = DOG.Select(row => row[0]).ToArray();
                double[] BN = DOG.Select(row => row[1]).ToArray();
                double[] DOD = DOG.Select(row => row[2]).ToArray();
                double[] Pvek_d = DOG.Select(row => row[7]).ToArray();
                int nd = DOG.GetLength(0);

                double S0 = 1;

                
                double[] DNK = KEN.Select(row => row[0]).ToArray();
                double[] BNK = KEN.Select(row => row[1]).ToArray();
                double[] KENAR = KEN.Select(row => row[2]).ToArray();
                double[] Pvek_k = KEN.Select(row => row[7]).ToArray();
                int nk = KEN.GetLength(0);

                
            }
        }

        static double[][] ReadExcelData(string filePath, string sheetName)
        {
            Console.WriteLine($"Reading data from sheet: {sheetName}");

            return new double[0][];
        }

        static void DeleteFileIfExists(string filePath)
        {           
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

    }
}