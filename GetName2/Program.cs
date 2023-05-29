using System.Data;
using System.Reflection;
using System.Text;
using ExcelDataReader;

namespace GetName2
{
    internal class Program
    {
        static List<TianRenDi> TianRenDiList = new List<TianRenDi>()
        {
            //new TianRenDi() { Tian = 1, Ren = 1, Di = 1 },
            new TianRenDi() { Tian = 2, Ren = 2, Di = 2 },

            //new TianRenDi() { Tian = 1, Ren = 1, Di = 3 },
            new TianRenDi() { Tian = 2, Ren = 2, Di = 4 },

            //new TianRenDi() { Tian = 1, Ren = 1, Di = 5 },
            new TianRenDi() { Tian = 2, Ren = 2, Di = 6 },
            
            //new TianRenDi() { Tian = 1, Ren = 3, Di = 5 },
            new TianRenDi() { Tian = 2, Ren = 4, Di = 6 },
            
            //new TianRenDi() { Tian = 1, Ren = 5, Di = 3 },
            new TianRenDi() { Tian = 2, Ren = 6, Di = 4 },

            //new TianRenDi() { Tian = 3, Ren = 1, Di = 1 },
            new TianRenDi() { Tian = 4, Ren = 2, Di = 2 },

            //new TianRenDi() { Tian = 3, Ren = 1, Di = 3 },
            new TianRenDi() { Tian = 4, Ren = 2, Di = 4 },

            //new TianRenDi() { Tian = 3, Ren = 1, Di = 5 },
            new TianRenDi() { Tian = 4, Ren = 2, Di = 6 },

            //new TianRenDi() { Tian = 3, Ren = 3, Di = 1 },
            new TianRenDi() { Tian = 4, Ren = 4, Di = 2 },

            //new TianRenDi() { Tian = 3, Ren = 5, Di = 3 },
            new TianRenDi() { Tian = 4, Ren = 6, Di = 4 },

            //new TianRenDi() { Tian = 3, Ren = 5, Di = 5 },
            new TianRenDi() { Tian = 4, Ren = 6, Di = 6 },

            //new TianRenDi() { Tian = 5, Ren = 3, Di = 5 },
            new TianRenDi() { Tian = 6, Ren = 4, Di = 6 },

            //new TianRenDi() { Tian = 5, Ren = 5, Di = 3 },
            new TianRenDi() { Tian = 6, Ren = 6, Di = 4 },

            //new TianRenDi() { Tian = 5, Ren = 5, Di = 7 },
            new TianRenDi() { Tian = 6, Ren = 6, Di = 8 },

            //new TianRenDi() { Tian = 5, Ren = 7, Di = 5 },
            new TianRenDi() { Tian = 6, Ren = 8, Di = 6 },

            //new TianRenDi() { Tian = 7, Ren = 5, Di = 5 },
            new TianRenDi() { Tian = 8, Ren = 6, Di = 6 },

            //new TianRenDi() { Tian = 7, Ren = 5, Di = 7 },
            new TianRenDi() { Tian = 8, Ren = 6, Di = 8 },

            //new TianRenDi() { Tian = 9, Ren = 1, Di = 1 },
            new TianRenDi() { Tian = 0, Ren = 2, Di = 2 },

            //new TianRenDi() { Tian = 9, Ren = 1, Di = 5 },
            new TianRenDi() { Tian = 0, Ren = 2, Di = 6 },

            //new TianRenDi() { Tian = 9, Ren = 7, Di = 5 },
            new TianRenDi() { Tian = 0, Ren = 8, Di = 6 },
        };
        static Dictionary<int, int> ModNumMap = new Dictionary<int, int>()
        {
            {1, 2},
            {2, 2},
            {3, 4},
            {4, 4},
            {5, 6},
            {6, 6},
            {7, 8},
            {8, 8},
            {9, 0},
            {0, 0},
        };
        static Dictionary<int, bool> RenDiWaiZong = new Dictionary<int, bool>()
        {
            { 1 , true },
            { 3 , true },
            { 5 , true },
            { 6 , true }, //
            { 7 , true },
            { 8 , true },
            { 11, true },
            { 13, true },
            { 15, true },
            { 16, true },
            { 18, true },
            { 24, true },
            { 25, true },
            { 30, true }, //
            { 31, true },
            { 32, true },
            { 35, true },
            { 37, true },
            { 38, true }, //
            { 41, true },
            { 45, true },
            { 47, true },
            { 48, true },
            { 49, true }, //
            { 51, true }, //
            { 52, true },
            { 55, true }, //
            { 57, true },
            { 58, true }, //
            { 61, true },
            { 63, true },
            { 65, true },
            { 67, true },
            { 68, true },
            { 71, true }, //
            { 73, true }, //
            { 75, true }, //
            { 81, true },
        };

        public static int XingCount;
        public static Dictionary<char, WordInfo> wordInfoMap = new Dictionary<char, WordInfo>();

        static StringBuilder stringBuilder = new StringBuilder();
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("仅支持单行，3个字的姓名。可自行修改。");
            Console.Write("请输入库Excel路径：");
            string excelPath = Console.ReadLine();
            var dataSet = ExcelToDataSet(excelPath, false);
            List<WordInfo> words = new List<WordInfo>();
            for (int i = 1; i <= 30; ++i)
            {
                string tableName = string.Format("Word{0}", i);
                DataTable dt = dataSet.Tables[tableName];
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    var row = dt.Rows[j];
                    words.Add(new WordInfo()
                    {
                        SimpleChar = row[0].ToString()[0],
                        PinYin = row[1].ToString(),
                        TraditionalStr = row[2].ToString(),
                        BiHua = int.Parse(row[3].ToString()),
                        ShuXing = row[4].ToString()[0]
                    });
                    if (words[words.Count - 1].BiHua != i)
                    {
                        throw new Exception("???");
                    }
                }
            }

            foreach (WordInfo wordInfo in words)
            {
                wordInfoMap[wordInfo.SimpleChar] = wordInfo;
            }
            Console.WriteLine("All Words Count: {0}", wordInfoMap.Count);

            Console.Write("请输入姓：");
            string xingStr = Console.ReadLine();
            XingCount = wordInfoMap[xingStr[0]].BiHua;

            List<string> names = new List<string>();
            System.Object oooo = new object();
            System.Object bbbb = new object();
            int currentIndex = 0;
            for (int i = 0; i < Environment.ProcessorCount; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback((obj) =>
                {
                RESTART:
                    string name = "";
                    lock (oooo)
                    {
                        if (currentIndex >= words.Count) return;
                        name = name + words[currentIndex++].SimpleChar;

                        Console.WriteLine(string.Format("------------------ {0} ------------------", currentIndex));
                    }

                    List<string> tempNames = new List<string>();
                    for (int j = 0; j < words.Count; j++)
                    {
                        string name2 = name + words[j].SimpleChar;
                        if (IsMatch(name2))
                        {
                            tempNames.Add(name2);
                        }
                    }
                    lock (bbbb)
                    {
                        for (int i = 0; i < tempNames.Count; i++)
                        {
                            PrintStrokeNumber(tempNames[i]);
                        }
                        names.AddRange(tempNames);
                    }
                    goto RESTART;
                }));
            }

            while (true)
            {
                Thread.Sleep(1000);
                if (currentIndex >= words.Count)
                {
                    if (File.Exists("1.txt"))
                    {
                        File.Delete("1.txt");
                    }
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter("1.txt"))
                    {
                        file.Write(stringBuilder.ToString());
                    }
                    break;
                }
            }
            Console.WriteLine("所有数据已写入1.txt。");
            Console.ReadKey();
        }

        static void PrintStrokeNumber(string name)
        {
            Console.Write(name);
            Console.Write(" ");
            for (int i = 0; i < name.Length; i++)
            {
                Console.Write(wordInfoMap[name[i]].ShuXing);
            }
            Console.WriteLine();

            stringBuilder.Append(name);
            stringBuilder.Append(" ");
            for (int i = 0; i < name.Length; i++)
            {
                stringBuilder.Append(wordInfoMap[name[i]].ShuXing);
            }
            stringBuilder.AppendLine();
            for (int i = 0; i < name.Length; i++)
            {
                string pinyin = wordInfoMap[name[i]].PinYin;

                Console.Write(pinyin);
                Console.Write(" ");
                stringBuilder.Append(pinyin);
                stringBuilder.Append(" ");
            }

            Console.WriteLine();
            stringBuilder.AppendLine();

            for (int i = 0; i < name.Length; i++)
            {
                string c = wordInfoMap[name[i]].TraditionalStr;
                Console.Write(c);
                Console.Write(wordInfoMap[name[i]].BiHua);
                Console.Write(" ");

                stringBuilder.Append(c);
                stringBuilder.Append(wordInfoMap[name[i]].BiHua);
                stringBuilder.Append(" ");
            }
            Console.WriteLine();
            Console.Write(string.Format("人：{0}，地：{1}，外：{2}，总：{3}", Ren(name), Di(name), Wai(name), Zong(name)));
            Console.WriteLine();
            Console.WriteLine();

            stringBuilder.AppendLine();
            stringBuilder.Append(string.Format("人：{0}，地：{1}，外：{2}，总：{3}", Ren(name), Di(name), Wai(name), Zong(name)));
            stringBuilder.AppendLine();
            stringBuilder.AppendLine();
        }

        static bool IsMatch(string name)
        {
            int ren = Ren(name);
            if (!RenDiWaiZong.ContainsKey(ren)) return false;
            int di = Di(name);
            if (!RenDiWaiZong.ContainsKey(di)) return false;
            //int wai = Wai(name);
            //if (!RenDiWaiZong.ContainsKey(wai)) return false;
            int zong = Zong(name);
            if (!RenDiWaiZong.ContainsKey(zong)) return false;

            int tian = ModNumMap[Tian() % 10];
            ren = ModNumMap[ren % 10];
            di = ModNumMap[di % 10];

            for (int i = 0; i < TianRenDiList.Count; i++)
            {
                if (TianRenDiList[i].Tian == tian && TianRenDiList[i].Ren == ren && TianRenDiList[i].Di == di)
                {
                    return true;
                }
            }
            return false;
        }

        static int Tian()
        {
            return XingCount + 1;
        }

        static int Ren(string str)
        {
            return XingCount + wordInfoMap[str[0]].BiHua;
        }

        static int Di(string str)
        {
            int count = 0;
            for (int i = 0; i < str.Length; i++)
            {
                count = count + wordInfoMap[str[i]].BiHua;
            }
            return count;
        }

        static int Zong(string str)
        {
            return XingCount + Di(str);
        }

        static int Wai(string str)
        {
            return Zong(str) - Ren(str) + 1;
        }

        public static DataSet ExcelToDataSet(string physicalPath, bool FirstRowAsColumnNames)
        {
            FileStream stream = File.Open(physicalPath, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader;

            if (System.IO.Path.GetExtension(physicalPath) == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            DataSet result = excelReader.AsDataSet();

            excelReader.Close();
            stream.Close();

            return result;
        }

    }

    public class WordInfo
    {
        public char SimpleChar;
        public string PinYin;
        public string TraditionalStr;
        public int BiHua;
        public char ShuXing;
    }

    public class TianRenDi
    {
        public int Tian { get; set; }
        public int Ren { get; set; }
        public int Di { get; set; }
    }

}