using System;
using System.Collections.Generic;
using System.Text;
//using System.Windows.Forms;
using System.IO;

namespace Power_Flow.类文件
{
    class PF_Calc
    {
        static short LINEMAX = 5000;              //最大线路数
        static short GENERATORMAX = 500;          //最大发电机数
        static short LOADMAX = 2000;		        //最大负荷数
        static short NODEMAX = 2000;		        //最大节点数
        static byte SWINGMAX = 20;		        //最大平衡节点数
        static short PVMAX = 500;			        //最大PV节点数
        static byte NODEFACTOR = 10;		        //导纳矩阵中非零非对角元素的个数相对于
        //最大节点数(NODEMAX)的倍数.
        static float Deg_to_Rad = 0.017453292f;    //度到弧度的转换系数
        static float Rad_to_Deg = 57.29577951f;	//弧度到度的转换系数
        static float SinglePai = 3.14159265f;	    //圆周率
        static float DoublePai = 6.2831853f;	    //两倍的圆周率

        short[] Node_Name_NewtoOld = new short[NODEMAX];//新节点名（号）-->旧节点名（号）
        //——保存按出线数由小到大排序的节点号
        short[] Node_Flag = new short[NODEMAX];			//节点类型标志：0-平衡,1-PQ,2-PV
        short[,] Line_NodeName = new short[LINEMAX, 2];	//线路的左、右节点新名（号）
        short[] Line_No_NewtoOld = new short[LINEMAX];	//新线路号-->旧线路号
        short[] Line_Flag = new short[LINEMAX];			//新线路的类型标志:0,1,2说明同Line结构
        short[] Gen_NodeName = new short[GENERATORMAX];	//发电机节点的新节点名（号）
        short[] Gen_No_NewtoOld = new short[GENERATORMAX];	//新发电机顺序号-->旧发电机顺序号
        short[,] Gen_SWNode = new short[SWINGMAX, 2];	//平衡节点数据：0-新节点名（号）;
        //1-对应的旧发电机顺序号
        short[,] Gen_PVNode = new short[PVMAX, 2];		//发电机PV节点数据：0-新节点名（号）;
        //1-对应的旧发电机顺序号
        short[,] Gen_PQNode = new short[GENERATORMAX, 2];//发电机PQ节点数据：0-新节点名（号）;
        //1-对应的旧发电机顺序号
        short[] Load_NodeName = new short[LOADMAX];		//负荷节点的新节点名（号）
        short[] Load_No_NewtoOld = new short[LOADMAX];	//新负荷顺序号-->旧负荷顺序号
        bool Sav_result = false;

        public Line[] LLine;
        public Generator[] GGen;
        public Load[] LLoad;
        StreamWriter sw;
        //读取数据
        private void Data_Input(out short Num_Line, out short Num_Gen, 
            out short Num_Load, out float Eps, out short Iter_Max,
        out short VolIni_Flag, out short VolRes_Flag, string path)
        {
            //DateTime BeginTime = DateTime.Now;
            //*********这段代码为输出参数赋初值***********//
            Num_Line = 0;
            Num_Gen = 0;
            Num_Load = 0;
            Eps = 1.0e-5f;
            Iter_Max = 200;
            VolIni_Flag = 0;
            VolRes_Flag = 0;
            //*********结束*****************************//
            //bool sign = true;
            TemporaryCulture.Start();
            FileStream stream = File.OpenRead(path);
            StreamReader input = new StreamReader(stream);
            ushort rows = 0;
            ushort j = 0; ushort p = 0; ushort q = 0;
            string dataLine = "";
            if ((dataLine = input.ReadLine()) != null)
            {
                //dataLine = dataLine.Trim().Replace(" ", ",");
                dataLine = System.Text.RegularExpressions.Regex.Replace(dataLine.Trim(), @" +", " ");

                string[] parts = dataLine.Split(' ');
                Num_Line = Convert.ToInt16(parts[0]);
                Num_Gen = Convert.ToInt16(parts[1]);
                Num_Load = Convert.ToInt16(parts[2]);
                Eps = Convert.ToSingle(parts[3]);
                Iter_Max = Convert.ToInt16(parts[4]);
                VolIni_Flag = Convert.ToInt16(parts[5]);
                VolRes_Flag = Convert.ToInt16(parts[6]);
                if (Convert.ToInt16(parts[7]) != 0)
                    Sav_result = true;
                else
                    Sav_result = false;
            }
            else
            {
                Console.WriteLine("no data find!");              
                Environment.Exit(1);
            }

            if (Num_Line > LINEMAX)
            {
                Console.WriteLine("Line Numbers > " + LINEMAX + "!");             
                Environment.Exit(1);
            }
            if (Num_Gen > GENERATORMAX)
            {
                Console.WriteLine("Generator Numbers > " + GENERATORMAX + "!");              
                Environment.Exit(1);
            }
            if (Num_Load > LOADMAX)
            {
                Console.WriteLine("Load Numbers > " + LOADMAX + "!");               
                Environment.Exit(1);
            }

            LLine = new Line[Num_Line];
            GGen = new Generator[Num_Gen];
            LLoad = new Load[Num_Load];
            while ((dataLine = input.ReadLine()) != null)
            {
                rows++;
                dataLine = System.Text.RegularExpressions.Regex.Replace(dataLine.Trim(), @" +", " ");
                //dataLine = dataLine.Trim().Replace(" ", ",");
                string[] parts = dataLine.Split(' ');
                if (rows <= Num_Line)//读线路参数
                {
                    LLine[j] = new Line();
                    LLine[j].Node_No[0] = Int16.Parse(parts[0]);
                    LLine[j].Node_No[1] = Int16.Parse(parts[1]);
                    LLine[j].Flag = byte.Parse(parts[2]);
                    for (short n = 0; n < 3; n++)
                        LLine[j].RXBK[n] = float.Parse(parts[n + 3]);
                    j++;
                }
                else if ((rows > Num_Line) && (rows <= Num_Line + Num_Gen))//读发电机参数
                {
                    GGen[p] = new Generator();
                    GGen[p].Node_No = Int16.Parse(parts[0]);
                    GGen[p].Flag = byte.Parse(parts[1]);
                    GGen[p].PQV[0] = float.Parse(parts[2]);
                    GGen[p].PQV[1] = float.Parse(parts[3]);
                    p++;
                }
                else if ((rows > Num_Line + Num_Gen) && (rows <= Num_Gen + Num_Line + Num_Load))//读负荷参数
                {
                    LLoad[q] = new Load();
                    LLoad[q].Node_No = short.Parse(parts[0]);
                    LLoad[q].Flag = byte.Parse(parts[1]);
                    for (short n = 0; n < 6; n++)
                    {
                        LLoad[q].ABC[n] = float.Parse(parts[n + 2]);
                    }
                    q++;
                }
            }

            TemporaryCulture.Stop();
            stream.Close();
            input.Close();
        }

        //序号处理子程
        void Node_Sequen(out short Num_Node, int Num_Line, int Num_Gen, int Num_Load,
                         out short Num_Swing, out short Num_GPV, out short Num_GPQ)
        {
            Num_Node = 0;
            Num_Swing = 0;
            Num_GPV = 0;
            Num_GPQ = 0;
            //bool sign = true;
            short i, j, Flag, temp, np;
            short[,] Node_Name = new short[NODEMAX, 2];			//0-节点名(号);1-节点出线数 

            //统计各节点的出线数
            //for (i = 0; i < NODEMAX; i++) Node_Name[i, 1] = 0; //节点出线数初始化为0
            Array.Clear(Node_Name, 0, Node_Name.Length);//节点出线数初始化为0
            for (i = 0; i < Num_Line; i++)
            {
                if (LLine[i].Node_No[0] == LLine[i].Node_No[1])
                    continue;		//接地支路(左右节点相同)不在出线统计之内
                Flag = 0;							//左节点出线数分析开始
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[0] == Node_Name[j, 0])//该节点已经在节点
                    {										//数组中出现,只需
                        Node_Name[j, 1]++;					//出线数加1。
                        Flag = 1;
                    };
                    if (Flag == 1) break;
                }
                if (Flag == 0)										//该节点还没在
                {												//节点数组中出
                    Node_Name[Num_Node, 0] = LLine[i].Node_No[0];	//现,需将该节
                    Node_Name[Num_Node, 1]++;					//点名(号)添加
                    Num_Node++;									//到节点数组中,
                    if (Num_Node > NODEMAX)						//然后该节点的
                    {											//出线数加1,并将节点数也加1。
                        Console.WriteLine("Nodes Number > " + NODEMAX + "!");                     
                        Environment.Exit(1);
                    }
                }
                Flag = 0;							//右节点出线数分析开始
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[1] == Node_Name[j, 0])//该节点已经在节点	
                    {										//数组中出现,只需
                        Node_Name[j, 1]++;					//出线数加1
                        Flag = 1;
                    };
                    if (Flag == 1) break;
                }
                if (Flag == 0)										//该节点还没在
                {												//节点数组中出
                    Node_Name[Num_Node, 0] = LLine[i].Node_No[1];	//现,需将该节
                    Node_Name[Num_Node, 1]++;					//点名(号)添加
                    Num_Node++;									//到节点数组中,
                    if (Num_Node > NODEMAX)						//然后该节点的
                    {											//出线数加1,并
                                                                //将节点数也加
                        Console.WriteLine("Node Numbers > " + NODEMAX + "!");                     
                        Environment.Exit(1);
                    }
                }
            }
            //节点出线数统计完毕
            //根据出线数由小到大的顺序对节点进行排序(冒泡算法)
            for (i = 0; i < Num_Node - 1; i++)
            {
                np = i;
                for (short m = (short)(i + 1); m < Num_Node; m++)
                    if (Node_Name[np, 1] > Node_Name[m, 1]) np = m;
                temp = Node_Name[i, 0];
                Node_Name[i, 0] = Node_Name[np, 0];
                Node_Name[np, 0] = temp;
                temp = Node_Name[i, 1];
                Node_Name[i, 1] = Node_Name[np, 1];
                Node_Name[np, 1] = temp;
            }
            //平衡节点统计:总数及各节点的名(号)
            short[] Node_Name_Swing = new short[SWINGMAX];
            for (i = 0; i < Num_Gen; i++)
            {
                if (GGen[i].Flag == 0)
                {
                    Node_Name_Swing[Num_Swing] = GGen[i].Node_No;
                    Num_Swing++;
                    if (Num_Swing > SWINGMAX)
                    {
                        Console.WriteLine("Swinging Generators Number > " + "SWINGMAX!");                      
                        Environment.Exit(1);
                    }
                }
            }
            //根据出线数由小到大的顺序对节点排序,并将平衡节点排在最后(序号最大)
            int Nswing = 0, Nnode = 0;
            for (i = 0; i < Num_Node; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Swing; j++)
                {
                    if (Node_Name[i, 0] == Node_Name_Swing[j]) Flag = 1;
                    if (Flag == 1) break;		//Flag=1时,表示该节点为平衡节点,
                }							//需排在靠后的位置上。
                if (Flag == 0)
                {
                    Node_Name_NewtoOld[Nnode] = Node_Name[i, 0];
                    Nnode++;
                }
                else	//最后的各平衡节点间也仍然按出线数由小到大的顺序排列
                {
                    Node_Name_NewtoOld[Num_Node - Num_Swing + Nswing] =
                        Node_Name[i, 0];
                    Nswing++;
                }
            }
            //新线路类型标志赋初值
            for (i = 0; i < Num_Line; i++) Line_Flag[i] = LLine[i].Flag;
            //线路名(号)处理:变成新的节点名(号)且左节点的绝对值小于右节点的绝对值
            for (i = 0; i < Num_Line; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[0] == Node_Name_NewtoOld[j])//左节点处理
                    {
                        Line_NodeName[i, 0] = j;	//赋新名(号)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[1] == Node_Name_NewtoOld[j])//右节点处理		
                    {
                        Line_NodeName[i, 1] = j;	//赋新名(号)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
                if (Line_NodeName[i, 0] > Line_NodeName[i, 1])//左节点的绝对值小于
                {										   //右节点的绝对值处理
                    if (LLine[i].Flag == 1) Line_Flag[i] = 2;		//变压器的非标准
                    if (LLine[i].Flag == 2) Line_Flag[i] = 1;		//变比侧发生变化
                    temp = Line_NodeName[i, 0];
                    Line_NodeName[i, 0] = Line_NodeName[i, 1];
                    Line_NodeName[i, 1] = temp;
                }
            }
            //线路排序:按照左节点的绝对值由小到大、若左节点的绝对值相等则按照右节
            //点的绝对值由小到大顺序排序（双排序冒泡算法）
            for (i = 0; i < Num_Line; i++) Line_No_NewtoOld[i] = i;
            for (i = 0; i < Num_Line - 1; i++)
            {
                np = i;
                for (j = (short)(i + 1); j < Num_Line; j++)
                {
                    if (Line_NodeName[j, 0] < Line_NodeName[np, 0]
                        || (Line_NodeName[j, 0] == Line_NodeName[np, 0]
                        && Line_NodeName[j, 1] < Line_NodeName[np, 1]))
                    {
                        np = j;
                    }
                }
                temp = Line_NodeName[np, 0];
                Line_NodeName[np, 0] = Line_NodeName[i, 0];
                Line_NodeName[i, 0] = temp;
                temp = Line_NodeName[np, 1];
                Line_NodeName[np, 1] = Line_NodeName[i, 1];
                Line_NodeName[i, 1] = temp;
                temp = Line_No_NewtoOld[np];
                Line_No_NewtoOld[np] = Line_No_NewtoOld[i];
                Line_No_NewtoOld[i] = temp;
                temp = Line_Flag[np];
                Line_Flag[np] = Line_Flag[i];
                Line_Flag[i] = temp;
            }
            //发电机节点名(号)处理:变成新的节点名(号)
            for (i = 0; i < Num_Gen; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (GGen[i].Node_No == Node_Name_NewtoOld[j])
                    {
                        Gen_NodeName[i] = j;						//赋新名(号)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
            }
            //发电机排序：按照新节点名（号）由小到大的顺序排序，并找出新发电机序号
            //对应的旧发电机序号
            for (i = 0; i < Num_Gen; i++) Gen_No_NewtoOld[i] = i;
            for (i = 0; i < Num_Gen - 1; i++)
            {
                np = i;
                for (j = (short)(i + 1); j < Num_Gen; j++)
                {
                    if (Gen_NodeName[j] < Gen_NodeName[np])
                    {
                        np = j;
                    }
                }
                temp = Gen_NodeName[np];
                Gen_NodeName[np] = Gen_NodeName[i];
                Gen_NodeName[i] = temp;
                temp = Gen_No_NewtoOld[np];
                Gen_No_NewtoOld[np] = Gen_No_NewtoOld[i];
                Gen_No_NewtoOld[i] = temp;
            }
            //负荷节点名(号)处理:变成新的节点名(号)
            for (i = 0; i < Num_Load; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLoad[i].Node_No == Node_Name_NewtoOld[j])
                    {
                        Load_NodeName[i] = j;						//赋新名(号)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
            }
            //负荷排序：按照新节点名（号）由小到大的顺序排序，并找出新负荷序号
            //对应的旧负荷序号
            for (i = 0; i < Num_Load; i++) Load_No_NewtoOld[i] = i;
            for (i = 0; i < Num_Load - 1; i++)
            {
                np = i;
                for (j = (short)(i + 1); j < Num_Load; j++)
                {
                    if (Load_NodeName[j] < Load_NodeName[np])
                    {
                        np = j;
                    }
                }
                temp = Load_NodeName[np];
                Load_NodeName[np] = Load_NodeName[i];
                Load_NodeName[i] = temp;
                temp = Load_No_NewtoOld[np];
                Load_No_NewtoOld[np] = Load_No_NewtoOld[i];
                Load_No_NewtoOld[i] = temp;
            }
            //从发电机节点数据中归纳出平衡节点、PV节点、PQ节点的新节点名（号）和对
            //应的旧发电机序号，并对平衡节点和PV节点修正其节点类型标志
            for (i = 0; i < Num_Node; i++) Node_Flag[i] = 1;	//节点类型赋初值1(PQ节点)
            Nswing = 0;
            for (i = 0; i < Num_Gen; i++)
            {
                j = Gen_No_NewtoOld[i];					//发电机节点旧顺序号
                if (GGen[j].Flag == 0)
                {
                    Gen_SWNode[Nswing, 0] = Gen_NodeName[i];	//发电机节点名称
                    Gen_SWNode[Nswing, 1] = j;
                    Node_Flag[Gen_NodeName[i]] = 0;
                    Nswing++;
                }
                else if (GGen[j].Flag == 1)
                {
                    Gen_PQNode[Num_GPQ, 0] = Gen_NodeName[i];	//发电机节点名称 新旧对应
                    Gen_PQNode[Num_GPQ, 1] = j;
                    (Num_GPQ)++;
                }
                else if (GGen[j].Flag == 2)
                {
                    Gen_PVNode[Num_GPV, 0] = Gen_NodeName[i];	//发电机节点名称
                    Gen_PVNode[Num_GPV, 1] = j;
                    Node_Flag[Gen_NodeName[i]] = 2;
                    (Num_GPV)++;
                    if (Num_GPV > PVMAX)
                    {
                        Console.WriteLine("PV Generators Number > " + PVMAX + "!");
                        Environment.Exit(1);
                    }
                }
            }
        }


        double[,] Y_Diag = new double[NODEMAX, 2];				//节点导纳阵的对角元:0-实部;
        //1-虚部。
        double[,] Y_UpTri = new double[NODEMAX * NODEFACTOR, 2];	//节点导纳阵上三角的非零元:
        //0-实部;1-虚部。
        int[] Foot_Y_UpTri = new int[NODEMAX * NODEFACTOR];	//上三角按行压缩存储的非零元的
        //列足码。
        int[] Num_Y_UpTri = new int[NODEMAX];				//上三角各行非零元素的个数
        int[] No_First_Y_UpTri = new int[NODEMAX];			//上三角各行第一个非零元素在
        //Y_UpTri中的顺序号。
        int[] Foot_Y_DownTri = new int[NODEMAX * NODEFACTOR];	//下三角按行压缩存储的非零元的
        //列足码。
        int[] Num_Y_DownTri = new int[NODEMAX];				//下三角各行非零元素的个数
        int[] No_First_Y_DownTri = new int[NODEMAX];		//下三角各行第一个非零元素在按
        //行压缩存储序列中的顺序号
        int[] No_Y_DownTri_RowtoCol = new int[NODEMAX * NODEFACTOR];	//下三角某行非零元所对
        //应的按列压缩存储序列
        //中的序号

        //形成节点导纳矩阵1(不包括线路充电容纳及非标准变比的影响)子程
        void Y_Bus1(int Num_Node, int Num_Line, int Num_Swing)
        {
            int i, j, k, k_old, Flag, l;
            double X, B;							//线路参数工作单元
            l = 0;
            //for (i = 0; i < Num_Node - Num_Swing; i++)	//初始化
            //{
            //    Y_Diag[i, 1] = 0.0;
            //    Num_Y_UpTri[i] = 0;
            //}
            Array.Clear(Y_Diag, 0, 2 * (Num_Node - Num_Swing));//
            Array.Clear(Num_Y_UpTri, 0, Num_Node - Num_Swing);
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0];		//线路左节点
                j = Line_NodeName[k, 1];		//线路右节点
                if (i >= Num_Node - Num_Swing)	//左右节点均为平衡节点，对导纳阵无
                    break;					//影响。
                k_old = Line_No_NewtoOld[k];	//对应的旧线路顺序号
                X = LLine[k_old].RXBK[1];		//取线路电抗值
                B = -1.0 / X;					//不计线路电阻后的线路支路电纳
                if (j >= Num_Node - Num_Swing)	//左为普通节点，右为平衡节点
                    Y_Diag[i, 1] = Y_Diag[i, 1] + B;//只计及左节点（普通节点）的自电纳，不计互电纳
                else						//左、右节点均为普通节点
                {
                    Flag = 0;
                    if (k > 0 && (i == Line_NodeName[k - 1, 0])
                        && (j == Line_NodeName[k - 1, 1])) Flag = 1;	//多回线
                    Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                    if (i != j)								//非接地支路
                    {
                        Y_Diag[j, 1] = Y_Diag[j, 1] + B;
                        if (Flag == 0)							//第一回线
                        {
                            Y_UpTri[l, 1] = -B;
                            Foot_Y_UpTri[l] = j;
                            Num_Y_UpTri[i]++;
                            l++;
                            if (l > NODEMAX * NODEFACTOR)
                            {
                                Console.WriteLine("Number of none-zero elements of up_triangle > " + 
                                    NODEMAX * NODEFACTOR + "!");                              
                                Environment.Exit(1);
                            }
                        }
                        else										//多回线
                        {
                            Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - B;
                        }
                    }
                }
            }

            No_First_Y_UpTri[0] = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
                No_First_Y_UpTri[i + 1] = No_First_Y_UpTri[i] + Num_Y_UpTri[i];
            //导纳矩阵下三角按行压缩存储时各行非零元的个数、每行第一个非零元的序
            //号、按行压缩存储时非零元的列足码、某一非零元所对应的下三角阵按列压
            //缩存储时相同非零元的序号，上述这些值的求取目的在于快速地处理下面迭
            //代过程中修正方程Jacobi矩阵的按行压缩形式下的取值、消去和回代运算。
            int[] Row_Down = new int[NODEMAX];		//下三角某列非零元序号的下限工作单元
            int[] Row_Up = new int[NODEMAX];		//下三角某列非零元序号的上限工作单元
            int li;
            //for (i = 0; i < Num_Node - Num_Swing; i++)//下三角各行非零元个数数组清零
            //    Num_Y_DownTri[i] = 0;
            Array.Clear(Num_Y_DownTri, 0, Num_Node - Num_Swing);//
            for (j = 0; j < Num_Node - Num_Swing; j++)//该循环统计下三角各行非零元个数
            {
                for (k = No_First_Y_UpTri[j]; k < No_First_Y_UpTri[j + 1]; k++)
                {								//针对下三角第j列非零元作处理
                    i = Foot_Y_UpTri[k];			//行足码
                    Num_Y_DownTri[i]++;			//下三角第i行非零元个数增1
                }
                Row_Down[j] = No_First_Y_UpTri[j];
                Row_Up[j] = No_First_Y_UpTri[j + 1];
            }
            No_First_Y_DownTri[0] = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)	//下三角各行第一个非零元的序号
                No_First_Y_DownTri[i + 1] = No_First_Y_DownTri[i] + Num_Y_DownTri[i];
            for (i = 1; i < Num_Node - Num_Swing; i++)	//该循环确定下三角各行非零元的
            {									//列足码。
                li = No_First_Y_DownTri[i];		//下三角第i行第一个非零元序号
                for (j = 0; j < i; j++)				//该循环搜寻下三角第0~i-1列中
                {								//行号为i的非零元。

                    if ((k = Row_Down[j]) < Row_Up[j])//如果上三角第j行有非零的非对角元素。这段代码用以替换下面一段代码
                    {
                        if (Foot_Y_UpTri[k] == i)//表示下三角第i行有非零的非对角元素
                        {
                            Foot_Y_DownTri[li] = j;//记录i行第li个非零元的列足码
                            No_Y_DownTri_RowtoCol[li] = k;//记录该元素在下三角按
                            //列压缩存储序列中序号
                            li++;				//序号计数器增1，备下次使用
                            Row_Down[j]++;
                        }
                    }
                    //以下代码不合理——修改
                    //for(k=Row_Down[j];k<Row_Up[j];k++)
                    //{
                    //    if(Foot_Y_UpTri[k]==i)
                    //    {
                    //        Foot_Y_DownTri[li]=j;//记录i行第li个非零元的列足码
                    //        No_Y_DownTri_RowtoCol[li]=k;//记录该元素在下三角按
                    //                                    //列压缩存储序列中序号
                    //        li++;				//序号计数器增1，备下次使用
                    //        Row_Down[j]++;
                    //    }
                    //    break;
                    //}
                }
            }
        }


        //形成节点导纳矩阵2(包括线路充电容纳及非标准变比的影响)子程
        void Y_Bus2(int Num_Node, int Num_Line, int Num_Load, int Num_Swing)
        {
            int i, j, k, k_old, Flag, l;
            double R, X, Z, G, B, BK;

            l = 0;
            //for (i = 0; i < Num_Node - Num_Swing; i++)	//初始化
            //{
            //    Y_Diag[i, 0] = 0.0;
            //    Y_Diag[i, 1] = 0.0;
            //}
            Array.Clear(Y_Diag, 0, 2 * (Num_Node - Num_Swing));//
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0];		//线路左节点
                j = Line_NodeName[k, 1];		//线路右节点
                if (i >= Num_Node - Num_Swing)	//左右节点均为平衡节点，对导纳阵无
                    break;					//影响。
                k_old = Line_No_NewtoOld[k];	//对应的旧线路顺序号
                R = LLine[k_old].RXBK[0];		//取线路电阻值
                X = LLine[k_old].RXBK[1];		//取线路电抗值
                BK = LLine[k_old].RXBK[2];	//取线路容纳半值或变压器变比值
                Z = R * R + X * X;
                G = R / Z;						//电导
                B = -X / Z;						//电纳
                if (j >= Num_Node - Num_Swing)	//左为普通节点,右为平衡节点
                {
                    if (Line_Flag[k] == 0)					//普通支路
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                    }
                    else if (Line_Flag[k] == 1)			//非标准变比在左侧节点
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + 1.0 / BK / BK * G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + 1.0 / BK / BK * B;
                    }
                    else if (Line_Flag[k] == 2)			//非标准变比在右侧节点
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                    }
                }
                else								//左、右节点均为普通节点
                {
                    Flag = 0;
                    if (k > 0 && (i == Line_NodeName[k - 1, 0])
                        && (j == Line_NodeName[k - 1, 1])) Flag = 1;	//多回线
                    if (i == j)				//接地支路(变压器支路不直接接地)
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                    }
                    else										//非接地支路
                    {
                        if (Line_Flag[k] == 0)						//普通支路
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + B + BK;
                            if (Flag == 0)							//第一回线
                            {
                                Y_UpTri[l, 0] = -G;
                                Y_UpTri[l, 1] = -B;
                                l++;
                            }
                            else								//多回线
                            {
                                Y_UpTri[l, 0] = Y_UpTri[l, 0] - G;
                                Y_UpTri[l, 1] = Y_UpTri[l, 1] - B;
                            }
                        }
                        else if (Line_Flag[k] == 1)		//非标准变比在左侧节点
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + 1.0 / BK / BK * G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + 1.0 / BK / BK * B;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + B;
                            if (Flag == 0)							//第一回线
                            {
                                Y_UpTri[l, 0] = -1.0 / BK * G;
                                Y_UpTri[l, 1] = -1.0 / BK * B;
                                l++;
                            }
                            else								//多回线
                            {
                                Y_UpTri[l - 1, 0] = Y_UpTri[l - 1, 0] - 1.0 / BK * G;
                                Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - 1.0 / BK * B;
                            }
                        }
                        else							//非标准变比在右侧节点
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + 1.0 / BK / BK * G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + 1.0 / BK / BK * B;
                            if (Flag == 0)							//第一回线
                            {
                                Y_UpTri[l, 0] = -1.0 / BK * G;
                                Y_UpTri[l, 1] = -1.0 / BK * B;
                                l++;
                            }
                            else								//多回线
                            {
                                Y_UpTri[l - 1, 0] = Y_UpTri[l - 1, 0] - 1.0 / BK * G;
                                Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - 1.0 / BK * B;
                            }
                        }
                    }
                }
            }

            //将负荷静特性中阻抗成份计入导纳矩阵的对角元中
            for (i = 0; i < Num_Load; i++)
            {
                k = Load_No_NewtoOld[i];
                if (LLoad[k].Flag == 1)
                {
                    j = Load_NodeName[i];
                    if (j < Num_Node - Num_Swing)
                    {
                        Y_Diag[j, 0] = Y_Diag[j, 0] + LLoad[k].ABC[0];
                        Y_Diag[j, 1] = Y_Diag[j, 1] - LLoad[k].ABC[1];
                    }
                }
            }
        }


        //复数A和B相乘得C（直角坐标形式）
        void Comp_Mul(out double C0, out double C1,
            double A1, double A2, double B1, double B2)
        {

            //C=new double[2];A=new double[2];B=new double[2];
            C0 = A1 * B1 - A2 * B2; C1 = A1 * B2 + A2 * B1;
            //C[0]=A[0]*B[0]-A[1]*B[1];	C[1]=A[0]*B[1]+A[1]*B[0];
        }


        //复数A和B相除得C（直角坐标形式）
        void Comp_Div(double[] C, double[] A, double[] B)
        {
            //C=new double[2];A=new double[2];B=new double[2];
            double tt;
            double[] t = new double[2];
            tt = B[0] * B[0] + B[1] * B[1];
            if (tt == 0.0)
            {
                Console.WriteLine("Divided by zero!!");
                Environment.Exit(1);
            }
            else
            {
                t[0] = B[0];
                t[1] = -B[1];
                Comp_Mul(out C[0], out C[1], A[0], A[1], t[0], t[1]);
                C[0] = C[0] / tt;
                C[1] = C[1] / tt;
            }
        }


        double[,] Fact_Diag = new double[NODEMAX, 2];				//因子表对角元素:0-有功因
        //子表;1-无功因子表。
        double[,] Fact_UpTri = new double[NODEMAX * NODEFACTOR, 2];	//因子表上三角非零元素:0-
        //有功因子表;1-无功因子表。
        int[,] Foot_Fact_UpTri = new int[NODEMAX * NODEFACTOR, 2];	//因子表上三角非零元列足码
        int[,] Num_Fact_UpTri = new int[NODEMAX, 2];				//因子表上三角各行非零非对
        //角元的个数。
        //形成节点导纳矩阵因子表
        void Factor_Table(int Num_Node, int Num_Swing, int Num_GPV, int IterFlag)
        {
            int i;							//因子表正在形成的所在行号
            int im;							//因子表正在消去的所在行号
            int j;							//列足码暂存单元
            int k;							//临时记数单元
            int ix;							//因子表上三角元素地址（序号）计数
            double[] Y_Work = new double[NODEMAX];			//工作数组
            double Temp1, Temp2;				//临时工作单元
            int kgpv;

            for (i = 0; i < Num_Node - Num_Swing; i++)//遍历所有行，形成完整的因子表，整个计算过程以行号i为循环变量
            {
                if (IterFlag == 1 && Node_Flag[i] == 2)		//无功迭代对应的PV节点
                {
                    Num_Fact_UpTri[i, IterFlag] = 0;
                    Fact_Diag[i, IterFlag] = 0.0;
                }
                else
                {
                    //for (k = i + 1; k < Num_Node - Num_Swing; k++)
                    //    Y_Work[k] = 0.0;
                    Array.Clear(Y_Work, i + 1, Num_Node - Num_Swing - i - 1);//
                    Y_Work[i] = Y_Diag[i, 1];

                    for (k = No_First_Y_UpTri[i]; k < No_First_Y_UpTri[i + 1]; k++)//上三角各行第一个非零元素在
                                                                                   //Y_UpTri中的顺序号。
                    {
                        j = Foot_Y_UpTri[k];//上三角按行压缩存储的非零元的
                                            //列足码。
                        Y_Work[j] = Y_UpTri[k, 1];//获取第i行的数据
                    }
                    if (IterFlag == 1)
                    {
                        for (kgpv = 0; kgpv < Num_GPV; kgpv++)
                        {
                            j = Gen_PVNode[kgpv, 0];			//PV节点号
                            Y_Work[j] = 0.0;
                        }
                    }

                    ix = 0;
                    for (im = 0; im < i; im++)//在形成因子表第i行各元素时，工作数组应该与i-1行以前已形成的各行因子表元素进行消去运算
                    {
                        for (k = 0; k < Num_Fact_UpTri[im, IterFlag]; k++)//因子表上三角各行非零非对角元的个数。
                        {
                            if (Foot_Fact_UpTri[ix, IterFlag] != i)/*当列足码等于i时，要执行消去运算。
                                                                   在一般情况下，当工作数组中待消行行号为i，而第im（im<i）行
                                                                    中没有足码为i的元素时（这就表明Uim,i=0），即可省去im行
                                                                    的消去运算过程*/
                                ix++;
                            else
                            {
                                Temp1 = Fact_UpTri[ix, IterFlag]
                                    / Fact_Diag[im, IterFlag];
                                for (short m =(short)k; m < Num_Fact_UpTri[im, IterFlag]; m++)//
                                { 
                                    j = Foot_Fact_UpTri[ix, IterFlag];
                                    Temp2 = Temp1 * Fact_UpTri[ix, IterFlag];
                                    Y_Work[j] = Y_Work[j] - Temp2;
                                    ix++;
                                }
                            }
                        }
                    }
                    Fact_Diag[i, IterFlag] = 1.0 / Y_Work[i];
                    Temp1 = Fact_Diag[i, IterFlag];
                    k = 0;
                    for (j = i + 1; j < Num_Node - Num_Swing; j++)//只形成上三角
                    {
                        if (Math.Abs(Y_Work[j]) != 0.0)
                        {
                            Fact_UpTri[ix, IterFlag] = Y_Work[j] * Temp1;//上三角阵元素，压缩存放
                            Foot_Fact_UpTri[ix, IterFlag] = j;//上三角阵元素压缩存放后的列足码
                            k++;
                            ix++;
                        }
                    }
                    Num_Fact_UpTri[i, IterFlag] = k;//因子表上三角各行非零非对角元的个数。  
                }
            }
        }

        //方程AX=t求解
        void Equ_Calculation(int Num_Node, int Num_Swing,
                             double[] Power_Dis_Correct, int IterFlag)
        {
            //Power_Dis_Correct=new double[NODEMAX];
            int i, j, k, ix;							//参见Factor_Table子程说明
            double Temp1, Temp2;						//临时工作单元

            ix = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)		//前代运算开始
            {
                Temp1 = Power_Dis_Correct[i];			//将右端量送入临时工作单元
                for (k = 0; k < Num_Fact_UpTri[i, IterFlag]; k++)
                {
                    j = Foot_Fact_UpTri[ix, IterFlag];
                    Temp2 = Temp1 * Fact_UpTri[ix, IterFlag];
                    Power_Dis_Correct[j] = Power_Dis_Correct[j] - Temp2;
                    ix++;
                }
                Power_Dis_Correct[i] = Temp1 * Fact_Diag[i, IterFlag];
            }
            for (i = Num_Node - Num_Swing - 1; i >= 0; i--)	//回代运算开始
            {
                Temp1 = Power_Dis_Correct[i];
                for (k = 0; k < Num_Fact_UpTri[i, IterFlag]; k++)
                {
                    ix--;
                    j = Foot_Fact_UpTri[ix, IterFlag];
                    Temp2 = Power_Dis_Correct[j] * Fact_UpTri[ix, IterFlag];
                    Temp1 = Temp1 - Temp2;
                }
                Power_Dis_Correct[i] = Temp1;				//最终得未知数的解
            }
        }


        double[,] Voltage = new double[NODEMAX, 2];						//节点电压:0-相位角;
        //1-幅值。
        double[,] Current_Const = new double[SWINGMAX * NODEFACTOR, 2];	//常数电流:0-实部;
        //1-虚部。
        int[] Node_Name_Curr_Const = new int[SWINGMAX * NODEFACTOR];	//常数电流的节点名(号)
        double[,] Power_Const = new double[NODEMAX, 2];					//各节点注入功率不变部
        //分:0-实部;1-虚部。

        void Voltage_Initial(int Num_Node, int Num_Swing,string path)
        {
            TemporaryCulture.Start();
            FileStream stream = File.OpenRead(path);
            StreamReader input = new StreamReader(stream);
            string dataLine = "";
            ushort index = 0;
            while ((dataLine = input.ReadLine()) != null)
            {

                dataLine = System.Text.RegularExpressions.Regex.Replace(dataLine.Trim(), @" +", " ");
                if (dataLine != "")
                {
                    string[] parts = dataLine.Split(' ');
                    Voltage[index, 0] = double.Parse(parts[0]);
                    Voltage[index, 1] = double.Parse(parts[1]);
                    index++;
                }
            }
            if (index != Num_Node - Num_Swing)
            {
                Console.WriteLine("Lines of voltage initial data do not match the number:" + 
                    (Num_Node - Num_Swing).ToString() + 
                    "\r\nplease check the voltage initial data!\r\nSmooth start will be used instead!");
                for (short i = 0; i < Num_Node - Num_Swing; i++)
                {
                    Voltage[i, 0] = 0.0;				//电压相位的初值：0.0(弧度)
                    Voltage[i, 1] = 1.0;				//电压幅值的初值：1.0
                }
            }
            TemporaryCulture.Stop();
            stream.Close();
            input.Close();
        }


        //初始化子程
        void Initial(int Num_Node, int Num_Line, int Num_Load, int Num_Swing,
                     int Num_GPV, int Num_GPQ, out int Num_Cur_Const,
                     double[,] DVolt)
        {
            Num_Cur_Const = 0;
            int i, j, jl, jr, k, kk;
            int Flag;
            int kg_old, kl_old;					//发电机、负荷旧顺序号工作单元
            int kl;								//负荷计数临时工作单元
            int kgpv;							//发电机PV节点计数临时工作单元
            int kgpq;							//发电机PQ节点计数临时工作单元
            double R, X, Z, Ang;					//线路参数及平衡节点电压相角临
            //时单元。
            double[] yij, V_Temp, I_Temp;	//临时工作单元
            yij = new double[2];
            V_Temp = new double[2];
            I_Temp = new double[2];
            //for (i = 0; i < Num_Node - Num_Swing; i++)
            //{
            //    DVolt[i, 0] = 0.0;				//电压相量变化量赋初值0.0
            //    DVolt[i, 1] = 0.0;
            //    Power_Const[i, 0] = 0.0;			//各节点注入功率赋初值
            //    Power_Const[i, 1] = 0.0;
            //}
            Array.Clear(DVolt, 0, 2 * (Num_Node - Num_Swing));//
            Array.Clear(Power_Const, 0, 2 * (Num_Node - Num_Swing));//
            //else if(VolIni_Flag==1)
            //if (VolIni_Flag == 1) Voltage_Initial(Num_Node, Num_Swing);
            for (kgpv = 0; kgpv < Num_GPV; kgpv++)		//发电机PV节点的电压幅值=VG
            {
                i = Gen_PVNode[kgpv, 0];
                kg_old = Gen_PVNode[kgpv, 1];
                Voltage[i, 1] = GGen[kg_old].PQV[1];
            }

            for (i = 0; i < Num_Line; i++)				//常数电流信息求解
            {
                jl = Line_NodeName[i, 0];			//线路左节点
                jr = Line_NodeName[i, 1];			//线路右节点
                if (jl < Num_Node - Num_Swing
                    && jr >= Num_Node - Num_Swing)	//jl为普通节点,jr为平衡节点。
                {
                    k = Line_No_NewtoOld[i];		//对应的旧线路号
                    R = LLine[k].RXBK[0];
                    X = LLine[k].RXBK[1];
                    Z = R * R + X * X;
                    if (Line_Flag[i] == 0)
                    {
                        R = R / Z;
                        X = -X / Z;
                    }
                    else
                    {
                        R = R / Z / LLine[k].RXBK[2];
                        X = -X / Z / LLine[k].RXBK[2];
                    }
                    yij[0] = R;
                    yij[1] = X;				//至此求得支路导纳除以非标准变比

                    Flag = 0;
                    for (k = 0; k < Num_Swing; k++)
                    {
                        if (Gen_SWNode[k, 0] == jr)
                        {
                            kk = Gen_SWNode[k, 1];
                            Ang = GGen[kk].PQV[1] * Deg_to_Rad;
                            V_Temp[0] = GGen[kk].PQV[0] * Math.Cos(Ang);
                            V_Temp[1] = GGen[kk].PQV[0] * Math.Sin(Ang);
                            Flag = 1;
                        }
                        if (Flag == 1) break;
                    }						//至此求得相应平衡节点电压的实虚部

                    Flag = 0;
                    for (j = 0; j < Num_Cur_Const; j++)
                    {
                        if (Node_Name_Curr_Const[j] == jl) Flag = 1;
                        if (Flag == 1) break;
                    }
                    if (Flag == 0)					//新增常数电流节点
                    {
                        Node_Name_Curr_Const[Num_Cur_Const] = jl;

                        Comp_Mul(out Current_Const[Num_Cur_Const, 0], out Current_Const[Num_Cur_Const, 1], 
                            yij[0], yij[1], V_Temp[0], V_Temp[1]);
                        Num_Cur_Const++;
                        if (Num_Cur_Const > SWINGMAX * NODEFACTOR)
                        {
                            Console.WriteLine("Number of constant-current nodes > " + SWINGMAX * NODEFACTOR + "!");
                        }
                    }
                    else						//该常数电流节点已经出现过
                    {
                        Comp_Mul(out I_Temp[0], out I_Temp[1], yij[0], yij[1], V_Temp[0], V_Temp[1]);
                        Current_Const[j, 0] = Current_Const[j, 0] + I_Temp[0];
                        Current_Const[j, 1] = Current_Const[j, 1] + I_Temp[1];
                    }
                }
            }
            //输出常数电流节点数据

            //各节点注入功率不变部分求值处理
            kgpv = 0;
            kgpq = 0;
            kl = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
            {
                if (kgpv < Num_GPV && i == Gen_PVNode[kgpv, 0])	//发电机PV节点
                {
                    kg_old = Gen_PVNode[kgpv, 1];				//发电机旧顺序号
                    if (kl < Num_Load && i == Load_NodeName[kl])	//负荷节点
                    {
                        kl_old = Load_No_NewtoOld[kl];		//负荷旧顺序号
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0]
                            - LLoad[kl_old].ABC[4];			//有功部分加PG和C1
                        kl++;
                    }
                    else									//非负荷节点
                    {
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0];			//有功部分加入PG
                    }
                    kgpv++;
                }
                else if (kgpq < Num_GPQ && i == Gen_PQNode[kgpq, 0])	//发电机PQ节点
                {
                    kg_old = Gen_PQNode[kgpq, 1];				//发电机旧顺序号
                    if (kl < Num_Load && i == Load_NodeName[kl])	//负荷节点
                    {
                        kl_old = Load_No_NewtoOld[kl];		//负荷旧顺序号
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0]
                            - LLoad[kl_old].ABC[4];			//有功部分加PG和C1
                        Power_Const[i, 1] = Power_Const[i, 1]
                            + GGen[kg_old].PQV[1]
                            - LLoad[kl_old].ABC[5];			//无功部分加QG和C2
                        kl++;
                    }
                    else									//非负荷节点
                    {
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0];			//有功部分加入PG
                        Power_Const[i, 1] = Power_Const[i, 1]
                            + GGen[kg_old].PQV[1];			//无功部分加入QG
                    }
                    kgpq++;
                }
                else					//既非发电机PV节点，又非发电机PQ节点
                {
                    if (kl < Num_Load && i == Load_NodeName[kl])	//负荷节点
                    {
                        kl_old = Load_No_NewtoOld[kl];		//负荷旧顺序号
                        Power_Const[i, 0] = Power_Const[i, 0]
                            - LLoad[kl_old].ABC[4];			//有功部分加入C1
                        Power_Const[i, 1] = Power_Const[i, 1]
                            - LLoad[kl_old].ABC[5];			//无功部分加入C2
                        kl++;
                    }
                }
            }
            //各节点注入功率不变部分处理结果输出:新节点名(号),有功功率,无功功率

            //功率失配量输出部分的标题输出
            //打开失配量输出磁盘文件
               Console.WriteLine(String.Format("{0,15:D}", "Iterating No")+
               String.Format("{0,15:D}", "P_Dis_Max") + String.Format("{0,6:D}", "Node") +
               String.Format("{0,15:D}", "Q_Dis_Max") + String.Format("{0,6:D}", "Node"));
            Console.WriteLine(String.Format("{0,15:D}", "==============") +
               String.Format("{0,15:D}", "==============") + String.Format("{0,6:D}", "=====") +
               String.Format("{0,15:D}", "==============") + String.Format("{0,6:D}", "====="));
        }

        //求节点功率失配量子程(PV节点的dQi=0)
        void PowerDis_Comp(int Num_Node, int Num_Load, int Num_Swing, int Num_GPV,
                           int Num_Cur_Const, double[,] Power_Dis,
                           double[,] Pij_Sum, double[,] DVolt,
                           int Num_Iteration,
                           out double Power_Dis_Max)
        {
            int i, j, k, li, kl_old, kl, kgpv, ki, k1;
            double V, Ang;					//节点i电压幅值和相位
            double[] VV = new double[2];					//节点i电压实部和虚部
            double[] V_Temp = new double[2];				//节点电压临时工作单元（实、虚部）
            double[] Cur_Count = new double[2];			//注入电流统计单元（实、虚部）
            double[] Cur_Temp = new double[2];				//注入电流临时工作单元（实、虚部）
            double Ix, Iy;					//电流（实、虚部）
            double PP, QQ;					//有功、无功
            int ipmax, iqmax;				//最大有功、无功失配量对应的节点号
            ipmax = 0; iqmax = 0;
            double P_Dis_Max, Q_Dis_Max;		//最大有功、无功失配量

            kl = 0;
            kgpv = 0;
            ki = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
            {
                Power_Dis[i, 0] = Power_Const[i, 0];	//将节点i注入功率不变部分
                Power_Dis[i, 1] = Power_Const[i, 1];	//送入失配量单元。
                Ang = Voltage[i, 0];	 				//上次迭代后节点i电压相位
                V = Voltage[i, 1];					//和幅值。
                VV[0] = V * Math.Cos(Ang);					//上次迭代后节点i电压实部
                VV[1] = V * Math.Sin(Ang);					//和虚部。
                //下面求Pij和Qij及从节点i发出的所有Pij和Qij之和
                Cur_Count[0] = 0.0;
                Cur_Count[1] = 0.0;
                for (k = No_First_Y_DownTri[i];
                    k < No_First_Y_DownTri[i + 1]; k++)	//下三角第i行非零元循环
                {
                    j = Foot_Y_DownTri[k];	//下三角第i行当前非零元的列足码
                    V_Temp[0] = Voltage[j, 1] * Math.Cos(Voltage[j, 0]);
                    V_Temp[1] = Voltage[j, 1] * Math.Sin(Voltage[j, 0]);
                    li = No_Y_DownTri_RowtoCol[k];	//对应的按列压缩存储序列
                    //中同一非零元的序号。

                    Comp_Mul(out Cur_Temp[0], out Cur_Temp[1], 
                        Y_UpTri[li, 0], Y_UpTri[li, 1], V_Temp[0], V_Temp[1]);
                    Cur_Count[0] = Cur_Count[0] + Cur_Temp[0];
                    Cur_Count[1] = Cur_Count[1] + Cur_Temp[1];
                }
                Comp_Mul(out Cur_Temp[0], out Cur_Temp[1],
                    Y_Diag[i, 0], Y_Diag[i, 1], VV[0], VV[1]);
                Cur_Count[0] = Cur_Count[0] + Cur_Temp[0];
                Cur_Count[1] = Cur_Count[1] + Cur_Temp[1];
                for (k = No_First_Y_UpTri[i]; k < No_First_Y_UpTri[i + 1]; k++)
                {
                    j = Foot_Y_UpTri[k];
                    V_Temp[0] = Voltage[j, 1] * Math.Cos(Voltage[j, 0]);
                    V_Temp[1] = Voltage[j, 1] * Math.Sin(Voltage[j, 0]);
                    Comp_Mul(out Cur_Temp[0], out Cur_Temp[1], 
                        Y_UpTri[k, 0], Y_UpTri[k, 1], V_Temp[0], V_Temp[1]);
                    Cur_Count[0] = Cur_Count[0] + Cur_Temp[0];
                    Cur_Count[1] = Cur_Count[1] + Cur_Temp[1];
                }
                Cur_Count[1] = -Cur_Count[1];			//电流取共軛
                Comp_Mul(out Pij_Sum[i, 0], out Pij_Sum[i, 1],
                    VV[0], VV[1], Cur_Count[0], Cur_Count[1]);	//至此,求得从节点i发出的所
                //有Pij和Qij之和。

                if (kgpv < Num_GPV && i == Gen_PVNode[kgpv, 0])	//发电机PV节点
                {
                    if (kl < Num_Load && i == Load_NodeName[kl])	//负荷节点
                    {
                        kl_old = Load_No_NewtoOld[kl];		//负荷旧顺序号
                        if (LLoad[kl_old].Flag == 1)			//计及负荷静特性
                            Power_Dis[i, 0] = Power_Dis[i, 0]
                            - LLoad[kl_old].ABC[2] * V;
                        kl++;
                    }
                    if (ki < Num_Cur_Const && i == Node_Name_Curr_Const[ki])
                    {										//含常数电流节点
                        Power_Dis[i, 0] = Power_Dis[i, 0]
                            + V * (Current_Const[ki, 0] * Math.Cos(Ang)
                            + Current_Const[ki, 1] * Math.Sin(Ang));
                        ki++;
                    }
                    Power_Dis[i, 0] = Power_Dis[i, 0] - Pij_Sum[i, 0];
                    kgpv++;
                }
                else					//PQ节点(包括发电机、负荷及联络节点）
                {
                    if (kl < Num_Load && i == Load_NodeName[kl])	//负荷节点
                    {
                        kl_old = Load_No_NewtoOld[kl];		//负荷旧顺序号
                        if (LLoad[kl_old].Flag == 1)			//计及负荷静特性
                        {
                            Power_Dis[i, 0] = Power_Dis[i, 0]
                                - LLoad[kl_old].ABC[2] * V;
                            Power_Dis[i, 1] = Power_Dis[i, 1]
                                - LLoad[kl_old].ABC[3] * V;
                        }
                        kl++;
                    }
                    if (ki < Num_Cur_Const && i == Node_Name_Curr_Const[ki])
                    {										//含常数电流节点
                        Power_Dis[i, 0] = Power_Dis[i, 0]
                            + V * (Current_Const[ki, 0] * Math.Cos(Ang)
                            + Current_Const[ki, 1] * Math.Sin(Ang));
                        Power_Dis[i, 1] = Power_Dis[i, 1]
                            + V * (Current_Const[ki, 0] * Math.Sin(Ang)
                            - Current_Const[ki, 1] * Math.Cos(Ang));
                        ki++;
                    }
                    Power_Dis[i, 0] = Power_Dis[i, 0] - Pij_Sum[i, 0];
                    Power_Dis[i, 1] = Power_Dis[i, 1] - Pij_Sum[i, 1];
                }
            }

            P_Dis_Max = 0.0;
            Q_Dis_Max = 0.0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
            {
                if (Math.Abs(Power_Dis[i, 0]) > P_Dis_Max)
                {
                    P_Dis_Max = Math.Abs(Power_Dis[i, 0]);
                    ipmax = i;
                }
                if (Math.Abs(Power_Dis[i, 1]) > Q_Dis_Max)
                {
                    Q_Dis_Max = Math.Abs(Power_Dis[i, 1]);
                    iqmax = i;
                }
            }
            //Power_Dis_Max=__max(P_Dis_Max,Q_Dis_Max);
            if (P_Dis_Max > Q_Dis_Max)
                Power_Dis_Max = P_Dis_Max;
            else
                Power_Dis_Max = Q_Dis_Max;

            Console.WriteLine(String.Format("{0,15:D}", Num_Iteration) +
                          String.Format("{0,15:f8}", P_Dis_Max) + String.Format("{0,6:D}", Node_Name_NewtoOld[ipmax]) +
                          String.Format("{0,15:f8}", Q_Dis_Max) + String.Format("{0,6:D}", Node_Name_NewtoOld[iqmax]));
            //节点功率一次偏差修正项
            for (k = 0; k < Num_Cur_Const; k++)			//常数电流修正项处理
            {
                i = Node_Name_Curr_Const[k];			//节点号
                if (i < Num_Node - Num_Swing)
                {
                    Ix = Current_Const[k, 0];
                    Iy = Current_Const[k, 1];
                    Ang = Voltage[i, 0];
                    PP = (-Ix * Math.Sin(Ang) + Iy * Math.Cos(Ang)) * DVolt[i, 0]
                        + (Ix * Math.Cos(Ang) + Iy * Math.Sin(Ang)) * DVolt[i, 1];
                    QQ = (Ix * Math.Cos(Ang) + Iy * Math.Sin(Ang)) * DVolt[i, 0]
                        + (Ix * Math.Sin(Ang) - Iy * Math.Cos(Ang)) * DVolt[i, 1];
                    Power_Dis[i, 0] = Power_Dis[i, 0] - PP * 0.1;
                    if (Node_Flag[i] != 2) Power_Dis[i, 1] = Power_Dis[i, 1] - QQ * 0.1;
                }
            }

            for (i = 0; i < Num_Node - Num_Swing; i++)	//i节点所有出线的Pij、Qij之和
            {									//修正项处理。
                V = Voltage[i, 1];
                PP = Pij_Sum[i, 1] * DVolt[i, 0] - Pij_Sum[i, 0] * DVolt[i, 1];
                QQ = -Pij_Sum[i, 0] * DVolt[i, 0] - Pij_Sum[i, 1] * DVolt[i, 1];
                //		Power_Dis[i,0]=Power_Dis[i,0]-PP/V;
                //		if(Node_Flag[i]!=2)Power_Dis[i,1]=Power_Dis[i,1]-QQ/V;
                Power_Dis[i, 0] = Power_Dis[i, 0] - PP * 0.1;
                if (Node_Flag[i] != 2) Power_Dis[i, 1] = Power_Dis[i, 1] - QQ * 0.1;
            }

            for (k = 0; k < Num_Load; k++)					//负荷静特性之修正项处理
            {
                i = Load_NodeName[k];
                k1 = Load_No_NewtoOld[k];
                if (i < Num_Node - Num_Swing && LLoad[k1].Flag == 1)
                {
                    PP = -LLoad[k1].ABC[2] * DVolt[i, 1];
                    QQ = -LLoad[k1].ABC[3] * DVolt[i, 1];
                    Power_Dis[i, 0] = Power_Dis[i, 0] - PP * 0.5;
                    if (Node_Flag[i] != 2) Power_Dis[i, 1] = Power_Dis[i, 1] - QQ * .5;
                }
            }
        }

        //电压保存(按内部节点号)子程
        void Voltage_Reserve(int Num_Node, int Num_Swing,string fileName)
        {
            sw = new StreamWriter(fileName);
            for (short i = 0; i < Num_Node - Num_Swing; i++)
            {
                sw.WriteLine(String.Format("{0,-12:f6}", Voltage[i, 0]) + 
                    String.Format("{0,-12:f6}", Voltage[i, 1]));
            }
            sw.Close();
        }


        //结果输出子程
        void Result_Output(int Num_Node, int Num_Line, int Num_Gen, int Num_Load,
                           int Num_Swing, int Num_Iteration, double Duration,string fileName)
        {
            int i, j, k, kg_old, kl_old, k_old;
            double BK;					//线路参数临时工作单元
            double[] Z = new double[2];
            double[,] S_Count = new double[NODEMAX, 2];		//节点所有出线功率累加数组:
            //0-实部,1-虚部。
            double[,] S_Line = new double[LINEMAX, 4];		//线路功率:0-Pij,1-Qij,2-Pji,3-Qji
            double[,] S_Node = new double[NODEMAX, 4];		//节点功率:0-PG,1-QG,2-PL,3-QL
            double[,] DS_Node = new double[NODEMAX, 2];		//节点功率失配量:0-有功失配量,
            //1-无功失配量。
            double[] S_T, V_T, I_T, I1_T, I2_T;	//临时工作单元
            S_T = new double[2];
            V_T = new double[2];
            I_T = new double[2];
            I1_T = new double[2];
            I2_T = new double[2];

            double V, t, Angle, Vi, Vj, Angi, Angj;

            //将平衡节点的电压值送入数组Voltage[,]中
            for (i = 0; i < Num_Swing; i++)
            {
                j = Gen_SWNode[i, 0];
                kg_old = Gen_SWNode[i, 1];
                Angle = GGen[kg_old].PQV[1] * Deg_to_Rad;
                Voltage[j, 0] = Angle;
                Voltage[j, 1] = GGen[kg_old].PQV[0];
            }		//至此可以直接使用系统中所有节点电压的相位和幅值

            //初值化处理
            //for (i = 0; i < Num_Node; i++)
            //{
            //    S_Count[i, 0] = 0.0;			//节点所有出线功率累加数组赋初值0
            //    S_Count[i, 1] = 0.0;
            //    S_Node[i, 0] = 0.0;			//节点功率数组赋初值0
            //    S_Node[i, 1] = 0.0;
            //    S_Node[i, 2] = 0.0;
            //    S_Node[i, 3] = 0.0;
            //}
            Array.Clear(S_Count, 0, Num_Node * 2);//
            Array.Clear(S_Node, 0, 4 * Num_Node);
            //求线路潮流及各节点的所有出线功率累加
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0];		//取线路左节点
                j = Line_NodeName[k, 1];		//取线路右节点
                k_old = Line_No_NewtoOld[k];	//对应的旧线路顺序号
                Z[0] = LLine[k_old].RXBK[0];	//取线路的电阻值
                Z[1] = LLine[k_old].RXBK[1];	//取线路的电抗值
                BK = LLine[k_old].RXBK[2];	//取线路容纳半值或变压器非标准变比
                if (Line_Flag[k] == 0)			//普通支路
                {
                    if (i != j)				//非接地支路
                    {
                        Vi = Voltage[i, 1];
                        Vj = Voltage[j, 1];
                        Angi = Voltage[i, 0];
                        Angj = Voltage[j, 0];
                        V_T[0] = Vi * Math.Cos(Angi) - Vj * Math.Cos(Angj);
                        V_T[1] = Vi * Math.Sin(Angi) - Vj * Math.Sin(Angj);
                        Comp_Div(I_T, V_T, Z);
                        I1_T[0] = I_T[0] - BK * Vi * Math.Sin(Angi);			//Iij
                        I1_T[1] = I_T[1] + BK * Vi * Math.Cos(Angi);
                        I1_T[1] = -I1_T[1];						//取Iij的共轭
                        I2_T[0] = -I_T[0] - BK * Vj * Math.Sin(Angj);		//Iji
                        I2_T[1] = -I_T[1] + BK * Vj * Math.Cos(Angj);
                        I2_T[1] = -I2_T[1];						//取Iji的共轭
                        V_T[0] = Vi * Math.Cos(Angi);
                        V_T[1] = Vi * Math.Sin(Angi);
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);					//求Sij
                        S_Line[k, 0] = S_T[0];
                        S_Line[k, 1] = S_T[1];
                        S_Count[i, 0] = S_Count[i, 0] + S_T[0];	//节点i出线
                        S_Count[i, 1] = S_Count[i, 1] + S_T[1];	//功率累加。
                        V_T[0] = Vj * Math.Cos(Angj);
                        V_T[1] = Vj * Math.Sin(Angj);
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]);					//求Sji
                        S_Line[k, 2] = S_T[0];
                        S_Line[k, 3] = S_T[1];
                        S_Count[j, 0] = S_Count[j, 0] + S_T[0];	//节点j出线
                        S_Count[j, 1] = S_Count[j, 1] + S_T[1];	//功率累加。
                    }
                    else					//接地支路
                    {
                        Vi = Voltage[i, 1];
                        Angi = Voltage[i, 0];
                        V_T[0] = Vi * Math.Cos(Angi);
                        V_T[1] = Vi * Math.Sin(Angi);
                        Comp_Div(I_T, V_T, Z);
                        I1_T[0] = I_T[0] - BK * Vi * Math.Sin(Angi);	//求接地支路电流Iii
                        I1_T[1] = I_T[1] + BK * Vi * Math.Cos(Angi);
                        I1_T[1] = -I1_T[1];				//取Iii的共轭
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);
                        S_Line[k, 0] = S_T[0];		//求Sii(由节点i到地节点)
                        S_Line[k, 1] = S_T[1];
                        S_Line[k, 2] = 0.0;//求Sii(由地节点到节点i,其值等于零)
                        S_Line[k, 3] = 0.0;
                        S_Count[i, 0] = S_Count[i, 0] + S_T[0];	//节点i出线
                        S_Count[i, 1] = S_Count[i, 1] + S_T[1];	//功率累加。
                    }
                }
                else if (Line_Flag[k] == 1)	//非标准变比在节点i侧的变压器支路
                {
                    Vi = Voltage[i, 1];
                    Vj = Voltage[j, 1];
                    Angi = Voltage[i, 0];
                    Angj = Voltage[j, 0];
                    V_T[0] = Vi * Math.Cos(Angi) - Vj * Math.Cos(Angj);
                    V_T[1] = Vi * Math.Sin(Angi) - Vj * Math.Sin(Angj);
                    Comp_Div(I_T, V_T, Z);
                    I_T[0] = I_T[0] / BK;
                    I_T[1] = I_T[1] / BK;
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Div(I1_T, V_T, Z);
                    t = (1.0 / BK - 1.0) / BK;
                    I1_T[0] = I1_T[0] * t;
                    I1_T[1] = I1_T[1] * t;
                    I1_T[0] = I1_T[0] + I_T[0];						//Iij
                    I1_T[1] = I1_T[1] + I_T[1];
                    I1_T[1] = -I1_T[1];							//取Iij的共轭
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Div(I2_T, V_T, Z);
                    t = 1.0 - 1.0 / BK;
                    I2_T[0] = I2_T[0] * t - I_T[0];					//Iji
                    I2_T[1] = I2_T[1] * t - I_T[1];
                    I2_T[1] = -I2_T[1];							//取Iji的共轭
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);						//求Sij
                    S_Line[k, 0] = S_T[0];
                    S_Line[k, 1] = S_T[1];
                    S_Count[i, 0] = S_Count[i, 0] + S_T[0];		//节点i出线
                    S_Count[i, 1] = S_Count[i, 1] + S_T[1];		//功率累加。
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]);						//求Sji
                    S_Line[k, 2] = S_T[0];
                    S_Line[k, 3] = S_T[1];
                    S_Count[j, 0] = S_Count[j, 0] + S_T[0];		//节点j出线
                    S_Count[j, 1] = S_Count[j, 1] + S_T[1];		//功率累加。
                }
                else if (Line_Flag[k] == 2)	//非标准变比在节点j侧的变压器支路
                {
                    Vi = Voltage[i, 1];
                    Vj = Voltage[j, 1];
                    Angi = Voltage[i, 0];
                    Angj = Voltage[j, 0];
                    V_T[0] = Vi * Math.Cos(Angi) - Vj * Math.Cos(Angj);
                    V_T[1] = Vi * Math.Sin(Angi) - Vj * Math.Sin(Angj);
                    Comp_Div(I_T, V_T, Z);
                    I_T[0] = I_T[0] / BK;
                    I_T[1] = I_T[1] / BK;
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Div(I1_T, V_T, Z);
                    t = 1.0 - 1.0 / BK;
                    I1_T[0] = I1_T[0] * t;
                    I1_T[1] = I1_T[1] * t;
                    I1_T[0] = I1_T[0] + I_T[0];						//Iij
                    I1_T[1] = I1_T[1] + I_T[1];
                    I1_T[1] = -I1_T[1];							//取Iij的共轭
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Div(I2_T, V_T, Z);
                    t = (1.0 / BK - 1.0) / BK;
                    I2_T[0] = I2_T[0] * t - I_T[0];					//Iji
                    I2_T[1] = I2_T[1] * t - I_T[1];
                    I2_T[1] = -I2_T[1];							//取Iji的共轭
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);						//求Sij
                    S_Line[k, 0] = S_T[0];
                    S_Line[k, 1] = S_T[1];
                    S_Count[i, 0] = S_Count[i, 0] + S_T[0];			//节点i出线
                    S_Count[i, 1] = S_Count[i, 1] + S_T[1];			//功率累加。
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]);						//求Sji
                    S_Line[k, 2] = S_T[0];
                    S_Line[k, 3] = S_T[1];
                    S_Count[j, 0] = S_Count[j, 0] + S_T[0];			//节点j出线
                    S_Count[j, 1] = S_Count[j, 1] + S_T[1];			//功率累加。
                }
                else						//没有此种线路类型
                {
                    Console.WriteLine("There is not this line type!");
                    Environment.Exit(1);
                }
            }

            //将节点功率失配量赋初值为节点出线功率累加的负值
            for (i = 0; i < Num_Node; i++)
            {
                DS_Node[i, 0] = -S_Count[i, 0];
                DS_Node[i, 1] = -S_Count[i, 1];
            }

            //求取负荷节点功率PL和QL
            for (i = 0; i < Num_Load; i++)
            {
                j = Load_NodeName[i];							//负荷节点名称
                kl_old = Load_No_NewtoOld[i];					//负荷节点旧顺序号
                if (LLoad[kl_old].Flag == 1)					//计及静特性
                {
                    V = Voltage[j, 1];
                    S_Node[j, 2] = LLoad[kl_old].ABC[0] * V * V
                        + LLoad[kl_old].ABC[2] * V
                        + LLoad[kl_old].ABC[4];
                    S_Node[j, 3] = LLoad[kl_old].ABC[1] * V * V
                        + LLoad[kl_old].ABC[3] * V
                        + LLoad[kl_old].ABC[5];
                }
                else if (LLoad[kl_old].Flag == 0)					//不计静特性
                {
                    S_Node[j, 2] = LLoad[kl_old].ABC[4];
                    S_Node[j, 3] = LLoad[kl_old].ABC[5];
                }
                else								//没有此种类型的负荷节点
                {
                    Console.WriteLine("There is not this load node type!");
                    Environment.Exit(1);
                }
                DS_Node[j, 0] = DS_Node[j, 0] - S_Node[j, 2];	//功率失配量处理
                DS_Node[j, 1] = DS_Node[j, 1] - S_Node[j, 3];
            }

            //求取发电机节点功率PG和QG
            for (i = 0; i < Num_Gen; i++)
            {
                j = Gen_NodeName[i];						//发电机节点名称
                kg_old = Gen_No_NewtoOld[i];				//发电机节点旧顺序号
                if (GGen[kg_old].Flag == 0)				//平衡发电机节点
                {
                    S_Node[j, 0] = S_Count[j, 0] + S_Node[j, 2];
                    S_Node[j, 1] = S_Count[j, 1] + S_Node[j, 3];
                }
                else if (GGen[kg_old].Flag == 1)				//发电机PQ节点
                {
                    S_Node[j, 0] = GGen[kg_old].PQV[0];
                    S_Node[j, 1] = GGen[kg_old].PQV[1];
                }
                else if (GGen[kg_old].Flag == 2)				//发电机PV节点
                {
                    S_Node[j, 0] = GGen[kg_old].PQV[0];
                    S_Node[j, 1] = S_Count[j, 1] + S_Node[j, 3];
                }
                else								//没有此种类型的发电机节点
                {
                    Console.WriteLine("There is not this generator node type!");
                    Environment.Exit(1);
                }
                DS_Node[j, 0] = DS_Node[j, 0] + S_Node[j, 0];	//功率失配量处理
                DS_Node[j, 1] = DS_Node[j, 1] + S_Node[j, 1];
            }

            //节点潮流输出:序号,节点号,节点类型,V,Angle,PG,QG,PL,QL,节点有功失配
            //量,节点无功失配量。
            int[] Node_Name_Old = new int[NODEMAX];		//节点旧名（号）
            int[] Node_No_OldtoNew = new int[NODEMAX];	//排序后旧节点序号对应的新节点序号
            int I_Temp, np;					//临时变量
            int[] Node_Name_Voltage_MaxMin = new int[4];	//电压幅值最小、最大和电压
            //相位最小、最大节点号。
            double[] Voltage_Value_MaxMin = new double[4];		//最小、最大电压幅值和最小、
            //最大电压相位。
            for (i = 0; i < Num_Node; i++)
            {
                Node_Name_Old[i] = Node_Name_NewtoOld[i];
                Node_No_OldtoNew[i] = i;
            }
            for (i = 0; i < Num_Node - 1; i++)				//按照节点旧名由小到大排序
            {
                np = i;
                for (j = i + 1; j < Num_Node; j++)
                    if (Node_Name_Old[j] < Node_Name_Old[np]) np = j;
                I_Temp = Node_Name_Old[i];
                Node_Name_Old[i] = Node_Name_Old[np];
                Node_Name_Old[np] = I_Temp;
                I_Temp = Node_No_OldtoNew[i];
                Node_No_OldtoNew[i] = Node_No_OldtoNew[np];
                Node_No_OldtoNew[np] = I_Temp;
            }
            //屏幕输出部分
            Console.WriteLine("");
            Console.WriteLine(String.Format("{0,50:D}", "Node Flow Output"));
            Console.WriteLine(String.Format("{0,50:D}", "================"));
            Console.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-4:D}", "Bus")+
                String.Format("{0,-5:D}", "Type") + String.Format("{0,-8:D}", "Voltage")+
                String.Format("{0,-8:D}", "Angle") + String.Format("{0,-8:D}", "PG")+
                String.Format("{0,-8:D}", "QG") + String.Format("{0,-8:D}", "PL")+
                String.Format("{0,-8:D}", "QL") + String.Format("{0,-10:D}", "Bus_P_Dis")+
                String.Format("{0,-10:D}", "Bus_Q_Dis"));
            //磁盘输出部分
            if (Sav_result == true)
            {
                sw = new StreamWriter(fileName);
                sw.WriteLine(String.Format("{0,50:D}", "Node Flow Output"));
                sw.WriteLine(String.Format("{0,50:D}", "================"));
                sw.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-4:D}", "Bus") +
                    String.Format("{0,-5:D}", "Type") + String.Format("{0,-8:D}", "Voltage") +
                    String.Format("{0,-8:D}", "Angle") + String.Format("{0,-8:D}", "PG") +
                    String.Format("{0,-8:D}", "QG") + String.Format("{0,-8:D}", "PL") +
                    String.Format("{0,-8:D}", "QL") + String.Format("{0,-10:D}", "Bus_P_Dis") +
                    String.Format("{0,-10:D}", "Bus_Q_Dis"));
            }



            j = Node_No_OldtoNew[0];
            Node_Name_Voltage_MaxMin[0] = Node_Name_NewtoOld[j];
            Node_Name_Voltage_MaxMin[1] = Node_Name_NewtoOld[j];
            Node_Name_Voltage_MaxMin[2] = Node_Name_NewtoOld[j];
            Node_Name_Voltage_MaxMin[3] = Node_Name_NewtoOld[j];
            V = Voltage[j, 1];
            Angle = Voltage[j, 0] * Rad_to_Deg;
            Voltage_Value_MaxMin[0] = V;
            Voltage_Value_MaxMin[1] = V;
            Voltage_Value_MaxMin[2] = Angle;
            Voltage_Value_MaxMin[3] = Angle;
            for (i = 0; i < Num_Node; i++)
            {
                j = Node_No_OldtoNew[i];	//当前旧节点序号所对应的新节点顺序号
                k = Node_Name_NewtoOld[j];//新节点号所对应的旧节点名（号）
                V = Voltage[j, 1];
                Angle = Voltage[j, 0] * Rad_to_Deg;
                //屏幕输出部分
                Console.WriteLine(String.Format("{0,-3:D}", i) + String.Format("{0,-4:D}", k) +
                String.Format("{0,-5:D}", Node_Flag[j]) + String.Format("{0,-8:f4}", V) +
                String.Format("{0,-8:f2}", Angle) + String.Format("{0,-8:f4}", S_Node[j, 0]) +
                String.Format("{0,-8:f4}", S_Node[j, 1]) + String.Format("{0,-8:f4}", S_Node[j, 2]) +
                String.Format("{0,-8:f4}", S_Node[j, 3]) + String.Format("{0,-10:f6}", DS_Node[j, 0]) +
                String.Format("{0,-10:f6}", DS_Node[j, 1]));

                //磁盘输出部分
                if (Sav_result == true)
                {
                    sw.WriteLine(String.Format("{0,-3:D}", i) + String.Format("{0,-4:D}", k) +
                                   String.Format("{0,-5:D}", Node_Flag[j]) + String.Format("{0,-8:f4}", V) +
                                   String.Format("{0,-8:f2}", Angle) + String.Format("{0,-8:f4}", S_Node[j, 0]) +
                                   String.Format("{0,-8:f4}", S_Node[j, 1]) + String.Format("{0,-8:f4}", S_Node[j, 2]) +
                                   String.Format("{0,-8:f4}", S_Node[j, 3]) + String.Format("{0,-10:f6}", DS_Node[j, 0]) +
                                   String.Format("{0,-10:f6}", DS_Node[j, 1]));

                }
                if (i > 0)
                {
                    if (V < Voltage_Value_MaxMin[0])
                    {
                        Voltage_Value_MaxMin[0] = V;
                        Node_Name_Voltage_MaxMin[0] = k;
                    }
                    if (V > Voltage_Value_MaxMin[1])
                    {
                        Voltage_Value_MaxMin[1] = V;
                        Node_Name_Voltage_MaxMin[1] = k;
                    }
                    if (Angle < Voltage_Value_MaxMin[2])
                    {
                        Voltage_Value_MaxMin[2] = Angle;
                        Node_Name_Voltage_MaxMin[2] = k;
                    }
                    if (Angle > Voltage_Value_MaxMin[3])
                    {
                        Voltage_Value_MaxMin[3] = Angle;
                        Node_Name_Voltage_MaxMin[3] = k;
                    }
                }
            }

            //线路潮流输出:序号,左节点,右节点,线路类型,Pij,Qij,Pji,Qji,有功线损,
            //无功线损。
            int i_old, j_old;				//线路左、右节点旧号
            int[,] Node_Line = new int[LINEMAX, 2];		//线路左、右节点旧号工作数组
            int[] Type_Exchange = new int[LINEMAX];		//线路左、右节点旧号交换标志:
            //0-没交换,1-交换。
            int[] Line_No_OldtoNew = new int[LINEMAX];	//旧节点双排序后的线路号到新线路号
            //的转换数组。
            int kk, Line_Type;
            double DPLOSS = 0, DQLOSS = 0;		//系统总的有功、无功网损
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0];			//线路左节点新号
                j = Line_NodeName[k, 1];			//线路右节点新号
                i_old = Node_Name_NewtoOld[i];	//线路左节点旧号
                j_old = Node_Name_NewtoOld[j];	//线路右节点旧号
                if (i_old <= j_old)	//线路左节点旧号<=线路右节点旧号,不交换
                {
                    Node_Line[k, 0] = i_old;
                    Node_Line[k, 1] = j_old;
                    Type_Exchange[k] = 0;
                }
                else				//线路左节点旧号>线路右节点旧号,交换
                {
                    Node_Line[k, 0] = j_old;
                    Node_Line[k, 1] = i_old;
                    Type_Exchange[k] = 1;
                }
                Line_No_OldtoNew[k] = k;	//旧节点双排序后的线路号到新线路号的转
                //换数组赋初值
            }
            for (i = 0; i < Num_Line - 1; i++)
            {
                np = i;
                for (j = i + 1; j < Num_Line; j++)
                {
                    if (Node_Line[j, 0] < Node_Line[np, 0]
                        || (Node_Line[j, 0] == Node_Line[np, 0])
                        && Node_Line[j, 1] < Node_Line[np, 1]) np = j;
                }
                I_Temp = Node_Line[np, 0];
                Node_Line[np, 0] = Node_Line[i, 0];
                Node_Line[i, 0] = I_Temp;
                I_Temp = Node_Line[np, 1];
                Node_Line[np, 1] = Node_Line[i, 1];
                Node_Line[i, 1] = I_Temp;
                I_Temp = Type_Exchange[np];
                Type_Exchange[np] = Type_Exchange[i];
                Type_Exchange[i] = I_Temp;
                I_Temp = Line_No_OldtoNew[np];
                Line_No_OldtoNew[np] = Line_No_OldtoNew[i];
                Line_No_OldtoNew[i] = I_Temp;
            }

            //屏幕输出部分

            Console.WriteLine(String.Format("{0,50:D}", "Line Flow Output"));
            Console.WriteLine(String.Format("{0,50:D}", "================"));
            Console.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-5:D}", "LBus") +
                String.Format("{0,-5:D}", "RBus") + String.Format("{0,-5:D}", "Type") +
                String.Format("{0,-9:D}", "Pij") + String.Format("{0,-9:D}", "Qij") +
                String.Format("{0,-9:D}", "Pji") + String.Format("{0,-9:D}", "Qji") +
                String.Format("{0,-9:D}", "QL") + String.Format("{0,-10:D}", "P_Loss") +
                String.Format("{0,-10:D}", "Q_Loss"));
            //磁盘输出部分
            if (Sav_result == true)
            {
                sw.WriteLine("");
                sw.WriteLine(String.Format("{0,50:D}", "Line Flow Output"));
                sw.WriteLine(String.Format("{0,50:D}", "================"));
                sw.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-5:D}", "LBus") +
                    String.Format("{0,-5:D}", "RBus") + String.Format("{0,-5:D}", "Type") +
                    String.Format("{0,-9:D}", "Pij") + String.Format("{0,-9:D}", "Qij") +
                    String.Format("{0,-9:D}", "Pji") + String.Format("{0,-9:D}", "Qji") +
                    String.Format("{0,-9:D}", "QL") + String.Format("{0,-10:D}", "P_Loss") +
                    String.Format("{0,-10:D}", "Q_Loss"));
            }

            for (k = 0; k < Num_Line; k++)
            {
                kk = Line_No_OldtoNew[k];			//对应的新线路号
                Line_Type = Line_Flag[kk];		//线路类型标志
                if (Type_Exchange[k] == 1)	//线路左、右节点旧号发生过交换,需对变
                {						//压器支路线路类型标志做修改。
                    if (Line_Type == 1) Line_Type = 2;
                    else if (Line_Type == 2) Line_Type = 1;
                    //屏幕输出部分
                    Console.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
            String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
            String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
            String.Format("{0,-10:f4}", S_Line[kk, 3]));

                    //磁盘输出部分
                    if (Sav_result == true)
                    {
                        sw.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
                                   String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
                                   String.Format("{0,-10:f4}", S_Line[kk, 3]));

                    }
                }
                else
                {
                    //屏幕输出部分
                    Console.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
            String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
            String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
            String.Format("{0,-10:f4}", S_Line[kk, 3]));
                    //磁盘输出部分
                    if (Sav_result == true)
                    {
                        sw.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
                                   String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
                                   String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
                                   String.Format("{0,-10:f4}", S_Line[kk, 3]));
                    }
                }
                DPLOSS = DPLOSS + S_Line[kk, 0] + S_Line[kk, 2];
                DQLOSS = DQLOSS + S_Line[kk, 1] + S_Line[kk, 3];
            }

            //系统总体性能指标输出:有功网损,无功网损,最低电压值及其节点名,最高电压
            //值及其节点名,最小电压相角及其节点名,最大电压相角及其节点名,收敛次数,
            //计算时间。

            //屏幕输出部分
            Console.WriteLine(String.Format("{0,60:D}", "System Characteristics Index Output"));
            Console.WriteLine(String.Format("{0,60:D}", "==================================="));
            Console.WriteLine(String.Format("{0,-8:D}", "P_LOSS") + String.Format("{0,-8:D}", "Q_LOSS") +
                String.Format("{0,-8:D}", "Vmin") + String.Format("{0,-4:D}", "Bus") +
                String.Format("{0,-8:D}", "Vmax") + String.Format("{0,-4:D}", "Bus") +
                String.Format("{0,-8:D}", "ANGmin") + String.Format("{0,-4:D}", "Bus") +
                String.Format("{0,-8:D}", "ANGmax") + String.Format("{0,-4:D}", "Bus") +
                String.Format("{0,-8:D}", "Num_It") + String.Format("{0,-9:D}", "Time(s)"));
            Console.WriteLine(String.Format("{0,-8:f4}", DPLOSS) + String.Format("{0,-8:f4}", DQLOSS) +
                String.Format("{0,-8:f4}", Voltage_Value_MaxMin[0]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[0]) +
                String.Format("{0,-8:f4}", Voltage_Value_MaxMin[1]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[1]) +
                String.Format("{0,-8:f2}", Voltage_Value_MaxMin[2]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[2]) +
                String.Format("{0,-8:f2}", Voltage_Value_MaxMin[3]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[3]) +
                String.Format("{0,-8:D}", Num_Iteration) + String.Format("{0,-9:f3}", Duration));

            //磁盘输出部分
            if (Sav_result == true)
            {
                sw.WriteLine("");
                sw.WriteLine(String.Format("{0,60:D}", "System Characteristics Index Output"));
                sw.WriteLine(String.Format("{0,60:D}", "==================================="));
                sw.WriteLine(String.Format("{0,-8:D}", "P_LOSS") + String.Format("{0,-8:D}", "Q_LOSS") +
                    String.Format("{0,-8:D}", "Vmin") + String.Format("{0,-4:D}", "Bus") +
                    String.Format("{0,-8:D}", "Vmax") + String.Format("{0,-4:D}", "Bus") +
                    String.Format("{0,-8:D}", "ANGmin") + String.Format("{0,-4:D}", "Bus") +
                    String.Format("{0,-8:D}", "ANGmax") + String.Format("{0,-4:D}", "Bus") +
                    String.Format("{0,-8:D}", "Num_It") + String.Format("{0,-9:D}", "Time(s)"));
                sw.WriteLine(String.Format("{0,-8:f4}", DPLOSS) + String.Format("{0,-8:f4}", DQLOSS) +
                    String.Format("{0,-8:f4}", Voltage_Value_MaxMin[0]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[0]) +
                    String.Format("{0,-8:f4}", Voltage_Value_MaxMin[1]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[1]) +
                    String.Format("{0,-8:f2}", Voltage_Value_MaxMin[2]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[2]) +
                    String.Format("{0,-8:f2}", Voltage_Value_MaxMin[3]) + String.Format("{0,-4:D}", Node_Name_Voltage_MaxMin[3]) +
                    String.Format("{0,-8:D}", Num_Iteration) + String.Format("{0,-9:f3}", Duration));
                sw.Close();
                sw.Dispose();
            }
        }


        //将角度换算到区间[-PAI，PAI]
        void TreatAngle(ref double angle)
        {
            //angle = 0;
            if (angle > SinglePai)
                angle = angle - DoublePai;
            if (angle < -SinglePai)
                angle = angle + DoublePai;
        }

        public void PF_Main(string path)
        {
            string fileName;
            short i, kgpv, kg_old;

            short Num_Node = 0;			//总节点数
            short Num_Line = 0;			//总线路数
            short Num_Gen = 0;			//总发电机数
            short Num_Load = 0;			//总负荷数
            short Num_Swing = 0;		//总平衡机节点数
            short Num_GPV = 0;			//发电机节点中的PV节点总数
            short Num_GPQ = 0;			//发电机节点中的PQ节点总数
            int Num_Cur_Const = 0;	//常数电流节点总数
            short Iter_Max;			//迭代次数最大值
            short Num_Iteration = 0;	//迭代次数
            short VolIni_Flag;		//读电压初值标志:1-读;0-不读
            short VolRes_Flag;		//保留电压(作为下次初值)标志
            //:1-保留;0-不保留。

            float Eps = 1e-5f;		//节点功率失配量值收敛限值
            double Power_Dis_Max;	//最大节点功率失配量值
            double Duration;		//存放计算时间(s)

            double[,] Power_Dis = new double[NODEMAX, 2];		//功率失配量dP、dQ
            double[,] Pij_Sum = new double[NODEMAX, 2];			//节点所有出线功率累加
            double[] Power_Dis_Correct = new double[NODEMAX];	//dP/V(dSita.V)、dQ/V(dV)
            double[,] DVolt = new double[NODEMAX, 2];			//电压修正量dSita.V、dV

            Data_Input(out Num_Line, out Num_Gen, out Num_Load, out Eps, out Iter_Max,
        out VolIni_Flag, out VolRes_Flag, path);					//数据输入
            DateTime BeginTime = DateTime.Now;
            Node_Sequen(out Num_Node, Num_Line, Num_Gen, Num_Load,
                out Num_Swing, out Num_GPV, out Num_GPQ);			//序号处理
            Y_Bus1(Num_Node, Num_Line, Num_Swing);	//第一导纳阵
            Factor_Table(Num_Node, Num_Swing, Num_GPV, 0);		//形成第一因子表

            Y_Bus2(Num_Node, Num_Line, Num_Load, Num_Swing);	//第二导纳阵
            Factor_Table(Num_Node, Num_Swing, Num_GPV, 1);		//形成第二因子表
            if (VolIni_Flag == 1)
            {
                fileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "vol.ini";
                if (!File.Exists(fileName))
                {
                    sw = File.CreateText(fileName);
                    sw.Close();
                }
                Voltage_Initial(Num_Node, Num_Swing, fileName);
            }
            else
                for (short m = 0; m < Num_Node - Num_Swing; m++)
                {
                    Voltage[m, 0] = 0.0;				//电压相位的初值：0.0(弧度)
                    Voltage[m, 1] = 1.0;				//电压幅值的初值：1.0
                }
            Initial(Num_Node, Num_Line, Num_Load, Num_Swing,
                Num_GPV, Num_GPQ, out Num_Cur_Const, DVolt);				//初始化
        Iteration_Back:										//迭代开始点
            PowerDis_Comp(Num_Node, Num_Load, Num_Swing,
                          Num_GPV, Num_Cur_Const, Power_Dis,
                          Pij_Sum, DVolt,
                          Num_Iteration, out Power_Dis_Max);		//计算功率失配量
            if (Power_Dis_Max > Eps && Num_Iteration < Iter_Max)	//收敛判断
            {
                for (i = 0; i < Num_Node - Num_Swing; i++)
                    Power_Dis_Correct[i] = Convert.ToSingle(
                    Power_Dis[i, 0] / Voltage[i, 1]);			//取dP/V
                Equ_Calculation(Num_Node, Num_Swing, Power_Dis_Correct, 0);
                //有功求解

                for (i = 0; i < Num_Node - Num_Swing; i++)
                {
                    Voltage[i, 0] = Voltage[i, 0]
                        - Power_Dis_Correct[i] / Voltage[i, 1];	//修正相位
                    TreatAngle(ref Voltage[i, 0]);
                    DVolt[i, 0] = Power_Dis_Correct[i];	//保存相位差dSita.V
                    Power_Dis_Correct[i] =
                        Power_Dis[i, 1] / Voltage[i, 1];	//取dQ/V
                }
                Equ_Calculation(Num_Node, Num_Swing, Power_Dis_Correct, 1);
                //无功求解

                for (i = 0; i < Num_Node - Num_Swing; i++)
                {
                    Voltage[i, 1] = Voltage[i, 1]
                        - Power_Dis_Correct[i];			//修正幅值
                    DVolt[i, 1] = Power_Dis_Correct[i];	//保存幅值差dV
                }
                for (kgpv = 0; kgpv < Num_GPV; kgpv++)	//发电机PV节点的电压幅值=VG
                {
                    i = Gen_PVNode[kgpv, 0];
                    kg_old = Gen_PVNode[kgpv, 1];
                    Voltage[i, 1] = GGen[kg_old].PQV[1];
                    DVolt[i, 1] = 0.0f;					//PV节点电压幅值差=0.0
                }
                Num_Iteration++;						//迭代次数增1
                goto Iteration_Back;					//迭代折返点
            }
            //DateTime BeginTime = DateTime.Now;
            TimeSpan span = DateTime.Now.Subtract(BeginTime);
            Duration = span.TotalSeconds;
            //Console.WriteLine("计算完成！共用时：" + Duration.ToString() + "毫秒！");
            if (VolRes_Flag == 1)
            {
                fileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "vol.ini";
                if (!File.Exists(fileName))
                {
                    sw = File.CreateText(fileName);
                    sw.Close();
                }
                Voltage_Reserve(Num_Node, Num_Swing,fileName);	//保存电压(按内部号)
            }
            fileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + 
                Path.GetFileNameWithoutExtension(path)+"_PF.txt";
            if (!File.Exists(fileName))
            {
                sw = File.CreateText(fileName);
                sw.Close();
            }
            if (Num_Iteration != Iter_Max)
                Result_Output(Num_Node, Num_Line, Num_Gen, Num_Load,
                            Num_Swing, Num_Iteration, Duration,fileName);	//结果输出
            else
            {
                Console.WriteLine("======================================================================");
                Console.WriteLine("==========input data hasn't converged in" + Iter_Max + "iterations!===========");
                Console.WriteLine("======================================================================");
            }
            Console.WriteLine("press any key to quit！");
            Console.ReadKey();
        }
    }

    #region  输入数据结构定义
    class Line              //线路参数结构定义
    {
        private short[] _Node_No = new short[2];    //线路两端节点名(号):0-左节点;1-右节点
        private byte _Flag;                         //线路类型标志:0-普通支路;1、2-变压器支路。其中,		
        //1-非标准变比在左侧节点,2-非标准变比在右侧节点。
        //对接地支路，左右节点名（号）相同。
        private float[] _RXBK = new float[3];       //0-R;1-X;2-Bc/2 或 K

        public short[] Node_No
        {
            get { return _Node_No; }
            set { _Node_No = value; }
        }

        public byte Flag
        {
            get { return _Flag; }
            set { _Flag = value; }
        }

        public float[] RXBK
        {
            get { return _RXBK; }
            set { _RXBK = value; }
        }
    }

    class Generator     //发电机参数结构定义
    {
        private short _Node_No;//发电机节点名(号)
        private byte _Flag;//发电机节点类型标志:0-平衡节点;1-PQ节点;2-PV节点
        private float[] _PQV = new float[2];//对平衡节点，0-V,1-Angle;对PQ节点,0-P,1-Q;对PV节
                                            //点,0-P,1-V。

        public short Node_No
        {
            get { return _Node_No; }
            set { _Node_No = value; }
        }

        public byte Flag
        {
            get { return _Flag; }
            set { _Flag = value; }
        }

        public float[] PQV
        {
            get { return _PQV; }
            set { _PQV = value; }
        }
    }

    class Load         //负荷参数结构定义
    {
        private short _Node_No;//负荷节点名(号)
        private byte _Flag;//负荷节点静特性标志:0-不计静特性;1-计静特性。
        private float[] _ABC = new float[6];//PL=a1*V**V+b1*V+c1,QL=a2*V*V+b2*V+c2。
                                            //0-a1;1-a2;2-b1;3-b2;4-c1;5-c2。

        public short Node_No
        {
            get { return _Node_No; }
            set { _Node_No = value; }
        }

        public byte Flag
        {
            get { return _Flag; }
            set { _Flag = value; }
        }

        public float[] ABC
        {
            get { return _ABC; }
            set { _ABC = value; }
        }
    }
    #endregion
}
