using System;
using System.Collections.Generic;
using System.Text;
//using System.Windows.Forms;
using System.IO;

namespace Power_Flow.ClassFile
{
    class PF_Calc
    {
        static short LINEMAX = 5000;              //Maximum number of lines
        static short GENERATORMAX = 500;          //Maximum number of generators
        static short LOADMAX = 2000;		        //Maximum load
        static short NODEMAX = 2000;		        //Maximum number of nodes
        static byte SWINGMAX = 20;		        //Maximum number of balanced nodes
        static short PVMAX = 500;			        //Maximum number of PV nodes
        static byte NODEFACTOR = 10;		        //The number of non-zero off-diagonal elements in the admittance matrix relative to
                                                    //The multiple of the maximum number of nodes (NODEMAX)
        static float Deg_to_Rad = 0.017453292f;    //Degree to radians conversion factor
        static float Rad_to_Deg = 57.29577951f;	//Radius to degree conversion factor
        static float SinglePai = 3.14159265f;	    //PI
        static float DoublePai = 6.2831853f;	    //Double the pi

        short[] Node_Name_NewtoOld = new short[NODEMAX];//New node name (number) --> old node name (number)
        //- Save the node number sorted by the number of outgoing lines from small to large
        short[] Node_Flag = new short[NODEMAX];			//Node type flag: 0-balance, 1-PQ, 2-PV
        short[,] Line_NodeName = new short[LINEMAX, 2];	//The new name (number) of the left and right nodes of the line
        short[] Line_No_NewtoOld = new short[LINEMAX];	//New line number --> old line numbe
        short[] Line_Flag = new short[LINEMAX];			//Type mark of new line: 0, 1, 2 Description with Line structure
        short[] Gen_NodeName = new short[GENERATORMAX];	//New node name (number) of the generator node
        short[] Gen_No_NewtoOld = new short[GENERATORMAX];	//New generator sequence number --> old generator sequence number
        short[,] Gen_SWNode = new short[SWINGMAX, 2];	//Balance node data: 0 - new node name (number);
        //1-Corresponding old generator sequence number
        short[,] Gen_PVNode = new short[PVMAX, 2];		//Generator PV node data: 0 - new node name (number);
        //1-Corresponding old generator sequence number
        short[,] Gen_PQNode = new short[GENERATORMAX, 2];//Generator PQ node data: 0 - new node name (number);
        //1-Corresponding old generator sequence number
        short[] Load_NodeName = new short[LOADMAX];		//The new node name (number) of the load node
        short[] Load_No_NewtoOld = new short[LOADMAX];	//New load sequence number --> old load sequence number
        bool Sav_result = false;

        public Line[] LLine;
        public Generator[] GGen;
        public Load[] LLoad;
        StreamWriter sw;
        //Read data
        private void Data_Input(out short Num_Line, out short Num_Gen, 
            out short Num_Load, out float Eps, out short Iter_Max,
        out short VolIni_Flag, out short VolRes_Flag, string path)
        {
            //DateTime BeginTime = DateTime.Now;
            //*********This code assigns an initial value to the output parameter.***********//
            Num_Line = 0;
            Num_Gen = 0;
            Num_Load = 0;
            Eps = 1.0e-5f;
            Iter_Max = 200;
            VolIni_Flag = 0;
            VolRes_Flag = 0;
            //*********End*****************************//
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
                if (rows <= Num_Line)//Read line parameters
                {
                    LLine[j] = new Line();
                    LLine[j].Node_No[0] = Int16.Parse(parts[0]);
                    LLine[j].Node_No[1] = Int16.Parse(parts[1]);
                    LLine[j].Flag = byte.Parse(parts[2]);
                    for (short n = 0; n < 3; n++)
                        LLine[j].RXBK[n] = float.Parse(parts[n + 3]);
                    j++;
                }
                else if ((rows > Num_Line) && (rows <= Num_Line + Num_Gen))//Read generator parameters
                {
                    GGen[p] = new Generator();
                    GGen[p].Node_No = Int16.Parse(parts[0]);
                    GGen[p].Flag = byte.Parse(parts[1]);
                    GGen[p].PQV[0] = float.Parse(parts[2]);
                    GGen[p].PQV[1] = float.Parse(parts[3]);
                    p++;
                }
                else if ((rows > Num_Line + Num_Gen) && (rows <= Num_Gen + Num_Line + Num_Load))//Read load parameter
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

        //Serial number processing subroutine
        void Node_Sequen(out short Num_Node, int Num_Line, int Num_Gen, int Num_Load,
                         out short Num_Swing, out short Num_GPV, out short Num_GPQ)
        {
            Num_Node = 0;
            Num_Swing = 0;
            Num_GPV = 0;
            Num_GPQ = 0;
            //bool sign = true;
            short i, j, Flag, temp, np;
            short[,] Node_Name = new short[NODEMAX, 2];			//0-Node name (number); 1-node outgoing line number 

            //Count the number of outgoing lines of each node
            //for (i = 0; i < NODEMAX; i++) Node_Name[i, 1] = 0; //The number of node outlets is initialized to 0.
            Array.Clear(Node_Name, 0, Node_Name.Length);//The number of node outlets is initialized to 0.
            for (i = 0; i < Num_Line; i++)
            {
                if (LLine[i].Node_No[0] == LLine[i].Node_No[1])
                    continue;		//Grounding branch (same left and right nodes) are not within the outlet statistics
                Flag = 0;							//Left node line number analysis begins
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[0] == Node_Name[j, 0])//The node is already at the node
                    {										//Appear in the array, just
                        Node_Name[j, 1]++;					//Add 1 to the number of outlets.
                        Flag = 1;
                    };
                    if (Flag == 1) break;
                }
                if (Flag == 0)										//The node is not yet in
                {												//Out of the node array
                    Node_Name[Num_Node, 0] = LLine[i].Node_No[0];	//Now, the section is needed
                    Node_Name[Num_Node, 1]++;					//Add a name (number)
                    Num_Node++;									//Into the node array,
                    if (Num_Node > NODEMAX)						//Then the node's
                    {											//Add 1 to the number of lines and increase the number of nodes by 1.
                        Console.WriteLine("Nodes Number > " + NODEMAX + "!");                     
                        Environment.Exit(1);
                    }
                }
                Flag = 0;							//Right node output number analysis begins
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[1] == Node_Name[j, 0])//The node is already at the node
                    {										//Appear in the array, just
                        Node_Name[j, 1]++;					//Number of outlets plus 1
                        Flag = 1;
                    };
                    if (Flag == 1) break;
                }
                if (Flag == 0)										//The node is not yet in
                {												//Out of the node array
                    Node_Name[Num_Node, 0] = LLine[i].Node_No[1];	//Now, the section is needed
                    Node_Name[Num_Node, 1]++;					//Add a name (number)
                    Num_Node++;									//Into the node array,
                    if (Num_Node > NODEMAX)						//Then the node's
                    {											//The number of outlets is increased by 1, and
                                                                //Add the number of nodes
                        Console.WriteLine("Node Numbers > " + NODEMAX + "!");                     
                        Environment.Exit(1);
                    }
                }
            }
            //The number of node outlets is counted.
            //Sort nodes according to the number of outgoing lines from small to large (bubble algorithm)
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
            //Balance node statistics: total number and name of each node (number)
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
            //Sort the nodes according to the number of outgoing lines from small to large, and rank the balanced nodes at the end (the largest serial number)
            int Nswing = 0, Nnode = 0;
            for (i = 0; i < Num_Node; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Swing; j++)
                {
                    if (Node_Name[i, 0] == Node_Name_Swing[j]) Flag = 1;
                    if (Flag == 1) break;		//Flag=1When the node is a balanced node,
                }							//Need to be placed in the back position.
                if (Flag == 0)
                {
                    Node_Name_NewtoOld[Nnode] = Node_Name[i, 0];
                    Nnode++;
                }
                else	//The last balance nodes are still arranged in the order of the number of outgoing lines.
                {
                    Node_Name_NewtoOld[Num_Node - Num_Swing + Nswing] =
                        Node_Name[i, 0];
                    Nswing++;
                }
            }
            //New line type flag assignment initial value
            for (i = 0; i < Num_Line; i++) Line_Flag[i] = LLine[i].Flag;
            //Line name (number) processing: becomes the new node name (number) and the absolute value of the left node is less than the absolute value of the right node
            for (i = 0; i < Num_Line; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[0] == Node_Name_NewtoOld[j])//Left node processing
                    {
                        Line_NodeName[i, 0] = j;	//Assign a new name (number)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLine[i].Node_No[1] == Node_Name_NewtoOld[j])//Right node processing		
                    {
                        Line_NodeName[i, 1] = j;	//Assign a new name (number)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
                if (Line_NodeName[i, 0] > Line_NodeName[i, 1])//The absolute value of the left node is less than
                {										   //Absolute value processing of the right node
                    if (LLine[i].Flag == 1) Line_Flag[i] = 2;		//Non-standard transformer
                    if (LLine[i].Flag == 2) Line_Flag[i] = 1;		//Change on the ratio side
                    temp = Line_NodeName[i, 0];
                    Line_NodeName[i, 0] = Line_NodeName[i, 1];
                    Line_NodeName[i, 1] = temp;
                }
            }
            //Line sorting: according to the absolute value of the left node from small to large, if the absolute value of the left node is equal, follow the right section
            //The absolute value of the point is sorted from small to large (double sorting bubble algorithm)
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
            //Generator node name (number) processing: become a new node name (number)
            for (i = 0; i < Num_Gen; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (GGen[i].Node_No == Node_Name_NewtoOld[j])
                    {
                        Gen_NodeName[i] = j;						//Assign a new name (number)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
            }
            //Generator sequencing: sort by new node name (number) from small to large, and find the new generator serial number
            //Corresponding old generator serial number
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
            //Load node name (number) processing: become the new node name (number)
            for (i = 0; i < Num_Load; i++)
            {
                Flag = 0;
                for (j = 0; j < Num_Node; j++)
                {
                    if (LLoad[i].Node_No == Node_Name_NewtoOld[j])
                    {
                        Load_NodeName[i] = j;						//Assign a new name (number)
                        Flag = 1;
                    }
                    if (Flag == 1) break;
                }
            }
            //Load sorting: sort by new node name (number) from small to large, and find the new load serial number
            //Corresponding old load serial number
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
            //The new node name (number) and pair of balanced node, PV node, and PQ node are summarized from the generator node data.
            //The old generator serial number should be corrected, and the node type flag should be corrected for the balance node and the PV node.
            for (i = 0; i < Num_Node; i++) Node_Flag[i] = 1;	//The node type is assigned an initial value of 1 (PQ node)
            Nswing = 0;
            for (i = 0; i < Num_Gen; i++)
            {
                j = Gen_No_NewtoOld[i];					//Generator node old sequence number
                if (GGen[j].Flag == 0)
                {
                    Gen_SWNode[Nswing, 0] = Gen_NodeName[i];	//Generator node name
                    Gen_SWNode[Nswing, 1] = j;
                    Node_Flag[Gen_NodeName[i]] = 0;
                    Nswing++;
                }
                else if (GGen[j].Flag == 1)
                {
                    Gen_PQNode[Num_GPQ, 0] = Gen_NodeName[i];	//Generator node name
                    Gen_PQNode[Num_GPQ, 1] = j;
                    (Num_GPQ)++;
                }
                else if (GGen[j].Flag == 2)
                {
                    Gen_PVNode[Num_GPV, 0] = Gen_NodeName[i];	//Generator node name
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


        double[,] Y_Diag = new double[NODEMAX, 2];				//Diagonal element of node admittance array: 0 - real part;
        //1-The imaginary part.
        double[,] Y_UpTri = new double[NODEMAX * NODEFACTOR, 2];	//The non-zero element of the triangle on the node admittance array:
        //0-Real part; 1- imaginary part.
        int[] Foot_Y_UpTri = new int[NODEMAX * NODEFACTOR];	//The upper triangle compresses the stored non-zero elements by row
        //List the code.
        int[] Num_Y_UpTri = new int[NODEMAX];				//The number of non-zero elements in each line of the upper triangle
        int[] No_First_Y_UpTri = new int[NODEMAX];			//The first non-zero element in each row of the upper triangle is
        //Y_UpTri The sequence number in .
        int[] Foot_Y_DownTri = new int[NODEMAX * NODEFACTOR];	//The lower triangle compresses the stored non-zero elements by row
        //List the code.
        int[] Num_Y_DownTri = new int[NODEMAX];				//The number of non-zero elements in each row of the lower triangle
        int[] No_First_Y_DownTri = new int[NODEMAX];		//The first non-zero element in each row of the lower triangle is pressed
        //The sequence number in the row compression storage sequence
        int[] No_Y_DownTri_RowtoCol = new int[NODEMAX * NODEFACTOR];	//The lower triangle is not zero
        //Should compress the storage sequence by column
        //Serial number

        //Forming a node admittance matrix 1 (excluding the effects of line charging accommodation and non-standard ratio)
        void Y_Bus1(int Num_Node, int Num_Line, int Num_Swing)
        {
            int i, j, k, k_old, Flag, l;
            double X, B;							//Line parameter work unit
            l = 0;
            
            Array.Clear(Y_Diag, 0, 2 * (Num_Node - Num_Swing));//
            Array.Clear(Num_Y_UpTri, 0, Num_Node - Num_Swing);
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0];		//Line left node
                j = Line_NodeName[k, 1];		//Line right node
                if (i >= Num_Node - Num_Swing)	//The left and right nodes are balanced nodes, and there is no admittance array.
                    break;					//influences.
                k_old = Line_No_NewtoOld[k];	//Corresponding old line sequence number
                X = LLine[k_old].RXBK[1];		//Line reactance value
                B = -1.0 / X;					//Line branch susceptance after line resistance
                if (j >= Num_Node - Num_Swing)	//Left is a normal node, right is a balanced node
                    Y_Diag[i, 1] = Y_Diag[i, 1] + B;//Only self-slave of the left node (ordinary node) is considered, excluding mutual susceptance
                else						//Left and right nodes are common nodes
                {
                    Flag = 0;
                    if (k > 0 && (i == Line_NodeName[k - 1, 0])
                        && (j == Line_NodeName[k - 1, 1])) Flag = 1;	//Multiple loops
                    Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                    if (i != j)								//Ungrounded branch
                    {
                        Y_Diag[j, 1] = Y_Diag[j, 1] + B;
                        if (Flag == 0)							//First line
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
                        else										//Multiple loops
                        {
                            Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - B;
                        }
                    }
                }
            }

            No_First_Y_UpTri[0] = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
                No_First_Y_UpTri[i + 1] = No_First_Y_UpTri[i] + Num_Y_UpTri[i];
            //The number of non-zero elements in each row of the admittance matrix under the triangle compression storage, the first non-zero element order of each row
            //No., non-zero element column code when compressed by row, lower triangle array corresponding to a non-zero element
            //The same non-zero element number when storing, the purpose of these values ​​is to quickly process the following
            //In the process of generation, the values ​​of the Jacobian matrix are corrected, and the values ​​are subtracted and returned.
            int[] Row_Down = new int[NODEMAX];		//The lower bound work unit of a non-zero element number in a lower triangle
            int[] Row_Up = new int[NODEMAX];		//Upper bound work unit of a non-zero element number in a lower triangle
            int li;
            //for (i = 0; i < Num_Node - Num_Swing; i++)//The lower triangle rows are non-zero element arrays cleared
            //    Num_Y_DownTri[i] = 0;
            Array.Clear(Num_Y_DownTri, 0, Num_Node - Num_Swing);//
            for (j = 0; j < Num_Node - Num_Swing; j++)//The loop counts the number of non-zero elements in each triangle
            {
                for (k = No_First_Y_UpTri[j]; k < No_First_Y_UpTri[j + 1]; k++)
                {								//For the third triangle of the lower triangle, the non-zero element is processed.
                    i = Foot_Y_UpTri[k];			//Line code
                    Num_Y_DownTri[i]++;			//The number of non-zero elements in the i-th row of the lower triangle is increased by one.
                }
                Row_Down[j] = No_First_Y_UpTri[j];
                Row_Up[j] = No_First_Y_UpTri[j + 1];
            }
            No_First_Y_DownTri[0] = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)	//The first non-zero element number in each row of the lower triangle
                No_First_Y_DownTri[i + 1] = No_First_Y_DownTri[i] + Num_Y_DownTri[i];
            for (i = 1; i < Num_Node - Num_Swing; i++)	//The loop determines that the lower triangles are non-zero
            {									//List the code.
                li = No_First_Y_DownTri[i];		//The first non-zero element number of the i-th row of the lower triangle
                for (j = 0; j < i; j++)				//The loop searches for the lower triangle in column 0~i-1
                {								//A non-zero element with line number i.

                    if ((k = Row_Down[j]) < Row_Up[j])//If the jth row of the upper triangle has a non-zero non-diagonal element. This code is used to replace the following piece of code
                    {
                        if (Foot_Y_UpTri[k] == i)//Indicates that the i-th row of the lower triangle has a non-diagonal non-diagonal element
                        {
                            Foot_Y_DownTri[li] = j;//Record the i-th non-zero-element column of the i-line
                            No_Y_DownTri_RowtoCol[li] = k;//Record the element in the lower triangle
                            //Column compression storage sequence number
                            li++;				//The serial number counter is incremented by 1, ready for the next use.
                            Row_Down[j]++;
                        }
                    }
                    
                }
            }
        }


        //Forming a node admittance matrix 2 (including the effects of line charging accommodation and non-standard variation ratio)
        void Y_Bus2(int Num_Node, int Num_Line, int Num_Load, int Num_Swing)
        {
            int i, j, k, k_old, Flag, l;
            double R, X, Z, G, B, BK;

            l = 0;
            
            Array.Clear(Y_Diag, 0, 2 * (Num_Node - Num_Swing));//
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0]; //Line left node
                j = Line_NodeName[k, 1]; //Line right node
                if (i >= Num_Node - Num_Swing) //The left and right nodes are balanced nodes, no for the admittance array
                    break; //Impact.
                k_old = Line_No_NewtoOld[k]; // corresponding old line sequence number
                R = LLine[k_old].RXBK[0]; // take the line resistance value
                X = LLine[k_old].RXBK[1]; //Get the line reactance value
                BK = LLine[k_old].RXBK[2]; // Take the line to accommodate the half value or the transformer ratio
                Z = R * R + X * X;
                G = R / Z; // conductance
                B = -X / Z; //Wiener
                if (j >= Num_Node - Num_Swing) //Left is the normal node, right is the balance node
                {
                    if (Line_Flag[k] == 0)					//Ordinary branch
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                    }
                    else if (Line_Flag[k] == 1)			//Non-standard ratio is on the left node
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + 1.0 / BK / BK * G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + 1.0 / BK / BK * B;
                    }
                    else if (Line_Flag[k] == 2)			//Non-standard ratio in the right node
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                    }
                }
                
                else //left and right nodes are ordinary nodes
                {
                    Flag = 0;
                    if (k > 0 && (i == Line_NodeName[k - 1, 0])
                        && (j == Line_NodeName[k - 1, 1])) Flag = 1; //Multiple loops
                    if (i == j) //ground branch (transformer branch is not directly grounded)
                    {
                        Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                        Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                    }
                    else //non-grounded branch
                    {
                        if (Line_Flag[k] == 0) // ordinary branch
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + B + BK;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + B + BK;
                            if (Flag == 0) //first round
                            {
                                Y_UpTri[l, 0] = -G;
                                Y_UpTri[l, 1] = -B;
                                l++;
                            }
                            else //multiple loops
                            {
                                Y_UpTri[l, 0] = Y_UpTri[l, 0] - G;
                                Y_UpTri[l, 1] = Y_UpTri[l, 1] - B;
                            }
                        }
                        else if (Line_Flag[k] == 1)		//Non-standard ratio is on the left node
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + 1.0 / BK / BK * G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + 1.0 / BK / BK * B;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + B;
                            if (Flag == 0)							//First line
                            {
                                Y_UpTri[l, 0] = -1.0 / BK * G;
                                Y_UpTri[l, 1] = -1.0 / BK * B;
                                l++;
                            }
                            else								//Multiple loops
                            {
                                Y_UpTri[l - 1, 0] = Y_UpTri[l - 1, 0] - 1.0 / BK * G;
                                Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - 1.0 / BK * B;
                            }
                        }
                        else							//Non-standard ratio in the right node
                        {
                            Y_Diag[i, 0] = Y_Diag[i, 0] + G;
                            Y_Diag[i, 1] = Y_Diag[i, 1] + B;
                            Y_Diag[j, 0] = Y_Diag[j, 0] + 1.0 / BK / BK * G;
                            Y_Diag[j, 1] = Y_Diag[j, 1] + 1.0 / BK / BK * B;
                            if (Flag == 0)							//First line
                            {
                                Y_UpTri[l, 0] = -1.0 / BK * G;
                                Y_UpTri[l, 1] = -1.0 / BK * B;
                                l++;
                            }
                            else								//Multiple loops
                            {
                                Y_UpTri[l - 1, 0] = Y_UpTri[l - 1, 0] - 1.0 / BK * G;
                                Y_UpTri[l - 1, 1] = Y_UpTri[l - 1, 1] - 1.0 / BK * B;
                            }
                        }
                    }
                }
            }

            //Calculate the impedance component of the static characteristics of the load into the diagonal elements of the admittance matrix
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

        //The complex numbers A and B are multiplied by C (orthogonal coordinate form)
        void Comp_Mul(out double C0, out double C1,
            double A1, double A2, double B1, double B2)
        {

            //C=new double[2];A=new double[2];B=new double[2];
            C0 = A1 * B1 - A2 * B2; C1 = A1 * B2 + A2 * B1;
            //C[0]=A[0]*B[0]-A[1]*B[1];	C[1]=A[0]*B[1]+A[1]*B[0];
        }


        //Complex A and B are divided by C (orthogonal coordinate form)
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


        double[,] Fact_Diag = new double[NODEMAX, 2];				//Factor table diagonal element: 0 - active factor
        //Sub-table; 1-reactive factor table.
        double[,] Fact_UpTri = new double[NODEMAX * NODEFACTOR, 2];	//Triangular non-zero elements on the factor table: 0-
        //Active factor table; 1-reactive factor table.
        int[,] Foot_Fact_UpTri = new int[NODEMAX * NODEFACTOR, 2];	//Triangle non-zero element on the factor table
        int[,] Num_Fact_UpTri = new int[NODEMAX, 2];				//The triangles on the factor table are non-zero and non-pair
        //The number of corners.
        //Node admittance matrix factor table
        void Factor_Table(int Num_Node, int Num_Swing, int Num_GPV, int IterFlag)
        {
            
            int i; //The line number where the factor table is being formed
            int im; //The line number where the factor table is being erased
            int j; // column foot code temporary storage unit
            int k; //temporary counting unit
            int ix; // the triangle element address (serial number) count on the factor table
            Double[] Y_Work = new double[NODEMAX]; //Working array
            Double Temp1, Temp2; //temporary work unit
            int kgpv;

            for (i = 0; i < Num_Node - Num_Swing; i++) // traverse all rows to form a complete factor table, the entire calculation process with the line number i as a loop variable
            {
                if (IterFlag == 1 && Node_Flag[i] == 2) //Reactive PV node corresponding to reactive iteration
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

                    for (k = No_First_Y_UpTri[i]; k < No_First_Y_UpTri[i + 1]; k++)//The first non-zero element in each row of the upper triangle is
                                                                                   //The sequence number in Y_UpTri.
                    {
                        j = Foot_Y_UpTri[k];//The upper triangle compresses the stored non-zero elements by row
                                            //List the code.
                        Y_Work[j] = Y_UpTri[k, 1];//Get the data of the i-th row
                    }
                    if (IterFlag == 1)
                    {
                        for (kgpv = 0; kgpv < Num_GPV; kgpv++)
                        {
                            j = Gen_PVNode[kgpv, 0];			//PV node number
                            Y_Work[j] = 0.0;
                        }
                    }

                    ix = 0;
                    for (im = 0; im < i; im++)//When the PV node number forms the elements in the i-th row of the factor table, the working array should be erased with the row factor table elements that have been formed before the i-1 line.
                    {
                        for (k = 0; k < Num_Fact_UpTri[im, IterFlag]; k++)//The number of non-zero non-diagonal elements in each line of the triangle on the factor table.
                        {
                            if (Foot_Fact_UpTri[ix, IterFlag] != i)/*
                                                                    When the column code is equal to i, the erasure operation is performed.
                                                                   In general, when the line number to be erased in the working array is i, the first im (im < i) line
                                                                    When there is no element with the foot code i (this means Uim, i=0), you can save the im line.
                                                                    Elimination process*/
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
                    for (j = i + 1; j < Num_Node - Num_Swing; j++)//Only form the upper triangle
                    {
                        if (Math.Abs(Y_Work[j]) != 0.0)
                        {
                            Fact_UpTri[ix, IterFlag] = Y_Work[j] * Temp1;//Upper triangular array element, compressed storage
                            Foot_Fact_UpTri[ix, IterFlag] = j;//The upper triangle element compresses the stored column code
                            k++;
                            ix++;
                        }
                    }
                    Num_Fact_UpTri[i, IterFlag] = k;//The number of non-zero non-diagonal elements in each line of the triangle on the factor table.  
                }
            }
        }

        //Equation AX=t
        void Equ_Calculation(int Num_Node, int Num_Swing,
                             double[] Power_Dis_Correct, int IterFlag)
        {
            //Power_Dis_Correct=new double[NODEMAX];
            int i, j, k, ix;							//See the Factor_Table subroutine description
            double Temp1, Temp2;						//Temporary work unit

            ix = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)		//The previous generation operation begins
            {
                Temp1 = Power_Dis_Correct[i];			//Send the right end amount to the temporary work unit
                for (k = 0; k < Num_Fact_UpTri[i, IterFlag]; k++)
                {
                    j = Foot_Fact_UpTri[ix, IterFlag];
                    Temp2 = Temp1 * Fact_UpTri[ix, IterFlag];
                    Power_Dis_Correct[j] = Power_Dis_Correct[j] - Temp2;
                    ix++;
                }
                Power_Dis_Correct[i] = Temp1 * Fact_Diag[i, IterFlag];
            }
            for (i = Num_Node - Num_Swing - 1; i >= 0; i--)	//Back to generation
            {
                Temp1 = Power_Dis_Correct[i];
                for (k = 0; k < Num_Fact_UpTri[i, IterFlag]; k++)
                {
                    ix--;
                    j = Foot_Fact_UpTri[ix, IterFlag];
                    Temp2 = Power_Dis_Correct[j] * Fact_UpTri[ix, IterFlag];
                    Temp1 = Temp1 - Temp2;
                }
                Power_Dis_Correct[i] = Temp1;				//The solution to the unknown
            }
        }


        double[,] Voltage = new double[NODEMAX, 2];						//Node voltage: 0-phase angle;
        //1-Amplitude.
        double[,] Current_Const = new double[SWINGMAX * NODEFACTOR, 2];	//Constant current: 0-real part;
        //1-The imaginary part.
        int[] Node_Name_Curr_Const = new int[SWINGMAX * NODEFACTOR];	//Constant current node name (number)
        double[,] Power_Const = new double[NODEMAX, 2];					//Injecting power invariant parts of each node
        //Points: 0 - real part; 1 - imaginary part。

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
                    Voltage[i, 0] = 0.0;				//Initial value of voltage phase: 0.0 (radian)
                    Voltage[i, 1] = 1.0;				//Initial value of voltage amplitude: 1.0
                }
            }
            TemporaryCulture.Stop();
            stream.Close();
            input.Close();
        }


        //Initialization subroutine
        void Initial(int Num_Node, int Num_Line, int Num_Load, int Num_Swing,
                     int Num_GPV, int Num_GPQ, out int Num_Cur_Const,
                     double[,] DVolt)
        {
            Num_Cur_Const = 0;
            int i, j, jl, jr, k, kk;
            int Flag;
            int kg_old, kl_old;					//Generator, load old sequence number work unit            
            int kl; // load count temporary work unit
            int kgpv; //Generator PV node count temporary work unit
            int kgpq; //Generator PQ node count temporary work unit
            Double R, X, Z, Ang; //Line parameters and balance node voltage phase angle
            //Time unit.
            Double[] yij, V_Temp, I_Temp; //temporary unit of work
            yij = new double[2];
            V_Temp = new double[2];
            I_Temp = new double[2];
            //for (i = 0; i < Num_Node - Num_Swing; i++)
            //{
            //    DVolt[i, 0] = 0.0;				//The voltage phasor change amount is given the initial value of 0.0
            //    DVolt[i, 1] = 0.0;
            //    Power_Const[i, 0] = 0.0;			//Injecting power initial value of each node
            //    Power_Const[i, 1] = 0.0;
            //}
            Array.Clear(DVolt, 0, 2 * (Num_Node - Num_Swing));//
            Array.Clear(Power_Const, 0, 2 * (Num_Node - Num_Swing));//
            //else if(VolIni_Flag==1)
            //if (VolIni_Flag == 1) Voltage_Initial(Num_Node, Num_Swing);
            for (kgpv = 0; kgpv < Num_GPV; kgpv++)		//Generator PV node voltage amplitude = VG
            {
                i = Gen_PVNode[kgpv, 0];
                kg_old = Gen_PVNode[kgpv, 1];
                Voltage[i, 1] = GGen[kg_old].PQV[1];
            }

            for (i = 0; i < Num_Line; i++)				//Constant current information solution
            {
                jl = Line_NodeName[i, 0];			//Line left node
                jr = Line_NodeName[i, 1];			//Line right node
                if (jl < Num_Node - Num_Swing
                    && jr >= Num_Node - Num_Swing)	//Jl is a normal node and jr is a balanced node.
                {
                    k = Line_No_NewtoOld[i];		//Corresponding old line number
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
                    yij[1] = X;				//So far, the branch admittance is divided by the non-standard ratio

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
                    }						//So far, the real part of the corresponding equilibrium node voltage is obtained.

                    Flag = 0;
                    for (j = 0; j < Num_Cur_Const; j++)
                    {
                        if (Node_Name_Curr_Const[j] == jl) Flag = 1;
                        if (Flag == 1) break;
                    }
                    if (Flag == 0)					//New constant current node
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
                    else						//This constant current node has appeared
                    {
                        Comp_Mul(out I_Temp[0], out I_Temp[1], yij[0], yij[1], V_Temp[0], V_Temp[1]);
                        Current_Const[j, 0] = Current_Const[j, 0] + I_Temp[0];
                        Current_Const[j, 1] = Current_Const[j, 1] + I_Temp[1];
                    }
                }
            }
            //Output constant current node data

            //Each node injection power constant part evaluation
            kgpv = 0;
            kgpq = 0;
            kl = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
            {
                if (kgpv < Num_GPV && i == Gen_PVNode[kgpv, 0])	//Generator PV node
                {
                    kg_old = Gen_PVNode[kgpv, 1];				//Generator old serial number
                    if (kl < Num_Load && i == Load_NodeName[kl])	//Load node
                    {
                        kl_old = Load_No_NewtoOld[kl];		//Load old sequence number
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0]
                            - LLoad[kl_old].ABC[4];			//Active part plus PG and C1
                        kl++;
                    }
                    else									//Non-load node
                    {
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0];			//Active part added to PG
                    }
                    kgpv++;
                }
                else if (kgpq < Num_GPQ && i == Gen_PQNode[kgpq, 0])	//Generator PQ node
                {
                    kg_old = Gen_PQNode[kgpq, 1];				//Generator old serial number
                    if (kl < Num_Load && i == Load_NodeName[kl])	//Load node
                    {
                        kl_old = Load_No_NewtoOld[kl];		//Load old sequence number
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0]
                            - LLoad[kl_old].ABC[4];			//Active part plus PG and C1
                        Power_Const[i, 1] = Power_Const[i, 1]
                            + GGen[kg_old].PQV[1]
                            - LLoad[kl_old].ABC[5];			//Reactive part plus QG and C2
                        kl++;
                    }
                    else									//Non-load node
                    {
                        Power_Const[i, 0] = Power_Const[i, 0]
                            + GGen[kg_old].PQV[0];			//Active part added to PG
                        Power_Const[i, 1] = Power_Const[i, 1]
                            + GGen[kg_old].PQV[1];			//Reactive part joins QG
                    }
                    kgpq++;
                }
                else					//Neither the generator PV node nor the generator PQ node
                {
                    if (kl < Num_Load && i == Load_NodeName[kl])	//Load node
                    {
                        kl_old = Load_No_NewtoOld[kl];		//Load old sequence number
                        Power_Const[i, 0] = Power_Const[i, 0]
                            - LLoad[kl_old].ABC[4];			//Active part added to C1
                        Power_Const[i, 1] = Power_Const[i, 1]
                            - LLoad[kl_old].ABC[5];			//C2
                        kl++;
                    }
                }
            }
            //Each node injection power constant part processing result output: new node name (number), active power, reactive power

            
            //The heading output of the power mismatch output section
            // Open the mismatch output disk file
               Console.WriteLine(String.Format("{0,15:D}", "Iterating No")+
               String.Format("{0,15:D}", "P_Dis_Max") + String.Format("{0,6:D}", "Node") +
               String.Format("{0,15:D}", "Q_Dis_Max") + String.Format("{0,6:D}", "Node"));
            Console.WriteLine(String.Format("{0,15:D}", "==============") +
               String.Format("{0,15:D}", "==============") + String.Format("{0,6:D}", "=====") +
               String.Format("{0,15:D}", "==============") + String.Format("{0,6:D}", "====="));
        }

        //Find the node power mismatch quantum path (dQi=0 of PV node)
        void PowerDis_Comp(int Num_Node, int Num_Load, int Num_Swing, int Num_GPV,
                           int Num_Cur_Const, double[,] Power_Dis,
                           double[,] Pij_Sum, double[,] DVolt,
                           int Num_Iteration,
                           out double Power_Dis_Max)
        {
            int i, j, k, li, kl_old, kl, kgpv, ki, k1;
            double V, Ang;					//Node i voltage amplitude and phase
            double[] VV = new double[2];					//Node i voltage real and imaginary parts
            double[] V_Temp = new double[2];				//Node voltage temporary working unit (real and imaginary)
            double[] Cur_Count = new double[2];			//Node voltage temporary working unit (real and imaginary)
            double[] Cur_Temp = new double[2];				//Injection current temporary working unit (real and imaginary)
            double Ix, Iy;					//Current (real and imaginary)
            double PP, QQ;					//Active, reactive
            int ipmax, iqmax;				//The node number corresponding to the maximum active power and reactive mismatch
            ipmax = 0; iqmax = 0;
            double P_Dis_Max, Q_Dis_Max;		//Maximum active and reactive mismatch

            kl = 0;
            kgpv = 0;
            ki = 0;
            for (i = 0; i < Num_Node - Num_Swing; i++)
            {
               
                Power_Dis[i, 0] = Power_Const[i, 0]; //Inject the node i into the power constant
                Power_Dis[i, 1] = Power_Const[i, 1]; // Feed the mismatch unit.
                Ang = Voltage[i, 0]; //node i voltage phase after last iteration
                V = Voltage[i, 1]; // and amplitude.
                VV[0] = V * Math.Cos(Ang); //The real part of the node i voltage after the last iteration
                VV[1] = V * Math.Sin(Ang); // and imaginary parts.
                //Follow the sum of Pij and Qij and all Pij and Qij from node i
                Cur_Count[0] = 0.0;
                Cur_Count[1] = 0.0;
                for (k = No_First_Y_DownTri[i];
                    k < No_First_Y_DownTri[i + 1]; k++) //The lower i-th row is a non-zero loop
                {
                    j = Foot_Y_DownTri[k]; //lower i-th row of the current non-zero-element column code
                    V_Temp[0] = Voltage[j, 1] * Math.Cos(Voltage[j, 0]);
                    V_Temp[1] = Voltage[j, 1] * Math.Sin(Voltage[j, 0]);
                    li = No_Y_DownTri_RowtoCol[k]; // Corresponding compressed storage sequence by column
                    // The sequence number of the same non-zero element.

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
                Cur_Count[1] = -Cur_Count[1]; //current conjugate
                Comp_Mul(out Pij_Sum[i, 0], out Pij_Sum[i, 1],
                    VV[0], VV[1], Cur_Count[0], Cur_Count[1]); //At this point, find the location from node i
                //The sum of Pij and Qij.

                if (kgpv < Num_GPV && i == Gen_PVNode[kgpv, 0]) //Generator PV node
                {
                    if (kl < Num_Load && i == Load_NodeName[kl]) //load node
                    {
                        kl_old = Load_No_NewtoOld[kl]; //load old sequence number
                        if (LLoad[kl_old].Flag == 1) // take into account the static characteristics of the load
                            Power_Dis[i, 0] = Power_Dis[i, 0]
                            - LLoad[kl_old].ABC[2] * V;
                        kl++;
                    }
                    if (ki < Num_Cur_Const && i == Node_Name_Curr_Const[ki])
                    {										//Constant current node
                        Power_Dis[i, 0] = Power_Dis[i, 0]
                            + V * (Current_Const[ki, 0] * Math.Cos(Ang)
                            + Current_Const[ki, 1] * Math.Sin(Ang));
                        ki++;
                    }
                    Power_Dis[i, 0] = Power_Dis[i, 0] - Pij_Sum[i, 0];
                    kgpv++;
                }
                else					
                                //PQ node (including generator, load and contact node)
                {
                    if (kl < Num_Load && i == Load_NodeName[kl]) //load node
                    {
                        kl_old = Load_No_NewtoOld[kl]; //load old sequence number
                        if (LLoad[kl_old].Flag == 1) // take into account the static characteristics of the load
                        {
                            Power_Dis[i, 0] = Power_Dis[i, 0]
                                - LLoad[kl_old].ABC[2] * V;
                            Power_Dis[i, 1] = Power_Dis[i, 1]
                                - LLoad[kl_old].ABC[3] * V;
                        }
                        kl++;
                    }
                    if (ki < Num_Cur_Const && i == Node_Name_Curr_Const[ki])
                    {										//Constant current node
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
            //Node power primary deviation correction term
            for (k = 0; k < Num_Cur_Const; k++)			//Constant current correction term processing
            {
                i = Node_Name_Curr_Const[k];			//Node number
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

            for (i = 0; i < Num_Node - Num_Swing; i++)    //The sum of Pij and Qij of all outgoing lines of the i-node
            {                                             //Revision item processing.
                V = Voltage[i, 1];
                PP = Pij_Sum[i, 1] * DVolt[i, 0] - Pij_Sum[i, 0] * DVolt[i, 1];
                QQ = -Pij_Sum[i, 0] * DVolt[i, 0] - Pij_Sum[i, 1] * DVolt[i, 1];
                //		Power_Dis[i,0]=Power_Dis[i,0]-PP/V;
                //		if(Node_Flag[i]!=2)Power_Dis[i,1]=Power_Dis[i,1]-QQ/V;
                Power_Dis[i, 0] = Power_Dis[i, 0] - PP * 0.1;
                if (Node_Flag[i] != 2) Power_Dis[i, 1] = Power_Dis[i, 1] - QQ * 0.1;
            }

            for (k = 0; k < Num_Load; k++)					//Correction of load static characteristics
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

        //Voltage save (by internal node number) subroutine
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


        //Result output subroutine
        void Result_Output(int Num_Node, int Num_Line, int Num_Gen, int Num_Load,
                           int Num_Swing, int Num_Iteration, double Duration,string fileName)
        {
            int i, j, k, kg_old, kl_old, k_old;
            double BK;					//Line parameter temporary work unit
            double[] Z = new double[2];
            double[,] S_Count = new double[NODEMAX, 2];		//All the outlet power accumulation arrays of the nodes:
            //0-Real part, 1- imaginary part.
            double[,] S_Line = new double[LINEMAX, 4];		//Line power: 0-Pij, 1-Qij, 2-Pji, 3-Qji
            double[,] S_Node = new double[NODEMAX, 4];		//Node power: 0-PG, 1-QG, 2-PL, 3-QL
            double[,] DS_Node = new double[NODEMAX, 2];		//Node power mismatch: 0 - active mismatch,
            //1- Reactive mismatch.
            double[] S_T, V_T, I_T, I1_T, I2_T;	//Temporary work unit
            S_T = new double[2];
            V_T = new double[2];
            I_T = new double[2];
            I1_T = new double[2];
            I2_T = new double[2];

            double V, t, Angle, Vi, Vj, Angi, Angj;

            //Send the voltage value of the balanced node to the array Voltage[,]
            for (i = 0; i < Num_Swing; i++)
            {
                j = Gen_SWNode[i, 0];
                kg_old = Gen_SWNode[i, 1];
                Angle = GGen[kg_old].PQV[1] * Deg_to_Rad;
                Voltage[j, 0] = Angle;
                Voltage[j, 1] = GGen[kg_old].PQV[0];
            }		//At this point, you can directly use the phase and amplitude of all node voltages in the system.

            
            Array.Clear(S_Count, 0, Num_Node * 2);//
            Array.Clear(S_Node, 0, 4 * Num_Node);
            //Find the line trend and all the outgoing power of each node
            for (k = 0; k < Num_Line; k++)
            {
                
                i = Line_NodeName[k, 0]; // take the left node of the line
                j = Line_NodeName[k, 1]; // take the right node of the line
                k_old = Line_No_NewtoOld[k]; // corresponding old line sequence number
                Z[0] = LLine[k_old].RXBK[0]; //Receive the resistance value of the line
                Z[1] = LLine[k_old].RXBK[1]; //Receive the reactance value of the line
                BK = LLine[k_old].RXBK[2]; // Take the line to accommodate the half value or the transformer non-standard ratio
                if (Line_Flag[k] == 0) // ordinary branch
                {
                    if (i != j) //Ungrounded branch
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
                        I1_T[1] = -I1_T[1];						//Take the conjugate of Iij
                        I2_T[0] = -I_T[0] - BK * Vj * Math.Sin(Angj);		//Iji
                        I2_T[1] = -I_T[1] + BK * Vj * Math.Cos(Angj);
                        I2_T[1] = -I2_T[1];						//Take the conjugate of Iji
                        V_T[0] = Vi * Math.Cos(Angi);
                        V_T[1] = Vi * Math.Sin(Angi);
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);					//Seeking Sij
                        S_Line[k, 0] = S_T[0];
                        S_Line[k, 1] = S_T[1];
                        S_Count[i, 0] = S_Count[i, 0] + S_T[0];	//Node i outgoing
                        S_Count[i, 1] = S_Count[i, 1] + S_T[1];	//The power is accumulated.
                        V_T[0] = Vj * Math.Cos(Angj);
                        V_T[1] = Vj * Math.Sin(Angj);
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]);					//Seeking Sji
                        S_Line[k, 2] = S_T[0];
                        S_Line[k, 3] = S_T[1];
                        S_Count[j, 0] = S_Count[j, 0] + S_T[0];	//Node j outgoing
                        S_Count[j, 1] = S_Count[j, 1] + S_T[1];	//The power is accumulated.
                    }
                    else					//Grounding branch
                    {
                        Vi = Voltage[i, 1];
                        Angi = Voltage[i, 0];
                        V_T[0] = Vi * Math.Cos(Angi);
                        V_T[1] = Vi * Math.Sin(Angi);
                        Comp_Div(I_T, V_T, Z);
                        I1_T[0] = I_T[0] - BK * Vi * Math.Sin(Angi);	//Find the grounding branch current Iii
                        I1_T[1] = I_T[1] + BK * Vi * Math.Cos(Angi);
                        I1_T[1] = -I1_T[1];				//Take the conjugate of Iii
                        Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]);
                        S_Line[k, 0] = S_T[0];		//Seeking Sii (from node i to the ground node)
                        S_Line[k, 1] = S_T[1];
                        S_Line[k, 2] = 0.0; // Find Sii (from ground node to node i, its value is equal to zero)
                        S_Line[k, 3] = 0.0;
                        S_Count[i, 0] = S_Count[i, 0] + S_T[0]; //node i outgoing
                        S_Count[i, 1] = S_Count[i, 1] + S_T[1]; // Power accumulation.
                    }
                }
                else if (Line_Flag[k] == 1) // Transformer branch with non-standard ratio on node i side
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
                    I1_T[1] = -I1_T[1];							//Take the conjugate of Iij
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Div(I2_T, V_T, Z);
                    t = 1.0 - 1.0 / BK;
                    I2_T[0] = I2_T[0] * t - I_T[0];					//Iji
                    I2_T[1] = I2_T[1] * t - I_T[1];
                    I2_T[1] = -I2_T[1]; //take the conjugate of Iji
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]); //See Sij
                    S_Line[k, 0] = S_T[0];
                    S_Line[k, 1] = S_T[1];
                    S_Count[i, 0] = S_Count[i, 0] + S_T[0]; //node i outgoing
                    S_Count[i, 1] = S_Count[i, 1] + S_T[1]; // Power accumulation.
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]); //Seeking Sji
                    S_Line[k, 2] = S_T[0];
                    S_Line[k, 3] = S_T[1];
                    S_Count[j, 0] = S_Count[j, 0] + S_T[0];		//node j outgoing
                    S_Count[j, 1] = S_Count[j, 1] + S_T[1]; //Power accumulation.
                }
                else if (Line_Flag[k] == 2) // Transformer branch with non-standard ratio on node j side
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
                    I1_T[1] = -I1_T[1]; // take the conjugate of Iij
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Div(I2_T, V_T, Z);
                    t = (1.0 / BK - 1.0) / BK;
                    I2_T[0] = I2_T[0] * t - I_T[0]; //Iji
                    I2_T[1] = I2_T[1] * t - I_T[1];
                    I2_T[1] = -I2_T[1]; //take the conjugate of Iji
                    V_T[0] = Vi * Math.Cos(Angi);
                    V_T[1] = Vi * Math.Sin(Angi);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I1_T[0], I1_T[1]); //See Sij
                    S_Line[k, 0] = S_T[0];
                    S_Line[k, 1] = S_T[1];
                    S_Count[i, 0] = S_Count[i, 0] + S_T[0]; //node i outgoing
                    S_Count[i, 1] = S_Count[i, 1] + S_T[1]; // Power accumulation.
                    V_T[0] = Vj * Math.Cos(Angj);
                    V_T[1] = Vj * Math.Sin(Angj);
                    Comp_Mul(out S_T[0], out S_T[1], V_T[0], V_T[1], I2_T[0], I2_T[1]); //Seeking Sji
                    S_Line[k, 2] = S_T[0];
                    S_Line[k, 3] = S_T[1];
                    S_Count[j, 0] = S_Count[j, 0] + S_T[0]; //node j outgoing
                    S_Count[j, 1] = S_Count[j, 1] + S_T[1]; //Power accumulation.
                }
                else // no such line type
                {
                    Console.WriteLine("There is not this line type!");
                    Environment.Exit(1);
                }
            }

            
            // Add the initial value of the node power mismatch to the negative value of the node outlet power accumulation
            for (i = 0; i < Num_Node; i++)
            {
                DS_Node[i, 0] = -S_Count[i, 0];
                DS_Node[i, 1] = -S_Count[i, 1];
            }

            // Find the load node power PL and QL
            for (i = 0; i < Num_Load; i++)
            {
                j = Load_NodeName[i]; //Load node name
                kl_old = Load_No_NewtoOld[i]; //Load node old sequence number
                if (LLoad[kl_old].Flag == 1) // take into account static characteristics
                {
                    V = Voltage[j, 1];
                    S_Node[j, 2] = LLoad[kl_old].ABC[0] * V * V
                        + LLoad[kl_old].ABC[2] * V
                        + LLoad[kl_old].ABC[4];
                    S_Node[j, 3] = LLoad[kl_old].ABC[1] * V * V
                        + LLoad[kl_old].ABC[3] * V
                        + LLoad[kl_old].ABC[5];
                }
                
                else if (LLoad[kl_old].Flag == 0) // does not count static characteristics
                {
                    S_Node[j, 2] = LLoad[kl_old].ABC[4];
                    S_Node[j, 3] = LLoad[kl_old].ABC[5];
                }
                else // there is no such type of load node
                {
                    Console.WriteLine("There is not this load node type!");
                    Environment.Exit(1);
                }
                DS_Node[j, 0] = DS_Node[j, 0] - S_Node[j, 2]; //Power mismatch processing
                DS_Node[j, 1] = DS_Node[j, 1] - S_Node[j, 3];
            }

            
            // Find generator node power PG and QG
            for (i = 0; i < Num_Gen; i++)
            {
                j = Gen_NodeName[i]; //Generator node name
                kg_old = Gen_No_NewtoOld[i]; // Generator node old sequence number
                if (GGen[kg_old].Flag == 0) //balance generator node
                {
                    S_Node[j, 0] = S_Count[j, 0] + S_Node[j, 2];
                    S_Node[j, 1] = S_Count[j, 1] + S_Node[j, 3];
                }
                else if (GGen[kg_old].Flag == 1) //Generator PQ node
                {
                    S_Node[j, 0] = GGen[kg_old].PQV[0];
                    S_Node[j, 1] = GGen[kg_old].PQV[1];
                }                
                else if (GGen[kg_old].Flag == 2) //Generator PV node
                {
                    S_Node[j, 0] = GGen[kg_old].PQV[0];
                    S_Node[j, 1] = S_Count[j, 1] + S_Node[j, 3];
                }
                else // no such type of generator node
                {
                    Console.WriteLine("There is not this generator node type!");
                    Environment.Exit(1);
                }
                DS_Node[j, 0] = DS_Node[j, 0] + S_Node[j, 0]; //Power mismatch processing
                DS_Node[j, 1] = DS_Node[j, 1] + S_Node[j, 1];
            }

            //Node power flow output: sequence number, node number, node type, V, Angle, PG, QG, PL, QL, node active mismatch
            // Quantity, node reactive mismatch.
            int[] Node_Name_Old = new int[NODEMAX]; //node old name (number)
            int[] Node_No_OldtoNew = new int[NODEMAX]; //New node number corresponding to the old node number after sorting
            int I_Temp, np; //temporary variable
            int[] Node_Name_Voltage_MaxMin = new int[4]; //voltage amplitude minimum, maximum and voltage
            //The smallest phase, the largest node number.
            Double[] Voltage_Value_MaxMin = new double[4]; //minimum, maximum voltage amplitude and minimum,
            //Maximum voltage phase.
            for (i = 0; i < Num_Node; i++)
            {
                Node_Name_Old[i] = Node_Name_NewtoOld[i];
                Node_No_OldtoNew[i] = i;
            }
            for (i = 0; i < Num_Node - 1; i++) // sort by the old name of the node from small to large
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
            //Screen output section
            Console.WriteLine("");
            Console.WriteLine(String.Format("{0,50:D}", "Node Flow Output"));
            Console.WriteLine(String.Format("{0,50:D}", "================"));
            Console.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-4:D}", "Bus")+
                String.Format("{0,-5:D}", "Type") + String.Format("{0,-8:D}", "Voltage")+
                String.Format("{0,-8:D}", "Angle") + String.Format("{0,-8:D}", "PG")+
                String.Format("{0,-8:D}", "QG") + String.Format("{0,-8:D}", "PL")+
                String.Format("{0,-8:D}", "QL") + String.Format("{0,-10:D}", "Bus_P_Dis")+
                String.Format("{0,-10:D}", "Bus_Q_Dis"));
            //Disk output section
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
                
                j = Node_No_OldtoNew[i]; //The new node sequence number corresponding to the current old node number
                k = Node_Name_NewtoOld[j];//The old node name (number) corresponding to the new node number
                V = Voltage[j, 1];
                Angle = Voltage[j, 0] * Rad_to_Deg;
                // screen output section
                Console.WriteLine(String.Format("{0,-3:D}", i) + String.Format("{0,-4:D}", k) +
                String.Format("{0,-5:D}", Node_Flag[j]) + String.Format("{0,-8:f4}", V) +
                String.Format("{0,-8:f2}", Angle) + String.Format("{0,-8:f4}", S_Node[j, 0]) +
                String.Format("{0,-8:f4}", S_Node[j, 1]) + String.Format("{0,-8:f4}", S_Node[j, 2]) +
                String.Format("{0,-8:f4}", S_Node[j, 3]) + String.Format("{0,-10:f6}", DS_Node[j, 0]) +
                String.Format("{0,-10:f6}", DS_Node[j, 1]));

                //Disk output section
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

            
            //Line trend output: serial number, left node, right node, line type, Pij, Qij, Pji, Qji, active line loss,
            // reactive line loss.
            int i_old, j_old; //The left and right nodes of the line are old
            int[,] Node_Line = new int[LINEMAX, 2]; //Line left and right nodes old working array
            int[] Type_Exchange = new int[LINEMAX]; //Line left and right node old number exchange flag:
            //0- no exchange, 1-exchange.
            int[] Line_No_OldtoNew = new int[LINEMAX]; //The old node double-sorted line number to the new line number
            // conversion array.
            int kk, Line_Type;
            Double DPLOSS = 0, DQLOSS = 0; // total active and reactive power loss of the system
            for (k = 0; k < Num_Line; k++)
            {
                i = Line_NodeName[k, 0]; // line left node new number
                j = Line_NodeName[k, 1]; // line right node new number
                i_old = Node_Name_NewtoOld[i]; //The left number of the line left node
                j_old = Node_Name_NewtoOld[j]; // line right node old number
                if (i_old <= j_old) // line left node old number <= line right node old number, no exchange
                {
                    Node_Line[k, 0] = i_old;
                    Node_Line[k, 1] = j_old;
                    Type_Exchange[k] = 0;
                }
                else				
                //Line left node old number> Line right node old number, exchange
                {
                    Node_Line[k, 0] = j_old;
                    Node_Line[k, 1] = i_old;
                    Type_Exchange[k] = 1;
                }
                Line_No_OldtoNew[k] = k; //The old node double-sorted line number to the new line number
                // Change the initial value of the array
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

            //Screen output section

            Console.WriteLine(String.Format("{0,50:D}", "Line Flow Output"));
            Console.WriteLine(String.Format("{0,50:D}", "================"));
            Console.WriteLine(String.Format("{0,-3:D}", "No") + String.Format("{0,-5:D}", "LBus") +
                String.Format("{0,-5:D}", "RBus") + String.Format("{0,-5:D}", "Type") +
                String.Format("{0,-9:D}", "Pij") + String.Format("{0,-9:D}", "Qij") +
                String.Format("{0,-9:D}", "Pji") + String.Format("{0,-9:D}", "Qji") +
                String.Format("{0,-9:D}", "QL") + String.Format("{0,-10:D}", "P_Loss") +
                String.Format("{0,-10:D}", "Q_Loss"));
            //Disk output section
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
                kk = Line_No_OldtoNew[k];			
                //  Corresponding new line number
                Line_Type = Line_Flag[kk]; //Line type flag
                if (Type_Exchange[k] == 1) //The left and right nodes of the line have been exchanged, and they need to be changed.
                { // The voltage divider line type flag is modified.
                    if (Line_Type == 1) Line_Type = 2;
                    else if (Line_Type == 2) Line_Type = 1;
                    // screen output section
                    Console.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
            String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
            String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
            String.Format("{0,-10:f4}", S_Line[kk, 3]));

                    //disk output section
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
                    //Screen output section
                    Console.WriteLine(String.Format("{0,-3:D}", k) + String.Format("{0,-5:D}", Node_Line[k, 0]) +
            String.Format("{0,-5:D}", Node_Line[k, 1]) + String.Format("{0,-5:D}", Line_Type) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-9:f4}", S_Line[kk, 1]) +
            String.Format("{0,-9:f4}", S_Line[kk, 2]) + String.Format("{0,-9:f4}", S_Line[kk, 3]) +
            String.Format("{0,-9:f4}", S_Line[kk, 0]) + String.Format("{0,-10:f4}", S_Line[kk, 2]) +
            String.Format("{0,-10:f4}", S_Line[kk, 3]));
                    //Disk output section
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

            // System overall performance indicators output: active network loss, reactive power loss, minimum voltage value and its node name, the highest voltage
            //value and its node name, minimum voltage phase angle and its node name, maximum voltage phase angle and its node name, number of convergence,
            //calculating time.

            // screen output section
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

            //Disk output section
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


        //Convert angle to interval [-PAI, PAI]
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

            short Num_Node = 0; // total number of nodes
            short Num_Line = 0; // total number of lines
            short Num_Gen = 0; // total number of generators
            short Num_Load = 0; // total load
            short Num_Swing = 0; // total balancer node number
            short Num_GPV = 0; //The total number of PV nodes in the generator node
            short Num_GPQ = 0; //The total number of PQ nodes in the generator node
            int Num_Cur_Const = 0; //The total number of constant current nodes
            short Iter_Max; //maximum number of iterations
            short Num_Iteration = 0; //number of iterations
            short VolIni_Flag; //Read voltage initial value flag: 1-read; 0-not read
            short VolRes_Flag; //Retain voltage (as the next initial value) flag
            //: 1-reserved; 0-not reserved.
            float Eps = 1e-5f; //node power mismatch magnitude convergence limit
            double Power_Dis_Max; //Maximum node power mismatch value
            double Duration; //Store calculation time (s)

            double[,] Power_Dis = new double[NODEMAX, 2]; //power mismatch dP, dQ
            double[,] Pij_Sum = new double[NODEMAX, 2]; //All the outlet power of the node is accumulated
            double[] Power_Dis_Correct = new double[NODEMAX]; //dP/V(dSita.V), dQ/V(dV)
            double[,] DVolt = new double[NODEMAX, 2]; //Voltage correction amount dSita.V, dV

            Data_Input(out Num_Line, out Num_Gen, out Num_Load, out Eps, out Iter_Max, out VolIni_Flag, out VolRes_Flag, path); //Data input
            DateTime BeginTime = DateTime.Now;
            Node_Sequen(out Num_Node, Num_Line, Num_Gen, Num_Load, out Num_Swing, out Num_GPV, out Num_GPQ); //Serial number processing
            Y_Bus1(Num_Node, Num_Line, Num_Swing); //First admittance array
            Factor_Table(Num_Node, Num_Swing, Num_GPV, 0); //Form the first factor table

            Y_Bus2(Num_Node, Num_Line, Num_Load, Num_Swing); //Second admittance array
            Factor_Table(Num_Node, Num_Swing, Num_GPV, 1); //Form the second factor table
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
                    Voltage[m, 0] = 0.0; //The initial value of the voltage phase: 0.0 (radian)
                    Voltage[m, 1] = 1.0; //The initial value of the voltage amplitude: 1.0
                }
            Initial(Num_Node, Num_Line, Num_Load, Num_Swing,
                Num_GPV, Num_GPQ, out Num_Cur_Const, DVolt); //Initialize
        Iteration_Back: //iteration start point
            PowerDis_Comp(Num_Node, Num_Load, Num_Swing,
                          Num_GPV, Num_Cur_Const, Power_Dis,
                          Pij_Sum, DVolt,
                          Num_Iteration, out Power_Dis_Max); // Calculate the power mismatch
            if (Power_Dis_Max > Eps && Num_Iteration < Iter_Max) //Convergence judgment
            {
                for (i = 0; i < Num_Node - Num_Swing; i++)
                    Power_Dis_Correct[i] = Convert.ToSingle(
                    Power_Dis[i, 0] / Voltage[i, 1]); //take dP/V
                Equ_Calculation(Num_Node, Num_Swing, Power_Dis_Correct, 0);
                //Active solution

                for (i = 0; i < Num_Node - Num_Swing; i++)
                {
                    Voltage[i, 0] = Voltage[i, 0]- Power_Dis_Correct[i] / Voltage[i, 1]; //Correct phase
                    TreatAngle(ref Voltage[i, 0]);
                    DVolt[i, 0] = Power_Dis_Correct[i]; //Save phase difference dSita.V
                    Power_Dis_Correct[i] =
                        Power_Dis[i, 1] / Voltage[i, 1]; //take dQ/V
                }
                Equ_Calculation(Num_Node, Num_Swing, Power_Dis_Correct, 1);
                //Reactive solution

                for (i = 0; i < Num_Node - Num_Swing; i++)
                {
                    Voltage[i, 1] = Voltage[i, 1]
                        - Power_Dis_Correct[i]; //Correct amplitude
                    DVolt[i, 1] = Power_Dis_Correct[i]; //Save the amplitude difference dV
                }
                for (kgpv = 0; kgpv < Num_GPV; kgpv++) //voltage amplitude of the PV node of the generator = VG
                {
                    i = Gen_PVNode[kgpv, 0];
                    kg_old = Gen_PVNode[kgpv, 1];
                    Voltage[i, 1] = GGen[kg_old].PQV[1];
                    DVolt[i, 1] = 0.0f; //PV node voltage amplitude difference = 0.0
                }
                Num_Iteration++; // number of iterations increased by 1
                goto Iteration_Back; //iteration of the return point
            }
            //DateTime BeginTime = DateTime.Now;
            TimeSpan span = DateTime.Now.Subtract(BeginTime);
            Duration = span.TotalSeconds;
            //Console.WriteLine("Complete calculation! Shared:" + Duration.ToString() + "millisecond!");
            if (VolRes_Flag == 1)
            {
                fileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "vol.ini";
                if (!File.Exists(fileName))
                {
                    sw = File.CreateText(fileName);
                    sw.Close();
                }
                Voltage_Reserve(Num_Node, Num_Swing, fileName); //Save voltage (by internal number)
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
                            Num_Swing, Num_Iteration, Duration, fileName);	//Result output
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

    #region  Input data structure definition
    class Line              // Line parameter structure definition
    {
        private short[] _Node_No = new short[2]; //node name (number) at both ends of the line: 0-left node; 1-right node
        private byte _Flag; //Line type flag: 0-common branch; 1, 2-transformer branch. among them,	the 
                            //1-non-standard ratio is on the left node, and the 2-non-standard ratio is on the right node.
                             // For the ground branch, the left and right node names (numbers) are the same.
        private float[] _RXBK = new float[3];       //0-R;1-X;2-Bc/2 or K

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

    class Generator     // Generator parameter structure definition
    {
        private short _Node_No;//generator node name (number)
        private byte _Flag;//generator node type flag: 0-balance node; 1-PQ node; 2-PV node
        private float[] _PQV = new float[2];//for balanced nodes, 0-V, 1-Angle; for PQ nodes, 0-P, 1-Q; for PV sections
                                            //Point, 0-P, 1-V.

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

    class Load         // Load parameter structure definition
    {
        private short _Node_No;//load node name (number)
        private byte _Flag; // load node static characteristic flag: 0 - does not count static characteristics; 1 - static characteristics.
        private float[] _ABC = new float[6];//PL=a1*V**V+b1*V+c1,QL=a2*V*V+b2*V+c2.
                                            //0-a1; 1-a2; 2-b1; 3-b2; 4-c1; 5-c2.

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
