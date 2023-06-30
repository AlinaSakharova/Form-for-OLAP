using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            GetDimensions();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void GetDimensions()
        {
            /*using (AdomdConnection mdConn = new AdomdConnection())
            {
                mdConn.ConnectionString = "Data Source=LAPTOP-TMCGM84D; Initial Catalog=MultidimensionalProject3";
                mdConn.Open();

                foreach (CubeDef cube in mdConn.Cubes) 
                {
                    if (cube.Type != CubeType.Cube) continue;

                    foreach (Dimension dimension in cube.Dimensions)
                    {
                        foreach (Hierarchy hierarhcy in dimension.Hierarchies)
                        {
                            if (!hierarhcy.Name.Contains("ID") && !hierarhcy.Name.Contains("Id") &&
                                !hierarhcy.Name.Contains("Measures")) checkedListBox1.Items.Add(hierarhcy.UniqueName
                                    + ".[" + hierarhcy.Name + "]"); 
                        }
                    }
                }
            }
            */
            // prepare adomd connection
            using (AdomdConnection mdConn = new AdomdConnection())
            {
                string conn = Properties.Settings.Default.connectionString;
                mdConn.ConnectionString = Properties.Settings.Default.connectionString;
                mdConn.Open();

                // перебор кубов
                foreach (CubeDef cube in mdConn.Cubes)
                {
                    if (cube.Type != CubeType.Cube) continue;

                    // перебор измерений
                    foreach (Dimension dimension in cube.Dimensions)
                    {
                        // перебор иерархий
                        foreach (Hierarchy hierarchy in dimension.Hierarchies)
                        {
                            if (!hierarchy.Name.Contains("Id") && !hierarchy.Name.Contains("Measures"))
                            {
                                checkedListBox1.Items.Add(hierarchy.UniqueName + ".[" + hierarchy.Name + "]");
                            }
                        }
                    }

                    foreach (Measure measure in cube.Measures)
                    {
                        lstMeasures.Items.Add(measure.UniqueName);
                    }
                }

            }
        }

         private void BuildQuery()
        {
            
        }


        private void checkedListBox1_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0) return;

            string query = "SELECT NON EMPTY {[Measures].[Вес Кг]} ON COLUMNS, NON EMPTY {(";

            foreach (string s in checkedListBox1.CheckedItems)
            {
                query += " " + s.ToString() + "ALLMEMBERS ";
            }
            query = query.Remove(query.Length - 2);

            query += " )} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Trucking] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

        }

        private void UpdateChart(string mdxQuery)
        {
            chart1.Series.Clear();    
            // prepare adomd connection
            using (AdomdConnection mdConn = new AdomdConnection())
            {
                mdConn.ConnectionString = Properties.Settings.Default.connectionString;
                mdConn.Open();
                AdomdCommand mdCommand = mdConn.CreateCommand();
                mdCommand.CommandText = mdxQuery;  // << MDX Query 

                // выполняем запрос, получаем CellSet
                CellSet cs = mdCommand.ExecuteCellSet();

                // our method supports only 2-Axes CellSets
                if (cs.Axes.Count != 2) return;

                TupleCollection tuplesOnColumns = cs.Axes[0].Set.Tuples;//меры
                TupleCollection tuplesOnRows = cs.Axes[1].Set.Tuples;//измерения

                // 2 дублирующиеся структуры данных для графиков и таблицы с данными
                var Data = new DataTable();
                var DataChart = new DataTable();
                List<KeyValuePair<string, int>> ChartData = new List<KeyValuePair<string, int>>();

                if (tuplesOnColumns.Count > 0 && tuplesOnRows.Count > 0)
                {
                    //попробовать программно добавить series
                    //chart1.Series.Add();
                    //chart1.Series["data"].XValueMember = tuplesOnRows[0].Members[0].ParentLevel.Name; //имена измерений
                    //chart1.Series["data"].YValueMembers = tuplesOnColumns[0].Members[0].Caption; //меры вес
                    treeView1.Nodes.Clear();

                    var node = treeView1.Nodes.Add("tuplesOnRows[" + tuplesOnRows.GetType().Name + "]");
                    TreeNode nodeCurrent, nodeSub, nodeSubSub;

                    foreach (var tuple in tuplesOnRows)
                    {
                        string name = tuple.ToString();

                        if (name.StartsWith("Microsoft.AnalysisServices."))
                        {
                            name = name.Substring(27);
                        }

                        nodeCurrent = node.Nodes.Add(name);

                        nodeSub = nodeCurrent.Nodes.Add("Members");
                        foreach (var member in tuple.Members)
                        {
                            nodeSubSub = nodeSub.Nodes.Add(member.Name);

                            nodeSubSub.Nodes.Add("Caption").Nodes.Add(member.Caption);
                            nodeSubSub.Nodes.Add("ParentLevel.Name").Nodes.Add(member.ParentLevel.Name);
                        }
                    }

                    node = treeView1.Nodes.Add("tuplesOnColumns[" + tuplesOnColumns.GetType().Name + "]");

                    foreach (var tuple in tuplesOnColumns)
                    {
                        string name = tuple.ToString();

                        if (name.StartsWith("Microsoft.AnalysisServices."))
                        {
                            name = name.Substring(27);
                        }

                        nodeCurrent = node.Nodes.Add(name);

                        nodeSub = nodeCurrent.Nodes.Add("Members");
                        foreach (var member in tuple.Members)
                        {
                            nodeSubSub = nodeSub.Nodes.Add(member.Name);

                            nodeSubSub.Nodes.Add("Caption").Nodes.Add(member.Caption);
                            nodeSubSub.Nodes.Add("ParentLevel.Name").Nodes.Add(member.ParentLevel.Name);
                        }
                    }

                    foreach (TreeNode n in treeView1.Nodes)
                    {
                        foreach (TreeNode nn in n.Nodes)
                        {
                            nn.ExpandAll();
                        }
                    }

                    // Создаем заголовки в таблице на основе названий
                    for (int m = 0; m < tuplesOnRows[0].Members.Count; m++)

                    {
                        Data.Columns.Add(tuplesOnRows[0].Members[m].ParentLevel.Name);

                    }

                    for (int m = 0; m < tuplesOnColumns.Count; m++)
                    {
                        Data.Columns.Add(tuplesOnColumns[m].Members[0].Caption);


                        chart1.Series.Add(tuplesOnColumns[m].Members[0].Caption);
                        chart1.Series[m].XValueMember = tuplesOnRows[0].Members[0].ParentLevel.Name; 
                        chart1.Series[m].YValueMembers = tuplesOnColumns[m].Members[0].Caption; 
                    }

                    // Выводим строки с данными
                    for (int row = 0; row < tuplesOnRows.Count; row++)
                    {
                        int m = 0; string key = "";
                        // объект - для занесения в таблицу данных
                        var o = new object[tuplesOnRows[row].Members.Count + tuplesOnColumns.Count];

                        //проверить как заеполняется таблица
                        for (; m < tuplesOnRows[row].Members.Count; m++)
                        {
                            // заносим названия измерений для таблицы (по колонкам)
                            o[m] = tuplesOnRows[row].Members[m].Caption;
                            // создаем строку со списком измерений (для графиков)
                            key += tuplesOnRows[row].Members[m].Caption + " ";
                        }
                        // заносим непосредственно значения метрики (кол-во обращений) в последнюю колонку строки (для таблицы)
                        for (int col = 0; col < tuplesOnColumns.Count; col++)
                        {
                            ///
                            o[m + col] = cs.Cells[col, row].Value.ToString();
                            ChartData.Add(new KeyValuePair<string, int>(key, Convert.ToInt32(cs.Cells[col, row].Value)));

                        }
                        //добавляем строку таблицы
                        Data.Rows.Add(o);
                        //для графиков - добавляем в список новую пару "Ключ-значение", где ключ - имя строки со списком измерений, значение - кол-вол обращений
                       // ChartData.Add(new KeyValuePair<string, int>(key, Convert.ToInt32(cs.Cells[0, row].Value)));
                    }
                }
                // обновляем визуализацию данных
                try
                {
                    dataGridView1.DataSource = Data.DefaultView;
                    chart1.DataSource = Data.DefaultView;
                    chart1.DataBind();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }


        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0) return;
            if (lstMeasures.CheckedItems.Count == 0) return;

            // первая часть запроса: количсество обращений
            //string query = "SELECT NON EMPTY { [Measures].[Price all] } ON COLUMNS, NON EMPTY { ";

            string query = "SELECT NON EMPTY { ";

            foreach (var i in lstMeasures.CheckedItems)
            {
                query += i.ToString() + ", ";
            }
            query = query.Remove(query.Length - 2);
            query += " } ON COLUMNS, NON EMPTY { (";

            // перебираем все отмеченные измерения и добавляем их к запросу
            foreach (var i in checkedListBox1.CheckedItems)
            {
                query += " " + i.ToString() + ".ALLMEMBERS *";
            }

            // удаляем последний пробел со звездочкой
            query = query.Remove(query.Length - 2);

            //финальная часть запроса
            query += ") }  DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Trucking] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";
            UpdateChart(query);

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (AdomdConnection conn = new AdomdConnection())
            {
                conn.ConnectionString = Properties.Settings.Default.connectionString;
               // conn.Open();

                //AdomdCommand cmd = conn.CreateCommand();
                //conn.ConnectionString = "Data Source=LAPTOP-TMCGM84D; Initial Catalog=MultidimensionalProject3";
                conn.Open();
                AdomdCommand cmd = conn.CreateCommand();
                cmd.CommandText = "<Process xmlns=\"http://schemas.microsoft.com/analysisservices/2003/engine\">\r\n" +
                @"<Object>
                <DatabaseID>MultidimensionalProject3</DatabaseID>
                <CubeID>Trucking</CubeID>
                </Object>
                <Type>ProcessFull</Type>
                <WriteBackTableCreation>UseExisting</WriteBackTableCreation>
                </Process>";

                if (cmd.ExecuteNonQuery() == 1)
                {
                    MessageBox.Show("Куб успешно обновлен", "Обновление куба", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Не удалось обновить куб", "Обновление куба", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }

}

