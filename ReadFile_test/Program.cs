using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace ReadFile_test
{
    class Program
    {
        private static System.Data.DataTable datatable = new DataTable("datatable");
        static void Main(string[] args)
        {
            string path = System.Environment.CurrentDirectory;

            //验证路径及文件
            string filename = pathjudge(ref path);

            //从excel中读取并存入datatable中
            ReadFrom(path, filename);

            
        }

        private static string pathjudge(ref string path)
        {
            //显示当前目录
            Console.Write("当前路径为" + path + '\n');

            //判断是否存在该路径
            while (!Directory.Exists(@path))
            {
                path = Console.ReadLine();
            }
            Console.Write("请输入需要导入的文件名（包括文件名后缀）" + '\n');
            string filename;
            filename = Console.ReadLine();

            //判断文件是否在path目录下
            while (!File.Exists(path + "\\" + filename))
            {
                Console.Write("找不到文件！" + "\n" + "请重新输入！" + "\n");
                filename = Console.ReadLine();
            }

            return filename;
        }

        static void ReadFrom(string path,string file)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            //必须要使用绝对路径打开？如果可以的话可以直接放入当前所在路径或者是path这个参量
            //注意此处的单斜杠在c#中有具体的功能，如果是要当作字符使用则需要在前面再加上一个斜杠 
            //可以在此处加上对于excel版本的判断，对于旧版本的支持
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path + "\\" + file);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //将excel表格中的数据导入到datatable中
            InitColumn2(xlRange, rowCount, colCount);

            //将数据插入datatable中
            InsertData(xlRange, rowCount, colCount);

            //将数据从datatable中插入到SQL Server
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                Console.Write("Please input the datasource\n");
                builder.DataSource = Console.ReadLine();
                Console.Write("Please input the useid\n");
                builder.UserID = Console.ReadLine();
                Console.Write("Please input the password\n");
                builder.Password = Console.ReadLine();
                Console.Write("Please input the initialcatalog\n");
                builder.InitialCatalog = Console.ReadLine();

                //connect to sql server
                Console.Write("connecting to SQL Server...");
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    Console.WriteLine("Done.");

                    //create a new database list
                    Console.Write("Dropping and creating database 'list'...");
                    string sql = "drop database if exists [list];create database [list]";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.ExecuteNonQuery();
                        Console.WriteLine("Done.");
                    }

                    //    //create a new table and insert the data from the datatable



                    //创建一个新的table
                    Console.Write("Creatiing data table from datatable, press any key to continute...");
                    Console.ReadKey(true);
                    StringBuilder sb = new StringBuilder();
                    sb.Append("use list;");
                    sb.Append("if not exists (select * from sysobjects where name='datatable' and xtype='U') create table datatable (");
                    foreach(DataColumn columns in datatable.Columns)
                    {
                        sb.Append(" ");
                        //插入的数据的列名称
                        sb.Append(columns.ColumnName.ToString());
                        sb.Append(" ");
                        //插入列的数据类型
                        if (columns.DataType == System.Type.GetType("System.String"))
                            sb.Append("nvarchar(50)");
                        if (columns.DataType == System.Type.GetType("System.Int64"))
                            sb.Append("bigint");
                        else
                        {
                            if (columns.DataType == System.Type.GetType("System.Int32"))
                                sb.Append("int");
                            else
                            {
                                if (columns.DataType == System.Type.GetType("System.Int16"))
                                    sb.Append("smallint");
                                else
                                    if (columns.DataType == System.Type.GetType("System.Double"))
                                        sb.Append("real");
                            } 
                        }
                        if (columns.DataType==System.Type.GetType("System.DateTime"))
                            sb.Append("datetime");
                        sb.Append(",");
                    }
                    //移除最后一个','
                    sb = sb.Remove( sb.Length - 1,1);
                    sb.Append(");");
                    sql = sb.ToString();
                    Console.Write(sql);
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.ExecuteNonQuery();
                        Console.Write("List creation is done.");
                    }

                    //test
                    foreach (DataRow row2 in datatable.Rows)
                    {
                        Console.Write(row2.ToString());
                        StringBuilder jd = new StringBuilder();
                        jd.Append("SELECT * FROM LIST WHERE ");
                        jd.Append(datatable.Columns[0].ColumnName);
                        jd.Append(" like '");
                        jd.Append(row2[datatable.Columns[0].ColumnName]);
                        jd.Append("';");
                        sql = jd.ToString();
                        using (SqlCommand command = new SqlCommand(sql, connection))
                        {
                            command.ExecuteNonQuery();
                            Console.Write("receive data from sql server\n");
                            SqlDataAdapter rec = new SqlDataAdapter(command);

                            System.Data.DataTable temp = new DataTable();
                            rec.Fill(temp);
                            if ((temp.Rows[1]) == (row2))
                            {
                                Console.Write("The record is existed! Do you want ro recover it? y/n:");
                                char flag = Convert.ToChar(Console.ReadLine());
                                if (flag == 'y')
                                {
                                    StringBuilder di = new StringBuilder();
                                    di.Append("delect from datatable where ");
                                    di.Append(datatable.Columns[0].ColumnName);
                                    di.Append(" = '");
                                    di.Append(row2[datatable.Columns[0]].ToString());
                                    di.Append("';");
                                    sql = di.ToString();
                                    using (SqlCommand command2 = new SqlCommand(sql, connection))
                                    {
                                        command.ExecuteNonQuery();
                                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                                        {
                                            bulkCopy.DestinationTableName =
                                                "dbo.datatable";

                                            try
                                            {
                                                // Write from the source to the destination.
                                                bulkCopy.WriteToServer(datatable);
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //insert the data using the datatable
                                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                                    {
                                        bulkCopy.DestinationTableName =
                                            "dbo.datatable";

                                        try
                                        {
                                            // Write from the source to the destination.
                                            bulkCopy.WriteToServer(datatable);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    ////insert the data using the datatable
                    //using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    //{
                    //    bulkCopy.DestinationTableName =
                    //        "dbo.datatable";

                    //    try
                    //    {
                    //        // Write from the source to the destination.
                    //        bulkCopy.WriteToServer(datatable);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        Console.WriteLine(ex.Message);
                    //    }
                    //}
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

        private static void InsertData(Excel.Range xlRange, int rowCount, int colCount)
        {
            //将数据从excel中读取到datatable中
            DataRow row;
            for (int i = 2; i <= rowCount; i++)
            {
                row = datatable.NewRow();
                for (int j = 0; j <= colCount - 1; j++)
                {
                    //Console.Write(datatable.Columns[i].ColumnName+"\r\n"
                    //cells从1开始计数
                    if (xlRange.Cells[i, j + 1].Value2 != null)
                    {
                        //如果是日期型数据则另外处理存储
                        if (xlRange.Cells[i, j + 1].value is DateTime)
                        {
                            string strValue = xlRange.Cells[i, j + 1].Value2.ToString(); //获取得到数字值
                            //注意数据表中含有的日期数据精确到了小时，所以表示日期应该用double，而不是表示日的int32
                            string strDate = DateTime.FromOADate(Convert.ToDouble(strValue)).ToString("s");
                            ////转成sql server能接受的数据格式
                            //strDate = strDate.Replace("T", " ");
                            row[datatable.Columns[j].ColumnName] = strDate;
                            Console.Write(strDate + "\r\n");
                            Console.Write(datatable.Columns[j].ColumnName + " ");
                        }
                        //将相应列的数据导入到datatable中，注意位置的对应关系，在column中从0开始计数
                        else
                        {
                            row[datatable.Columns[j].ColumnName] = xlRange.Cells[i, j + 1].Value2;
                            Console.Write(xlRange.Cells[i, j + 1].Value2.ToString() + " ");
                            Console.Write(datatable.Columns[j].ColumnName + "\r\n");
                        }
                    }
                }
            }
        }

        private static void InitColumn2(Excel.Range xlRange, int rowCount, int colCount)
        {
            //将excel表格中的数据导入到datatable中 version 2
            //Console.Write(rowCount + "\r\n");
            //Console.Write(colCount + "\r\n'");
            datatable = new DataTable("datatable");
            DataColumn column;
            for (int i = 1; i <= colCount; i++)
            {
                if (xlRange.Cells[1, i].Value2 != null && xlRange.Cells[1, i] != null)
                {
                    column = new DataColumn();
                    column.ColumnName = xlRange.Cells[1, i].Value2.ToString();
                    //if (xlRange.Cells[2, i].Value2 != null && !(xlRange.Cells[2, i].value is DateTime))
                    //    column.DataType = xlRange.Cells[2, i].value.GetType();
                    //else
                    //{
                    //    if (xlRange.Cells[2, i].value is Int32)
                    //        column.DataType = System.Type.GetType("System.Int");
                    //    else
                    //        column.DataType = System.Type.GetType("System.String");
                    //}
                    if (xlRange.Cells[2, i].value != null)
                    {
                        if (xlRange.Cells[2, i].value is DateTime)
                            column.DataType = System.Type.GetType("System.DateTime");
                        else
                            column.DataType = xlRange.Cells[2, i].Value2.GetType();
                    }
                    else
                        column.DataType = System.Type.GetType("System.String");
                    column.ReadOnly = false;
                    column.Unique = false;
                    datatable.Columns.Add(column);
                   
                }
            }
        }

        private static void InitColumn()
        {
            //将excel表格中的数据导入到datatable中 version 1
            //创建并初始化一个datatable
            //声明一个datatable对象
            DataColumn column;
            //DataRow row;
            //创建新的一列
            //此处不能通过第二行的数据对数据类型对列的数据类型进行定义，很多数据都显示为string类型
            //可否尝试通过将通过string to int来判断是否为非string类型数据？
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "自定义码";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "名称";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "规格";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "数量";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "单位";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "购入价";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "购入金额";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "零售价";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "零售金额";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "批号";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "有效期";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "批准文号";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "发票号";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "发票日期";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "生产厂家";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "生产日期";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "送货单号";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "备注";
            column.ReadOnly = false;
            column.Unique = false;
            datatable.Columns.Add(column);
        }
    }

}
