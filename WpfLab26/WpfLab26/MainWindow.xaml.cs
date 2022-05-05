using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using Microsoft.Office.Interop.Word;
using System.Reflection;


namespace WpfLab26
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string connectionString;
        SqlDataAdapter adapter;
        System.Data.DataTable ActorsTable;
        public MainWindow()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        }

        private void updateButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateDB();
        }

        private void deleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (actorsGrid.SelectedItems != null)
            {
                for (int i = 0; i < actorsGrid.SelectedItems.Count; i++)
                {
                    DataRowView datarowView = actorsGrid.SelectedItems[i] as DataRowView;
                    if (datarowView != null)
                    {
                        DataRow dataRow = (DataRow)datarowView.Row;
                        dataRow.Delete();
                    }
                }
            }
            UpdateDB();

        }
        private void UpdateDB() 
        {
            SqlCommandBuilder comandbuilder = new SqlCommandBuilder(adapter);
            adapter.Update(ActorsTable);
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string sql = "SELECT * FROM Actors";
            ActorsTable = new System.Data.DataTable();
            SqlConnection connection = null;
            try
            {
                connection = new SqlConnection(connectionString);
                SqlCommand command = new SqlCommand(sql, connection);
                adapter = new SqlDataAdapter(command);

                // установка команды на добавление для вызова хранимой процедуры
                adapter.InsertCommand = new SqlCommand("sp_InsertActor", connection);
                adapter.InsertCommand.CommandType = CommandType.StoredProcedure;
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@fullname", SqlDbType.VarChar, 50, "FullName"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@dateofbirth", SqlDbType.Date, 0, "DateOfBirth"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@gender", SqlDbType.VarChar, 1, "Gender"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@genres", SqlDbType.VarChar, 50, "Genres"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@numberoffilms", SqlDbType.Int, 0, "NumberOfFilms"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@numberofoscars", SqlDbType.Int, 0, "NumberOfOscars"));
                adapter.InsertCommand.Parameters.Add(new SqlParameter("@kinopoiskmark", SqlDbType.Decimal, 2, "KinopoiskMark"));
                SqlParameter parameter = adapter.InsertCommand.Parameters.Add("@Id", SqlDbType.Int, 0, "Id");
                parameter.Direction = ParameterDirection.Output;

                connection.Open();
                adapter.Fill(ActorsTable);
                actorsGrid.ItemsSource = ActorsTable.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection != null)
                    connection.Close();
            }
        }

        private void actorsGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            UpdateDB();
        }

        private void Sord1Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = null;
            string sql = "SELECT * FROM Actors ORDER BY FullName, Genres";
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);
            adapter = new SqlDataAdapter(command);
            connection.Open();
            ActorsTable.Clear();
            adapter.Fill(ActorsTable);
            connection.Close();
        }

        private void Sord2Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = null;
            string sql = "SELECT FullName FROM Actors GROUP BY FullName HAVING COUNT(DateOfBirth) > 1";
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);
            adapter = new SqlDataAdapter(command);
            connection.Open();
            ActorsTable.Clear();
            adapter.Fill(ActorsTable);
            connection.Close();

        }

        private void Sord3Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = null;
            string sql = "SELECT FullName FROM Actors WHERE NumberOfOscars = (SELECT MAX(NumberOfOscars) from Actors)";
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);
            adapter = new SqlDataAdapter(command);
            connection.Open();
            ActorsTable.Clear();
            adapter.Fill(ActorsTable);
            connection.Close();

        }

        private void Sord4Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = null;
            string sql = "SELECT * FROM Actors";
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);
            adapter = new SqlDataAdapter(command);
            connection.Open();
            ActorsTable.Clear();
            adapter.Fill(ActorsTable);
            connection.Close();

        }

        private void Sord5Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection = null;
            string sql = "SELECT * FROM Actors WHERE Gender = 'м'";
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);
            adapter = new SqlDataAdapter(command);
            connection.Open();
            ActorsTable.Clear();
            adapter.Fill(ActorsTable);
            connection.Close();
        }

        private void BtnExportToWord_Click(object sender, RoutedEventArgs e)
        {
            int RowCount = ActorsTable.Rows.Count;
            int ColumnCount = ActorsTable.Columns.Count;
            Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

            //Добавление строк и ячеек
            int r = 0;
            for (int c = 0; c <= ColumnCount - 1; c++)
            {
                for (r = 0; r <= RowCount - 1; r++)
                {
                    DataArray[r, c] = ActorsTable.Rows[r].ItemArray[c];
                }
            }

            Microsoft.Office.Interop.Word.Document oDoc = new Microsoft.Office.Interop.Word.Document();
            oDoc.Application.Visible = true;

            //Ориентация листа
            oDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;


            dynamic oRange = oDoc.Content.Application.Selection.Range;
            string oTemp = "";
            for (r = 0; r <= RowCount - 1; r++)
            {
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oTemp = oTemp + DataArray[r, c] + "\t";

                }
            }

            //Формат таблицы
            oRange.Text = oTemp;
            object oMissing = Missing.Value;
            object Separator = Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
            object ApplyBorders = true;
            object AutoFit = true;
            object AutoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent;

            oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                  Type.Missing, Type.Missing, ref ApplyBorders,
                                  Type.Missing, Type.Missing, Type.Missing,
                                  Type.Missing, Type.Missing, Type.Missing,
                                  Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

            oRange.Select();

            oDoc.Application.Selection.Tables[1].Select();
            oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
            oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
            oDoc.Application.Selection.Tables[1].Rows[1].Select();
            oDoc.Application.Selection.InsertRowsAbove(1);
            oDoc.Application.Selection.Tables[1].Rows[1].Select();

            //Стиль заголовка таблицы
            oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 2;
            oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
            oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

            //add header row manually

            //for (int c = 0; c <= ColumnCount - 1; c++)
            //{
            //    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = moviesTable.Columns[c].HeaderText;
            //}

            //Стили таблицы
            oDoc.Application.Selection.Tables[1].Rows[1].Select();
            oDoc.Application.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            oDoc.Application.Selection.Tables[1].Borders.Enable = 1;



            //Текст шапки
            foreach (Microsoft.Office.Interop.Word.Section section in oDoc.Application.ActiveDocument.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.Text = "Вывод БД";
                headerRange.Font.Size = 16;
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            //Сохранение файла

            oDoc.SaveAs("export.doc", ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
        }

    }
}
