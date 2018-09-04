using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProfPrepOTZ
{
    public partial class MainWind : Form
    {
        private double[,] keyTable = new double[3, 3]
        {
            {348.4, 482.4, 616.4 },
            {416.0, 576.0, 736.0 },
            {400.4, 554.4, 708.4 }
        };



        public MainWind()
        {
            InitializeComponent();
            addRowTable(tableExp);
            addRowTable(resTable);
            addRowTable(medTable);
            chooseGroup.SelectedIndex = 0;
            zvanie.SelectedIndex = 0;
        }


        private void tableExp_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (tableExp.CurrentCellAddress.X == 0)
            {
                try { Convert.ToDateTime(tableExp.CurrentCell.Value).GetType(); }
                catch { tableExp.CurrentCell.Value = 0; MessageBox.Show("Неправильный формат даты \nДолжно быть: 00.00.0000"); }
                return;
            }
            if (tableExp.CurrentCell.Value.ToString().Split('.').Length == 2) tableExp.CurrentCell.Value = tableExp.CurrentCell.Value.ToString().Replace('.', ',');
            try { Convert.ToDouble(tableExp.CurrentCell.Value).GetType(); }
            catch { tableExp.CurrentCell.Value = 0; MessageBox.Show("Неправильный ввод"); }
            tableExp[tableExp.Columns.Count - 1, tableExp.CurrentCellAddress.Y].Value = CalcSumm();
            medTable[tableExp.CurrentCellAddress.X - 1, 0].Value = CalcMed();
            CalcRes();
        }

        private void addRow_Click(object sender, EventArgs e)
        {
            addRowTable(tableExp);
        }

        private void addRowTable(DataGridView table)
        {
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(table);
            for (int ii = 0; ii <= table.Columns.Count - 1; ii++)
            {
                row.Cells[ii].Value = 0;
            }
            if(table.Name == "tableExp") row.Cells[0].Value = DateTime.Now.ToShortDateString();
            table.Rows.Add(row);
            inputNum.Text = table.Rows.Count.ToString();
        }


        //Расчет общего балла
        private void mainCalc_Click(object sender, EventArgs e)
        {
            double sum = 0;
            resultNum.Visible = true;
            verbRes.Visible = true;
            foreach (DataGridViewCell cell in resTable.Rows[0].Cells)
            {
                sum += Convert.ToDouble(cell.Value);
            }
            resultNum.Text = sum.ToString();
            int ind = chooseGroup.SelectedIndex;
            if (sum < keyTable[ind, 0])
            {
                verbRes.Text = "Критический уровень";
                return;
            }
            if (sum >= keyTable[ind, 0] & sum < keyTable[ind, 1])
            {
                verbRes.Text = "Низкий уровень";
                return;
            }
            if (sum >= keyTable[ind, 1] & sum < keyTable[ind, 2])
            {
                verbRes.Text = "Средний уровень";
                return;
            }
            if (sum > keyTable[ind, 2])
            {
                verbRes.Text = "Высший уровень";
                return;
            }
        }

        //Подсчет суммы баллов
        private double CalcSumm()
        {
            int jj = tableExp.CurrentCellAddress.Y;
            double sum = 0;
            for (int ii = 1; ii < tableExp.Columns.Count - 1; ii++)
            {
                sum = sum + Convert.ToDouble(tableExp[ii, jj].Value);
            }
            return sum;
        }

        //Подсчет среднего арифметического
        private double CalcMed()
        {
            double med = 0;
            int xx = tableExp.CurrentCellAddress.X;
            int conut = 0;
            for (int yy = 0; yy <= tableExp.Rows.Count - 1; yy++)
            {
                if (tableExp[xx, yy].Value.ToString() != "") conut++;
                med = med + Convert.ToDouble(tableExp[xx, yy].Value);
            }
            med = Math.Round((double)med / conut, 1);
            if (med <= 2.5) return 0;
            else return med;
        }

        //Подсчет результатов
        private void CalcRes()
        {
            for (int yy = 0; yy <= medTable.Columns.Count - 1; yy++)
            {
                resTable[yy, 0].Value = Math.Round(Convert.ToDouble(medTable[yy, 0].Value) * resTable.Columns[yy].MinimumWidth);
            }
        }

        private void writeFile_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Visible = true;
            Excel.Workbook exBook = exApp.Workbooks.Add();
            Excel.Range oRange;
            Excel.Worksheet workSheetStat = (Excel.Worksheet)exApp.Worksheets.Add();
            workSheetStat.Name = zvanie.Text + " " + nameOffic.Text;
            string str4Group = "";
            switch (chooseGroup.SelectedIndex)
            {
                case 0:
                    str4Group = " (группа командиров взводов)";
                    break;
                case 1:
                    str4Group = " (группа командиров батарей)";
                    break;
                case 2:
                    str4Group = " (группа командиров дивизионов)";
                    break;
            }
            //Первый ряд
            workSheetStat.Cells[1, 1] = "Сводная таблица результатов ПДП" + str4Group;
            oRange = workSheetStat.Range[workSheetStat.Cells[1, 1], workSheetStat.Cells[1, 16]];
            oRange.Merge(Type.Missing);
            //oRange.HorizontalAlignment = HorizontalAlignment.Center;
            //Второй ряд
            workSheetStat.Cells[2, 1] = "Дата проверки";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 1], workSheetStat.Cells[3, 1]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 2] = "Т";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 2], workSheetStat.Cells[2, 3]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 4] = "СУО";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 4], workSheetStat.Cells[2, 6]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 7] = "СП";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 7], workSheetStat.Cells[2, 8]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 9] = "ТП";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 9], workSheetStat.Cells[2, 10]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 11] = "Вожд";
            oRange = workSheetStat.Range[workSheetStat.Cells[2, 11], workSheetStat.Cells[2, 12]];
            oRange.Merge(Type.Missing);
            oRange.HorizontalAlignment = HorizontalAlignment.Center;
            workSheetStat.Cells[2, 13] = "ОП";
            workSheetStat.Cells[2, 14] = "ФП";
            workSheetStat.Cells[2, 15] = "Моб";
            workSheetStat.Cells[2, 16] = "ОГП";
            //Третий ряд
            workSheetStat.Cells[3, 2] = "Устн";
            workSheetStat.Cells[3, 3] = "Практ";
            workSheetStat.Cells[3, 4] = "ПРЗ";
            workSheetStat.Cells[3, 5] = "ВМП";
            workSheetStat.Cells[3, 6] = "ВОгЗ";
            workSheetStat.Cells[3, 7] = "Устн";
            workSheetStat.Cells[3, 8] = "Практ";
            workSheetStat.Cells[3, 9] = "Устн";
            workSheetStat.Cells[3, 10] = "Практ";
            workSheetStat.Cells[3, 11] = "Устн";
            workSheetStat.Cells[3, 12] = "Практ";
            workSheetStat.Cells[3, 13] = "Практ";
            workSheetStat.Cells[3, 14] = "Практ";
            workSheetStat.Cells[3, 15] = "Устн";
            workSheetStat.Cells[3, 16] = "Устн";
            //Заполнение из таблицы ввода результатов
            for (int ii = 0; ii < tableExp.Rows.Count; ii++)
            {
                for(int jj = 0; jj < tableExp.Columns.Count-1; jj++)
                {
                    workSheetStat.Cells[ii + 3, jj+1] = tableExp[jj, ii].Value.ToString();
                }
            }
            //Заполнение из средних значений
            workSheetStat.Cells[tableExp.Rows.Count + 3, 1] = "Среднее значение";
            for (int ii = 0; ii < medTable.Columns.Count; ii++)
            {
                workSheetStat.Cells[tableExp.Rows.Count + 3, ii + 2] = medTable[ii, 0].Value.ToString();
            }
            //Заполнение из результата
            workSheetStat.Cells[tableExp.Rows.Count + 4, 1] = "Результат";
            for (int ii = 0; ii < resTable.Columns.Count; ii++)
            {
                workSheetStat.Cells[tableExp.Rows.Count + 4, ii + 2] = resTable[ii, 0].Value.ToString();
            }

            //Заполнение результата
            string str4res = "Общий результат: " + resultNum.Text + " - " + verbRes.Text;
            workSheetStat.Cells[tableExp.Rows.Count + 5, 1] = str4res;
            oRange = workSheetStat.Range[workSheetStat.Cells[tableExp.Rows.Count + 5, 1], workSheetStat.Cells[tableExp.Rows.Count + 5, 16]];
            oRange.Merge(Type.Missing);



            //SaveFileDialog save = new SaveFileDialog();
            //save.Title = ("Сохранить как...");
            //save.Filter = "Excel Document (*.xlsx) | *.xlsx";
            //save.OverwritePrompt = true;
            //if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{

            //}

        }
    }
}
