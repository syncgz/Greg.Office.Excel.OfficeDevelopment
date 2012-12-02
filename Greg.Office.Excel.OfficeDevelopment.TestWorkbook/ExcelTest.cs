using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace Greg.Office.Excel.OfficeDevelopment.TestWorkbook
{
    public class ExcelTest
    {
        public static void TestFoo()
        {
            AddNamedRange();

            FillExistingNamedRange();

            //AddListObject();

            //BindListObject1();

            BindListObject2();

            BindNamedRange();

            BindNamedRangeByBindingSource();

            AddChart1();

            AddChart2();

            AddActionPane();

            AddNewWorkSheet();

            CopySheet();
        }

        public static void AddNamedRange()
        {
            // Dostep do obiektow Officowych mamy poprzez klase statyczna Globals - jest ona utworzona przez designer.

            Microsoft.Office.Tools.Excel.NamedRange namedRange =
                Globals.Arkusz1.Controls.AddNamedRange(Globals.Arkusz1.get_Range("A1", "E1"), "MyRange");

            // ustawienie wartosci dla calego rangu
            namedRange.Value2 = "Named range value";
        }

        public static void FillExistingNamedRange()
        {
            Globals.Arkusz1.namedRange1.Value2 = "Existing NR";
        }

        public static ListObject AddListObject()
        {
            return Globals.Arkusz2.Controls.AddListObject(Globals.Arkusz2.get_Range("A1", "D1"), "MyListObject");
        }

        public static void BindListObject1()
        {
            //
            // Uwaga!!!!!!!!! - ListObject zawsze dopasowuje sie do zrodla, czyli jezeli zadeklarujemy zrodlo jako 5 kolumnowe a w mappingu przekazemy tylko dwie kolumny to tylko te kolumny
            // zostana wyswietlone
            //
            var lo = AddListObject();

            // podlaczenie zrodla i zmapowaine kolumn wyswietlanych i zrodlowych
            lo.SetDataBinding(CreateDataSource(),null,"Name","Surname");

            // ustawienie nazw kolumn zgodnie z kolumnami datasourca
            lo.AutoSetDataBoundColumnHeaders = true;

            // jezeli chcemy aby zmiany ktore zostana dokonan przez usera w listobjectcie nie mialy odwzorowania w source to odpalamy funkcje:
            lo.Disconnect();

        }

        public static void BindListObject2()
        {
            var lo = AddListObject();

            lo.DataSource = CreateDataSource();

            lo.AutoSetDataBoundColumnHeaders = true;
        }

        public static void BindNamedRange()
        {
            // caly named range uzupelnia sie tylko pierwsza wartoscia z listy...
            Microsoft.Office.Tools.Excel.NamedRange namedRange =
                Globals.Arkusz1.Controls.AddNamedRange(Globals.Arkusz1.get_Range("A3", "A6"), "MyRange1");

            var list = CreateDataSource();

            namedRange.DataBindings.Add(new Binding("Value2", list, "Age"));
        }

        public static void BindNamedRangeByBindingSource()
        {
            // podlaczenie named range do datasource poprzez binding source umozliwia nam przewijanie dancych!!!!
            Microsoft.Office.Tools.Excel.NamedRange namedRange =
                Globals.Arkusz1.Controls.AddNamedRange(Globals.Arkusz1.get_Range("B3", "B6"), "MyRange2");

            var list = CreateDataSource();

            BindingSource bs = new BindingSource(list,null);

            // zbindowanie powoduje ze w named range ustawia sie pierwsza wartosc
            namedRange.DataBindings.Add(new Binding("Value2", bs, "Age"));

            // ustawienie kolejnej wartosci z listy
            bs.MoveNext();
        }

        public static void AddChart1()
        {
            // utworzenie obiektu wykresu
            var chart = Globals.Arkusz3.Controls.AddChart(25, 110, 200, 150, "employees1");

            // ustawienie odpowiedniego typu wykresu
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlAreaStacked;

            // ustawienie zrodla danych dla wykresu
            Excel1.Range chartRange = Globals.Arkusz3.get_Range("A1", "B1");
            chart.SetSourceData(chartRange, Missing.Value);
        }

        public static void AddChart2()
        {
            // utworzenie obiektu wykresu
            var chart = Globals.Arkusz3.Controls.AddChart(25, 110, 200, 150, "employees2");

            // ustawienie odpowiedniego typu wykresu
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlAreaStacked;

            // ustawienie zrodla danych dla wykresu
            Excel1.Range chartRange = Globals.Arkusz3.get_Range("A1", "F6");
            chart.SetSourceData(chartRange, Missing.Value);

        }

        public static void AddActionPane()
        {
            // actionpane control tworzymy tak samo jak customowa kontrolke!!!! w windows forms
            // nastepnie tworzymy ja
            ActionsPaneControl1 actionsPane = new ActionsPaneControl1();

            // dodajemy do kolekcji ActionPane dla danego workbooka
            Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane);
        }

        public static void AddNewWorkSheet()
        {
            var sheet = (Excel1.Worksheet)Globals.ThisWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }

        public static void CopySheet()
        {
            Globals.Arkusz1.Copy(Missing.Value, Globals.ThisWorkbook.Sheets[3]);

            var abc = Globals.ThisWorkbook.Worksheets;
        }


        private static Object CreateDataSource()
        {
            List<Person> list = new List<Person>();

            list.Add(new Person() { Age = 1, Name = "A", Surname = "A" });
            list.Add(new Person() { Age = 2, Name = "A", Surname = "A" });
            list.Add(new Person() { Age = 3, Name = "A", Surname = "A" });
            list.Add(new Person() { Age = 4, Name = "A", Surname = "A" });

            return list;
        }

        //private static Object CreateDataSource1()
        //{
        //    List<Container> list = new List<Container>();


        //}

    }

    public class Person
    {
        public String Name { get; set; }

        public String Surname { get; set; }

        public int Age { get; set; }
    }

    public class Container
    {
        public DateTime Data { get; set; }

        public int Value { get; set; }
    }
}
