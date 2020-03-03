using Aspose.Cells;

using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace testLinqToExcel2
{
    class Program
    {
        static void Main(string[] args)
        {

            Workbook wb = new Workbook(@"C:\TemplateTravel100.xlsb");

            Worksheet ws = wb.Worksheets[2];

            /* System.Collections.IEnumerator iEnum = ws.Cells.GetEnumerator();

             while (iEnum.MoveNext())
             {               

                 Aspose.Cells.Cell cell = (Aspose.Cells.Cell)iEnum.Current;


                 Console.WriteLine(cell.Value);
             }*/

            System.Collections.IEnumerator iEnumRow = ws.Cells.Rows.GetEnumerator();

            RowCollection rows = ws.Cells.Rows;
            List<registro> registros = new List<registro>();

            foreach (Aspose.Cells.Row row in rows)
            {
                registro registro = new registro();
                registro.status= row.GetCellOrNull(0).Value.ToString();
                registro.dob= row.GetCellOrNull(1).Value.ToString();
                registro.firstName= row.GetCellOrNull(2).Value.ToString();
                registro.lastName= row.GetCellOrNull(3).Value.ToString();
                registro.nationality= row.GetCellOrNull(4).Value.ToString();
                registro.passport= row.GetCellOrNull(5).Value.ToString();
                registro.email= row.GetCellOrNull(6).Value.ToString();
                registro.address= row.GetCellOrNull(7).Value.ToString();
                registro.neighborhood= row.GetCellOrNull(8).Value.ToString();
                registro.city= row.GetCellOrNull(9).Value.ToString();
                registro.state= row.GetCellOrNull(10).Value.ToString();
                registro.zipCode= row.GetCellOrNull(11).Value.ToString();
                registro.telephone= row.GetCellOrNull(12).Value.ToString();
                registro.telephone1= row.GetCellOrNull(13).Value.ToString();
                registro.telephone2= row.GetCellOrNull(14).Value.ToString();
                registro.spouseCheck= row.GetCellOrNull(15).Value.ToString();
                registro.childCheck= row.GetCellOrNull(16).Value.ToString();
                registro.companyName = row.GetCellOrNull(17).Value.ToString();

                registros.Add(registro);
                Console.WriteLine($"{registro.status} # {registro.dob} # {registro.firstName} # {registro.lastName} # {registro.nationality} # {registro.passport} # {registro.email} # {registro.address} # {registro.neighborhood} # {registro.city}" +
                    $"{registro.state} # {registro.zipCode} # {registro.telephone} # {registro.telephone1} # {registro.telephone2} # {registro.spouseCheck} # {registro.childCheck} # {registro.companyName}");
            }
            
            Console.ReadKey();

           


            

            

            /*
            using (var excelQueryFactory = new ExcelQueryFactory(@"C:\TemplateTravel100.xlsb"))
            {
                //access your worksheet LINQ way
                var worksheet = excelQueryFactory.Worksheet(@"data");
                //var rows= worksheet.
            }
            */
        }

        public class registro
        {
            public string status { get; set; }
            public string dob { get; set; }
            public string firstName { get; set; }
            public string lastName { get; set; }
            public string nationality { get; set; }
            public string passport { get; set; }
            public string email { get; set; }
            public string address { get; set; }
            public string neighborhood { get; set; }
            public string city { get; set; }
            public string state { get; set; }
            public string zipCode { get; set; }
            public string telephone { get; set; }
            public string telephone1 { get; set; }
            public string telephone2 { get; set; }
            public string spouseCheck { get; set; }
            public string childCheck { get; set; }
            public string companyName { get; set; }
            public registro() {
                
            }

            public registro(string status,string dob,string firstName,string lastName,string nationality,string passport,string email,string address,string neighborhood,string city,string state,string zipCode,string telephone,
                string telephone1,string telephone2,string spouseCheck,string childCheck,string companyName)
            {
                this.status = status;
                this.dob = dob;
                this.firstName = firstName;
                this.lastName = lastName;
                this.nationality = nationality;
                this.passport = passport;
                this.email = email;
                this.address = address;
                this.neighborhood = neighborhood;
                this.city = city;
                this.state = state;
                this.zipCode = zipCode;
                this.telephone = telephone;
                this.telephone1 = telephone1;
                this.telephone2 = telephone2;
                this.spouseCheck = spouseCheck;
                this.childCheck = childCheck;
                this.companyName = companyName;
            
            }

        }
    }

    

}
