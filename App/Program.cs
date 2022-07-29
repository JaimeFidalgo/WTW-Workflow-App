using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.SqlServer;
using System.Linq;
using Microsoft.Office.Interop.Excel;



namespace App
{
    class Program
    {
        static void Main(string[] args)
        {
            // START: CONNECTION TO DATABASE

            using (var db = new AppLogicContext())
            {
                 // END: CONNECTION TO DATABASE 

                Console.WriteLine("1. Connected to DataBase: WTW-WorkFlow-Example");

                // START: GET DATA FROM DATABASE

                var users = db.Users.AsNoTracking();

                //END: GET DATA FROM DATABASE

                List<App.Users> noValidUsers = new List<App.Users>();
                List<App.Users> validUsers = new List<App.Users>();

                // START: FILTER DATA

                if (users != null)
                {

                    foreach (var user in users)
                    {
                        if (user.Age <= 0)
                        {
                            noValidUsers.Add(user);
                        }
                        else
                        {
                            validUsers.Add(user);
                        }
                    }
                    Console.WriteLine("2. Filtered Data");
                }

                // END: FILTER DATA

                // START: INSERT DATA TO PROPER DATABASE

                foreach (var noVal in noValidUsers)
                {
                    var userToAdd = new IncorrectUsers
                    {
                        Name = noVal.Name,
                        Surname = noVal.Surname,
                        Age = noVal.Age,
                        InsuredAmount = noVal.InsuredAmount
                    };

                    db.IncorrectUsers.Add(userToAdd);
                    var transactionState = db.SaveChanges();
                }

                foreach (var valid in validUsers)
                {
                    var userToAdd = new ValidatedUsers
                    {
                        Name = valid.Name,
                        Surname = valid.Surname,
                        Age = valid.Age,
                        InsuredAmount = valid.InsuredAmount
                    };

                    db.ValidatedUsers.Add(userToAdd);
                    var transactionState = db.SaveChanges();
                }

                Console.WriteLine("3. Opened Excel App");

                Application xlApp = new Application(); 
                xlApp.Visible = true;
                Workbook sampleWorkbook = xlApp.Workbooks.Open(@"C:/Users/jaime/Desktop/.NET/01-Proyectos/WTW-WorkFlow-Example/RAFM-Simulator.xlsx");
                Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sampleWorkbook.Worksheets["Data"];

                string[] columns = { "A", "B", "C", "D", "E" };

                //START:INTRODUCE DATA INTO EXCEL

                for (var i = 0; i < validUsers.Count; i++)
                {

                    worksheet.Range["A" + (2 + i)].Value = validUsers[i].UserId;
                }

                for (var i = 0; i < validUsers.Count; i++)
                {

                    worksheet.Range["B" + (2 + i)].Value = validUsers[i].Name;
                }

                for (var i = 0; i < validUsers.Count; i++)
                {

                    worksheet.Range["C" + (2 + i)].Value = validUsers[i].Surname;
                }

                for (var i = 0; i < validUsers.Count; i++)
                {

                    worksheet.Range["D" + (2 + i)].Value = validUsers[i].Age;
                }


                for (var i = 0; i < validUsers.Count; i++)
                {

                    worksheet.Range["E" + (2 + i)].Value = validUsers[i].InsuredAmount;
                }

                //END: ITRODUCE DATA INTO EXCEL

                Console.WriteLine("4. Inserted Values to Excel ");
                Console.WriteLine("5. Generated results by Excel App");

                //START: RETRIEVE RESULTS FROM EXCEL

                string[] columnsReport = { "A", "B", "C", "D", "E", "F", "G" };

                for (var row = 0; row < validUsers.Count; row++)
                {

                    var UserToAdd = new FinalReport
                    {

                        Name = (String)worksheet.Range["B" + (row + 2)].Value,
                        Surname = (String)worksheet.Range["C" + (row + 2)].Value,
                        Age = Convert.ToInt32(worksheet.Range["D" + (row + 2)].Value),
                        InsuredAmount = Convert.ToDecimal(worksheet.Range["E" + (row + 2)].Value),
                        DeathProbability = Convert.ToDecimal(worksheet.Range["F" + (row + 2)].Value),
                        Provision = Convert.ToDecimal(worksheet.Range["G" + (row + 2)].Value)

                    };

                    //END: RETRIEVE RESULTS FROM EXCEL
                    //START: INTRODUCE RESULTS INTO DATABASE

                    db.FinalReport.Add(UserToAdd);
                    var transactionState = db.SaveChanges();
                    // Console.WriteLine(transactionState);

                    //END: INTRODUCE RESULTS INTO DATABASE

                    //START: SHOW RESULTS IN CONSOLE
                }
                Console.WriteLine("6. Retrieved Data from Excel and Inserted into DataBase");
                Console.WriteLine("7. Closed connection to the Database");
                Console.WriteLine("");
                Console.WriteLine("*********RESULTS*********");
                Console.WriteLine("");

                var report = db.FinalReport.AsNoTracking();
                decimal totalProvision = 0;
                foreach (var repo in report)
                {
                    Console.WriteLine($"User: {repo.Name} {repo.Surname}" + " --- " + $"Provision: {repo.Provision}  euros");
                    totalProvision += repo.Provision;
                }
                Console.WriteLine("");
                Console.WriteLine("Total Provision" + " --- " + totalProvision + " euros");
                Console.WriteLine("");
                Console.WriteLine("*********INCORRECT DATA*********");
                Console.WriteLine("");

                var incorrects = db.IncorrectUsers.AsNoTracking();

                foreach (var inc in incorrects)
                {
                    Console.WriteLine($"Data from User: {inc.Name} {inc.Surname} is invalid");
                }

                Console.WriteLine("");
                Console.WriteLine("Warn the Insurance Company to check");
                Console.WriteLine("");
            }

            //END: SHOW RESULTS IN CONSOLE
        }
        //CLOSE CONNECTION TO DATABASE

    }
















    //TODO: Exporta data to excel and retrieve the results to pass them to the database.
}


