using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Net.NetworkInformation;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Configuration;
using System.Reflection;
using System.Management;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Net.Mail; 


namespace VS_NFD_als_Programm
{
    class Program
    {

       
        static void Main(string[] args)
        {


                        bool pingerfolg;
                        string Servername = "172.20.50.160";

                        
                        pingerfolg = pingdev(Servername);

                        if (pingerfolg == true)
                        {
                            Console.WriteLine("Server erreichbar");
                        }

                        if (pingerfolg == false)
                        {
                            Console.WriteLine("Server nicht erreichbar");

                            sendMail("noreply@banst-pt.de", "Helpdesk-SAP@BAnst-PT.de",
                                  "SAP Archivierung: NFS-Share/An_TSI ist nicht erreichbar", "Bitte T-System kontaktieren!",
                                  "st-mail01.it-banst.int", 25);
                        }

                        int datenanzahl;
                        string ordnername = @"X:\an_tsi";

                        datenanzahl = CountFiles(ordnername);

                        if (datenanzahl == 0)
                        {
                            

                            Console.WriteLine("Der Ordner ist leer");


                            //schreibe CSV Datei
                            string[] fileArray = Directory.GetFiles(@"\\BN-file02\d$\Data\FiBuScan", "010010*.pdf").Select(path => Path.GetFileName(path)).ToArray();
                            string[] fileArrayBK30 = Directory.GetFiles(@"\\BN-file02\d$\Data\FiBuScan", "060030*.pdf").Select(path => Path.GetFileName(path)).ToArray();

                            System.IO.File.WriteAllLines(@"X:\an_tsi\bk10\test.csv", fileArray);

                            System.IO.File.WriteAllLines(@"X:\an_tsi\bk30\test.csv", fileArrayBK30);


                            using (StreamWriter rdr = new StreamWriter(@"X:\an_tsi\bk10\test.csv"))
                            {
                                const string description = "SAP_Eingang";
                                const string separator = "#_#";

                                DirectoryInfo di = new DirectoryInfo(@"\\BN-file02\d$\Data\FiBuScan");
                                FileInfo[] files = di.GetFiles("010010*.pdf");
                                foreach (FileInfo fi in files)
                                {
                                    string[] columns = new string[4];
                                    columns[0] = description;
                                    columns[1] = fi.CreationTime.ToString("yyyy.MM.dd hh:mm:ss");
                                    columns[2] = fi.Name.Substring(0, 20);
                                    columns[3] = fi.Name;
                                    string line = String.Join(separator, columns);

                                    rdr.WriteLine(line);

                                }

                            }

                            //CSV Datei umbennen
                            RenameFile(@"X:\an_tsi\bk10\test.csv", @"X:\an_tsi\bk10\renamed.csv");

                            
                            //BK30
                            using (StreamWriter rdr = new StreamWriter(@"X:\an_tsi\bk30\test.csv"))
                            {
                                const string description = "SAP_Eingang_BK30";
                                const string separator = "#_#";

                                DirectoryInfo di = new DirectoryInfo(@"\\BN-file02\d$\Data\FiBuScan");
                                FileInfo[] files = di.GetFiles("060030*.pdf");
                                foreach (FileInfo fi in files)
                                {
                                    string[] columns = new string[4];
                                    columns[0] = description;
                                    columns[1] = fi.CreationTime.ToString("yyyy.MM.dd hh:mm:ss");
                                    columns[2] = fi.Name.Substring(0, 20);
                                    columns[3] = fi.Name;
                                    string line = String.Join(separator, columns);

                                    rdr.WriteLine(line);

                                }

                            }

                            //CSV Datei umbennen
                            RenameFile(@"X:\an_tsi\bk30\test.csv", @"X:\an_tsi\bk30\renamed.csv");



                            //Schreibe BK10 ready.txt
                            int pdf = Directory.GetFiles(@"\\BN-file02\d$\Data\FiBuScan", "010010*.pdf").Length;
                            string textfile = @"X:\an_tsi\bk10\ready.txt";


                            using (TextWriter writer = File.CreateText(textfile))
                            {
                                writer.WriteLine("Dokumententyp pro Stapel");
                                writer.WriteLine("");
                                writer.WriteLine(pdf);

                            }


                            //Schreibe BK30 ready.txt
                            int pdfBK30 = Directory.GetFiles(@"\\BN-file02\d$\Data\FiBuScan", "060030*.pdf").Length;
                            string textfileBK30 = @"X:\an_tsi\bk30\ready.txt";


                            using (TextWriter writer = File.CreateText(textfileBK30))
                            {
                                writer.WriteLine("Dokumententyp pro Stapel");
                                writer.WriteLine("");
                                writer.WriteLine(pdfBK30);

                            }




                            //Alle PDFs werden nach Share "C:\test2" kopiert
                            Console.WriteLine("Kopiere Dateien");

                            string sourcePath = @"\\BN-file02\d$\Data\FiBuScan";
                            string targetPathBK10 = @"X:\an_tsi\BK10";
                            string targetPathBK30 = @"X:\an_tsi\BK30";
                            string archivPath = @"\\BN-file02\d$\Data\FiBuScan\Archiv";
                            

                            var extensions = new[] { ".pdf" };

                            var pdfsBK10 = (from file in Directory.EnumerateFiles(sourcePath, "010010*.pdf")
                                        where extensions.Contains(Path.GetExtension(file), StringComparer.InvariantCultureIgnoreCase) // comment this out if you don't want to filter extensions
                                        select new
                                        {
                                            Source = file,
                                            Destination = Path.Combine(targetPathBK10, Path.GetFileName(file)),
                                            DestinationArchiv = Path.Combine(archivPath, Path.GetFileName(file))
                                        });

                           var pdfsBK30 = (from file in Directory.EnumerateFiles(sourcePath, "060030*.pdf")
                                           where extensions.Contains(Path.GetExtension(file), StringComparer.InvariantCultureIgnoreCase) // comment this out if you don't want to filter extensions
                                           select new
                                           {
                                               Source = file,
                                               Destination = Path.Combine(targetPathBK30, Path.GetFileName(file)),
                                               DestinationArchiv = Path.Combine(archivPath, Path.GetFileName(file))
                                           });






                foreach (var file in pdfsBK10)
                            {

                                if (File.Exists(file.Destination))
                                {
                                    Console.WriteLine("BK10: Datei " + file.Source + " bereits im Share vorhanden");
                                }

                                else
                                {
                                     File.Copy(file.Source, file.Destination, true);
                  
                                    //Kopierte PDFs werden ins Archiv verschoben
                                    File.Move(file.Source, file.DestinationArchiv);
 
                                }
                            }

                foreach (var file in pdfsBK30)
                {

                    if (File.Exists(file.Destination))
                    {
                        Console.WriteLine("BK30: Datei " + file.Source + " bereits im Share vorhanden");
                    }

                    else
                    {
                        File.Copy(file.Source, file.Destination, true);
                                           
                        //Kopierte PDFs werden ins Archiv verschoben
                        File.Move(file.Source, file.DestinationArchiv);

                    }
                }





            }

                        if (datenanzahl > 0)

                        {
                        Console.WriteLine("Der Ordner enthält Dateien");

                        sendMail("noreply@banst-pt.de", "Helpdesk-SAP@BAnst-PT.de",
                                  "SAP Archivierung: NFS-Share/An_TSI ist voll", "Bitte NFS-Share/An_TSI prüfen!",
                                  "st-mail01.it-banst.int", 25);


                      
                        }

                        

                        var csv = Directory.GetFiles(@"X:\an_banstpt\bk10", "*.csv");
                        var csvBK30 = Directory.GetFiles(@"X:\an_banstpt\bk30", "*.csv");
                        var csvDestination = Directory.GetFiles(@"X:\an_banstpt\bk10", "*.csv").Select(path => Path.GetFileName(path)).ToArray();
                        var csvDestinationBK30 = Directory.GetFiles(@"X:\an_banstpt\bk30", "*.csv").Select(path => Path.GetFileName(path)).ToArray();

                        if (csv.Length == 1 || csv.Length > 1)
                        {


                            Console.WriteLine("Kopiere CSV Dateien von TSI");
                            System.IO.File.Copy(csv[0], @"\\BN-file02\Groups\Finanzen\FiBu-WiPl\1Finanzbuchführung\Archivierung\Ausgabeprotokolle\" + csvDestination[0], true);

                            string directoryPath = @"X:\an_banstpt\bk10";
                            Directory.GetFiles(directoryPath).ToList().ForEach(File.Delete);
                            Directory.GetDirectories(directoryPath).ToList().ForEach(Directory.Delete);

                            Console.WriteLine("Job ohne Fehler abgeschlossen");
                        }


                        else
                        {

                            Console.WriteLine("BK10 CSV von TSI fehlt! Email wird gesendet!");

                            sendMail("noreply@banst-pt.de", "Helpdesk-SAP@BAnst-PT.de",
                                  "SAP Archivierung: BK10 CSV in NFS-Share/An_BAnstPT fehlt", "Bitte T-System kontaktieren, dass BK10 CSV fehlt!",
                                  "st-mail01.it-banst.int", 25);

                            Console.WriteLine("!!!Job mit Fehler abgeschlossen!!!");

                        }

                        //BK30 CSV
                        if (csvBK30.Length == 1 || csvBK30.Length > 1)
                        {


                            Console.WriteLine("Kopiere BK30 CSV Dateien von TSI");
                            System.IO.File.Copy(csvBK30[0], @"\\BN-file02\Groups\Finanzen\FiBu-WiPl\1Finanzbuchführung\Archivierung\Ausgabeprotokolle\" + csvDestinationBK30[0], true);

                            string directoryPath = @"X:\an_banstpt\bk30";
                            Directory.GetFiles(directoryPath).ToList().ForEach(File.Delete);
                            Directory.GetDirectories(directoryPath).ToList().ForEach(Directory.Delete);

                            Console.WriteLine("Job ohne Fehler abgeschlossen");
                        }


                        else
                        {

                            Console.WriteLine("CSV BK30 von TSI fehlt! Email wird gesendet!");

                            sendMail("noreply@banst-pt.de", "Helpdesk-SAP@BAnst-PT.de",
                                  "SAP Archivierung: BK30 CSV in NFS-Share/An_BAnstPT fehlt", "Bitte T-System kontaktieren, dass BK30 CSV fehlt!",
                                  "st-mail01.it-banst.int", 25);

                            Console.WriteLine("!!!Job mit Fehler abgeschlossen!!!");

                        }











        }

        //Methode um CSV umzubennen
        private static void RenameFile(string a, string b)
        {
            
            string lcDate = DateTime.Now.ToString("yyyyMMdd_HH_mm");
            b = lcDate;
            string lcNew = lcDate + Path.GetExtension(a);
            File.Move(a, Path.Combine(Path.GetDirectoryName(a), lcNew));
        }


        public static void sendMail(string absender, string empfaenger,
                              string betreff, string nachricht,
                              string server, int port)
        {
            MailMessage Email = new MailMessage();

            //Absender konfigurieren
            Email.From = new MailAddress(absender);

            //Empfänger konfigurieren
            Email.To.Add(empfaenger);

            //Betreff einrichten
            Email.Subject = betreff;

            //Hinzufügen der eigentlichen Nachricht
            Email.Body = nachricht;

            //Ausgangsserver initialisieren
            SmtpClient MailClient = new SmtpClient(server, port);

            //if (user.Length > 0 && user != string.Empty)
            //{
                //Login konfigurieren
              //  MailClient.Credentials = new System.Net.NetworkCredential(
                //                                                            user, passwort);
            //}

            //Email absenden
            MailClient.Send(Email);
        }
 
      

        public static bool pingdev(string servername)
        {
            bool pingerfolg;
            try
            {
                Ping Sender = new Ping();

                PingReply Result = Sender.Send(servername);

                if (Result.Status == IPStatus.Success)
                    pingerfolg = true;
                else
                    pingerfolg = false;
            }

            catch
            {

                pingerfolg = false;
            }

            return pingerfolg;
        }
       
          // using System.IO;

/// <summary>
/// Counts the files.
/// </summary>
        /// <param name="path">The path.</param>
/// <returns></returns>
        private static int CountFiles(string path)
{
    DirectoryInfo di = new DirectoryInfo(path);
    return di.GetFiles().Length;
}  
       
        }
    }


