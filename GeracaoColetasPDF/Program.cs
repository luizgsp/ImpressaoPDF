using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using GeracaoColetasPDF.Entities;

namespace GeracaoColetasPDF
{
    public class Program
    {
        //private StreamReader streamToPrint;
        private static string FirstLettre { get; set; }

        private static bool Print { get; set; }

        private static string Seller { get; set; }
        private static string Buyer { get; set; }
        private static string Customer { get; set; }
        private static string NumOrder { get; set; }
        private static string CnpjCpf { get; set; }
        private static string EmailCustomer { get; set; }
        private static string Note1 { get; set; }
        private static string Note2 { get; set; }
        private static BaseColor FontColor { get; set; }
        private static bool AlternateColor { get; set; }
        private static int Count { get; set; } = 0;


        static void Main(string[] args)
        {
            Process aProcess = Process.GetCurrentProcess();
            string aProcName = aProcess.ProcessName;

            if (Process.GetProcessesByName(aProcName).Length > 1)
            {
                Console.WriteLine("O programa já está em execução!");
                System.Threading.Thread.Sleep(5000);
                return;
            }

            while (ListaArquivos() != 0)
            {
                Console.WriteLine(@"\t\t\t\tRelatorio impresso...");
            }
        }

        private static int ListaArquivos()
        {

            System.Threading.Thread.Sleep(2000);

            Config config = new Config();
            config.GetConfig();

            Console.Clear();
            Console.WriteLine("\t\t\tMinas Ferramentas Ltda.");
            Console.WriteLine();
            Console.WriteLine("\t      Impressao e Envio de emails de Coletas em PDF");
            Console.WriteLine();
            Console.WriteLine("\t     Belo Horizonte, {0}", DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy HH:mm:ss"));
            Console.WriteLine();


            DirectoryInfo directory = new DirectoryInfo(config.SourcePath);
            if (!directory.Exists)
            {
                Console.WriteLine("Não encontrado o diretorio especificado!");
                Print = false;
                return 0;
            }

            FileInfo[] files = directory.GetFiles("*.txt");
            Console.WriteLine();
            Console.WriteLine(@"Arquivos.: {0}", files.Count());
            Console.WriteLine();
            if (files.Count() == 0)
            {
                try
                {
                    files = directory.GetFiles("*.pdf");
                    if (files.Count() > 0)
                    {
                        foreach (FileInfo fileinfo in files)
                        {
                            System.Threading.Thread.Sleep(5000);
                            File.Move(config.SourcePath + fileinfo, config.TargetPath + fileinfo);
                            Console.WriteLine(@"Movido o Arquivo {0} para a pasta Impressos.", fileinfo);

                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    System.Threading.Thread.Sleep(10000);

                }
                return 1;
            }

            foreach (FileInfo fileinfo in files)
            {
                decimal FileLenth = fileinfo.Length / 1024;
                string DadosDoArquivo = fileinfo.Name + "\t" + FileLenth.ToString("##,##0.00") + " kb" +
                    "\t" + fileinfo.LastWriteTime;
                if (FileLenth > 0)
                {
                    Console.WriteLine(@"Txt {0}", DadosDoArquivo);
                    Count++;
                    //if (Imprime)
                    GerarPDF(fileinfo.Name);
                    // this.Refresh();
                }
                else
                {
                    fileinfo.Delete();
                }
            }

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("\t\t\t\tAguarde....");
            return 1;
        }

        

        //============= gera PDF para enviar por e-mail ============================================================================================
        
    }
}
